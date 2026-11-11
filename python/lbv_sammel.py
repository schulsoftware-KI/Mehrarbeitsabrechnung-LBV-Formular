# -*- coding: utf-8 -*-
"""
lbv_sammel.py - Mehrere LBV-PDFs zu EINER Sammeldatei fuer alle Lehrer zusammenfuegen.

Verwendbar fuer
  - Einzelformulare  (LBV_Mehrarbeit_<Kuerzel>_<MM>_<JJJJ>.pdf, aus lbv_fill)
  - Gesamtformulare  (LBV_Mehrarbeit_<Kuerzel>_Gesamt.pdf,      aus lbv_merge)

WICHTIG - warum "flatten":
  Alle LBV-Formulare verwenden dieselben Feldnamen (LBV_Nachname, FeldA_Input1 ...).
  Haengt man sie einfach aneinander, zeigt ein PDF-Viewer gleichnamige Felder
  mit demselben Wert an - z.B. auf jeder Seite den Namen des ersten Lehrers.
  Darum werden die sichtbaren Feldinhalte jeder Quelldatei fest in den
  Seiteninhalt uebernommen und die Formularfelder danach entfernt.
  Uebernommen wird genau das, was auch beim Drucken sichtbar ist
  (Print-Flag gesetzt, Hidden-Flag nicht gesetzt).
  Die Einzeldateien selbst bleiben unveraendert.

Aufruf: lbv_sammel.py <json_pfad>
JSON-Format:
{
  "ausgabe_pfad": "C:/.../LBV_Sammeldatei_Einzelformulare_2025-11-11.pdf",
  "log_pfad":     "C:/.../lbv_sammel_log.txt",          (optional)
  "quellen":      ["C:/.../LBV_Mehrarbeit_AB_09_2025.pdf", ...]
}
Die Reihenfolge in "quellen" ist die Reihenfolge in der Sammeldatei.

Exit-Codes:
  0 = Sammeldatei erstellt
  2 = Sammeldatei erstellt, aber einzelne Quelldateien fehlerhaft (siehe Log)
  1 = Fehler, keine Sammeldatei
"""
import sys
import os
import re
import json
import shutil
import tempfile
import traceback
from io import BytesIO

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    NameObject, ArrayObject, DictionaryObject, FloatObject,
    IndirectObject, StreamObject, DecodedStreamObject
)

_LOG = []


def log(msg):
    _LOG.append(str(msg))
    try:
        print(msg, flush=True)
    except Exception:
        pass


def log_schreiben(log_pfad):
    """Log in Windows-1252 schreiben, damit VBA (Line Input) Umlaute korrekt anzeigt."""
    if not log_pfad:
        return
    try:
        with open(log_pfad, 'w', encoding='cp1252', errors='replace') as f:
            f.write('\n'.join(_LOG) + '\n')
    except Exception:
        pass


# ----------------------------------------------------------------
#  Hilfsfunktionen
# ----------------------------------------------------------------
def _raw(d, key):
    """Eintrag eines Dictionaries ohne Aufloesen der indirekten Referenz."""
    if hasattr(d, 'raw_get'):
        try:
            return d.raw_get(key)
        except KeyError:
            return None
    return dict.get(d, key)


def _wert(d, key, standard=None):
    """Eintrag eines Dictionaries, indirekte Referenzen aufgeloest."""
    try:
        return d[key] if key in d else standard
    except Exception:
        return standard


def _als_referenz(writer, obj):
    """Gibt eine indirekte Referenz zurueck (Streams muessen indirekt sein)."""
    if isinstance(obj, IndirectObject):
        return obj
    return writer._add_object(obj)


def _zahlen(arr, standard):
    try:
        werte = [float(x) for x in arr]
        return werte if len(werte) == len(standard) else standard
    except Exception:
        return standard


def _ap_stream_waehlen(annot):
    """
    Liefert (Referenz, Stream) der Normal-Darstellung (/AP /N) eines Widgets.
    Bei Checkboxen/Radiobuttons ist /N ein Dictionary mit Zustaenden;
    dann entscheidet /AS (bzw. /V des Feldes).
    """
    ap = annot.get('/AP')
    if ap is None:
        return None, None
    ap = ap.get_object()
    n_raw = _raw(ap, '/N')
    if n_raw is None:
        return None, None
    n_obj = n_raw.get_object()
    if isinstance(n_obj, StreamObject):
        return n_raw, n_obj
    if isinstance(n_obj, DictionaryObject):
        zustand = annot.get('/AS')
        if zustand is None:
            zustand = annot.get('/V')
            if zustand is None and annot.get('/Parent') is not None:
                zustand = annot['/Parent'].get_object().get('/V')
        if zustand is not None and zustand in n_obj:
            s_raw = _raw(n_obj, zustand)
            s_obj = s_raw.get_object()
            if isinstance(s_obj, StreamObject):
                return s_raw, s_obj
    return None, None


def _font_ressourcen_ergaenzen(stream, dr):
    """
    Fehlen dem Darstellungs-Stream Schrift-Ressourcen, die Standard-Ressourcen
    des Formulars (/AcroForm /DR) eintragen - so wie Acrobat es beim Anzeigen macht.
    """
    if dr is None:
        return
    dr = dr.get_object()
    res = stream.get('/Resources')
    if res is None:
        neu = DictionaryObject()
        for k in dr.keys():
            neu[NameObject(k)] = _raw(dr, k)
        stream[NameObject('/Resources')] = neu
        return
    res = res.get_object()
    if '/Font' not in res and '/Font' in dr:
        res[NameObject('/Font')] = _raw(dr, '/Font')


def flatten_seite(writer, page, dr, zaehler):
    """
    Uebernimmt alle druckbaren Widget-Darstellungen der Seite als Form-XObjects
    in den Seiteninhalt und entfernt die Widgets. Gibt den neuen Zaehlerstand zurueck.
    """
    annots_raw = page.get('/Annots')
    if annots_raw is None:
        return zaehler
    annots = annots_raw.get_object()

    behalten = ArrayObject()
    ops = []
    xobjekte = []

    for ref in annots:
        annot = ref.get_object()
        if annot.get('/Subtype') != '/Widget':
            behalten.append(ref)              # Links o.ae. bleiben erhalten
            continue

        # Fehlt /F, gilt das Feld wie in lbv_fill/lbv_merge als druckbar (4)
        try:
            flags = int(_wert(annot, '/F', 4))
        except Exception:
            flags = 4
        if (flags & 2) or not (flags & 4):    # Hidden oder nicht druckbar
            continue

        s_ref, stream = _ap_stream_waehlen(annot)
        if stream is None:
            continue
        try:
            if not stream.get_data().strip():
                continue                      # leere Darstellung
        except Exception:
            continue

        rx1, ry1, rx2, ry2 = _zahlen(_wert(annot, '/Rect', []), [0, 0, 0, 0])
        rx1, rx2 = min(rx1, rx2), max(rx1, rx2)
        ry1, ry2 = min(ry1, ry2), max(ry1, ry2)
        if rx2 - rx1 <= 0 or ry2 - ry1 <= 0:
            continue

        bbox = _zahlen(_wert(stream, '/BBox', []), [0, 0, rx2 - rx1, ry2 - ry1])
        matrix = _zahlen(_wert(stream, '/Matrix', []), [1, 0, 0, 1, 0, 0])

        # PDF-Spezifikation 12.5.5: BBox mit Matrix transformieren,
        # umschliessendes Rechteck auf /Rect abbilden.
        a, b, c, d, e, f = matrix
        ecken = [(bbox[0], bbox[1]), (bbox[2], bbox[1]),
                 (bbox[0], bbox[3]), (bbox[2], bbox[3])]
        px = [a * x + c * y + e for x, y in ecken]
        py = [b * x + d * y + f for x, y in ecken]
        bx1, bx2, by1, by2 = min(px), max(px), min(py), max(py)
        if bx2 - bx1 <= 0 or by2 - by1 <= 0:
            continue
        sx = (rx2 - rx1) / (bx2 - bx1)
        sy = (ry2 - ry1) / (by2 - by1)
        tx = rx1 - bx1 * sx
        ty = ry1 - by1 * sy

        # Stream als vollstaendiges Form-XObject kennzeichnen
        stream[NameObject('/Type')] = NameObject('/XObject')
        stream[NameObject('/Subtype')] = NameObject('/Form')
        if '/BBox' not in stream:
            stream[NameObject('/BBox')] = ArrayObject([FloatObject(v) for v in bbox])
        _font_ressourcen_ergaenzen(stream, dr)

        zaehler += 1
        name = '/LbvFlat%d' % zaehler
        xobjekte.append((name, _als_referenz(writer, s_ref)))
        ops.append('q %.5f 0 0 %.5f %.5f %.5f cm %s Do Q' % (sx, sy, tx, ty, name))

    # Widgets entfernen (auch nicht gedruckte -> Seite sieht aus wie im Ausdruck)
    if len(behalten) > 0:
        page[NameObject('/Annots')] = behalten
    elif '/Annots' in page:
        del page['/Annots']

    if not ops:
        return zaehler

    # Ressourcen der Seite kopieren und XObjects eintragen
    res_alt = page.get('/Resources')
    res_neu = DictionaryObject()
    if res_alt is not None:
        res_alt = res_alt.get_object()
        for k in res_alt.keys():
            res_neu[NameObject(k)] = _raw(res_alt, k)
    xo_neu = DictionaryObject()
    xo_alt = res_neu.get('/XObject')
    if xo_alt is not None:
        xo_alt = xo_alt.get_object()
        for k in xo_alt.keys():
            xo_neu[NameObject(k)] = _raw(xo_alt, k)
    for name, ref in xobjekte:
        xo_neu[NameObject(name)] = ref
    res_neu[NameObject('/XObject')] = xo_neu
    page[NameObject('/Resources')] = res_neu

    # Seiteninhalt: alter Inhalt in q/Q kapseln, danach Feldinhalte zeichnen
    vor = DecodedStreamObject()
    vor.set_data(b'q\n')
    nach = DecodedStreamObject()
    nach.set_data(b'\nQ\n' + '\n'.join(ops).encode('latin-1') + b'\n')

    inhalt = ArrayObject()
    inhalt.append(writer._add_object(vor))
    c_raw = page.get('/Contents')
    if c_raw is not None:
        c_obj = c_raw.get_object()
        if isinstance(c_obj, ArrayObject):
            for teil in c_obj:
                inhalt.append(_als_referenz(writer, teil))
        else:
            inhalt.append(_als_referenz(writer, c_raw))
    inhalt.append(writer._add_object(nach))
    page[NameObject('/Contents')] = inhalt

    return zaehler


def flatten_datei(pfad):
    """Liest eine LBV-PDF, flattet alle Seiten und gibt die PDF als Bytes zurueck."""
    reader = PdfReader(pfad)
    w = PdfWriter()
    w.append(reader)

    root = w._root_object
    dr = None
    if '/AcroForm' in root:
        acro = root['/AcroForm'].get_object()
        dr = _raw(acro, '/DR')

    zaehler = 0
    for page in w.pages:
        zaehler = flatten_seite(w, page, dr, zaehler)

    if '/AcroForm' in root:
        del root['/AcroForm']

    buf = BytesIO()
    w.write(buf)
    buf.seek(0)
    return buf, len(w.pages)


_RE_EINZEL = re.compile(r'^LBV_Mehrarbeit_(.+)_(\d{2})_(\d{4})$', re.IGNORECASE)
_RE_GESAMT = re.compile(r'^LBV_Mehrarbeit_(.+)_Gesamt$', re.IGNORECASE)


def lesezeichen_titel(pfad):
    stamm = os.path.splitext(os.path.basename(pfad.replace('\\', '/')))[0]
    m = _RE_EINZEL.match(stamm)
    if m:
        return '%s - %s/%s' % (m.group(1), m.group(2), m.group(3))
    m = _RE_GESAMT.match(stamm)
    if m:
        return '%s - Gesamt' % m.group(1)
    return stamm


# ----------------------------------------------------------------
#  Hauptfunktion
# ----------------------------------------------------------------
def sammeln(data):
    ausgabe = data['ausgabe_pfad']
    quellen = data.get('quellen') or []

    if not quellen:
        log('FEHLER: Keine Quelldateien angegeben.')
        return 1

    gesamt = PdfWriter()
    ok = 0
    fehler = 0

    for src in quellen:
        name = os.path.basename(src.replace('\\', '/'))
        try:
            buf, seiten = flatten_datei(src)
            start = len(gesamt.pages)
            gesamt.append(PdfReader(buf), import_outline=False)
            try:
                gesamt.add_outline_item(lesezeichen_titel(src), start)
            except Exception:
                pass                          # Lesezeichen sind nicht wichtig
            ok += 1
            log('  OK    %s (%d Seite%s)' % (name, seiten, '' if seiten == 1 else 'n'))
        except Exception as e:
            fehler += 1
            log('  FEHLER %s: %s: %s' % (name, type(e).__name__, e))

    if ok == 0:
        log('FEHLER: Keine Quelldatei konnte verarbeitet werden.')
        return 1

    root = gesamt._root_object
    if '/AcroForm' in root:
        del root['/AcroForm']
    try:
        gesamt.page_mode = '/UseOutlines'     # Lesezeichenleiste beim Oeffnen zeigen
    except Exception:
        pass

    # Identische Objekte (Schriften, Formular-Hintergrund) nur einmal speichern
    if hasattr(gesamt, 'compress_identical_objects'):
        try:
            gesamt.compress_identical_objects()   # Standard: Duplikate + Waisen entfernen
        except Exception as e:
            log('Hinweis: Komprimierung uebersprungen (%s)' % e)

    # Erst in TEMP schreiben, dann verschieben (verhindert OneDrive-Konflikte)
    tmp_fd, tmp_pfad = tempfile.mkstemp(suffix='.pdf')
    os.close(tmp_fd)
    try:
        with open(tmp_pfad, 'wb') as f:
            gesamt.write(f)
        if os.path.exists(ausgabe):
            os.remove(ausgabe)
        shutil.move(tmp_pfad, ausgabe)
    except Exception:
        if os.path.exists(tmp_pfad):
            os.remove(tmp_pfad)
        raise

    log('OK: %d Datei(en), %d Seite(n) -> %s' % (ok, len(gesamt.pages), ausgabe))
    if fehler:
        log('ACHTUNG: %d Datei(en) konnten nicht uebernommen werden.' % fehler)
        return 2
    return 0


def main():
    if len(sys.argv) < 2:
        print('Aufruf: lbv_sammel.py <json_pfad>')
        return 1
    log_pfad = None
    try:
        with open(sys.argv[1], 'r', encoding='utf-8') as f:
            data = json.load(f)
        log_pfad = data.get('log_pfad')
        code = sammeln(data)
    except Exception as e:
        log('FEHLER: %s: %s' % (type(e).__name__, e))
        log(traceback.format_exc())
        code = 1
    log_schreiben(log_pfad)
    return code


if __name__ == '__main__':
    sys.exit(main())
