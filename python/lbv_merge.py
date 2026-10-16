# -*- coding: utf-8 -*-
"""
lbv_merge.py - Mehrere LBV-Monatsdateien zu einer Zusammenfassung zusammenführen
Liest val_a (MMJJ SSS S) und Merkmal aus jeder Quelldatei (FeldA_Input1),
schreibt sie in Zeilen FeldA_Input1..FeldA_Input6 der Vorlage.
Aufruf: lbv_merge.py <json_pfad>
JSON-Format:
{
  "vorlage_pfad": "...",
  "ausgabe_pfad": "...",
  "quellen": ["monat1.pdf", "monat2.pdf", ...]
}
"""
import sys, json
from reportlab.graphics.barcode import code39 as _rl_code39
from reportlab.pdfgen import canvas as _rl_canvas
REPORTLAB_OK = True
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    NameObject, TextStringObject, BooleanObject, DecodedStreamObject
)

MAX_ZEILEN = 6   # FeldA_Input1 bis FeldA_Input6

# Kästchen-Struktur für FeldA (identisch zu lbv_fill.py)
KAESTCHEN = (
    b'q\n0.5 0.5 126.55 13.428 re\ns\n'
    b'12.755 13.428 m\n12.755 0.5 l\n'
    b'25.51 13.428 m\n25.51 0.5 l\n'
    b'38.265 13.428 m\n38.265 0.5 l\n'
    b'51.02 13.428 m\n51.02 0.5 l\n'
    b'63.775 13.428 m\n63.775 0.5 l\n'
    b'76.53 13.428 m\n76.53 0.5 l\n'
    b'89.285 13.428 m\n89.285 0.5 l\n'
    b'102.04 13.428 m\n102.04 0.5 l\n'
    b'114.795 13.428 m\n114.795 0.5 l\n'
    b's\nQ\n'
)

BARCODE_PRAEFIX = "ABES23/201902/"


def make_ap(val):
    """AP-Stream für FeldA-Wertfeld (10 Zeichen: MMJJ SSS S)"""
    p = ['/Tx BMC ', 'q', '1 1 125.55 12.428 re', 'W', 'n',
         'BT', '/CoBo 12 Tf', '0 g', '3.96 4.162 Td']
    sp = False
    for i, ch in enumerate(val):
        if i > 0:
            if sp:
                p.append('10.300 0 Td')
            elif ch == ' ':
                p.append('15.330 0 Td')
            else:
                p.append('12.755 0 Td')
        p.append(f'({ch}) Tj')
        sp = (ch == ' ')
    p.extend(['ET', 'Q', 'EMC', ''])
    txt = '\n'.join(p).encode('latin-1')
    o = DecodedStreamObject()
    o._data = txt  # kein KAESTCHEN - keine Kaestchen im Ausdruck
    return o


def make_ap_merkmal(val):
    """AP-Stream fuer Merkmal-Feld (1 Zeichen) - exakt nach Adobe-Standard"""
    # Box-Rahmen + Text, Werte aus manuell befuellter Vorlage extrahiert
    KAESTCHEN_M = b'q\n0.5 0.5 11.746 13.428 re\ns\ns\nQ\n'
    p = ['/Tx BMC ', 'q', '1 1 10.746 12.428 re', 'W', 'n',
         'BT', '/CoBo 12 Tf', '0 g', '2.821 4.162 Td']
    if val:
        p.append(f'({val[0]}) Tj')
    p.extend(['ET', 'Q', 'EMC', ''])
    txt = '\n'.join(p).encode('latin-1')
    o = DecodedStreamObject()
    o._data = txt  # kein KAESTCHEN_M - kein Kaestchen um das A
    return o


def add_barcode(writer, pers_nr):
    """
    Fuegt Code-39 (3of9) Barcode an exakter Position der PDF-Vorlage ein.
    Exakte Masse aus LBV_ABes_23_2019_02_release.pdf:
      LBV_Barcode_9of3: Rect=[10.56, 83.43, 71.43, 533.53] pt, /MK/R=270
      LBV_Barcode_Text: Rect=[0.00, 218.88, 35.21, 326.69] pt, /MK/R=270
    Benoetigt: pip install reportlab
    """
    if not REPORTLAB_OK:
        print("Warnung: reportlab nicht installiert - Barcode wird uebersprungen.")
        return
    code39 = _rl_code39
    canvas = _rl_canvas
    try:
        from io import BytesIO
        from pypdf import PdfReader as PR
        pers_nr_clean = pers_nr.replace(" ", "").replace("-", "")
        barcode_val   = BARCODE_PRAEFIX + pers_nr_clean

        page_w = float(writer.pages[0].mediabox.width)
        page_h = float(writer.pages[0].mediabox.height)

        buf = BytesIO()
        c   = canvas.Canvas(buf, pagesize=(page_w, page_h))

        # --- Exakte Feldkoordinaten aus der Vorlage ---
        bx1, by1, bx2, by2 = 10.56, 83.43, 71.43, 533.53   # LBV_Barcode_9of3
        field_w = bx2 - bx1   # 60.87 pt -> Balkenhoehe
        field_h = by2 - by1   # 450.10 pt -> Gesamtlaenge

        # Barcode-Objekt erstellen um exakte Breite zu ermitteln
        # Dann barWidth so skalieren dass der Barcode genau field_h lang wird
        probe = code39.Standard39(barcode_val, barHeight=1, barWidth=1,
                                  humanReadable=False, quiet=False)
        natural_width = probe.width   # Breite bei barWidth=1
        bar_w = (field_h - 2) / natural_width   # Skalierungsfaktor
        # Barcode-Breite: 55% der Feldbreite (Benutzeranforderung)
        bar_h = field_w * 0.385

        bc = code39.Standard39(
            barcode_val,
            barHeight     = bar_h,
            barWidth      = bar_w,
            humanReadable = False,
            quiet         = False,
        )

        # Platzierung: Anfang oben (by2), Rotation -90 Grad (CW)
        # Lokales x -> Seite -y (nach unten), lokales y -> Seite +x (nach rechts)
        # translate zu (bx1, by2): obere linke Ecke des Feldes
        # => Barcode laeuft von oben (by2) nach unten (by1)
        c.saveState()
        c.translate(bx1 + 25.515, by2 - 1)
        c.rotate(-90)
        bc.drawOn(c, 0, 0)
        c.restoreState()

        # --- Barcode-Text (LBV_Barcode_Text Feld) ---
        # Feld: [0.00, 218.88, 35.21, 326.69] pt
        # rotate(-90): Kopf nach rechts neigen zum Lesen (von oben nach unten)
        tx1, ty1, tx2, ty2 = 0.0, 218.88, 35.21, 326.69
        t_cx = (tx1 + tx2) / 2.0
        t_cy = (ty1 + ty2) / 2.0
        c.saveState()
        c.translate(t_cx, t_cy)
        c.rotate(-90)
        c.setFont("Helvetica-Bold", 8.925)
        c.setFillColorRGB(0, 0, 0)
        c.drawCentredString(0, 0, barcode_val)
        c.restoreState()

        c.save()
        buf.seek(0)

        # Overlay ueber das Formular legen (over=True: barcode wird oben gerendert)
        overlay_reader = PR(buf)
        overlay_page   = overlay_reader.pages[0]
        writer.pages[0].merge_page(overlay_page, over=True)
        print(f"Barcode: {barcode_val}")

    except Exception as e:
        print(f"Barcode-Fehler (nicht fatal): {type(e).__name__}: {e}", flush=True)





def get_val(fields, name):
    if name not in fields:
        return ""
    f = fields[name]
    v = ""
    if isinstance(f, dict):
        v = f.get("/V", "")
    else:
        v = str(f)
    v = str(v).strip()
    if v.startswith("/"):
        v = v[1:]
    return v


def lese_a_werte(pdf_pfad):
    """Liest val_a und val_merkmal aus FeldA_Input1 einer Monats-PDF"""
    reader = PdfReader(pdf_pfad)
    page = reader.pages[0]
    val_a = ''
    val_m = ''
    if '/Annots' not in page:
        return val_a, val_m
    for annot in page['/Annots']:
        obj = annot.get_object()
        t = str(obj.get('/T', ''))
        p_ref = obj.get('/Parent', None)
        if not p_ref:
            continue
        pt = str(p_ref.get_object().get('/T', ''))
        if pt == 'FeldA_Input1':
            v = str(obj.get('/V', '')).strip().lstrip('/')
            if t == '1' and v:
                val_a = v
            elif t == '2' and v:
                val_m = v
    return val_a, val_m


def lese_persoenliche_felder(pdf_pfad):
    """Liest persönliche Daten aus einer Monats-PDF (ohne FeldA/FeldC).
    Bei Checkboxen/Radiobuttons (/FT = /Btn) bleibt der fuehrende '/'
    des Namenswerts (z.B. '/Ja') erhalten."""
    reader = PdfReader(pdf_pfad)
    fields = reader.get_fields() or {}
    result = {}
    for name, field in fields.items():
        if name.startswith('FeldA') or name.startswith('FeldC') or 'Füller' in name:
            continue
        if isinstance(field, dict):
            v = field.get('/V', '')
        else:
            v = str(field)
        raw = str(v).strip()

        # Checkbox / Radio: Namenswert inkl. '/' beibehalten
        if isinstance(field, dict) and str(field.get('/FT', '')) == '/Btn':
            if raw and raw not in ('/Off', 'Off'):
                result[name] = raw if raw.startswith('/') else '/' + raw
            continue

        v = raw.lstrip('/')
        if v:
            result[name] = v
    return result






def suppress_empty_widget_print(writer):
    """
    Setzt den Print-Flag leerer Widget-Annotationen auf False.
    Leere Felder (Kaestchen in Bereichen A/B/C) werden beim Drucken
    komplett unsichtbar - kein Rahmen, kein Inhalt.
    Gefuellte Felder (Name, Datum, FeldA-Werte etc.) bleiben sichtbar.
    """
    from pypdf.generic import NameObject, NumberObject

    page = writer.pages[0]
    if '/Annots' not in page:
        return

    for annot_ref in page['/Annots']:
        try:
            annot = annot_ref.get_object()
            if str(annot.get('/Subtype', '')) != '/Widget':
                continue

            # Eigenen Wert pruefen
            v = str(annot.get('/V', '')).strip().lstrip('/')

            # Eltern-Wert pruefen
            pv = ''
            parent_ref = annot.get('/Parent')
            if parent_ref:
                try:
                    pv = str(parent_ref.get_object().get('/V', '')).strip().lstrip('/')
                except Exception:
                    pass

            if not v and not pv:
                # Print-Bit (Bit 3 = Wert 4) aus /F entfernen
                try:
                    f_val = int(str(annot.get('/F', 4)))
                except Exception:
                    f_val = 4
                new_f = f_val & ~4   # Print=false
                annot[NameObject('/F')] = NumberObject(new_f)
        except Exception:
            pass





def _leere_felder_ap_loeschen(writer):
    """Leert AP-Streams leerer Widget-Felder direkt.
    Zuverlässiger als Print-Flag da es auf Bildschirm und Druck wirkt."""
    from pypdf.generic import NameObject, DecodedStreamObject

    page = writer.pages[0]
    if '/Annots' not in page:
        return

    for annot_ref in page['/Annots']:
        try:
            annot = annot_ref.get_object()
            if str(annot.get('/Subtype', '')) != '/Widget':
                continue

            v = str(annot.get('/V', '')).strip().lstrip('/')
            pv = ''
            parent_ref = annot.get('/Parent')
            if parent_ref:
                try:
                    pv = str(parent_ref.get_object().get('/V', '')).strip().lstrip('/')
                except Exception:
                    pass

            if not v and not pv:
                if '/AP' in annot:
                    try:
                        ap = annot['/AP'].get_object()
                        if '/N' in ap:
                            n = ap['/N'].get_object()
                            if hasattr(n, '_data'):
                                n._data = b''
                    except Exception:
                        pass
        except Exception:
            pass


def setze_checkbox(writer, feldname):
    """Setzt eine Checkbox direkt auf Zustand 'an' (V, AS, Print-Flag)."""
    from pypdf.generic import NameObject, NumberObject
    gefunden = False
    for annot_ref in writer.pages[0].get('/Annots', []):
        obj = annot_ref.get_object()
        if str(obj.get('/Subtype', '')) != '/Widget':
            continue
        # vollqualifizierten Feldnamen aufbauen
        teile, node = [], obj
        while node is not None:
            t = node.get('/T')
            if t is not None:
                teile.append(str(t))
            p = node.get('/Parent')
            node = p.get_object() if p else None
        name = '.'.join(reversed(teile))
        if name != feldname and not name.endswith('.' + feldname):
            continue

        # Name des "An"-Zustands aus dem AP-Dictionary lesen
        an = '/Ja'
        try:
            zustaende = list(obj['/AP'].get_object()['/N'].get_object().keys())
            print(f"Checkbox {name}: Zustaende im PDF = {zustaende}")
            kandidaten = [z for z in zustaende if str(z) != '/Off']
            if kandidaten:
                an = str(kandidaten[0])
        except Exception as e:
            print(f"Checkbox {name}: AP nicht lesbar ({e}), nehme /Ja")

        obj[NameObject('/V')] = NameObject(an)
        obj[NameObject('/AS')] = NameObject(an)
        par = obj.get('/Parent')
        if par:
            par.get_object()[NameObject('/V')] = NameObject(an)
        try:
            f_val = int(str(obj.get('/F', 4)))
        except Exception:
            f_val = 4
        obj[NameObject('/F')] = NumberObject(f_val | 4)   # Print-Flag an
        print(f"Checkbox {name} gesetzt auf {an}")
        gefunden = True
    if not gefunden:
        print(f"WARNUNG: Checkbox-Feld '{feldname}' nicht gefunden")


def merge_lbv(json_pfad):
    with open(json_pfad, encoding='utf-8') as f:
        data = json.load(f)

    vorlage = data['vorlage_pfad']
    ausgabe = data['ausgabe_pfad']
    quellen = data['quellen']

    if not quellen:
        print("FEHLER: Keine Quelldateien angegeben", file=sys.stderr)
        sys.exit(1)

    felder = lese_persoenliche_felder(quellen[0])
    pers_nr = felder.get('LBV_Personalnummer', '')
    print(f"Personalnummer: {pers_nr}")

    # Merkmal-Fallback aus erster Quelldatei
    _, merkmal_fallback = lese_a_werte(quellen[0])

    a_werte = []
    for src in quellen:
        val_a, val_m = lese_a_werte(src)
        if not val_m:
            val_m = merkmal_fallback
        dateiname = src.replace('\\', '/').split('/')[-1]
        if val_a:
            a_werte.append((val_a, val_m))
            print(f"  OK  {dateiname}: A={val_a}  Merkmal={val_m!r}")
        else:
            print(f"  WARN  Kein A-Wert in {dateiname}", file=sys.stderr)

    if not a_werte:
        print("FEHLER: Keine A-Werte in den Quelldateien gefunden.", file=sys.stderr)
        sys.exit(1)

    if len(a_werte) > MAX_ZEILEN:
        print(f"Warnung: {len(a_werte)} Monate, Template hat nur {MAX_ZEILEN} Zeilen. "
              f"Nur erste {MAX_ZEILEN} werden eingetragen.", file=sys.stderr)
        a_werte = a_werte[:MAX_ZEILEN]

    reader = PdfReader(vorlage)
    writer = PdfWriter()
    writer.append(reader)

    felder['LBV_Checkbox_Nein'] = '/Ja'

    writer.update_page_form_field_values(
        writer.pages[0], felder, auto_regenerate=False
    )

    page = writer.pages[0]
    for annot in page['/Annots']:
        obj = annot.get_object()
        t = str(obj.get('/T', ''))
        p_ref = obj.get('/Parent', None)
        if not p_ref:
            continue
        pt = str(p_ref.get_object().get('/T', ''))

        for z, (val_a, val_m) in enumerate(a_werte, start=1):
            if pt != f'FeldA_Input{z}':
                continue
            if t == '1' and val_a:
                # Nur /V setzen - AP wird unten via update_page_form_field_values generiert
                obj[NameObject('/V')] = TextStringObject(val_a)
            elif t == '2' and val_m:
                # Nur /V setzen - AP wird unten via update_page_form_field_values generiert
                obj[NameObject('/V')] = TextStringObject(val_m)

    # Schritt 1: auto_regenerate=True erstellt korrekte AP-Struktur (BBox, Resources)
    felda_update = {}
    for _z, (_va, _vm) in enumerate(a_werte, start=1):
        if _va:
            felda_update[f'FeldA_Input{_z}.1'] = _va
        if _vm:
            felda_update[f'FeldA_Input{_z}.2'] = _vm
    if felda_update:
        writer.update_page_form_field_values(
            writer.pages[0], felda_update, auto_regenerate=True
        )

    # Schritt 2: AP-Daten durch eigene Versionen ersetzen
    # t='1': eigene Zeichenabstaende (make_ap)
    # t='2': byte-identisch mit Adobe (make_ap_merkmal)
    for _annot in writer.pages[0]['/Annots']:
        _obj = _annot.get_object()
        _t = str(_obj.get('/T', ''))
        _pr = _obj.get('/Parent')
        if not _pr:
            continue
        _pt = str(_pr.get_object().get('/T', ''))
        for _z, (_va, _vm) in enumerate(a_werte, start=1):
            if _pt != f'FeldA_Input{_z}':
                continue
            if _t == '1' and _va and '/AP' in _obj:
                try:
                    _ap = _obj['/AP'].get_object()
                    if '/N' in _ap:
                        _n = _ap['/N'].get_object()
                        if hasattr(_n, '_data'):
                            _n._data = make_ap(_va).get_data()
                            if '/Filter' in _n:
                                del _n['/Filter']
                except Exception:
                    pass
            elif _t == '2' and _vm and '/AP' in _obj:
                try:
                    _ap = _obj['/AP'].get_object()
                    if '/N' in _ap:
                        _n = _ap['/N'].get_object()
                        if hasattr(_n, '_data'):
                            _n._data = make_ap_merkmal(_vm).get_data()
                            if '/Filter' in _n:
                                del _n['/Filter']
                except Exception:
                    pass

    writer._root_object['/AcroForm'][NameObject('/NeedAppearances')] = BooleanObject(False)

    suppress_empty_widget_print(writer)
    _leere_felder_ap_loeschen(writer)

    # Checkbox "Nein" ganz am Schluss direkt im PDF setzen (V, AS, Print-Flag)
    setze_checkbox(writer, 'LBV_Checkbox_Nein')

    # Barcode einfügen (benötigt reportlab)
    if pers_nr:
        add_barcode(writer, pers_nr)

    # Erst in TEMP schreiben, dann verschieben (verhindert OneDrive-Konflikte)
    import tempfile, shutil, os
    tmp_fd, tmp_pfad = tempfile.mkstemp(suffix='.pdf')
    os.close(tmp_fd)
    try:
        with open(tmp_pfad, 'wb') as f:
            writer.write(f)
        if os.path.exists(ausgabe):
            os.remove(ausgabe)
        shutil.move(tmp_pfad, ausgabe)
    except Exception:
        if os.path.exists(tmp_pfad):
            os.remove(tmp_pfad)
        raise

    print(f"OK: {len(a_werte)} Monate zusammengefasst -> {ausgabe}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Aufruf: lbv_merge.py <json_pfad>", file=sys.stderr)
        sys.exit(1)
    merge_lbv(sys.argv[1])
