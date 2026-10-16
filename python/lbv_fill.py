# -*- coding: utf-8 -*-
"""
lbv_fill.py - LBV Mehrarbeitsvergütung Formular befüllen
Liest Daten aus einer JSON-Datei und schreibt sie in die LBV-Vorlage-PDF.
Aufruf: lbv_fill.exe <json_pfad>
"""
import sys
import json
from reportlab.graphics.barcode import code39 as _rl_code39
from reportlab.pdfgen import canvas as _rl_canvas
REPORTLAB_OK = True
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    NameObject, TextStringObject, BooleanObject,
    DecodedStreamObject, ArrayObject, FloatObject,
    NumberObject, DictionaryObject
)

# Kästchen-Struktur für FeldA (von Adobe generiert, fest)
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

# Barcode-Konstanten (aus LBV-Formular LBV(A/Bes)23_2019_02)
BARCODE_PRAEFIX = "ABES23/201902/"


def make_ap(val):
    """AP-Stream für FeldA (10 Zeichen: MMJJ SSS S)"""
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




def fill_lbv(json_pfad):
    # JSON laden
    with open(json_pfad, 'r', encoding='utf-8') as f:
        data = json.load(f)

    vorlage_pfad = data['vorlage_pfad']
    ausgabe_pfad = data['ausgabe_pfad']
    felder       = data['felder']
    val_a1       = data.get('val_a1', '')
    val_merkmal  = data.get('val_merkmal', '')
    val_c1       = data.get('val_c1', '')

    # Personalnummer für Barcode
    pers_nr = felder.get('LBV_Personalnummer', '')

    # PDF öffnen
    reader = PdfReader(vorlage_pfad)
    writer = PdfWriter()
    writer.append(reader)

    # Einfache Felder befüllen
    writer.update_page_form_field_values(
        writer.pages[0], felder, auto_regenerate=False
    )

    # FeldA und FeldC per AP-Stream befüllen
    page = writer.pages[0]
    for annot in page['/Annots']:
        obj = annot.get_object()
        t   = str(obj.get('/T', ''))
        p   = obj.get('/Parent', None)
        if not p:
            continue
        pt = str(p.get_object().get('/T', ''))

        for z in range(1, 7):
            # FeldA Zeile 1: Monat/Jahr/Stunden
            if ('FeldA_Input' + str(z)) == pt and t == '1' and val_a1 and z == 1:
                obj[NameObject('/V')] = TextStringObject(val_a1)
                if '/AP' in obj and '/N' in obj['/AP']:
                    n = obj['/AP']['/N'].get_object()
                    n._data = make_ap(val_a1).get_data()
                    if '/Filter' in n:
                        del n['/Filter']

            # FeldA Zeile 1: Merkmal
            if ('FeldA_Input' + str(z)) == pt and t == '2' and val_merkmal and z == 1:
                obj[NameObject('/V')] = TextStringObject(val_merkmal)
                ap_stream = make_ap_merkmal(val_merkmal)
                # /CoBo Font-Ressource aus AcroForm /DR hinzufuegen
                try:
                    dr_fonts = writer._root_object['/AcroForm'].get_object()['/DR'].get_object().get('/Font')
                    if dr_fonts:
                        cobo = dr_fonts.get_object().get('/CoBo')
                        if cobo:
                            fd = DictionaryObject()
                            fd[NameObject('/CoBo')] = cobo
                            rd = DictionaryObject()
                            rd[NameObject('/Font')] = fd
                            ap_stream[NameObject('/Resources')] = rd
                except Exception:
                    pass
                ap_ref = writer._add_object(ap_stream)
                if '/AP' in obj:
                    obj['/AP'].get_object()[NameObject('/N')] = ap_ref
                else:
                    ap_dict = DictionaryObject()
                    ap_dict[NameObject('/N')] = ap_ref
                    obj[NameObject('/AP')] = ap_dict

            # FeldC Zeile 1: Teilzeit-Angestellte
            if val_c1 and ('FeldC_Input' + str(z)) == pt and t == '1' and z == 1:
                obj[NameObject('/V')] = TextStringObject(val_c1)

    writer._root_object['/AcroForm'][NameObject('/NeedAppearances')] = BooleanObject(False)

    # Leere Formularfelder: # Kaestchen-Raster beim Drucken unterdruecken
    suppress_empty_widget_print(writer)

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
        if os.path.exists(ausgabe_pfad):
            os.remove(ausgabe_pfad)
        shutil.move(tmp_pfad, ausgabe_pfad)
    except Exception:
        if os.path.exists(tmp_pfad):
            os.remove(tmp_pfad)
        raise

    print(f"OK: {ausgabe_pfad}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Aufruf: lbv_fill.exe <json_pfad>")
        sys.exit(1)
    fill_lbv(sys.argv[1])
