Attribute VB_Name = "Modul1"
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr


' ================================================================
'  HILFSFUNKTION: Spalte in Uebersicht per Kopfname finden
' ================================================================
Private Function SpalteUebersicht(ws3 As Worksheet, kopf As String) As Integer
    Dim c As Integer
    Dim lastCol As Integer
    lastCol = ws3.Cells(1, ws3.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(ws3.Cells(1, c).Value) = kopf Then
            SpalteUebersicht = c
            Exit Function
        End If
    Next c
    SpalteUebersicht = 0
End Function

' ================================================================
'  HILFSFUNKTION: Spalte "PDF erstellen" in Uebersicht sicherstellen
' ================================================================
Public Function SpaltePDF() As Integer
    Dim wsUe As Worksheet
    Set wsUe = ThisWorkbook.Sheets(Chr(220) & "bersicht")
    Dim col As Integer
    col = SpalteUebersicht(wsUe, "PDF erstellen")
    If col = 0 Then
        col = wsUe.Cells(1, wsUe.Columns.Count).End(xlToLeft).Column + 1
        wsUe.Cells(1, col).Value = "PDF erstellen"
        wsUe.Cells(1, col).Font.Bold = True
        Dim lastRow As Long
        lastRow = wsUe.Cells(wsUe.Rows.Count, "A").End(xlUp).Row
        Dim r As Long
        For r = 2 To lastRow
            If Trim(wsUe.Cells(r, 1).Value) <> "" Then
                wsUe.Cells(r, col).Value = 0
            End If
        Next r
        wsUe.Columns(col).AutoFit
    End If
    SpaltePDF = col
End Function

' ================================================================
'  HILFSFUNKTIONEN: Korrektur der Vertretungen
'  Spalten "Korrektur" und "Korrekturhinweis" der Uebersicht.
'  Die Zahl in "Korrektur" wird zu den Vertretungen addiert, also
'      Anzahl abrechenbare Vertretungen = Summenwert + Korrektur
'  Das gilt fuer "PDF - Nur Abrechnungsfaehige", die Mehrarbeits-
'  uebersicht (Einzel- und Sammel-PDF) und das LBV-Formular.
'  Die Spalten "zu zaehlende Vertretungen" (frueher "Vertretungen") und
'  "Summenwert" behalten den Wert aus
'  Tabelle1; die Korrektur wird jeweils aktuell dazugerechnet, d.h.
'  eine nachtraeglich geaenderte Korrektur wirkt ohne neues Einlesen.
'  Leer = 0, negative Werte (Abzug) sind erlaubt.
' ================================================================
Private Function ZelleHatInhalt(ByVal v As Variant) As Boolean
    If IsError(v) Then ZelleHatInhalt = True: Exit Function
    ZelleHatInhalt = (Trim(CStr(v)) <> "")
End Function

' Korrekturwert einer Zeile (leer, keine Zahl oder Spalte fehlt = 0)
Private Function KorrekturWert(ws3 As Worksheet, ByVal r As Long, ByVal colKorr As Integer) As Double
    KorrekturWert = 0
    If colKorr = 0 Then Exit Function
    Dim v As Variant
    v = ws3.Cells(r, colKorr).Value
    If IsError(v) Then Exit Function
    If Trim(CStr(v)) = "" Then Exit Function
    If IsNumeric(v) Then KorrekturWert = CDbl(v)
End Function

' Korrekturhinweis einer Zeile ("" wenn leer oder Spalte fehlt)
Private Function KorrekturHinweis(ws3 As Worksheet, ByVal r As Long, ByVal colHinweis As Integer) As String
    KorrekturHinweis = ""
    If colHinweis = 0 Then Exit Function
    Dim v As Variant
    v = ws3.Cells(r, colHinweis).Value
    If IsError(v) Then Exit Function
    KorrekturHinweis = Trim(CStr(v))
End Function

' Prueft, ob in Spalte "Korrektur" nur Zahlen stehen.
'   colPDF > 0: nur Lehrer mit "PDF erstellen" = 1, sonst alle Lehrer
'   False: Meldung wurde angezeigt, aufrufendes Makro bricht ab
Private Function KorrekturPruefen(ws3 As Worksheet, ByVal colKorr As Integer, _
                                  ByVal colPDF As Integer, ByVal lastRow As Long) As Boolean
    KorrekturPruefen = True
    If colKorr = 0 Then Exit Function
    Dim r As Long, v As Variant, liste As String
    liste = ""
    For r = 2 To lastRow
        If Trim(CStr(ws3.Cells(r, 1).Value)) <> "" Then
            Dim pruefen As Boolean: pruefen = True
            If colPDF > 0 Then pruefen = (ws3.Cells(r, colPDF).Value = 1)
            If pruefen Then
                v = ws3.Cells(r, colKorr).Value
                If IsError(v) Then
                    liste = liste & "  " & Trim(CStr(ws3.Cells(r, 1).Value)) & ":  (Fehlerwert)" & Chr(10)
                ElseIf Trim(CStr(v)) <> "" Then
                    If Not IsNumeric(v) Then
                        liste = liste & "  " & Trim(CStr(ws3.Cells(r, 1).Value)) & ":  " & Trim(CStr(v)) & Chr(10)
                    End If
                End If
            End If
        End If
    Next r
    If liste = "" Then Exit Function
    MsgBox "In Spalte 'Korrektur' der " & Chr(220) & "bersicht steht keine Zahl bei:" & Chr(10) & Chr(10) & _
           liste & Chr(10) & _
           "Bitte eine Zahl eintragen (z.B. 2 oder -1) oder die Zelle leeren." & Chr(10) & _
           "Es wurde nichts erstellt.", vbExclamation, "Korrektur"
    KorrekturPruefen = False
End Function

' Zahl ohne unnoetige Nachkommastellen, z.B. 3 / -2 / 1,5
Private Function ZahlText(ByVal x As Double) As String
    If x = Fix(x) Then
        ZahlText = CStr(CLng(x))
    Else
        ZahlText = Format(x, "0.0#")
    End If
End Function

' ================================================================
'  HILFSMAKRO 1: Alle Lehrer auf PDF=1 setzen
' ================================================================
Sub PDFAlleMarkieren()
    Dim wsUe As Worksheet
    Set wsUe = ThisWorkbook.Sheets(Chr(220) & "bersicht")
    Dim col As Integer
    col = SpaltePDF()
    Dim lastRow As Long
    lastRow = wsUe.Cells(wsUe.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsUe.Cells(r, 1).Value) <> "" Then
            wsUe.Cells(r, col).Value = 1
        End If
    Next r
    Call PDFSpalteFormatieren
End Sub

' ================================================================
'  HILFSMAKRO 2: Nur Abrechnungsfaehige auf PDF=1 setzen
' ================================================================
Sub PDFNurAbrechnungsfaehige()
    Dim wsUe As Worksheet
    Set wsUe = ThisWorkbook.Sheets(Chr(220) & "bersicht")
    Dim colPDF As Integer
    Dim colSumme As Integer
    Dim colBagatell As Integer
    colPDF = SpaltePDF()
    colSumme = SpalteUebersicht(wsUe, "Summenwert")
    colBagatell = SpalteUebersicht(wsUe, "Bagatellgrenze")
    Dim colPersNr As Integer
    colPersNr = SpalteUebersicht(wsUe, "Pers.Nr")
    Dim colKorr As Integer
    colKorr = SpalteUebersicht(wsUe, "Korrektur")

    Dim lastRow As Long
    lastRow = wsUe.Cells(wsUe.Rows.Count, "A").End(xlUp).Row
    If Not KorrekturPruefen(wsUe, colKorr, 0, lastRow) Then Exit Sub
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsUe.Cells(r, 1).Value) <> "" Then
            Dim summe As Double
            Dim bagatell As Double
            summe = 0: bagatell = 0
            If colSumme > 0 And IsNumeric(wsUe.Cells(r, colSumme).Value) Then
                summe = CDbl(wsUe.Cells(r, colSumme).Value)
            End If
            ' Korrektur zu den Vertretungen addieren
            summe = summe + KorrekturWert(wsUe, r, colKorr)
            If colBagatell > 0 And IsNumeric(wsUe.Cells(r, colBagatell).Value) Then
                bagatell = CDbl(wsUe.Cells(r, colBagatell).Value)
            End If
            wsUe.Cells(r, colPDF).Value = IIf(bagatell > 0 And summe > bagatell, 1, 0)
        End If
    Next r
    Call PDFSpalteFormatieren
End Sub

' ================================================================
'  HILFSMAKRO 3: Auswahl leeren
' ================================================================
Sub PDFAuswahlLeeren()
    Dim wsUe As Worksheet
    Set wsUe = ThisWorkbook.Sheets(Chr(220) & "bersicht")
    Dim col As Integer
    col = SpaltePDF()
    Dim lastRow As Long
    lastRow = wsUe.Cells(wsUe.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsUe.Cells(r, 1).Value) <> "" Then
            wsUe.Cells(r, col).Value = 0
        End If
    Next r
    Call PDFSpalteFormatieren
End Sub

' ================================================================
'  Formatierung der PDF-Spalte (gruen=1, weiss=0)
' ================================================================
Public Sub PDFSpalteFormatieren()
    Dim wsUe As Worksheet
    Set wsUe = ThisWorkbook.Sheets(Chr(220) & "bersicht")
    Dim col As Integer
    col = SpalteUebersicht(wsUe, "PDF erstellen")
    If col = 0 Then Exit Sub
    Dim lastRow As Long
    lastRow = wsUe.Cells(wsUe.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsUe.Cells(r, 1).Value) <> "" Then
            With wsUe.Cells(r, col)
                If .Value = 1 Then
                    .Interior.Color = RGB(180, 230, 180)
                    .Font.Bold = True
                Else
                    .Interior.ColorIndex = xlNone
                    .Font.Bold = False
                    .Value = 0
                End If
            End With
        End If
    Next r
End Sub

' ================================================================
'  MAKRO 1: Kuerzel ersetzen
' ================================================================
Sub NamenskuerzelErsetzenUndZellenLoeschen()
    Dim wsData As Worksheet, wsMapping As Worksheet
    Dim lastRowData As Long, lastRowMapping As Long
    Dim i As Long, j As Long
    Dim kuerzel As String, pseudonym As String

    Set wsData = ThisWorkbook.Sheets("Tabelle1")
    Set wsMapping = ThisWorkbook.Sheets("Pseudonyme")
    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastRowMapping = wsMapping.Cells(wsMapping.Rows.Count, "A").End(xlUp).Row

    Dim colName As Integer, colPseudonym As Integer
    colName = 0: colPseudonym = 0
    Dim lastCol As Integer
    lastCol = wsMapping.Cells(1, wsMapping.Columns.Count).End(xlToLeft).Column
    Dim k As Integer
    For k = 1 To lastCol
        Select Case LCase(Trim(wsMapping.Cells(1, k).Value))
            Case "name":      colName = k
            Case "pseudonym": colPseudonym = k
        End Select
    Next k
    If colName = 0 Or colPseudonym = 0 Then
        MsgBox "Spalten 'Name'/'Pseudonym' nicht gefunden!", vbCritical: Exit Sub
    End If

    Dim ersetzungen As Long: ersetzungen = 0
    For i = 1 To lastRowData
        kuerzel = Trim(wsData.Cells(i, "A").Value)
        If kuerzel <> "" Then
            For j = 2 To lastRowMapping
                If LCase(kuerzel) = LCase(Trim(wsMapping.Cells(j, colName).Value)) Then
                    pseudonym = Trim(wsMapping.Cells(j, colPseudonym).Value)
                    wsData.Cells(i, "A").Value = pseudonym
                    Dim lastColRow As Integer
                    lastColRow = wsData.Cells(i, wsData.Columns.Count).End(xlToLeft).Column
                    If lastColRow > 1 Then
                        wsData.Range(wsData.Cells(i, 2), wsData.Cells(i, lastColRow)).ClearContents
                    End If
                    ersetzungen = ersetzungen + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    MsgBox "Fertig! " & ersetzungen & " K" & Chr(252) & "rzel ersetzt.", vbInformation
End Sub

' ================================================================
'  MAKRO 2: Kuerzel wiederherstellen
' ================================================================
Sub NamenskuerzelWiederherstellen()
    Dim wsData As Worksheet, wsMapping As Worksheet
    Dim lastRowData As Long, lastRowMapping As Long
    Dim i As Long, j As Long, zellwert As String

    Set wsData = ThisWorkbook.Sheets("Tabelle1")
    Set wsMapping = ThisWorkbook.Sheets("Pseudonyme")
    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastRowMapping = wsMapping.Cells(wsMapping.Rows.Count, "A").End(xlUp).Row

    Dim colName As Integer, colPseudonym As Integer
    colName = 0: colPseudonym = 0
    Dim lastCol As Integer
    lastCol = wsMapping.Cells(1, wsMapping.Columns.Count).End(xlToLeft).Column
    Dim k As Integer
    For k = 1 To lastCol
        Select Case LCase(Trim(wsMapping.Cells(1, k).Value))
            Case "name":      colName = k
            Case "pseudonym": colPseudonym = k
        End Select
    Next k
    If colName = 0 Or colPseudonym = 0 Then
        MsgBox "Spalten nicht gefunden!", vbCritical: Exit Sub
    End If

    Dim wiederherstellungen As Long: wiederherstellungen = 0
    For i = 1 To lastRowData
        zellwert = Trim(wsData.Cells(i, "A").Value)
        If zellwert <> "" Then
            For j = 2 To lastRowMapping
                If LCase(zellwert) = LCase(Trim(wsMapping.Cells(j, colPseudonym).Value)) Then
                    wsData.Cells(i, "A").Value = Trim(wsMapping.Cells(j, colName).Value)
                    wiederherstellungen = wiederherstellungen + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    MsgBox "Fertig! " & wiederherstellungen & " K" & Chr(252) & "rzel wiederhergestellt.", vbInformation
End Sub

' ================================================================
'  MAKRO 3: Summen in Tabelle Uebersicht
' ================================================================
Sub SummenInTabelleUebersicht()
    Dim ws1 As Worksheet, ws3 As Worksheet
    Dim lastRow1 As Long, lastRow3 As Long
    Dim i As Long, zellwert As String
    Dim aktuellesKuerzel As String
    Dim summenwert As Double
    Dim entfaelleZaehlend As Long, entfaelleNichtZaehlen As Long
    Dim vertretungen As Long, vertretungenNichtZaehlen As Long

    Dim prfxE  As String: prfxE = "Entf" & Chr(228) & "lle: "
    Dim prfxEN As String: prfxEN = "Entfall nicht zu z" & Chr(228) & "hlen: "
    Dim prfxV  As String: prfxV = "Vertretungen: "
    Dim prfxVN As String: prfxVN = "Vertretung nicht zu z" & Chr(228) & "hlen: "

    Set ws1 = ThisWorkbook.Sheets("Tabelle1")
    Dim wsExists As Boolean: wsExists = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = Chr(220) & "bersicht" Then wsExists = True: Exit For
    Next ws
    If Not wsExists Then MsgBox "'" & Chr(220) & "bersicht' nicht gefunden!", vbCritical: Exit Sub

    Set ws3 = ThisWorkbook.Sheets(Chr(220) & "bersicht")
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

    ' Korrektur und Korrekturhinweis werden von Hand eingetragen und beim
    ' Einlesen nicht ueberschrieben. Stehen dort noch Eintraege (z.B. aus dem
    ' Vormonat), nachfragen - sonst wuerden sie erneut abgerechnet.
    Dim colKorr As Integer:  colKorr = SpalteUebersicht(ws3, "Korrektur")
    Dim colKorrH As Integer: colKorrH = SpalteUebersicht(ws3, "Korrekturhinweis")
    If (colKorr > 0 Or colKorrH > 0) And lastRow3 > 1 Then
        Dim anzKorr As Long: anzKorr = 0
        Dim rk As Long
        For rk = 2 To lastRow3
            Dim hatKorr As Boolean: hatKorr = False
            If colKorr > 0 Then hatKorr = ZelleHatInhalt(ws3.Cells(rk, colKorr).Value)
            If colKorrH > 0 And Not hatKorr Then hatKorr = ZelleHatInhalt(ws3.Cells(rk, colKorrH).Value)
            If hatKorr Then anzKorr = anzKorr + 1
        Next rk
        If anzKorr > 0 Then
            Dim antwortK As VbMsgBoxResult
            antwortK = MsgBox("In den Spalten 'Korrektur' / 'Korrekturhinweis' stehen noch Eintr" & Chr(228) & "ge bei " & _
                anzKorr & " Lehrer(n)." & Chr(10) & Chr(10) & _
                "Die Korrektur wird zu den Vertretungen addiert. Stammen die Eintr" & Chr(228) & "ge aus einem " & _
                "fr" & Chr(252) & "heren Monat, w" & Chr(252) & "rden sie noch einmal abgerechnet." & Chr(10) & Chr(10) & _
                "[Ja]" & Chr(9) & "Korrekturen behalten (gelten f" & Chr(252) & "r den eingelesenen Monat)" & Chr(10) & _
                "[Nein]" & Chr(9) & "Korrekturen und Hinweise l" & Chr(246) & "schen" & Chr(10) & _
                "[Abbrechen]" & Chr(9) & "nichts einlesen", _
                vbYesNoCancel + vbQuestion, "Summen einlesen")
            If antwortK = vbCancel Then Exit Sub
            If antwortK = vbNo Then
                If colKorr > 0 Then ws3.Range(ws3.Cells(2, colKorr), ws3.Cells(lastRow3, colKorr)).ClearContents
                If colKorrH > 0 Then ws3.Range(ws3.Cells(2, colKorrH), ws3.Cells(lastRow3, colKorrH)).ClearContents
            End If
        End If
    End If

    Dim colSumme As Integer, colEntfG As Integer, colEntfN As Integer
    Dim colVert As Integer, colVertN As Integer
    colSumme = 0: colEntfG = 0: colEntfN = 0: colVert = 0: colVertN = 0
    Dim lastColUe As Integer
    lastColUe = ws3.Cells(1, ws3.Columns.Count).End(xlToLeft).Column
    Dim c As Integer
    For c = 1 To lastColUe
        Select Case Trim(ws3.Cells(1, c).Value)
            Case "Summenwert":                                                 colSumme = c
            Case "Entf" & Chr(228) & "lle gesamt":                            colEntfG = c
            Case "nicht zu z" & Chr(228) & "hlende Entf" & Chr(228) & "lle": colEntfN = c
            Case "zu z" & Chr(228) & "hlende Vertretungen", "Vertretungen":    If colVert = 0 Then colVert = c
            Case "nicht zu z" & Chr(228) & "hlende Vertretungen":             colVertN = c
        End Select
    Next c

    If colSumme = 0 Then lastColUe = lastColUe + 1: colSumme = lastColUe: ws3.Cells(1, colSumme).Value = "Summenwert"
    If colEntfG = 0 Then lastColUe = lastColUe + 1: colEntfG = lastColUe: ws3.Cells(1, colEntfG).Value = "Entf" & Chr(228) & "lle gesamt"
    If colEntfN = 0 Then lastColUe = lastColUe + 1: colEntfN = lastColUe: ws3.Cells(1, colEntfN).Value = "nicht zu z" & Chr(228) & "hlende Entf" & Chr(228) & "lle"
    If colVert = 0 Then lastColUe = lastColUe + 1: colVert = lastColUe: ws3.Cells(1, colVert).Value = "zu z" & Chr(228) & "hlende Vertretungen"
    If colVertN = 0 Then lastColUe = lastColUe + 1: colVertN = lastColUe: ws3.Cells(1, colVertN).Value = "nicht zu z" & Chr(228) & "hlende Vertretungen"

    ws3.Rows(1).Font.Bold = True
    If lastRow3 > 1 Then
        ws3.Range(ws3.Cells(2, colSumme), ws3.Cells(lastRow3, colSumme)).ClearContents
        ws3.Range(ws3.Cells(2, colEntfG), ws3.Cells(lastRow3, colEntfG)).ClearContents
        ws3.Range(ws3.Cells(2, colEntfN), ws3.Cells(lastRow3, colEntfN)).ClearContents
        ws3.Range(ws3.Cells(2, colVert), ws3.Cells(lastRow3, colVert)).ClearContents
        ws3.Range(ws3.Cells(2, colVertN), ws3.Cells(lastRow3, colVertN)).ClearContents
    End If

    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    aktuellesKuerzel = ""
    summenwert = 0: entfaelleZaehlend = 0: entfaelleNichtZaehlen = 0
    vertretungen = 0: vertretungenNichtZaehlen = 0
    Dim nichtGefunden As String: nichtGefunden = ""
    Dim rawVal As String, sp As Integer, j As Long

    For i = 1 To lastRow1 - 1
        zellwert = Trim(ws1.Cells(i, 1).Value)
        If zellwert = "" Then GoTo Weiter
        Dim naechste As String: naechste = ""
        For j = i + 1 To Application.Min(i + 5, lastRow1)
            naechste = Trim(ws1.Cells(j, 1).Value)
            If naechste <> "" Then Exit For
        Next j
        If InStr(naechste, "Vertretungen / Entf") > 0 Then
            If aktuellesKuerzel <> "" Then
                Call SchreibeInUebersicht(ws3, lastRow3, aktuellesKuerzel, _
                    colSumme, colEntfG, colEntfN, colVert, colVertN, _
                    summenwert, entfaelleZaehlend, entfaelleNichtZaehlen, _
                    vertretungen, vertretungenNichtZaehlen, nichtGefunden)
            End If
            aktuellesKuerzel = zellwert
            summenwert = 0: entfaelleZaehlend = 0: entfaelleNichtZaehlen = 0
            vertretungen = 0: vertretungenNichtZaehlen = 0
            GoTo Weiter
        End If
        If Left(zellwert, 6) = "Summe:" Then
            rawVal = Trim(Mid(zellwert, 7)): sp = InStr(rawVal, " "): If sp > 0 Then rawVal = Left(rawVal, sp - 1)
            If IsNumeric(rawVal) Then summenwert = CDbl(rawVal)
            GoTo Weiter
        End If
        If Left(zellwert, Len(prfxE)) = prfxE Then
            rawVal = Trim(Mid(zellwert, Len(prfxE) + 1)): sp = InStr(rawVal, " "): If sp > 0 Then rawVal = Left(rawVal, sp - 1)
            If IsNumeric(rawVal) Then entfaelleZaehlend = entfaelleZaehlend + CLng(rawVal)
            GoTo Weiter
        End If
        If Left(zellwert, Len(prfxEN)) = prfxEN Then
            rawVal = Trim(Mid(zellwert, Len(prfxEN) + 1)): sp = InStr(rawVal, " "): If sp > 0 Then rawVal = Left(rawVal, sp - 1)
            If IsNumeric(rawVal) Then entfaelleNichtZaehlen = entfaelleNichtZaehlen + CLng(rawVal)
            GoTo Weiter
        End If
        If Left(zellwert, Len(prfxV)) = prfxV Then
            rawVal = Trim(Mid(zellwert, Len(prfxV) + 1)): sp = InStr(rawVal, " "): If sp > 0 Then rawVal = Left(rawVal, sp - 1)
            If IsNumeric(rawVal) Then vertretungen = CLng(rawVal)
            GoTo Weiter
        End If
        If Left(zellwert, Len(prfxVN)) = prfxVN Then
            rawVal = Trim(Mid(zellwert, Len(prfxVN) + 1)): sp = InStr(rawVal, " "): If sp > 0 Then rawVal = Left(rawVal, sp - 1)
            If IsNumeric(rawVal) Then vertretungenNichtZaehlen = CLng(rawVal)
            GoTo Weiter
        End If
Weiter:
    Next i

    If aktuellesKuerzel <> "" Then
        Call SchreibeInUebersicht(ws3, lastRow3, aktuellesKuerzel, _
            colSumme, colEntfG, colEntfN, colVert, colVertN, _
            summenwert, entfaelleZaehlend, entfaelleNichtZaehlen, _
            vertretungen, vertretungenNichtZaehlen, nichtGefunden)
    End If

    ws3.Range(ws3.Columns(1), ws3.Columns(lastColUe)).AutoFit
    If nichtGefunden = "" Then
        MsgBox "Fertig! Alle K" & Chr(252) & "rzel eingetragen.", vbInformation
    Else
        MsgBox "Nicht in " & Chr(220) & "bersicht gefunden:" & Chr(10) & nichtGefunden, vbExclamation
    End If
End Sub

Private Sub SchreibeInUebersicht(ws3 As Worksheet, lastRow3 As Long, _
    kuerzel As String, colSumme As Integer, colEntfG As Integer, colEntfN As Integer, _
    colVert As Integer, colVertN As Integer, summenwert As Double, _
    entfaelleZaehlend As Long, entfaelleNichtZaehlen As Long, _
    vertretungen As Long, vertretungenNichtZaehlen As Long, ByRef nichtGefunden As String)
    Dim r As Long
    For r = 2 To lastRow3
        If LCase(Trim(ws3.Cells(r, 1).Value)) = LCase(kuerzel) Then
            ws3.Cells(r, colSumme).Value = summenwert
            ws3.Cells(r, colEntfG).Value = entfaelleZaehlend + entfaelleNichtZaehlen
            ws3.Cells(r, colEntfN).Value = entfaelleNichtZaehlen
            ws3.Cells(r, colVert).Value = vertretungen
            ws3.Cells(r, colVertN).Value = vertretungenNichtZaehlen
            Exit Sub
        End If
    Next r
    nichtGefunden = nichtGefunden & IIf(nichtGefunden = "", "", Chr(10)) & kuerzel
End Sub

' ================================================================
'  MAKRO 4: PDF-Erstellung
' ================================================================
Sub PDFErstellen()
    Dim ws1  As Worksheet
    Dim wsUe As Worksheet
    Dim wsInfo As Worksheet

    Set ws1 = ThisWorkbook.Sheets("Tabelle1")
    Set wsUe = ThisWorkbook.Sheets(Chr(220) & "bersicht")

    Dim colPDF As Integer
    colPDF = SpaltePDF()

    Dim jahreszahl As String: jahreszahl = ""
    Dim wsI As Worksheet
    For Each wsI In ThisWorkbook.Sheets
        If LCase(wsI.Name) = "informationen" Then Set wsInfo = wsI: Exit For
    Next wsI
    If Not wsInfo Is Nothing Then
        Dim ri As Long
        For ri = 1 To wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
            If LCase(Trim(wsInfo.Cells(ri, 1).Value)) = "jahreszahl" Then
                jahreszahl = Trim(wsInfo.Cells(ri, 2).Value): Exit For
            End If
        Next ri
    End If

    Dim colSumme As Integer, colEntfG As Integer, colEntfN As Integer
    Dim colVert  As Integer, colVertN As Integer, colBagatell As Integer
    colSumme = SpalteUebersicht(wsUe, "Summenwert")
    colEntfG = SpalteUebersicht(wsUe, "Entf" & Chr(228) & "lle gesamt")
    colEntfN = SpalteUebersicht(wsUe, "nicht zu z" & Chr(228) & "hlende Entf" & Chr(228) & "lle")
    ' Spalte heisst jetzt "zu zaehlende Vertretungen" (alter Name wird weiter erkannt)
    colVert = SpalteUebersicht(wsUe, "zu z" & Chr(228) & "hlende Vertretungen")
    If colVert = 0 Then colVert = SpalteUebersicht(wsUe, "Vertretungen")
    colVertN = SpalteUebersicht(wsUe, "nicht zu z" & Chr(228) & "hlende Vertretungen")
    colBagatell = SpalteUebersicht(wsUe, "Bagatellgrenze")
    Dim colPersNr As Integer
    colPersNr = SpalteUebersicht(wsUe, "Pers.Nr")
    Dim colKorr As Integer, colKorrH As Integer
    colKorr = SpalteUebersicht(wsUe, "Korrektur")
    colKorrH = SpalteUebersicht(wsUe, "Korrekturhinweis")

    Dim lastRowUe As Long
    lastRowUe = wsUe.Cells(wsUe.Rows.Count, "A").End(xlUp).Row

    Dim anzahlMarkiert As Integer: anzahlMarkiert = 0
    Dim r As Long
    For r = 2 To lastRowUe
        If wsUe.Cells(r, colPDF).Value = 1 Then anzahlMarkiert = anzahlMarkiert + 1
    Next r
    If anzahlMarkiert = 0 Then
        MsgBox "Keine Lehrer markiert.", vbExclamation: Exit Sub
    End If
    If Not KorrekturPruefen(wsUe, colKorr, colPDF, lastRowUe) Then Exit Sub

    Dim speicherOrdner As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Speicherordner f" & Chr(252) & "r PDFs w" & Chr(228) & "hlen"
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Kein Ordner gew" & Chr(228) & "hlt. Abbruch.", vbInformation: Exit Sub
        End If
        speicherOrdner = .SelectedItems(1)
    End With
    If Right(speicherOrdner, 1) <> "\" Then speicherOrdner = speicherOrdner & "\"

    Dim lastRow1 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    Dim pdfErstellt As Integer: pdfErstellt = 0
    Dim sammelBlaetter() As String
    ReDim sammelBlaetter(0)

    For r = 2 To lastRowUe
        If wsUe.Cells(r, colPDF).Value <> 1 Then GoTo NaechsterLehrer

        Dim kuerzel As String: kuerzel = Trim(wsUe.Cells(r, 1).Value)
        If kuerzel = "" Then GoTo NaechsterLehrer

        Dim vollname As String:    vollname = ""
        Dim zeitbereich As String: zeitbereich = ""
        Dim detailZeilen() As String
        Dim anzDetail As Integer:  anzDetail = 0
        ReDim detailZeilen(0)

        Dim inBlock As Boolean: inBlock = False
        Dim i As Long
        For i = 1 To lastRow1
            Dim zv As String: zv = Trim(ws1.Cells(i, 1).Value)
            If Not inBlock Then
                If LCase(zv) = LCase(kuerzel) Then
                    vollname = Trim(ws1.Cells(i, 2).Value)
                    inBlock = True
                End If
            Else
                If InStr(zv, "Vertretungen / Entf") > 0 Then
                    Dim dp As Integer: dp = InStr(zv, ":")
                    If dp > 0 Then zeitbereich = Trim(Mid(zv, dp + 1))
                ElseIf Left(zv, 10) = "Vertretung" And InStr(zv, "/") = 0 Then
                    ' Alle Vertretungs-Zusammenfassungszeilen ignorieren
                ElseIf Left(zv, 4) = "Entf" Or Left(zv, 5) = "Summe" Or _
                       Left(zv, 12) = "Bereitschaft" Or Left(zv, 13) = "Pausenaufsich" Or _
                       Left(zv, 3) = "z" & Chr(228) & "h" Or zv = "zu" Or zv = "nicht" Or _
                       Left(zv, 5) = "z" & Chr(228) & "hle" Then
                    ' Zusammenfassungszeilen und Fortsetzungsfragmente ignorieren
                ElseIf zv = "Datum" Then
                ElseIf zv <> "" Then
                    Dim zvN As String: zvN = ""
                    Dim ii As Long
                    For ii = i + 1 To Application.Min(i + 5, lastRow1)
                        zvN = Trim(ws1.Cells(ii, 1).Value)
                        If zvN <> "" Then Exit For
                    Next ii
                    If InStr(zvN, "Vertretungen / Entf") > 0 Then
                        Exit For
                    Else
                        Dim zeile As String: zeile = ""
                        Dim s As Integer
                        For s = 1 To 9
                            zeile = zeile & IIf(s > 1, Chr(9), "") & Trim(ws1.Cells(i, s).Value)
                        Next s
                        ReDim Preserve detailZeilen(anzDetail)
                        detailZeilen(anzDetail) = zeile
                        anzDetail = anzDetail + 1
                    End If
                End If
            End If
        Next i

        Dim summeWert As Double, entfGesamt As Long, entfNicht As Long
        Dim vertWert As Long, vertNicht As Long, bagatell As Double
        Dim persNr As String: persNr = ""
        If colSumme > 0 And IsNumeric(wsUe.Cells(r, colSumme).Value) Then summeWert = CDbl(wsUe.Cells(r, colSumme).Value)
        If colBagatell > 0 And IsNumeric(wsUe.Cells(r, colBagatell).Value) Then bagatell = CDbl(wsUe.Cells(r, colBagatell).Value)
        If colPersNr > 0 And Trim(wsUe.Cells(r, colPersNr).Value) <> "" Then persNr = Trim(wsUe.Cells(r, colPersNr).Value)
        If colEntfG > 0 And IsNumeric(wsUe.Cells(r, colEntfG).Value) Then entfGesamt = CLng(wsUe.Cells(r, colEntfG).Value)
        If colEntfN > 0 And IsNumeric(wsUe.Cells(r, colEntfN).Value) Then entfNicht = CLng(wsUe.Cells(r, colEntfN).Value)
        If colVert > 0 And IsNumeric(wsUe.Cells(r, colVert).Value) Then vertWert = CLng(wsUe.Cells(r, colVert).Value)
        If colVertN > 0 And IsNumeric(wsUe.Cells(r, colVertN).Value) Then vertNicht = CLng(wsUe.Cells(r, colVertN).Value)
        Dim korrWert As Double:    korrWert = KorrekturWert(wsUe, r, colKorr)
        Dim korrHinweis As String: korrHinweis = KorrekturHinweis(wsUe, r, colKorrH)

        Call ErstellePDF(speicherOrdner, kuerzel, vollname, persNr, zeitbereich, jahreszahl, _
            detailZeilen, anzDetail, summeWert, entfGesamt, entfNicht, vertWert, vertNicht, bagatell, _
            korrWert, korrHinweis)
        pdfErstellt = pdfErstellt + 1

        ' Blatt fuer Sammel-PDF merken (wird nach der Schleife zusammengefuehrt)
        ReDim Preserve sammelBlaetter(pdfErstellt - 1)
        sammelBlaetter(pdfErstellt - 1) = kuerzel

NaechsterLehrer:
    Next r

    ' Sammel-PDF aus allen erstellten Einzelblaettern erzeugen
    If pdfErstellt > 1 Then
        Call ErstelleSammelPDF(speicherOrdner, sammelBlaetter, pdfErstellt, jahreszahl, zeitbereich, _
            ws1, wsUe, colSumme, colBagatell, colEntfG, colEntfN, colVert, colVertN, colPersNr, lastRowUe, lastRow1, _
            colKorr, colKorrH)
    End If

    MsgBox pdfErstellt & " PDF(s) erstellt in:" & Chr(10) & speicherOrdner, vbInformation
End Sub

' ================================================================
'  PDF eines Lehrers erstellen
'  summeWert = Summenwert aus der Uebersicht (ohne Korrektur),
'  korrWert/korrHinweis = Spalten Korrektur/Korrekturhinweis
' ================================================================
Private Sub ErstellePDF(ordner As String, kuerzel As String, vollname As String, _
    persNr As String, zeitbereich As String, jahreszahl As String, detailZeilen() As String, anzDetail As Integer, _
    summeWert As Double, entfGesamt As Long, entfNicht As Long, _
    vertWert As Long, vertNicht As Long, bagatell As Double, _
    korrWert As Double, korrHinweis As String)

    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "_PDF_TMP_" Then ws.Delete: Exit For
    Next ws
    Application.DisplayAlerts = True

    Dim wsTmp As Worksheet
    Set wsTmp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTmp.Name = "_PDF_TMP_"
    wsTmp.Cells.Font.Name = "Arial"
    wsTmp.Cells.Font.Size = 10

    Dim z As Long: z = 1

    ' Titel
    With wsTmp.Cells(z, 1)
        .Value = "Mehrarbeits" & Chr(252) & "bersicht " & zeitbereich & " " & jahreszahl
        .Font.Bold = True: .Font.Size = 13
    End With
    z = z + 2

    ' Name
    Dim nameZeile As String
    nameZeile = kuerzel & IIf(vollname <> "", "   " & vollname, "")
    If persNr <> "" Then nameZeile = nameZeile & "   (Pers.Nr: " & persNr & ")"
    With wsTmp.Cells(z, 1)
        .Value = nameZeile
        .Font.Size = 20
        .Font.Bold = True
    End With
    z = z + 2

    ' Tabellenkopf (ab Zeile z, Tabelle beginnt in Spalte B=2)
    ' Spalte A bleibt leer um Titeltext nicht zu stoeren
    Dim tabelleStartSpalte As Integer: tabelleStartSpalte = 1
    Dim hdrs(8) As String
    hdrs(0) = "Datum": hdrs(1) = "Stunde": hdrs(2) = "Art": hdrs(3) = "Wert"
    hdrs(4) = Chr(90) & Chr(228) & "hler": hdrs(5) = "Fach"
    hdrs(6) = "Klasse(n)": hdrs(7) = "Grund": hdrs(8) = "Text"
    Dim h As Integer
    For h = 0 To 8
        With wsTmp.Cells(z, tabelleStartSpalte + h)
            .Value = hdrs(h): .Font.Bold = True
            .Interior.Color = RGB(180, 200, 230)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
    Next h
    z = z + 1

    ' Detailzeilen
    Dim d As Integer
    For d = 0 To anzDetail - 1
        Dim felder() As String
        felder = Split(detailZeilen(d), Chr(9))
        Dim f As Integer
        For f = 0 To Application.Min(UBound(felder), 8)
            wsTmp.Cells(z, tabelleStartSpalte + f).Value = Trim(felder(f))
        Next f
        ' Zeilenumbruch fuer alle Zellen der Detailzeile aktivieren
        wsTmp.Range(wsTmp.Cells(z, tabelleStartSpalte), wsTmp.Cells(z, tabelleStartSpalte + 8)).WrapText = True
        If d Mod 2 = 1 Then
            wsTmp.Range(wsTmp.Cells(z, tabelleStartSpalte), wsTmp.Cells(z, tabelleStartSpalte + 8)).Interior.Color = RGB(242, 242, 242)
        End If
        z = z + 1
    Next d

    ' Trennlinie
    wsTmp.Range(wsTmp.Cells(z - 1, tabelleStartSpalte), wsTmp.Cells(z - 1, tabelleStartSpalte + 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    z = z + 1

    ' Kennzahlen inkl. Korrektur
    ' ("Zur Kenntnis genommen" und Unterschriftsfelder entfallen)
    Dim hinweisZeile As Long
    z = SchreibeKennzahlen(wsTmp, z, bagatell, entfGesamt, entfNicht, vertWert, vertNicht, _
                           summeWert, korrWert, korrHinweis, hinweisZeile)

    ' --- Spaltenbreiten ---
    ' AutoFit fuer alle Spalten
    wsTmp.Columns("A:I").AutoFit

    ' Spalten 1 (Datum), 2 (Stunde), 5 (Zaehler) auf Datenwert-Breite begrenzen.
    ' Nur Zeilen der Detailtabelle auswerten - keine Titel-/Kennzahlen-Zeilen.
    Dim maxBreiteDatum   As Double: maxBreiteDatum = 0
    Dim maxBreiteStunde  As Double: maxBreiteStunde = 0
    Dim maxBreiteZaehler As Double: maxBreiteZaehler = 0
    Dim zRow As Long
    Dim tabelleStartZeile As Long: tabelleStartZeile = 6 ' erste Datenzeile nach Kopf

    ' Ausschluss-Prefixe fuer Nicht-Datenwerte in Spalte 1
    Dim skip(8) As String
    skip(0) = "Mehr": skip(1) = "Von": skip(2) = "Bag": skip(3) = "Entf"
    skip(4) = "Vert": skip(5) = "Anz":  skip(6) = "Korr": skip(7) = "erf"
    skip(8) = "zu z"

    For zRow = tabelleStartZeile To z
        Dim datumVal   As String: datumVal = CStr(wsTmp.Cells(zRow, 1).Value)
        Dim stundeVal  As String: stundeVal = CStr(wsTmp.Cells(zRow, 2).Value)
        Dim zaehlerVal As String: zaehlerVal = CStr(wsTmp.Cells(zRow, 5).Value)

        ' Pruefen ob Zeile eine echte Detailzeile ist (kein Titel/Kennzahl)
        Dim istDetail As Boolean: istDetail = True
        Dim sk As Integer
        For sk = 0 To 8
            If Left(datumVal, Len(skip(sk))) = skip(sk) Then
                istDetail = False: Exit For
            End If
        Next sk

        If istDetail And datumVal <> "" Then
            Dim brDatum As Double: brDatum = Len(datumVal) * 5.5 + 4
            If brDatum > maxBreiteDatum Then maxBreiteDatum = brDatum
        End If

        If stundeVal <> "" And Left(stundeVal, 3) <> "___" Then
            Dim brStunde As Double: brStunde = Len(stundeVal) * 5.5 + 4
            If brStunde > maxBreiteStunde Then maxBreiteStunde = brStunde
        End If

        If zaehlerVal <> "" And Left(zaehlerVal, 3) <> "___" Then
            Dim brZaehler As Double: brZaehler = Len(zaehlerVal) * 5.5 + 4
            If brZaehler > maxBreiteZaehler Then maxBreiteZaehler = brZaehler
        End If
    Next zRow

    ' Mindestbreite je Kopftext
    If maxBreiteDatum < 28 Then maxBreiteDatum = 28       ' "Datum"
    If maxBreiteStunde < 28 Then maxBreiteStunde = 28     ' "Stunde"
    If maxBreiteZaehler < 28 Then maxBreiteZaehler = 28   ' "Zaehler"

    ' Obergrenze damit kein Langtext die Spalte aufblaeht
    If maxBreiteDatum > 45 Then maxBreiteDatum = 45
    If maxBreiteStunde > 50 Then maxBreiteStunde = 50
    If maxBreiteZaehler > 45 Then maxBreiteZaehler = 45

    wsTmp.Columns(1).ColumnWidth = maxBreiteDatum / 6
    wsTmp.Columns(2).ColumnWidth = maxBreiteStunde / 6
    wsTmp.Columns(5).ColumnWidth = maxBreiteZaehler / 6

    ' Zeilenhoehe des Korrekturhinweises (verbundene Zelle) erst jetzt,
    ' wenn die Spaltenbreiten feststehen
    HinweisZeileAnpassen wsTmp, hinweisZeile

    ' Seiteneinrichtung: Hochformat mit automatischem Zeilenumbruch
    With wsTmp.PageSetup
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PaperSize = xlPaperA4
        .Zoom = False
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
    End With

    Dim dateiname As String
    ' Monat aus Zeitbereich extrahieren z.B. '1.9. - 30.9.' -> '09'
    Dim monatStr As String: monatStr = ""
    If zeitbereich <> "" Then
        Dim ztParts() As String: ztParts = Split(zeitbereich, "-")
        If UBound(ztParts) >= 0 Then
            Dim letztesDat As String: letztesDat = Trim(ztParts(UBound(ztParts)))
            Dim dtParts() As String: dtParts = Split(letztesDat, ".")
            If UBound(dtParts) >= 1 Then
                Dim mNum As String: mNum = Trim(dtParts(1))
                If IsNumeric(mNum) Then monatStr = Right("0" & mNum, 2)
            End If
        End If
    End If
    Dim dateiSuffixPDF As String
    If monatStr <> "" Then
        dateiSuffixPDF = monatStr & "_" & jahreszahl
    Else
        dateiSuffixPDF = jahreszahl
    End If
    dateiname = ordner & "Mehrarbeit_" & kuerzel & "_" & dateiSuffixPDF & ".pdf"
    wsTmp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dateiname, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    Application.DisplayAlerts = False
    wsTmp.Delete
    Application.DisplayAlerts = True
End Sub

' ================================================================
'  Kennzahlen der Mehrarbeitsuebersicht (Einzel- und Sammel-PDF)
'  Schreibt ab Zeile z in Spalte A und gibt die naechste freie Zeile zurueck.
'  Die Korrektur wird zu den Vertretungen addiert:
'      Anzahl abrechenbare Vertretungen = Summenwert + Korrektur
'  Die Zeilen "Korrektur Vertretungen" und "Korrekturhinweis" erscheinen
'  nur, wenn eine Korrektur bzw. ein Hinweis eingetragen ist.
'  hinweisZeile (Rueckgabe): Zeile des Korrekturhinweises, 0 = keiner
' ================================================================
Private Function SchreibeKennzahlen(ws As Worksheet, ByVal z As Long, _
    ByVal bagatell As Double, ByVal entfGesamt As Long, ByVal entfNicht As Long, _
    ByVal vertWert As Long, ByVal vertNicht As Long, ByVal summeWert As Double, _
    ByVal korrWert As Double, ByVal korrHinweis As String, ByRef hinweisZeile As Long) As Long

    hinweisZeile = 0

    Dim bagatellText As String
    bagatellText = IIf(bagatell > 0, Format(bagatell, "0.00"), "k.A.")
    ws.Cells(z, 1).Value = "Bagatellgrenze:  " & bagatellText
    z = z + 1
    ws.Cells(z, 1).Value = "Entf" & Chr(228) & "lle gesamt:  " & entfGesamt & _
        "          nicht zu z" & Chr(228) & "hlende Entf" & Chr(228) & "lle:  " & entfNicht
    z = z + 1
    ws.Cells(z, 1).Value = "zu z" & Chr(228) & "hlende Vertretungen:  " & vertWert & _
        "          nicht zu z" & Chr(228) & "hlende Vertretungen:  " & vertNicht
    z = z + 1
    ' "erfasste Vertretungen" = unkorrigierter Summenwert aus der Uebersicht
    ws.Cells(z, 1).Value = "erfasste Vertretungen:  " & ZahlText(summeWert)
    z = z + 1

    ' Korrektur (wird zu den Vertretungen addiert) und Hinweis
    If korrWert <> 0 Or korrHinweis <> "" Then
        ws.Cells(z, 1).Value = "Korrektur Vertretungen:  " & IIf(korrWert > 0, "+", "") & ZahlText(korrWert)
        z = z + 1
        If korrHinweis <> "" Then
            ' Ueber die ganze Tabellenbreite, mit Zeilenumbruch bei langem Text
            With ws.Range(ws.Cells(z, 1), ws.Cells(z, 9))
                .Merge
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = True
            End With
            ws.Cells(z, 1).Value = "Korrekturhinweis:  " & korrHinweis
            hinweisZeile = z
            z = z + 1
        End If
    End If

    Dim abrechenbar As Double
    abrechenbar = summeWert + korrWert
    With ws.Cells(z, 1)
        .Value = "Anzahl abrechenbare Vertretungen:  " & ZahlText(abrechenbar) & _
                 IIf(korrWert <> 0, "   (inkl. Korrektur)", "")
        .Font.Bold = True
        If bagatell > 0 And abrechenbar > bagatell Then .Font.Color = RGB(0, 128, 0)
    End With
    z = z + 1

    SchreibeKennzahlen = z
End Function

' ----------------------------------------------------------------
' Zeilenhoehe der verbundenen Zeile mit dem Korrekturhinweis setzen
' (Excel passt verbundene Zellen nicht selbst an). Grobe Schaetzung:
' ein Zeichen je Einheit Spaltenbreite. Erst aufrufen, wenn die
' Spaltenbreiten feststehen. zeile <= 0: nichts zu tun.
' ----------------------------------------------------------------
Private Sub HinweisZeileAnpassen(ws As Worksheet, ByVal zeile As Long)
    If zeile <= 0 Then Exit Sub
    Dim breite As Double, c As Integer
    breite = 0
    For c = 1 To 9
        breite = breite + ws.Columns(c).ColumnWidth
    Next c
    If breite < 20 Then breite = 20
    Dim teile() As String, t As Long, anz As Long
    teile = Split(Replace(CStr(ws.Cells(zeile, 1).Value), Chr(13), ""), Chr(10))
    anz = 0
    For t = 0 To UBound(teile)
        anz = anz + Int(Len(teile(t)) / breite) + 1
    Next t
    If anz < 1 Then anz = 1
    If anz > 30 Then anz = 30
    ws.Rows(zeile).RowHeight = 13 * anz
End Sub

' ================================================================
'  MAKRO 5: Tabelle1 importieren
' ================================================================
Sub Tabelle1Importieren()
    Dim quelldatei As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Tabelle1-Quelldatei w" & Chr(228) & "hlen"
        .Filters.Clear
        .Filters.Add "Excel-Dateien", "*.xlsx; *.xlsm; *.xls; *.xlsb"
        .Filters.Add "Alle Dateien", "*.*"
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Keine Datei gew" & Chr(228) & "hlt. Abbruch.", vbInformation: Exit Sub
        End If
        quelldatei = .SelectedItems(1)
    End With

    If MsgBox("Tabelle1 wird vollst" & Chr(228) & "ndig gel" & Chr(246) & "scht und neu eingelesen." & Chr(10) & Chr(10) & _
              "Quelle: " & quelldatei & Chr(10) & Chr(10) & "Fortfahren?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Dim wbQuelle As Workbook
    On Error GoTo FehlerImport
    Set wbQuelle = Workbooks.Open(quelldatei, ReadOnly:=True)

    Dim wsQuelle As Worksheet: Set wsQuelle = wbQuelle.Sheets(1)
    Dim wsZiel As Worksheet:   Set wsZiel = ThisWorkbook.Sheets("Tabelle1")
    wsZiel.Cells.ClearContents
    wsZiel.Cells.ClearFormats

    Dim lastRowQ As Long: lastRowQ = wsQuelle.UsedRange.Rows.Count
    Dim lastColQ As Long: lastColQ = wsQuelle.UsedRange.Columns.Count

    wsQuelle.Range(wsQuelle.Cells(1, 1), wsQuelle.Cells(lastRowQ, lastColQ)).Copy
    wsZiel.Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    wbQuelle.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Tabelle1 importiert: " & lastRowQ & " Zeilen.", vbInformation
    Exit Sub

FehlerImport:
    Application.ScreenUpdating = True
    If Not wbQuelle Is Nothing Then wbQuelle.Close SaveChanges:=False
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub

' ================================================================
'  MAKRO 6: Buttons auf Informationen, Pseudonyme und Uebersicht
'  Einmalig ausfuehren - loescht vorhandene Makro-Buttons und
'  erstellt alle neu.
'  Tabelle Informationen: nur "Buttons erstellen", oben in Spalte SPALTE_INFO.
' ================================================================
Sub ButtonsErstellen()

    ' Spalte, in der der Button auf Tabelle Informationen beginnt
    Const SPALTE_INFO As String = "D"

    ' Buttons werden als Shapes (abgerundete Rechtecke) angelegt,
    ' damit sie auf allen Excel-Versionen zuverlaessig funktionieren.

    ' Zuerst alle vorhandenen Makro-Buttons auf den drei Blaettern loeschen
    Dim sheetNamen(2) As String
    sheetNamen(0) = "Informationen"
    sheetNamen(1) = "Pseudonyme"
    sheetNamen(2) = Chr(220) & "bersicht"

    Dim sn As Integer
    For sn = 0 To 2
        Dim wsB As Worksheet
        Dim gefunden As Boolean: gefunden = False
        Dim wsT As Worksheet
        For Each wsT In ThisWorkbook.Sheets
            If wsT.Name = sheetNamen(sn) Then
                Set wsB = wsT: gefunden = True: Exit For
            End If
        Next wsT
        If Not gefunden Then GoTo NaechstesBlatt

        ' Alle Shapes mit Makrozuweisung loeschen
        Dim shp As Shape
        Dim toDelete() As String
        Dim delCount As Integer: delCount = 0
        ReDim toDelete(0)
        For Each shp In wsB.Shapes
            If shp.OnAction <> "" Then
                ReDim Preserve toDelete(delCount)
                toDelete(delCount) = shp.Name
                delCount = delCount + 1
            End If
        Next shp
        Dim d As Integer
        For d = 0 To delCount - 1
            wsB.Shapes(toDelete(d)).Delete
        Next d

NaechstesBlatt:
    Next sn

    ' ------------------------------------------------------------
    '  Button-Definitionen: Name, Caption, Makro, Blatt
    ' ------------------------------------------------------------
    ' Format: Blattname | Caption | Makroname
    Dim btnDefs(24) As String   ' max 25 Buttons
    Dim btnCount As Integer: btnCount = 0

    ' Tabelle Informationen (blau) - nur "Buttons erstellen"
    ' (Tabelle1 importieren / Summen einlesen gibt es auf der Uebersicht)
    btnDefs(btnCount) = "Informationen|Buttons erstellen|ButtonsErstellen":                                           btnCount = btnCount + 1

    ' Tabelle Pseudonyme (gruen)
    btnDefs(btnCount) = "Pseudonyme|K" & Chr(252) & "rzel ersetzen|NamenskuerzelErsetzenUndZellenLoeschen":           btnCount = btnCount + 1
    btnDefs(btnCount) = "Pseudonyme|K" & Chr(252) & "rzel wiederherstellen|NamenskuerzelWiederherstellen":            btnCount = btnCount + 1

    ' Tabelle Uebersicht - blaue Buttons
    btnDefs(btnCount) = Chr(220) & "bersicht|Tabelle1 importieren|Tabelle1Importieren":                               btnCount = btnCount + 1
    btnDefs(btnCount) = Chr(220) & "bersicht|Summen einlesen|SummenInTabelleUebersicht":                              btnCount = btnCount + 1
    btnDefs(btnCount) = Chr(220) & "bersicht|K" & Chr(252) & "rzel ersetzen|NamenskuerzelErsetzenUndZellenLoeschen":  btnCount = btnCount + 1
    btnDefs(btnCount) = Chr(220) & "bersicht|K" & Chr(252) & "rzel wiederherstellen|NamenskuerzelWiederherstellen":   btnCount = btnCount + 1

    ' Tabelle Uebersicht - orange Buttons (PDF)
    btnDefs(btnCount) = Chr(220) & "bersicht|PDF - Alle markieren|PDFAlleMarkieren":                                  btnCount = btnCount + 1
    btnDefs(btnCount) = Chr(220) & "bersicht|PDF - Nur Abrechnungsf" & Chr(228) & "hige|PDFNurAbrechnungsfaehige":   btnCount = btnCount + 1
    btnDefs(btnCount) = Chr(220) & "bersicht|PDF - Auswahl leeren|PDFAuswahlLeeren":                                  btnCount = btnCount + 1
    btnDefs(btnCount) = Chr(220) & "bersicht|PDFs erstellen|PDFErstellen":                                            btnCount = btnCount + 1
    btnDefs(btnCount) = Chr(220) & "bersicht|PDFs drucken|PDFsDrucken":                                               btnCount = btnCount + 1
    btnDefs(btnCount) = Chr(220) & "bersicht|LBV-Formular erstellen|LBVFormularErstellen":                             btnCount = btnCount + 1
    ' LBV zusammenfassen (Modul2) - bisher von Hand angelegt, jetzt fest dabei
    btnDefs(btnCount) = Chr(220) & "bersicht|Merge LBV|LBVZusammenfassen":                                             btnCount = btnCount + 1
    ' LBV-Sammeldatei (Modul2): Einzel- ODER Gesamtformulare aller Lehrer in eine PDF
    btnDefs(btnCount) = Chr(220) & "bersicht|LBV-Sammeldatei|LBVSammeldateiErstellen":                                 btnCount = btnCount + 1

    ' ------------------------------------------------------------
    '  Buttons platzieren
    ' ------------------------------------------------------------
    ' Buttons werden nebeneinander in Zeile 1, ab Spalte A
    ' (oder rechts neben letztem Inhalt) platziert.
    ' Je Blatt werden alle zugehoerigen Buttons gesammelt.

    Dim btnBreite  As Double: btnBreite = 130    ' Punkte
    Dim btnHoehe   As Double: btnHoehe = 22      ' Punkte
    Dim btnAbstand As Double: btnAbstand = 6     ' Punkte zwischen Buttons
    Dim btnTop     As Double: btnTop = 4         ' Punkte vom oberen Rand

    Dim letztesSHeet As String: letztesSHeet = ""
    Dim btnLeft As Double: btnLeft = 4

    Dim b As Integer
    For b = 0 To btnCount - 1
        Dim parts() As String
        parts = Split(btnDefs(b), "|")
        Dim zielBlatt  As String: zielBlatt = parts(0)
        Dim caption    As String: caption = parts(1)
        Dim makroName  As String: makroName = parts(2)

        ' Zielblatt suchen
        Dim wsZiel As Worksheet
        Set wsZiel = Nothing
        Dim wsX As Worksheet
        For Each wsX In ThisWorkbook.Sheets
            If wsX.Name = zielBlatt Then Set wsZiel = wsX: Exit For
        Next wsX
        If wsZiel Is Nothing Then GoTo NaechsterButton

        ' X-Position zuruecksetzen wenn neues Blatt
        ' (Informationen: linke Kante von Spalte SPALTE_INFO, sonst ab Spalte A)
        If zielBlatt <> letztesSHeet Then
            If zielBlatt = "Informationen" Then
                btnLeft = wsZiel.Columns(SPALTE_INFO).Left + 4
            Else
                btnLeft = 4
            End If
            letztesSHeet = zielBlatt
        End If

        ' Shape erstellen
        Dim newShp As Shape
        Set newShp = wsZiel.Shapes.AddShape( _
            msoShapeRoundedRectangle, btnLeft, btnTop, btnBreite, btnHoehe)

        With newShp
            .Name = "Btn_" & makroName
            .OnAction = makroName

            ' Farbe: blau fuer Import/Summen/Kuerzel, orange fuer PDF, gruen fuer Pseudonyme
            If makroName = "PDFAlleMarkieren" Or makroName = "PDFNurAbrechnungsfaehige" Or _
               makroName = "PDFAuswahlLeeren" Or makroName = "PDFErstellen" Then
                .Fill.ForeColor.RGB = RGB(197, 90, 17)     ' Orange fuer PDF-Buttons
            ElseIf zielBlatt = "Pseudonyme" Then
                .Fill.ForeColor.RGB = RGB(112, 173, 71)    ' Gruen fuer Pseudonyme
            Else
                .Fill.ForeColor.RGB = RGB(68, 114, 196)    ' Blau fuer alle anderen
            End If
            .Fill.Transparency = 0
            .Line.Visible = msoFalse

            ' Text
            With .TextFrame2.TextRange
                .Text = caption
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Font.Bold = msoTrue
                .Font.Size = 9
            End With
            .TextFrame.HorizontalAlignment = xlHAlignCenter
            .TextFrame.VerticalAlignment = xlVAlignCenter
        End With

        btnLeft = btnLeft + btnBreite + btnAbstand

NaechsterButton:
    Next b

    MsgBox "Buttons wurden erfolgreich erstellt!", vbInformation

End Sub

' ================================================================
'  Sammel-PDF: alle Einzelblaetter auf ein temporaeres Blatt sammeln
'  und als eine PDF exportieren
' ================================================================
Private Sub ErstelleSammelPDF(ordner As String, sammelK() As String, anzahl As Integer, _
    jahreszahl As String, zeitbereich As String, ws1 As Worksheet, wsUe As Worksheet, _
    colSumme As Integer, colBagatell As Integer, colEntfG As Integer, _
    colEntfN As Integer, colVert As Integer, colVertN As Integer, _
    colPersNr As Integer, lastRowUe As Long, lastRow1 As Long, _
    colKorr As Integer, colKorrH As Integer)

    ' Temporaeres Sammelblatt anlegen
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "_SAMMEL_TMP_" Then ws.Delete: Exit For
    Next ws
    Application.DisplayAlerts = True

    Dim wsSammel As Worksheet
    Set wsSammel = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsSammel.Name = "_SAMMEL_TMP_"
    wsSammel.Cells.Font.Name = "Arial"
    wsSammel.Cells.Font.Size = 10

    Dim sammelZeile As Long: sammelZeile = 1
    Dim idx As Integer
    ' Zeilen mit Korrekturhinweis (Hoehe wird am Ende angepasst)
    Dim hinweisZeilen() As Long
    Dim anzHinweise As Long: anzHinweise = 0
    ReDim hinweisZeilen(0)

    For idx = 0 To anzahl - 1
        Dim kuerzel As String: kuerzel = sammelK(idx)

        ' Wir bauen den Inhalt direkt neu auf (gleiche Logik wie ErstellePDF)
        Dim vollname As String:    vollname = ""
        zeitbereich = ""
        Dim persNr As String:      persNr = ""
        Dim detailZeilen() As String
        Dim anzDetail As Integer: anzDetail = 0
        ReDim detailZeilen(0)

        ' Vollname und Zeitbereich aus Tabelle1
        Dim inBlock As Boolean: inBlock = False
        Dim i As Long
        For i = 1 To lastRow1
            Dim zv As String: zv = Trim(ws1.Cells(i, 1).Value)
            If Not inBlock Then
                If LCase(zv) = LCase(kuerzel) Then
                    vollname = Trim(ws1.Cells(i, 2).Value): inBlock = True
                End If
            Else
                If InStr(zv, "Vertretungen / Entf") > 0 Then
                    Dim dp As Integer: dp = InStr(zv, ":")
                    If dp > 0 Then zeitbereich = Trim(Mid(zv, dp + 1))
                ElseIf Left(zv, 10) = "Vertretung" And InStr(zv, "/") = 0 Then
                    ' Alle Vertretungs-Zusammenfassungszeilen ignorieren
                ElseIf Left(zv, 4) = "Entf" Or Left(zv, 5) = "Summe" Or _
                       Left(zv, 12) = "Bereitschaft" Or Left(zv, 13) = "Pausenaufsich" Or _
                       Left(zv, 3) = "z" & Chr(228) & "h" Or zv = "zu" Or zv = "nicht" Or _
                       Left(zv, 5) = "z" & Chr(228) & "hle" Then
                    ' Zusammenfassungszeilen und Fortsetzungsfragmente ignorieren
                ElseIf zv = "Datum" Then
                ElseIf zv <> "" Then
                    Dim zvN As String: zvN = ""
                    Dim ii As Long
                    For ii = i + 1 To Application.Min(i + 5, lastRow1)
                        zvN = Trim(ws1.Cells(ii, 1).Value)
                        If zvN <> "" Then Exit For
                    Next ii
                    If InStr(zvN, "Vertretungen / Entf") > 0 Then
                        Exit For
                    Else
                        Dim zeile As String: zeile = ""
                        Dim s As Integer
                        For s = 1 To 9
                            zeile = zeile & IIf(s > 1, Chr(9), "") & Trim(ws1.Cells(i, s).Value)
                        Next s
                        ReDim Preserve detailZeilen(anzDetail)
                        detailZeilen(anzDetail) = zeile
                        anzDetail = anzDetail + 1
                    End If
                End If
            End If
        Next i

        ' Werte aus Uebersicht
        Dim summeWert As Double, entfGesamt As Long, entfNicht As Long
        Dim vertWert As Long, vertNicht As Long, bagatell As Double
        Dim korrWert As Double:    korrWert = 0
        Dim korrHinweis As String: korrHinweis = ""
        Dim ueZ As Long
        For ueZ = 2 To lastRowUe
            If LCase(Trim(wsUe.Cells(ueZ, 1).Value)) = LCase(kuerzel) Then
                korrWert = KorrekturWert(wsUe, ueZ, colKorr)
                korrHinweis = KorrekturHinweis(wsUe, ueZ, colKorrH)
                If colSumme > 0 And IsNumeric(wsUe.Cells(ueZ, colSumme).Value) Then summeWert = CDbl(wsUe.Cells(ueZ, colSumme).Value)
                If colBagatell > 0 And IsNumeric(wsUe.Cells(ueZ, colBagatell).Value) Then bagatell = CDbl(wsUe.Cells(ueZ, colBagatell).Value)
                If colEntfG > 0 And IsNumeric(wsUe.Cells(ueZ, colEntfG).Value) Then entfGesamt = CLng(wsUe.Cells(ueZ, colEntfG).Value)
                If colEntfN > 0 And IsNumeric(wsUe.Cells(ueZ, colEntfN).Value) Then entfNicht = CLng(wsUe.Cells(ueZ, colEntfN).Value)
                If colVert > 0 And IsNumeric(wsUe.Cells(ueZ, colVert).Value) Then vertWert = CLng(wsUe.Cells(ueZ, colVert).Value)
                If colVertN > 0 And IsNumeric(wsUe.Cells(ueZ, colVertN).Value) Then vertNicht = CLng(wsUe.Cells(ueZ, colVertN).Value)
                If colPersNr > 0 And Trim(wsUe.Cells(ueZ, colPersNr).Value) <> "" Then persNr = Trim(wsUe.Cells(ueZ, colPersNr).Value)
                Exit For
            End If
        Next ueZ

        ' --- Inhalt auf Sammelblatt schreiben ---
        ' Seitenumbruch vor jedem Lehrer (ausser dem ersten)
        If idx > 0 Then
            wsSammel.HPageBreaks.Add Before:=wsSammel.Rows(sammelZeile)
        End If

        ' Titel
        With wsSammel.Cells(sammelZeile, 1)
            .Value = "Mehrarbeits" & Chr(252) & "bersicht " & zeitbereich & " " & jahreszahl
            .Font.Bold = True: .Font.Size = 13
        End With
        sammelZeile = sammelZeile + 2

        ' Name
        Dim nameZeile As String
        nameZeile = kuerzel & IIf(vollname <> "", "   " & vollname, "")
        If persNr <> "" Then nameZeile = nameZeile & "   (Pers.Nr: " & persNr & ")"
        With wsSammel.Cells(sammelZeile, 1)
            .Value = nameZeile
            .Font.Size = 20
            .Font.Bold = True
        End With
        sammelZeile = sammelZeile + 2

        ' Tabellenkopf
        Dim hdrs(8) As String
        hdrs(0) = "Datum": hdrs(1) = "Stunde": hdrs(2) = "Art": hdrs(3) = "Wert"
        hdrs(4) = Chr(90) & Chr(228) & "hler": hdrs(5) = "Fach"
        hdrs(6) = "Klasse(n)": hdrs(7) = "Grund": hdrs(8) = "Text"
        Dim h As Integer
        For h = 0 To 8
            With wsSammel.Cells(sammelZeile, h + 1)
                .Value = hdrs(h): .Font.Bold = True
                .Interior.Color = RGB(180, 200, 230)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
            End With
        Next h
        sammelZeile = sammelZeile + 1

        ' Detailzeilen
        Dim d As Integer
        For d = 0 To anzDetail - 1
            Dim felder() As String
            felder = Split(detailZeilen(d), Chr(9))
            Dim f As Integer
            For f = 0 To Application.Min(UBound(felder), 8)
                wsSammel.Cells(sammelZeile, f + 1).Value = Trim(felder(f))
            Next f
            wsSammel.Range(wsSammel.Cells(sammelZeile, 1), wsSammel.Cells(sammelZeile, 9)).WrapText = True
            If d Mod 2 = 1 Then
                wsSammel.Range(wsSammel.Cells(sammelZeile, 1), wsSammel.Cells(sammelZeile, 9)).Interior.Color = RGB(242, 242, 242)
            End If
            sammelZeile = sammelZeile + 1
        Next d

        ' Trennlinie
        wsSammel.Range(wsSammel.Cells(sammelZeile - 1, 1), wsSammel.Cells(sammelZeile - 1, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        sammelZeile = sammelZeile + 1

        ' Kennzahlen inkl. Korrektur
        ' ("Zur Kenntnis genommen" und Unterschriftsfelder entfallen)
        Dim hinweisZeile As Long
        sammelZeile = SchreibeKennzahlen(wsSammel, sammelZeile, bagatell, entfGesamt, entfNicht, _
                                         vertWert, vertNicht, summeWert, korrWert, korrHinweis, hinweisZeile)
        If hinweisZeile > 0 Then
            ReDim Preserve hinweisZeilen(anzHinweise)
            hinweisZeilen(anzHinweise) = hinweisZeile
            anzHinweise = anzHinweise + 1
        End If
        sammelZeile = sammelZeile + 1

    Next idx

    ' Spaltenbreiten anpassen
    wsSammel.Columns("A:I").AutoFit
    wsSammel.Columns(1).ColumnWidth = 7
    wsSammel.Columns(2).ColumnWidth = 7
    wsSammel.Columns(5).ColumnWidth = 5

    ' Zeilenhoehe der Korrekturhinweise (verbundene Zellen)
    Dim hz As Long
    For hz = 0 To anzHinweise - 1
        HinweisZeileAnpassen wsSammel, hinweisZeilen(hz)
    Next hz

    ' Seiteneinrichtung
    With wsSammel.PageSetup
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PaperSize = xlPaperA4
        .Zoom = False
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
    End With

    ' Sammel-PDF exportieren
    Dim sammelDatei As String
    ' Monat aus Zeitbereich fuer Sammel-PDF
    Dim sammelMonat As String: sammelMonat = ""
    If zeitbereich <> "" Then
        Dim szParts() As String: szParts = Split(zeitbereich, "-")
        If UBound(szParts) >= 0 Then
            Dim slDat As String: slDat = Trim(szParts(UBound(szParts)))
            Dim sdParts() As String: sdParts = Split(slDat, ".")
            If UBound(sdParts) >= 1 Then
                Dim smNum As String: smNum = Trim(sdParts(1))
                If IsNumeric(smNum) Then sammelMonat = Right("0" & smNum, 2)
            End If
        End If
    End If
    Dim sammelSuffix As String
    If sammelMonat <> "" Then
        sammelSuffix = sammelMonat & "_" & jahreszahl
    Else
        sammelSuffix = jahreszahl
    End If
    sammelDatei = ordner & "Mehrarbeit_ALLE_" & sammelSuffix & ".pdf"
    wsSammel.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sammelDatei, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    ' Sammelblatt loeschen
    Application.DisplayAlerts = False
    wsSammel.Delete
    Application.DisplayAlerts = True

End Sub
' ================================================================
'  MAKRO 8: LBV-Formular direkt in Original-PDF befuellen
'  Nutzt Python/pypdf um das Original-Formular zu befuellen
' ================================================================
Sub LBVFormularErstellen()

    Dim ws1  As Worksheet
    Dim wsUe As Worksheet
    Dim wsInfo As Worksheet

    Set ws1 = ThisWorkbook.Sheets("Tabelle1")
    Set wsUe = ThisWorkbook.Sheets(Chr(220) & "bersicht")

    Dim wsI As Worksheet
    For Each wsI In ThisWorkbook.Sheets
        If LCase(wsI.Name) = "informationen" Then Set wsInfo = wsI: Exit For
    Next wsI

    Dim jahreszahl      As String: jahreszahl = ""
    Dim aktenzeichen    As String: aktenzeichen = ""
    Dim dienststelle    As String: dienststelle = ""
    Dim datumStr        As String: datumStr = ""
    Dim bearbeiter      As String: bearbeiter = ""
    Dim telefon         As String: telefon = ""
    Dim schulnummer     As String: schulnummer = ""
    Dim minutenJeStunde As String: minutenJeStunde = ""
    Dim schluessel       As String: schluessel = "2"        ' Standard: Schluessel 2 (Gymnasium)
    Dim feldC1          As String: feldC1 = ""              ' (bisher ungenutzt, bleibt leer)

    If Not wsInfo Is Nothing Then
        ' Werte per Zeilenbeschriftung suchen - robust gegen eingefuegte Zeilen
        Dim infoRow As Long
        Dim infoLastRow As Long
        infoLastRow = wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
        For infoRow = 1 To infoLastRow
            Dim infoKey As String
            Dim infoVal As String
            infoKey = LCase(Trim(CStr(wsInfo.Cells(infoRow, 1).Value)))
            infoVal = Trim(CStr(wsInfo.Cells(infoRow, 2).Value))
            Select Case infoKey
                Case "jahreszahl"
                    jahreszahl = infoVal
                Case "aktenzeichen"
                    aktenzeichen = infoVal
                Case "dienststelle", "dienststelle, anschrift"
                    dienststelle = infoVal
                Case "datum"
                    datumStr = infoVal
                Case "bearbeiter"
                    bearbeiter = infoVal
                Case "telefon", "telefon-durchwahl"
                    telefon = infoVal
                Case "schulnummer"
                    If infoVal <> "" Then schulnummer = infoVal
                Case "minuten", "anzahl der minuten je unterrichtsstunde"
                    If IsNumeric(infoVal) Then
                        minutenJeStunde = Right("0000" & CStr(CLng(CDbl(infoVal) * 100)), 5)
                    ElseIf Len(infoVal) = 5 Then
                        minutenJeStunde = infoVal
                    End If
                Case "standardschl" & Chr(252) & "ssel", "schluessel"
                    If infoVal <> "" And IsNumeric(infoVal) Then
                        Dim schlNum As Long
                        schlNum = CLng(CDbl(infoVal))
                        If schlNum >= 1 And schlNum <= 4 Then schluessel = CStr(schlNum)
                    End If
            End Select
        Next infoRow
    End If

    Dim colPDF      As Integer: colPDF = SpaltePDF()
    Dim colSumme    As Integer: colSumme = SpalteUebersicht(wsUe, "Summenwert")
    Dim colPersNr   As Integer: colPersNr = SpalteUebersicht(wsUe, "Pers.Nr")
    Dim colSollUe   As Integer: colSollUe = SpalteUebersicht(wsUe, "Soll/Woche")
    Dim colSchlUe   As Integer: colSchlUe = SpalteUebersicht(wsUe, "Schl" & Chr(252) & "ssel")
    Dim colStatus   As Integer: colStatus = SpalteUebersicht(wsUe, "Status")
    Dim colGeburt   As Integer: colGeburt = SpalteUebersicht(wsUe, "Geburtsdatum")
    Dim colAktzeich As Integer: colAktzeich = SpalteUebersicht(wsUe, "Aktenzeichen")
    Dim colKorr     As Integer: colKorr = SpalteUebersicht(wsUe, "Korrektur")

    Dim lastRowUe As Long
    lastRowUe = wsUe.Cells(wsUe.Rows.Count, "A").End(xlUp).Row

    Dim anzahlMarkiert As Integer: anzahlMarkiert = 0
    Dim r As Long
    For r = 2 To lastRowUe
        If wsUe.Cells(r, colPDF).Value = 1 Then anzahlMarkiert = anzahlMarkiert + 1
    Next r
    If anzahlMarkiert = 0 Then
        MsgBox "Keine Lehrer markiert.", vbExclamation: Exit Sub
    End If
    If Not KorrekturPruefen(wsUe, colKorr, colPDF, lastRowUe) Then Exit Sub

    ' Vorlagepfad aus Tabelle Informationen
    Dim vorlagePfad As String
    Dim vorlagenOrdner As String: vorlagenOrdner = ""
    Dim vorlagenDatei As String: vorlagenDatei = ""
    If Not wsInfo Is Nothing Then
        Dim viRow As Long
        For viRow = 1 To wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
            Dim viKey As String: viKey = LCase(Trim(CStr(wsInfo.Cells(viRow, 1).Value)))
            If viKey = "pfad zur manuell bef" & Chr(252) & "llten vorlage" Then vorlagenOrdner = Trim(CStr(wsInfo.Cells(viRow, 2).Value))
            If viKey = "dateiname vorlage" Then vorlagenDatei = Trim(CStr(wsInfo.Cells(viRow, 2).Value))
        Next viRow
    End If
    If vorlagenOrdner = "" Or vorlagenDatei = "" Then
        MsgBox "Bitte Pfad und Dateiname der Vorlage in Tabelle Informationen (Zeilen 10+11) eintragen.", vbExclamation: Exit Sub
    End If
    If Right(vorlagenOrdner, 1) <> "\" Then vorlagenOrdner = vorlagenOrdner & "\"
    vorlagePfad = vorlagenOrdner & vorlagenDatei & ".pdf"
    If Dir(vorlagePfad) = "" Then
        MsgBox "Vorlagen-PDF nicht gefunden:" & Chr(10) & vorlagePfad, vbCritical: Exit Sub
    End If

    Dim speicherOrdner As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Speicherordner f" & Chr(252) & "r LBV-Formulare w" & Chr(228) & "hlen"
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Kein Ordner gew" & Chr(228) & "hlt. Abbruch.", vbInformation: Exit Sub
        End If
        speicherOrdner = .SelectedItems(1)
    End With
    If Right(speicherOrdner, 1) <> "\" Then speicherOrdner = speicherOrdner & "\"

    Dim lastRow1 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    Dim erstellt As Integer: erstellt = 0
    Dim fehlerListe As String: fehlerListe = ""
    Dim ohneStunden As String: ohneStunden = ""   ' Lehrer mit weniger als 1 abrechenbaren Vertretung

    For r = 2 To lastRowUe
        If wsUe.Cells(r, colPDF).Value <> 1 Then GoTo NaechsterLehrerLBV

        Dim kuerzel  As String: kuerzel = Trim(wsUe.Cells(r, 1).Value)
        Dim persNrL  As String: persNrL = ""
        Dim summeL   As Double: summeL = 0
        Dim zeitbereichL As String: zeitbereichL = ""

        If kuerzel = "" Then GoTo NaechsterLehrerLBV

        persNrL = Trim(CStr(wsUe.Cells(r, colPersNr).Value))
        Dim geburtsdatumL As String: geburtsdatumL = ""
        If colGeburt > 0 Then
            Dim gebRaw As Variant: gebRaw = wsUe.Cells(r, colGeburt).Value
            If IsDate(gebRaw) Then
                geburtsdatumL = Format(CDate(gebRaw), "DD.MM.YYYY")
            ElseIf Trim(CStr(gebRaw)) <> "" Then
                geburtsdatumL = Trim(CStr(gebRaw))
            End If
        End If
        ' Aktenzeichen pro Lehrer
        Dim aktenzeichenL As String: aktenzeichenL = aktenzeichen
        If colAktzeich > 0 Then
            Dim aktRaw As String: aktRaw = Trim(CStr(wsUe.Cells(r, colAktzeich).Value))
            If aktRaw <> "" Then aktenzeichenL = aktRaw
        End If
        If colSumme > 0 Then
            Dim summeRaw As String: summeRaw = Trim(CStr(wsUe.Cells(r, colSumme).Value))
            If IsNumeric(summeRaw) Then summeL = CDbl(summeRaw)
        End If
        ' Korrektur zu den Vertretungen addieren -> Stundenzahl im LBV-Formular
        summeL = summeL + KorrekturWert(wsUe, r, colKorr)
        ' Formular nur mit mindestens 1 abrechenbaren Stunde
        ' (sonst stuende z.B. "0-2" bzw. "000" im Feld A)
        If CLng(summeL) < 1 Then
            ohneStunden = ohneStunden & IIf(ohneStunden = "", "", ", ") & kuerzel & " (" & ZahlText(summeL) & ")"
            GoTo NaechsterLehrerLBV
        End If
        ' Sollstunden fuer Merkmal A (Teilzeit < 25.5 Std)
        Dim sollL As Double: sollL = 0
        Dim sollRawL As String: sollRawL = ""
        If colSollUe > 0 Then sollRawL = Trim(CStr(wsUe.Cells(r, colSollUe).Value))
        If sollRawL <> "" And IsNumeric(sollRawL) Then sollL = CDbl(sollRawL)
        ' Status: B=Beamter, A=Angestellter - aus Spalte Status der Uebersicht
        Dim statusL As String: statusL = "B"  ' Standard: Beamter
        If colStatus > 0 Then
            Dim statusRaw As String: statusRaw = UCase(Trim(CStr(wsUe.Cells(r, colStatus).Value)))
            If statusRaw = "A" Or statusRaw = "TV" Or statusRaw = "TARIF" Then statusL = "A"
        End If
        Dim merkmalL As String: merkmalL = ""
        If statusL = "B" And sollL > 0 And sollL < 25.5 Then merkmalL = "A"
        ' Teilzeit-Angestellte (<25.5 Std): Fehlermeldung
        Dim inAbschnittC As Boolean
        inAbschnittC = False
        If statusL = "A" And sollL > 0 And sollL < 25.5 Then
            fehlerListe = fehlerListe & kuerzel & " (Tarifbesch" & Chr(228) & "ftigte in Teilzeit - bitte manuell eingeben)" & Chr(10)
            GoTo NaechsterLehrerLBV
        End If
        ' Schluessel: aus Uebers.-Spalte falls vorhanden, sonst globaler Schluessel
        Dim schluesselL As String: schluesselL = schluessel
        If colSchlUe > 0 Then
            Dim schlUeRaw As String: schlUeRaw = Trim(CStr(wsUe.Cells(r, colSchlUe).Value))
            If schlUeRaw <> "" And IsNumeric(schlUeRaw) Then
                Dim schlUeNum As Long: schlUeNum = CLng(CDbl(schlUeRaw))
                If schlUeNum >= 1 And schlUeNum <= 4 Then schluesselL = CStr(schlUeNum)
            End If
        End If

        Dim vollnameL As String: vollnameL = ""
        Dim i As Long
        Dim inBlockL As Boolean: inBlockL = False
        For i = 1 To lastRow1
            Dim zvL As String: zvL = Trim(ws1.Cells(i, 1).Value)
            ' Kuerzel-Zeile: Vollname aus Spalte B lesen
            If LCase(zvL) = LCase(kuerzel) Then
                vollnameL = Trim(ws1.Cells(i, 2).Value)
                inBlockL = True
            End If
            ' Zeitbereich aus naechster Vertretungs-Zeile nach dem Kuerzel
            If inBlockL And InStr(zvL, "Vertretungen / Entf") > 0 And zeitbereichL = "" Then
                Dim dpL As Integer: dpL = InStr(zvL, ":")
                If dpL > 0 Then
                    zeitbereichL = Trim(Mid(zvL, dpL + 1))
                    Exit For  ' Zeitbereich gefunden - fertig
                End If
            End If
        Next i

        ' Name/Vorname aus Pseudonyme Spalten C und D
        Dim nachname As String: nachname = ""
        Dim vorname  As String: vorname = ""
        Dim wsPseudo As Worksheet
        Set wsPseudo = ThisWorkbook.Sheets("Pseudonyme")
        Dim pRow As Long
        Dim lastPRow As Long: lastPRow = wsPseudo.Cells(wsPseudo.Rows.Count, "A").End(xlUp).Row
        For pRow = 2 To lastPRow
            If LCase(Trim(wsPseudo.Cells(pRow, 1).Value)) = LCase(kuerzel) Or _
               LCase(Trim(wsPseudo.Cells(pRow, 2).Value)) = LCase(kuerzel) Then
                nachname = Trim(CStr(wsPseudo.Cells(pRow, 3).Value))
                vorname = Trim(CStr(wsPseudo.Cells(pRow, 4).Value))
                Exit For
            End If
        Next pRow
        If nachname = "" And vollnameL <> "" Then
            If InStr(vollnameL, " ") > 0 Then
                Dim t() As String: t = Split(vollnameL, " ")
                vorname = t(UBound(t))
                nachname = Left(vollnameL, Len(vollnameL) - Len(vorname) - 1)
            Else
                nachname = vollnameL
            End If
        End If

        ' Zeitbereich -> MMJJ
        Dim monatJahr As String: monatJahr = ""
        If zeitbereichL <> "" Then
            Dim zt() As String: zt = Split(zeitbereichL, "-")
            If UBound(zt) >= 0 Then
                Dim ld As String: ld = Trim(zt(UBound(zt)))
                Dim dt() As String: dt = Split(ld, ".")
                If UBound(dt) >= 1 Then
                    monatJahr = Right("0" & Trim(dt(1)), 2) & Right(jahreszahl, 2)
                End If
            End If
        End If

        Dim stundenStr As String
        stundenStr = Right("00" & CStr(CLng(summeL)), 3)
        Dim feldA1 As String
        feldA1 = monatJahr & " " & stundenStr & " " & schluesselL

        Dim ausgabePfad As String
        ' Dateiname: Monat+Jahr aus FeldA falls vorhanden, sonst nur Jahr
        Dim dateiSuffix As String
        If monatJahr <> "" Then
            ' monatJahr = MMJJ z.B. '0925' -> '09_2025'
            dateiSuffix = Left(monatJahr, 2) & "_" & Left(jahreszahl, 2) & Right(jahreszahl, 2)
        Else
            dateiSuffix = jahreszahl
        End If
        ausgabePfad = speicherOrdner & "LBV_Mehrarbeit_" & SafeFileName(kuerzel) & "_" & dateiSuffix & ".pdf"

        ' JSON-Datei schreiben
        Dim jsonPfad As String
        jsonPfad = Environ("TEMP") & "\lbv_data_" & SafeFileName(kuerzel) & ".json"

        Dim fJson As Integer: fJson = FreeFile
        Open jsonPfad For Output As #fJson
        Print #fJson, "{"
        Print #fJson, "  " & Chr(34) & "vorlage_pfad" & Chr(34) & ": " & Chr(34) & Replace(vorlagePfad, "\", "/") & Chr(34) & ","
        Print #fJson, "  " & Chr(34) & "ausgabe_pfad" & Chr(34) & ": " & Chr(34) & Replace(ausgabePfad, "\", "/") & Chr(34) & ","
        Print #fJson, "  " & Chr(34) & "val_a1" & Chr(34) & ": " & Chr(34) & feldA1 & Chr(34) & ","
        Print #fJson, "  " & Chr(34) & "val_merkmal" & Chr(34) & ": " & Chr(34) & merkmalL & Chr(34) & ","
        Print #fJson, "  " & Chr(34) & "val_c1" & Chr(34) & ": " & Chr(34) & feldC1 & Chr(34) & ","
        Print #fJson, "  " & Chr(34) & "felder" & Chr(34) & ": {"
        Print #fJson, "    " & Chr(34) & "LBV_Nachname" & Chr(34) & ": " & Chr(34) & EscapeStr(nachname) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "LBV_Vorname" & Chr(34) & ": " & Chr(34) & EscapeStr(vorname) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "LBV_Geburtsdatum" & Chr(34) & ": " & Chr(34) & EscapeStr(geburtsdatumL) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "LBV_Personalnummer" & Chr(34) & ": " & Chr(34) & EscapeStr(persNrL) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "LBV_Schulnummer" & Chr(34) & ": " & Chr(34) & EscapeStr(schulnummer) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "LBV_Stundentaktung" & Chr(34) & ": " & Chr(34) & EscapeStr(minutenJeStunde) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "Dienststelle" & Chr(34) & ": " & Chr(34) & EscapeStr(dienststelle) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "Aktenzeichen" & Chr(34) & ": " & Chr(34) & EscapeStr(aktenzeichenL) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "Datum" & Chr(34) & ": " & Chr(34) & EscapeStr(datumStr) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "Bearbeiter" & Chr(34) & ": " & Chr(34) & EscapeStr(bearbeiter) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "Telefon" & Chr(34) & ": " & Chr(34) & EscapeStr(telefon) & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "LBV_Checkbox_Nein" & Chr(34) & ": " & Chr(34) & "/Ja" & Chr(34) & ","
        Print #fJson, "    " & Chr(34) & "LBV_Merkmal_FeldA_1" & Chr(34) & ": " & Chr(34) & merkmalL & Chr(34)
        Print #fJson, "  }"
        Print #fJson, "}"
        Close #fJson

        ' EXE aufrufen
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        Dim exitCode As Long
        Dim logPfad As String
        logPfad = Environ("TEMP") & "\lbv_error_" & SafeFileName(kuerzel) & ".txt"
        ' exePfad aus Tabelle Informationen lesen (Zeile mit 'Pfad lbv_fill.exe')
        Dim exePfad As String: exePfad = ""
        If Not wsInfo Is Nothing Then
            Dim exeRow As Long
            For exeRow = 1 To wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
                If LCase(Trim(CStr(wsInfo.Cells(exeRow, 1).Value))) = "pfad lbv_fill.exe" Then
                    exePfad = Trim(CStr(wsInfo.Cells(exeRow, 2).Value))
                    Exit For
                End If
            Next exeRow
        End If
        ' Falls nicht gespeichert: Dialog anzeigen
        If exePfad = "" Or Dir(exePfad) = "" Then
            With Application.FileDialog(msoFileDialogFilePicker)
                .Title = "lbv_fill.exe ausw" & Chr(228) & "hlen"
                .Filters.Clear
                .Filters.Add "Anwendungen", "*.exe"
                .AllowMultiSelect = False
                .Show
                If .SelectedItems.Count = 0 Then
                    MsgBox "Kein Pfad gew" & Chr(228) & "hlt. Abbruch.", vbInformation: Exit Sub
                End If
                exePfad = .SelectedItems(1)
            End With
            ' Pfad in Tabelle Informationen speichern
            If Not wsInfo Is Nothing Then
                Dim newRow As Long
                newRow = wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row + 1
                ' Pruefe ob Zeile bereits existiert
                Dim exeRowSave As Long: exeRowSave = 0
                For exeRow = 1 To newRow - 1
                    If LCase(Trim(CStr(wsInfo.Cells(exeRow, 1).Value))) = "pfad lbv_fill.exe" Then
                        exeRowSave = exeRow: Exit For
                    End If
                Next exeRow
                If exeRowSave = 0 Then exeRowSave = newRow
                wsInfo.Cells(exeRowSave, 1).Value = "Pfad lbv_fill.exe"
                wsInfo.Cells(exeRowSave, 2).Value = exePfad
            End If
        End If
        ' EXE ueber temporaere Batch-Datei aufrufen (vermeidet cmd-Quoting-Probleme)
        Dim useExe As Boolean: useExe = False
        If Len(Dir(exePfad)) > 0 Then
            Dim batPfad As String
            batPfad = Environ("TEMP") & "\lbv_run_" & SafeFileName(kuerzel) & ".bat"
            Dim fBat As Integer: fBat = FreeFile
            Open batPfad For Output As #fBat
            Print #fBat, "@echo off"
            Print #fBat, Chr(34) & exePfad & Chr(34) & " " & Chr(34) & jsonPfad & Chr(34)
            Print #fBat, "exit /b %errorlevel%"
            Close #fBat
            exitCode = wsh.Run("cmd /c " & Chr(34) & batPfad & Chr(34), 0, True)
            If exitCode = 0 Or Len(Dir(ausgabePfad)) > 0 Then useExe = True
        Else
            MsgBox "lbv_fill.exe nicht gefunden:" & Chr(10) & exePfad & Chr(10) & Chr(10) & _
                   "Bitte Pfad in Tabelle Informationen pr" & Chr(252) & "fen.", vbCritical, "EXE fehlt"
            GoTo NaechsterLehrerLBV
        End If
        ' Python-Fallback: nur wenn .py vorhanden UND EXE fehlgeschlagen
        If Not useExe Then
            Dim pySkript As String
            pySkript = Left(exePfad, InStrRev(exePfad, "\")) & "lbv_fill.py"
            If Dir(pySkript) <> "" Then
                Dim pyCmd As String: pyCmd = "py"
                If wsh.Run("cmd /c py --version > nul 2>&1", 0, True) <> 0 Then pyCmd = "python"
                exitCode = wsh.Run("cmd /c " & pyCmd & " """ & pySkript & """ """ & jsonPfad & """ > """ & logPfad & """ 2>&1", 0, True)
                If exitCode = 0 Then useExe = True
            End If
        End If


        ' Kurz warten damit Datei geschrieben werden kann
        Application.Wait Now + TimeValue("00:00:02")
        If exitCode = 0 Then
            erstellt = erstellt + 1
        Else
            fehlerListe = fehlerListe & kuerzel & " (Exit: " & exitCode & ")" & Chr(10)
            ' Log-Datei lesen und anzeigen
            Dim diagMsg As String
            diagMsg = "Ausgabe: " & ausgabePfad & Chr(10) & "Log: " & logPfad & Chr(10)
            If Dir(logPfad) <> "" Then
                Dim fLog As Integer: fLog = FreeFile
                Dim logText As String: logText = ""
                Dim logLine As String
                Open logPfad For Input As #fLog
                Dim logCount As Integer: logCount = 0
                Do While Not EOF(fLog) And logCount < 30
                    Line Input #fLog, logLine
                    logText = logText & logLine & Chr(10)
                    logCount = logCount + 1
                Loop
                Close #fLog
                diagMsg = diagMsg & Chr(10) & logText
            End If
            MsgBox diagMsg, vbCritical, "Fehler bei " & kuerzel
        End If

NaechsterLehrerLBV:
    Next r

    Dim meldungLBV As String
    meldungLBV = erstellt & " LBV-Formular(e) erstellt in:" & Chr(10) & speicherOrdner
    If ohneStunden <> "" Then
        meldungLBV = meldungLBV & Chr(10) & Chr(10) & _
                     "Kein Formular, da abrechenbare Vertretungen inkl. Korrektur unter 1:" & Chr(10) & ohneStunden
    End If
    If fehlerListe <> "" Then
        meldungLBV = meldungLBV & Chr(10) & Chr(10) & "Fehler bei:" & Chr(10) & fehlerListe
    End If
    MsgBox meldungLBV, IIf(fehlerListe = "", vbInformation, vbExclamation)
End Sub

' Hilfsfunktion: Umlaute fuer Dateinamen ersetzen
Private Function SafeFileName(s As String) As String
    Dim r As String: r = s
    r = Replace(r, Chr(196), "Ae")
    r = Replace(r, Chr(214), "Oe")
    r = Replace(r, Chr(220), "Ue")
    r = Replace(r, Chr(228), "ae")
    r = Replace(r, Chr(246), "oe")
    r = Replace(r, Chr(252), "ue")
    r = Replace(r, Chr(223), "ss")
    SafeFileName = r
End Function

' Hilfsfunktion: Text JSON-konform escapen.
' WICHTIG: ByVal - sonst wird die uebergebene Variable (z.B. dienststelle)
' bei jedem Aufruf veraendert und die Backslashes vermehren sich.
Private Function EscapeStr(ByVal s As String) As String
    Dim i As Long, c As Long
    Dim result As String
    For i = 1 To Len(s)
        c = AscW(Mid(s, i, 1))
        If c < 0 Then c = c + 65536
        Select Case c
            Case 92:  result = result & "\\"            ' Backslash
            Case 34:  result = result & "\" & Chr(34)   ' Anfuehrungszeichen
            Case 10:  result = result & "\n"            ' Zeilenumbruch
            Case 13:                                    ' CR ignorieren
            Case 9:   result = result & "\t"
            Case Is < 32, Is > 126
                result = result & "\u" & Right("0000" & Hex(c), 4)
            Case Else
                result = result & Mid(s, i, 1)
        End Select
    Next i
    EscapeStr = result
End Function


' ================================================================
'  MAKRO 7: Alle PDFs in einem Ordner drucken (via ShellExecute)
'  Drucker: Zeile "Druckeradresse" in Tabelle Informationen (Spalte B).
'  Erlaubt sind Druckername, Netzwerkpfad (\\server\drucker) oder
'  IP-Adresse eines auf diesem PC installierten Druckers.
'  Leer oder Zeile fehlt: Windows-Standarddrucker (wie bisher).
'  Druckprogramm: Zeile "Druckprogramm" in Tabelle Informationen:
'    leer / automatisch     -> SumatraPDF, falls gefunden, sonst Standard-PDF-Programm
'    SumatraPDF             -> immer SumatraPDF
'    Acrobat                -> Adobe Acrobat (Reader) direkt, egal welches
'                              Programm in Windows Standard ist (auch "Adobe")
'    C:\...\SumatraPDF.exe  -> diese SumatraPDF- bzw. Acrobat-Datei
'    Standard-PDF-Programm  -> immer das Standard-PDF-Programm von Windows
' ================================================================
Sub PDFsDrucken()

    ' Drucker aus Tabelle Informationen ermitteln
    Dim druckerEingabe As String: druckerEingabe = DruckerAusInformationen()
    Dim druckerName As String: druckerName = ""
    If druckerEingabe <> "" Then
        Dim druckerListe As String: druckerListe = ""
        druckerName = DruckerAufloesen(druckerEingabe, druckerListe)
        If druckerName = "" Then
            If MsgBox("Der Drucker aus Tabelle Informationen wurde auf diesem PC nicht gefunden:" & Chr(10) & _
                      "  " & druckerEingabe & Chr(10) & Chr(10) & _
                      "Installierte Drucker:" & Chr(10) & druckerListe & Chr(10) & _
                      "Stattdessen den Standarddrucker verwenden?", _
                      vbYesNo + vbExclamation, "Drucker nicht gefunden") = vbNo Then Exit Sub
        End If
    End If
    Dim druckerText As String
    druckerText = IIf(druckerName = "", "Standarddrucker", druckerName)

    ' Druckweg laut Zeile "Druckprogramm" in Tabelle Informationen.
    ' SumatraPDF druckt ohne Fenster direkt auf den gewuenschten Drucker,
    ' unabhaengig vom Standard-PDF-Programm (Edge kann nicht per Makro drucken).
    ' Standard-PDF-Programm: per ShellExecute, z.B. Adobe Acrobat (Reader).
    Dim druckProgramm As String: druckProgramm = InfoWertLesen("druckprogramm")
    Dim dpKurz As String
    dpKurz = LCase(Replace(Replace(Replace(druckProgramm, " ", ""), "-", ""), "_", ""))
    Dim sumatraPfad As String: sumatraPfad = ""
    Dim acrobatPfad As String: acrobatPfad = ""
    Dim dpTreffer As String
    Select Case True
        Case dpKurz = "" Or dpKurz = "automatisch" Or dpKurz = "auto"
            sumatraPfad = SumatraSuchen()

        Case dpKurz = "sumatra" Or dpKurz = "sumatrapdf"
            sumatraPfad = SumatraSuchen()
            If sumatraPfad = "" Then
                If MsgBox("In Tabelle Informationen ist als Druckprogramm SumatraPDF eingestellt," & Chr(10) & _
                          "es wurde aber keine SumatraPDF-Datei gefunden (gesucht u.a. im Ordner von lbv_fill.exe)." & Chr(10) & Chr(10) & _
                          "Stattdessen das Standard-PDF-Programm verwenden?", _
                          vbYesNo + vbExclamation, "PDFs drucken") = vbNo Then Exit Sub
            End If

        Case Right(dpKurz, 4) = ".exe"
            ' Pfad zu SumatraPDF oder Acrobat angegeben
            dpTreffer = ""
            On Error Resume Next
            dpTreffer = Dir(Trim(druckProgramm))
            On Error GoTo 0
            If dpTreffer = "" Then
                MsgBox "Das Druckprogramm aus Tabelle Informationen wurde nicht gefunden:" & Chr(10) & _
                       Trim(druckProgramm), vbExclamation, "PDFs drucken"
                Exit Sub
            End If
            If InStr(LCase(dpTreffer), "sumatra") > 0 Then
                sumatraPfad = Trim(druckProgramm)
            ElseIf InStr(LCase(dpTreffer), "acro") > 0 Then
                acrobatPfad = Trim(druckProgramm)
            Else
                MsgBox "Als Pfad bei 'Druckprogramm' werden nur SumatraPDF und" & Chr(10) & _
                       "Adobe Acrobat (Acrobat.exe / AcroRd32.exe) unterst" & Chr(252) & "tzt.", vbExclamation, "PDFs drucken"
                Exit Sub
            End If

        Case InStr(dpKurz, "acrobat") > 0 Or InStr(dpKurz, "adobe") > 0
            ' Acrobat bzw. Acrobat Reader DIREKT aufrufen - egal, welches Programm
            ' in Windows als Standard-PDF-Programm eingestellt ist
            acrobatPfad = AcrobatSuchen()
            If acrobatPfad = "" Then
                If MsgBox("In Tabelle Informationen ist als Druckprogramm Acrobat eingestellt," & Chr(10) & _
                          "Adobe Acrobat bzw. Acrobat Reader wurde auf diesem PC aber nicht gefunden." & Chr(10) & Chr(10) & _
                          "Stattdessen automatisch w" & Chr(228) & "hlen (SumatraPDF, sonst Standard-PDF-Programm)?", _
                          vbYesNo + vbExclamation, "PDFs drucken") = vbNo Then Exit Sub
                sumatraPfad = SumatraSuchen()
            End If

        Case dpKurz Like "standard*"
            sumatraPfad = ""            ' bewusst das Standard-PDF-Programm von Windows

        Case Else
            MsgBox "Unbekannter Eintrag bei 'Druckprogramm' in Tabelle Informationen:" & Chr(10) & _
                   "  " & druckProgramm & Chr(10) & Chr(10) & _
                   "Erlaubt sind:" & Chr(10) & _
                   "  (leer) oder automatisch" & Chr(10) & _
                   "  SumatraPDF" & Chr(10) & _
                   "  Acrobat" & Chr(10) & _
                   "  Standard-PDF-Programm" & Chr(10) & _
                   "  Pfad zur Programmdatei (C:\...\SumatraPDF.exe oder Acrobat.exe)", vbExclamation, "PDFs drucken"
            Exit Sub
    End Select
    Dim druckWegText As String
    If sumatraPfad <> "" Then
        druckWegText = "SumatraPDF" & Chr(10) & "  (" & sumatraPfad & ")"
    ElseIf acrobatPfad <> "" Then
        druckWegText = "Adobe Acrobat" & Chr(10) & "  (" & acrobatPfad & ")"
    Else
        druckWegText = "Standard-PDF-Programm"
    End If

    ' Acrobat braucht beim Drucken ohne Dialog (/t) immer einen Druckernamen
    Dim acroDrucker As String: acroDrucker = ""
    If acrobatPfad <> "" Then
        acroDrucker = IIf(druckerName <> "", druckerName, StandardDrucker())
        If acroDrucker = "" Then
            MsgBox "Der Windows-Standarddrucker konnte nicht ermittelt werden." & Chr(10) & _
                   "Bitte in Tabelle Informationen bei 'Druckeradresse' den Drucker eintragen.", _
                   vbExclamation, "PDFs drucken"
            Exit Sub
        End If
        druckerText = acroDrucker
    End If

    ' Ordner mit PDFs waehlen
    Dim druckOrdner As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Ordner mit PDFs zum Drucken w" & Chr(228) & "hlen"
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Kein Ordner gew" & Chr(228) & "hlt. Abbruch.", vbInformation
            Exit Sub
        End If
        druckOrdner = .SelectedItems(1)
    End With
    If Right(druckOrdner, 1) <> "\" Then druckOrdner = druckOrdner & "\"

    ' Alle PDF-Dateien im Ordner auflisten
    Dim dateiname As String
    Dim pdfDateien() As String
    Dim anzahl As Integer: anzahl = 0
    ReDim pdfDateien(0)

    dateiname = Dir(druckOrdner & "*.pdf")
    Do While dateiname <> ""
        ReDim Preserve pdfDateien(anzahl)
        pdfDateien(anzahl) = dateiname
        anzahl = anzahl + 1
        dateiname = Dir()
    Loop

    If anzahl = 0 Then
        MsgBox "Keine PDF-Dateien im gew" & Chr(228) & "hlten Ordner gefunden.", vbExclamation
        Exit Sub
    End If

    ' Bestaetigung mit Liste der Dateien
    Dim liste As String
    liste = "Folgende " & anzahl & " PDF(s) werden gedruckt:" & Chr(10) & Chr(10)
    Dim i As Integer
    For i = 0 To Application.Min(anzahl - 1, 9)
        liste = liste & "  " & pdfDateien(i) & Chr(10)
    Next i
    If anzahl > 10 Then liste = liste & "  ... und " & (anzahl - 10) & " weitere" & Chr(10)
    liste = liste & Chr(10) & "Drucker: " & druckerText & Chr(10) & _
            "Druck " & Chr(252) & "ber: " & druckWegText & Chr(10) & Chr(10) & "Fortfahren?"

    If MsgBox(liste, vbYesNo + vbQuestion, "PDFs drucken") = vbNo Then Exit Sub

    ' PDFs an den Drucker senden
    Dim gedruckt As Integer: gedruckt = 0
    Dim fehler As String:    fehler = ""
    ' Kennt das PDF-Programm "printto" nicht, wird der Windows-Standarddrucker
    ' voruebergehend auf den Zieldrucker umgestellt und am Ende zurueckgesetzt.
    Dim nurPrint As Boolean: nurPrint = (druckerName = "")
    Dim umgestellt As Boolean: umgestellt = False
    Dim alterStandard As String: alterStandard = ""

    Dim wsh As Object
    If sumatraPfad <> "" Or acrobatPfad <> "" Then Set wsh = CreateObject("WScript.Shell")

    On Error GoTo DruckFehler
    For i = 0 To anzahl - 1
        Dim vollPfad As String
        vollPfad = druckOrdner & pdfDateien(i)

        If sumatraPfad <> "" Then
            ' --- Weg 1: SumatraPDF (wartet, bis der Auftrag uebergeben ist) ---
            ' "shrink": nur verkleinern, falls der Drucker nicht randlos drucken kann
            '           (der LBV-Barcode liegt sehr nah am linken Rand)
            Dim sumCmd As String, sumExit As Long, sumLog As String, sumMeldung As String
            sumLog = Environ("TEMP") & "\lbv_druck_log.txt"
            On Error Resume Next
            Kill sumLog
            On Error GoTo DruckFehler
            sumCmd = Chr(34) & sumatraPfad & Chr(34) & " -silent " & _
                     IIf(druckerName = "", "-print-to-default", "-print-to " & Chr(34) & druckerName & Chr(34)) & _
                     " -print-settings " & Chr(34) & "shrink" & Chr(34) & " " & Chr(34) & vollPfad & Chr(34)
            ' Start ueber cmd (wie bei lbv_fill.exe): kein Windows-Startdialog fuer
            ' heruntergeladene Programme, und Fehlermeldungen von Windows landen im Log.
            ' Aeussere Anfuehrungszeichen: cmd /c entfernt das erste und letzte.
            sumExit = wsh.Run("cmd /c " & Chr(34) & sumCmd & " > " & Chr(34) & sumLog & Chr(34) & " 2>&1" & Chr(34), 0, True)
            If sumExit = 0 Then
                gedruckt = gedruckt + 1
            Else
                ' SumatraPDF schreibt mit -silent nichts. Steht etwas im Log, kam die
                ' Meldung von Windows: Programm startet gar nicht -> abbrechen.
                sumMeldung = TextDateiLesen(sumLog, 8)
                If sumMeldung <> "" Then
                    MsgBox "SumatraPDF konnte nicht gestartet werden (Code " & sumExit & "):" & Chr(10) & _
                           sumatraPfad & Chr(10) & Chr(10) & _
                           "Meldung von Windows:" & Chr(10) & sumMeldung & Chr(10) & _
                           "Tipp: Datei einmal per Doppelklick starten. Bei 'Computer wurde gesch" & Chr(252) & "tzt'" & Chr(10) & _
                           "bzw. Sicherheitswarnung: 'Weitere Informationen' > 'Trotzdem ausf" & Chr(252) & "hren'." & Chr(10) & _
                           "Oder: Rechtsklick > Eigenschaften > unten 'Zulassen' anhaken > OK.", _
                           vbExclamation, "PDFs drucken"
                    Exit Sub
                End If
                fehler = fehler & pdfDateien(i) & " (SumatraPDF-Fehlercode: " & sumExit & ")" & Chr(10)
            End If
            GoTo NaechstesPDF
        End If

        If acrobatPfad <> "" Then
            ' --- Weg 2: Adobe Acrobat (Reader) direkt ---
            ' /t = ohne Dialog auf den angegebenen Drucker, /s = ohne Startbild,
            ' /o = ohne Oeffnen-Dialog, /h = minimiert.
            ' Acrobat bleibt danach geoeffnet und druckt im Hintergrund -> nicht auf
            ' das Programmende warten (sonst haengt Excel), sondern kurz pausieren.
            Dim acroCmd As String
            acroCmd = Chr(34) & acrobatPfad & Chr(34) & " /s /o /h /t " & _
                      Chr(34) & vollPfad & Chr(34) & " " & Chr(34) & acroDrucker & Chr(34)
            wsh.Run acroCmd, 7, False          ' 7 = minimiert, ohne Fokus
            gedruckt = gedruckt + 1
            ' Erster Auftrag: Acrobat muss erst starten
            If i = 0 Then
                Application.Wait Now + TimeValue("00:00:06")
            Else
                Application.Wait Now + TimeValue("00:00:03")
            End If
            GoTo NaechstesPDF
        End If

        ' --- Weg 3: Standard-PDF-Programm per ShellExecute ---
        ' "printto" -> bestimmter Drucker (Druckername in Anfuehrungszeichen)
        ' "print"   -> Standarddrucker
        Dim result As LongPtr
        If Not nurPrint Then
            result = ShellExecute(0, "printto", vollPfad, Chr(34) & druckerName & Chr(34), druckOrdner, 0)   ' 0 = SW_HIDE
            ' 31 = SE_ERR_NOASSOC: PDF-Programm kennt "printto" nicht
            '      -> ab jetzt "print" mit voruebergehend umgestelltem Standarddrucker
            If result = 31 Then
                alterStandard = StandardDrucker()
                If LCase(alterStandard) <> LCase(druckerName) Then
                    If Not StandardDruckerSetzen(druckerName) Then
                        MsgBox "Der Drucker '" & druckerName & "' konnte nicht vor" & Chr(252) & "bergehend als " & _
                               "Standarddrucker eingestellt werden." & Chr(10) & "Es wurde nichts gedruckt.", _
                               vbExclamation, "PDFs drucken"
                        Exit Sub
                    End If
                    umgestellt = True
                End If
                nurPrint = True
            End If
        End If
        If nurPrint Then
            result = ShellExecute(0, "print", vollPfad, vbNullString, druckOrdner, 0)
        End If

        ' Auch "print" unbekannt: dieses PDF-Programm kann gar nicht per Makro drucken
        If result = 31 Then
            If umgestellt And alterStandard <> "" Then StandardDruckerSetzen alterStandard
            MsgBox "Das Standard-PDF-Programm (z.B. Edge) unterst" & Chr(252) & "tzt kein Drucken per Makro." & Chr(10) & Chr(10) & _
                   "Abhilfe (eine davon gen" & Chr(252) & "gt):" & Chr(10) & _
                   "- SumatraPDF (kostenlos, portable Version) als SumatraPDF.exe" & Chr(10) & _
                   "  in den Ordner von lbv_fill.exe legen" & Chr(10) & _
                   "- Adobe Acrobat Reader als Standard-PDF-Programm einstellen", vbExclamation, "PDFs drucken"
            Exit Sub
        End If

        If result > 32 Then
            gedruckt = gedruckt + 1
        Else
            fehler = fehler & pdfDateien(i) & " (Fehlercode: " & result & ")" & Chr(10)
        End If

        ' Kurze Pause damit der Drucker-Spooler nicht ueberlastet wird
        Application.Wait Now + TimeValue("00:00:02")
NaechstesPDF:
    Next i
    On Error GoTo 0

    ' Abschlussmeldung
    Dim meldung As String
    If fehler = "" Then
        meldung = gedruckt & " PDF(s) wurden an den Drucker gesendet:" & Chr(10) & druckerText
    Else
        meldung = gedruckt & " PDF(s) gesendet an: " & druckerText & Chr(10) & Chr(10) & _
                  "Fehler bei:" & Chr(10) & fehler
    End If
    If acrobatPfad <> "" Then
        meldung = meldung & Chr(10) & Chr(10) & _
                  "Acrobat druckt im Hintergrund weiter und bleibt danach ge" & Chr(246) & "ffnet." & Chr(10) & _
                  "Es kann geschlossen werden, sobald alles gedruckt ist."
    End If

    If umgestellt Then
        ' Erst zuruecksetzen, wenn das PDF-Programm alle Auftraege uebernommen hat -
        ' sonst landen die letzten Dokumente evtl. noch auf dem alten Standarddrucker
        MsgBox meldung & Chr(10) & Chr(10) & _
               "Daf" & Chr(252) & "r wurde der Windows-Standarddrucker vor" & Chr(252) & "bergehend umgestellt." & Chr(10) & _
               "Bitte erst auf OK klicken, wenn alle Dokumente in der Druckwarteschlange" & Chr(10) & _
               "stehen bzw. gedruckt werden. Danach wird der bisherige Standarddrucker" & Chr(10) & _
               "wieder eingestellt: " & IIf(alterStandard <> "", alterStandard, "(unbekannt)"), _
               IIf(fehler = "", vbInformation, vbExclamation), "PDFs drucken"
        If alterStandard = "" Then
            MsgBox "Der bisherige Standarddrucker konnte nicht ermittelt werden." & Chr(10) & _
                   "Bitte in den Windows-Einstellungen wieder den gew" & Chr(252) & "nschten Standarddrucker w" & Chr(228) & "hlen.", _
                   vbExclamation, "PDFs drucken"
        ElseIf Not StandardDruckerSetzen(alterStandard) Then
            MsgBox "Der bisherige Standarddrucker konnte nicht wieder eingestellt werden:" & Chr(10) & alterStandard & Chr(10) & _
                   "Bitte in den Windows-Einstellungen pr" & Chr(252) & "fen.", vbExclamation, "PDFs drucken"
        End If
    Else
        MsgBox meldung, IIf(fehler = "", vbInformation, vbExclamation), "PDFs drucken"
    End If
    Exit Sub

DruckFehler:
    ' Fehlertext ZUERST sichern - StandardDruckerSetzen setzt Err zurueck
    Dim fehlerText As String
    fehlerText = "Fehler beim Drucken (" & Err.Number & "):" & Chr(10) & Err.Description
    If sumatraPfad <> "" Then fehlerText = fehlerText & Chr(10) & Chr(10) & "Programm: " & sumatraPfad
    If umgestellt And alterStandard <> "" Then StandardDruckerSetzen alterStandard
    MsgBox fehlerText, vbCritical, "PDFs drucken"
End Sub

' ----------------------------------------------------------------
' Hilfsfunktion: aktuellen Windows-Standarddrucker ermitteln ("" = unbekannt)
' ----------------------------------------------------------------
Private Function StandardDrucker() As String
    Dim wmi As Object, liste As Object, d As Object
    Dim nm As String
    StandardDrucker = ""
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If wmi Is Nothing Then Exit Function
    Set liste = wmi.ExecQuery("SELECT Name FROM Win32_Printer WHERE Default = TRUE")
    If liste Is Nothing Then Exit Function
    For Each d In liste
        nm = "": nm = d.Name & ""
        If nm <> "" Then
            StandardDrucker = nm
            Exit For
        End If
    Next d
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Windows-Standarddrucker setzen (True = erfolgreich)
' ----------------------------------------------------------------
Private Function StandardDruckerSetzen(ByVal zielDrucker As String) As Boolean
    Dim net As Object
    StandardDruckerSetzen = False
    On Error Resume Next
    Set net = CreateObject("WScript.Network")
    If net Is Nothing Then Exit Function
    Err.Clear
    net.SetDefaultPrinter zielDrucker
    StandardDruckerSetzen = (Err.Number = 0)
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: SumatraPDF suchen ("" = nicht gefunden)
' Reihenfolge: Ordner von lbv_fill.exe, Ordner dieser Arbeitsmappe,
' Standard-Installationsorte. Dateiname "SumatraPDF*.exe" (auch mit
' Versionsnummer, z.B. SumatraPDF-3.5.2-64.exe), aber kein Installer.
' ----------------------------------------------------------------
Private Function SumatraSuchen() As String
    Dim kandidaten As String, exePfad As String
    Dim ordner() As String, i As Long, treffer As String

    SumatraSuchen = ""
    kandidaten = ""
    exePfad = InfoWertLesen("pfad lbv_fill.exe")
    If InStrRev(exePfad, "\") > 0 Then kandidaten = kandidaten & "|" & Left(exePfad, InStrRev(exePfad, "\"))
    If ThisWorkbook.Path <> "" Then kandidaten = kandidaten & "|" & ThisWorkbook.Path & "\"
    If Environ("LOCALAPPDATA") <> "" Then kandidaten = kandidaten & "|" & Environ("LOCALAPPDATA") & "\SumatraPDF\"
    If Environ("ProgramFiles") <> "" Then kandidaten = kandidaten & "|" & Environ("ProgramFiles") & "\SumatraPDF\"
    If Environ("ProgramW6432") <> "" Then kandidaten = kandidaten & "|" & Environ("ProgramW6432") & "\SumatraPDF\"
    If Environ("ProgramFiles(x86)") <> "" Then kandidaten = kandidaten & "|" & Environ("ProgramFiles(x86)") & "\SumatraPDF\"
    If kandidaten = "" Then Exit Function

    ordner = Split(Mid(kandidaten, 2), "|")
    On Error Resume Next            ' Dir scheitert z.B. bei OneDrive-URLs
    For i = 0 To UBound(ordner)
        treffer = ""
        treffer = Dir(ordner(i) & "SumatraPDF*.exe")
        Do While treffer <> ""
            If InStr(LCase(treffer), "install") = 0 Then
                SumatraSuchen = ordner(i) & treffer
                Exit Function
            End If
            treffer = ""
            treffer = Dir()
        Loop
    Next i
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Adobe Acrobat bzw. Acrobat Reader suchen ("" = nicht gefunden)
' 1) registrierte Programmpfade (App Paths), 2) uebliche Installationsordner.
' ProgramW6432 = 64-Bit-Programmordner, auch wenn Office 32-bittig ist.
' ----------------------------------------------------------------
Private Function AcrobatSuchen() As String
    Dim wsh As Object, kandidaten As String, wert As String, treffer As String
    Dim regPfade(5) As String, ordner(2) As String, teile() As String
    Dim i As Long, j As Long

    AcrobatSuchen = ""
    kandidaten = ""
    On Error Resume Next

    regPfade(0) = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe\"
    regPfade(1) = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe\"
    regPfade(2) = "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe\"
    regPfade(3) = "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe\"
    regPfade(4) = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe\"
    regPfade(5) = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe\"
    Set wsh = CreateObject("WScript.Shell")
    If Not wsh Is Nothing Then
        For i = 0 To 5
            wert = ""
            wert = wsh.RegRead(regPfade(i))      ' Fehler = Eintrag fehlt -> bleibt ""
            wert = Trim(Replace(wert, Chr(34), ""))
            If wert <> "" Then kandidaten = kandidaten & "|" & wert
        Next i
    End If

    ordner(0) = Environ("ProgramW6432")
    ordner(1) = Environ("ProgramFiles")
    ordner(2) = Environ("ProgramFiles(x86)")
    For j = 0 To 2
        If ordner(j) <> "" Then
            kandidaten = kandidaten & "|" & ordner(j) & "\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
            kandidaten = kandidaten & "|" & ordner(j) & "\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
            kandidaten = kandidaten & "|" & ordner(j) & "\Adobe\Acrobat Reader\Reader\AcroRd32.exe"
        End If
    Next j
    If kandidaten = "" Then Exit Function

    teile = Split(Mid(kandidaten, 2), "|")
    For i = 0 To UBound(teile)
        treffer = ""
        treffer = Dir(teile(i))
        If treffer <> "" Then
            AcrobatSuchen = teile(i)
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Wert aus Tabelle Informationen lesen
' (Spalte A = Beschriftung, Gross-/Kleinschreibung und ":" egal)
' ----------------------------------------------------------------
Private Function InfoWertLesen(ByVal schluessel As String) As String
    Dim wsInfo As Worksheet, r As Long, k As String
    InfoWertLesen = ""
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Worksheets("Informationen")
    On Error GoTo 0
    If wsInfo Is Nothing Then Exit Function
    For r = 1 To wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
        k = LCase(Trim(CStr(wsInfo.Cells(r, 1).Value)))
        If Right(k, 1) = ":" Then k = Trim(Left(k, Len(k) - 1))
        If k = LCase(schluessel) Then
            InfoWertLesen = Trim(CStr(wsInfo.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Textdatei lesen (max. maxZeilen nicht-leere Zeilen)
' "" wenn die Datei fehlt oder leer ist
' ----------------------------------------------------------------
Private Function TextDateiLesen(ByVal pfad As String, ByVal maxZeilen As Long) As String
    Dim f As Integer, zeile As String, txt As String, n As Long
    Dim vorhanden As String
    TextDateiLesen = ""
    On Error Resume Next
    vorhanden = ""
    vorhanden = Dir(pfad)
    If vorhanden = "" Then Exit Function
    f = FreeFile
    Open pfad For Input As #f
    If Err.Number <> 0 Then Exit Function
    Do While Not EOF(f)
        Line Input #f, zeile
        If Trim(zeile) <> "" And n < maxZeilen Then txt = txt & Trim(zeile) & Chr(10): n = n + 1
    Loop
    Close #f
    TextDateiLesen = txt
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Eintrag "Druckeradresse" aus Tabelle Informationen
' (Spalte A = Beschriftung, Spalte B = Wert). "" wenn leer/fehlend.
' ----------------------------------------------------------------
Private Function DruckerAusInformationen() As String
    Dim wsInfo As Worksheet
    Dim r As Long
    Dim schluessel As String
    DruckerAusInformationen = ""
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Worksheets("Informationen")
    On Error GoTo 0
    If wsInfo Is Nothing Then Exit Function
    For r = 1 To wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
        schluessel = LCase(Trim(CStr(wsInfo.Cells(r, 1).Value)))
        If Right(schluessel, 1) = ":" Then schluessel = Trim(Left(schluessel, Len(schluessel) - 1))
        Select Case schluessel
            Case "druckeradresse", "drucker-adresse", "drucker", "druckername", "drucker-ip", "drucker ip"
                DruckerAusInformationen = Trim(CStr(wsInfo.Cells(r, 2).Value))
                Exit Function
        End Select
    Next r
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Eingabe (Druckername, \\server\drucker oder IP-Adresse)
' einem auf diesem PC installierten Drucker zuordnen.
' Rueckgabe: Windows-Druckername oder "" (nicht gefunden).
' liste: alle installierten Drucker (fuer die Fehlermeldung)
' ----------------------------------------------------------------
Private Function DruckerAufloesen(ByVal eingabe As String, ByRef liste As String) As String
    Dim suche As String, sucheAdr As String
    Dim wmi As Object, ports As Object, port As Object, alle As Object, d As Object
    Dim portTreffer As String
    Dim nm As String, freigabe As String, server As String, portListe As String
    Dim teile() As String, p As Long, pt As String
    Dim treffer As String

    DruckerAufloesen = ""
    liste = ""
    suche = LCase(Trim(eingabe))
    ' Adresse ohne "http://", "https://" und abschliessendes "/"
    sucheAdr = Replace(Replace(suche, "https://", ""), "http://", "")
    If Right(sucheAdr, 1) = "/" Then sucheAdr = Left(sucheAdr, Len(sucheAdr) - 1)

    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If wmi Is Nothing Then
        ' WMI nicht verfuegbar: Eingabe unveraendert als Druckername verwenden
        DruckerAufloesen = Trim(eingabe)
        Exit Function
    End If

    ' 1) Standard-TCP/IP-Ports mit dieser Adresse sammeln
    portTreffer = "|"
    Set ports = wmi.ExecQuery("SELECT Name, HostAddress FROM Win32_TCPIPPrinterPort")
    If Not ports Is Nothing Then
        For Each port In ports
            nm = "": nm = LCase(Trim(port.Name & ""))
            pt = "": pt = LCase(Trim(port.HostAddress & ""))
            If pt = sucheAdr Or nm = sucheAdr Then portTreffer = portTreffer & nm & "|"
        Next port
    End If

    ' 2) Installierte Drucker durchsuchen
    treffer = ""
    Set alle = wmi.ExecQuery("SELECT Name, ShareName, SystemName, PortName FROM Win32_Printer")
    If Not alle Is Nothing Then
        For Each d In alle
            nm = "": nm = Trim(d.Name & "")
            freigabe = "": freigabe = LCase(Trim(d.ShareName & ""))
            server = "": server = LCase(Trim(d.SystemName & ""))
            portListe = "": portListe = LCase(d.PortName & "")
            liste = liste & "  " & nm & vbLf

            If treffer = "" Then
                ' a) exakter Druckername (Netzwerkdrucker heissen "\\server\freigabe")
                If LCase(nm) = suche Then treffer = nm
                ' b) Freigabename oder \\server\freigabe
                If treffer = "" And freigabe <> "" Then
                    If freigabe = suche Then treffer = nm
                    If server <> "" Then
                        If "\\" & Replace(server, "\\", "") & "\" & freigabe = suche Then treffer = nm
                    End If
                End If
                ' c) Anschluss (Port) mit der IP-Adresse, z.B. "192.168.1.50" oder "IP_192.168.1.50"
                If treffer = "" Then
                    teile = Split(portListe, ",")
                    For p = 0 To UBound(teile)
                        pt = Trim(teile(p))
                        If pt <> "" Then
                            If pt = sucheAdr Or pt = "ip_" & sucheAdr Or _
                               InStr(portTreffer, "|" & pt & "|") > 0 Then
                                treffer = nm
                                Exit For
                            End If
                        End If
                    Next p
                End If
            End If
        Next d
    End If
    On Error GoTo 0

    DruckerAufloesen = treffer
End Function

' ============================================================
' MAKRO: Mehrarbeit-PDFs per E-Mail versenden (Outlook)
' Absender: Zeile "Outlook-Absender" in Tabelle Informationen (Spalte B).
'   Leer oder Zeile fehlt -> Outlook-Standardkonto.
'   Adresse in Outlook nicht eingerichtet -> Abbruch vor dem Versand.
' ============================================================
Public Sub MehrarbeitPDFsVersenden()
    Dim ws3 As Worksheet
    Set ws3 = Nothing
    Dim wbName As String
    wbName = ThisWorkbook.Name

    ' Tabelle Uebersicht finden
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(LCase(ws.Name), "bersicht") > 0 Then
            Set ws3 = ws
            Exit For
        End If
    Next ws
    If ws3 Is Nothing Then
        MsgBox "Tabelle " & Chr(220) & "bersicht nicht gefunden.", vbCritical: Exit Sub
    End If

    ' Spalten ermitteln
    Dim colKuerzel As Integer, colEmail As Integer
    colKuerzel = 1  ' Spalte A = Kuerzel
    colEmail = SpalteUebersicht(ws3, "emailadresse")
    If colEmail = 0 Then
        MsgBox "Spalte 'emailadresse' nicht in " & Chr(220) & "bersicht gefunden.", vbCritical: Exit Sub
    End If

    ' Ordner mit PDFs auswaehlen
    Dim pdfOrdner As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Ordner mit Mehrarbeit-PDFs w" & Chr(228) & "hlen"
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Kein Ordner gew" & Chr(228) & "hlt. Abbruch.", vbInformation: Exit Sub
        End If
        pdfOrdner = .SelectedItems(1)
    End With
    If Right(pdfOrdner, 1) <> "" Then pdfOrdner = pdfOrdner & ""

    ' Outlook-Objekt erstellen
    Dim olApp As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        MsgBox "Outlook konnte nicht gestartet werden.", vbCritical: Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row

    ' Absenderkonto aus Tabelle Informationen (Zeile "Outlook-Absender")
    Dim absenderAdr As String
    absenderAdr = InfoWertLesen("outlook-absender")

    Dim zielKonto As Object
    Dim olAccount As Object
    Dim kontenListe As String: kontenListe = ""
    If absenderAdr <> "" Then
        For Each olAccount In olApp.Session.Accounts
            kontenListe = kontenListe & "  " & olAccount.SmtpAddress & Chr(10)
            If LCase(Trim(olAccount.SmtpAddress)) = LCase(absenderAdr) Then
                Set zielKonto = olAccount
                Exit For
            End If
        Next olAccount
        If zielKonto Is Nothing Then
            MsgBox "Das Absenderkonto aus Tabelle Informationen (Zeile 'Outlook-Absender')" & Chr(10) & _
                   "ist in Outlook nicht eingerichtet:" & Chr(10) & "  " & absenderAdr & Chr(10) & Chr(10) & _
                   "Vorhandene Outlook-Konten:" & Chr(10) & kontenListe & Chr(10) & _
                   "Es wurde nichts gesendet.", vbExclamation, "E-Mail-Versand"
            Exit Sub
        End If
    End If

    Dim absenderText As String
    If zielKonto Is Nothing Then
        absenderText = "Outlook-Standardkonto (Zeile 'Outlook-Absender' ist leer)"
    Else
        absenderText = absenderAdr
    End If

    ' Rueckfrage: die Mails werden ohne Vorschau sofort gesendet
    If MsgBox("Die E-Mails werden ohne Vorschau sofort gesendet." & Chr(10) & Chr(10) & _
              "Absender:  " & absenderText & Chr(10) & _
              "PDF-Ordner:  " & pdfOrdner & Chr(10) & Chr(10) & _
              "Versand starten?", vbOKCancel + vbQuestion, "E-Mail-Versand") = vbCancel Then Exit Sub

    Dim gesendet As Integer, keinEmail As Integer, keinPDF As Integer
    gesendet = 0: keinEmail = 0: keinPDF = 0

    ' Alle Lehrer durchgehen
    Dim r As Long
    For r = 2 To lastRow
        Dim kuerzel As String
        kuerzel = Trim(ws3.Cells(r, colKuerzel).Value)
        If kuerzel = "" Then GoTo NaechsterLehrer

        Dim emailAdr As String
        emailAdr = Trim(ws3.Cells(r, colEmail).Value)
        If emailAdr = "" Then
            keinEmail = keinEmail + 1
            GoTo NaechsterLehrer
        End If

        ' Alle passenden PDFs finden: Mehrarbeit_<kuerzel>_*.pdf
        Dim suchMuster As String
        If Right(pdfOrdner, 1) <> "\" Then pdfOrdner = pdfOrdner & "\"
        suchMuster = pdfOrdner & "Mehrarbeit_" & kuerzel & "_*.pdf"
        Dim gefundenePDF As String
        gefundenePDF = Dir(suchMuster)

        If gefundenePDF = "" Then
            keinPDF = keinPDF + 1
            GoTo NaechsterLehrer
        End If

        ' E-Mail erstellen und versenden
        Do While gefundenePDF <> ""
            Dim vollPfad2 As String
            vollPfad2 = pdfOrdner & gefundenePDF

            Dim mail As Object
            Set mail = olApp.CreateItem(0)
            With mail
                .To = emailAdr
                If Not zielKonto Is Nothing Then Set .SendUsingAccount = zielKonto
                .Subject = "Mehrarbeitsverg" & Chr(252) & "tung - " & gefundenePDF
                .Body = "Sehr geehrte Damen und Herren," & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "anbei erhalten Sie Ihre Mehrarbeitsverg" & Chr(252) & "tungsabrechnung." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                        "Mit freundlichen Gr" & Chr(252) & Chr(223) & "en"
                .Attachments.Add vollPfad2
                .Send
            End With
            Set mail = Nothing
            gesendet = gesendet + 1

            gefundenePDF = Dir()
        Loop

NaechsterLehrer:
    Next r

    MsgBox "Fertig!" & Chr(10) & _
           "Absender: " & absenderText & Chr(10) & _
           "Gesendet: " & gesendet & Chr(10) & _
           "Keine E-Mail-Adresse: " & keinEmail & Chr(10) & _
           "Keine PDF gefunden: " & keinPDF, vbInformation, "E-Mail-Versand"
End Sub

' ============================================================
' MAKRO: Mehrarbeit-PDFs per SMTP versenden
' ============================================================
Public Sub MehrarbeitPDFsVersendenSMTP()

    ' Tabelle Informationen finden
    Dim wsInfo As Worksheet
    Set wsInfo = Nothing
    Dim wsI As Worksheet
    For Each wsI In ThisWorkbook.Worksheets
        If LCase(wsI.Name) Like "*informationen*" Then
            Set wsInfo = wsI: Exit For
        End If
    Next wsI
    If wsInfo Is Nothing Then
        MsgBox "Tabelle Informationen nicht gefunden.", vbCritical: Exit Sub
    End If

    ' SMTP-Parameter aus Tabelle Informationen lesen
    Dim smtpServer As String, smtpPort As String
    Dim verschl As String, benutzer As String, passwort As String
    Dim exePfadFill As String

    Dim lastInfoRow As Long
    lastInfoRow = wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
    Dim iRow As Long
    For iRow = 1 To lastInfoRow
        Dim label As String
        label = LCase(Trim(wsInfo.Cells(iRow, 1).Value))
        Select Case label
            Case "smtp-server:", "smtp-server"
                smtpServer = Trim(wsInfo.Cells(iRow, 2).Value)
            Case "smtp-port:", "smtp-port"
                smtpPort = Trim(wsInfo.Cells(iRow, 2).Value)
            Case "verschl" & Chr(252) & "sselung:", "verschl" & Chr(252) & "sselung"
                verschl = Trim(wsInfo.Cells(iRow, 2).Value)
            Case "benutzername:", "benutzername"
                benutzer = Trim(wsInfo.Cells(iRow, 2).Value)
            Case "passwort:", "passwort"
                passwort = Trim(wsInfo.Cells(iRow, 2).Value)
            Case "pfad lbv_fill.exe"
                exePfadFill = Trim(wsInfo.Cells(iRow, 2).Value)
        End Select
    Next iRow

    ' Pflichtfelder pruefen
    If smtpServer = "" Or smtpPort = "" Or benutzer = "" Or passwort = "" Then
        MsgBox "SMTP-Parameter unvollst" & Chr(228) & "ndig. Bitte Tabelle Informationen pr" & _
               Chr(252) & "fen (Zeilen SMTP-Server, SMTP-Port, Benutzername, Passwort).", _
               vbCritical, "SMTP-Konfiguration fehlt"
        Exit Sub
    End If

    ' Pfad lbv_email.exe ableiten
    Dim emailExe As String
    emailExe = Replace(exePfadFill, "lbv_fill.exe", "lbv_email.exe")
    If Len(Dir(emailExe)) = 0 Then
        MsgBox "lbv_email.exe nicht gefunden:" & Chr(10) & emailExe, vbCritical, "EXE fehlt"
        Exit Sub
    End If

    ' Tabelle Uebersicht finden
    Dim ws3 As Worksheet
    Set ws3 = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(LCase(ws.Name), "bersicht") > 0 Then
            Set ws3 = ws: Exit For
        End If
    Next ws
    If ws3 Is Nothing Then
        MsgBox "Tabelle " & Chr(220) & "bersicht nicht gefunden.", vbCritical: Exit Sub
    End If

    ' E-Mail-Spalte finden
    Dim colEmail As Integer
    colEmail = SpalteUebersicht(ws3, "emailadresse")
    If colEmail = 0 Then
        MsgBox "Spalte 'emailadresse' nicht in " & Chr(220) & "bersicht gefunden.", vbCritical: Exit Sub
    End If

    ' Ordner mit PDFs auswaehlen
    Dim pdfOrdner As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Ordner mit Mehrarbeit-PDFs w" & Chr(228) & "hlen"
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Kein Ordner gew" & Chr(228) & "hlt. Abbruch.", vbInformation: Exit Sub
        End If
        pdfOrdner = .SelectedItems(1)
    End With
    If Right(pdfOrdner, 1) <> "" Then pdfOrdner = pdfOrdner & ""

    ' Mails-Liste aufbauen
    Dim lastRow As Long
    lastRow = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row

    Dim mailsJson As String
    mailsJson = ""
    Dim anzMails As Integer
    anzMails = 0

    Dim r As Long
    For r = 2 To lastRow
        Dim kuerzel As String
        kuerzel = Trim(ws3.Cells(r, 1).Value)
        If kuerzel = "" Then GoTo NaechsterSMTP

        Dim emailAdr As String
        emailAdr = Trim(ws3.Cells(r, colEmail).Value)
        If emailAdr = "" Then GoTo NaechsterSMTP

        ' Passende PDFs suchen
        If Right(pdfOrdner, 1) <> "" Then pdfOrdner = pdfOrdner & ""
        Dim suchMuster As String
        If Right(pdfOrdner, 1) <> "\" Then pdfOrdner = pdfOrdner & "\"
        suchMuster = pdfOrdner & "Mehrarbeit_" & kuerzel & "_*.pdf"
        Dim gefunden As String
        gefunden = Dir(suchMuster)

        Do While gefunden <> ""
            Dim vollPfad As String
            vollPfad = Replace(pdfOrdner & gefunden, Chr(92), "/")

            If mailsJson <> "" Then mailsJson = mailsJson & "," & Chr(10)
            mailsJson = mailsJson & "    {" & Chr(10) & _
                "      " & Chr(34) & "an" & Chr(34) & ": " & Chr(34) & emailAdr & Chr(34) & "," & Chr(10) & _
                "      " & Chr(34) & "pdf_pfad" & Chr(34) & ": " & Chr(34) & vollPfad & Chr(34) & "," & Chr(10) & _
                "      " & Chr(34) & "dateiname" & Chr(34) & ": " & Chr(34) & gefunden & Chr(34) & Chr(10) & _
                "    }"
            anzMails = anzMails + 1
            gefunden = Dir()
        Loop

NaechsterSMTP:
    Next r

    If anzMails = 0 Then
        MsgBox "Keine PDFs gefunden.", vbInformation: Exit Sub
    End If

    ' JSON schreiben
    Dim jsonPfad As String
    jsonPfad = Environ("TEMP") & "\lbv_email.json"

    Dim fNr As Integer
    fNr = FreeFile
    Open jsonPfad For Output As #fNr
    Print #fNr, "{"
    Print #fNr, "  " & Chr(34) & "smtp_server" & Chr(34) & ": " & Chr(34) & smtpServer & Chr(34) & ","
    Print #fNr, "  " & Chr(34) & "smtp_port" & Chr(34) & ": " & smtpPort & ","
    Print #fNr, "  " & Chr(34) & "verschluesselung" & Chr(34) & ": " & Chr(34) & verschl & Chr(34) & ","
    Print #fNr, "  " & Chr(34) & "benutzername" & Chr(34) & ": " & Chr(34) & benutzer & Chr(34) & ","
    Print #fNr, "  " & Chr(34) & "passwort" & Chr(34) & ": " & Chr(34) & passwort & Chr(34) & ","
    Print #fNr, "  " & Chr(34) & "absender" & Chr(34) & ": " & Chr(34) & benutzer & Chr(34) & ","
    Print #fNr, "  " & Chr(34) & "mails" & Chr(34) & ": ["
    Print #fNr, mailsJson
    Print #fNr, "  ]"
    Print #fNr, "}"
    Close #fNr

    ' lbv_email.exe ausfuehren
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim logPfad As String
    logPfad = Environ("TEMP") & "\lbv_email_log.txt"

    Dim exitCode As Long
    ' EXE direkt aufrufen
    exitCode = wsh.Run(Chr(34) & emailExe & Chr(34) & " " & Chr(34) & jsonPfad & Chr(34), 0, True)

    If exitCode = 0 Then
        MsgBox "E-Mail-Versand abgeschlossen.", vbInformation, "Fertig"
    Else
        MsgBox "Fehler beim Versand (Exit-Code: " & exitCode & ")", vbCritical, "E-Mail-Fehler"
    End If

End Sub
