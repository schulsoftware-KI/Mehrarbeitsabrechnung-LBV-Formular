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

    Dim lastRow As Long
    lastRow = wsUe.Cells(wsUe.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim(wsUe.Cells(r, 1).Value) <> "" Then
            Dim summe As Double
            Dim bagatell As Double
            summe = 0: bagatell = 0
            If colSumme > 0 And IsNumeric(wsUe.Cells(r, colSumme).Value) Then
                summe = CDbl(wsUe.Cells(r, colSumme).Value)
            End If
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
            Case "Vertretungen":                                               colVert = c
            Case "nicht zu z" & Chr(228) & "hlende Vertretungen":             colVertN = c
        End Select
    Next c

    If colSumme = 0 Then lastColUe = lastColUe + 1: colSumme = lastColUe: ws3.Cells(1, colSumme).Value = "Summenwert"
    If colEntfG = 0 Then lastColUe = lastColUe + 1: colEntfG = lastColUe: ws3.Cells(1, colEntfG).Value = "Entf" & Chr(228) & "lle gesamt"
    If colEntfN = 0 Then lastColUe = lastColUe + 1: colEntfN = lastColUe: ws3.Cells(1, colEntfN).Value = "nicht zu z" & Chr(228) & "hlende Entf" & Chr(228) & "lle"
    If colVert = 0 Then lastColUe = lastColUe + 1: colVert = lastColUe: ws3.Cells(1, colVert).Value = "Vertretungen"
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
    colVert = SpalteUebersicht(wsUe, "Vertretungen")
    colVertN = SpalteUebersicht(wsUe, "nicht zu z" & Chr(228) & "hlende Vertretungen")
    colBagatell = SpalteUebersicht(wsUe, "Bagatellgrenze")
    Dim colPersNr As Integer
    colPersNr = SpalteUebersicht(wsUe, "Pers.Nr")

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

        Call ErstellePDF(speicherOrdner, kuerzel, vollname, persNr, zeitbereich, jahreszahl, _
            detailZeilen, anzDetail, summeWert, entfGesamt, entfNicht, vertWert, vertNicht, bagatell)
        pdfErstellt = pdfErstellt + 1

        ' Blatt fuer Sammel-PDF merken (wird nach der Schleife zusammengefuehrt)
        ReDim Preserve sammelBlaetter(pdfErstellt - 1)
        sammelBlaetter(pdfErstellt - 1) = kuerzel

NaechsterLehrer:
    Next r

    ' Sammel-PDF aus allen erstellten Einzelblaettern erzeugen
    If pdfErstellt > 1 Then
        Call ErstelleSammelPDF(speicherOrdner, sammelBlaetter, pdfErstellt, jahreszahl, zeitbereich, _
            ws1, wsUe, colSumme, colBagatell, colEntfG, colEntfN, colVert, colVertN, colPersNr, lastRowUe, lastRow1)
    End If

    MsgBox pdfErstellt & " PDF(s) erstellt in:" & Chr(10) & speicherOrdner, vbInformation
End Sub

' ================================================================
'  PDF eines Lehrers erstellen
' ================================================================
Private Sub ErstellePDF(ordner As String, kuerzel As String, vollname As String, _
    persNr As String, zeitbereich As String, jahreszahl As String, detailZeilen() As String, anzDetail As Integer, _
    summeWert As Double, entfGesamt As Long, entfNicht As Long, _
    vertWert As Long, vertNicht As Long, bagatell As Double)

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

    ' Kennzahlen
    Dim bagatellText As String
    bagatellText = IIf(bagatell > 0, Format(bagatell, "0.00"), "k.A.")
    wsTmp.Cells(z, 1).Value = "Bagatellgrenze:  " & bagatellText
    z = z + 1
    wsTmp.Cells(z, 1).Value = "Entf" & Chr(228) & "lle gesamt:  " & entfGesamt & _
        "          nicht zu z" & Chr(228) & "hlende Entf" & Chr(228) & "lle:  " & entfNicht
    z = z + 1
    wsTmp.Cells(z, 1).Value = "Vertretungen:  " & vertWert & _
        "          nicht zu z" & Chr(228) & "hlende Vertretungen:  " & vertNicht
    z = z + 1
    With wsTmp.Cells(z, 1)
        .Value = "Anzahl abrechenbare Vertretungen:  " & summeWert
        .Font.Bold = True
        If bagatell > 0 And summeWert > bagatell Then .Font.Color = RGB(0, 128, 0)
    End With
    z = z + 1

    ' Abstand vor "Zur Kenntnis genommen"
    z = z + 4

    ' Zur Kenntnis genommen
    wsTmp.Cells(z, 1).Value = "Zur Kenntnis genommen:": wsTmp.Cells(z, 1).Font.Bold = True
    z = z + 1

    ' Abstand nach "Zur Kenntnis genommen" (Unterschrift-Zeilen)
    z = z + 4

    ' Unterschriftszeilen
    wsTmp.Cells(z, 1).Value = "________________________________"
    wsTmp.Cells(z, 9).Value = "________________________________"
    wsTmp.Cells(z, 9).HorizontalAlignment = xlRight
    z = z + 1
    wsTmp.Cells(z, 1).Value = "Datum / Unterschrift Lehrkraft"
    wsTmp.Cells(z, 9).Value = "Datum / Unterschrift Schulleitung"
    wsTmp.Cells(z, 9).HorizontalAlignment = xlRight

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
    skip(4) = "Vert": skip(5) = "Anz":  skip(6) = "Zur": skip(7) = "Dat"
    skip(8) = "___"

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
' ================================================================
Sub ButtonsErstellen()

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

    ' Tabelle Informationen (blau)
    btnDefs(btnCount) = "Informationen|Tabelle1 importieren|Tabelle1Importieren":                                     btnCount = btnCount + 1
    btnDefs(btnCount) = "Informationen|Summen einlesen|SummenInTabelleUebersicht":                                    btnCount = btnCount + 1
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
        If zielBlatt <> letztesSHeet Then
            btnLeft = 4
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
    colPersNr As Integer, lastRowUe As Long, lastRow1 As Long)

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
        Dim ueZ As Long
        For ueZ = 2 To lastRowUe
            If LCase(Trim(wsUe.Cells(ueZ, 1).Value)) = LCase(kuerzel) Then
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

        ' Kennzahlen
        Dim bagatellText As String
        bagatellText = IIf(bagatell > 0, Format(bagatell, "0.00"), "k.A.")
        wsSammel.Cells(sammelZeile, 1).Value = "Bagatellgrenze:  " & bagatellText
        sammelZeile = sammelZeile + 1
        wsSammel.Cells(sammelZeile, 1).Value = "Entf" & Chr(228) & "lle gesamt:  " & entfGesamt & _
            "          nicht zu z" & Chr(228) & "hlende Entf" & Chr(228) & "lle:  " & entfNicht
        sammelZeile = sammelZeile + 1
        wsSammel.Cells(sammelZeile, 1).Value = "Vertretungen:  " & vertWert & _
            "          nicht zu z" & Chr(228) & "hlende Vertretungen:  " & vertNicht
        sammelZeile = sammelZeile + 1
        With wsSammel.Cells(sammelZeile, 1)
            .Value = "Anzahl abrechenbare Vertretungen:  " & summeWert
            .Font.Bold = True
            If bagatell > 0 And summeWert > bagatell Then .Font.Color = RGB(0, 128, 0)
        End With
        sammelZeile = sammelZeile + 1
        sammelZeile = sammelZeile + 4

        ' Unterschrift
        wsSammel.Cells(sammelZeile, 1).Value = "Zur Kenntnis genommen:"
        wsSammel.Cells(sammelZeile, 1).Font.Bold = True
        sammelZeile = sammelZeile + 5
        wsSammel.Cells(sammelZeile, 1).Value = "________________________________"
        wsSammel.Cells(sammelZeile, 9).Value = "________________________________"
        wsSammel.Cells(sammelZeile, 9).HorizontalAlignment = xlRight
        sammelZeile = sammelZeile + 1
        wsSammel.Cells(sammelZeile, 1).Value = "Datum / Unterschrift Lehrkraft"
        wsSammel.Cells(sammelZeile, 9).Value = "Datum / Unterschrift Schulleitung"
        wsSammel.Cells(sammelZeile, 9).HorizontalAlignment = xlRight
        sammelZeile = sammelZeile + 2

    Next idx

    ' Spaltenbreiten anpassen
    wsSammel.Columns("A:I").AutoFit
    wsSammel.Columns(1).ColumnWidth = 7
    wsSammel.Columns(2).ColumnWidth = 7
    wsSammel.Columns(5).ColumnWidth = 5

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

    If fehlerListe = "" Then
        MsgBox erstellt & " LBV-Formular(e) erstellt in:" & Chr(10) & speicherOrdner, vbInformation
    Else
        MsgBox erstellt & " erstellt. Fehler bei:" & Chr(10) & fehlerListe, vbExclamation
    End If
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
' ================================================================
Sub PDFsDrucken()

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
    liste = liste & Chr(10) & "Standarddrucker wird verwendet." & Chr(10) & Chr(10) & "Fortfahren?"

    If MsgBox(liste, vbYesNo + vbQuestion, "PDFs drucken") = vbNo Then Exit Sub

    ' PDFs an Standarddrucker senden
    Dim gedruckt As Integer: gedruckt = 0
    Dim fehler As String:    fehler = ""

    For i = 0 To anzahl - 1
        Dim vollPfad As String
        vollPfad = druckOrdner & pdfDateien(i)

        ' ShellExecute mit "print" schickt die Datei an den Standarddrucker
        Dim result As LongPtr
        result = ShellExecute( _
            0, _
            "print", _
            vollPfad, _
            vbNullString, _
            druckOrdner, _
            0)          ' 0 = SW_HIDE - kein Fenster

        If result > 32 Then
            gedruckt = gedruckt + 1
        Else
            fehler = fehler & pdfDateien(i) & " (Fehlercode: " & result & ")" & Chr(10)
        End If

        ' Kurze Pause damit der Drucker-Spooler nicht ueberlastet wird
        Application.Wait Now + TimeValue("00:00:02")
    Next i

    ' Abschlussmeldung
    If fehler = "" Then
        MsgBox gedruckt & " PDF(s) wurden an den Drucker gesendet.", vbInformation
    Else
        MsgBox gedruckt & " PDF(s) gesendet." & Chr(10) & Chr(10) & _
               "Fehler bei:" & Chr(10) & fehler, vbExclamation
    End If

End Sub

' ============================================================
' MAKRO: Mehrarbeit-PDFs per E-Mail versenden
' ============================================================
' ============================================================
' MAKRO: Mehrarbeit-PDFs per E-Mail versenden (Outlook)
' Nur Lehrer, die in der Uebersicht bei "PDF erstellen" auf 1 stehen.
' Absenderkonto: Zeile "Outlook-Absender" in der Tabelle Informationen
' (leer = Standardkonto von Outlook).
' ============================================================
Public Sub MehrarbeitPDFsVersenden()
    Dim ws3 As Worksheet
    Set ws3 = Nothing

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

    ' Absenderadresse aus Tabelle Informationen lesen
    Dim absenderAdr As String: absenderAdr = ""
    Dim wsInfo As Worksheet
    Set wsInfo = Nothing
    Dim wsI As Worksheet
    For Each wsI In ThisWorkbook.Worksheets
        If LCase(wsI.Name) Like "*informationen*" Then
            Set wsInfo = wsI: Exit For
        End If
    Next wsI
    If Not wsInfo Is Nothing Then
        Dim iRow As Long
        For iRow = 1 To wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
            Select Case LCase(Trim(wsInfo.Cells(iRow, 1).Value))
                Case "outlook-absender", "outlook-absender:"
                    absenderAdr = Trim(CStr(wsInfo.Cells(iRow, 2).Value))
                    Exit For
            End Select
        Next iRow
    End If

    ' Spalten ermitteln
    Dim colKuerzel As Integer, colEmail As Integer, colPDF As Integer
    colKuerzel = 1  ' Spalte A = Kuerzel
    colEmail = SpalteUebersicht(ws3, "emailadresse")
    If colEmail = 0 Then
        MsgBox "Spalte 'emailadresse' nicht in " & Chr(220) & "bersicht gefunden.", vbCritical: Exit Sub
    End If
    colPDF = SpaltePDF()   ' Spalte "PDF erstellen" (wird bei Bedarf angelegt)

    Dim lastRow As Long
    lastRow = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row

    ' Pruefen, ob ueberhaupt Lehrer markiert sind
    Dim anzahlMarkiert As Long: anzahlMarkiert = 0
    Dim rz As Long
    For rz = 2 To lastRow
        If Trim(ws3.Cells(rz, colKuerzel).Value) <> "" And ws3.Cells(rz, colPDF).Value = 1 Then
            anzahlMarkiert = anzahlMarkiert + 1
        End If
    Next rz
    If anzahlMarkiert = 0 Then
        MsgBox "Keine Lehrer in der Spalte 'PDF erstellen' markiert.", vbExclamation: Exit Sub
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
    If Right(pdfOrdner, 1) <> "\" Then pdfOrdner = pdfOrdner & "\"

    ' Outlook-Objekt erstellen
    Dim olApp As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        MsgBox "Outlook konnte nicht gestartet werden.", vbCritical: Exit Sub
    End If

    ' Absenderkonto suchen
    ' - Adresse eingetragen: Konto muss in Outlook existieren, sonst Abbruch
    ' - Adresse leer: Standardkonto von Outlook wird verwendet
    Dim zielKonto As Object
    Set zielKonto = Nothing
    If absenderAdr <> "" Then
        Dim olAccount As Object
        Dim kontenListe As String: kontenListe = ""
        For Each olAccount In olApp.Session.Accounts
            kontenListe = kontenListe & "  " & olAccount.SmtpAddress & Chr(10)
            If LCase(olAccount.SmtpAddress) = LCase(absenderAdr) Then
                Set zielKonto = olAccount
                Exit For
            End If
        Next olAccount
        If zielKonto Is Nothing Then
            MsgBox "Das Absenderkonto '" & absenderAdr & "' ist in Outlook nicht eingerichtet." & _
                   Chr(10) & Chr(10) & "Vorhandene Konten:" & Chr(10) & kontenListe & Chr(10) & _
                   "Bitte die Zeile 'Outlook-Absender' in der Tabelle Informationen pr" & Chr(252) & "fen.", _
                   vbCritical, "Absenderkonto fehlt"
            Exit Sub
        End If
    End If

    Dim gesendet As Integer, keinEmail As Integer, keinPDF As Integer, nichtMarkiert As Integer
    gesendet = 0: keinEmail = 0: keinPDF = 0: nichtMarkiert = 0

    Dim r As Long
    For r = 2 To lastRow
        Dim kuerzel As String
        kuerzel = Trim(ws3.Cells(r, colKuerzel).Value)
        If kuerzel = "" Then GoTo NaechsterLehrer

        ' Nur markierte Lehrer ("PDF erstellen" = 1)
        If ws3.Cells(r, colPDF).Value <> 1 Then
            nichtMarkiert = nichtMarkiert + 1
            GoTo NaechsterLehrer
        End If

        Dim emailAdr As String
        emailAdr = Trim(ws3.Cells(r, colEmail).Value)
        If emailAdr = "" Then
            keinEmail = keinEmail + 1
            GoTo NaechsterLehrer
        End If

        ' Alle passenden PDFs finden: Mehrarbeit_<kuerzel>_*.pdf
        Dim suchMuster As String
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
           "Absender: " & IIf(absenderAdr = "", "Outlook-Standardkonto", absenderAdr) & Chr(10) & _
           "Gesendet: " & gesendet & Chr(10) & _
           "Keine E-Mail-Adresse: " & keinEmail & Chr(10) & _
           "Keine PDF gefunden: " & keinPDF & Chr(10) & _
           "Nicht markiert (" & Chr(252) & "bersprungen): " & nichtMarkiert, vbInformation, "E-Mail-Versand"
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
