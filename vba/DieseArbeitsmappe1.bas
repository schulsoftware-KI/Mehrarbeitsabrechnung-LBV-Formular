Attribute VB_Name = "DieseArbeitsmappe1"

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

    ' Nur in der Übersicht reagieren
    If Sh.Name <> Chr(220) & "bersicht" Then Exit Sub

    ' Spalte "PDF erstellen" ermitteln
    Dim colPDF As Integer
    colPDF = 0
    Dim c As Integer
    Dim lastCol As Integer
    lastCol = Sh.Cells(1, Sh.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(Sh.Cells(1, c).Value) = "PDF erstellen" Then
            colPDF = c
            Exit For
        End If
    Next c

    ' Nur reagieren wenn in der PDF-Spalte geklickt wurde (nicht Kopfzeile)
    If colPDF = 0 Then Exit Sub
    If Target.Column <> colPDF Then Exit Sub
    If Target.Row = 1 Then Exit Sub
    If Trim(Sh.Cells(Target.Row, 1).Value) = "" Then Exit Sub

    ' Doppelklick abfangen (kein Editiermodus)
    Cancel = True

    ' Wert umschalten 0 <-> 1
    If Sh.Cells(Target.Row, colPDF).Value = 1 Then
        Sh.Cells(Target.Row, colPDF).Value = 0
    Else
        Sh.Cells(Target.Row, colPDF).Value = 1
    End If

    ' Formatierung aktualisieren
    Call Modul1.PDFSpalteFormatieren

End Sub
