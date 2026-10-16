Attribute VB_Name = "Modul2"
Option Explicit

' ============================================================
' Makro 9: LBVZusammenfassen
' Liest alle LBV_Mehrarbeit_<Kuerzel>_<MM>_<JJJJ>.pdf
' eines Ordners, gruppiert nach Lehrer-Kuerzel und erstellt
' fuer jeden Lehrer eine Zusammenfassungs-PDF mit allen
' Monatseintraegen (Bereich A) in der gleichen Vorlage.
' Benoetigt: lbv_merge.exe ODER lbv_merge.py im selben Ordner wie lbv_fill.exe
' ============================================================
Sub LBVZusammenfassen()

    ' --- Tabelle Informationen lesen ---
    Dim wsInfo As Worksheet
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Worksheets("Informationen")
    On Error GoTo 0

    Dim vorlagenOrdner As String: vorlagenOrdner = ""
    Dim vorlagenDatei  As String: vorlagenDatei = ""
    Dim exePfad        As String: exePfad = ""

    If Not wsInfo Is Nothing Then
        Dim iRow As Long
        For iRow = 1 To wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
            Dim iKey As String
            iKey = LCase(Trim(CStr(wsInfo.Cells(iRow, 1).Value)))
            Dim iVal As String
            iVal = Trim(CStr(wsInfo.Cells(iRow, 2).Value))
            Select Case iKey
                Case "pfad zur manuell bef" & Chr(252) & "llten vorlage"
                    vorlagenOrdner = iVal
                Case "dateiname vorlage"
                    vorlagenDatei = iVal
                Case "pfad lbv_fill.exe"
                    exePfad = iVal
            End Select
        Next iRow
    End If

    If vorlagenOrdner = "" Or vorlagenDatei = "" Then
        MsgBox "Bitte Pfad und Dateiname der Vorlage in Tabelle Informationen eintragen.", vbExclamation
        Exit Sub
    End If
    If Right(vorlagenOrdner, 1) <> "\" Then vorlagenOrdner = vorlagenOrdner & "\"
    Dim vorlagePfad As String
    vorlagePfad = vorlagenOrdner & vorlagenDatei & ".pdf"
    If Dir(vorlagePfad) = "" Then
        MsgBox "Vorlagen-PDF nicht gefunden:" & Chr(10) & vorlagePfad, vbCritical
        Exit Sub
    End If

    ' --- lbv_merge.exe und lbv_merge.py lokalisieren ---
    Dim skriptOrdner As String: skriptOrdner = ""
    If exePfad <> "" Then skriptOrdner = Left(exePfad, InStrRev(exePfad, "\"))

    Dim mergeExe As String: mergeExe = ""
    Dim mergePy  As String: mergePy = ""

    ' EXE suchen (neben lbv_fill.exe, dann neben XLSM)
    If skriptOrdner <> "" And Dir(skriptOrdner & "lbv_merge.exe") <> "" Then
        mergeExe = skriptOrdner & "lbv_merge.exe"
    ElseIf Dir(ThisWorkbook.Path & "\lbv_merge.exe") <> "" Then
        mergeExe = ThisWorkbook.Path & "\lbv_merge.exe"
    End If

    ' PY suchen (neben lbv_fill.exe, dann neben XLSM)
    If skriptOrdner <> "" And Dir(skriptOrdner & "lbv_merge.py") <> "" Then
        mergePy = skriptOrdner & "lbv_merge.py"
    ElseIf Dir(ThisWorkbook.Path & "\lbv_merge.py") <> "" Then
        mergePy = ThisWorkbook.Path & "\lbv_merge.py"
    End If

    If mergeExe = "" And mergePy = "" Then
        Dim suchOrdner As String
        suchOrdner = IIf(skriptOrdner <> "", skriptOrdner, ThisWorkbook.Path & "\")
        MsgBox "Weder lbv_merge.exe noch lbv_merge.py gefunden." & Chr(10) & Chr(10) & _
               "Gesucht in: " & suchOrdner & Chr(10) & Chr(10) & _
               "Bitte eine der Dateien dort ablegen.", vbCritical
        Exit Sub
    End If

    ' --- Python-Befehl ermitteln (fuer PY-Fallback) ---
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim pyCmd As String: pyCmd = ""
    If mergePy <> "" Then
        pyCmd = "py"
        If wsh.Run("cmd /c py --version > nul 2>&1", 0, True) <> 0 Then
            If wsh.Run("cmd /c python --version > nul 2>&1", 0, True) = 0 Then
                pyCmd = "python"
            Else
                pyCmd = ""
            End If
        End If
    End If

    ' --- Quellordner waehlen ---
    Dim quellOrdner As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Ordner mit LBV-Formularen w" & Chr(228) & "hlen"
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        quellOrdner = .SelectedItems(1)
    End With
    If Right(quellOrdner, 1) <> "\" Then quellOrdner = quellOrdner & "\"

    ' --- Alle LBV_Mehrarbeit_*.pdf einlesen ---
    Dim alleDateien() As String
    Dim anzahl As Integer: anzahl = 0
    Dim datei As String
    datei = Dir(quellOrdner & "LBV_Mehrarbeit_*.pdf")
    Do While datei <> ""
        ' Zusammenfassungs-Dateien (_Gesamt.pdf) ueberspringen
        If InStr(LCase(datei), "_gesamt") = 0 Then
            ReDim Preserve alleDateien(anzahl)
            alleDateien(anzahl) = datei
            anzahl = anzahl + 1
        End If
        datei = Dir()
    Loop

    If anzahl = 0 Then
        MsgBox "Keine LBV_Mehrarbeit_*.pdf Dateien gefunden in:" & Chr(10) & quellOrdner, vbInformation
        Exit Sub
    End If

    ' --- Nach Kuerzel gruppieren ---
    Dim kuerzelListe() As String
    Dim kuerzelAnz As Integer: kuerzelAnz = 0
    Dim i As Integer, j As Integer

    For i = 0 To anzahl - 1
        Dim k As String: k = KuerzelAusDateiname(alleDateien(i))
        If k = "" Then GoTo WeiterGruppieren
        Dim kGefunden As Boolean: kGefunden = False
        For j = 0 To kuerzelAnz - 1
            If kuerzelListe(j) = k Then kGefunden = True: Exit For
        Next j
        If Not kGefunden Then
            ReDim Preserve kuerzelListe(kuerzelAnz)
            kuerzelListe(kuerzelAnz) = k
            kuerzelAnz = kuerzelAnz + 1
        End If
WeiterGruppieren:
    Next i

    If kuerzelAnz = 0 Then
        MsgBox "Kein gueltiges Dateinamen-Muster gefunden." & Chr(10) & _
               "Erwartet: LBV_Mehrarbeit_<Kuerzel>_<MM>_<JJJJ>.pdf", vbExclamation
        Exit Sub
    End If

    ' --- Zusammenfassung anzeigen und bestaetigen ---
    Dim uebersicht As String: uebersicht = kuerzelAnz & " Lehrer gefunden:" & Chr(10) & Chr(10)
    For i = 0 To kuerzelAnz - 1
        Dim cnt As Integer: cnt = 0
        For j = 0 To anzahl - 1
            If KuerzelAusDateiname(alleDateien(j)) = kuerzelListe(i) Then cnt = cnt + 1
        Next j
        uebersicht = uebersicht & "  " & kuerzelListe(i) & ":  " & cnt & " Monat(e)" & Chr(10)
    Next i
    uebersicht = uebersicht & Chr(10) & "Zusammenfassungs-PDFs werden im gleichen Ordner gespeichert."
    If MsgBox(uebersicht, vbOKCancel + vbQuestion, "LBV zusammenfassen") = vbCancel Then Exit Sub

    ' --- Pro Lehrer verarbeiten ---
    Dim erstellt As Integer: erstellt = 0
    Dim fehlerListe As String: fehlerListe = ""

    For i = 0 To kuerzelAnz - 1
        Dim kue As String: kue = kuerzelListe(i)

        ' Dateien dieses Lehrers sammeln
        Dim quellen() As String
        Dim quellenAnz As Integer: quellenAnz = 0
        For j = 0 To anzahl - 1
            If KuerzelAusDateiname(alleDateien(j)) = kue Then
                ReDim Preserve quellen(quellenAnz)
                quellen(quellenAnz) = quellOrdner & alleDateien(j)
                quellenAnz = quellenAnz + 1
            End If
        Next j

        If quellenAnz < 1 Then GoTo NaechsterLehrer

        ' Nur bei mehreren Dateien sinnvoll - bei einer Datei trotzdem erlauben
        ' Chronologisch sortieren (aelteste zuerst)
        If quellenAnz > 1 Then
            quellen = DateienChronologisch(quellen, quellenAnz)
        End If

        ' Ausgabepfad
        Dim ausgabePfad As String
        ausgabePfad = quellOrdner & "LBV_Mehrarbeit_" & kue & "_Gesamt.pdf"

        ' JSON schreiben
        Dim jsonPfad As String
        jsonPfad = Environ("TEMP") & "\lbv_merge_" & kue & ".json"
        Dim logPfad As String
        logPfad = Environ("TEMP") & "\lbv_merge_log_" & kue & ".txt"

        Dim fj As Integer: fj = FreeFile
        Open jsonPfad For Output As #fj
        Print #fj, "{"
        Print #fj, "  " & Chr(34) & "vorlage_pfad" & Chr(34) & ": " & Chr(34) & Replace(vorlagePfad, "\", "/") & Chr(34) & ","
        Print #fj, "  " & Chr(34) & "ausgabe_pfad" & Chr(34) & ": " & Chr(34) & Replace(ausgabePfad, "\", "/") & Chr(34) & ","
        Print #fj, "  " & Chr(34) & "quellen" & Chr(34) & ": ["
        Dim q As Integer
        For q = 0 To quellenAnz - 1
            Dim komma As String: komma = IIf(q < quellenAnz - 1, ",", "")
            Print #fj, "    " & Chr(34) & Replace(quellen(q), "\", "/") & Chr(34) & komma
        Next q
        Print #fj, "  ]"
        Print #fj, "}"
        Close #fj

        ' Batch-Datei + EXE/PY-Aufruf (EXE bevorzugt, Python-Fallback)
        Dim batPfad As String
        batPfad = Environ("TEMP") & "\lbv_merge_run_" & kue & ".bat"
        Dim exitCode As Long
        Dim usedExe As Boolean: usedExe = False

        ' --- EXE versuchen ---
        If mergeExe <> "" Then
            Dim fb1 As Integer: fb1 = FreeFile
            Open batPfad For Output As #fb1
            Print #fb1, "@echo off"
            Print #fb1, Chr(34) & mergeExe & Chr(34) & " " & Chr(34) & jsonPfad & Chr(34)
            Print #fb1, "exit /b %errorlevel%"
            Close #fb1
            exitCode = wsh.Run(Chr(34) & batPfad & Chr(34), 0, True)
            Application.Wait Now + TimeValue("00:00:02")
            If exitCode = 0 And Dir(ausgabePfad) <> "" Then usedExe = True
            If exitCode = 0 And Dir(ausgabePfad) = "" Then exitCode = -1
        End If

        ' --- Python-Fallback ---
        If Not usedExe And mergePy <> "" Then
            If pyCmd = "" Then
                MsgBox "EXE fehlgeschlagen und Python nicht installiert." & Chr(10) & _
                       "Bitte Python installieren oder lbv_merge.exe bereitstellen.", _
                       vbCritical, "Fehler bei " & kue
                GoTo NaechsterLehrer
            End If
            Dim fb2 As Integer: fb2 = FreeFile
            Open batPfad For Output As #fb2
            Print #fb2, "@echo off"
            Print #fb2, pyCmd & " " & Chr(34) & mergePy & Chr(34) & " " & Chr(34) & jsonPfad & Chr(34) & " > " & Chr(34) & logPfad & Chr(34) & " 2>&1"
            Print #fb2, "exit /b %errorlevel%"
            Close #fb2
            exitCode = wsh.Run(Chr(34) & batPfad & Chr(34), 0, True)
            Application.Wait Now + TimeValue("00:00:02")
            If exitCode = 0 And Dir(ausgabePfad) <> "" Then usedExe = True
            If exitCode = 0 And Dir(ausgabePfad) = "" Then exitCode = -1
        End If

        If usedExe Then
            erstellt = erstellt + 1
        Else
            fehlerListe = fehlerListe & kue & " (Exit: " & exitCode & ")" & Chr(10)
            Dim diagMsg As String
            If exitCode = -1 Then
                diagMsg = "Skript lief durch, aber keine PDF erstellt:" & Chr(10) & ausgabePfad
            Else
                diagMsg = "Exit-Code: " & exitCode
            End If
            If Dir(logPfad) <> "" Then
                Dim fl As Integer: fl = FreeFile
                Dim ll As String, lt As String: lt = ""
                Open logPfad For Input As #fl
                Dim lc As Integer: lc = 0
                Do While Not EOF(fl) And lc < 40
                    Line Input #fl, ll
                    lt = lt & ll & Chr(10)
                    lc = lc + 1
                Loop
                Close #fl
                diagMsg = diagMsg & Chr(10) & Chr(10) & lt
            End If
            MsgBox diagMsg, vbCritical, "Fehler bei " & kue
        End If

NaechsterLehrer:
    Next i

    If fehlerListe = "" Then
        MsgBox erstellt & " Zusammenfassungs-PDF(s) erstellt in:" & Chr(10) & quellOrdner, vbInformation
    Else
        MsgBox erstellt & " erstellt. Fehler bei:" & Chr(10) & fehlerListe, vbExclamation
    End If

End Sub

' ----------------------------------------------------------------
' Hilfsfunktion: Kuerzel aus Dateinamen extrahieren
' Schema: LBV_Mehrarbeit_<Kuerzel>_<MM>_<JJJJ>.pdf
' ----------------------------------------------------------------
Private Function KuerzelAusDateiname(dateiname As String) As String
    Dim s As String: s = dateiname
    ' .pdf entfernen
    If LCase(Right(s, 4)) = ".pdf" Then s = Left(s, Len(s) - 4)
    ' Praefix pruefen
    Const PRAEFIX As String = "LBV_Mehrarbeit_"
    If Left(s, Len(PRAEFIX)) <> PRAEFIX Then
        KuerzelAusDateiname = ""
        Exit Function
    End If
    s = Mid(s, Len(PRAEFIX) + 1)   ' "<Kuerzel>_<MM>_<JJJJ>"
    ' Letzten zwei Teile (_MM_JJJJ) entfernen
    Dim parts() As String: parts = Split(s, "_")
    If UBound(parts) < 2 Then
        KuerzelAusDateiname = ""    ' kein gueltiges Muster
        Exit Function
    End If
    ' Letzter Teil muss 4-stellige Jahreszahl sein, vorletzter 2-stelliger Monat
    Dim jjjj As String: jjjj = parts(UBound(parts))
    Dim mm   As String: mm = parts(UBound(parts) - 1)
    If Len(jjjj) <> 4 Or Not IsNumeric(jjjj) Then
        KuerzelAusDateiname = ""
        Exit Function
    End If
    If Len(mm) <> 2 Or Not IsNumeric(mm) Then
        KuerzelAusDateiname = ""
        Exit Function
    End If
    ' Kuerzel = alle Teile ausser den letzten beiden
    Dim result As String: result = ""
    Dim p As Integer
    For p = 0 To UBound(parts) - 2
        result = result & IIf(result = "", "", "_") & parts(p)
    Next p
    KuerzelAusDateiname = result
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Dateien chronologisch sortieren (aelteste zuerst)
' Sortierkey: JJJJ * 100 + MM aus Dateinamen
' ----------------------------------------------------------------
Private Function DateienChronologisch(dateien() As String, anzahl As Integer) As String()
    Dim i As Integer, j As Integer, temp As String
    For i = 0 To anzahl - 2
        For j = 0 To anzahl - 2 - i
            If DateiSortKey(dateien(j)) > DateiSortKey(dateien(j + 1)) Then
                temp = dateien(j)
                dateien(j) = dateien(j + 1)
                dateien(j + 1) = temp
            End If
        Next j
    Next i
    DateienChronologisch = dateien
End Function

Private Function DateiSortKey(pfad As String) As Long
    ' Gibt JJJJ * 100 + MM zurueck fuer chronologische Sortierung
    Dim dateiname As String
    Dim slashPos As Integer: slashPos = InStrRev(pfad, "\")
    If slashPos > 0 Then
        dateiname = Mid(pfad, slashPos + 1)
    Else
        dateiname = pfad
    End If
    If LCase(Right(dateiname, 4)) = ".pdf" Then dateiname = Left(dateiname, Len(dateiname) - 4)
    Dim parts() As String: parts = Split(dateiname, "_")
    If UBound(parts) >= 2 Then
        Dim jjjj As String: jjjj = parts(UBound(parts))
        Dim mm   As String: mm = parts(UBound(parts) - 1)
        If IsNumeric(jjjj) And IsNumeric(mm) Then
            DateiSortKey = CLng(jjjj) * 100 + CLng(mm)
            Exit Function
        End If
    End If
    DateiSortKey = 0
End Function
