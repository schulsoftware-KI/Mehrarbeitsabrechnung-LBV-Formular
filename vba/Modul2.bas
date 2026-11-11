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
    uebersicht = uebersicht & Chr(10) & "Zusammenfassungs-PDFs werden im gleichen Ordner gespeichert." & Chr(10) & _
                 "Danach wird die LBV-Sammeldatei aller Gesamtformulare angeboten."
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

    If erstellt = 0 Then
        MsgBox erstellt & " erstellt. Fehler bei:" & Chr(10) & fehlerListe, vbExclamation
        Exit Sub
    End If

    ' --- Anschliessend automatisch die LBV-Sammeldatei aller Gesamtformulare ---
    ' (Modus "Gesamt", gleicher Ordner - Modus und Ordner werden nicht erneut abgefragt.
    '  Die Erfolgsmeldung des Zusammenfassens erscheint im ersten Dialog der Sammeldatei.)
    Dim einleitung As String
    If fehlerListe = "" Then
        einleitung = erstellt & " Zusammenfassungs-PDF(s) erstellt in:" & Chr(10) & quellOrdner & Chr(10) & Chr(10)
    Else
        MsgBox erstellt & " erstellt. Fehler bei:" & Chr(10) & fehlerListe, vbExclamation
        einleitung = "Hinweis: F" & Chr(252) & "r Lehrer mit Fehler beim Zusammenfassen wird ggf. ein " & _
                     Chr(228) & "lteres _Gesamt.pdf aus dem Ordner verwendet." & Chr(10) & Chr(10)
    End If
    LBVSammeldateiIntern 2, quellOrdner, einleitung

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

' ============================================================
' Makro 10: LBVSammeldateiErstellen
' Fasst wahlweise
'   - alle LBV-Einzelformulare  LBV_Mehrarbeit_<Kuerzel>_<MM>_<JJJJ>.pdf
'   - alle LBV-Gesamtformulare  LBV_Mehrarbeit_<Kuerzel>_Gesamt.pdf
'     (Ergebnis von "Merge LBV" / LBVZusammenfassen)
' eines Ordners zu EINER PDF-Datei fuer alle Lehrer zusammen.
' Reihenfolge: Kuerzel alphabetisch, Einzelformulare je Lehrer chronologisch.
' Die Sammeldatei bekommt je Formular ein Lesezeichen (z.B. "AB - 09/2025").
'
' Die Formularfelder werden in der Sammeldatei fest in die Seiten
' uebernommen ("flatten"). Grund: alle LBV-Formulare haben dieselben
' Feldnamen - ohne Flatten wuerde ein PDF-Viewer auf jeder Seite die
' Werte des ersten Lehrers anzeigen. Die Quelldateien bleiben unveraendert.
'
' Wird auch automatisch am Ende von LBVZusammenfassen ("Merge LBV")
' aufgerufen - dann fest mit Gesamtformularen und demselben Ordner.
'
' Benoetigt: lbv_sammel.exe ODER lbv_sammel.py im selben Ordner wie
'            lbv_fill.exe (oder im Ordner dieser Arbeitsmappe)
' ============================================================
Sub LBVSammeldateiErstellen()
    LBVSammeldateiIntern 0, "", ""
End Sub

' ----------------------------------------------------------------
' Eigentliche Ausfuehrung
'   vorgabeModus:  0 = nachfragen, 1 = Einzelformulare, 2 = Gesamtformulare
'   vorgabeOrdner: "" = Ordner per Dialog waehlen, sonst dieser Ordner
'   einleitung:    Text, der dem ersten Dialog vorangestellt wird
' ----------------------------------------------------------------
Private Sub LBVSammeldateiIntern(ByVal vorgabeModus As Long, ByVal vorgabeOrdner As String, _
                                 ByVal einleitung As String)

    Dim q As String: q = Chr(34)

    ' --- Pfad lbv_fill.exe aus Tabelle Informationen (dort liegen die Hilfsprogramme) ---
    Dim wsInfo As Worksheet
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Worksheets("Informationen")
    On Error GoTo 0

    Dim exePfad As String: exePfad = ""
    Dim iRow As Long
    If Not wsInfo Is Nothing Then
        For iRow = 1 To wsInfo.Cells(wsInfo.Rows.Count, "A").End(xlUp).Row
            If LCase(Trim(CStr(wsInfo.Cells(iRow, 1).Value))) = "pfad lbv_fill.exe" Then
                exePfad = Trim(CStr(wsInfo.Cells(iRow, 2).Value))
                Exit For
            End If
        Next iRow
    End If
    Dim skriptOrdner As String: skriptOrdner = ""
    If exePfad <> "" Then skriptOrdner = Left(exePfad, InStrRev(exePfad, "\"))

    ' --- lbv_sammel.exe / lbv_sammel.py suchen (neben lbv_fill.exe, dann neben XLSM) ---
    Dim sammelExe As String: sammelExe = HilfsdateiSuchen(skriptOrdner, "lbv_sammel.exe")
    Dim sammelPy As String:  sammelPy = HilfsdateiSuchen(skriptOrdner, "lbv_sammel.py")
    If sammelExe = "" And sammelPy = "" Then
        MsgBox einleitung & "LBV-Sammeldatei: Weder lbv_sammel.exe noch lbv_sammel.py gefunden." & vbLf & vbLf & _
               "Gesucht in: " & IIf(skriptOrdner <> "", skriptOrdner, ThisWorkbook.Path & "\") & vbLf & vbLf & _
               "Bitte eine der Dateien dort ablegen (am besten neben lbv_fill.exe).", vbCritical
        Exit Sub
    End If

    ' --- Auswahl: Einzel- oder Gesamtformulare (entfaellt bei Vorgabe) ---
    Dim modusGesamt As Boolean
    If vorgabeModus = 0 Then
        Dim antwort As VbMsgBoxResult
        antwort = MsgBox("Welche LBV-Formulare sollen zu EINER Datei f" & Chr(252) & "r alle Lehrer " & _
                         "zusammengefasst werden?" & vbLf & vbLf & _
                         "[Ja]" & vbTab & "alle EINZELformulare (je Lehrer und Monat)" & vbLf & _
                         vbTab & "LBV_Mehrarbeit_<K" & Chr(252) & "rzel>_<MM>_<JJJJ>.pdf" & vbLf & vbLf & _
                         "[Nein]" & vbTab & "alle GESAMTformulare (Ergebnis von 'Merge LBV')" & vbLf & _
                         vbTab & "LBV_Mehrarbeit_<K" & Chr(252) & "rzel>_Gesamt.pdf", _
                         vbYesNoCancel + vbQuestion, "LBV-Sammeldatei erstellen")
        If antwort = vbCancel Then Exit Sub
        modusGesamt = (antwort = vbNo)
    Else
        modusGesamt = (vorgabeModus = 2)
    End If
    Dim artText As String
    artText = IIf(modusGesamt, "Gesamtformulare", "Einzelformulare")

    ' --- Quellordner (Vorgabe oder per Dialog) ---
    Dim quellOrdner As String
    If vorgabeOrdner <> "" Then
        quellOrdner = vorgabeOrdner
    Else
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Ordner mit LBV-" & artText & "n w" & Chr(228) & "hlen"
            .Show
            If .SelectedItems.Count = 0 Then Exit Sub
            quellOrdner = .SelectedItems(1)
        End With
    End If
    If Right(quellOrdner, 1) <> "\" Then quellOrdner = quellOrdner & "\"

    ' --- Passende Dateien einlesen ---
    Dim dateien() As String, kuerzelArr() As String, sortKeys() As String
    Dim anzahl As Long: anzahl = 0
    Dim datei As String, k As String
    datei = Dir(quellOrdner & "LBV_Mehrarbeit_*.pdf")
    Do While datei <> ""
        k = ""
        If modusGesamt Then
            k = KuerzelAusGesamtname(datei)
        ElseIf InStr(LCase(datei), "_gesamt") = 0 Then
            k = KuerzelAusDateiname(datei)
        End If
        If k <> "" Then
            ReDim Preserve dateien(anzahl)
            ReDim Preserve kuerzelArr(anzahl)
            ReDim Preserve sortKeys(anzahl)
            dateien(anzahl) = datei
            kuerzelArr(anzahl) = k
            ' Sortierung: Kuerzel, dann JJJJMM (Gesamtformulare: 000000)
            sortKeys(anzahl) = LCase(k) & Chr(1) & Format(DateiSortKey(datei), "000000")
            anzahl = anzahl + 1
        End If
        datei = Dir()
    Loop

    If anzahl = 0 Then
        MsgBox einleitung & "Keine LBV-" & artText & " gefunden in:" & vbLf & quellOrdner & vbLf & vbLf & _
               "Erwartet: " & IIf(modusGesamt, "LBV_Mehrarbeit_<K" & Chr(252) & "rzel>_Gesamt.pdf", _
                                              "LBV_Mehrarbeit_<K" & Chr(252) & "rzel>_<MM>_<JJJJ>.pdf"), vbInformation
        Exit Sub
    End If

    ' --- Sortieren (Einfuegesortierung, alle drei Listen parallel) ---
    Dim a As Long, b As Long
    Dim tmpS As String, tmpD As String, tmpK As String
    For a = 1 To anzahl - 1
        tmpS = sortKeys(a): tmpD = dateien(a): tmpK = kuerzelArr(a)
        b = a - 1
        Do While b >= 0
            If StrComp(sortKeys(b), tmpS, vbBinaryCompare) <= 0 Then Exit Do
            sortKeys(b + 1) = sortKeys(b): dateien(b + 1) = dateien(b): kuerzelArr(b + 1) = kuerzelArr(b)
            b = b - 1
        Loop
        sortKeys(b + 1) = tmpS: dateien(b + 1) = tmpD: kuerzelArr(b + 1) = tmpK
    Next a

    ' --- Uebersicht anzeigen und bestaetigen ---
    Const MAX_LISTE As Long = 25
    Dim liste As String: liste = ""
    Dim anzLehrer As Long: anzLehrer = 0
    Dim i As Long, j As Long
    i = 0
    Do While i < anzahl
        j = i
        Do While j + 1 < anzahl
            If LCase(kuerzelArr(j + 1)) <> LCase(kuerzelArr(i)) Then Exit Do
            j = j + 1
        Loop
        anzLehrer = anzLehrer + 1
        If anzLehrer <= MAX_LISTE Then
            If modusGesamt Then
                liste = liste & "  " & kuerzelArr(i) & vbLf
            Else
                liste = liste & "  " & kuerzelArr(i) & ":  " & (j - i + 1) & " Monat(e)" & vbLf
            End If
        End If
        i = j + 1
    Loop
    If anzLehrer > MAX_LISTE Then liste = liste & "  ... und " & (anzLehrer - MAX_LISTE) & " weitere" & vbLf

    If MsgBox(einleitung & anzahl & " " & artText & " von " & anzLehrer & " Lehrer(n) gefunden:" & vbLf & vbLf & _
              liste & vbLf & "Alle zu EINER PDF-Sammeldatei zusammenfassen?", _
              vbOKCancel + vbQuestion, "LBV-Sammeldatei erstellen") = vbCancel Then Exit Sub

    ' --- Speicherort / Dateiname der Sammeldatei ---
    Dim auswahl As Variant
    auswahl = Application.GetSaveAsFilename( _
        InitialFileName:=quellOrdner & "LBV_Sammeldatei_" & artText & "_" & Format(Date, "YYYY-MM-DD") & ".pdf", _
        FileFilter:="PDF-Dateien (*.pdf), *.pdf", _
        Title:="Sammeldatei speichern unter")
    If VarType(auswahl) = vbBoolean Then Exit Sub
    Dim ausgabePfad As String: ausgabePfad = CStr(auswahl)
    If LCase(Right(ausgabePfad, 4)) <> ".pdf" Then ausgabePfad = ausgabePfad & ".pdf"

    ' Name darf nicht wie ein Lehrer-Formular aussehen, sonst wird die
    ' Sammeldatei spaeter von "Merge LBV" bzw. hier selbst mit eingelesen
    Dim ausgabeName As String
    ausgabeName = Mid(ausgabePfad, InStrRev(ausgabePfad, "\") + 1)
    If LCase(Left(ausgabeName, 15)) = "lbv_mehrarbeit_" Then
        MsgBox "Der Dateiname der Sammeldatei darf nicht mit 'LBV_Mehrarbeit_' beginnen," & vbLf & _
               "sonst wird sie sp" & Chr(228) & "ter als Lehrer-Formular mit eingelesen." & vbLf & vbLf & _
               "Bitte einen anderen Namen w" & Chr(228) & "hlen.", vbExclamation
        Exit Sub
    End If

    ' Vorhandene Datei: nachfragen und vorab loeschen (damit spaeter eindeutig
    ' erkennbar ist, ob die neue Datei wirklich erstellt wurde)
    If Dir(ausgabePfad) <> "" Then
        If MsgBox("Die Datei existiert bereits:" & vbLf & ausgabePfad & vbLf & vbLf & _
                  Chr(220) & "berschreiben?", vbYesNo + vbQuestion, "LBV-Sammeldatei") = vbNo Then Exit Sub
        Dim killFehler As Long
        On Error Resume Next
        Kill ausgabePfad
        killFehler = Err.Number
        On Error GoTo 0
        If killFehler <> 0 Then
            MsgBox "Die vorhandene Datei kann nicht ersetzt werden." & vbLf & _
                   "Ist sie noch in einem PDF-Programm ge" & Chr(246) & "ffnet?" & vbLf & vbLf & ausgabePfad, vbExclamation
            Exit Sub
        End If
    End If

    ' --- JSON schreiben (Pfade als \u-Escapes -> Umlaute in Pfaden sind kein Problem) ---
    Dim jsonPfad As String: jsonPfad = Environ("TEMP") & "\lbv_sammel.json"
    Dim logPfad As String:  logPfad = Environ("TEMP") & "\lbv_sammel_log.txt"
    On Error Resume Next
    Kill logPfad
    On Error GoTo 0

    Dim fj As Integer: fj = FreeFile
    Open jsonPfad For Output As #fj
    Print #fj, "{"
    Print #fj, "  " & q & "ausgabe_pfad" & q & ": " & q & JsonText(Replace(ausgabePfad, "\", "/")) & q & ","
    Print #fj, "  " & q & "log_pfad" & q & ": " & q & JsonText(Replace(logPfad, "\", "/")) & q & ","
    Print #fj, "  " & q & "quellen" & q & ": ["
    For i = 0 To anzahl - 1
        Print #fj, "    " & q & JsonText(Replace(quellOrdner & dateien(i), "\", "/")) & q & IIf(i < anzahl - 1, ",", "")
    Next i
    Print #fj, "  ]"
    Print #fj, "}"
    Close #fj

    ' --- lbv_sammel.exe aufrufen, bei Fehlschlag Python-Fallback ---
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim exitCode As Long: exitCode = -1
    Dim erfolg As Boolean: erfolg = False
    Dim pyFehlt As Boolean: pyFehlt = False

    Application.StatusBar = "LBV-Sammeldatei wird erstellt (" & anzahl & " Formulare) ..."
    Application.Cursor = xlWait

    ' (On Error: falls z.B. ein Virenscanner den Programmstart blockiert)
    If sammelExe <> "" Then
        On Error Resume Next
        exitCode = wsh.Run(q & sammelExe & q & " " & q & jsonPfad & q, 0, True)
        If Err.Number <> 0 Then exitCode = -1: Err.Clear
        On Error GoTo 0
        erfolg = ((exitCode = 0 Or exitCode = 2) And Dir(ausgabePfad) <> "")
    End If

    If Not erfolg And sammelPy <> "" Then
        Dim pyCmd As String: pyCmd = PythonBefehl(wsh)
        If pyCmd = "" Then
            pyFehlt = True
        Else
            On Error Resume Next
            exitCode = wsh.Run(pyCmd & " " & q & sammelPy & q & " " & q & jsonPfad & q, 0, True)
            If Err.Number <> 0 Then exitCode = -1: Err.Clear
            On Error GoTo 0
            erfolg = ((exitCode = 0 Or exitCode = 2) And Dir(ausgabePfad) <> "")
        End If
    End If

    Application.Cursor = xlDefault
    Application.StatusBar = False

    ' --- Ergebnis melden ---
    If Not erfolg Then
        Dim diag As String
        diag = "Die Sammeldatei konnte nicht erstellt werden (Exit-Code: " & exitCode & ")." & vbLf
        If pyFehlt Then
            diag = diag & vbLf & "lbv_sammel.exe fehlt oder ist fehlgeschlagen, und Python ist nicht installiert." & vbLf & _
                   "Bitte lbv_sammel.exe neben lbv_fill.exe ablegen."
        End If
        diag = diag & vbLf & LogLesen(logPfad, False)
        MsgBox diag, vbCritical, "LBV-Sammeldatei"
        Exit Sub
    End If

    Dim meldung As String
    meldung = "Sammeldatei erstellt: " & anzahl & " " & artText & " von " & anzLehrer & " Lehrer(n)." & vbLf & vbLf & ausgabePfad
    If exitCode = 2 Then
        meldung = meldung & vbLf & vbLf & "ACHTUNG - nicht alle Dateien konnten " & Chr(252) & "bernommen werden:" & vbLf & _
                  LogLesen(logPfad, True)
    End If
    If MsgBox(meldung & vbLf & vbLf & "Sammeldatei jetzt " & Chr(246) & "ffnen?", _
              vbYesNo + IIf(exitCode = 2, vbExclamation, vbInformation), "LBV-Sammeldatei") = vbYes Then
        On Error Resume Next
        ThisWorkbook.FollowHyperlink ausgabePfad
        On Error GoTo 0
    End If

End Sub

' ----------------------------------------------------------------
' Hilfsfunktion: Kuerzel aus Gesamt-Dateinamen extrahieren
' Schema: LBV_Mehrarbeit_<Kuerzel>_Gesamt.pdf
' ----------------------------------------------------------------
Private Function KuerzelAusGesamtname(ByVal dateiname As String) As String
    Const PRAEFIX As String = "lbv_mehrarbeit_"
    Const SUFFIX As String = "_gesamt.pdf"
    KuerzelAusGesamtname = ""
    If LCase(Left(dateiname, Len(PRAEFIX))) <> PRAEFIX Then Exit Function
    If LCase(Right(dateiname, Len(SUFFIX))) <> SUFFIX Then Exit Function
    If Len(dateiname) <= Len(PRAEFIX) + Len(SUFFIX) Then Exit Function
    KuerzelAusGesamtname = Mid(dateiname, Len(PRAEFIX) + 1, Len(dateiname) - Len(PRAEFIX) - Len(SUFFIX))
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Hilfsprogramm suchen (neben lbv_fill.exe, dann neben XLSM)
' ----------------------------------------------------------------
Private Function HilfsdateiSuchen(ByVal skriptOrdner As String, ByVal dateiname As String) As String
    Dim treffer As String
    HilfsdateiSuchen = ""
    On Error Resume Next
    If skriptOrdner <> "" Then
        treffer = ""
        treffer = Dir(skriptOrdner & dateiname)
        If treffer <> "" Then
            HilfsdateiSuchen = skriptOrdner & dateiname
            Exit Function
        End If
    End If
    treffer = ""
    treffer = Dir(ThisWorkbook.Path & "\" & dateiname)
    If treffer <> "" Then HilfsdateiSuchen = ThisWorkbook.Path & "\" & dateiname
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Python-Befehl ermitteln ("py", "python" oder "")
' ----------------------------------------------------------------
Private Function PythonBefehl(wsh As Object) As String
    PythonBefehl = ""
    If wsh.Run("cmd /c py --version > nul 2>&1", 0, True) = 0 Then
        PythonBefehl = "py"
    ElseIf wsh.Run("cmd /c python --version > nul 2>&1", 0, True) = 0 Then
        PythonBefehl = "python"
    End If
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Log-Datei lesen (max. 15 Zeilen)
' nurFehler = True: nur Zeilen mit FEHLER/ACHTUNG
' ----------------------------------------------------------------
Private Function LogLesen(ByVal logPfad As String, ByVal nurFehler As Boolean) As String
    Dim txt As String: txt = ""
    Dim zeile As String
    Dim n As Long: n = 0
    LogLesen = ""
    If Dir(logPfad) = "" Then Exit Function
    Dim f As Integer: f = FreeFile
    Open logPfad For Input As #f
    Do While Not EOF(f)
        Line Input #f, zeile
        If Not nurFehler Or InStr(zeile, "FEHLER") > 0 Or InStr(zeile, "ACHTUNG") > 0 Then
            If n < 15 Then txt = txt & zeile & vbLf
            n = n + 1
        End If
    Loop
    Close #f
    If n > 15 Then txt = txt & "... (" & (n - 15) & " weitere Zeilen in " & logPfad & ")" & vbLf
    LogLesen = txt
End Function

' ----------------------------------------------------------------
' Hilfsfunktion: Text JSON-konform escapen (Umlaute -> \u00xx)
' ----------------------------------------------------------------
Private Function JsonText(ByVal s As String) As String
    Dim i As Long, c As Long
    Dim r As String: r = ""
    For i = 1 To Len(s)
        c = AscW(Mid(s, i, 1))
        If c < 0 Then c = c + 65536
        Select Case c
            Case 92:           r = r & "\\"
            Case 34:           r = r & "\" & Chr(34)
            Case 0 To 31:      r = r & "\u" & Right("0000" & Hex(c), 4)
            Case Is > 126:     r = r & "\u" & Right("0000" & Hex(c), 4)
            Case Else:         r = r & Mid(s, i, 1)
        End Select
    Next i
    JsonText = r
End Function
