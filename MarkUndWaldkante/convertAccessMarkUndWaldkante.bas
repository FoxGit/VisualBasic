''======================================================================================================
'' Programmm:       ConvertAccessMarkUndWaldkante
'' Beschreibung:    Extrahiert Daten für Mark, Marknähe, Waldkante und Splint aus einer Hemmenhofen-
''                  tabelle und schreibt sie einzelnd in einer neue. Schreibt zu jedem Fund wahlweise
''                  die Nummern oder wenn vorhanden die DG ODER nur die Nummern auf.
''======================================================================================================

Sub ConvertAccessMarkUndWaldkante()
    Dim minStartDate As Long
    Dim maxEndDate As Long
    Dim DCName As String
    Dim markWaldkanteSheetName As String
    Dim markWaldkanteSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim dateDiff As Long
    Dim nummernAndDG As Boolean

    'Nummern UND DG eintragen?
    nummernAndDG = False


    DCName = "DC"
    markWaldkanteSheetName = "Waldkante"

    'Quelltabelle vorhanden?
    If sheetExists(DCName) Then
        Set sourceSheet = Worksheets(DCName)
    End If

    If sourceSheet Is Nothing Then
        MsgBox "Tabelle " & DCName & "gefunden. ;(..."
        End
    End If

    sourceSheet.Activate

    minStartDate = getMinStartDate(sourceSheet)
    maxEndDate = getMaxEndDate(sourceSheet)

    dateDiff = maxEndDate - minStartDate

    Set markWaldkanteSheet = insertNewTable(dateDiff + 2, minStartDate, markWaldkanteSheetName)

    sourceSheet.Activate

    Call insertValues(dateDiff, sourceSheet, markWaldkanteSheet, nummernAndDG)

    markWaldkanteSheet.Activate

End Sub

''======================================================================================================
'' Funktion:        insertValues
'' Beschreibung:    Konvertiert Daten der Mark- und WaldkanteSpalte und schreibt sie in die Zieltabelle.
'' Parameter:       diff (long) - Anzahl der Jahre zwischen dem
''                  geringstes Startjahr und dem höchstes Endjahr der
''                  Quelltabelle
''                  sourceSheet (Worksheet) - Quelltabelle (DC oder DG)
''                  markWaldkanteSheet (Worksheet) - Zieltabelle (markWaldkanteSheet)
''                  nummernAndDG - Flag, ob nur Nummern oder Nummern und DG eingetragen werden
''======================================================================================================
Function insertValues(diff As Long, sourceSheet As Worksheet, markWaldkanteSheet As Worksheet, nummernAndDG As Boolean)
    Dim anfangsjahrString As String
    Dim markString As String
    Dim datierungString As String
    Dim nummerString As String
    Dim ortscodeString As String
    Dim getEndDat1e As Double
    Dim jahr As Long
    Dim anfangsjahrCell As Range
    Dim datierungCell As Range
    Dim markCell As Range
    Dim nummerCell As Range
    Dim ortscodeCell As Range
    Dim startJahr As Variant
    Dim waldkanteMatch As Variant
    Dim splintMatch As Variant
    Dim counter As Long
    Dim ortsCode As String
    Dim ortsCodeEndung As String
    Dim markNummer As String
    Dim splintNummer As String
    Dim sourceNummer As String
    Dim sourceDG As String
    Dim mark As Long
    Dim marknaehe As Long
    Dim waldkante As Long
    Dim splint As Long
    Dim DGSearchString As String
    Dim DGValueCell As Range
    Dim splintRow As Variant
    Dim splintTempCell As Range
    Dim nummerSearchString As String
    Dim nummerValueCell As Range
    Dim numberOfWerte As Long
    Dim sourcejahr As Long
    Dim markRow As Variant
    Dim destAnfangsjahrRange As Range
    Dim markTempCell As Range
    Dim destAnfangsjahrFound As Boolean
    Dim marknaeheNummer As String
    Dim waldkanteNummer As String
    Dim markWaldkanteTemp As String
    Dim marknaeheTemp As String
    Dim waldkanteRow As Variant
    Dim waldkanteTempCell As Range
    Dim waldkanteTemp As String
    Dim splintTemp As String


    'entsprechende Werte finden
    anfangsjahrString = "Anfangsjahr"
    'Spalte mit dem Namen Werte finden
    Set anfangsjahrCell = sourceSheet.Rows(1).Find(What:=anfangsjahrString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    markString = "Mark"
    'Spalte mit dem Namen Mark finden
    Set markCell = sourceSheet.Rows(1).Find(What:=markString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    datierungString = "Datierung"
    'Spalte mit dem Namen Datierung finden
    Set datierungCell = sourceSheet.Rows(1).Find(What:=datierungString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    nummerString = "Nummer"
    'Spalte mit dem Namen Nummer finden
    Set nummerCell = sourceSheet.Rows(1).Find(What:=nummerString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    ortscodeString = "Ortscode"
    'Spalte mit dem Namen Nummer finden
    Set ortscodeCell = sourceSheet.Rows(1).Find(What:=ortscodeString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'Spalte mit dem Namen Nummer finden
    nummerSearchString = "Nummer"
    Set nummerValueCell = sourceSheet.Rows(1).Find(What:=nummerSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'Spalte mit dem Namen DG finden
    DGSearchString = "DG"
    Set DGValueCell = sourceSheet.Rows(1).Find(What:=DGSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'zählen, wie viele Werte Zeilen es gibt, über die iterieren wir
    numberOfWerte = WorksheetFunction.Count(nummerValueCell.EntireColumn)
    'wir iterieren über die Werte der Quelltabelle
    For counter = 2 To numberOfWerte + 1

        'Anfangsjahres der aktuellen Zeile der Quelltabelle finden
        sourcejahr = Cells(counter, anfangsjahrCell.Column).Value

        'Reihe des Anfangsjahres in der Zieltabelle finden
        If sourcejahr <> 0 Then
            destAnfangsjahrFound = False
            Set destAnfangsjahrRange = markWaldkanteSheet.Range("A:A")

            'alle Zellen der Zieltabelle nach dem Anfangsjahr durchsuchen
            For Each markTempCell In destAnfangsjahrRange
                If markTempCell.Value = sourcejahr Then
                    destAnfangsjahrFound = True
                    markRow = markTempCell.Row
                End If
                'aus dem Loop springen, sobald der Wert gefunden wurde
                If destAnfangsjahrFound Then Exit For
            Next markTempCell

           'quelltabelle
            ortsCode = Cells(counter, ortscodeCell.Column).Value
            ortsCodeEndung = Mid(ortsCode, 6, Len(ortsCode))
            sourceNummer = Cells(counter, nummerCell.Column).Value
            sourceDG = Cells(counter, DGValueCell.Column).Value

            'zieltabelle
            markNummer = markWaldkanteSheet.Cells(markRow, 6).Value
            marknaeheNummer = markWaldkanteSheet.Cells(markRow, 7).Value

            'M A R K
            'Mark und Marknähe anhand des aktuellen Startjahres in der Quelltabelle finden
            If Cells(counter, markCell.Column).Value = "M" Then
                mark = markWaldkanteSheet.Cells(markRow, 2).Value

                'Markdaten in der aktuellen Zeile hochzählen
                If mark = 0 Then
                    markWaldkanteSheet.Cells(markRow, 2).Value = 1
                Else
                    markWaldkanteSheet.Cells(markRow, 2).Value = mark + 1
                End If

                If sourceDG <> "----" And nummernAndDG Then
                    markWaldkanteTemp = sourceDG
                Else
                    markWaldkanteTemp = sourceNummer
                End If

                If Trim(markNummer & vbNullString) = vbNullString Then
                    markWaldkanteSheet.Cells(markRow, 6).Value = markWaldkanteTemp & " " & ortsCodeEndung
                Else
                    markWaldkanteSheet.Cells(markRow, 6).Value = markNummer & ", " & markWaldkanteTemp & " " & ortsCodeEndung
                End If
            End If

            'M A R K N A E H E
             If Cells(counter, markCell.Column).Value = "Mn" Then
                marknaehe = markWaldkanteSheet.Cells(markRow, 3).Value

                'Markdaten in der aktuellen Zeile hochzählen
                If marknaehe = 0 Then
                    markWaldkanteSheet.Cells(markRow, 3).Value = 1
                Else
                    markWaldkanteSheet.Cells(markRow, 3).Value = mark + 1
                End If

                If sourceDG <> "----" And nummernAndDG Then
                    marknaeheTemp = sourceDG
                Else
                    marknaeheTemp = sourceNummer
                End If

                If Trim(marknaeheNummer & vbNullString) = vbNullString Then
                    markWaldkanteSheet.Cells(markRow, 7).Value = marknaeheTemp & " " & ortsCodeEndung
                Else
                    markWaldkanteSheet.Cells(markRow, 7).Value = marknaeheNummer & ", " & marknaeheTemp & " " & ortsCodeEndung
                End If
            End If

            'W A L D K A N T E
            'Waldkantedaten und Splint anhand des aktuellen Startjahres in der Quelltabelle finden
            'W/S Zahl -> Zahl extrahieren -> +1 bei dem Jahr in der Zieltabelle
            If Left$(Cells(counter, datierungCell.Column).Value, 1) = "W" Then

                destAnfangsjahrFound = False

                'alle Zellen der Zieltabelle nach dem Anfangsjahr durchsuchen
                For Each waldkanteTempCell In destAnfangsjahrRange
                    If waldkanteTempCell.Value = CInt(Mid(Cells(counter, datierungCell.Column).Value, 3, Len(Cells(counter, datierungCell.Column).Value))) Then
                        destAnfangsjahrFound = True
                        waldkanteRow = waldkanteTempCell.Row
                    End If
                    'aus dem Loop springen, sobald der Wert gefunden wurde
                    If destAnfangsjahrFound Then Exit For
                Next waldkanteTempCell

                If Trim(waldkanteRow & vbNullString) <> vbNullString Then
                    'für gefundende Waldkantedaten Wert hochzählen, Wert in der Zieltabelle hochzählen
                    waldkante = markWaldkanteSheet.Cells(waldkanteRow, 4).Value
                    waldkanteNummer = markWaldkanteSheet.Cells(waldkanteRow, 8).Value

                    If sourceDG <> "----" And nummernAndDG Then
                        waldkanteTemp = sourceDG
                    Else
                        waldkanteTemp = sourceNummer
                    End If

                    If waldkante = 0 Then
                        markWaldkanteSheet.Cells(waldkanteRow, 4).Value = 1
                    Else
                        markWaldkanteSheet.Cells(waldkanteRow, 4).Value = waldkante + 1
                    End If

                    If Trim(waldkanteNummer & vbNullString) = vbNullString Then
                        markWaldkanteSheet.Cells(waldkanteRow, 8).Value = waldkanteTemp & " " & ortsCodeEndung
                    Else
                        markWaldkanteSheet.Cells(waldkanteRow, 8).Value = waldkanteNummer & ", " & waldkanteTemp & " " & ortsCodeEndung
                    End If
                End If
            End If

            ' S P L I N T
            If Left$(Cells(counter, datierungCell.Column).Value, 1) = "S" Then

                destAnfangsjahrFound = False

                'alle Zellen der Zieltabelle nach dem Anfangsjahr durchsuchen
                For Each splintTempCell In destAnfangsjahrRange
                    If splintTempCell.Value = CInt(Mid(Cells(counter, datierungCell.Column).Value, 3, Len(Cells(counter, datierungCell.Column).Value))) Then
                        destAnfangsjahrFound = True
                        splintRow = splintTempCell.Row
                    End If
                    'aus dem Loop springen, sobald der Wert gefunden wurde
                    If destAnfangsjahrFound Then Exit For
                Next splintTempCell

                If Trim(splintRow & vbNullString) <> vbNullString Then
                    'für gefundende Waldkantedaten Wert hochzählen, Wert in der Zieltabelle hochzählen
                    splint = markWaldkanteSheet.Cells(splintRow, 5).Value
                    splintNummer = markWaldkanteSheet.Cells(splintRow, 9).Value

                    If sourceDG <> "----" And nummernAndDG Then
                        splintTemp = sourceDG
                    Else
                        splintTemp = sourceNummer
                    End If

                    If splint = 0 Then
                        markWaldkanteSheet.Cells(splintRow, 5).Value = 1
                    Else
                        markWaldkanteSheet.Cells(splintRow, 5).Value = splint + 1
                    End If

                    If Trim(splintNummer & vbNullString) = vbNullString Then
                        markWaldkanteSheet.Cells(splintRow, 9).Value = splintTemp & " " & ortsCodeEndung
                    Else
                        markWaldkanteSheet.Cells(splintRow, 9).Value = splintNummer & ", " & splintTemp & " " & ortsCodeEndung
                    End If
                End If
            End If
        End If

NextIterationi:
    Next counter

End Function

''======================================================================================================
'' Funktion:        insertNewTable
'' Beschreibung:    Füllt die Jahrespalte von dem geringsten Startjahr bis zum größten Endjahr
'' Parameter:       minStartDate (long) - geringstes Startjahr der Quelltabelle
''                  maxEndDate (long) - höchstes Endjahr der Quelltabelle
''                  newSheetName (String) - Name der Zieltabelle
'' Rückgabe:        markWaldkanteSheet (Worksheet) - Zieltabelle (markWaldkanteSheet)
''======================================================================================================
Function insertNewTable(diff As Long, minStartDate As Long, newSheetName As String) As Worksheet
    Dim wks As Worksheet
    Dim Rng As Range
    Dim counter As Long
    Dim daten() As Integer

    'falls es schon eine newSheetName Tabelle gibt -> Löschen
    If sheetExists(newSheetName) Then
        'Warnfenster für das Löschen des Datenblattes deaktivieren
        Application.DisplayAlerts = False
        Set wks = Worksheets(newSheetName)
        wks.Delete
        Application.DisplayAlerts = True
    End If

    Set wks = Worksheets.Add(Worksheets(1))
    wks.Name = newSheetName

    'Größe des Datenarrays festlegen
    ReDim daten(diff)

    'Daten ins Array eintragen
    For counter = 0 To diff
        daten(counter) = minStartDate + counter
    Next counter

    'Daten ins Worksheet übertragen
    wks.Range("A2:A" & diff).Value = WorksheetFunction.Transpose(daten)

    'Spaltenüberschriften
    wks.Cells(1, 1).Value = "Jahr"
    wks.Cells(1, 2).Value = "Mark"
    wks.Cells(1, 3).Value = "Marknähe"
    wks.Cells(1, 4).Value = "Waldkante"
    wks.Cells(1, 5).Value = "Splint"
    wks.Cells(1, 6).Value = "Marknummer"
    wks.Cells(1, 7).Value = "Marknähenummer"
    wks.Cells(1, 8).Value = "Waldkantenummer"
    wks.Cells(1, 9).Value = "Splintnummer"

    Set insertNewTable = wks

End Function

''======================================================================================================
'' Funktion:        columnLetter
'' Beschreibung:    Sucht zu einer Spaltenzahl den entsprechenden
''                  Buchstaben
'' Parameter:       ColumnNumber (Long) - Spaltenzahl
'' Rückgabe:        Spaltenbuchstabe (String)
''======================================================================================================
Function columnLetter(ColumnNumber As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    columnLetter = s
End Function

''======================================================================================================
'' Funktion:        getMinStartDate
'' Beschreibung:    Sucht den geringsten Wert in der Anfangsjahrspalte
'' Parameter:       sourceSheet (Sheet) - Quelltabelle
'' Rückgabe:        geringsten Wert in der Anfangsjahrspalte (Long)
''======================================================================================================
Function getMinStartDate(sourceSheet) As Long

    Dim anfangsjahrSearchString As String
    Dim sourceCell As Range
    Dim columnLetterVar As String
    Dim c As Range
    Dim minStartDate As String

    anfangsjahrSearchString = "Anfangsjahr"

    'Spalte mit dem Namen Anfangsjahr finden
    Set sourceCell = sourceSheet.Rows(1).Find(What:=anfangsjahrSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'Buchstaben der Spalte finden
    columnLetterVar = columnLetter(sourceCell.Column)

    'kleinstes Anfangsjahr finden

    minStartDate = "MIN(IF(" & columnLetterVar & "2:" & columnLetterVar & Rows.Count & "<>0," & columnLetterVar & "2:" & columnLetterVar & Rows.Count & "))"
    getMinStartDate = Evaluate(minStartDate)
End Function

''======================================================================================================
'' Funktion:        getMaxEndDate
'' Beschreibung:    Sucht den höchsten Wert in der Endjahrspalte
'' Parameter:       sourceSheet (Sheet) - Quelltabelle
'' Rückgabe:        höchsten Wert in der Endjahrspalte (Long)
''======================================================================================================
Function getMaxEndDate(sourceSheet) As Long

    Dim endjahrSearchString As String
    Dim endjahrCell As Range
    Dim columnLetterVar As String
    Dim splintMatch As Variant
    Dim numberOfsplintDatierungen As Long
    Dim counter As Long
    Dim splintDatierungCell As Range
    Dim splintDatierungSearchString As String
    Dim splintDatierungString As String
    Dim maxSplintDatierungJahr As Long
    Dim maxEndDate As Long
    Dim splintDatierungen() As Variant
    Dim splintDatierungJahr As Long
    ReDim splintDatierungen(0 To 0)


    'die Jahre der Splint Datierung können bis zu 18 größer sein als die Endjahre
    'daher, wenn vorhanden und größer, Splint Datierung als größten Wert nehmen
    endjahrSearchString = "Endjahr"
    'Spalte mit dem Namen Endjahr finden
    Set endjahrCell = sourceSheet.Rows(1).Find(What:=endjahrSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    splintDatierungSearchString = "Datierung"
    'Spalte mit dem Namen Endjahr finden
    Set splintDatierungCell = sourceSheet.Rows(1).Find(What:=splintDatierungSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'zählen, wie viele Datierungs Zeilen es gibt, über die iterieren wir
    numberOfsplintDatierungen = WorksheetFunction.Count(splintDatierungCell.EntireColumn)

    'iterieren über die Datierungsspalte
    For counter = 2 To numberOfsplintDatierungen + 1
        'Werte der Datierungsspalte
        splintDatierungString = Cells(counter, splintDatierungCell.Column).Value

        'Markdaten in der aktuellen Zeile hochzählen
        If Mid(splintDatierungString, 1, 1) = "S" Then
            splintDatierungJahr = Mid(Cells(counter, splintDatierungCell.Column).Value, 3, Len(splintDatierungString) - 2)

            ' change / adjust the size of array
            ReDim Preserve splintDatierungen(0 To UBound(splintDatierungen) + 1) As Variant

            ' add value on the end of the array
            splintDatierungen(UBound(splintDatierungen)) = splintDatierungJahr
        Else
            'aus dem Loop springen
           GoTo NextIterationCounter
        End If
NextIterationCounter:
    Next counter


    'größtes Datierungsjahr finden
    maxSplintDatierungJahr = Application.WorksheetFunction.Max(splintDatierungen)
    'größtes Endjahr finden
    maxEndDate = Application.WorksheetFunction.Max(sourceSheet.Columns(endjahrCell.Column))

    getMaxEndDate = Application.WorksheetFunction.Max(maxSplintDatierungJahr, maxEndDate)
End Function

''======================================================================================================
'' Funktion:        sheetExists
'' Beschreibung:    Prüft, ob eine Tabelle existiert
'' Parameter:       sheetToFind (String) - gesuchte Tabelle
'' Rückgabe:        Tabelle existiert oder nicht (Boolean)
''======================================================================================================
Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
