''======================================================================================================
'' Programmm:       ConvertAccessMarkUndWaldkante
'' Beschreibung:    Extrahiert Daten für Mark, Marknähe, Waldkante und Splint aus einer Hemmenhofen-
''                  tabelle und schreibt sie einzelnd in einer neue.
''======================================================================================================

Sub ConvertAccessMarkUndWaldkante()
    Dim minStartDate As Long
    Dim maxEndDate As Long
    Dim DCName As String
    Dim markWaldkanteSheetName As String
    Dim markWaldkanteSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim dateDiff As Long

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

    Call insertValues(dateDiff, sourceSheet, markWaldkanteSheet)

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
''======================================================================================================
Function insertValues(diff As Long, sourceSheet As Worksheet, markWaldkanteSheet As Worksheet)
    Dim anfangsjahrString As String
    Dim markString As String
    Dim datierungString As String
    Dim getEndDat1e As Double
    Dim jahr As Long
    Dim anfangsjahrCell As Range
    Dim datierungCell As Range
    Dim markCell As Range
    Dim startJahr As Variant
    Dim waldkanteMatch As Variant
    Dim splintMatch As Variant
    Dim counter As Long

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

    'wir iterieren über die Jahre der Zieltabelle
    For counter = 1 To diff + 1

        'aktuelles Jahr
        jahr = markWaldkanteSheet.Cells(counter + 1, 1).Value

        'Zeile des aktuellen Startjahres finden
        startJahr = Application.Match(jahr, Columns(anfangsjahrCell.Column), 0)

        'gibt es einen Eintrag für das Startjahr in der Quelltabelle?
        If Not IsError(startJahr) Then

            'M A R K
            'Mark und Marknähe anhand des aktuellen Startjahres in der Quelltabelle finden
            If Cells(startJahr, markCell.Column).Value = "M" Then
                'Markdaten in der aktuellen Zeile hochzählen
                If IsEmpty(markWaldkanteSheet.Cells(counter + 1, 2).Value) Then
                    markWaldkanteSheet.Cells(counter + 1, 2).Value = 1
                Else
                    markWaldkanteSheet.Cells(counter + 1, 2).Value = markWaldkanteSheet.Cells(counter + 1, 2).Value + 1
                End If
            End If

            'M A R K N A E H E
            If Cells(startJahr, markCell.Column).Value = "Mn" Then
                'Markdaten in der aktuellen Zeile hochzählen
                If IsEmpty(markWaldkanteSheet.Cells(counter + 1, 3).Value) Then
                    markWaldkanteSheet.Cells(counter + 1, 3).Value = 1
                Else
                    markWaldkanteSheet.Cells(counter + 1, 3).Value = markWaldkanteSheet.Cells(counter + 1, 3).Value + 1
                End If
            End If
        End If

        'W A L D K A N T E
        'Waldkantedaten und Splint anhand des aktuellen Startjahres in der Quelltabelle finden
        'W/S Zahl -> Zahl extrahieren -> +1 bei dem Jahr in der Zieltabelle
        waldkanteMatch = Application.Match("W " & jahr, Columns(datierungCell.Column), 0)
        splintMatch = Application.Match("S " & jahr, Columns(datierungCell.Column), 0)
        If IsError(waldkanteMatch) And IsError(splintMatch) Then
             Debug.Print n
             GoTo NextIterationi
        Else
            Debug.Print Y
        End If

        'Waldkante setzen
        If Not IsError(waldkanteMatch) Then
            'für gefundende Waldkantedaten Wert hochzählen, Wert in der Zieltabelle hochzählen
            If IsEmpty(markWaldkanteSheet.Cells(counter + 1, 4).Value) Then
                markWaldkanteSheet.Cells(counter + 1, 4).Value = 1
            Else
                markWaldkanteSheet.Cells(counter + 1, 4).Value = markWaldkanteSheet.Cells(counter + 1, 4).Value + 1
            End If
        End If

        'Splint setzen
        If Not IsError(splintMatch) Then
            'für gefundende Splint Wert, Wert in der Zieltabelle hochzählen
            If IsEmpty(markWaldkanteSheet.Cells(counter + 1, 5).Value) Then
                markWaldkanteSheet.Cells(counter + 1, 5).Value = 1
            Else
                markWaldkanteSheet.Cells(counter + 1, 5).Value = markWaldkanteSheet.Cells(counter + 1, 5).Value + 1
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
    numberOfsplintDatierungen = 9

    'iterieren über die Datierungsspalte
    For counter = 2 To numberOfsplintDatierungen + 1
        'Werte der Datierungsspalte
        splintDatierungString = Cells(counter, splintDatierungCell.Column).Value

        'Markdaten in der aktuellen Zeile hochzählen
        If Mid(splintDatierungString, 1, 1) = "S" Then
            splintDatierungJahr = Mid(Cells(counter, splintDatierungCell.Column).Value, 3, Len(splintDatierungString) - 2)

            ' Array Größe anpassen
            ReDim Preserve splintDatierungen(0 To UBound(splintDatierungen) + 1) As Variant

            ' Jahr ins Array schieben
            splintDatierungen(UBound(splintDatierungen)) = splintDatierungJahr
        Else
            'aus dem Loop springen
           GoTo NextIterationCounter
        End If
NextIterationCounter:
    Next counter


    'größtes Datierungsjahr finden
    maxSplintDatierungJahr = Application.WorksheetFunction.max(splintDatierungen)
    'größtes Endjahr finden
    maxEndDate = Application.WorksheetFunction.max(sourceSheet.Columns(endjahrCell.Column))

    getMaxEndDate = Application.WorksheetFunction.max(maxSplintDatierungJahr, maxEndDate)
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

