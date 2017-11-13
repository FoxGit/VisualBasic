''======================================================================================================
'' Programmm:       ConvertAccessHemmenhofenToStandDyna
'' Beschreibung:    Konvertiert Accessdaten im "Hemmenhofen-Format" für die Benutzung von standDyna.
''                  Stellt außerdem die Belegungsdichte in einer eigenen Tabelle dar.
''======================================================================================================

Sub ConvertAccessHemmenhofenToStandDyna()
    Dim minStartDate As Long
    Dim maxEndDate As Long
    Dim sourceSheet As Worksheet
    Dim dynaSheet As Worksheet
    Dim belegungsdichteSheet As Worksheet
    Dim dynaSheetName As String
    Dim DCName As String
    Dim DGName As String
    Dim belegungsdichteSheetName As String
    Dim dateDiff As Long

    dynaSheetName = "standDyna"
    DCName = "DC"
    DGName = "DG"
    belegungsdichteSheetName = "Belegungsdichte"

    If sheetExists(DCName) Then
        Set sourceSheet = Worksheets(DCName)
    End If

    If sheetExists(DGName) Then
        Set sourceSheet = Worksheets(DGName)
    End If

    If sourceSheet Is Nothing Then
        MsgBox "Weder Tabelle " & DCName & " noch DG " & DCName & "gefunden. ;(..."
        End
    End If

    sourceSheet.Activate

    minStartDate = getMinStartDate(sourceSheet)
    maxEndDate = getMaxEndDate(sourceSheet)

    dateDiff = maxEndDate - minStartDate

    Set dynaSheet = insertNewTable(dateDiff + 2, minStartDate, dynaSheetName)

    Set belegungsdichteSheet = insertNewTable(dateDiff + 2, minStartDate, belegungsdichteSheetName)

    sourceSheet.Activate

    Call insertValues(dateDiff, sourceSheet, dynaSheet, belegungsdichteSheet)

    dynaSheet.Activate

End Sub


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

    Set insertNewTable = wks

End Function


''======================================================================================================
'' Funktion:        insertValues
'' Beschreibung:    Konvertiert Accessdaten im "Hemmenhofen-Format"
''                  für die Benutzung von standDyna
'' Parameter:       diff (long) - Anzahl der Jahre zwischen dem
''                  geringstes Startjahr und dem höchstes Endjahr der
''                  Quelltabelle
''                  sourceSheet (Worksheet) - Quelltabelle (DC oder DG)
''                  dynaSheet (Worksheet) - Zieltabelle (dynaSheet)
''                  belegungsdichteSheet (Worksheet) - Zieltabelle2 (Belegungsdichte)
''======================================================================================================
Function insertValues(diff As Long, sourceSheet As Worksheet, dynaSheet As Worksheet, belegungsdichteSheet As Worksheet)

    Dim anfangsjahrSearchString As String
    Dim werteSearchString As String
    Dim nummerSearchString As String

    Dim targetRow As Variant

    Dim dynaAnfangsjahrRange As Range
    Dim dynaTempCell As Range
    Dim dynaAnfangsjahrFound As Boolean

    Dim numberOfWerte As Long
    Dim columnCounter As Long
    Dim nummer As Long
    Dim sourcejahr As Long

    Dim errormessage As String

    Dim belegungTargetRow As Variant
    Dim belegungAnfangsjahrFound As Boolean
    Dim belegungsdichteAnfangsjahrRange As Range
    Dim belegungsdichteTempCell As Range

    'interessante Zellen finden
    'Spalte mit dem Namen Anfangsjahr finden
    anfangsjahrSearchString = "Anfangsjahr"
    Dim anfangsjahrValueCell As Range
    Set anfangsjahrValueCell = sourceSheet.Rows(1).Find(What:=anfangsjahrSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'Spalte mit dem Namen Werte finden
    werteSearchString = "Werte"
    Dim werteValueCell As Range
    Set werteValueCell = sourceSheet.Rows(1).Find(What:=werteSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'Spalte mit dem Namen Nummer finden
    nummerSearchString = "Nummer"
    Dim nummerValueCell As Range
    Set nummerValueCell = sourceSheet.Rows(1).Find(What:=nummerSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)


    columnCounter = 1
    'zählen, wie viele Werte Zeilen es gibt, über die iterieren wir
    numberOfWerte = WorksheetFunction.Count(nummerValueCell.EntireColumn)

    'wir beginnen in der Zeiltabelle erst in der zweiten Spalte
    For i = 2 To numberOfWerte + 1
        'Wert des Feldes Nummer in der derzeitigen Reihe holen
        nummer = Cells(i, nummerValueCell.Column).Value

        'Zeile des aktuellen Anfangsjahres finden
        sourcejahr = Cells(i, anfangsjahrValueCell.Column).Value

        'Reihe des Anfangsjahres in der Zieltabelle finden
        If sourcejahr <> 0 Then
            dynaAnfangsjahrFound = False
            Set dynaAnfangsjahrRange = dynaSheet.Range("A:A")

            'alle Zellen der Zieltabelle nach dem Anfangsjahr durchsuchen
            For Each dynaTempCell In dynaAnfangsjahrRange
                If dynaTempCell.Value = sourcejahr Then
                    dynaAnfangsjahrFound = True
                    targetRow = dynaTempCell.Row
                    columnCounter = columnCounter + 1
                End If
                'aus dem Loop springen, sobald der Wert gefunden wurde
                If dynaAnfangsjahrFound Then Exit For
            Next dynaTempCell

            'aus dem komischen Werte Feld eine Liste an Werten machen
            Dim values As String
            values = Cells(i, werteValueCell.Column).Value
            Dim WrdArray() As String
            WrdArray() = Split(values, vbCrLf)
            Dim anzahl As Long
            anzahl = UBound(WrdArray())
            'letztes Zeichen des letzten Elementes des Arrays löschen, weil es ein Zeilenumbruch/Leerzeichen oder sowas ist
            WrdArray(anzahl) = Left$(WrdArray(anzahl), Len(WrdArray(anzahl)) - 1)

            'fehlerhafte Werte (0 und 1) herausfiltern
            Dim withError As Boolean
            withError = False
            Dim element As Variant

            For Each element In WrdArray
                If element = 0 Or element = "1" Then
                    withError = True
                End If
            Next element

            If withError Then
                errormessage = errormessage + "Der Datensatz " & nummer & " enthält ungültige Werte (0 oder 1)" & vbNewLine
            Else
                'Werteliste einfügen
                dynaSheet.Range(dynaSheet.Cells(targetRow, columnCounter), dynaSheet.Cells(targetRow + anzahl, columnCounter)).Value = WorksheetFunction.Transpose(WrdArray)
                dynaSheet.Cells(1, columnCounter).Value = nummer

                'Eintrag für gefundene Werte in Belegungstabelle vornehmen
                belegungAnfangsjahrFound = False
                Set belegungsdichteAnfangsjahrRange = belegungsdichteSheet.Range("A:A")

                'alle Zellen der Zieltabelle nach dem Anfangsjahr durchsuchen
                For Each belegungsdichteTempCell In belegungsdichteAnfangsjahrRange
                    If belegungsdichteTempCell.Value = sourcejahr Then
                        belegungAnfangsjahrFound = True
                        belegungTargetRow = belegungsdichteTempCell.Row
                    End If
                    'aus dem Loop springen, sobald der Wert gefunden wurde
                    If belegungAnfangsjahrFound Then Exit For
                Next belegungsdichteTempCell

                Dim belegungsdichteRange As Range, currentCell As Range
                Set belegungsdichteRange = belegungsdichteSheet.Range(belegungsdichteSheet.Cells(targetRow, 2), belegungsdichteSheet.Cells(targetRow + anzahl, 2))
                For Each currentCell In belegungsdichteRange
                    currentCell.Value = currentCell.Value + 1
                Next
            End If
        End If
    Next i

    'Dummy Spalte generieren
    dynaSheet.Cells(1, columnCounter + 1).Value = "Dummy"
    Dim j As Integer

    For j = 2 To diff + 2
        dynaSheet.Cells(j, columnCounter + 1).Value = "1"
    Next j

    If errormessage <> vbNullString Then
        MsgBox errormessage
    End If

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

    anfangsjahrSearchString = "Anfangsjahr"

    'Spalte mit dem Namen Anfangsjahr finden
    Set sourceCell = sourceSheet.Rows(1).Find(What:=anfangsjahrSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'Buchstaben der Spalte finden
    columnLetterVar = columnLetter(sourceCell.Column)

    'kleinstes Anfangsjahr finden
    Dim c As Range
    Dim minStartDate As String
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
    Dim sourceCell As Range
    Dim columnLetterVar As String

    endjahrSearchString = "Endjahr"

    'Spalte mit dem Namen Anfangsjahr finden
    Set sourceCell = sourceSheet.Rows(1).Find(What:=endjahrSearchString, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    'größtes Endjahr finden
    getMaxEndDate = Application.WorksheetFunction.Max(sourceSheet.Columns(sourceCell.Column))
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
