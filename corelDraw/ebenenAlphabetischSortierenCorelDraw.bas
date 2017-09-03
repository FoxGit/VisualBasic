''======================================================================================================
'' Programmm:       Kurvenebenen_aufsteigendBuchstabenUndZahlen
'' Beschreibung:    Sortiert die Ebenen in Corel Draw im Objektemanager alphabetisch.
''                  Dabei werden NurZahlen nach unten sortiert.
'' Veränderungen----------------------------------------------------------------------------------------
'' Datum        Entwickler      Veränderung
'' 03.09.2017   Jan Hansen      Initialer Build
''======================================================================================================
Sub Kurvenebenen_aufsteigendBuchstabenUndZahlen()

Dim lyr As Layer
Dim counter As Long
Dim currentLayerName As String
Dim lastLayerName As String
Dim currentLayerNameLen As Integer
Dim lastLayerNameLen As Integer
Dim tempStringCounter As Integer
Dim currentLayerOnlyNumbers As Boolean
Dim lastLayerOnlyNumbers As Boolean
Dim currentLayerOnlyNumbersSameLength As Boolean
Dim lastLayerOnlyNumbersSameLength As Boolean
Dim currentLayerCharacterSameLength As String
Dim lastLayerCharacterSameLength As String
Dim currentLength As Integer
Dim currentLayerCharacter As String
Dim lastLayerCharacter As String
                                        
'iterieren über alle Ebenen
For Each lyr In ActivePage.Layers
    'nicht zu weit laufen
    If lyr.Index < ActivePage.Layers.Count Then
        Debug.Print lyr.Index
        'ja, bubblesort, weil wegen....
        For counter = 1 To lyr.Index
            currentLayerName = ActivePage.Layers(counter).Name
            currentLayerNameLen = Len(currentLayerName)
            lastLayerName = ActivePage.Layers(counter + 1).Name
            lastLayerNameLen = Len(lastLayerName)
            
            'string haben die gleiche Länge? Dann können wir sie fast normal vergleichen....
            If currentLayerNameLen = lastLayerNameLen Then
                currentLayerOnlyNumbersSameLength = True
                lastLayerOnlyNumbersSameLength = True
                
                'sind da welche nur aus Zahlen?
                For i = 1 To currentLayerNameLen
                    currentLayerCharacterSameLength = Mid(currentLayerName, i, 1)
                    lastLayerCharacterSameLength = Mid(lastLayerName, i, 1)
                    If Not IsNumeric(currentLayerCharacterSameLength) Then
                        currentLayerOnlyNumbersSameLength = False
                    End If
                    If Not IsNumeric(lastLayerCharacterSameLength) Then
                        lastLayerOnlyNumbersSameLength = False
                    End If
                Next
                'dann wird die Zahl nach unten verschoben.
                If currentLayerOnlyNumbersSameLength And Not lastLayerOnlyNumbersSameLength Then
                    Debug.Print "nur Nummern aktuelle " & currentLayerName & " tauschen mit " & lastLayerName
                    ActivePage.Layers(n).MoveBelow ActivePage.Layers(n + 1)
                ElseIf lastLayerOnlyNumbersSameLength And Not currentLayerOnlyNumbersSameLength Then
                    Debug.Print "nur Nummern untere, nichts machen"
                'keiner der beiden nur Zahlenwerte? normaler Vergleich :D
                ElseIf StrComp(currentLayerName, lastLayerName, vbTextCompare) = 1 Then
                    Debug.Print "gleiche Länge: " & currentLayerName & " is bigger than " & lastLayerName
                    ActivePage.Layers(n).MoveBelow ActivePage.Layers(n + 1)
                End If
            'Strings unterschiedlich lang? :O
            ElseIf currentLayerNameLen <> lastLayerNameLen Then
                'because Corel Draw kennt kein Math.min -.-
                If currentLayerNameLen < lastLayerNameLen Then
                    currentLength = currentLayerNameLen
                Else
                    currentLength = lastLayerNameLen
                End If
                
                currentLayerOnlyNumbers = True
                lastLayerOnlyNumbers = True
            
                'dann nehmen wir die Steinzeitmethode und vergleichen die einzelnen Charakters
                For i = 1 To currentLength
                    currentLayerCharacter = Mid(currentLayerName, i, 1)
                    lastLayerCharacter = Mid(lastLayerName, i, 1)
                    'sind es Buchstaben?
                    If Not IsNumeric(currentLayerCharacter) And Not IsNumeric(lastLayerCharacter) Then
                        'ist der aktuelle Buchstabe grüßer? -> sortieren
                        If currentLayerCharacter > lastLayerCharacter Then
                            Debug.Print "nur Buchstaben: " & currentLayerName & " ist größer als " & lastLayerName
                            ActivePage.Layers(n).MoveBelow ActivePage.Layers(n + 1)
                            currentLayerOnlyNumbers = False
                            lastLayerOnlyNumbers = False
                            Exit For
                        'ist er kleiner? -> nichts tun
                        ElseIf currentLayerCharacter < lastLayerCharacter Then
                            currentLayerOnlyNumbers = False
                            lastLayerOnlyNumbers = False
                            Exit For
                        End If
                    End If
                    'Charakter der aktuellen Zeile eine Zahl und der dadrunter ein Buchstabe? -> sortieren
                    If IsNumeric(currentLayerCharacter) And Not IsNumeric(lastLayerCharacter) Then
                        Debug.Print currentLayerName & " ist numerisch, daher größer als " & lastLayerName
                        lastLayerOnlyNumbers = False
                        ActivePage.Layers(n).MoveBelow ActivePage.Layers(n + 1)
                        Exit For
                    End If
                    'andersrum? -> nichts tun
                    If Not IsNumeric(currentLayerCharacter) And IsNumeric(lastLayerCharacter) Then
                        currentLayerOnlyNumbers = False
                        Exit For
                    End If
                Next i
                'beides nur zahlen? -> der längerer gewinnt und muss sortiert werden
                If currentLayerOnlyNumbers And lastLayerOnlyNumbers And currentLayerNameLen > lastLayerNameLen Then
                    Debug.Print "nur Nummern: aktueller " & currentLayerName & " länger als " & lastLayerName
                    ActivePage.Layers(n).MoveBelow ActivePage.Layers(n + 1)
                End If
            End If
        Next counter
    End If
Next lyr
  
End Sub

