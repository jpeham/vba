Sub KN_extract()
' bekommt Tabelle mit Kontakten
' fügt ganz vorne eine neue Spalte ein | [TODO] vorne oder hinten?
' geht jede Zeile durch, sucht die Zelle mit "Kundennummer"
' und kopiert die Nummer, die in der Zelle rechts daneben steht in die erste Spalte
'
' UND wenn kein custom field gefunden mit Kundennummer
' -> sucht in Notizen Feld nach Kundennummer und fischt sie dort raus

Dim debugBool As Boolean
debugBool = False

' only do this one time
Dim firstTimeBool As Boolean
firstTimeBool = False
'
If firstTimeBool Then Cells(1, 1).Activate
If firstTimeBool Then Selection.EntireColumn.Insert

customColumn = Range(Cells(1, 1), Cells(1, 200)).Find("Custom Field 1 - Type").Column
lastColumn = 200

If True Then
    For iRow = 3 To 10
        If debugBool Then MsgBox ("searching row " + Str(iRow))
        Cells(iRow, 168).Activate
        Set Suchbereich = Range(Cells(iRow, customColumn), Cells(iRow, lastColumn)).Find("Kundennummer")
        If Suchbereich Is Nothing Then
            ' search through notes cell for "Kundennummer"
            
            Dim notes As String
            notes = Cells(iRow, 27).Value ' column 27 = Notizen
            If debugBool Then MsgBox TypeName(notes)
            If debugBool Then MsgBox notes
            
            ' check if "Kundennummer" is in Notes
            If InStr(notes, "Kundennummer") Then ' if not 0  | if found
                startChar = InStr(notes, "Kundennummer") ' Type: long
                startChar = startChar + 12 ' for the word "Kundennummer" = 12 chars
                
                If debugBool Then MsgBox Mid(notes, startChar, 6)
                ' check if first char after  Kundennummer = ":"
                If InStr(Mid(notes, startChar, 1), ":") Then startChar = startChar + 1
                If debugBool Then MsgBox Mid(notes, startChar, 6)
                ' check if next char = " "
                If InStr(Mid(notes, startChar, 1), " ") Then startChar = startChar + 1
                If debugBool Then MsgBox Mid(notes, startChar, 6)
                
                ' take following 6 chars if 0-9
                ' If Regex.IsMatch(number, "^[0-9 ]+$") Then
                ' IsNumeric(number)
                If debugBool Then MsgBox IsNumeric(Mid(notes, startChar, 6))
                If IsNumeric(Mid(notes, startChar, 6)) Then Cells(iRow, 1) = Mid(notes, startChar, 6)
                
                ' TODO use exit to avoid nesting ifs?
                
            End If
        Else ' = searchterm found
            Cells(iRow, 1) = Cells(iRow, 169) ' copy KN to first column
        End If
    Next
End If

End Sub