Sub KN_extract()
' bekommt Tabelle mit Kontakten
' fügt ganz vorne eine neue Spalte ein falls noch nicht da| [TODO] vorne oder hinten?
' geht jede Zeile durch, sucht die Zelle mit "Kundennummer"
' und kopiert die Nummer, die in der Zelle rechts daneben steht in die erste Spalte
'
' UND wenn kein custom field gefunden mit Kundennummer
' -> sucht in Notizen Feld nach Kundennummer und fischt sie dort raus

Dim debugBool As Boolean
debugBool = False

' check if first column is Name column
' if so, add column and name it "Kundennummer"
columnName = Cells(1, 1).Value
If InStr(columnName, "Name") > 0 Then
    Cells(1, 1).EntireColumn.Insert
    Cells(1, 1) = "Kundennummer"
    columnName = Cells(1, 1).Value
End If

' define column with first custom field for search for Kundennummer
' range for search will then be from customColumn to lastColumn
customColumn = Range(Cells(1, 1), Cells(1, 200)).Find("Custom Field 1 - Type").Column
lastColumn = 200

lastRow = 5985

For iRow = 3 To lastRow
    If debugBool Then MsgBox ("searching row " + Str(iRow))
    Cells(iRow, 168).Activate
    Set r = Range(Cells(iRow, customColumn), Cells(iRow, lastColumn)).Find("Kundennummer")
    If r Is Nothing Then
        ' search through notes cell for Type cell "Kundennummer"
        
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
            If debugBool Then MsgBox IsNumeric(Mid(notes, startChar, 6))
            If IsNumeric(Mid(notes, startChar, 6)) Then Cells(iRow, 1) = Mid(notes, startChar, 6)
            
            ' TODO use exit to avoid nesting ifs?
            
        End If
    Else ' = searchterm found, result = r
        ' take found Type cell and copy Value cell to KN cell
        Cells(iRow, 1) = Cells(iRow, r.Column + 1) ' copy KN to first column
    End If
Next

Cells(1, 1).Activate

End Sub