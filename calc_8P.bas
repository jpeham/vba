Sub calc_8P()
'
' calc_8P Makro
' Einfügen von Zeile und Berechnung von 8%
'
' Tastenkombination: Strg+d
' from JP 20210526
'
' mit xlDown bis ans Ende des Blocks gehen
' Zeile einfügen
' function mit Summe und plus 8% einfügen
' bzw direkt mit Referenz auf Zelle links daneben? TODO
' Zellen mergen und Text dazwischen einfügen

' for debugging
Dim debugBool As Boolean
debugBool = True
Dim debugBoolCells As Boolean
debugBoolCells = False
Dim debugBoolFuture As Boolean
debugBoolFuture = True

' Set Suchbereich = Range("F:F").Find("total")
' Suchbereich.Select

' save startCell
If debugBoolCells Then MsgBox ActiveCell.Row
startColumn = ActiveCell.Column
startRow = ActiveCell.Row

Selection.End(xlDown).Select 'jump down to last line of block
' save endCell
If debugBoolCells Then MsgBox ActiveCell.Row
endColumn = ActiveCell.Column
endRow = ActiveCell.Row

If debugBoolCells Then MsgBox ("start: " + Str(startColumn) + Str(startRow) + ", end: " + Str(endColumn) + Str(endRow))

ActiveCell.Offset(1, 0).Range("A1").Select ' go down 1 to insert above
Selection.EntireRow.Insert ' insert row above activeCell
ActiveCell.Offset(0, -3).Range("A1").Select ' go left 3
ActiveCell.FormulaR1C1 = "8%"
ActiveCell.Offset(0, 1).Range("A1").Select ' go right 1
tempRow = ActiveCell.Row
tempColumn = ActiveCell.Column
Range(Cells(tempRow, tempColumn), Cells(tempRow, tempColumn + 1)).Merge
ActiveCell.FormulaR1C1 = "add text here"
ActiveCell.Offset(0, 1).Range("A1").Select ' go right 1
ActiveCell.Formula = "=SUM(F" & startRow & ":F" & endRow & ")*8%"

ActiveSheet.Cells(startRow, startColumn).Select

End Sub