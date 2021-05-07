Attribute VB_Name = "Module1"
Sub update_prices()
    ' take current table, loop over rows
    ' switch to other open table, search for article number
    ' copy price, switch back to first table and paste price
    ' update "Code" column to log changes
    ' -1 = could not be found = no change,
    ' 2 = found and in % range = changed,
    ' 3 = found but not in % range = no change,
    ' 4 = found but new price is lower = no change
    '
    ' Keyboard shortcut = Strg + d
    ' from JP 20210507
    
    ' for debugging
    Dim debugBool As Boolean
    debugBool = False
    debugBoolS = True ' debug bool for single problems
    
    ' declare iRows for loop over rows
    Dim x As Integer
    
    ' Set numrows = number of rows of data.
    'NumRows = Range("A2", Range("A2").End(xlDown)).Rows.Count
    rowStart = 1000 ' where to start looping ' first row = 28
    NumRows = 4091 ' to where to loop ' last row = 4091
    ' for debug: 1220 - 1224
    
    iRow = rowStart
    rowCode = 16 ' column for log ' P = 16
    rowVK = 4 ' column to paste price ' D = 4
    ' nope rowVKtoCpy = 6 ' column to copy price from ' not used because of dynamic offset from active cell from search result
    
    ' declare String searchTerm
    Dim searchTerm As String
    searchTerm = "asdf" ' TODO blind declaration necessary?
    
    ' declare String result for search result
    Dim searchResult As String
    searchResult = "no" ' TODO blind declaration necessary?
    
    'declare interval for prozent range
    Dim minProz As Double
    Dim maxProz As Double
    minProz = 0.0278 '2.78
    maxProz = 0.0436 '4.36
    
    ' ### loop over rows
    ' Establish "For" loop to loop "numrows" number of times.
    For iRows = rowStart To NumRows ' TODO eliminate one of iRows and rowStart?
    
        ' ### take ArtNr as searchTerm
        artNum = Cells(iRow, 1)
        searchTerm = artNum
        If debugBool Then MsgBox ("Row " + Str(iRows) + ": " + artNum)
        ' take first 5 digits (but 6 because leading space
        searchTerm = Left(searchTerm, 6)
        ' get rid of leading space
        searchTerm = Right(searchTerm, 5)
        
        ' ### switch to other table
        ActiveWindow.ActivateNext
        
        ' ### search for searchTerm
        ' MsgBox (TypeName(searchTerm))
        Set foundCell = Cells.Find(What:=searchTerm, LookIn:=xlFormulas)
        
        ' test if returned object is Range, if Nothing then no search result
        If foundCell Is Nothing Then ' foundCell is Nothing in case no search result found
            ' ### mark cell as not found
            ' switch back to first table
            ActiveWindow.ActivateNext
            ' ### mark cell as done, log with -1
            Cells(iRow, rowCode) = "-1"
            
        Else ' = search result found
            foundCell.Activate
            
            ' ### offset to price column
            ActiveCell.Offset(0, 3).Select
            ' ### copy price
            newPrice = ActiveCell.Value
            
            ' ### switch back to first table
            ActiveWindow.ActivateNext
            
            ' get old price for comparison
            oldPrice = Cells(iRow, rowVK)
            
            ' ### check if new price is higher than old price
            ' and only update price if higher
            If newPrice <= oldPrice Then
                ' mark with "4"
                Cells(iRow, rowCode) = "4"
            Else ' new price is higher than old price
                ' TODO switch case wie viel teurer newPrice ist
                ' zwischen 2,78 - 4,36 % Preiserhöhung übernehmen
                ' darunter und darüber eigenen Log Code geben
                diff = newPrice - oldPrice
                proz = diff / oldPrice
                'If debugBoolS Then MsgBox ("diff = " + Str(diff))
                If debugBool Then MsgBox ("proz = " + Str(proz))
                
                ' ### paste price
                ' prüfen of Erhöhung in gewisser Prozent Spanne ist
                If minProz < proz And proz <= maxProz Then
                    'MsgBox (TypeName(minProz) + TypeName(proz) + TypeName(maxProz))
                    'MsgBox (Str(minProz) + "<" + Str(proz) + "<=" + Str(maxProz))
                    If debugBool Then MsgBox "in % range"
                    'Cells(iRow, rowVK) = newPrice ' only update price if new price is higher
                    Cells(iRow, 17) = newPrice
                    Cells(iRow, 18) = proz
                    
                    ' ### mark cell as done with 2
                    Cells(iRow, rowCode) = "2"
                Else
                    Cells(iRow, 17) = newPrice
                    Cells(iRow, 18) = proz
                    ' ### mark cell as done with 3
                    Cells(iRow, rowCode) = "3"
                End If
                
            End If
            
        ' Activate Cell to see Progress
        Cells(iRow, rowCode).Activate
        
        End If
        
        ' ### iterate
        iRow = iRow + 1
    Next
End Sub