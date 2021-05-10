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
    
    ' ----------------------------------README----------------------------------
    ' to do before use:
    
    ' make sure columns line up:
    ' reveicing table needs:
    ' 1st column: article number
    ' 4th column: price
    columnVK = 4 ' column to paste price ' D = 4
    
    columnChanged = 14 ' column N ' to log if changed for 2021
    
    ' if wanted, id for sorting, log for debug
    ' add column O = 15 for id, and iterate to give id
    ' add column P = 16 for logs
    columnLog = 16 ' column for log ' P = 16
    ' debug codes:
    ' -1 = Artikelnummer in Suche nicht gefunden
    '  2 = ArtNr gefunden und aktualisiert, weil nicht billiger und in %-Rahmen
    '  3 = ArtNr gefunden aber außerhalb von %-Rahmen (2,78% - 4,36%), also nicht aktualisiert
    '  4 = ArtNr gefunden aber billiger als aktueller Preis, also nicht aktualisiert
    '
    Dim debugBoolColumns As Boolean
    debugBoolColumns = False
    ' if debugBoolColumns is true then:
    ' add column Q = 17 for zw = zwischenspeicher,
    ' used here for newPrice before overwriting oldPrice
    columnZW = 17
    ' add column R = 18 for percent calculations
    columnProz = 18
    '
    ' make sure only 2 tables are open:
    ' and activate receiving table
    '
    ' Set numrows = number of rows of data.
    'NumRows = Range("A2", Range("A2").End(xlDown)).Rows.Count
    rowStart = 28 ' where to start looping ' first row = 28
    NumRows = 40 ' to where to loop ' last row = 4091
    '
    'declare interval for prozent range
    Dim minProz As Double
    Dim maxProz As Double
    minProz = 0.0278 '2.78
    maxProz = 0.0436 '4.36
    
    ' ---------------------------------------------------------------------------
    
    ' for debugging
    Dim debugBool As Boolean
    debugBool = False
    debugBoolS = False ' debug bool for single problems
    
    ' declare iRows for loop over rows
    Dim x As Integer 'TODO still in use?
    
    ' declare String searchTerm
    Dim searchTerm As String
    searchTerm = "asdf" ' TODO blind declaration necessary?
    
    ' declare String result for search result
    Dim searchResult As String
    searchResult = "no" ' TODO blind declaration necessary?
    
    iRow = rowStart
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
            If debugBoolColumns Then Cells(iRow, columnLog) = "-1"
            
        Else ' = search result found
            foundCell.Activate
            
            ' ### offset to price column
            ActiveCell.Offset(0, 3).Select
            ' ### copy price
            newPrice = ActiveCell.Value
            
            ' ### switch back to first table
            ActiveWindow.ActivateNext
            
            ' get old price for comparison
            oldPrice = Cells(iRow, columnVK)
            
            ' ### check if new price is higher than old price
            ' and only update price if higher
            If newPrice <= oldPrice Then
                ' mark with "4"
                If debugBoolColumns Then Cells(iRow, columnLog) = "4"
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
                    If Not debugBoolColumns Then Cells(iRow, columnVK) = newPrice ' only update price if new price is higher
                    If debugBoolColumns Then Cells(iRow, columnZW) = newPrice
                    If debugBoolColumns Then Cells(iRow, columnProz) = proz
                    
                    ' mark that row has changed
                    Cells(iRow, columnChanged) = "1"
                    
                    ' ### mark cell as done with 2
                    If debugBoolColumns Then Cells(iRow, columnLog) = "2"
                Else
                    If debugBoolColumns Then Cells(iRow, columnZW) = newPrice
                    If debugBoolColumns Then Cells(iRow, columnProz) = proz
                    ' ### mark cell as done with 3
                    If debugBoolColumns Then Cells(iRow, columnLog) = "3"
                End If
                
            End If
            
        ' Activate Cell to see Progress
        Cells(iRow, columnVK).Activate
        If debugBoolColumns Then Cells(iRow, columnLog).Activate
        
        End If
        
        ' ### iterate
        iRow = iRow + 1
    Next
End Sub
