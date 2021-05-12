Sub update_EKs()
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
    ' from JP 20210511
    
    ' ----------------------------------README----------------------------------
    ' to do before use:
    
    ' make sure columns line up:
    ' reveicing table needs:
    ' 1st column: article number
    ' 4th column: price
    columnVK = 8 ' column to paste price ' D = 4, H = 8
    
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
    ' add column Q = 17 for zs = zwischenspeicher,
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
    rowStart = 3000 ' where to start looping ' first row = 28
    NumRows = 4091 ' to where to loop ' last row = 4091
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
    
    ' declare String searchTerm
    Dim searchTerm As String
    
    ' declare String result for search result
    Dim searchResult As String
    
    iRow = rowStart
    ' ### loop over rows
    ' Establish "For" loop to loop "numrows" number of times.
    For iRows = rowStart To NumRows
    
        ' ### take ArtNr as searchTerm
        artNum = Cells(iRow, 1)
        searchTerm = artNum
        If debugBoolS Then MsgBox ("Row " + Str(iRows) + ": " + artNum)
        ' take first 12 digits
        searchTerm = Left(searchTerm, 12)
        
        If debugBoolS Then MsgBox searchTerm
        
        ' ### switch to other table
        ActiveWindow.ActivateNext
        
        ' ### search for searchTerm
        Set foundCell = Cells.Find(What:=searchTerm, LookIn:=xlFormulas)
        
        ' test if returned object is Range, if Nothing then no search result
        If foundCell Is Nothing Then ' foundCell is Nothing in case no search result found
            ' switch back to first table
            ActiveWindow.ActivateNext
            ' ### mark cell as not found, log with -1
            If debugBoolColumns Then Cells(iRow, columnLog) = "-1"
            
        Else ' = search result found
            foundCell.Activate
            
            ' ### offset to price column & copy price
            ActiveCell.Offset(0, 8).Select
            newPrice = ActiveCell.Value
            
            ' ### switch back to first table
            ActiveWindow.ActivateNext
            
            ' get old price for comparison
            oldPrice = Cells(iRow, columnVK)
            
            ' ### check if new price is higher than old price
            ' and only update price if higher
            If newPrice <= oldPrice Then
                If debugBoolColumns Then Cells(iRow, columnZW) = newPrice
                ' mark with "4"
                If debugBoolColumns Then Cells(iRow, columnLog) = "4"
            Else ' new price is higher than old price
                ' wie viel teurer newPrice ist
                ' zwischen 2,78 - 4,36 % Preiserhöhung übernehmen
                ' darunter und darüber eigenen Log Code geben
                diff = newPrice - oldPrice
                proz = diff / oldPrice
                'If debugBoolS Then MsgBox ("diff = " + Str(diff))
                If debugBool Then MsgBox ("proz = " + Str(proz))
                
                ' ### paste price
                ' prüfen of Erhöhung in gewisser Prozent Spanne ist
                If minProz < proz And proz <= maxProz Then
                    If debugBool Then MsgBox "in % range"
                    If Not debugBoolColumns Then Cells(iRow, columnVK) = newPrice ' only update price if new price is higher
                    If debugBoolColumns Then Cells(iRow, columnZW) = newPrice
                    If debugBoolColumns Then Cells(iRow, columnProz) = proz
                    
                    ' mark that row has changed
                    ' 1 = VK was changed
                    ' 2 = EK has changed
                    ' 3 = VK + EK has changed
                    If Cells(iRow, columnChanged) = "1" Then
                        Cells(iRow, columnChanged) = "3"
                    Else
                        Cells(iRow, columnChanged) = "2"
                    End If
                    
                    ' ### mark cell as done with 2
                    If debugBoolColumns Then Cells(iRow, columnLog) = "2"
                Else
                    If debugBoolColumns Then Cells(iRow, columnZW) = newPrice
                    If debugBoolColumns Then Cells(iRow, columnProz) = proz
                    ' ### mark cell as done with 3
                    If debugBoolColumns Then Cells(iRow, columnLog) = "3"
                End If
                
            End If
            
        End If
        
        ' Activate Cell to see Progress
        Cells(iRow, columnVK).Activate
        If debugBoolColumns Then Cells(iRow, columnLog).Activate
        
        ' ### iterate
        iRow = iRow + 1
    Next
End Sub