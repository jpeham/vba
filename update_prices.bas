Attribute VB_Name = "Module1"
Sub update_prices()
    
    ' declare iRows for loop over rows
    Dim x As Integer
    ' Set numrows = number of rows of data.
    'NumRows = Range("A2", Range("A2").End(xlDown)).Rows.Count
    rowStart = 3000
    ' rowStart = 28
    NumRows = 4091
    ' 1224 4091
    iRow = rowStart
    rowCode = 20
    rowVK = 4
    ' nope rowVKtoCpy = 6
    
    ' declare String searchTerm
    Dim searchTerm As String
    searchTerm = "asdf"
    
    ' declare String result for search result
    Dim searchResult As String
    searchResult = "no"
    
    ' ### loop over rows
    ' Establish "For" loop to loop "numrows" number of times.
    For iRows = rowStart To NumRows
    
        ' Dim price As Sting
        
        ' ### take ArtNr as searchTerm
        ' MsgBox ("| ArtNr is " + Cells(iRow, 1))
        ' searchTerm = Str(Cells(iRow, 1))
        searchTerm = Cells(iRow, 1)
        ' MsgBox ("searchTerm = " + searchTerm)
        ' take first 5 digits (but 6 because leading space
        searchTerm = Left(searchTerm, 6)
        ' MsgBox ("left of searchTerm = >" + searchTerm + "<")
        ' get rid of leading space
        searchTerm = Right(searchTerm, 5)
        ' MsgBox ("left of searchTerm = >" + searchTerm + "<")
        ' MsgBox "Type of SearchTerm = " + TypeName(searchTerm)
        
        ' ### switch to other table
        ActiveWindow.ActivateNext
        
        ' ### search for searchTerm
        ' nope Set asdf0 = Cells.Find(What:="12412.31.001", LookIn:=xlFormulas)
        ' MsgBox (TypeName(searchTerm))
        Set foundCell = Cells.Find(What:=searchTerm, LookIn:=xlFormulas)
        
        ' test if returned object is Range, if Nothing then no search result
        If foundCell Is Nothing Then
            ' MsgBox "search result of " + searchTerm + " not found"
            ' ### mark cell as not found
            ' ### switch back to first table
            ActiveWindow.ActivateNext
            ' ### mark cell as done
            Cells(iRow, rowCode) = "6"
            
        Else
            ' MsgBox "searchResult = " + foundCell
            ' MsgBox ("Type of searchResult = " + TypeName(foundCell))
            ' MsgBox (foundCell)
            foundCell.Activate
            
            ' ### offset to price column
            ActiveCell.Offset(0, 3).Select
            ' ### copy price
            ' MsgBox (ActiveCell)
            Price = ActiveCell.Value
            ' MsgBox (price)
            
            ' ### switch back to first table
            ActiveWindow.ActivateNext
            
            ' ### paste price
            Cells(iRow, rowVK) = Price
            
            ' ### check if new price is higher than old price
            oldPrice = Cells(iRow, 5)
            If Price < oldPrice Then
                ' mark with "billiger"
                Cells(iRow, rowCode) = "B"
            Else
                ' ### mark cell as done with 2
                Cells(iRow, rowCode) = "2"
            End If
            
        ' Activate Cell to see Progress
        Cells(iRow, rowCode).Activate
        
        End If
        
        ' ### iterate
        iRow = iRow + 1
    Next
End Sub