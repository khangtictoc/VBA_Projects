Attribute VB_Name = "BillingServiceWorksheet"
Function CreateBillingServiceWorksheet(userWsName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim checkbox As Object
    Dim svcList As Variant
    svcList = Array("Phi Dien", _
        "Phi QL", _
        "Phi nuoc", _
        "Phi nuoc & XLNT", _
        "Nuoc uong", _
        "Phong", _
        "Xe")
    
    svcList_Length = UBound(svcList) - LBound(svcList) + 1
    
    ' Create a new worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Billing Service" & GetCurrentMonthAndYear()
    
    ' Add headers with bold text
    ws.Cells(1, 1).Value = "No. #"
    ws.Cells(1, 2).Value = "Billing Service"
    ws.Cells(1, 3).Value = "Price"
    ws.Cells(1, 4).Value = "Person In Charge (PIC)"
    ws.Rows(1).Font.Bold = True ' Make the headers bold

    ' Align the data to the center horizontally and vertically
    ws.Columns(1).HorizontalAlignment = xlVAlignCenter
    ws.Columns(2).HorizontalAlignment = xlVAlignCenter
    ws.Columns(3).HorizontalAlignment = xlVAlignCenter
    ws.Columns(4).HorizontalAlignment = xlVAlignCenter
    ws.Columns(1).VerticalAlignment = xlVAlignCenter
    ws.Columns(2).VerticalAlignment = xlVAlignCenter
    ws.Columns(3).VerticalAlignment = xlVAlignCenter
    ws.Columns(4).VerticalAlignment = xlVAlignCenter
    
    ' Set the column widths for better visibility
    ws.Columns(1).ColumnWidth = 10
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 15
    ws.Columns(4).ColumnWidth = 25

    ' Optional: Set the number format for the index column
    ws.Columns(1).NumberFormat = "0"
    
    ' Generate row's data
    For i = 2 To svcList_Length + 1 ' Start from row 2
        ' Get the cell mapped to each column
        Set indexCell = ws.Cells(i, 1)
        Set svcCell = ws.Cells(i, 2)
        Set priceCell = ws.Cells(i, 3)
        Set picCell = ws.Cells(i, 4)
        
        ' Fill with values with predefined conditions
        indexCell.Value = i - 1
        svcCell.Value = svcList(i - 2)
        priceCell.Value = ""
        picCell.Value = ""
        
        ' Create Data Validation for choosing "Person In Charge"
        Set refWs = ThisWorkbook.Sheets(userWsName)
        With ws.Range(picCell.Address).Validation 
            .Delete ' Remove any existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='" & refWs.Name & "'!B2:B3"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Next i

    ' Create border for the table
    BlackOutlineCells("D" & svcList_Length + 1)

    CreateBillingServiceWorksheet = ws.Name
End Function



