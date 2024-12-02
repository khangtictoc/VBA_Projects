Attribute VB_Name = "SummaryWorksheet"
Function CreateSummaryWorksheet(svWsName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim checkbox As Object
    Dim userList As Variant
    userList = Array("Tran Hoang Khang", _
    "Nguyen Vo Bao Huy")
    
    userList_Length = UBound(userList) - LBound(userList) + 1
    
    ' Create a new worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Summary" & GetCurrentMonthAndYear()
    
    ' Add headers with bold text
    ws.Cells(1, 1).Value = "No. #"
    ws.Cells(1, 2).Value = "Billing User"
    ws.Cells(1, 3).Value = "Have Paid"
    ws.Cells(1, 4).Value = "Must Paid"
    ws.Cells(1, 5).Value = "Remaining"
    ws.Rows(1).Font.Bold = True ' Make the headers bold

    ' Align the data to the center horizontally and vertically
    ws.Columns(1).HorizontalAlignment = xlVAlignCenter
    ws.Columns(2).HorizontalAlignment = xlVAlignCenter
    ws.Columns(3).HorizontalAlignment = xlVAlignCenter
    ws.Columns(4).HorizontalAlignment = xlVAlignCenter
    ws.Columns(5).HorizontalAlignment = xlVAlignCenter
    ws.Columns(1).VerticalAlignment = xlVAlignCenter
    ws.Columns(2).VerticalAlignment = xlVAlignCenter
    ws.Columns(3).VerticalAlignment = xlVAlignCenter
    ws.Columns(4).VerticalAlignment = xlVAlignCenter
    ws.Columns(5).VerticalAlignment = xlVAlignCenter
    
    ' Set the column widths for better visibility
    ws.Columns(1).ColumnWidth = 10
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 15
    ws.Columns(4).ColumnWidth = 15
    ws.Columns(5).ColumnWidth = 15
    
    ' Optional: Set the number format for the index column
    ws.Columns(1).NumberFormat = "0"
    
    ' Generate row's data
    For i = 2 To userList_Length + 1 ' Start from row 2
        ' Get the cell mapped to each column
        Set indexCell = ws.Cells(i, 1)
        Set userCell = ws.Cells(i, 2)
        Set havePaidCell = ws.Cells(i, 3)
        Set mustPaidCell = ws.Cells(i, 4)
        Set remainingCell = ws.Cells(i, 5)
        
        ' Fill with values with predefined conditions
        indexCell.Value = i - 1
        userCell.Value = userList(i - 2)
        havePaidCell.Value = ""
        mustPaidCell.Value = ""
        remainingCell.Value = ""

        havePaidCell.Value = "=SUMIF('" & svWsName & "'!$D$2:$D$8, B" & i & ",'" &  svWsName & "'!$C$2:$C$8)"
        mustPaidCell.Value = "=SUM('" & svWsName & "'!$C$2:$C$8) /2"
        remainingCell.Value = "=C" & i & "-D" & i
    Next i

    ' Create border for the table
    BlackOutlineCells("E" & userList_Length + 1)
End Function

