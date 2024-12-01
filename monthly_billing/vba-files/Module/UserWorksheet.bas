Attribute VB_Name = "UserWorksheet"
Sub CreateUserWorksheet()
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
    ws.Name = "BillingData" ' Name the new sheet (you can change this name)
    
    ' Add headers with bold text
    ws.Cells(1, 1).Value = "No. #"
    ws.Cells(1, 2).Value = "Billing User"
    ws.Cells(1, 3).Value = "is Activated ?"
    ws.Rows(1).Font.Bold = True ' Make the headers bold
    
    ' Set the column widths for better visibility
    ws.Columns(1).ColumnWidth = 10
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 15 ' Width for the checkboxes column
    
    ' Optional: Set the number format for the index column
    ws.Columns(1).NumberFormat = "0"
    
    ' Determine the number of rows you want to fill
    lastRow = 20 ' Change 20 to however many rows you need
    
    ' Populate the index (No. #) in the first column with auto-incrementation
    For i = 2 To userList_Length + 1 ' Start from row 2
        ' Get the cell in the third column for the checkbox
        Set checkboxCell = ws.Cells(i, 3)
        
        ws.Cells(i, 1).Value = i - 1 ' Index starts from 1 and increments automatically
        ws.Cells(i, 2).Value = userList(i - 2)  ' Set the "Billing User" column as blank
        
        ' Add a checkbox in the third column (is Activated?)
        Set checkbox = ws.CheckBoxes.Add(checkboxCell.Left, checkboxCell.Top, checkboxCell.Width, checkboxCell.Height)
        
        ' ' Center the checkbox within the cell
        ' checkbox.Left = checkboxCell.Left + (checkboxCell.Width - checkbox.Width) / 2
        ' checkbox.Top = checkboxCell.Top + (checkboxCell.Height - checkbox.Height) / 2
        checkbox.Top = checkboxCell.Top + (checkboxCell.Height / 2#) - (checkbox.Height / 2#)
        ' checkbox.Left = checkboxCell.Left + (checkboxCell.Width / 2#) - (checkbox.Width / 2#)
        
        checkbox.Name = "Checkbox_" & i ' Give each checkbox a unique name
        checkbox.Caption = ""  ' Remove the caption for a clean checkbox
        checkbox.Value = xlOn ' Initially set the checkbox to unchecked
        checkbox.LinkedCell = ws.Cells(i, 3).Address
        ws.Cells(i, 3).NumberFormat = ";;;"
    Next i
End Sub

