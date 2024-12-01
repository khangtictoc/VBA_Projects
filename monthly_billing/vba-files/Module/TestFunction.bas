Attribute VB_Name = "TestFunction"

Public Sub TestFunction()
    Dim ws As Worksheet
    Set refWs = ThisWorkbook.Sheets("Billing User(12_2024)") ' Adjust the sheet name as needed
    Set targetWs = ThisWorkbook.Sheets("Billing Service(12_2024)") 

    With targetWs.Range("D2").Validation ' Adjust the cell reference as needed
        .Delete ' Remove any existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='" & refWs.Name & "'!B2:B3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub

