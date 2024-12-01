Attribute VB_Name = "Utilities"
Function GetCurrentMonthAndYear()
    Dim currentMonth As Integer
    Dim currentYear As Integer
    
    currentMonth = Month(Date)
    currentYear = Year(Date)
    GetCurrentMonthAndYear = "(" & currentMonth & "_" & currentYear & ")"
End Function

Function BlackOutlineCells(RightBottomIndex As String)
    Dim rng As Range

    Set rng = Range("A1:" & RightBottomIndex)

    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Function
