VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthlyBilling 
   Caption         =   "Monthly Billing"
   ClientHeight    =   4070
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7220
   OleObjectBlob   =   "MonthlyBilling.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MonthlyBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    CreateBillingServiceWorksheet(CreateUserWorksheet)
    CreateSummaryWorksheet
End Sub

Private Sub createBilling_btn_Click()
    CreateBillingServiceWorksheet
End Sub

Private Sub createServiceSummary_btn_Click()
    CreateSummaryWorksheet
End Sub

Private Sub createUser_btn_Click()
    CreateUserWorksheet
End Sub
