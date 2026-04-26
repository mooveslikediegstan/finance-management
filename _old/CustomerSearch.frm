VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerSearch 
   ClientHeight    =   5680
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   9450.001
   OleObjectBlob   =   "CustomerSearch.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CustomerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedCustomer As Customer
Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    UpdateListBox "", ""
End Sub

Sub UpdateListBox(CriteriaText As String, CriteriaCode As String)

    Dim CustomerList As Variant: CustomerList = getCustomerListForDisplay("CustomerTable", CriteriaText, CriteriaCode)

    With Me.CustomerSearchListBox
        
        .ColumnCount = GetArrayColumnSize(CustomerList)
        .Clear
        .AddItem
        .List = CustomerList
        .ColumnWidths = "60;200;200;150;20"
        .ColumnHeads = False
        .TextAlign = fmTextAlignCenter
        
    End With
    
End Sub

Private Sub NameSearchCriteria_Change()
    UpdateListBox Me.NameSearchCriteria.Text, Me.CodeSearchCriteria.Text
End Sub
Private Sub CodeSearchCriteria_Change()
    UpdateListBox Me.NameSearchCriteria.Text, Me.CodeSearchCriteria.Text
End Sub

Private Sub ConfirmCustomer_Click()

    If Me.CustomerSearchListBox.ListIndex <= 0 Then Exit Sub
    
    Set SelectedCustomer = New Customer
    
    SelectedCustomer.LoadFromDatabaseToObject ID:=Me.CustomerSearchListBox.List(Me.CustomerSearchListBox.ListIndex, 0)
    
    Cancelled = False
    
    Me.Hide
    
End Sub


Private Sub CustomerSearchListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    ConfirmCustomer_Click

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True ' Impede o "X" de descarregar o form
        Cancelled = True
        Me.Hide
    End If
End Sub
