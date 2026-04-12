VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SupplierSearch 
   ClientHeight    =   5660
   ClientLeft      =   345
   ClientTop       =   1635
   ClientWidth     =   9405.001
   OleObjectBlob   =   "SupplierSearch.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "SupplierSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedSupplier As Supplier
Public Cancelled As Boolean


Private Sub UserForm_Initialize()
    Me.UpdateListBox "", ""
End Sub

Sub UpdateListBox(CriteriaText As String, CriteriaCode As String)

    Dim SupplierList As Variant: SupplierList = getSupplierListForDisplay("SupplierTable", CriteriaText, CriteriaCode)

    With Me.SupplierSearchListBox
        
        .ColumnCount = GetArrayColumnSize(SupplierList)
        .Clear
        .AddItem
        .List = SupplierList
        .ColumnWidths = "60;150;150;100;100;25;100"
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

Private Sub ConfirmSupplier_Click()

    If Me.SupplierSearchListBox.ListIndex <= 0 Then Exit Sub
    
    Set SelectedSupplier = New Supplier
    
    SelectedSupplier.LoadFromDatabaseToObject ID:=Me.SupplierSearchListBox.List(Me.SupplierSearchListBox.ListIndex, 0)
    
    Cancelled = False
    
    Me.Hide
    
End Sub


Private Sub SupplierSearchListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    ConfirmSupplier_Click

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True ' Impede o "X" de descarregar o form
        Cancelled = True
        Me.Hide
    End If
End Sub
