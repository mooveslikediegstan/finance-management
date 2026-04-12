VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SupplierRegistry 
   ClientHeight    =   3990
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   5175
   OleObjectBlob   =   "SupplierRegistry.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "SupplierRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mySupplier As Supplier
Public Cancelled As Boolean






Private Sub RegisterNewSupplier_Click()

   Set mySupplier = InstaciateSupplierFromForm
   
   mySupplier.CreateFromObjectToDatabase
   
   MsgBox "Fornecedor cadastrado com sucesso !!", vbOKOnly
   
   CleanForm
   
   Me.Hide
    
End Sub

Private Sub ModifySupplier_Click()

    Dim SupplierFromDB As Supplier
    Set SupplierFromDB = New Supplier
    SupplierFromDB.LoadFromDatabaseToObject ID:=Me.txtSupplierID

    Dim SupplierFromForm As Supplier
    Set SupplierFromForm = InstaciateSupplierFromForm
    
    Dim ModificationCount As Long
    ModificationCount = 0
    
    If SupplierFromForm.Name <> SupplierFromDB.Name Then
        SupplierFromDB.UpdateFieldInDatabase ColumnToModify:=SupplierFromDB.clName, NewValue:=SupplierFromForm.Name
        ModificationCount = ModificationCount + 1
    End If
    If SupplierFromForm.FantasyName <> SupplierFromDB.FantasyName Then
        SupplierFromDB.UpdateFieldInDatabase ColumnToModify:=SupplierFromDB.clFantasyName, NewValue:=SupplierFromForm.FantasyName
        ModificationCount = ModificationCount + 1
    End If
    If SupplierFromForm.CNPJ <> SupplierFromDB.CNPJ Then
        SupplierFromDB.UpdateFieldInDatabase ColumnToModify:=SupplierFromDB.clCNPJ, NewValue:=SupplierFromForm.CNPJ
        ModificationCount = ModificationCount + 1
    End If
    If SupplierFromForm.City <> SupplierFromDB.City Then
        SupplierFromDB.UpdateFieldInDatabase ColumnToModify:=SupplierFromDB.clCity, NewValue:=SupplierFromForm.City
        ModificationCount = ModificationCount + 1
    End If
    If SupplierFromForm.UF <> SupplierFromDB.UF Then
        SupplierFromDB.UpdateFieldInDatabase ColumnToModify:=SupplierFromDB.clUF, NewValue:=SupplierFromForm.UF
        ModificationCount = ModificationCount + 1
    End If
    If SupplierFromForm.Contact <> SupplierFromDB.Contact Then
        SupplierFromDB.UpdateFieldInDatabase ColumnToModify:=SupplierFromDB.clContact, NewValue:=SupplierFromForm.Contact
        ModificationCount = ModificationCount + 1
    End If
    
    
    If ModificationCount > 0 Then
        MsgBox "Fornecedor alterado com sucesso !!"
        CleanForm
        Me.Hide
    End If
    
    
End Sub

Private Sub CleanForm()

    Me.txtSupplierID = vbNullString
    Me.txtName = vbNullString
    Me.txtFantasyName = vbNullString
    Me.txtCNPJ = vbNullString
    Me.txtCity = vbNullString
    Me.cbxUF = vbNullString
    Me.txtContact = vbNullString

End Sub

Private Function InstaciateSupplierFromForm() As Supplier

    Dim tempSupplier As New Supplier
    
    tempSupplier.SupplierID = Me.txtSupplierID
    tempSupplier.Name = UCase(Me.txtName)
    tempSupplier.FantasyName = UCase(Me.txtFantasyName)
    tempSupplier.CNPJ = NormalizeCNPJNumber(Me.txtCNPJ)
    tempSupplier.City = UCase(Me.txtCity)
    tempSupplier.UF = UCase(Me.cbxUF)
    tempSupplier.Contact = Me.txtContact
    
    Set InstaciateSupplierFromForm = tempSupplier
    
    Set tempSupplier = Nothing

End Function

Private Sub btnSearchSupplier_Click()
    
    FillSupplierInfoToForm SelectedSupplier:=getSupplier
    
End Sub

Private Sub FillSupplierInfoToForm(ByVal SelectedSupplier As Supplier)
    
    If SelectedSupplier Is Nothing Then Exit Sub
    
    Me.txtSupplierID = SelectedSupplier.SupplierID
    Me.txtName = SelectedSupplier.Name
    Me.txtFantasyName = SelectedSupplier.FantasyName
    Me.txtCNPJ = CStr(SelectedSupplier.CNPJ)
    Me.txtCity = SelectedSupplier.City
    Me.cbxUF = SelectedSupplier.UF
    Me.txtContact = SelectedSupplier.Contact

End Sub

Private Function NormalizeCNPJNumber(ByVal TextCNPJ As String) As LongLong

    TextCNPJ = Replace(Trim(TextCNPJ), ".", "")
    TextCNPJ = Replace(TextCNPJ, "/", "")
    TextCNPJ = Replace(TextCNPJ, "-", "")
    
    NormalizeCNPJNumber = CLngLng(TextCNPJ)
    
End Function

Private Sub txtCNPJ_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    fApenasNumerosKeyPress KeyAscii
End Sub


Private Sub CancelRegistry_Click()
    CleanForm
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Cancelled = True
        Me.Hide
    End If
End Sub
