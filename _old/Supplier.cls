VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clSupplierID As Long
Public clName As Long
Public clFantasyName As Long
Public clCNPJ As Long
Public clCity As Long
Public clUF As Long
Public clContact As Long

'// PROPRIDEDADES //'

Private pSupplierID As String
Private pName As String
Private pFantasyName As String
Private pCNPJ As LongLong
Private pCity As String
Private pUF As String
Private pContact As String

Public Property Get SupplierID() As String
   SupplierID = pSupplierID
End Property
Public Property Let SupplierID(Value As String)
   pSupplierID = Value
End Property

Public Property Get Name() As String
   Name = pName
End Property
Public Property Let Name(Value As String)
   pName = Value
End Property

Public Property Get FantasyName() As String
   FantasyName = pFantasyName
End Property
Public Property Let FantasyName(Value As String)
   pFantasyName = Value
End Property

Public Property Get CNPJ() As LongLong
   CNPJ = pCNPJ
End Property
Public Property Let CNPJ(Value As LongLong)
   pCNPJ = Value
End Property

Public Property Get City() As String
   City = pCity
End Property
Public Property Let City(Value As String)
   pCity = Value
End Property

Public Property Get UF() As String
   UF = pUF
End Property
Public Property Let UF(Value As String)
   pUF = Value
End Property

Public Property Get Contact() As String
   Contact = pContact
End Property
Public Property Let Contact(Value As String)
   pContact = Value
End Property

Private Sub Class_Initialize()

    InstanceTables
    
    LoadClassColumnsIndexes TargetObject:=Me, _
                            ColumnsNameList:=GetListOfColumnIndexes, _
                            DatabaseTable:=tbSupplier, _
                            TechNameTable:=M_BancoDeDados.TechNameSupplier
    
End Sub

Private Function GetListOfColumnIndexes()
    
    GetListOfColumnIndexes = Array("clSupplierID", "clName", "clFantasyName", "clCNPJ", "clCity", "clUF", "clContact")

End Function

Public Sub LoadFromDatabaseToObject(ByVal ID As String)
    
    Dim RecordExists As Boolean: RecordExists = False
    
    Dim RecordRow As Long
    
    RecordRow = SearchRowID(ID, tbSupplier, clSupplierID)
    
    If RecordRow <> -1 Then RecordExists = True
        
    If RecordExists = False Then
       Exit Sub
    End If
    
    With tbSupplier
        SupplierID = .DataBodyRange(RecordRow, clSupplierID)
        Name = .DataBodyRange(RecordRow, clName)
        FantasyName = .DataBodyRange(RecordRow, clFantasyName)
        CNPJ = .DataBodyRange(RecordRow, clCNPJ)
        City = .DataBodyRange(RecordRow, clCity)
        UF = .DataBodyRange(RecordRow, clUF)
        Contact = .DataBodyRange(RecordRow, clContact)
    End With
    
End Sub

Public Sub CreateFromObjectToDatabase()
    
    Dim RecordExists As Boolean: RecordExists = False
    
    Dim RecordRow As Long
    
    RecordRow = SearchRowID(SupplierID, tbSupplier, clSupplierID)
    
    If RecordRow <> -1 Then RecordExists = True
        
    If RecordExists = True Then
        MsgBox "Fornecedor já cadastrado."
        Exit Sub
    End If
    
    tbSupplier.ListRows.Add
    
    With tbSupplier
        RecordRow = .DataBodyRange.Rows.count
        
        .DataBodyRange(RecordRow, clSupplierID) = SupplierID
        .DataBodyRange(RecordRow, clName) = Name
        .DataBodyRange(RecordRow, clFantasyName) = FantasyName
        .DataBodyRange(RecordRow, clCNPJ) = CNPJ
        .DataBodyRange(RecordRow, clCity) = City
        .DataBodyRange(RecordRow, clUF) = UF
        .DataBodyRange(RecordRow, clContact) = Contact
        
    End With
    
End Sub

Public Sub UpdateFieldInDatabase(ColumnToModify As Long, NewValue As Variant)

    Dim RecordRow As Long
    RecordRow = SearchRowID(SupplierID, tbSupplier, clSupplierID)
    
    UpdateTableField Table:=tbSupplier, RowID:=RecordRow, ColumnID:=ColumnToModify, Value:=NewValue

End Sub

