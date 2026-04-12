VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clCustomerID As Long
Public clCustomerName As Long
Public clCustomerShortName As Long
Public clCity As Long
Public clUF As Long

'// PROPRIDEDADES //'

Private pCustomerID As String
Private pCustomerName As String
Private pCustomerShortName As String
Private pCity As String
Private pUF As String
Private pProjects As Collection

Public Property Get CustomerID() As String
   CustomerID = pCustomerID
End Property
Public Property Let CustomerID(Value As String)
   pCustomerID = Value
End Property

Public Property Get CustomerName() As String
   CustomerName = pCustomerName
End Property
Public Property Let CustomerName(Value As String)
   pCustomerName = Value
End Property

Public Property Get CustomerShortName() As String
   CustomerShortName = pCustomerShortName
End Property
Public Property Let CustomerShortName(Value As String)
   pCustomerShortName = Value
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

Public Property Get Projects() As Collection
   Set Projects = pProjects
End Property
Public Property Set Projects(Coll As Collection)
   Set pProjects = Coll
End Property

Private Sub Class_Initialize()

    InstanceTables
    
    LoadClassColumnsIndexes TargetObject:=Me, _
                            ColumnsNameList:=GetListOfColumnIndexes, _
                            DatabaseTable:=tbCustomer, _
                            TechNameTable:=M_BancoDeDados.TechNameCustomer
    
    Set pProjects = New Collection
    
End Sub

Private Function GetListOfColumnIndexes()
    
    GetListOfColumnIndexes = Array("clCustomerID", _
                                "clCustomerName", _
                                "clCustomerShortName", _
                                "clCity", _
                                "clUF")

End Function

Public Sub LoadFromDatabaseToObject(ByVal ID As String)
    
    Dim RecordExists As Boolean: RecordExists = False
    
    Dim RecordRow As Long
    
    RecordRow = SearchRowID(ID, tbCustomer, clCustomerID)
    
    If RecordRow <> -1 Then RecordExists = True
        
    If RecordExists = False Then
       Exit Sub
    End If
    
    With tbCustomer
        CustomerID = .DataBodyRange(RecordRow, clCustomerID)
        CustomerName = .DataBodyRange(RecordRow, clCustomerName)
        CustomerShortName = .DataBodyRange(RecordRow, clCustomerShortName)
        City = .DataBodyRange(RecordRow, clCity)
        UF = .DataBodyRange(RecordRow, clUF)
    End With
    
End Sub

Sub LoadFromFormToObject(ByVal QtdParcelas As Long) '/INCOMPLETE

    'NUMPR = UCase(Trim(formControle.txt_Cadastro_NUMPR))
    'DATAPR = formControle.txt_Cadastro_DATAPR
    'CODFORNEC = formControle.txt_Cadastro_CODFORNEC '// CODIGO DO FORNECEDOR
    'ALOCREC = formControle.cbx_Cadastro_ALOCREC
    'REQ = formControle.txt_Cadastro_REQARIBA

    
End Sub


Private Sub Class_Terminate()

    Set pProjects = Nothing
    
End Sub

