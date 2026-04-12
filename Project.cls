VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// INDEX DAS COLUNAS DA TABELA //'
Public clProjectID As Long
Public clCustomerID As Long

'// PROPRIDEDADES //'
Private pProjectID As String
Private pCustomerID As String


'// GETTERS AND SETTERS ///////////////////
Public Property Get ProjectID() As String
   ProjectID = pProjectID
End Property
Public Property Let ProjectID(Value As String)
   pProjectID = Value
End Property

Public Property Get CustomerID() As String
   CustomerID = pCustomerID
End Property
Public Property Let CustomerID(Value As String)
   pCustomerID = Value
End Property

Private Sub Class_Initialize()

    InstanceTables
    
    LoadClassColumnsIndexes TargetObject:=Me, _
                            ColumnsNameList:=GetListOfColumnIndexes, _
                            DatabaseTable:=tbProject, _
                            TechNameTable:=M_BancoDeDados.TechNameProject
      
End Sub

Private Function GetListOfColumnIndexes()
    
    GetListOfColumnIndexes = Array("clProjectID", "clCustomerID")

End Function

Public Sub LoadFromDatabaseToObject(ByVal ID As String)

    
End Sub


