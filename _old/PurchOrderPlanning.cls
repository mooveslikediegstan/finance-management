VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PurchOrderPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// INDEX DAS COLUNAS DA TABELA //'
Public clPurchaseRequisitionID As Long
Public clPlannedDate As Long
Public clAmount As Long

'// PROPRIDEDADES //'
Private pPurchaseRequisitionID As String
Private pPlannedDate As Date
Private pAmount As Currency


'// GETTERS AND SETTERS ///////////////////
Public Property Get PurchaseRequisitionID() As String
   PurchaseRequisitionID = pPurchaseRequisitionID
End Property
Public Property Let PurchaseRequisitionID(Value As String)
   pPurchaseRequisitionID = Value
End Property

Public Property Get PlannedDate() As Date
   PlannedDate = pPlannedDate
End Property
Public Property Let PlannedDate(Value As Date)
   pPlannedDate = Value
End Property

Public Property Get Amount() As Currency
   Amount = pAmount
End Property
Public Property Let Amount(Value As Currency)
   pAmount = Value
End Property

Private Sub Class_Initialize()

    InstanceTables
    
    LoadClassColumnsIndexes TargetObject:=Me, _
                            ColumnsNameList:=GetListOfColumnIndexes, _
                            DatabaseTable:=tbPurchaseOrderPlanning, _
                            TechNameTable:=M_BancoDeDados.TechNamePurchaseOrderPlanning
      
End Sub

Private Function GetListOfColumnIndexes()
    
    GetListOfColumnIndexes = Array("clPurchaseRequisitionID", "clPlannedDate", "clAmount")

End Function

