VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// INDEX DAS COLUNAS DA TABELA //'

Public clPurchaseRequisitionID As Long '// NUMERO DA SOLICITAÇÃO REFERENTE AO PEDIDO
Public clPurchaseOrderID As Long '// NÚMERO DA PO
Public clInvoiceNumber As Long '// NUMERO DA NOTA FISCAL
Public clPostingDate As Long '// DATA DA POSTAGEM DA NOTA (COMPETÊNCIA DA NF)
Public clPaymentDate As Long '// DATA PREVISTA PARA PAGAMENTO
Public clTotalAmount As Long '// VALOR DA NOTA FISCAL
Public clBMNumber As Long '// NUMERO DO BOLETIM DE MEDIÇÃO DAQUELE PROJETO


'// PROPRIDEDADES //' (VERIFICAR DESCRIÇÕES ACIMA)
Public pPurchaseRequisitionID As String
Public pPurchaseOrderID As String
Public pInvoiceNumber As String
Public pPostingDate As Date
Public pPaymentDate As Date
Public pTotalAmount As Currency
Public pBMNumber As String

'// GETTERS AND SETTERS ///////////////////

Public Property Get PurchaseRequisitionID() As String
   PurchaseRequisitionID = pPurchaseRequisitionID
End Property
Public Property Let PurchaseRequisitionID(Value As String)
   pPurchaseRequisitionID = Value
End Property

Public Property Get PurchaseOrderID() As String
   PurchaseOrderID = pPurchaseOrderID
End Property
Public Property Let PurchaseOrderID(Value As String)
   pPurchaseOrderID = Value
End Property

Public Property Get InvoiceNumber() As String
   InvoiceNumber = pInvoiceNumber
End Property
Public Property Let InvoiceNumber(Value As String)
   pInvoiceNumber = Value
End Property

Public Property Get PostingDate() As Date
   PostingDate = pPostingDate
End Property
Public Property Let PostingDate(Value As Date)
   pPostingDate = Value
End Property

Public Property Get PaymentDate() As Date
   PaymentDate = pPaymentDate
End Property
Public Property Let PaymentDate(Value As Date)
   pPaymentDate = Value
End Property

Public Property Get TotalAmount() As Currency
   TotalAmount = pTotalAmount
End Property
Public Property Let TotalAmount(Value As Currency)
   pTotalAmount = Value
End Property

Public Property Get BMNumber() As String
   BMNumber = pBMNumber
End Property
Public Property Let BMNumber(Value As String)
   pBMNumber = Value
End Property

Private Sub Class_Initialize()

    InstanceTables
    
    LoadClassColumnsIndexes TargetObject:=Me, _
                            ColumnsNameList:=GetListOfColumnIndexes, _
                            DatabaseTable:=tbInvoice, _
                            TechNameTable:=M_BancoDeDados.TechNameInvoice
      
End Sub

Private Function GetListOfColumnIndexes()
    
    GetListOfColumnIndexes = Array("clPurchaseRequisitionID", "clPurchaseOrderID", "clInvoiceNumber", "clPostingDate", "clPaymentDate", "clTotalAmount", "clBMNumber")

End Function

Public Sub LoadFromDatabaseToObject(ByVal ID As String)
    
    'Dim RecordExists As Boolean: RecordExists = False
    '
    'Dim RecordRow As Long
    '
    'RecordRow = SearchRowID(ID, tbCustomer, clCustomerID)
    '
    'If RecordRow <> -1 Then RecordExists = True
    '
    'If RecordExists = False Then
    '   Exit Sub
    'End If
    '
    'With tbCustomer
    '    CustomerID = .DataBodyRange(RecordRow, clCustomerID)
    '    CustomerName = .DataBodyRange(RecordRow, clCustomerName)
    '    CustomerShortName = .DataBodyRange(RecordRow, clCustomerShortName)
    '    City = .DataBodyRange(RecordRow, clCity)
    '    UF = .DataBodyRange(RecordRow, clUF)
    'End With
    
End Sub


