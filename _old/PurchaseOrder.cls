VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// INDEX DAS COLUNAS DA TABELA //'
Public clNUMPR As Long '// NÚMERO DA PR / CONCUR / DEBIT NOTES
Public clDATAPR As Long '// DATA DE INPUT NA PLANILHA
Public clDATAPLAN As Long
Public clREQ As Long '// REQUISITANTE
Public clCODFORNEC As Long 'ID CHAMADO
Public clPREVFC As Long '// PREVISTO EM FORECAST OU NÃO
Public clDESCRIC As Long '// DESCRIÇÃO DO SERVIÇO
Public clPROJECT As Long '// PROJETO
Public clCLIENTE '// CLIENTE
Public clTOTALBRL As Long '// VALOR EM REAIS
Public clSALDOBRL As Long '// SALDO DA SC
Public clALOCREC As Long '// ALOCAÇÃO DE RECURSO (GRUPO DE CUSTO)
Public clFORNECEDOR As Long
Public clCNPJFORN As Long
Public clNUMPO As Long '// NÚMERO DA PURCHASE ORDER
Public clDATAPO As Long '// DATA DE RECEBIMENTO DA PURCHASE ORDER
Public clNUMNF As Long '// NÚMERO DA NOTA FISCAL
Public clENTREGANF As Long '// ENTREGA NF

Public plRegistro As Long

'// PROPRIDEDADES //'

Private pNUMPR As String '// NÚMERO DA PR / CONCUR / DEBIT NOTES
Private pDATAPR As Date '// DATA DE INPUT NA PLANILHA
Private pDATAPLAN As Date '// DATA PLANEJADA PARA AQUELE SERVIÇO SER EXECUTADO '// ADICIONADO EM 05/DEZ/2024 CONFORMA ALINHAMENTO COM TIME DE ASSISTÊNCIA TECNICA
Private pREQ As String '// REQUISITANTE
Private pCODFORNEC As String '// MODIFICADO PARA O NÚMERO DE ID DO CHAMADO DA ASSISTENCIA
Private pPREVFC As String '// PREVISTO EM FORECAST OU NÃO
Private pDESCRIC As String '// DESCRIÇÃO DO SERVIÇO
Private pPROJECT As String
Private pCLIENTE
Private pTOTALBRL As Currency '// VALOR EM REAIS
Private pSALDOBRL As Currency '// SALDO DA SC
Private pALOCREC As String '// ALOCAÇÃO DE RECURSO (GRUPO DE CUSTO)
Private pFORNECEDOR As String
Private pCNPJFORN As String '// NÚMERO DA CONTA
Private pNUMPO As String '// NÚMERO DA PURCHASE ORDER
Private pDATAPO As Date '// DATA DE RECEBIMENTO DA PURCHASE ORDER
Private pNUMNF As String '// NUMERO DA NF
Private pENTREGANF As String '// ENTREGA NF

Private pPlanningDates As Collection






Public Property Get NUMPR() As String
   NUMPR = pNUMPR
End Property
Public Property Let NUMPR(Value As String)
   pNUMPR = Value
End Property

Public Property Get DATAPR() As Date
   DATAPR = pDATAPR
End Property
Public Property Let DATAPR(Value As Date)
   pDATAPR = Value
End Property

Public Property Get DATAPLAN() As Date
   DATAPLAN = pDATAPLAN
End Property
Public Property Let DATAPLAN(Value As Date)
   pDATAPLAN = Value
End Property

Public Property Get REQ() As String
   REQ = pREQ
End Property
Public Property Let REQ(Value As String)
   pREQ = Value
End Property

Public Property Get CODFORNEC() As String
   CODFORNEC = pCODFORNEC
End Property
Public Property Let CODFORNEC(Value As String)
   pCODFORNEC = Value
End Property

Public Property Get PREVFC() As String
   PREVFC = pPREVFC
End Property
Public Property Let PREVFC(Value As String)
   pPREVFC = Value
End Property

Public Property Get DESCRIC() As String
   DESCRIC = pDESCRIC
End Property
Public Property Let DESCRIC(Value As String)
   pDESCRIC = Value
End Property

Public Property Get Project() As String
   Project = pPROJECT
End Property
Public Property Let Project(Value As String)
   pPROJECT = Value
End Property

Public Property Get Cliente() As String
   Cliente = pCLIENTE
End Property
Public Property Let Cliente(Value As String)
   pCLIENTE = Value
End Property

Public Property Get TOTALBRL() As Currency
   TOTALBRL = pTOTALBRL
End Property
Public Property Let TOTALBRL(Value As Currency)
   pTOTALBRL = Value
End Property

Public Property Get SALDOBRL() As Currency
   SALDOBRL = pSALDOBRL
End Property
Public Property Let SALDOBRL(Value As Currency)
   pSALDOBRL = Value
End Property

Public Property Get ALOCREC() As String
   ALOCREC = pALOCREC
End Property
Public Property Let ALOCREC(Value As String)
   pALOCREC = Value
End Property

Public Property Get Fornecedor() As String
   Fornecedor = pFORNECEDOR
End Property
Public Property Let Fornecedor(Value As String)
   pFORNECEDOR = Value
End Property

Public Property Get CNPJFORN() As String
   CNPJFORN = pCNPJFORN
End Property
Public Property Let CNPJFORN(Value As String)
   pCNPJFORN = Value
End Property

Public Property Get NUMPO() As String
   NUMPO = pNUMPO
End Property
Public Property Let NUMPO(Value As String)
   pNUMPO = Value
End Property

Public Property Get DATAPO() As Date
   DATAPO = pDATAPO
End Property
Public Property Let DATAPO(Value As Date)
   pDATAPO = Value
End Property

Public Property Get NUMNF() As String
   NUMNF = pNUMNF
End Property
Public Property Let NUMNF(Value As String)
   pNUMNF = Value
End Property

Public Property Get ENTREGANF() As String
   ENTREGANF = pENTREGANF
End Property
Public Property Let ENTREGANF(Value As String)
   pENTREGANF = Value
End Property

Public Property Get lRegistro() As Long
   lRegistro = plRegistro
End Property
Public Property Let lRegistro(Value As Long)
   plRegistro = Value
End Property

Public Property Get PlanningDates() As Collection
   Set PlanningDates = pPlanningDates
End Property
Public Property Set PlanningDates(Coll As Collection)
   Set pPlanningDates = Coll
End Property


Private Sub Class_Initialize()

    InstanceTables
    
    LoadClassColumnsIndexes TargetObject:=Me, _
                            ColumnsNameList:=GetListOfColumnIndexes, _
                            DatabaseTable:=tbPurchaseOrder, _
                            TechNameTable:=M_BancoDeDados.TechNamePurchaseOrder
    
    Set pPlanningDates = New Collection
      
End Sub

Private Function GetListOfColumnIndexes()
    
    GetListOfColumnIndexes = Array("clNUMPR", "clDATAPR", _
                                    "clDATAPLAN", "clREQ", _
                                    "clCODFORNEC", "clPREVFC", _
                                    "clDESCRIC", "clPROJECT", _
                                    "clCLIENTE", "clTOTALBRL", _
                                    "clSALDOBRL", "clALOCREC", _
                                    "clFORNECEDOR", "clCNPJFORN", _
                                    "clNUMPO", "clDATAPO", _
                                    "clNUMNF", "clENTREGANF")

End Function

Public Sub LoadFromDatabaseToObject(ByVal ID As String) '//OK
    
    Dim ExisteRegistro As Boolean: ExisteRegistro = False
    
    lRegistro = SearchRowID(ID, M_BancoDeDados.tbPurchaseOrder, clNUMPR)
    
    If lRegistro <> -1 Then ExisteRegistro = True
    
    If ExisteRegistro = False Then
        lRegistro = SearchRowID(ID, M_BancoDeDados.tbPurchaseOrder, clNUMPO)
        If lRegistro <> -1 Then ExisteRegistro = True
    End If
    
    If ExisteRegistro = False Then
       Exit Sub
    End If
    
    With M_BancoDeDados.tbPurchaseOrder
    
        NUMPR = Trim(.DataBodyRange(lRegistro, clNUMPR))
        DATAPR = .DataBodyRange(lRegistro, clDATAPR)
        DATAPLAN = .DataBodyRange(lRegistro, clDATAPLAN)
        REQ = .DataBodyRange(lRegistro, clREQ)
        CODFORNEC = .DataBodyRange(lRegistro, clCODFORNEC) '// ID DO CHAMADO
        PREVFC = .DataBodyRange(lRegistro, clPREVFC)
        DESCRIC = .DataBodyRange(lRegistro, clDESCRIC)
        Project = .DataBodyRange(lRegistro, clPROJECT)
        Cliente = .DataBodyRange(lRegistro, clCLIENTE)
        TOTALBRL = .DataBodyRange(lRegistro, clTOTALBRL)
        SALDOBRL = .DataBodyRange(lRegistro, clSALDOBRL)
        ALOCREC = .DataBodyRange(lRegistro, clALOCREC)
        Fornecedor = .DataBodyRange(lRegistro, clFORNECEDOR)
        CNPJFORN = .DataBodyRange(lRegistro, clCNPJFORN)
        NUMPO = Trim(.DataBodyRange(lRegistro, clNUMPO))
        DATAPO = .DataBodyRange(lRegistro, clDATAPO)
        NUMNF = .DataBodyRange(lRegistro, clNUMNF)
        ENTREGANF = .DataBodyRange(lRegistro, clENTREGANF)
        
    End With
    
    LoadPlanningDates ID:=NUMPR
    
End Sub

Sub LoadPlanningDates(ByVal ID As String)

    Dim PlanningDateArray As Variant
    PlanningDateArray = tbPurchaseOrderPlanning.DataBodyRange
    
    Dim PlannedDate As PurchOrderPlanning
    
    Dim RecordRow As Long
    For RecordRow = LBound(PlanningDateArray) To UBound(PlanningDateArray)
        
        If PlanningDateArray(RecordRow, 1) = NUMPR Then
            
            Set PlannedDate = New PurchOrderPlanning
            
            PlannedDate.PurchaseRequisitionID = PlanningDateArray(RecordRow, PlannedDate.clPurchaseRequisitionID)
            PlannedDate.PlannedDate = PlanningDateArray(RecordRow, PlannedDate.clPlannedDate)
            PlannedDate.Amount = PlanningDateArray(RecordRow, PlannedDate.clAmount)
            
            PlanningDates.Add PlannedDate
            
            Set PlannedDate = Nothing
            
        End If
        
    Next RecordRow

End Sub

Sub LoadFromFormToObject(ByVal ParcelQuantity As Long)
    
    Dim PlannedDate As PurchOrderPlanning, Parcel As Long
    
    NUMPR = UCase(Trim(formControle.txt_Cadastro_NUMPR))
    DATAPR = formControle.txt_Cadastro_DATAPR
    
    If formControle.chbx_Cadastro_POGUARDACHUVA.Value = True And ParcelQuantity > 0 Then
    
        For Parcel = 1 To ParcelQuantity
        
            Set PlannedDate = New PurchOrderPlanning
            
            PlannedDate.PurchaseRequisitionID = NUMPR
            PlannedDate.PlannedDate = CDate(formControle.lbx_DATAPLANSC.List(Parcel, 0))
            PlannedDate.Amount = CCur(formControle.lbx_DATAPLANSC.List(Parcel, 1))
            
            PlanningDates.Add PlannedDate
            
            Set PlannedDate = Nothing
            
        Next Parcel
        
        DATAPLAN = CDate(formControle.lbx_DATAPLANSC.List(1, 0))
        
    Else:
        
        Set PlannedDate = New PurchOrderPlanning
            
        PlannedDate.PurchaseRequisitionID = NUMPR
        PlannedDate.PlannedDate = CDate(formControle.txt_Cadastro_DATAPLAN)
        PlannedDate.Amount = CCur(formControle.txt_Cadastro_TOTALBRL)
        
        PlanningDates.Add PlannedDate
        
        Set PlannedDate = Nothing
    
        DATAPLAN = formControle.txt_Cadastro_DATAPLAN '// DATA PLANEJADA
        
    End If
    
    CODFORNEC = formControle.txt_Cadastro_CODFORNEC '// CODIGO DO FORNECEDOR
    ALOCREC = formControle.cbx_Cadastro_ALOCREC
    REQ = formControle.txt_Cadastro_REQARIBA
    PREVFC = formControle.cbx_Cadastro_PREVFC
    DESCRIC = UCase(formControle.txt_Cadastro_DESCRIC)
    Project = formControle.cbx_Cadastro_PROJECT
    Cliente = formControle.cbx_Cadastro_CLIENTE
    TOTALBRL = formControle.txt_Cadastro_TOTALBRL
    ALOCREC = formControle.cbx_Cadastro_ALOCREC
    Fornecedor = UCase(formControle.cbx_Cadastro_FORNECEDOR)
    CNPJFORN = formControle.txt_Cadastro_CNPJFORN
    ENTREGANF = "PENDENTE"
    
End Sub

'/////////////////////////////
Public Sub CadastrarPR()

    Call S_Desproteger_Planilha
    
    Dim uLinha As Long, uLinhaPLANSC As Long, QtdParcelas As Long, ListIndexCounter As Long
    
    QtdParcelas = formControle.lbx_DATAPLANSC.ListCount '// o cabeçalho é o primeiro registro do listbox, por isso -1
    
    Me.LoadFromFormToObject (QtdParcelas)
    
    M_BancoDeDados.tbPurchaseOrder.ListRows.Add
    uLinha = M_BancoDeDados.tbPurchaseOrder.ListRows.count
    
    With M_BancoDeDados.tbPurchaseOrder
    
        .DataBodyRange(uLinha, clNUMPR) = NUMPR
        .DataBodyRange(uLinha, clDATAPR) = DATAPR
        .DataBodyRange(uLinha, clDATAPLAN) = DATAPLAN
        .DataBodyRange(uLinha, clCODFORNEC) = CODFORNEC
        .DataBodyRange(uLinha, clREQ) = REQ
        .DataBodyRange(uLinha, clPREVFC) = PREVFC
        .DataBodyRange(uLinha, clDESCRIC) = DESCRIC
        .DataBodyRange(uLinha, clPROJECT) = Project
        .DataBodyRange(uLinha, clCLIENTE) = Cliente
        .DataBodyRange(uLinha, clTOTALBRL) = CCur(TOTALBRL)
        .DataBodyRange(uLinha, clALOCREC) = ALOCREC
        .DataBodyRange(uLinha, clFORNECEDOR) = Fornecedor
        .DataBodyRange(uLinha, clCNPJFORN) = CNPJFORN
        .DataBodyRange(uLinha, clENTREGANF) = ENTREGANF
        
    End With
    
    If formControle.chbx_Cadastro_POGUARDACHUVA.Value = True And QtdParcelas > 0 Then
    
        M_BancoDeDados.tbPurchaseOrderPlanning.ListRows.Add
        uLinhaPLANSC = M_BancoDeDados.tbPurchaseOrderPlanning.ListRows.count
        
        For ListIndexCounter = 2 To QtdParcelas
            M_BancoDeDados.tbPurchaseOrderPlanning.DataBodyRange(uLinhaPLANSC, 1) = CStr(NUMPR) '// num da SC
            M_BancoDeDados.tbPurchaseOrderPlanning.DataBodyRange(uLinhaPLANSC, 2) = CDate(formControle.lbx_DATAPLANSC.List(ListIndexCounter - 1, 0)) '// data planejada
            M_BancoDeDados.tbPurchaseOrderPlanning.DataBodyRange(uLinhaPLANSC, 3) = CCur(formControle.lbx_DATAPLANSC.List(ListIndexCounter - 1, 1)) '// valor
            uLinhaPLANSC = uLinhaPLANSC + 1
        Next ListIndexCounter
    
    Else:
        
        M_BancoDeDados.tbPurchaseOrderPlanning.ListRows.Add
        uLinhaPLANSC = M_BancoDeDados.tbPurchaseOrderPlanning.ListRows.count
        M_BancoDeDados.tbPurchaseOrderPlanning.DataBodyRange(uLinhaPLANSC, 1) = CStr(NUMPR) '// Num da SC
        M_BancoDeDados.tbPurchaseOrderPlanning.DataBodyRange(uLinhaPLANSC, 2) = CDate(DATAPLAN) '//Data planejada
        M_BancoDeDados.tbPurchaseOrderPlanning.DataBodyRange(uLinhaPLANSC, 3) = CCur(TOTALBRL) '// valor
    
    End If
    
    Call Histórico_Modificações_Cadastro
   
    '// APAGANDO AS INFORMAÇÕES DO FORMULÁRIO
    formControle.txt_Cadastro_NUMPR = ""
    formControle.txt_Cadastro_DATAPR = ""
    formControle.txt_Cadastro_DATAPLAN = ""
    formControle.txt_Cadastro_CODFORNEC = ""
    formControle.cbx_Cadastro_ALOCREC = ""
    formControle.txt_Cadastro_REQARIBA = ""
    formControle.cbx_Cadastro_PREVFC = ""
    formControle.txt_Cadastro_DESCRIC = ""
    formControle.cbx_Cadastro_PROJECT = ""
    formControle.cbx_Cadastro_CLIENTE = ""
    formControle.txt_Cadastro_TOTALBRL = ""
    formControle.cbx_Cadastro_ALOCREC = ""
    formControle.cbx_Cadastro_FORNECEDOR = ""
    formControle.txt_Cadastro_CNPJFORN = ""
    formControle.lbx_DATAPLANSC.Clear
    formControle.chbx_Cadastro_POGUARDACHUVA.Value = False
    
    Call S_Proteger_Planilha

End Sub

Sub Atualizar_PO()

   If formControle.txt_attPO_NUMPO = vbNullString Then
      MsgBox "Insira o número da PO", vbCritical
      Exit Sub
   End If
   
   If Len(formControle.txt_attPO_DATAPO) <> 10 Then
      MsgBox "Insira corretamente a data", vbCritical
      Exit Sub
   End If
    
   NUMPR = UCase(Trim(formControle.txt_attPO_NUMPR))
   
   Me.LoadFromDatabaseToObject (NUMPR)
    If NUMPO <> vbNullString Then
        MsgBox "Pedido de compra já cadastrado para essa solicitação de compra.", vbCritical, "Cadastro inválido"
        Exit Sub
    End If
    
   NUMPO = UCase(Trim(formControle.txt_attPO_NUMPO))
   
   DATAPO = formControle.txt_attPO_DATAPO
   
   UpdateTableField Table:=M_BancoDeDados.tbPurchaseOrder, RowID:=lRegistro, ColumnID:=clNUMPO, Value:=CStr(NUMPO)
   
   UpdateTableField Table:=M_BancoDeDados.tbPurchaseOrder, RowID:=lRegistro, ColumnID:=clDATAPO, Value:=CDate(DATAPO)
   
   Me.Histórico_Modificações_AttPO
   
   MsgBox "Atualização da PO concluída.", vbOKOnly
   
   formControle.txt_attPO_NUMPR = vbNullString
   formControle.txt_attPO_NUMPO = vbNullString
   formControle.txt_attPO_DATAPO = vbNullString
   
   formControle.txt_Menu_NUMPR = vbNullString
   formControle.MultiPage1.Value = 0
   
   Call S_Proteger_Planilha
   
End Sub
'/////////////////////////////////////////////

Sub Atualizar_NF()

    If formControle.txt_ATTNF_NUMNF = "" Then
       MsgBox "Insira o número da NF", vbCritical
       Exit Sub
    End If
    
    If formControle.txt_ATTNF_TOTALBRL = 0 Or formControle.txt_ATTNF_TOTALBRL = "" Then
        MsgBox "Insira o valor da NF!!", vbCritical
        Exit Sub
    End If
    
    If formControle.cbx_ATTNF_ENTREGANF = "PENDENTE" Then
        MsgBox "O status da entrega de nota fiscal deverá ser 'PARCIAL' OU 'ENTREGUE'", vbCritical
        Exit Sub
    End If
    

    Dim NF As Invoice: Set NF = New Invoice
        
    Me.LoadFromDatabaseToObject (UCase(Trim(formControle.txt_attNF_NUMPR)))
    
     If SALDOBRL < CCur(formControle.txt_ATTNF_TOTALBRL) Then
         MsgBox "Valor da nota é maior do que o saldo remanescente na Ordem de Compra. Saldo restante: R$ " & SALDOBRL, vbCritical
         Exit Sub
     End If
     
    NF.InvoiceNumber = formControle.txt_ATTNF_NUMNF
    
    Call S_Desproteger_Planilha
    
    InstanceTables
       
    If M_BancoDeDados.tbPurchaseOrder.DataBodyRange(lRegistro, clNUMNF) = "" Then
        M_BancoDeDados.tbPurchaseOrder.DataBodyRange(lRegistro, clNUMNF) = CStr(NF.InvoiceNumber)
    Else:
        M_BancoDeDados.tbPurchaseOrder.DataBodyRange(lRegistro, clNUMNF) = M_BancoDeDados.tbPurchaseOrder.DataBodyRange(lRegistro, clNUMNF) & "/" & CStr(NF.InvoiceNumber)
    End If
    
    Dim lRow As Long
    
    If tbInvoice.DataBodyRange Is Nothing Then
    
        tbInvoice.ListRows.Add
        
    Else:
    
        tbInvoice.ListRows.Add
        
        NF.TotalAmount = CCur(formControle.txt_ATTNF_TOTALBRL)
        
        NF.PostingDate = CDate(formControle.txt_ATTNF_DATAPOSTNF)
        
        If formControle.txt_ATTNF_DATAPREVPGTONF <> "" Then NF.PaymentDate = CDate(formControle.txt_ATTNF_DATAPREVPGTONF)
        
        NF.BMNumber = formControle.txt_ATTNF_NUMBM
        
        lRow = tbInvoice.DataBodyRange.Rows.count
        With tbInvoice
            .DataBodyRange(lRow, NF.clPurchaseOrderID) = CStr(NUMPO)
            .DataBodyRange(lRow, NF.clPurchaseRequisitionID) = M_BancoDeDados.tbPurchaseOrder.DataBodyRange(lRegistro, clNUMPR)
            .DataBodyRange(lRow, NF.clInvoiceNumber) = CStr(NF.InvoiceNumber)
            .DataBodyRange(lRow, NF.clPostingDate) = NF.PostingDate
            .DataBodyRange(lRow, NF.clPaymentDate) = NF.PaymentDate
            .DataBodyRange(lRow, NF.clBMNumber) = NF.BMNumber
            .DataBodyRange(lRow, NF.clTotalAmount) = NF.TotalAmount
        End With
        M_BancoDeDados.tbPurchaseOrder.DataBodyRange(Me.lRegistro, clENTREGANF) = formControle.cbx_ATTNF_ENTREGANF
    
    End If
    
    Call Histórico_Modificações_AttNF
    
    MsgBox "Atualização da NF concluída.", vbOKOnly
    
    formControle.txt_attNF_NUMPR = ""

    formControle.txt_ATTNF_NUMNF = ""

    formControle.txt_ATTNF_TOTALBRL = ""
    
    formControle.txt_Menu_NUMPR = ""
    
    formControle.MultiPage1.Value = 0
    
    formControle.txt_ATTNF_DATAPOSTNF = ""
    
    formControle.txt_ATTNF_DATAPREVPGTONF = ""
    
    formControle.txt_ATTNF_NUMBM = ""
    
    Call S_Proteger_Planilha
   
End Sub
'/////////////////////////////////////


Sub Atualizar_Registro()


   If formControle.cbx_attGR_STATUS = "" Then
      MsgBox "Não é possível deixar o 'Status' da requisição em branco.", vbCritical
      Exit Sub
   End If
   
   Call S_Desproteger_Planilha
   
   Dim PR As PurchaseOrder
   Set PR = New PurchaseOrder
   NUMPR = formControle.txt_attGR_NUMPR
   
   LoadFromDatabaseToObject (NUMPR)
   
   NUMPR = formControle.txt_attGR_NUMPR
   ALOCREC = formControle.cbx_attGR_ALOCREC
   DATAPR = formControle.txt_attGR_DATAPR
   REQ = formControle.cbx_attGR_REQ
   DESCRIC = formControle.txt_attGR_DESCRIC
   Fornecedor = formControle.cbx_attGR_FORNECEDOR
   Project = formControle.txt_attGR_PROJECT
   Cliente = formControle.cbx_attGR_CLIENTE
   CNPJFORN = formControle.txt_attGR_CNPJFORN
   TOTALBRL = formControle.txt_attGR_TOTALBRL
   NUMPO = formControle.txt_attGR_NUMPO
   DATAPO = formControle.txt_attGR_DATAPO
   NUMNF = formControle.txt_attGR_NUMNF
   PREVFC = formControle.cbx_attGR_PREVFC
   CODFORNEC = formControle.txt_attGR_DATAPREVSC
   


   MsgBox "Atualização concluída.", vbOKOnly
   
   formControle.txt_attGR_NUMPR = ""
   formControle.cbx_attGR_ALOCREC = ""
   formControle.txt_attGR_DATAPR = ""
   formControle.cbx_attGR_REQ = ""
   formControle.cbx_attGR_REQARIBA = ""
   formControle.txt_attGR_DESCRIC = ""
   formControle.cbx_attGR_FORNECEDOR = ""
   formControle.txt_attGR_PROJECT = ""
   formControle.cbx_attGR_CLIENTE = ""
   formControle.txt_attGR_CNPJFORN = ""
   formControle.txt_attGR_TOTALBRL = ""
   formControle.txt_attGR_NUMPO = ""
   formControle.txt_attGR_DATAPO = ""
   formControle.txt_attGR_NUMNF = ""
   formControle.cbx_attGR_PREVFC = ""
   formControle.txt_attGR_DATAPREVSC = ""
   formControle.txt_Menu_NUMPR = ""
   formControle.MultiPage1.Value = 0
   
   Call S_Proteger_Planilha
   
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////




'// HISTÓRICOS DE MODIFICAÇÃO //

Sub Histórico_Modificações_Cadastro()


   Dim tbHistórico As ListObject
   Set tbHistórico = ThisWorkbook.Sheets("Histórico de Modificações").ListObjects("tbHistórico")
   tbHistórico.ListRows.Add
   
   Dim lRow As Long
   lRow = tbHistórico.DataBodyRange.Rows.count
   Dim Hoje As Date
   Hoje = Date

   With tbHistórico '// PREENCHE OS CAMPOS DA TABELA COM AS INFORMAÇÕES INPUTADAS
   
    .DataBodyRange(lRow, 1) = Application.UserName
    .DataBodyRange(lRow, 2) = NUMPR
    .DataBodyRange(lRow, 3) = "Cadastro"
    .DataBodyRange(lRow, 4) = Hoje
    .DataBodyRange(lRow, 5) = DATAPR
    .DataBodyRange(lRow, 6) = CODFORNEC
    .DataBodyRange(lRow, 7) = REQ
    .DataBodyRange(lRow, 8) = ""
    .DataBodyRange(lRow, 9) = DESCRIC
    .DataBodyRange(lRow, 10) = Project
    .DataBodyRange(lRow, 11) = Cliente
    .DataBodyRange(lRow, 12) = Fornecedor
    .DataBodyRange(lRow, 13) = CNPJFORN
    .DataBodyRange(lRow, 14) = TOTALBRL
    .DataBodyRange(lRow, 15) = ALOCREC
    .DataBodyRange(lRow, 16) = "n/a"
    .DataBodyRange(lRow, 17) = "n/a"
    .DataBodyRange(lRow, 18) = "n/a"
    .DataBodyRange(lRow, 19) = "n/a"
    .DataBodyRange(lRow, 20) = "n/a"
    .DataBodyRange(lRow, 21) = "n/a"
    .DataBodyRange(lRow, 22) = "Ativo"
    
   End With

End Sub

Sub Histórico_Modificações_AttPO()

   Dim lRow As Long
   
   Dim tbHistórico As ListObject
   Set tbHistórico = ThisWorkbook.Sheets("Histórico de Modificações").ListObjects("tbHistórico")
   
   tbHistórico.ListRows.Add
   lRow = tbHistórico.DataBodyRange.Rows.count
   
   With tbHistórico '// PREENCHE OS CAMPOS DA TABELA COM AS INFORMAÇÕES INPUTADAS
   
    .DataBodyRange(lRow, 1) = Application.UserName
    .DataBodyRange(lRow, 2) = NUMPR
    .DataBodyRange(lRow, 3) = "Atualização PO"
    .DataBodyRange(lRow, 4) = Date
    .DataBodyRange(lRow, 5) = "n/a"
    
    .DataBodyRange(lRow, 6) = "n/a" 'DATAPREVSC
    .DataBodyRange(lRow, 7) = "n/a" 'REQ
    .DataBodyRange(lRow, 8) = "n/a" 'REQARIBA
    .DataBodyRange(lRow, 9) = "n/a" 'DESCRIC
    .DataBodyRange(lRow, 10) = "n/a" 'PROJECT
    .DataBodyRange(lRow, 11) = "n/a" 'CLIENTE
    .DataBodyRange(lRow, 12) = "n/a" 'FORNECEDOR
    .DataBodyRange(lRow, 13) = "n/a" 'CNPJFORN
    .DataBodyRange(lRow, 14) = "n/a" 'TOTALBRL
    .DataBodyRange(lRow, 15) = "n/a" 'ALOCREC
    .DataBodyRange(lRow, 16) = NUMPO
    .DataBodyRange(lRow, 17) = DATAPO
    .DataBodyRange(lRow, 18) = "n/a"
    .DataBodyRange(lRow, 19) = "n/a"
    .DataBodyRange(lRow, 20) = "n/a"
    .DataBodyRange(lRow, 21) = "n/a"
    .DataBodyRange(lRow, 22) = "n/a"
    
   End With

End Sub

Sub Histórico_Modificações_AttNF()

   Dim lRow As Long
   
   Dim tbHistórico As ListObject
   Set tbHistórico = ThisWorkbook.Sheets("Histórico de Modificações").ListObjects("tbHistórico")
   
   tbHistórico.ListRows.Add
   lRow = tbHistórico.DataBodyRange.Rows.count

   With tbHistórico '// PREENCHE OS CAMPOS DA TABELA COM AS INFORMAÇÕES INPUTADAS
   
   
    .DataBodyRange(lRow, 1) = Application.UserName
    .DataBodyRange(lRow, 2) = NUMPR
    .DataBodyRange(lRow, 3) = "Atualização NF"
    .DataBodyRange(lRow, 4) = Date
    .DataBodyRange(lRow, 5) = "n/a"
    
    .DataBodyRange(lRow, 6) = "n/a" 'DATAPREVSC
    .DataBodyRange(lRow, 7) = "n/a" 'REQ
    .DataBodyRange(lRow, 8) = "n/a" 'REQARIBA
    .DataBodyRange(lRow, 9) = "n/a" 'DESCRIC
    .DataBodyRange(lRow, 10) = "n/a" 'PROJECT
    .DataBodyRange(lRow, 11) = "n/a" 'CLIENTE
    .DataBodyRange(lRow, 12) = "n/a" 'FORNECEDOR
    .DataBodyRange(lRow, 13) = "n/a" 'CNPJFORN
    .DataBodyRange(lRow, 14) = "n/a" 'TOTALBRL
    .DataBodyRange(lRow, 15) = "n/a" 'ALOCREC
    .DataBodyRange(lRow, 16) = NUMPO 'NUMPO
    .DataBodyRange(lRow, 17) = "n/a"
    .DataBodyRange(lRow, 18) = NUMNF
    .DataBodyRange(lRow, 19) = "n/a"
    .DataBodyRange(lRow, 20) = Date
    .DataBodyRange(lRow, 21) = "n/a"
    .DataBodyRange(lRow, 22) = "n/a"
    
   End With

End Sub

Sub Histórico_Modificações_AttGR()

   Dim lRow As Long
   
   Dim tbRegistro As ListObject
   Set tbRegistro = ThisWorkbook.Sheets("Controle").ListObjects("tb_Registro")
   
   Dim tbHistórico As ListObject
   Set tbHistórico = ThisWorkbook.Sheets("Histórico de Modificações").ListObjects("tbHistórico")

   tbHistórico.ListRows.Add
   lRow = tbHistórico.DataBodyRange.Rows.count

   With tbHistórico '// PREENCHE OS CAMPOS DA TABELA COM AS INFORMAÇÕES INPUTADAS
   
      .DataBodyRange(lRow, 1) = Application.UserName
      .DataBodyRange(lRow, 2) = NUMPR
      .DataBodyRange(lRow, 3) = "Atualização Geral"
      .DataBodyRange(lRow, 4) = Date
      .DataBodyRange(lRow, clALOCREC + 3) = ALOCREC
      .DataBodyRange(lRow, clDATAPR + 3) = DATAPR
      .DataBodyRange(lRow, clREQ + 3) = REQ

      .DataBodyRange(lRow, clDESCRIC + 3) = DESCRIC
      .DataBodyRange(lRow, clPROJECT + 3) = Project
      .DataBodyRange(lRow, clCLIENTE + 3) = Cliente
      .DataBodyRange(lRow, clCNPJFORN + 3) = CNPJFORN
      .DataBodyRange(lRow, clTOTALBRL + 3) = TOTALBRL
      .DataBodyRange(lRow, clFORNECEDOR + 3) = Fornecedor
      .DataBodyRange(lRow, clNUMPO + 3) = CStr(NUMPO)
      If DATAPO <> "" Then .DataBodyRange(lRow, clDATAPO + 3) = CDate(DATAPO)
      .DataBodyRange(lRow, clNUMNF + 3) = NUMNF
      .DataBodyRange(lRow, clPREVFC + 3) = PREVFC
      If CODFORNEC <> "" Then .DataBodyRange(lRow, clCODFORNEC + 3) = CDate(CODFORNEC)
      .DataBodyRange(lRow, clENTREGANF + 3) = ENTREGANF
    
   End With
      
End Sub

Private Sub Class_Terminate()
    Set pPlanningDates = Nothing
End Sub

