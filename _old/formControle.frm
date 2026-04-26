VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formControle 
   ClientHeight    =   6530
   ClientLeft      =   195
   ClientTop       =   750
   ClientWidth     =   9840.001
   OleObjectBlob   =   "formControle.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "formControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_Menu_AbrirMD_Click()
   Call B_GoTo_MasterData
End Sub



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////// CADASTRO //////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'// BOTÃO PARA IR PARA TELA DE CADASTRO
Private Sub btnMenu_CadastroPR_Click()
   '// DATA DO ALOCRECEMA
   Dim Hoje As String
   Hoje = CStr(Format(Date, "dd/mm/yyyy"))
   txt_Cadastro_DATAPR = Hoje

   
   txt_Cadastro_REQARIBA.Value = UCase(Application.UserName)
   cbx_Cadastro_ENTREGANF = "PENDENTE"
   '// TROCA TELA PARA CADASTRO
   MultiPage1.Value = 1
End Sub

Private Sub btnCalendario_DtAbertura_Click()
    txt_Cadastro_DATAPR = Format(GetCalendario, "dd/mm/yyyy")
End Sub

Private Sub CadastrarFornecedor_Click()

    SupplierRegistry.FrameFornecedor.Caption = "Cadastrar Fornecedor"
    SupplierRegistry.btnSearchSupplier.Visible = False
    SupplierRegistry.ModifySupplier.Visible = False
    SupplierRegistry.RegisterNewSupplier.Visible = True
    SupplierRegistry.Show
    
End Sub
Private Sub AlterarFornecedor_Click()

    SupplierRegistry.FrameFornecedor.Caption = "Modificar Fornecedor"
    SupplierRegistry.txtSupplierID.Locked = True
    SupplierRegistry.btnSearchSupplier.Visible = True
    SupplierRegistry.ModifySupplier.Visible = True
    SupplierRegistry.RegisterNewSupplier.Visible = False
    SupplierRegistry.Show
    
End Sub

Private Sub cbx_attGR_STATUS_DropButtonClick()

   cbx_attGR_STATUS.List = Array("Ativo", "Cancelado")

End Sub

Private Sub cbx_ATTNF_ENTREGANF_DropButtonClick()
    cbx_ATTNF_ENTREGANF.List = Array("PENDENTE", "PARCIAL", "ENTREGUE")
End Sub



Private Sub cbx_ATTNF_MESCOMP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         KeyAscii = 0
End Sub
Private Sub cbx_ATTNF_ANOCOMP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         KeyAscii = 0
End Sub
Private Sub cbx_ATTNF_ENTREGANF_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         KeyAscii = 0
End Sub

Private Sub cbx_attPO_DataManual_Change()

   If cbx_attPO_DataManual.Value = False Then txt_attPO_DATAPO.Locked = True
   If cbx_attPO_DataManual.Value = True Then txt_attPO_DATAPO.Locked = False

End Sub


'// ALOCAÇÃO DE RECURSO UTILIZADO
Private Sub cbx_Cadastro_ALOCREC_DropButtonClick()
   cbx_Cadastro_ALOCREC.List = Array("Miscelanious", "Travel Expenses", "Payroll", "Contractors Cost", "Insurance")
End Sub
Private Sub cbx_Cadastro_ALOCREC_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub


'// ENTREGA NF
Private Sub cbx_Cadastro_ENTREGANF_DropButtonClick()
    cbx_Cadastro_ENTREGANF.List = Array("SIM", "NÃO")
End Sub
Private Sub cbx_Cadastro_ENTREGANF_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub



Private Sub CommandButton3_Click()

End Sub

'//FORNECEDOR

Private Sub OpenSupplierSearch_Click()

    Dim slSupplier As Supplier
    
    Set slSupplier = getSupplier
    
    If slSupplier Is Nothing Then Exit Sub
    
    formControle.txt_Cadastro_CODFORNEC = slSupplier.SupplierID
    formControle.cbx_Cadastro_FORNECEDOR = slSupplier.Name
    formControle.txt_Cadastro_CNPJFORN = slSupplier.CNPJ
    
    Set slSupplier = Nothing

End Sub


Private Sub OpenCustomerSearch_Click()

    Dim slCustomer As Customer
    
    Set slCustomer = getCustomer
    
    If slCustomer Is Nothing Then Exit Sub
    
    formControle.txt_Cadastro_CODCLIENTE = slCustomer.CustomerID
    formControle.cbx_Cadastro_CLIENTE = slCustomer.CustomerShortName
    
    Set slCustomer = Nothing

End Sub

Private Sub cbx_Cadastro_CLIENTE_Change()
    cbx_Cadastro_PROJECT.Clear
End Sub

Private Sub cbx_Cadastro_PROJECT_DropButtonClick() '// SELEÇÃO DO NUMERO DO PROJETO BASEADO NO CLIENTE
    
    cbx_Cadastro_PROJECT.List = getProjectsFromCustomer(txt_Cadastro_CODCLIENTE)

End Sub









'// VALOR TOTAL DA PURCHASE REQUEST
Private Sub txt_Cadastro_TOTALBRL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   Select Case KeyAscii
      Case 44, 48 To 57
      Case Else
         KeyAscii = 0
   End Select
End Sub

'// BOTÃO 'VOLTAR' DO CADASTRO DE NOVA PURCHASE ORDER
Private Sub btn_Cadastro_Voltar_Click()
   Call A_Apagar_Info_CadastroPR
   MultiPage1.Value = 0
End Sub


Private Sub chbx_Cadastro_POGUARDACHUVA_Change()
    
    If Me.chbx_Cadastro_POGUARDACHUVA.Value = True Then

        formControle.btn_addDATA.Visible = True
        formControle.btn_rmvDATA.Visible = True
        
        lbl_Datas_Planejadas.Visible = False
        txt_Cadastro_DATAPLAN = ""
        txt_Cadastro_DATAPLAN.Visible = False
    
        formControle.lbx_DATAPLANSC.Visible = True
        formControle.lbx_DATAPLANSC.ColumnCount = 2
        formControle.lbx_DATAPLANSC.TextAlign = fmTextAlignCenter
        formControle.lbx_DATAPLANSC.ColumnWidths = "100;100"
        formControle.lbx_DATAPLANSC.AddItem
        formControle.lbx_DATAPLANSC.List(0, 0) = "Data"
        formControle.lbx_DATAPLANSC.List(0, 1) = "Valor"
    
    Else
    
        formControle.btn_addDATA.Visible = False
        formControle.btn_rmvDATA.Visible = False
        
        lbl_Datas_Planejadas.Visible = True
        txt_Cadastro_DATAPLAN.Visible = True
        
        formControle.lbx_DATAPLANSC.Visible = False
        formControle.lbx_DATAPLANSC.Clear
    
    End If

End Sub

Private Sub btn_addDATA_Click()
    formDataPlan.Show
End Sub
Private Sub btn_rmvDATA_Click()

    If Me.lbx_DATAPLANSC.ListIndex > 0 Then
          
        Dim index As Long
        index = Me.lbx_DATAPLANSC.ListIndex
        Me.lbx_DATAPLANSC.RemoveItem (index)
    
    Else
    End If
    
End Sub

'// SUB PARA CADASTRAR NOVA PURCHASE ORDER
Private Sub btn_Cadastro_RealizarCadastro_Click()
   Call A_CadastroPR
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////// ATUALIZAÇÃO PO ////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'// IR PARA A TELA DE ATUALIZAÇÃO DE PO
Private Sub btnMenu_AttPO_Click()
   Call A_Tela_Att_PO
End Sub

'//N° DA PO
Private Sub txt_ATTPO_NUMPO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
   End If
End Sub

Private Sub btnCalendarioDataPO_Click()
    txt_attPO_DATAPO = Format(GetCalendario, "dd/mm/yyyy")
End Sub

'// DATA DE ATUALIZAÇÃO DA PO
Private Sub txt_ATTPO_DATAPO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   txt_attPO_DATAPO.MaxLength = 10
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
   End If
 End Sub
Private Sub txt_ATTPO_DATAPO_Change()
   If Len(txt_attPO_DATAPO) = 2 Or Len(txt_attPO_DATAPO) = 5 Then
      txt_attPO_DATAPO.Text = txt_attPO_DATAPO.Text & "/"
      SendKeys "{End}", True
   End If
End Sub

'// EXECUTAR ATUALIZAÇÃO DA PO
Private Sub img_attPO_CadastrarPO_Click()
   Call A_Atualizar_PO
End Sub

'// BOTÃO VOLTAR DA TELA ATT PO
Private Sub img_attPO_Voltar_Click()
   txt_attPO_NUMPR = ""
   txt_attPO_NUMPO = ""
   txt_attPO_DATAPO = ""
   txt_Menu_NUMPR = ""
   MultiPage1.Value = 0
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////// ATUALIZAÇÃO NF ////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub btnMenu_AttNF_Click()
   Call A_Tela_Att_NF
End Sub

'// TIPO DA NOTA FISCAL
Private Sub cbx_Cadastro_PREVFC_DropButtonClick()
   cbx_Cadastro_PREVFC.List = Array("SIM", "NÃO")
End Sub

'// VALOR DA NF
Private Sub txt_ATTNF_TOTALBRL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   Select Case KeyAscii
      Case 44, 48 To 57
      Case Else
         KeyAscii = 0
   End Select
End Sub


Private Sub btnCalendarPostDate_Click()
    txt_ATTNF_DATAPOSTNF = Format(GetCalendario, "dd/mm/yyyy")
End Sub

Private Sub btnCalendarPayDate_Click()
    txt_ATTNF_DATAPREVPGTONF = Format(GetCalendario, "dd/mm/yyyy")
End Sub



'// BOTÃO VOLTAR DA TELA ATT NF
Private Sub img_attNF_Voltar_Click()
   txt_attNF_NUMPR = ""
   txt_ATTNF_NUMNF = ""
   txt_Cadastro_CODFORNEC = ""
   txt_Menu_NUMPR = ""
   MultiPage1.Value = 0
End Sub

Private Sub img_attNF_AtualizarNF_Click()
   Call A_Atualizar_NF
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////// ATUALIZAÇÃO GERAL DE REGISTRO /////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Private Sub btnMenu_AttGR_Click()
   Call A_Tela_Att_GERAL
End Sub

'// ALOCRECEMA UTILIZADO
Private Sub cbx_attGR_ALOCREC_DropButtonClick()
   cbx_attGR_ALOCREC.List = Array("ARIBA", "CONCUR", "DEBT NOTES")
End Sub
Private Sub cbx_attGR_ALOCREC_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

'// REQUISITANTE
Private Sub cbx_attGR_REQ_DropButtonClick()
   Dim tbColaborador As ListObject: Set tbColaborador = ThisWorkbook.Sheets("Login").ListObjects("tbUsuarios")
   cbx_attGR_REQ.List = tbColaborador.ListColumns(3).DataBodyRange.Value
   Set tbColaborador = Nothing
End Sub
Private Sub cbx_attGR_REQ_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

'// REQUISITANTE ARIBA
Private Sub cbx_attGR_REQARIBA_DropButtonClick()
   Dim tbColaborador As ListObject: Set tbColaborador = ThisWorkbook.Sheets("Login").ListObjects("tbUsuarios")
   cbx_attGR_REQARIBA.List = tbColaborador.ListColumns(3).DataBodyRange.Value
   Set tbColaborador = Nothing
End Sub
Private Sub cbx_attGR_REQARIBA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

'//FORNECEDOR
Private Sub cbx_attGR_FORNECEDOR_DropButtonClick()
   Dim tbFornecedor As ListObject: Set tbFornecedor = ThisWorkbook.Sheets("MasterData").ListObjects("tb_FORNECEDOR")
   cbx_attGR_FORNECEDOR.List = tbFornecedor.ListColumns(1).DataBodyRange.Value
   Set tbFornecedor = Nothing
End Sub
'Private Sub cbx_attGR_FORNECEDOR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'   KeyAscii = 0
'End Sub

'//PARAMETRIZANDO A INTERNAL ORDER PELA DESCRIÇÃO
Private Sub cbx_attGR_CLIENTE_DropButtonClick()
   Dim tbIO As ListObject: Set tbIO = ThisWorkbook.Sheets("MasterData").ListObjects("tb_IO")
   cbx_attGR_CLIENTE.List = tbIO.ListColumns(2).DataBodyRange.Value
   Set tbIO = Nothing
End Sub
Private Sub cbx_attGR_CLIENTE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub
Private Sub cbx_attGR_CLIENTE_Change()
   Dim tbIO As ListObject: Set tbIO = ThisWorkbook.Sheets("MasterData").ListObjects("tb_IO")
   Dim Descric_IO As String: Descric_IO = cbx_attGR_CLIENTE.Value
   Dim IO_Encontrado As Boolean: IO_Encontrado = False
   Dim i As Long
   
   For i = 1 To tbIO.DataBodyRange.Rows.count
      If Descric_IO = tbIO.DataBodyRange(i, 2) Then
         IO_Encontrado = True
         Exit For
      End If
   Next i
   If IO_Encontrado = True Then
      txt_attGR_PROJECT = tbIO.DataBodyRange(i, 1)
      txt_attGR_CNPJFORN = tbIO.DataBodyRange(i, 4)
   Else:
      txt_attGR_PROJECT = ""
      txt_attGR_CNPJFORN = ""
   End If
   Set tbIO = Nothing
End Sub


'// VALOR TOTAL DA PURCHASE REQUEST
Private Sub txt_attGR_TOTALBRL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   Select Case KeyAscii
      Case 44, 48 To 57
      Case Else
         KeyAscii = 0
   End Select
End Sub

'//N° DA PO
Private Sub txt_attGR_NUMPO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
   End If
End Sub

'// DATA DE ATUALIZAÇÃO DA PO
Private Sub txt_attGR_DATAPO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   txt_attGR_DATAPO.MaxLength = 10
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
   End If
 End Sub
Private Sub txt_attGR_DATAPO_Change()
   If Len(txt_attGR_DATAPO) = 2 Or Len(txt_attGR_DATAPO) = 5 Then
      txt_attGR_DATAPO.Text = txt_attGR_DATAPO.Text & "/"
      SendKeys "{End}", True
   End If
End Sub

'// TIPO DA NOTA FISCAL
Private Sub cbx_attGR_PREVFC_DropButtonClick()
   cbx_attGR_PREVFC.List = Array("Produto", "Serviço")
End Sub

'// ATUALIZAÇÃO DA DATA DE PEDIDO DE INPUT DA NF
Private Sub txt_attGR_DATAPREVSCKeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   txt_attGR_DATAPREVSC.MaxLength = 10
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0
   End If
 End Sub
Private Sub txt_attGR_DATAPREVSC_Change()
   If Len(txt_attGR_DATAPREVSC) = 2 Or Len(txt_attGR_DATAPREVSC) = 5 Then
      txt_attGR_DATAPREVSC.Text = txt_attGR_DATAPREVSC.Text & "/"
      SendKeys "{End}", True
   End If
End Sub

Private Sub img_attGR_Voltar_Click()
   Call Voltar_Registro_Geral
End Sub

Private Sub img_attGR_AtualizarRegistro_Click()
   Call A_Atualizar_Registro
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Private Sub UserForm_Initialize()

   MultiPage1.Value = 0
   MultiPage1.Style = fmTabStyleNone

End Sub
