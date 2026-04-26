Attribute VB_Name = "A_"
'Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// CADASTRO DE NOVA PR (REQUISIÇÃO) ////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub A_CadastroPR()

    If formControle.txt_Cadastro_NUMPR = vbNullString Then
       MsgBox "Cadastre o N° da Requisição", vbCritical, "Campo obrigatório"
       Exit Sub
    End If
    
    If formControle.cbx_Cadastro_ALOCREC = vbNullString Then
       MsgBox "Selecione o sistema utilizado", vbCritical, "Campo obrigatório"
       Exit Sub
    End If

    If formControle.txt_Cadastro_DESCRIC = vbNullString Then
          MsgBox "Coloque a descrição da requisição", vbCritical, "Campo obrigatório"
       Exit Sub
    End If
    If formControle.cbx_Cadastro_FORNECEDOR = vbNullString Then
       MsgBox "Selecione ou insira o nome do fornecedor", vbCritical, "Campo obrigatório"
       Exit Sub
    End If
    If formControle.cbx_Cadastro_CLIENTE = vbNullString Then
       MsgBox "Selecione a Internal Order", vbCritical, "Campo obrigatório"
       Exit Sub
    End If
    If formControle.txt_Cadastro_TOTALBRL = vbNullString Then
       MsgBox "Insira o valor da requisição", vbCritical, "Campo obrigatório"
       Exit Sub
    End If
    
    Dim iRow As Long
    
    Dim PR_Existe As Boolean: PR_Existe = False
    
    InstanceTables
    
    iRow = SearchRowID(formControle.txt_Cadastro_NUMPR, M_BancoDeDados.tbPurchaseOrder, 1)
    
    If iRow <> -1 Then PR_Existe = True

    If PR_Existe = True Then
        MsgBox "Solicitação de compra já cadastrada!", vbCritical, "Cadastro inválido"
        Exit Sub
    End If
    
    Dim QtdParcelas As Long
    QtdParcelas = formControle.lbx_DATAPLANSC.ListCount - 1 '// o cabeçalho é o primeiro registro do listbox, por isso -1

    If formControle.chbx_Cadastro_POGUARDACHUVA.Value = True And QtdParcelas <= 0 Then
    
        MsgBox "Datas inconsistentes!", vbCritical, "Cadastro inválido"
        Exit Sub
    
    End If
    
    Dim ValorTotalSC As Currency, ValorParcialPlanejado As Currency, i As Long
    
    ValorTotalSC = CCur(formControle.txt_Cadastro_TOTALBRL)

    For i = 2 To formControle.lbx_DATAPLANSC.ListCount
        ValorParcialPlanejado = ValorParcialPlanejado + CCur(formControle.lbx_DATAPLANSC.List(i - 1, 1))
    Next i
    If formControle.chbx_Cadastro_POGUARDACHUVA.Value = True And ValorParcialPlanejado <> ValorTotalSC Then
     
        MsgBox "O valor da SC e as parcelas inseridas não somam o mesmo valor!", vbCritical, "Cadastro inválido"
        Exit Sub
        
    End If
     
    Dim PR As PurchaseOrder
    
    Set PR = New PurchaseOrder
     
    PR.CadastrarPR
    
    Set PR = Nothing
    
    MsgBox "Registro adicionado ao controle!", vbOKOnly, "Cadastro concluído"
    
    formControle.MultiPage1.Value = 0

End Sub

Sub A_Apagar_Info_CadastroPR()
   '// APAGANDO AS INFORMAÇÕES DO FORMULÁRIO
    formControle.txt_Cadastro_NUMPR = ""
    formControle.cbx_Cadastro_ALOCREC = ""
    formControle.txt_Cadastro_DATAPR = ""
    
    formControle.txt_Cadastro_REQARIBA = ""
    formControle.txt_Cadastro_DESCRIC = ""
    
    formControle.cbx_Cadastro_CLIENTE = ""
    formControle.txt_Cadastro_CODCLIENTE = ""
    formControle.cbx_Cadastro_PROJECT = ""
    
    formControle.txt_Cadastro_CODFORNEC = ""
    formControle.cbx_Cadastro_FORNECEDOR = ""
    formControle.txt_Cadastro_CNPJFORN = ""
    
    formControle.txt_Cadastro_TOTALBRL = ""
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// ATUALIZAÇÃO DA PURCHASE ORDER /////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub A_Tela_Att_PO()

   Dim NUMEROPR As Boolean
   NUMEROPR = True
   
   If formControle.txt_Menu_NUMPR = "" Then
      NUMEROPR = False
      formControle.txt_Menu_NUMPR.BackColor = &HC0C0FF
      MsgBox "Insira o N° da Requisição", vbCritical
      Exit Sub
   End If
   If NUMEROPR = True Then formControle.txt_Menu_NUMPR.BackColor = &HFFFFFF
   
   InstanceTables
   
   Dim lRegistro As Long, ExisteRegistro As Boolean
   ExisteRegistro = False
   
   
   '// VERIFICA SE O REGISTRO EXISTE NO CONTROLE
   
   lRegistro = SearchRowID(formControle.txt_Menu_NUMPR, M_BancoDeDados.tbPurchaseOrder, 1)
   
   If lRegistro <> -1 Then ExisteRegistro = True
   
   If ExisteRegistro = False Then
      MsgBox "Registro não cadastrado", vbCritical
      formControle.txt_Menu_NUMPR = vbNullString
      Exit Sub
   End If
   
   Dim PR As PurchaseOrder: Set PR = New PurchaseOrder
   
   PR.NUMPR = formControle.txt_Menu_NUMPR
   
   PR.LoadFromDatabaseToObject (PR.NUMPR)
   
   If PR.NUMPO <> vbNullString And Not IsNull(PR.DATAPO) Then
      MsgBox "A PO já foi atualizada na data : " & PR.DATAPO & " !"
      Exit Sub
   End If

   formControle.MultiPage1.Value = 2
   formControle.txt_attPO_NUMPR = formControle.txt_Menu_NUMPR
   Dim Hoje As String
   Hoje = Format(Date, "dd/mm/yyyy")
   formControle.txt_attPO_NUMPO = M_BancoDeDados.tbPurchaseOrder.DataBodyRange(lRegistro, PR.clNUMPO)
   formControle.txt_attPO_DATAPO = Hoje
   
End Sub

Sub A_Atualizar_PO()

    S_Desproteger_Planilha
    
    Dim Registro As New PurchaseOrder
    Registro.Atualizar_PO
    
    S_Proteger_Planilha
    
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// ATUALIZAÇÃO DA NOTA FISCAL /////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub A_Tela_Att_NF()

   Dim NUMERONF As Boolean
   NUMERONF = True
   
   If formControle.txt_Menu_NUMPR = "" Then
      NUMERONF = False
      formControle.txt_Menu_NUMPR.BackColor = &HC0C0FF
      MsgBox "Insira o N° da PO", vbCritical
      Exit Sub
   End If
   

   If NUMERONF = True Then formControle.txt_Menu_NUMPR.BackColor = &HFFFFFF
   
   Dim lRegistro As Long, ExisteRegistro As Boolean
   ExisteRegistro = False
   
   InstanceTables
   
   Dim SC As PurchaseOrder
   Set SC = New PurchaseOrder
   
   
   '// VERIFICA SE O REGISTRO EXISTE NO CONTROLE
   
   lRegistro = SearchRowID(formControle.txt_Menu_NUMPR, M_BancoDeDados.tbPurchaseOrder, SC.clNUMPO)
   If lRegistro <> -1 Then ExisteRegistro = True
   
   If ExisteRegistro = False Then
      MsgBox "A PO ainda não está cadastrada ou você não inseriu um N° válido de PO.", vbCritical
      formControle.txt_Menu_NUMPR = ""
      Exit Sub
   End If
   
   Dim Registro As New PurchaseOrder
   Registro.NUMPR = M_BancoDeDados.tbPurchaseOrder.DataBodyRange(lRegistro, SC.clNUMPR)
   Registro.LoadFromDatabaseToObject (Registro.NUMPR)
   
   Dim Registro_PONF As New Invoice
   
   Dim TotalGasto As Currency, i As Long
   TotalGasto = 0
   If tbInvoice.DataBodyRange Is Nothing Then tbInvoice.ListRows.Add
   
   For i = 1 To tbInvoice.DataBodyRange.Rows.count
      If tbInvoice.DataBodyRange(i, 1) = Registro.NUMPO Then
         TotalGasto = TotalGasto + tbInvoice.DataBodyRange(i, Registro_PONF.TotalAmount)
     End If
   Next i
   
   If TotalGasto < Registro.TOTALBRL Then
      formControle.MultiPage1.Value = 3
      formControle.txt_attNF_NUMPR = formControle.txt_Menu_NUMPR
      Dim Hoje As String
      Hoje = Format(Date, "dd/mm/yyyy")
      formControle.txt_ATTNF_DATAPOSTNF = Hoje
      formControle.cbx_ATTNF_ENTREGANF = Registro.ENTREGANF
      
   ElseIf TotalGasto >= Registro.TOTALBRL Then
      MsgBox "O somatório das Notas Fiscais é igual ou excede o valor da PO!", vbExclamation
      Exit Sub
   End If

End Sub
Sub A_Atualizar_NF()

   Dim Registro As New PurchaseOrder
   Registro.Atualizar_NF
   
End Sub






'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// ATUALIZAÇÃO GERAL DE REGISTRO /////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub A_Tela_Att_GERAL()

   Dim NUMERONF As Boolean
   NUMERONF = True
   
   If formControle.txt_Menu_NUMPR = "" Then
      NUMERONF = False
      formControle.txt_Menu_NUMPR.BackColor = &HC0C0FF
      MsgBox "Insira o N° da Requisição", vbCritical
      Exit Sub
   End If
   If NUMERONF = True Then formControle.txt_Menu_NUMPR.BackColor = &HFFFFFF
   
   Dim tbRegistro As ListObject: Set tbRegistro = ThisWorkbook.Sheets("Controle").ListObjects("tb_Registro")
   Dim lRegistro As Long, ExisteRegistro As Boolean
   ExisteRegistro = False
   
   
   '// VERIFICA SE O REGISTRO EXISTE NO CONTROLE
   For lRegistro = 1 To tbRegistro.DataBodyRange.Rows.count
      If formControle.txt_Menu_NUMPR = tbRegistro.DataBodyRange(lRegistro, 1) Then
         ExisteRegistro = True
         Exit For
      End If
   Next lRegistro
   If ExisteRegistro = False Then
      MsgBox "Registro não cadastrado", vbCritical
      formControle.txt_Menu_NUMPR = ""
      Exit Sub
   End If
   
   Dim PR As PurchaseOrder
   Set PR = New PurchaseOrder
   PR.NUMPR = formControle.txt_Menu_NUMPR
   PR.LoadFromDatabaseToObject (PR.NUMPR)
   
   If PR.ALOCREC = "CONCUR" Or PR.ALOCREC = "DEBT NOTES" Then
      formControle.txt_Menu_NUMPR = ""
      MsgBox "Não é possível atualizar registros para os sistemas CONCUR e DEBT NOTES !"
      Exit Sub
   End If
   
   formControle.txt_attGR_NUMPR = PR.NUMPR
   formControle.cbx_attGR_ALOCREC = PR.ALOCREC
   formControle.txt_attGR_DATAPR = Format(PR.DATAPR, "dd/mm/yyyy")
   formControle.cbx_attGR_REQ = PR.REQ

   formControle.txt_attGR_DESCRIC = PR.DESCRIC
   formControle.cbx_attGR_FORNECEDOR = PR.Fornecedor
   formControle.txt_attGR_PROJECT = PR.Project
   formControle.cbx_attGR_CLIENTE = PR.Cliente
   formControle.txt_attGR_CNPJFORN = PR.CNPJFORN
   formControle.txt_attGR_TOTALBRL = Format(PR.TOTALBRL, "Currency")
   formControle.txt_attGR_NUMPO = PR.NUMPO
   formControle.txt_attGR_DATAPO = PR.DATAPO
   formControle.txt_attGR_NUMNF = PR.NUMNF
   formControle.cbx_attGR_PREVFC = PR.PREVFC
   formControle.txt_attGR_DATAPREVSC = PR.CODFORNEC
   'formControle.txt_attGR_FATURDIR = PR.FATURDIR
   
   formControle.MultiPage1.Value = 4
   
   Set PR = Nothing

   
End Sub

Sub A_Atualizar_Registro()
  
   Dim Registro As New PurchaseOrder
   Registro.Atualizar_Registro
   
End Sub

Sub Voltar_Registro_Geral()

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
   formControle.txt_attGR_FATURDIR = ""
   formControle.cbx_attGR_STATUS = ""
   formControle.txt_Menu_NUMPR = ""
   
   formControle.MultiPage1.Value = 0


End Sub

