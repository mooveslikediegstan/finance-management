VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formDataPlan 
   ClientHeight    =   1060
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   3330
   OleObjectBlob   =   "formDataPlan.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "formDataPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_OK_Click()
    On Error GoTo ErrHndlr
    Data = CDate(formDataPlan.txt_DATAPLANSC)

   
    If formDataPlan.txt_DATAPLANSC = "" Then
    
        MsgBox "Insira a data de planejamento", vbOKOnly
        Exit Sub
        
    ElseIf formDataPlan.txt_VALORPLANSC = "" Then
    
        MsgBox "Insira o valor planejado", vbOKOnly
        Exit Sub
        
    Else:
        formControle.lbx_DATAPLANSC.AddItem
        index = formControle.lbx_DATAPLANSC.ListCount
        formControle.lbx_DATAPLANSC.List(index - 1, 0) = Data
        formControle.lbx_DATAPLANSC.List(index - 1, 1) = CCur(formDataPlan.txt_VALORPLANSC)

        formDataPlan.txt_VALORPLANSC = ""
        formDataPlan.txt_DATAPLANSC = ""
    End If


ErrHndlr:
End Sub

Private Sub btn_Voltar_Click()

    txt_DATAPLANSC = ""
    txt_VALORPLANSC = ""
    
    formDataPlan.Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
    
        Cancel = True

    End If

End Sub

  

'Sub aa()
'  With FormCadastro.ListBox_Peças
'      .RowSource = tbFiltroAvanc.DataBodyRange.Address(, , , 1)
'      .ColumnCount = 3
'      .ColumnWidths = "100;180;100"
'      .ColumnHeads = True
'      .TextAlign = fmTextAlignCenter
'  End With
'End Sub
'// MACROS RELACIONADAS À LISTBOX (doubleclick para inserir em outro listbox)


'Private Sub ListBox_Peças_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'  Dim QtdUtilizadaSKU As Long
'  QtdUtilizadaSKU = Application.InputBox("Insira a quantidade utilizada nesse SKU:", "Quantidade por produto", 1, Type:=1)
'  If QtdUtilizadaSKU <= 0 Then Exit Sub'''''

  'Dim Index As Long
  'Index = FormCadastro.ListBox_Peças.ListIndex
 '
 ' FormCadastro.ListBox_PeçasSKU.ColumnCount = 3
 ' FormCadastro.ListBox_PeçasSKU.ColumnWidths = "100;180;80"
 ' Dim QtdIndex As Long
 '
 ' QtdIndex = FormCadastro.ListBox_PeçasSKU.ListCount
 '
 ' If QtdIndex >= 1 Then
 '   For i = 0 To QtdIndex - 1
 '     If FormCadastro.ListBox_PeçasSKU.List(i, 0) = CStr(ListBox_Peças.List(Index, 0)) Then
 '       MsgBox "Item já adicionado.", vbOKOnly
 '       Exit Sub
 '     End If
 '   Next i
 ' End If
 '
 ' FormCadastro.ListBox_PeçasSKU.AddItem
 ' FormCadastro.ListBox_PeçasSKU.List(QtdIndex, 0) = ListBox_Peças.List(Index, 0)
 ' FormCadastro.ListBox_PeçasSKU.List(QtdIndex, 1) = ListBox_Peças.List(Index, 1)
 ' FormCadastro.ListBox_PeçasSKU.List(QtdIndex, 2) = QtdUtilizadaSKU
 '
  'QtdIndex = FormCadastro.ListBox_PeçasSKU.ListCount
'End Sub


'/INSERE AS INFORMAÇÕES DO LISTBOX NUMA CAIXA DE TEXTO (INDEX 0 PRIMEIRA COLUNA, INDEX 1 SEGUNDA COLUNA)


'Private Sub ListBox_CadastroPeça_Moldes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'  Index = Me.ListBox_CadastroPeça_Moldes.ListIndex
'  Me.txt_CadastroPeça_Molde = Me.ListBox_CadastroPeça_Moldes.List(Index, 0)
'  Me.txt_CadastroPeça_DescriçãoMolde = Me.ListBox_CadastroPeça_Moldes.List(Index, 1)
'End Sub


Private Sub UserForm_Click()

End Sub
