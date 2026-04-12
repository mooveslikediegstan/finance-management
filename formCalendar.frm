VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCalendar 
   Caption         =   "Calendário"
   ClientHeight    =   2220
   ClientLeft      =   135
   ClientTop       =   705
   ClientWidth     =   2685
   OleObjectBlob   =   "formCalendar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "formCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    'A data inicial é a atual:
    lblHoje = "Hoje: " & Format(Date, "dd/mm/yyyy")
    sb = Year(Date) * 12 + Month(Date)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Impede que se dê Unload no formulário, caso contrário a linha que testa
    'frm.Tag na linha seguinte do módulo mdlCalendário dará erro, pois o objeto
    'deixará de existir. Ao invés de dar Unload, usa-se Hide para o objeto
    'continuar a existir na memória.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If
End Sub

Private Sub lblHoje_Click()
    'Quando se clica no Label do dia atual, o calendário atualiza-se
    'para o mês atual.
    
    'O modo de cálculo do mês em questão é o número de meses.
    'Como um ano possui 12 meses, o valor da ScrollBar é o número
    'total de meses:
    sb = Year(Date) * 12 + Month(Date)
    
End Sub

Private Sub sb_Change()
    'Deve-se atualizar o calendário ao alterar a ScrollBar.
    'O valor do calendário é uma divisão inteira (observe o símbolo \)
    'de anos e o resto do valor por 12 como mês:
    Atualizar DateSerial(sb \ 12, sb Mod 12, 1)
End Sub

Private Sub Atualizar(dt As Date)
    'Rotina que atualiza todos os Label do calendário
    
    Dim l As Long
    Dim c As Long
    Dim cInício As Long
    Dim dtDia As Date
    Dim ctrl As Control
    
    lblMêsAno = Format(dt, "mmmm yyyy")
    
    For l = 1 To 6                               'Linhas do calendário
        For c = 1 To 7                           'Colunas do calendário
            Set ctrl = Controls("l" & l & "c" & c)
            'O entendimento da linha abaixo é fundamental para entender como todos os
            'labels foram povoados:
            dtDia = DateSerial(Year(dt), Month(dt), (l - 1) * 7 + c - Weekday(dt) + 1)
            ctrl.Caption = Format(Day(dtDia), "00")
            ctrl.Tag = dtDia

            'Dias de um mês diferente do mês visualizado ficarão na cor cinza claro:
            If Month(dtDia) <> Month(dt) Then
                ctrl.ForeColor = &H808080
            Else
                ctrl.ForeColor = &H0
            End If
            'Realçar dia atual presente, caso esteja visível no calendário:
            If dtDia = Date Then
                ctrl.ForeColor = &HFF&
            End If
        Next c
    Next l

End Sub

