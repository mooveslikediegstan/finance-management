Attribute VB_Name = "Y_Calendar"
Option Explicit
'Vetor que armazena todos os Label de dia do Calendário
Dim Rótulos() As New Calendar

Function GetCalendario() As Date
        
    Dim lTotalRótulos As Long
    Dim ctrl As Control
    Dim frm As formCalendar
    
    Set frm = New formCalendar
    
    'Atribui cada um dos Label num elemento do vetor da classe
    For Each ctrl In frm.Controls
        If ctrl.Name Like "l?c?" Then
            lTotalRótulos = lTotalRótulos + 1
            ReDim Preserve Rótulos(1 To lTotalRótulos)
            Set Rótulos(lTotalRótulos).lblGrupo = ctrl
        End If
    Next ctrl

    frm.Show
    
    'Se a data escolhida for nula ou inválida, retorna-se a data atual:
    If IsDate(frm.Tag) Then
        GetCalendario = frm.Tag

    Else
        GetCalendario = Date
    End If
        
    Unload frm

End Function



