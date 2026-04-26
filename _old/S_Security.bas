Attribute VB_Name = "S_Security"
Sub S_Desproteger_Planilha()

    Senha = PegarSenha

    WSPURCHASEORDER.Unprotect Password:=Senha
    WSINVOICE.Unprotect Password:=Senha
    WSPURCHASEORDERPLANNING.Unprotect Password:=Senha

End Sub

Sub S_Proteger_Planilha()

    Senha = PegarSenha

    WSPURCHASEORDER.Protect Password:=Senha, UserInterfaceOnly:=True, AllowFiltering:=True
    WSINVOICE.Protect Password:=Senha, UserInterfaceOnly:=True, AllowFiltering:=True
    WSPURCHASEORDERPLANNING.Protect Password:=Senha, UserInterfaceOnly:=True, AllowFiltering:=True
   
End Sub

Function PegarSenha() As String

    PegarSenha = "_Agi.2026!"

End Function

Sub Exibir_Formulário_Manual()

   formControle.Show

End Sub



