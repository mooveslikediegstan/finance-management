Attribute VB_Name = "C_AnexosOld"
Public Arquivo_Anexado As Boolean
Public ANO As String

Option Explicit

Sub VAR_ANO()

   ANO = "2024"

End Sub

Sub C_Anexar_Orcamento()

   Call VAR_ANO

    Dim sourcePath As String
    Dim destinationPath As String
    Dim fileName As String, PR_Name As String
    Dim fullSourcePath As String
    Dim fullDestinationPath As String
    Dim userProfile As String
    
    Arquivo_Anexado = True
    
    MsgBox "Por favor, selecione o arquivo a anexar", vbInformation

    ' Prompt the user to select a file
    With Application.fileDialog(msoFileDialogFilePicker)
        .Title = "Selecione o arquivo do ORÇAMENTO"
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
        If .Show = -1 Then
            sourcePath = Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\"))
            fileName = mid(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") + 1)
        Else
            Arquivo_Anexado = False
            Exit Sub ' User canceled
        End If
    End With
    
    userProfile = Environ("USERPROFILE")
    

    
    
    ' Specify the destination directory
    destinationPath = userProfile & "\Versuni\ENG_DA - General\15_Compras\13_Planilha_de_Controle_SC\" & ANO & "\Cotações e Orçamentos\" ' Change to your desired destination
    
    PR_Name = formControle.txt_Cadastro_NUMPR
    
    ' Construct full paths
    fullSourcePath = sourcePath & fileName
    fullDestinationPath = destinationPath & PR_Name & " - " & fileName
    
    ' Check if the destination directory exists, create if not
    If Dir(destinationPath, vbDirectory) = "" Then
        MkDir destinationPath
    End If
    
    ' Transfer the file
    FileCopy fullSourcePath, fullDestinationPath
    
    ' Inform the user
    MsgBox "Orçamento anexado com sucesso!", vbInformation
    
End Sub

Sub C_Anexar_PO()

   Call VAR_ANO

    Dim sourcePath As String
    Dim destinationPath As String
    Dim fileName As String, PO_Name As String
    Dim fullSourcePath As String
    Dim fullDestinationPath As String
    Dim userProfile As String
    
   Arquivo_Anexado = True
    
    MsgBox "Por favor, selecione o arquivo a anexar", vbInformation

    ' Prompt the user to select a file
    With Application.fileDialog(msoFileDialogFilePicker)
        .Title = "Selecione o arquivo da PO"
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
        If .Show = -1 Then
            sourcePath = Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\"))
            fileName = mid(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") + 1)
        Else
            Arquivo_Anexado = False
            Exit Sub ' User canceled
        End If
    End With
    
    userProfile = Environ("USERPROFILE")
    
    
    ' Specify the destination directory
    destinationPath = userProfile & "\Versuni\ENG_DA - General\15_Compras\13_Planilha_de_Controle_SC\" & ANO & "\PO\" ' Change to your desired destination
    
    PO_Name = formControle.txt_attPO_NUMPO
    
    ' Construct full paths
    fullSourcePath = sourcePath & fileName
    fullDestinationPath = destinationPath & "PO" & PO_Name & " - " & fileName
    
    ' Check if the destination directory exists, create if not
    If Dir(destinationPath, vbDirectory) = "" Then
        MkDir destinationPath
    End If
    
    ' Transfer the file
    FileCopy fullSourcePath, fullDestinationPath
    
    ' Inform the user
    MsgBox "PO anexada com sucesso!", vbInformation
    
End Sub

Sub C_Anexar_NF()

   Call VAR_ANO

    Dim sourcePath As String
    Dim destinationPath As String
    Dim fileName As String, NF_Name As String
    Dim fullSourcePath As String
    Dim fullDestinationPath As String
    Dim userProfile As String
    
   Arquivo_Anexado = True

    ' Prompt the user to select a file
    
    MsgBox "Por favor, selecione o arquivo a anexar", vbInformation
    
    With Application.fileDialog(msoFileDialogFilePicker)
        .Title = "Selecione o arquivo da NF"
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
        If .Show = -1 Then
            sourcePath = Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\"))
            fileName = mid(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") + 1)
        Else
            Arquivo_Anexado = False
            Exit Sub ' User canceled
        End If
    End With
    
    userProfile = Environ("USERPROFILE")
    
    
    ' Specify the destination directory
    destinationPath = userProfile & "\Versuni\ENG_DA - General\15_Compras\13_Planilha_de_Controle_SC\" & ANO & "\NF\" ' Change to your desired destination
    
    
    NF_Name = formControle.txt_ATTNF_NUMNF
    
    ' Construct full paths
    fullSourcePath = sourcePath & fileName
    fullDestinationPath = destinationPath & NF_Name & " - " & fileName
    
    ' Check if the destination directory exists, create if not
    If Dir(destinationPath, vbDirectory) = "" Then
        MkDir destinationPath
    End If
    
    ' Transfer the file
    FileCopy fullSourcePath, fullDestinationPath
    
    ' Inform the user
    MsgBox "NF anexada com sucesso!", vbInformation
    
End Sub
