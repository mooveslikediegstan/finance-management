Attribute VB_Name = "Z_DatabaseConnection"
Option Explicit

Public DB As ADODB.Connection

Function Conectar()

    Dim fDB As ADODB.Connection
    Set fDB = New ADODB.Connection

    Dim dbpath As String
    dbpath = databasePath()
    
    With fDB
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbpath _
                & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
        .Open
    End With

    Set Conectar = fDB
    
End Function

Public Function databasePath() As String
    
    Dim userProfile As String, databaseFolder As String, databaseFile As String
    
    userProfile = GetUserProfile
    
    databaseFolder = GetDBFolder
    
    databaseFile = "\" & ThisWorkbook.Name

    databasePath = userProfile & databaseFolder & databaseFile 'retorna o caminho do BD desejado
    
End Function


Function IsTestWorkbook() As Boolean

    Dim ThisWBName As String
    
    ThisWBName = ThisWorkbook.Name
    
    If Left(ThisWBName, 4) = "TEST" Then IsTestWorkbook = True Else IsTestWorkbook = False

End Function

Function GetUserProfile()

    If IsTestWorkbook Then
        GetUserProfile = ""
    Else
        GetUserProfile = "" 'Environ("USERPROFILE")
    End If
    
End Function

Function GetDBFolder()

    If IsTestWorkbook Then
        GetDBFolder = "C:\Users\Diego.Stanzani\OneDrive - AgGrowth\06_Montagem\Att_Controle_Financeiro(20260401)"
    Else
        GetDBFolder = "C:\Users\Diego.Stanzani\OneDrive - AgGrowth\06_Montagem\Att_Controle_Financeiro(20260401)"
    End If
    
End Function


Sub Desconectar(ByRef DB)
    
    If Not DB Is Nothing Then DB.Close
    
    Set DB = Nothing
    
End Sub



