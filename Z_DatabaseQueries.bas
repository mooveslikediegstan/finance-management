Attribute VB_Name = "Z_DatabaseQueries"
Option Explicit

Public Function ExecQuery(ByVal SQL As String, Optional Database As ADODB.Connection) As ADODB.Recordset
    
    If Database Is Nothing Then
        If DB Is Nothing Then Set DB = Conectar
        Set Database = DB
    End If
    
    Set ExecQuery = Database.Execute(SQL)
    
End Function

Public Sub ExecuteUpdateCommand(ByVal SQL As String, Optional Database As ADODB.Connection)
    
    If Database Is Nothing Then
        If DB Is Nothing Then Set DB = Conectar
        Set Database = DB
    End If

    Database.Execute (SQL)
    
End Sub

Public Function ExecUpdateRS(ByVal SQL As String) As ADODB.Recordset

    If DB Is Nothing Then Set DB = Conectar
    
    Dim rs As New ADODB.Recordset
    
    rs.Open Source:=SQL, ActiveConnection:=DB, CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, Options:=adCmdText
            
    Set ExecUpdateRS = rs
    
End Function

Public Function GetScalar(ByVal SQL As String) As Variant

    Dim rs As ADODB.Recordset: Set rs = ExecQuery(SQL)
    
    If rs Is Nothing Then
        GetScalar = Null
    Else
        If rs.EOF Then
            GetScalar = Null
        Else
            GetScalar = rs.Fields(0).Value
        End If
        SafeCloseRS rs
    End If
    
End Function

Public Sub SafeCloseRS(ByRef rs As ADODB.Recordset)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
    End If
End Sub

Function getAddress(ByVal sTableName As String) As String
    With Range(sTableName & "[#All]")
        getAddress = "[" & .Parent.Name & "$" & .Address(False, False) & "]"
    End With
End Function
