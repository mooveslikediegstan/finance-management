Attribute VB_Name = "Z_Log"
    'Option Explicit

Function OldNewLogText(ByVal IdentifierName, ByVal Identifier, ByVal FieldName As String, oldValue As Variant, NewValue As Variant)

    OldNewLogText = "{" & IdentifierName & ":" & Identifier & "; " & FieldName & ": {old:" & oldValue & "; new:" & NewValue & "}}"
    
End Function

Function generate_LogText(ByRef RecSet)

    Dim collRecordSet As Collection: Set collRecordSet = New Collection
    
    Dim i As Long
    For i = 0 To RecSet.Fields.count - 1
        collRecordSet.Add RecSet.Fields(i).Name & ":" & RecSet.Fields(i)
    Next i
    
    Dim Record As Variant, logText As Variant
    
    For Each Record In collRecordSet
    
        If logText = "" Then
            logText = Record
        Else:
            logText = logText & ";" & Record
        End If
        
    Next Record
    
    generate_LogText = "{" + logText + "}"
    
End Function

Sub Create_Log(Table As String, Operation As String, logText As String)

    If DB Is Nothing Then Set DB = Conectar
    
    Dim LogRecordSet As New ADODB.Recordset

    'Set LogRecordSet = UPDATE_accessDB("log")
    
    LogRecordSet.AddNew
    
    LogRecordSet.Fields("table") = Table
    LogRecordSet.Fields("operation") = Operation
    LogRecordSet.Fields("log_text") = logText
    LogRecordSet.Fields("user") = Application.UserName
    LogRecordSet.Fields("timestamp") = Now()

    LogRecordSet.Update
    
    Set LogRecordSet = Nothing

End Sub

