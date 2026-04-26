Attribute VB_Name = "Z_DatabaseRecords"
Option Explicit

Public Function accessDB_to_array(RecSet As ADODB.Recordset) As Variant
    
    If Not RecSet.EOF And Not RecSet.BOF Then
        accessDB_to_array = arrayTranspose(RecSet.GetRows(294), 0)
    Else:
        accessDB_to_array = Array()
    End If
    
End Function

Public Function accessDB_to_recordset(ByVal Table As String, Optional ByVal filterField = "", Optional ByVal filterParameter As Variant = "", Optional ByVal customQueryString = "", Optional ByVal orderString = "") As Variant
            
    Dim filterString As String
    If filterField = "" Or filterParameter = "" Then
    Else
        filterString = " WHERE " & filterField & " = " & filterParameter
    End If
    
    Dim QueryString As String: QueryString = "SELECT * FROM " & Table & filterString & orderString

    If customQueryString = "" Then
        Set accessDB_to_recordset = ExecQuery(QueryString)
    Else
        Set accessDB_to_recordset = ExecQuery(customQueryString)
    End If
    
End Function

Public Function UPDATE_accessDB(ByVal Table As String, Optional ByVal filterField = "", Optional ByVal filterParameter = "", Optional ByVal customQueryString = "")
    
    Dim filterString As String
    If filterField = "" Or filterParameter = "" Then
    Else
        filterString = " AND " & Table & "_" & filterField & " = '" & filterParameter & "'"
    End If
    
    Dim QueryString As String: QueryString = "SELECT * FROM " & Table & " WHERE DELETE = False" & filterString
    
    If customQueryString = "" Then
        Set UPDATE_accessDB = ExecUpdateRS(QueryString)
    Else:
        Set UPDATE_accessDB = ExecUpdateRS(customQueryString)
    End If
    
End Function

























