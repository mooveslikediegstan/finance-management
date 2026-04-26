Attribute VB_Name = "Z_ClassQueries"
Option Explicit

Function GetRecordSetHeader(ByVal RecSet As ADODB.Recordset)
    
    Dim HeaderArray()
    ReDim HeaderArray(0 To 0, 0 To RecSet.Fields.count - 1)
    
    Dim j As Long
    For j = 0 To UBound(HeaderArray, 2)
        HeaderArray(0, j) = RecSet.Fields(j).Name
    Next j
    
    GetRecordSetHeader = HeaderArray
    
End Function

Function GetDataArrayFromDB(query As String, Optional ByVal WithHeader As Boolean = False)
    
    Dim RecSet As ADODB.Recordset: Set RecSet = ExecQuery(query)
    
    If WithHeader Then
        
        GetDataArrayFromDB = AppendArrays(GetRecordSetHeader(RecSet), accessDB_to_array(RecSet))
    
    Else:
    
        GetDataArrayFromDB = accessDB_to_array(RecSet)
    
        SafeCloseRS RecSet
    
    End If
    
End Function

Function GetCitiesFromState(StateComboBox As MSForms.Control) As Variant
    
    If TypeName(StateComboBox) <> "ComboBox" Then Exit Function
    
    GetCitiesFromState = GetDataArrayFromDB("SELECT city_name FROM CITY WHERE State = '" & StateComboBox.Value & "' AND DELETE = False ORDER BY city_name ASC")
    
End Function

Function GetAllStates() As Variant
        
    GetAllStates = GetDataArrayFromDB("SELECT DISTINCT state FROM CITY WHERE DELETE = False ORDER BY state ASC")
    
End Function

Function getProjectsFromCustomer(ByVal Criteria1 As String) As Variant

    Dim TableQueryRange As String
    
    TableQueryRange = getAddress("ProjectTable")
    
    Dim QryString As String
    
    QryString = "SELECT [Projeto] FROM " & TableQueryRange
                  
    If Criteria1 <> "" Then
        QryString = QryString & " WHERE [ID Cliente] LIKE '%" & Criteria1 & "%'"
    End If
    
    getProjectsFromCustomer = GetDataArrayFromDB(query:=QryString, WithHeader:=False)
    
End Function


Function ItemExistsInLocalDB(SearchParameter As Variant, SearchField As String, Table As String) As Boolean

    Dim QueryString As String
    QueryString = "SELECT DISTINCT " & SearchField & " FROM " & Table & " WHERE " & SearchField & " = '" & SearchParameter & "' AND DELETE = False"
    
    If IsNull(GetScalar(QueryString)) Then ItemExistsInLocalDB = False Else ItemExistsInLocalDB = True
    
End Function


Function getSupplierListForDisplay(ByVal sTableName As String, ByVal Criteria1 As String, Optional ByVal Criteria2 As String = "")
    
    Dim TableQueryRange As String
    
    TableQueryRange = getAddress(sTableName)
    
    Dim QryString As String
    
    QryString = "SELECT * FROM " & TableQueryRange
                  
    If Criteria1 <> "" Then
        QryString = QryString & " WHERE [NOME EMPRESARIAL] LIKE '%" & Criteria1 & "%' OR [NOME FANTASIA] LIKE '%" & Criteria1 & "%'"
    End If
    
    If Criteria2 <> "" And Criteria1 <> "" Then
        QryString = QryString & " OR [CÓDIGO] LIKE '%" & Criteria2 & "%'"
    ElseIf Criteria2 <> "" And Criteria1 = "" Then
        QryString = QryString & " WHERE [CÓDIGO] LIKE '%" & Criteria2 & "%'"
    End If
    
    getSupplierListForDisplay = GetDataArrayFromDB(query:=QryString, WithHeader:=True)

End Function

Function getCustomerListForDisplay(ByVal sTableName As String, ByVal Criteria1 As String, Optional ByVal Criteria2 As String = "")
    
    Dim TableQueryRange As String
    
    TableQueryRange = getAddress(sTableName)
    
    Dim QryString As String
    
    QryString = "SELECT * FROM " & TableQueryRange
                  
    If Criteria1 <> "" Then
        QryString = QryString & " WHERE [NOME] LIKE '%" & Criteria1 & "%' OR [NOME REDUZIDO] LIKE '%" & Criteria1 & "%'"
    End If
    
    If Criteria2 <> "" And Criteria1 <> "" Then
        QryString = QryString & " OR [ID CLIENTE] LIKE '%" & Criteria2 & "%'"
    ElseIf Criteria2 <> "" And Criteria1 = "" Then
        QryString = QryString & " WHERE [ID CLIENTE] LIKE '%" & Criteria2 & "%'"
    End If
    
    getCustomerListForDisplay = GetDataArrayFromDB(query:=QryString, WithHeader:=True)

End Function


























Function getListForDisplay(ByVal sTableName As String, ByVal TableQueryColumns As Variant, ByVal Criteria1 As String, ByVal TableFields As Variant)
    
    Dim TableQueryRange As String
    
    TableQueryRange = getAddress(sTableName)
    
    Dim sqlQuery As String
    
    sqlQuery = "SELECT * FROM " & TableQueryRange
                  
    If Criteria1 <> "" Then
        sqlQuery = sqlQuery & " WHERE [NOME EMPRESARIAL] LIKE '%" & Criteria1 & "%' OR [NOME FANTASIA] LIKE '%" & Criteria1 & "%'"
    End If
    
    
    getListForDisplay = GetDataArrayFromDB(sqlQuery, WithHeader:=True)

End Function























