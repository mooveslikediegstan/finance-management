Attribute VB_Name = "Z_ListObjectFunctions"
Option Explicit


Sub UpdateTableField(ByVal Table As ListObject, ByVal RowID As Long, ByVal ColumnID As Long, ByVal Value As Variant)

    Table.DataBodyRange(RowID, ColumnID) = Value

End Sub

Function GetColumnID(ByVal ColumnName As String, ByVal Dict As Dictionary) As Long
    
    GetColumnID = Dict(ColumnName)

End Function

Function SearchRowID(ByVal Identifier As Variant, ByVal Table As ListObject, ByVal ColumnID As Long) As Long 'get the FIRST row ID
    
    Identifier = Trim(Identifier)
    
    Dim arrTable As Variant: arrTable = Table.DataBodyRange
    
    Dim RowCounter As Long
    
    For RowCounter = LBound(arrTable) To UBound(arrTable)
        If arrTable(RowCounter, ColumnID) = Identifier Then
           SearchRowID = RowCounter
           Exit Function
        End If
    Next RowCounter
    
    SearchRowID = -1 'did not find any match
    
End Function

Function GetHeaderIndexList(ByVal Table As ListObject) As Dictionary

    Dim Dict As New Dictionary
    
    Dim HeaderColumn As Object
    
    For Each HeaderColumn In Table.ListColumns
    
        Dict.Add HeaderColumn.Name, HeaderColumn.index
    
    Next
    
    Set GetHeaderIndexList = Dict

End Function


Function GetTableObject(ByVal SheeetName As Worksheet, ByVal TableName As String) As ListObject
     
   Set GetTableObject = SheeetName.ListObjects(TableName)
   
End Function

Function GetTableDataBodyRange(ByVal Table As ListObject) As Range

    Set GetTableDataBodyRange = Table.DataBodyRange

End Function

Function GetArrayFromTable(ByVal Table As ListObject) As Variant
       
    GetArrayFromTable = Table.DataBodyRange
    
End Function


