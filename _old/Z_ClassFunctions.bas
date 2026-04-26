Attribute VB_Name = "Z_ClassFunctions"
Option Explicit

Public Sub LoadClassColumnsIndexes(TargetObject As Variant, ColumnsNameList As Variant, DatabaseTable As ListObject, TechNameTable As ListObject)
   
    Dim RecordTableHeader As Dictionary
    Set RecordTableHeader = GetHeaderIndexList(DatabaseTable)
    
    Dim TechNameHeader As Dictionary
    Set TechNameHeader = GetHeaderIndexList(TechNameTable)
    


    Dim ColumnName As String
    Dim ClassPropertyName As String

    Dim TechNameColumnID As Long
    
    Dim RecordTableColumnID As Long
    
    Dim i As Long
    For i = 0 To UBound(ColumnsNameList)
        
        ColumnName = ColumnsNameList(i)
        
        ClassPropertyName = Replace(ColumnName, "cl", "", 1, 2)
        
        TechNameColumnID = GetColumnID(ClassPropertyName, TechNameHeader)
        
        RecordTableColumnID = GetColumnID(TechNameTable.DataBodyRange(1, TechNameColumnID), RecordTableHeader)
        
        CallByName TargetObject, ColumnName, VbLet, RecordTableColumnID
        
    Next i

End Sub
