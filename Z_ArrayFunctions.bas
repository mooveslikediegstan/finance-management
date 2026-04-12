Attribute VB_Name = "Z_ArrayFunctions"
'@Folder("FinancialManagement")
Option Explicit


Public Function fApenasNumerosKeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Permite apenas números (0-9)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Function

Public Function fApenasNumerosEPontoKeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Permite apenas números (0-9) e ponto (.)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 44) Then
        KeyAscii = 0
    End If
End Function

Public Function IsInArray(valueToFind As Variant, arr As Variant) As Boolean

    Dim element As Variant
    
    For Each element In arr
    
        If element = valueToFind Then
            IsInArray = True
            Exit Function
        End If
        
    Next element
    
    IsInArray = False
    
End Function

Function IsArrayEmpty2(arr As Variant) As Boolean
    On Error Resume Next
    If IsEmpty(arr) Then IsArrayEmpty2 = True
    If UBound(arr) = -1 Then IsArrayEmpty2 = True
    On Error GoTo 0
End Function

Public Function arrayTranspose(ByVal arr As Variant, Optional ByVal sizeCorrection As Integer = 0) As Variant

    Dim temp_array As Variant
    
    ReDim temp_array(0 To 0, 0 To 0)
    
    Dim nCols As Long, nRows As Long
    nCols = UBound(arr)
    nRows = UBound(arr, 2) - sizeCorrection
        
    If nRows = 0 And arr(0, 0) = "" Then
        temp_array = Empty
        GoTo errhd
    End If
        
    ReDim temp_array(0 To nRows, 0 To nCols)
    
    Dim i As Long, j As Long
    For i = 0 To nRows
        For j = 0 To nCols
            temp_array(i, j) = arr(j, i)
        Next j
    Next i
errhd:

    arrayTranspose = temp_array
    If Not IsEmpty(temp_array) Then Erase temp_array
    Erase arr
    
End Function

Public Function arrayFilterBinary(ByRef dataArray, searchColumn As Long, searchValue As Variant)

    On Error GoTo errorHandler
    
    Dim firstMatch As Long
    Dim lastMatch As Long
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim resultArray() As Variant
    Dim matchCount As Long
    
    ' Find the first occurrence using binary search
    firstMatch = BinarySearchFirst(dataArray, searchColumn, searchValue, LBound(dataArray, 1), UBound(dataArray, 1))

    ' If not found, return empty array
    If firstMatch = -1 Then
        arrayFilterBinary = arrayFilterLinear(dataArray, searchColumn, searchValue)
        Exit Function
    End If

    ' Find the last occurrence by scanning linearly from first match
    lastMatch = firstMatch
    For i = firstMatch + 1 To UBound(dataArray, 1)
        If dataArray(i, searchColumn) = searchValue Then
            lastMatch = i
        Else
            Exit For                             ' Stop when value changes (thanks to sorting!)
        End If
    Next i
    
    ' Calculate how many rows match
    matchCount = lastMatch - firstMatch + 1
    
    ' Create result array with same number of columns
    ReDim resultArray(0 To matchCount - 1, 0 To UBound(dataArray, 2))
    
    ' Copy matching rows to result array
    For i = 0 To matchCount - 1
        For j = 0 To UBound(dataArray, 2)
            resultArray(i, j) = dataArray(firstMatch + i, j)
        Next j
    Next i
    
    arrayFilterBinary = resultArray
    Exit Function

errorHandler:
    ' Return empty array on error
    arrayFilterBinary = Array()
    
End Function

Function BinarySearchFirst(ByRef dataArray As Variant, ByRef searchColumn As Long, ByRef searchValue As Variant, low As Long, high As Long) As Long
    
    ' Binary search to find FIRST occurrence of searchValue
    Dim mid As Long
    Dim result As Long
    
    result = -1                                  ' Default: not found
    
    While low <= high
        mid = low + (high - low) \ 2
        
        If dataArray(mid, searchColumn) = searchValue Then
            result = mid                         ' Found a match, but need to find the FIRST one
            ' Continue searching left half to find first occurrence
            high = mid - 1
        ElseIf dataArray(mid, searchColumn) < searchValue Then
            low = mid + 1
        Else
            high = mid - 1
        End If
    Wend

    BinarySearchFirst = result

End Function



Function arrayFilterLinear(ByRef dataArray As Variant, searchColumn As Long, searchValue As Variant) As Variant
    ' Efficient linear scan for SORTED arrays
    ' Stops when it starts finding matches, stops when matches stop
    
    Dim i As Long, j As Long, k As Long
    Dim inMatchZone As Boolean
    Dim matchCount As Long
    Dim resultArray() As Variant
    Dim firstMatch As Long
    
    matchCount = 0
    inMatchZone = False
    firstMatch = -1
    
    ' Single pass through array
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        If dataArray(i, searchColumn) = searchValue Then
            If Not inMatchZone Then
                ' Found first match
                inMatchZone = True
                firstMatch = i
            End If
            matchCount = matchCount + 1
        ElseIf inMatchZone Then
            ' We were in match zone but value changed - we're done!
            Exit For
        End If
    Next i
    
    ' If no matches found
    If matchCount = 0 Or firstMatch = -1 Then
        arrayFilterLinear = Array()
        Exit Function
    End If
    
    ' Create result array
    ReDim resultArray(0 To matchCount - 1, 0 To UBound(dataArray, 2))
    
    ' Copy matching rows
    For i = 0 To matchCount - 1
        For j = 0 To UBound(dataArray, 2)
            resultArray(i, j) = dataArray(firstMatch + i, j)
        Next j
    Next i
    
    arrayFilterLinear = resultArray
End Function

Function AppendArrays(ByVal Array1 As Variant, ByVal Array2 As Variant)
    
    Dim ArrayRowSize As Integer: ArrayRowSize = (GetArrayRowSize(Array1) + GetArrayRowSize(Array2))
    Dim ArrayColSize As Integer: ArrayColSize = WorksheetFunction.Max(GetArrayColumnSize(Array1), GetArrayColumnSize(Array2))
    
    Dim resultArray()
    ReDim resultArray(0 To ArrayRowSize - 1, 0 To ArrayColSize - 1)
    
    Dim Array1RowCounter As Long: Array1RowCounter = 0
    Dim Array2RowCounter As Long: Array2RowCounter = 0
    
    Dim i As Long, j As Long
    For i = 0 To ArrayRowSize - 1
    
        If i < GetArrayRowSize(Array1) Then
        
            For j = 0 To ArrayColSize - 1
                resultArray(i, j) = Array1(Array1RowCounter, j)
            Next j
            
            Array1RowCounter = Array1RowCounter + 1
            
        Else:
        
            For j = 0 To ArrayColSize - 1
                resultArray(i, j) = Array2(Array2RowCounter, j)
            Next j
            
            Array2RowCounter = Array2RowCounter + 1
        
        End If
        
    Next i

    AppendArrays = resultArray

End Function

Function IsArrayEmpty(arr As Variant) As Boolean

    ' Purpose: Check if an array is empty (uninitialized or zero-length)
    ' Returns: True if empty, False otherwise
    Dim lb As Long, ub As Long
    
    ' First, check if the input is even an array
    If Not IsArray(arr) Then
        IsArrayEmpty = True ' Not an array ? treat as "empty" (optional behavior)
        Exit Function
    End If
    
    ' Handle uninitialized arrays (error when accessing LBound)
    On Error Resume Next
    lb = LBound(arr) ' Get lower bound
    ub = UBound(arr) ' Get upper bound
    On Error GoTo 0 ' Reset error handling
    
    ' If LBound/UBound failed (error), array is uninitialized (empty)
    ' If UBound < LBound, array is zero-length (empty)
    IsArrayEmpty = (lb > ub) Or (Err.Number <> 0)
    
End Function


Function GetArrayRowSize(ByVal fArray As Variant)

    GetArrayRowSize = UBound(fArray) - LBound(fArray) + 1

End Function

Function GetArrayColumnSize(ByVal fArray As Variant)
    
    
    If IsArrayEmpty(fArray) Then GoTo errHandlr
    
    GetArrayColumnSize = UBound(fArray, 2) - LBound(fArray, 2) + 1
    Exit Function
    
    
errHandlr:
    GetArrayColumnSize = 0
    
End Function



