Attribute VB_Name = "Y_SearchFunctions"
Option Explicit

Function getCustomer() As Customer

    Dim searchForm As CustomerSearch
    
    Set searchForm = New CustomerSearch
    
    searchForm.Show
    
    If Not searchForm.Cancelled Then
        Set getCustomer = searchForm.SelectedCustomer
    Else
        Set getCustomer = searchForm.SelectedCustomer ' Usuário cancelou
    End If
    
    Unload searchForm
    
End Function

Function getSupplier() As Supplier

    Dim searchForm As SupplierSearch
    
    Set searchForm = New SupplierSearch
    
    searchForm.Show
    
    If Not searchForm.Cancelled Then
        Set getSupplier = searchForm.SelectedSupplier
    Else
        Set getSupplier = searchForm.SelectedSupplier ' Usuário cancelou
    End If
    
    Unload searchForm

End Function
