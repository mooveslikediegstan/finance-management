Attribute VB_Name = "M_BancoDeDados"
Option Explicit

Public tbPurchaseOrder As ListObject
Public tbPurchaseOrderPlanning As ListObject
Public tbInvoice As ListObject
Public tbCustomer As ListObject
Public tbProject As ListObject
Public tbSupplier As ListObject

Public TechNamePurchaseOrder As ListObject
Public TechNamePurchaseOrderPlanning As ListObject
Public TechNameInvoice As ListObject
Public TechNameCustomer As ListObject
Public TechNameProject As ListObject
Public TechNameSupplier As ListObject

Public Sub InstanceTables()

    Set tbPurchaseOrder = GetTableObject(WSPURCHASEORDER, "PurchaseOrderTable")
    Set tbPurchaseOrderPlanning = GetTableObject(WSPURCHASEORDERPLANNING, "PurchaseOrderPlanningTable")
    Set tbInvoice = GetTableObject(WSINVOICE, "InvoiceTable")
    Set tbCustomer = GetTableObject(WSCUSTOMER, "CustomerTable")
    Set tbProject = GetTableObject(WSPROJECT, "ProjectTable")
    Set tbSupplier = GetTableObject(WSSUPPLIER, "SupplierTable")


    Set TechNamePurchaseOrder = GetTableObject(TechnicalNames, "TechNamePurchaseOrder")
    Set TechNamePurchaseOrderPlanning = GetTableObject(TechnicalNames, "TechNamePurchaseOrderPlanning")
    Set TechNameInvoice = GetTableObject(TechnicalNames, "TechNameInvoice")
    Set TechNameCustomer = GetTableObject(TechnicalNames, "TechNameCustomer")
    Set TechNameProject = GetTableObject(TechnicalNames, "TechNameProject")
    Set TechNameSupplier = GetTableObject(TechnicalNames, "TechNameSupplier")
    
End Sub

Public Sub RefreshInstancingInstanceTables()

    If tbPurchaseOrder Is Nothing Then Set tbPurchaseOrder = GetTableObject(WSPURCHASEORDER, "PurchaseOrderTable")
    If tbPurchaseOrderPlanning Is Nothing Then Set tbPurchaseOrderPlanning = GetTableObject(WSPURCHASEORDERPLANNING, "PurchaseOrderPlanningTable")
    If tbInvoice Is Nothing Then Set tbInvoice = GetTableObject(WSINVOICE, "InvoiceTable")
    If tbCustomer Is Nothing Then Set tbCustomer = GetTableObject(WSCUSTOMER, "CustomerTable")
    If tbProject Is Nothing Then Set tbProject = GetTableObject(WSPROJECT, "ProjectTable")
    If tbSupplier Is Nothing Then Set tbSupplier = GetTableObject(WSSUPPLIER, "SupplierTable")


    If TechNamePurchaseOrder Is Nothing Then Set TechNamePurchaseOrder = GetTableObject(TechnicalNames, "TechNamePurchaseOrder")
    If TechNamePurchaseOrderPlanning Is Nothing Then Set TechNamePurchaseOrderPlanning = GetTableObject(TechnicalNames, "TechNamePurchaseOrderPlanning")
    If TechNameInvoice Is Nothing Then Set TechNameInvoice = GetTableObject(TechnicalNames, "TechNameInvoice")
    If TechNameCustomer Is Nothing Then Set TechNameCustomer = GetTableObject(TechnicalNames, "TechNameCustomer")
    If TechNameProject Is Nothing Then Set TechNameProject = GetTableObject(TechnicalNames, "TechNameProject")
    If TechNameSupplier Is Nothing Then Set TechNameSupplier = GetTableObject(TechnicalNames, "TechNameSupplier")
    
End Sub

Public Sub KillTables()

    Set tbPurchaseOrder = Nothing
    Set tbPurchaseOrderPlanning = Nothing
    Set tbInvoice = Nothing
    Set tbCustomer = Nothing
    Set tbProject = Nothing
    Set tbSupplier = Nothing


    Set TechNamePurchaseOrder = Nothing
    Set TechNamePurchaseOrderPlanning = Nothing
    Set TechNameInvoice = Nothing
    Set TechNameCustomer = Nothing
    Set TechNameProject = Nothing
    Set TechNameSupplier = Nothing
    
End Sub
