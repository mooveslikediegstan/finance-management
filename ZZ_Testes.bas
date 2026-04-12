Attribute VB_Name = "ZZ_Testes"
Sub TMAIN()
    
    Set DB = Conectar
    
    TABLEADDRES = getAddress("SupplierTable")
    
    SQLstr = "SELECT * FROM " & TABLEADDRES
    
    Dim rs As New ADODB.Recordset
    
    Set rs = ExecQuery(SQLstr)
    
    a = accessDB_to_array(rs)
    
    Dim TESTE As PurchaseOrder
    
    Set TESTE = New PurchaseOrder
    
    TESTE.LoadFromDatabaseToObject ID:="991430"

End Sub

Sub testwbpath()
Dim fileDialog As fileDialog
Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
fileDialog.Show
    Debug.Print ThisWorkbook.Path
Debug.Print ThisWorkbook.FullName
Debug.Print ThisWorkbook.FullNameURLEncoded
End Sub
