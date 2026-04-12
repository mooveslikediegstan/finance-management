Attribute VB_Name = "B_"
Option Explicit

Sub B_GoTo_MasterData()

   ThisWorkbook.Sheets("MasterData").Visible = -1
   Application.Goto ThisWorkbook.Sheets("MasterData").Range("D1")
   Unload formControle
   
End Sub

Sub B_Voltar_MasterData()

   ThisWorkbook.Sheets("MasterData").Visible = 2
   Application.Goto ThisWorkbook.Sheets("Controle").Range("A1")
   formControle.Show
   
End Sub


