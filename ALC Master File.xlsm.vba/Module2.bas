Sub Open_HOS()
Application.ScreenUpdating = False
ActiveWorkbook.Unprotect Password:="contract"
Sheets("HOS_Modellable Customers").Visible = xlSheetVisible
' or xlSheetHidden or xlSheetVisible or xlSheetVeryHidden
Sheets("Modellable Customers").Visible = xlSheetVeryHidden
Range("D8").Select
ActiveCell.ClearContents

ActiveWorkbook.Protect Password:="contract", Structure:=True, Windows:=True
Application.ScreenUpdating = True
End Sub

Sub Open_MODELLABLE()
Application.ScreenUpdating = False
ActiveWorkbook.Unprotect Password:="contract"
Sheets("HOS_Modellable Customers").Visible = xlSheetVeryHidden
' or xlSheetHidden or xlSheetVisible or xlSheetVeryHidden
Sheets("Modellable Customers").Visible = xlSheetVisible
ActiveWorkbook.Protect Password:="contract", Structure:=True, Windows:=True
Application.ScreenUpdating = True
End Sub


Sub Clear_TERMS()
Application.ScreenUpdating = False
ActiveWorkbook.Unprotect Password:="contract"

Range("D12").Select
ActiveCell.FormulaR1C1 = "Select Term"
ActiveWorkbook.Protect Password:="contract", Structure:=True, Windows:=True
Application.ScreenUpdating = True
End Sub