Private Sub Worksheet_Change(ByVal Target As Range)
If ActiveWorkbook.Name = "ALC Master File1.xlsm" Then GoTo SkipAll Else
If ActiveCell.Value = "Select Term" Then Exit Sub Else
ActiveSheet.Unprotect Password:="contract"
Application.ScreenUpdating = False
If ActiveCell.Value = "Incentive Only" Then
    Call Change1
Else
    Call Change2
End If

If ActiveCell.Value = "PON Only" Then
ThisWorkbook.Unprotect Password:="contract"
Sheets("Incentives").Visible = False
ThisWorkbook.Protect Password:="contract"
Else
ThisWorkbook.Unprotect Password:="contract"
Sheets("Incentives").Visible = True
ThisWorkbook.Protect Password:="contract"
End If

If ActiveCell.Value = "HOSPITALITY" Then
    Call Open_HOS
Else
End If

If ActiveCell.Value = "INDEPENDENT RESTAURATEURS" Then
    Call Open_MODELLABLE
Else
End If

If ActiveCell.Value = "MDA" Then
ThisWorkbook.Unprotect Password:="contract"
Rows("10:11").EntireRow.Hidden = False
ThisWorkbook.Protect Password:="contract"
Else
End If

If ActiveCell.Offset(0, -1) = "Is this a MDA Renewal?" And ActiveCell.Value = "No" Then
If ActiveCell.Value = "Select Term" Then
ActiveSheet.Protect Password:="contract"
Exit Sub
Else
Call Clear_TERMS
End If
End If

If ActiveCell.Value = "PON Only" Then
If Range("D12").Value = "Select Term" Then
ActiveSheet.Protect Password:="contract"
Exit Sub
Else
Call Clear_TERMS
End If
End If

If Range("D8").Value <> "MDA" Then
ThisWorkbook.Unprotect Password:="contract"
On Error Resume Next
Rows("10:11").EntireRow.Hidden = True
ThisWorkbook.Protect Password:="contract"
End If

ActiveSheet.Protect Password:="contract"
SkipAll:
End Sub


