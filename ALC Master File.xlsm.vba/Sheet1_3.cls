Private Sub Worksheet_Change(ByVal Target As Range)

    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False


ActiveSheet.Unprotect Password:="contract"
    Range("D10:D30").Locked = False


'Statement that checks whether EXCLUDE is toggled anywhere
If Not Intersect(Target, Target.Worksheet.Range("E10:E30")) Is Nothing Then Exclude

'Clear Margin if Pricing Standard is toggled
Dim m As Long

m = ActiveCell.Row

If Target.Address = Range("D" & m).Address Then
    Range("K" & m).Value = ""
    
    
End If

'After macro re-lock column D depending on MDA/Incentive only

    If Range("E3") = "Incentive Only" Then
        If ActiveSheet.Range("D10:D30").Locked = False Then
            ActiveSheet.Range("D10:D30").Locked = True
        End If
    ActiveSheet.Protect Password:="contract"
    Else
    ActiveSheet.Unprotect Password:="contract"
    Worksheets("PIM Model").Range("D10,D13,D14,D15,D16,D17,D18,D19,D20,D21,D23,D28").Locked = True
    Worksheets("PIM Model").Range("D11,D12,D22,D24,D25,D26,D27,D29,D30").Locked = False
    ActiveSheet.Protect Password:="contract"
    End If

ActiveSheet.Protect Password:="contract"

    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True


End Sub

Sub Exclude()

If ActiveSheet.Unprotect = False Then
    ActiveSheet.Unprotect Password:="contract"
    End If
    

'Clear rows when Exclude is toggled

If ActiveCell.Value = "EXCLUDE" Then
    ActiveCell.Offset(0, -3).Value = 0
    ActiveCell.Offset(0, -2).Value = ""
    If ActiveCell.Offset(0, -1).Locked = True Then
        ActiveSheet.Unprotect Password:="contract"
        ActiveCell.Offset(0, -1).Locked = False
    End If
    ActiveCell.Offset(0, -1).Value = "%"
    ActiveCell.Offset(0, 6).Value = ""
   
End If

End Sub

