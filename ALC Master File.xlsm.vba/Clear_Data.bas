Sub ClearCustomerInputs()
Range("G6:G15").ClearContents
Range("M6:Q15").ClearContents
End Sub

Sub Incentives()

Range("G18:G20").ClearContents
Range("I14").Value = "Select Payment Method"
Range("I18").Value = 0.01
Range("I19").Value = 0.015
Range("I20").Value = 0.02
Range("E15").Value = "NO"

Range("G29:G31").ClearContents
Range("I25").Value = "Select Payment Method"
Range("I29").Value = 0.01
Range("I30").Value = 0.015
Range("I31").Value = 0.02
Range("E26").Value = "NO"

Range("I36").Value = "Select Payment Method"
Range("G40:G42").ClearContents
Range("I40").Value = 0.01
Range("I41").Value = 0.015
Range("I42").Value = 0.02
Range("E37").Value = "NO"

Range("I47").Value = "Select Payment Method"
Range("k50").ClearContents
Range("g50").Value = 0
Range("E48").Value = "NO"

Range("I55").Value = "Select Payment Method"
Range("k58").ClearContents
Range("E56").Value = "NO"

Range("I63").Value = "Select Payment Method"
Range("k66").ClearContents
Range("E64").Value = "NO"

Range("I71").Value = "Select Payment Method"
Range("E72").Value = "NO"



End Sub

Sub RequireInfoClear()

Range("F4:G4").ClearContents
Range("F6:G6").ClearContents
Range("F8:G12").ClearContents
Range("F14:G14").ClearContents

End Sub

Sub ClearPimModel()

ActiveSheet.Unprotect Password:="contract"
ActiveSheet.Range("d10:d30").Locked = False

If ActiveSheet.Range("E3").Value = "Incentive Only" Then
    Range("b10:b30").Value = 0
    Range("c10:c30").Value = ""
    Range("e10:e30").Value = "INCLUDE"
Else
    Range("b10:b30").Value = 0
    Range("c10:c30").Value = ""
    Range("d11:d12").Value = "%"
    Range("d22").Value = "%"
    Range("d24:d27").Value = "%"
    Range("d29:d30").Value = "%"
    Range("e10:e30").Value = "INCLUDE"
    Range("k10:k30").ClearContents
End If


ActiveSheet.Protect Password:="contract"

End Sub

Sub ClearWelcome()
Range("d6").Value = "INDEPENDENT RESTAURATEURS"
Range("d8").Value = "Select Deal Type"
Range("d10").Value = "Select Term"
Range("d12").Value = "Select # of Locations"
End Sub

Sub Change1()
Application.ScreenUpdating = False

    Sheets("PIM Model").Activate
    ActiveSheet.Unprotect Password:="contract"
    Range("D10:D30").Locked = False
    Range("D10:D30").Value = "%"
    Sheets("Welcome").Activate
    
End Sub

Sub Change2()
Application.ScreenUpdating = False

    Sheets("PIM Model").Activate
    ActiveSheet.Unprotect Password:="contract"
    ActiveSheet.Range("d10:d30").Locked = False
    ActiveSheet.Range("D10,D13,D14,D15,D16,D17,D18,D19,D20,D21,D23,D28").Locked = True
    ActiveSheet.Protect Password:="contract"
    Sheets("Welcome").Activate
    
End Sub

Sub ClearTracking()
Range("E16").Value = 0
End Sub