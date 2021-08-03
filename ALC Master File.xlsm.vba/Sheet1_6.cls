'Private Sub Worksheet_Activate()
'    Dim ToRow As Integer
'    ToRow = Range("R2") + 17
'    If Range("R2") = 0 Then GoTo SkipHere Else
'    Range("F18").Select
'    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'Customer Info'!C[-5]:C[-4],2,0),"""")"
'    Range("G18").Select
'    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],'Customer Info'!C[-6]:C[-4],3,0),"""")"
'    Range("H18").Select
'    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],'Customer Info'!C[-7]:C[-4],4,0),"""")"
'    Range("I18").Select
'    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],'Customer Info'!C[-8]:C[-4],5,0),"""")"
'    Range("J18").Select
'    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],'Customer Info'!C[-9]:C[-4],6,0),"""")"
'    Range("F18:J18").Select
'    Selection.Copy
'    Range("F18:J" & ToRow).Select
'    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
'    Range("D15").Select
'SkipHere:
'
'End Sub