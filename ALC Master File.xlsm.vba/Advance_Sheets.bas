Sub NextActiveSheet()
On Error Resume Next
Sheets(ActiveSheet.Index + 1).Activate
If Err.Number <> 0 Then Sheets(1).Activate
End Sub

Sub PrevActiveSheet()
On Error Resume Next
Sheets(ActiveSheet.Index - 1).Activate
If Err.Number <> 0 Then Sheets(Worksheets.Count).Activate
End Sub

Sub Submit()

Dim xWb As Workbook
Dim xStr As String
Dim xStrOldName As String
Dim xStrDate As String
Dim xFileName As String
Dim xCust As String
Dim xCust2 As String
Dim xFileDlg As FileDialog
Dim i As Variant
Dim booWorkbookSaved As Boolean
Application.DisplayAlerts = False

Set xWb = ActiveWorkbook
xStrOldName = xWb.Name
xStr = Left(xStrOldName, Len(xStrOldName) - 5)
xStrDate = Format(Now, "YYMMDD")
xCust = ActiveSheet.Range("F8")
xCust2 = Replace(xCust, ".", "")

xFileName = Application.GetSaveAsFilename(xStrDate & "-" & xCust2 & "-" & xStr, "Excel Macro-Enabled Workbook (*.xlsm),*.xlsm")
  
If xFileName = "False" Then
    Exit Sub
End If

xWb.SaveAs (xFileName)

'If Sheets("Incentives").Range("K66") >= 10000 Then
'Call Mail_over10K
'Else
Call Mail
'End If

Call sbMsgBox



Application.DisplayAlerts = True

End Sub


Sub Mail()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim mkt As String
    Dim Cust As String
    Dim Contact As String
    Dim xStrDate As String
    
    mkt = Worksheets("Welcome").Range("D4")
    Cust = Worksheets("Required Info to Submit").Range("F8")
    Contact = Worksheets("Required Info to Submit").Range("F4")
    xStrDate = Format(Now, "YYMMDD")
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

   strbody = "Hi USF Contract Team - Please see the attached A La Carte Model for:" & vbNewLine & vbNewLine & _
              "  Customer -  " & Cust & vbNewLine & _
              "  Market -  " & mkt & vbNewLine & _
              "  Date -  " & xStrDate & vbNewLine & _
              "  Market Contact -  " & Contact & vbNewLine & _
              vbNewLine & _
              "Thank you," & vbNewLine & _
              mkt & " Local Sales Team"

    On Error Resume Next
    With OutMail
        .To = "alacarte.shared@usfoods.com"
        .CC = ""
        .BCC = ""
        .Subject = "New " & mkt & " ALC Model For " & Cust
        .Body = strbody
        .Attachments.Add ActiveWorkbook.FullName
        
        .send   'or use .Display
    End With
    On Error GoTo 0
    

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub


Sub Mail_over10K()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim mkt As String
    Dim Cust As String
    Dim Contact As String
    Dim xStrDate As String
    
    mkt = Worksheets("Welcome").Range("D4")
    Cust = Worksheets("Required Info to Submit").Range("F8")
    Contact = Worksheets("Required Info to Submit").Range("F4")
    xStrDate = Format(Now, "YYMMDD")
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

   strbody = "Hi USF Contract Team - Please see the attached A La Carte Model for:" & vbNewLine & vbNewLine & _
              "  Customer -  " & Cust & vbNewLine & _
              "  Market -  " & mkt & vbNewLine & _
              "  Date -  " & xStrDate & vbNewLine & _
              "  Market Contact -  " & Contact & vbNewLine & _
              vbNewLine & _
              "Thank you," & vbNewLine & _
              mkt & " Local Sales Team"

    On Error Resume Next
    With OutMail
        .To = "profitmodelrequest.shared@usfoods.com"
        .CC = ""
        .BCC = ""
        .Subject = "New " & mkt & " ALC Model For " & Cust
        .Body = strbody
        .Attachments.Add ActiveWorkbook.FullName
        
        .display   'or use .Display
    End With
    On Error GoTo 0
    

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

Sub sbMsgBox()

MsgBox "Your A La Carte Model has been created and emailed to the USF Contract Team for review. If you have any questions or concerns please reach out to ALaCarte.Shared@usfoods.com. Thank you.", vbExclamation, "A La Carte Model Published"

End Sub