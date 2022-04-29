Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine

Set objExcel = CreateObject("Excel.Application")
filepath =  sFileSelected
objExcel.DisplayAlerts = True
On Error Resume Next
objExcel.Application.Run "'" & sFileSelected & "'!Module3.CreateUploadFile"
If objExcel.Workbooks.Count > 0 Then
	OutPut = MsgBox("Upload File Created Successfully", vbInformation, "Upload File")	
End If
objExcel.Quit
Set objExcel = Nothing
Set book = Nothing
Set wShell = Nothing 
Set oExec = Nothing 
	