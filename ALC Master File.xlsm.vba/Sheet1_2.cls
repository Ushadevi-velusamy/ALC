Private Sub Worksheet_Activate()

Dim i As Long
Dim FirstRow As Long

Application.ScreenUpdating = False

'Find my first row
If Worksheets("Welcome").Range("D14").Value = "Select # of Locations" Then
FirstRow = 6
Else: FirstRow = Worksheets("Welcome").Range("D14").Value + 6
End If

'Loop through range

For i = FirstRow To 15

Range("G" & i).ClearContents
Range("M" & i & ":" & "Q" & i).ClearContents

Next i

Application.ScreenUpdating = True


End Sub