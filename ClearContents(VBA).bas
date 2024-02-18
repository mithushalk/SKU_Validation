Attribute VB_Name = "ClearContents"
Sub ClearContents()

ThisWorkbook.Activate

Application.DisplayAlerts = False

'Delete all sheets except DataInput_Instructions & hidden CID_SUB IDs

On Error Resume Next
For Each ws In ThisWorkbook.Sheets
If ws.Name <> "DataInput_Instructions" And ws.Name <> "CID_SUBIDs" Then
Application.DisplayAlerts = False
ws.Delete
End If
Next ws
On Error GoTo 0

'Clear B3 cell

Sheets("DataInput_Instructions").Select
Range("B3").ClearContents

'Popup message

MsgBox "Data Cleared"

End Sub

Sub FormatData()

'Delete Aditional Columns from the File

ThisWorkbook.Activate

Workbooks.Open Filename:=ThisWorkbook.Path & "\" & "Rawfile"

Sheets("Page1_1").Select
Range("E:F,H:I,M:M").Delete Shift:=xlToLeft

End Sub
