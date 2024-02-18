Attribute VB_Name = "SKUMacro"
Sub Pending()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Check if user has selected the Advertiser

ThisWorkbook.Activate
Sheets("DataInput_Instructions").Select
If Range("B3").Value = "" Then
MsgBox "WARNING!  Choose Advertiser Name to proceed"
Exit Sub
End If

'Copy Data from Rawfile to ThisWorkbooks

ThisWorkbook.Activate

Workbooks.Open Filename:=ThisWorkbook.Path & "\" & "Rawfile"

Sheets("Page1_1").Select
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

'Create Raw Sheet

ThisWorkbook.Activate
Sheets.Add Before:=Sheets("DataInput_Instructions")
ActiveSheet.Name = "Raw"

Sheets("Raw").Select
Range("A1").PasteSpecial xlPasteValues
Columns("K:L").NumberFormat = "dd-mm-yyyy"

'Add Difference_Sale_Amount & Diff% columns

lr1 = 0
lr1 = Sheets("Raw").Cells(Rows.Count, 1).End(xlUp).Row

Sheets("Raw").Select
Range("T1").Value = "Difference_Sale_Amount"
Range("T2").Value = "=M2-Q2"
Range("U1").Value = "Diff%"
Range("U2").Value = "=(M2/Q2)-1"
Range("T2:U" & lr1).Select
Selection.FillDown
Columns("U:U").Style = "Percent"

Workbooks("Rawfile").Close SaveChanges:=False
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Create Working_Sheet

Sheets.Add Before:=Sheets("Raw")
ActiveSheet.Name = "Working_Sheet"

Sheets("Raw").Select

'Filter Data based on the Advertiser name

Range("A1:U1").AutoFilter Field:=2, Criteria1:=ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value

lr1 = 0
lr1 = Sheets("Raw").Cells(Rows.Count, 1).End(xlUp).Row
ThisWorkbook.Sheets("Raw").Range("XFD1").Value = Application.WorksheetFunction.CountIf(ThisWorkbook.Sheets("Raw").Range("B2:B" & lr1), ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value)

If Sheets("Raw").Range("XFD1").Value > 0 Then
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("Working_Sheet").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

Columns("K:L").NumberFormat = "dd-mm-yyyy"
Columns("H:I").NumberFormat = "0"
Columns("A:Z").EntireColumn.AutoFit
Else
MsgBox "No Advertiser of your choice"
Exit Sub
End If

Sheets("Raw").Select
Range("A1:U1").AutoFilter

'Pending

Sheets("Working_Sheet").Activate
Range("A1:U1").AutoFilter Field:=18, Criteria1:="Pending", Operator:=xlOr, Criteria2:="PP  "

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("R2:R" & lr1), "Pending") + Application.WorksheetFunction.CountIf(Range("R2:R" & lr1), "PP  ")

Sheets.Add Before:=Sheets("Raw")
ActiveSheet.Name = "Pending"

If Sheets("Working_Sheet").Range("XFD1").Text > 0 Then
Sheets("Working_Sheet").Select
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("Pending").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

'Delete Pending from Working Sheet

Sheets("Working_Sheet").Select
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).EntireRow.Delete

Sheets("Working_Sheet").Select
Range("A1:U1").AutoFilter

'Apply XLOOKUP formula to check OIDs are mapped by different Status

Sheets("Working_Sheet").Select
Range("V2").Value = "=XLOOKUP(H2,Pending!H:H,Pending!O:O,""NA"")"
Range("Q2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 5).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown

Sheets("Working_Sheet").Select
Range("A1:U1").AutoFilter

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("V2:V" & lr1), "<>NA")

If Sheets("Working_Sheet").Range("XFD1").Text > 0 Then
Range("A1:V1").AutoFilter Field:=22, Criteria1:="<>NA"
Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Pending").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Sheets("Working_Sheet").Select
Range("A1:V1").AutoFilter
End If
End If

Sheets("Working_Sheet").Select
Range("A1:U1").AutoFilter
Range("V:V").ClearContents

Application.DisplayAlerts = False
Application.ScreenUpdating = False

ThisWorkbook.Activate

Call Cancelled

End Sub

Sub Cancelled()

Sheets("Raw").Select
Sheets.Add Before:=Sheets("Raw")
ActiveSheet.Name = "Cancelled"

Sheets("Working_Sheet").Activate
Range("A1:U1").AutoFilter Field:=18, Criteria1:="Cancelled", Operator:=xlOr, Criteria2:="CL  "

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("R2:R" & lr1), "Cancelled") + Application.WorksheetFunction.CountIf(Range("R2:R" & lr1), "CL  ")

'CP Cancelled data

If Sheets("Working_Sheet").Range("XFD1").Text > 0 Then
Sheets("Working_Sheet").Select
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Cancelled").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("Working_Sheet").Select
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp

End If

Range("A1:U1").AutoFilter

'CP -ve data

Sheets("Working_Sheet").Activate
Range("A1:U1").AutoFilter Field:=17, Criteria1:="<0"

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("Q2:Q" & lr1), ">0")

If Sheets("Working_Sheet").Range("XFD1").Text > 1 Then

'Check if there is data or no

If Sheets("Cancelled").Range("A1").Value = "" Then

Sheets("Working_Sheet").Select
Range("A1:U" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Cancelled").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Else

Sheets("Working_Sheet").Select
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Cancelled").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

End If

Sheets("Working_Sheet").Select
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp

End If

Sheets("Working_Sheet").Select
Range("A1:U1").AutoFilter

Application.DisplayAlerts = False
Application.ScreenUpdating = False

ThisWorkbook.Activate

Call NA

End Sub

Sub NA()

'Add New Sheet

Sheets("Raw").Select
Sheets.Add Before:=Sheets("Raw")
ActiveSheet.Name = "NA"

'Filter Data

Sheets("Working_Sheet").Activate
Range("A1:U1").AutoFilter Field:=18, Criteria1:="#N/A"

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("R2:R" & lr1), "#N/A")

If Sheets("Working_Sheet").Range("XFD1").Text > 0 Then
Sheets("Working_Sheet").Select
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("NA").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("Working_Sheet").Select
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp

End If

Range("A1:U1").AutoFilter

Call Rejected

End Sub

Sub Rejected()

ThisWorkbook.Activate

Sheets("Raw").Select
Sheets.Add Before:=Sheets("Raw")
ActiveSheet.Name = "Rejected"

Sheets("DataInput_Instructions").Activate


'Does not apply for EMEA & US

If Range("B3").Value = "Dell Technologies Switzerland" Or _
Range("B3").Value = "Dell Technologies Germany" Or _
Range("B3").Value = "Dell Technologies Spain" Or _
Range("B3").Value = "Dell Technologies France" Or _
Range("B3").Value = "Dell Technologies Netherlands" Or _
Range("B3").Value = "Dell Technologies Sweden" Or _
Range("B3").Value = "Dell Technologies UK" Then

Else

Sheets("Working_Sheet").Activate
Range("A1:U1").AutoFilter Field:=19, Criteria1:="Reject  ", Operator:=xlOr, Criteria2:="Rejected"

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("S2:S" & lr1), "Reject  ") + Application.WorksheetFunction.CountIf(Range("S2:S" & lr1), "Rejected")

If Sheets("Working_Sheet").Range("XFD1").Text > 0 Then
Sheets("Working_Sheet").Select
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Rejected").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("Working_Sheet").Select
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp

End If

Range("A1:U1").AutoFilter

End If

Call No_Action_Required

End Sub


Sub No_Action_Required()

Sheets("Raw").Select
Sheets.Add Before:=Sheets("Raw")
ActiveSheet.Name = "No_Action_Required"

'Filter column " Diff %" for values between -5% and +5%

Sheets("Working_Sheet").Activate
Range("A1:U1").AutoFilter Field:=21, Criteria1:=">=-5%", Operator:=xlAnd, Criteria2:="<=5%"

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIfs(Range("U2:U" & lr1), ">=-5%", Range("U2:U" & lr1), "<=5%")

If Sheets("Working_Sheet").Range("XFD1").Text > 0 Then
Sheets("Working_Sheet").Select
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("No_Action_Required").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("Working_Sheet").Select
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp
End If

Range("A1:U1").AutoFilter

'Filter % above  or >=100%

Sheets("Working_Sheet").Activate
Range("A1:U1").AutoFilter Field:=21, Criteria1:=">=100%", Operator:=xlAnd

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIfs(Range("U2:U" & lr1), ">=100%")


If Sheets("Working_Sheet").Range("XFD1").Text > 0 Then
If Sheets("No_Action_Required").Range("A1").Value = "" Then
Sheets("Working_Sheet").Select
Range("A1:U" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("No_Action_Required").Select
Range("A1").Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Else
Sheets("Working_Sheet").Select
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("No_Action_Required").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit
End If


Sheets("Working_Sheet").Select
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp

End If

Range("A1:U1").AutoFilter

'Filter Difference in sale amount in -ve

Sheets("Working_Sheet").Activate
Range("A1:U1").AutoFilter Field:=20, Criteria1:="<0"

Sheets("Working_Sheet").Select
lr1 = 0
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("T2:T" & lr1), ">0")

If Sheets("Working_Sheet").Range("XFD1").Text > 1 Then
If Sheets("No_Action_Required").Range("A1").Value = "" Then
Sheets("Working_Sheet").Select
Range("A1:U" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("No_Action_Required").Select
Range("A1").Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Else
Sheets("Working_Sheet").Select
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("No_Action_Required").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit
End If


Sheets("Working_Sheet").Select
lr1 = Sheets("Working_Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:U" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp

End If

Sheets("Working_Sheet").Select
Range("A1:U1").AutoFilter
Range("A1").Select

Call OIDs_Lookup

End Sub

Sub OIDs_Lookup()

'Cancelled WorkingSheet

Sheets("Cancelled").Select

If IsEmpty(Range("A1").Value) Then
Else

Range("V2").Value = "=XLOOKUP(H2,Working_Sheet!H:H,Working_Sheet!H:H,""NA"")"
Range("Q2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 5).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown

Sheets("Cancelled").Select
lr1 = 0
lr1 = Sheets("Cancelled").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("V2:V" & lr1), "<>NA")

If Sheets("Cancelled").Range("XFD1").Text > 0 Then
Range("A1:V1").AutoFilter Field:=22, Criteria1:="<>NA"

'Dell sale amount cell to "Zero"
On Error Resume Next
ActiveSheet.Range("Q2:Q" & lr1).SpecialCells(xlCellTypeVisible).Value = 0
On Error GoTo 0

'CJ sale amount cell (finalSaleAmountUsd - Column L) to "-ve to the existing value"
For Each cell In ActiveSheet.Range("M2:M" & Cells(Rows.Count, 13).End(xlUp).Row).SpecialCells(xlCellTypeVisible)
If IsNumeric(cell.Value) Then
cell.Value = -Abs(cell.Value)
End If
Next cell

Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Working_Sheet").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("Cancelled").Select
lr1 = Sheets("Cancelled").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp
ActiveSheet.AutoFilterMode = False

End If

Range("V:V").ClearContents

End If

'Rejected WorkingSheet

Sheets("Rejected").Select

If IsEmpty(Range("A1").Value) Then
Else
Range("V2").Value = "=XLOOKUP(H2,Working_Sheet!H:H,Working_Sheet!H:H,""NA"")"
Range("Q2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 5).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown

Sheets("Rejected").Select
lr1 = 0
lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("V2:V" & lr1), "<>NA")

If Sheets("Rejected").Range("XFD1").Text > 0 Then
Range("A1:V1").AutoFilter Field:=22, Criteria1:="<>NA"

'Dell sale amount cell to "Zero"
On Error Resume Next
ActiveSheet.Range("Q2:Q" & lr1).SpecialCells(xlCellTypeVisible).Value = 0
On Error GoTo 0

'CJ sale amount cell (finalSaleAmountUsd - Column L) to "-ve to the existing value"
For Each cell In ActiveSheet.Range("M2:M" & Cells(Rows.Count, 13).End(xlUp).Row).SpecialCells(xlCellTypeVisible)
If IsNumeric(cell.Value) Then
cell.Value = -Abs(cell.Value)
End If
Next cell

Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Working_Sheet").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("Rejected").Select
lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp
ActiveSheet.AutoFilterMode = False

End If

Range("V:V").ClearContents

End If


'NA WorkingSheet

Sheets("NA").Select

If IsEmpty(Range("A1").Value) Then
Else
If IsEmpty(Range("A3").Value) Then
Range("V2").Value = "=XLOOKUP(H2,Working_Sheet!H:H,Working_Sheet!H:H,""NA"")"
Else
Range("V2").Value = "=XLOOKUP(H2,Working_Sheet!H:H,Working_Sheet!H:H,""NA"")"
Range("Q2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 5).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
End If


Sheets("NA").Select
lr1 = 0
lr1 = Sheets("NA").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("V2:V" & lr1), "<>NA")

If Sheets("NA").Range("XFD1").Text > 0 Then
Range("A1:V1").AutoFilter Field:=22, Criteria1:="<>NA"

'Dell sale amount cell to "Zero"
On Error Resume Next
ActiveSheet.Range("Q2:Q" & lr1).SpecialCells(xlCellTypeVisible).Value = 0
On Error GoTo 0

'CJ sale amount cell (finalSaleAmountUsd - Column L) to "-ve to the existing value"
For Each cell In ActiveSheet.Range("M2:M" & Cells(Rows.Count, 13).End(xlUp).Row).SpecialCells(xlCellTypeVisible)
If IsNumeric(cell.Value) Then
cell.Value = -Abs(cell.Value)
End If
Next cell

Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Working_Sheet").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("NA").Select
lr1 = Sheets("NA").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp
ActiveSheet.AutoFilterMode = False

End If

Range("V:V").ClearContents

End If

'Cancelled No_Action_Required

Sheets("Cancelled").Select

If IsEmpty(Range("A2").Value) Then
Else

Range("V2").Value = "=XLOOKUP(H2,No_Action_Required!H:H,No_Action_Required!H:H,""NA"")"
Range("Q2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 5).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown

Sheets("Cancelled").Select
lr1 = 0
lr1 = Sheets("Cancelled").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("V2:V" & lr1), "<>NA")

If Sheets("Cancelled").Range("XFD1").Text > 0 Then
Range("A1:V1").AutoFilter Field:=21, Criteria1:="<>NA"

'Dell sale amount cell to "Zero"
On Error Resume Next
ActiveSheet.Range("Q2:Q" & lr1).SpecialCells(xlCellTypeVisible).Value = 0
On Error GoTo 0

'CJ sale amount cell (finalSaleAmountUsd - Column L) to "-ve to the existing value"
For Each cell In ActiveSheet.Range("M2:M" & Cells(Rows.Count, 13).End(xlUp).Row).SpecialCells(xlCellTypeVisible)
If IsNumeric(cell.Value) Then
cell.Value = -Abs(cell.Value)
End If
Next cell

Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Working_Sheet").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("Cancelled").Select
lr1 = Sheets("Cancelled").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp
ActiveSheet.AutoFilterMode = False

End If

Range("V:V").ClearContents

End If


'Rejected No_Action_Required

Sheets("Rejected").Select

If IsEmpty(Range("A2").Value) Then
Else

Range("V2").Value = "=XLOOKUP(H2,No_Action_Required!H:H,No_Action_Required!H:H,""NA"")"
Range("Q2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 5).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown

Sheets("Rejected").Select
lr1 = 0
lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("V2:V" & lr1), "<>NA")

If Sheets("Rejected").Range("XFD1").Text > 0 Then
Range("A1:V1").AutoFilter Field:=22, Criteria1:="<>NA"

'Dell sale amount cell to "Zero"
On Error Resume Next
ActiveSheet.Range("Q2:Q" & lr1).SpecialCells(xlCellTypeVisible).Value = 0
On Error GoTo 0

'CJ sale amount cell (finalSaleAmountUsd - Column L) to "-ve to the existing value"
For Each cell In ActiveSheet.Range("M2:M" & Cells(Rows.Count, 13).End(xlUp).Row).SpecialCells(xlCellTypeVisible)
If IsNumeric(cell.Value) Then
cell.Value = -Abs(cell.Value)
End If
Next cell

Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Working_Sheet").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("Rejected").Select
lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp
ActiveSheet.AutoFilterMode = False

End If

Range("V:V").ClearContents

End If

'NA No_Action_Required

Sheets("NA").Select

If IsEmpty(Range("A2").Value) Then
Else
If IsEmpty(Range("A3").Value) Then
Range("V2").Value = "=XLOOKUP(H2,Working_Sheet!H:H,Working_Sheet!H:H,""NA"")"
Else
Range("V2").Value = "=XLOOKUP(H2,No_Action_Required!H:H,No_Action_Required!H:H,""NA"")"
Range("Q2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 5).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
End If

lr1 = 0
lr1 = Sheets("NA").Cells(Rows.Count, 1).End(xlUp).Row
Range("XFD1").Value = Application.WorksheetFunction.CountIf(Range("V2:V" & lr1), "<>NA")

If Sheets("NA").Range("XFD1").Text > 0 Then
Range("A1:V1").AutoFilter Field:=22, Criteria1:="<>NA"

'Dell sale amount cell to "Zero"
On Error Resume Next
ActiveSheet.Range("Q2:Q" & lr1).SpecialCells(xlCellTypeVisible).Value = 0
On Error GoTo 0

'CJ sale amount cell (finalSaleAmountUsd - Column L) to "-ve to the existing value"
For Each cell In ActiveSheet.Range("M2:M" & Cells(Rows.Count, 13).End(xlUp).Row).SpecialCells(xlCellTypeVisible)
If IsNumeric(cell.Value) Then
cell.Value = -Abs(cell.Value)
End If
Next cell

Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

Sheets("Working_Sheet").Select
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValuesAndNumberFormats
Columns("H:L").EntireColumn.AutoFit

Sheets("NA").Select
lr1 = Sheets("NA").Cells(Rows.Count, 1).End(xlUp).Row
Range("A2:V" & lr1).SpecialCells(xlCellTypeVisible).Delete xlUp
ActiveSheet.AutoFilterMode = False

End If

Range("V:V").ClearContents

End If

Range("V:V").ClearContents
Sheets("Working_Sheet").Select
Range("V:V").ClearContents

Call Creating_Pivot

End Sub

Sub Creating_Pivot()

Sheets("Working_Sheet").Select

If Range("A2").Value = "" Then
MsgBox "WARNING! No Data in Working_Sheet to create a Pivot "
Exit Sub
End If

Dim PTable As PivotTable
Dim PCache As PivotCache
Dim PRange As Range
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim LR As Long
Dim LC As Long

Application.ScreenUpdating = False

Sheets("Working_Sheet").Select
Worksheets.Add after:=ActiveSheet
ActiveSheet.Name = "Pivot"

On Error GoTo 0

Set PSheet = Worksheets("Pivot")
Set DSheet = Worksheets("Working_Sheet")

LR = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LC = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column

Set PRange = DSheet.Cells(1, 1).Resize(LR, LC)
Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange)
Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), TableName:="Pivot")

With PSheet.PivotTables("Pivot").PivotFields("advertiser")
.Subtotals(1) = False
.Orientation = xlPageField
.Position = 1
End With

With PSheet.PivotTables("Pivot").PivotFields("oid")
.Subtotals(1) = False
.Orientation = xlRowField
.Position = 1
End With

With PSheet.PivotTables("Pivot").PivotFields("sku")
.Subtotals(1) = False
.Orientation = xlRowField
.Position = 2
End With

With PSheet.PivotTables("Pivot").PivotFields("commission_id")
.Subtotals(1) = False
.Orientation = xlRowField
.Position = 3
End With

With PSheet.PivotTables("Pivot").PivotFields("finalSaleAmountUsd")
.Subtotals(1) = False
.Orientation = xlDataField
.Function = xlSum
.Position = 1
End With

With PSheet.PivotTables("Pivot").PivotFields("Dell_Final_Sale_Amount")
.Subtotals(1) = False
.Orientation = xlDataField
.Function = xlSum
.Position = 2
End With

ActiveSheet.PivotTables("Pivot").RowAxisLayout xlTabularRow
ActiveSheet.PivotTables("Pivot").ColumnGrand = True
ActiveSheet.PivotTables("Pivot").RowGrand = False

'Copy Data from Pivot to Formula Sheet

Sheets("Working_Sheet").Select
Worksheets.Add after:=ActiveSheet
ActiveSheet.Name = "Formula_Sheet"

Range("E1").Value = "Difference"
Range("F1").Value = "SKU_Difference"
Range("G1").Value = "Process"
Range("H1").Value = "Final"

Sheets("Pivot").Select
Range("B4").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToLeft)).Select
Selection.Copy

Sheets("Formula_Sheet").Select
Range("A1").PasteSpecial xlPasteValues

Range("B1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveCell.Offset(0, -1).Select
ActiveCell.Value = "Grand Total"

Sheets("Pivot").Select
Range("C4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("Formula_Sheet").Select
Range("I1").PasteSpecial xlPasteValues

Sheets("Pivot").Select
Range("D4:E4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("Formula_Sheet").Select
Range("C1").PasteSpecial xlPasteValues

'Add formulas E to H column & filldown

lr1 = 0
lr1 = Sheets("Formula_Sheet").Cells(Rows.Count, 9).End(xlUp).Row

Range("E2").Value = "=C2-D2"
Range("F2").Value = "=IF(A2<>"""", G2, """")"
Range("G2").Value = "=IF(A2<>"""", B2&"":""&E2, """")"
Range("E2:G" & lr1).Select
Selection.FillDown

Range("G3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

'G3 formula

Range("G3").Formula = "=IF(A3="""",G2&"";;""&B3&"":""&E3, IF(A3<>"""", B3&"":""&E3,""""))"
Sheets("Formula_Sheet").Select
Range("G3:G" & lr1).Select
Selection.FillDown

'H Column formula

Range("H2").Value = "=IF(ISBLANK(A3),"""",G2)"
Range("H2:H" & lr1).Select
Selection.FillDown

'Final column Text to Columns
Range("H1", Range("H" & Rows.Count).End(xlUp)).Copy
Range("H1").PasteSpecial Paste:=xlPasteValues
Range("H1", Range("H" & Rows.Count).End(xlUp)).TextToColumns Destination:=Range("H1"), DataType:=xlFixedWidth, FieldInfo:=Array(0, 1), TrailingMinusNumbers:=True
Application.CutCopyMode = False

'Batch sheet creation

Sheets("Formula_Sheet").Select
Range("I1").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToLeft)).Select
Selection.Copy

Sheets("Formula_Sheet").Select
Worksheets.Add after:=ActiveSheet
ActiveSheet.Name = "Batch"

Range("A1").PasteSpecial xlPasteValues
Columns("A:A").NumberFormat = "0"
Columns("I:A").EntireColumn.AutoFit
Range("A1:I1").Font.Bold = True

'Fill Down cells in column OID

lr1 = 0
lr1 = Sheets("Batch").Cells(Rows.Count, 2).End(xlUp).Row
Range("A1:A" & lr1).Select
x = Application.WorksheetFunction.CountA(Range("A1:A" & lr1))
If x < lr1 Then
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.FormulaR1C1 = "=R[-1]C"
End If

Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Range("A2").PasteSpecial xlPasteValues

'Filter column H (Final)

'With Sheets("Batch")
'Range("A1:I1").AutoFilter Field:=8, Criteria1:=""

lr1 = 0
lr1 = Sheets("Batch").Cells(Rows.Count, 1).End(xlUp).Row
Range("H1:H" & lr1).Select
x = Application.WorksheetFunction.CountA(Range("H1:H" & lr1))
If x < lr1 Then
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete
End If


'On Error Resume Next
'.AutoFilter.Range.Offset(1, 0).Resize(.AutoFilter.Range.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
'On Error GoTo 0
'Range("A1:I1").AutoFilter
'End With

Call Creating_Correction_file

End Sub

Sub Creating_Correction_file()

Application.DisplayAlerts = False

'Delete the file if it exists in the folder

On Error Resume Next
Kill ThisWorkbook.Path & "\CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value
On Error GoTo 0

'Create Correction File

Workbooks.Add
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value, FileFormat:=xlCSV

'CID & SUBID creation

Workbooks("SKU1.0").Activate
Sheets("DataInput_Instructions").Select

Range("XFD1").Value = "&CID=" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B4").Value
Range("XFD2").Value = "&SUBID=" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B5").Value
Range("XFD1:XFD2").Copy

Workbooks("CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("A1").PasteSpecial xlPasteValues

'Copy paste data from Batch to CorrectionFile

'Commision_id

Workbooks("SKU1.0").Activate
Sheets("Batch").Select
Range("I2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("B3").PasteSpecial xlPasteValues

'oid

Workbooks("SKU1.0").Activate
Sheets("Batch").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("C3").PasteSpecial xlPasteValues

'Final

Workbooks("SKU1.0").Activate
Sheets("Batch").Select
Range("H2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("D3").PasteSpecial xlPasteValues

Range("A3").Value = "OTHER"
Range("A3").AutoFill Destination:=Range("A3:A" & Cells(Rows.Count, "B").End(xlUp).Row)

'Unique IDs

ThisWorkbook.Activate
Dim A As String
Dim B As String
 
A = "CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value
B = ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value
 
ThisWorkbook.Sheets("DataInput_Instructions").Activate
Range("B6").Copy
Workbooks.Open Filename:=ThisWorkbook.Path & "\" & A
Sheets(1).Select
Range("XFD1").PasteSpecial xlPasteValues
Range("E3").Value = "=$XFD$1+1"
Range("E4").Value = "=E3+1"

Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 4).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
 
Range("E3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Range("E3").PasteSpecial xlPasteValues
Range("E3").Select
Selection.End(xlDown).Select
ActiveCell.Copy
 
ThisWorkbook.Activate
Worksheets("CID_SUBIDs").Visible = True
Sheets("CID_SUBIDs").Select
Range("A1:A13").Find(What:=B).Select
ActiveCell.Offset(0, 3).Select
Selection.PasteSpecial xlPasteValues
Worksheets("CID_SUBIDs").Visible = False

'Autofit & Number format

Workbooks("CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("XFD1").ClearContents
Columns("B:C").NumberFormat = "0"
Columns("I:A").EntireColumn.AutoFit

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Workbooks("CorrectionFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Close SaveChanges:=True

Call Creating_Reversal_file

End Sub

Sub Creating_Reversal_file()

Application.DisplayAlerts = False

'Delete the file if it exists in the folder

On Error Resume Next
Kill ThisWorkbook.Path & "\ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value
On Error GoTo 0

'Create a Reversal File

Workbooks.Add
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value, FileFormat:=xlCSV

'CID & SUBID creation

Workbooks("SKU1.0").Activate
Sheets("DataInput_Instructions").Select

Range("XFD1").Value = "&CID=" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B4").Value
Range("XFD2").Value = "&SUBID=" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B5").Value
Range("XFD1:XFD2").Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("A1").PasteSpecial xlPasteValues


'Copy paste data from Sheets to ReversalFile

'Cancelled

Workbooks("SKU1.0").Activate
Sheets("Cancelled").Select

'Add CommisionID

If Range("A2").Value <> "" Then
Range("I2").Select
lr1 = 0
lr1 = Sheets("Cancelled").Cells(Rows.Count, 1).End(xlUp).Row
Range("I2:I" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("B3").PasteSpecial xlPasteValues

'Add OID

Workbooks("SKU1.0").Activate
Sheets("Cancelled").Select
Range("H2").Select
lr1 = 0
lr1 = Sheets("Cancelled").Cells(Rows.Count, 1).End(xlUp).Row
Range("H2:H" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("C3").PasteSpecial xlPasteValues

'Add RETRUN

lr1 = 0
lr1 = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 2).End(xlUp).Row
Range("A3:A" & lr1).Value = "RETURN"

End If

'NA

Workbooks("SKU1.0").Activate
Sheets("NA").Select
If Range("A2").Value <> "" Then

'If Cancelled isnt available

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
If Range("A3").Value = "" Then

'Add CommisionID

Workbooks("SKU1.0").Activate
Sheets("NA").Select
Range("I2").Select
lr1 = 0
lr1 = Sheets("NA").Cells(Rows.Count, 1).End(xlUp).Row
Range("I2:I" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("B3").PasteSpecial xlPasteValues

'Add OID

Workbooks("SKU1.0").Activate
Sheets("NA").Select
Range("H2").Select
lr1 = 0
lr1 = Sheets("NA").Cells(Rows.Count, 1).End(xlUp).Row
Range("H2:H" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("C3").PasteSpecial xlPasteValues

'Add DUPO

lr1 = 0
lr1 = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 2).End(xlUp).Row
Range("A3:A" & lr1).Value = "DUPO"

Else


'Add CommissionID

Workbooks("SKU1.0").Activate
Sheets("NA").Select
lr1 = 0
lr1 = Sheets("NA").Cells(Rows.Count, 1).End(xlUp).Row
Range("I2:I" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
With ActiveSheet
lr1 = .Cells(Rows.Count, 1).End(xlUp).Row
If .Cells(lr1 + 1, 2) = "" Then
.Cells(lr1 + 1, 2).PasteSpecial xlPasteValues
Else
Range("B3").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValues
End If
End With


'Add OID

Workbooks("SKU1.0").Activate
Sheets("NA").Select
lr1 = 0
lr1 = Sheets("NA").Cells(Rows.Count, 1).End(xlUp).Row
Range("H2:H" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
With ActiveSheet
lr1 = .Cells(Rows.Count, 1).End(xlUp).Row
If .Cells(lr1 + 1, 3) = "" Then
.Cells(lr1 + 1, 3).PasteSpecial xlPasteValues
Else
Range("C3").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValues
End If
End With

'Add DUPO

lr1 = 0
lr1 = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 2).End(xlUp).Row
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Range(Selection, "A" & lr1).Select
Selection.Value = "DUPO"

End If
End If


'Rejected

Workbooks("SKU1.0").Activate
Sheets("Rejected").Select
If Range("A2").Value <> "" Then

'If Cancelled isnt available
Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
If Range("A3").Value = "" Then

'Add CommissionID

Workbooks("SKU1.0").Activate
Sheets("Rejected").Select
Range("I2").Select
lr1 = 0
lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
Range("I2:I" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("B3").PasteSpecial xlPasteValues

'Add OID

Workbooks("SKU1.0").Activate
Sheets("Rejected").Select
Range("H2").Select
lr1 = 0
lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
Range("H2:H" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Range("C3").PasteSpecial xlPasteValues

'Add OTHER

lr1 = 0
lr1 = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 2).End(xlUp).Row
Range("A3:A" & lr1).Value = "OTHER"

Else


'Add CommissionID

'Workbooks("SKU1.0").Activate
'Sheets("Rejected").Select
'Range("I2").Select
'lr1 = 0
'lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
'Range("I2:I" & lr1).SpecialCells(xlCellTypeVisible).Copy

'Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
'Range("B3").Select
'lr1 = 0
'lr1 = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row

'If Range("B4") = "" Then
'Range("B:B" & lr1).SpecialCells(xlCellTypeVisible).PasteSpecial xlPasteValues
'Else
'Range("B3").Select
'Selection.End(xlDown).Select
'ActiveCell.Offset(1, 0).Select
'Selection.PasteSpecial xlPasteValues
'End If



Workbooks("SKU1.0").Activate
Sheets("Rejected").Select
lr1 = 0
lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
Range("I2:I" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
With ActiveSheet
lr1 = .Cells(Rows.Count, 1).End(xlUp).Row
If .Cells(lr1 + 1, 2) = "" Then
.Cells(lr1 + 1, 2).PasteSpecial xlPasteValues
Else
Range("B3").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValues
End If
End With



'Add OID

'Workbooks("SKU1.0").Activate
'Sheets("Rejected").Select
'Range("H2").Select
'lr1 = 0
'lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
'Range("H2:H" & lr1).SpecialCells(xlCellTypeVisible).Copy

'Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
'Range("C3").Select
'lr1 = 0
'lr1 = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row

'If Range("C4") = "" Then
'Range("C4:C" & lr1).SpecialCells(xlCellTypeVisible).PasteSpecial xlPasteValues
'Else
'Range("C3").Select
'Selection.End(xlDown).Select
'ActiveCell.Offset(1, 0).Select
'Selection.PasteSpecial xlPasteValues
'End If


Workbooks("SKU1.0").Activate
Sheets("Rejected").Select
lr1 = 0
lr1 = Sheets("Rejected").Cells(Rows.Count, 1).End(xlUp).Row
Range("H2:H" & lr1).SpecialCells(xlCellTypeVisible).Copy

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
With ActiveSheet
lr1 = .Cells(Rows.Count, 1).End(xlUp).Row
If .Cells(lr1 + 1, 3) = "" Then
.Cells(lr1 + 1, 3).PasteSpecial xlPasteValues
Else
Range("C3").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Selection.PasteSpecial xlPasteValues
End If
End With


'Add OTHER

lr1 = 0
lr1 = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 2).End(xlUp).Row
Range("A1").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
Range(Selection, "A" & lr1).Select
Selection.Value = "OTHER"

End If
End If


Application.DisplayAlerts = False
Application.ScreenUpdating = False

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Activate
Columns("B:C").NumberFormat = "0"
Columns("A:C").EntireColumn.AutoFit

Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Save
Workbooks("ReversalFile-" & ThisWorkbook.Sheets("DataInput_Instructions").Range("B3").Value).Close 'SaveChanges:=True

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Workbooks("SKU1.0").Activate
Sheets("DataInput_Instructions").Move Before:=Sheets("Working_Sheet")
Sheets("DataInput_Instructions").Select
Range("A1").Select

MsgBox "Data Refreshed"

End Sub

