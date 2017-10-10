Attribute VB_Name = "AdditionalCosts"
Option Explicit

Sub Daily_AC()

Dim MyDate As Date
Dim Yesterday As Date
Dim parkedYesterday As Integer
Dim closedYesterday As Integer
Dim touchedYesterday As Integer

Worksheets("Sheet1").Activate

Dim park As Integer
park = Application.WorksheetFunction.CountIf(Range("AC:AC"), "Parked")

Dim over As Integer
over = Application.WorksheetFunction.Sum(Application.WorksheetFunction.CountIfs(Range("AC:AC"), "New", Range("AD:AD"), ">30"), _
        Application.WorksheetFunction.CountIfs(Range("AC:AC"), "Waiting for approval", Range("AD:AD"), ">30"))

Dim pend As Integer
pend = Application.WorksheetFunction.CountIf(Range("AC:AC"), "Waiting for approval") - over

Dim news As Integer
news = Application.WorksheetFunction.CountIf(Range("AC:AC"), "New")

MyDate = Date
Yesterday = Application.WorksheetFunction.WorkDay(Date, -1)
parkedYesterday = Application.WorksheetFunction.CountIf(Range("Z:Z"), Yesterday)
closedYesterday = Application.WorksheetFunction.CountIf(Range("AB:AB"), Yesterday)
touchedYesterday = Application.WorksheetFunction.CountIf(Range("Y:Y"), Yesterday)
MsgBox "Today it is " & MyDate & ". There are: " & vbNewLine & news & " new costs" & vbNewLine & park & " parked costs" & vbNewLine & pend & " pending costs (without overdues)" & _
        vbNewLine & over & " overdue costs (over 30 days)" & vbNewLine & vbNewLine & "Yesterday:" & vbNewLine & parkedYesterday & " cases parked" & _
        vbNewLine & closedYesterday & " cases closed" & vbNewLine & touchedYesterday & " cases worked with"
End Sub


Sub CopyRowsFromAbove()

Dim rowsToCopy As Long
Dim numRows As Long
Dim rng1 As Range, rng2 As Range, rng3 As Range
Dim trg1 As Range, trg2 As Range, trg3 As Range
Dim testCell As Range
Dim lastRowIndex As Long

numRows = Selection.Row
rowsToCopy = Selection.Rows.Count
lastRowIndex = firstFree(ActiveSheet, 1).Row - 1

Set testCell = Range("A" & Selection.Row)

If testCell.Value <> "" Then    'if selection on old rows (based on column A)
    Set rng1 = Range(Cells(numRows, 1), Cells(numRows + rowsToCopy - 1, 6))     'columns 1 to 6
    Set rng2 = Range(Cells(numRows, 29), Cells(numRows + rowsToCopy - 1, 30))   'columns 29 to 30
    Set rng3 = Rows(numRows & ":" & numRows + rowsToCopy - 1)                   'selected row formats
    
    Set trg1 = Range("A" & lastRowIndex)                                        'targets where to paste
    Set trg2 = Range("AC" & lastRowIndex)
    Set trg3 = trg1
Else                            'if selection on new rows (based on column A)
    Set rng1 = Range(Cells(lastRowIndex, 1), Cells(lastRowIndex, 6))            'columns 1 to 6
    Set rng2 = Range(Cells(lastRowIndex, 29), Cells(lastRowIndex, 30))          'columns 29 to 30
    Set rng3 = Rows(lastRowIndex)                                               'last row formats
    
    Set trg1 = Range("A" & lastRowIndex & ":A" & lastRowIndex + rowsToCopy)
    Set trg2 = Range("AC" & lastRowIndex & ":AD" & lastRowIndex + rowsToCopy)
    Set trg3 = Range("A" & lastRowIndex + 1 & ":AG" & lastRowIndex + rowsToCopy)
End If

Call CopyPaste(rng3, trg3, "formats")   'copy formats
Call CopyPaste(rng1, trg1, "values")    'copy values
Call CopyPaste(rng2, trg2, "values")

Range("A" & ActiveSheet.UsedRange.Rows.Count).Select

End Sub

Sub CopyPaste(what As Range, where As Range, Optional mode As String)

what.Select
what.Copy

If mode = "formats" Then
    where.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
ElseIf mode = "values" Then
    where.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Else
    where.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End If

End Sub

Sub checkDoubles()

Dim numRows As Long
Dim rowsToCheck As Long
Dim rangeToCheck As Range
Dim doubles As Long
Dim rowCounter As Long
Dim cell As Range

'Application.ScreenUpdating = False

numRows = Selection.Row
rowsToCheck = Selection.Rows.Count
rowCounter = 0

Set rangeToCheck = Range(Cells(numRows, 2), Cells(numRows + rowsToCheck - 1, 2)).SpecialCells(xlCellTypeVisible)

For Each cell In rangeToCheck
    rowCounter = rowCounter + 1
    Application.StatusBar = "Checking row: " & rowCounter & " out of " & rowsToCheck
    
    '=IF(COUNTIFS(G:G;G2;H:H;H2;Q:Q;Q2;R:R;R2)>1;"Doubles!")
    doubles = Application.WorksheetFunction.CountIfs(Range("G:G"), Cells(cell.Row, 7).Value, Range("H:H"), Cells(cell.Row, 8), Range("Q:Q"), Cells(cell.Row, 17).Value, Range("R:R"), Cells(cell.Row, 18).Value)
    If doubles > 1 Then
        cell.Value = "Doubles: " & doubles
    Else
        cell.Value = ""
    End If
Next cell


Application.ScreenUpdating = True
Application.StatusBar = False
End Sub

Function firstFree(works As Worksheet, column As Long) As Range
Dim cell As Range

    For Each cell In works.Columns(column).Cells
        If cell.Value = "" Then
            Set firstFree = cell
            Exit For
        End If
    Next cell
End Function
