Attribute VB_Name = "Shortcuts"
Option Explicit

Sub add_cost()
Attribute add_cost.VB_Description = "Inputs today's date into column Z (parked date) for every selected row."
Attribute add_cost.VB_ProcData.VB_Invoke_Func = "q\n14"
'Keyboard Shortcut: CTRL+Q

Dim area As Range
Dim flag As Integer
Dim question As Integer

flag = 0

For Each area In Selection 'if empty check
    If Cells(area.Row, 26) = "" Then
        flag = 0
    Else
        flag = 1
        GoTo test
    End If
Next

test:
If flag = 0 Then
Paste:
    For Each area In Selection
        Cells(area.Row, 26).Select
        ActiveCell.FormulaR1C1 = "=TODAY()"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Cells(area.Row + 1, area.Column).Activate
    Next
Else
    question = MsgBox("You are trying to paste today's date into unempty cells. Are you sure?", vbYesNo)
    If question = 6 Then
        GoTo Paste
    Else
        MsgBox ("Macro abandoned")
    End If
End If

End Sub

Sub last_action()
Attribute last_action.VB_Description = "Inputs today's date into column Y (last action date) for every selected row."
Attribute last_action.VB_ProcData.VB_Invoke_Func = "w\n14"
'Keyboard Shortcut: CTRL+W

Dim area As Range
Dim flag As Integer
Dim question As Integer

flag = 0

For Each area In Selection 'if empty check
    If Cells(area.Row, 25) = "" Then
        flag = 0
    Else
        flag = 1
        GoTo test
    End If
Next

test:
If flag = 0 Then
Paste:
    For Each area In Selection
        Cells(area.Row, 25).Select
        ActiveCell.FormulaR1C1 = "=TODAY()"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Cells(area.Row + 1, area.Column).Activate
    Next
Else
    question = MsgBox("You are trying to paste today's date into unempty cells. Are you sure?", vbYesNo)
    If question = 6 Then
        GoTo Paste
    Else
        MsgBox ("Macro abandoned")
    End If
End If

End Sub

Sub reject()
Attribute reject.VB_Description = "Inputs date into columns 22, 23, 25 as if we rejected cost for every selected row."
Attribute reject.VB_ProcData.VB_Invoke_Func = "r\n14"
'Keyboard Shortcut: CTRL+R

Dim area As Range
Dim flag As Integer
Dim question As Integer

flag = 0

For Each area In Selection 'if empty check: Decision / Date of decision / Latest action date
    If Cells(area.Row, 22) = "" And Cells(area.Row, 23) = "" Then
        flag = 0
    Else
        flag = 1
        GoTo test
    End If
Next

test:
If flag = 0 Then
Paste:
    For Each area In Selection
        Cells(area.Row, 22).Select  'Decision Key Contact Person
        ActiveCell.Value = "Rejected"
        Cells(area.Row, 23).Select  'Date of decision
        ActiveCell.FormulaR1C1 = "=TODAY()"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Cells(area.Row, 25).Select  'Latest action date
        ActiveCell.FormulaR1C1 = "=TODAY()"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Cells(area.Row + 1, area.Column).Activate
    Next
Else
    question = MsgBox("You are trying to paste new data into unempty cells. Are you sure?", vbYesNo)
    If question = 6 Then
        GoTo Paste
    Else
        MsgBox ("Macro abandoned")
    End If
End If

End Sub

