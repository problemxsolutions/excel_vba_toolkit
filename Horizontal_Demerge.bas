Attribute VB_Name = "Module3"
Sub Horizontal_Demerge()
Dim ParentWB As Workbook
Dim CurrentWS As Worksheet
Dim PrimRange As Range

Set ParentWB = ActiveWorkbook
Set CurrentWS = ParentWB.ActiveSheet

Set PrimRange = Application.InputBox(Prompt:="Unmerge Columns", Title:="Specify Range to Unmerge", Type:=8)
ProcessRange = PrimRange.Address

Range(ProcessRange).UnMerge
For Each i In Range(ProcessRange)
    i_row = i.Row
    i_col = i.Column
    i_address = Cells(i_row, i_col).Address
    
    If i = "" Then
        Range(i_address).Value = Range(i_address).Offset(0, -1).Value
    End If
Next
    MsgBox "Column Unmerge Complete"
    
End Sub
