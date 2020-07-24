Attribute VB_Name = "Module2"
Sub Vertical_Demerge()
Dim ParentWB As Workbook
Dim CurrentWS As Worksheet
Dim PrimRange As Range

Set ParentWB = ActiveWorkbook
Set CurrentWS = ParentWB.ActiveSheet

Set PrimRange = Application.InputBox(Prompt:="Unmerge Rows", _
                                        Title:="Specify Range to Unmerge", _
                                        Type:=8)
ProcessRange = PrimRange.Address

Range(ProcessRange).UnMerge
For Each i In Range(ProcessRange)
    i_row = i.Row
    i_col = i.Column
    i_address = Cells(i_row, i_col).Address
    
    If i = "" Then
        Range(i_address).Value = Range(i_address).Offset(-1, 0).Value
    End If
Next
    MsgBox "Row Unmerge Complete"
End Sub

