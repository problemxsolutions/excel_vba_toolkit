Attribute VB_Name = "Module1"
Sub DataGathering()
Dim RowHeader, ColHeader, Data As Range
Dim ParentWB As Workbook
Dim DT, RawWS, DistinctWS As Worksheet
Dim PrimRange As Range

Set ParentWB = ActiveWorkbook

On Error GoTo QuitRoutine
Set RowHeader = Application.InputBox(Prompt:="Select Row Header Range:", _
                                        Title:="Row Header Selection", _
                                        Type:=8)
currentName = RowHeader.Parent.Name
Set RawWS = ParentWB.Worksheets(currentName)

ModifyColumnHeaderSelection:
RawWS.Activate
Set ColHeader = Application.InputBox(Prompt:="Select Column Header Range:", _
                                        Title:="Column Header Selection", _
                                        Type:=8)
Set Data = Application.InputBox(Prompt:="Select Data Range:", _
                                Title:="Data Selection", _
                                Type:=8)
RawWS.Activate

Application.ScreenUpdating = False
RowHeader_add = RowHeader.Address
ColHeader_add = ColHeader.Address

testColRowCount = Range(ColHeader_add).Rows.Count
If testColRowCount > 1 Then MsgBox "Need to Develop this section"

If testColRowCount = 1 Then
    ColHeaderTitle1 = Application.InputBox(Prompt:="What would you like to name your Key Column Variable?", _
                                            Title:="Column Header Key Text", _
                                            Type:=2)
End If
On Error GoTo 0

UniqueString = "Gathered_Data_" & Format(Now(), "yyMMddhhmmss")
For Each iSheet In ParentWB.Worksheets
    If iSheet.Name = UniqueString Then
        UniqueString = UniqueString & "_2"
        ParentWB.Worksheets.Add.Name = UniqueString
        GoTo NextStep
    End If
Next

ParentWB.Worksheets.Add.Name = UniqueString
NextStep:
Set DT = ParentWB.Worksheets(UniqueString)

RH_num = Range(RowHeader_add).Columns.Count

Application.EnableEvents = False
DT.Range(DT.Cells(1, 1), DT.Cells(1, RH_num)).Value = RawWS.Range(RowHeader_add).Value

j = DT.UsedRange.Columns.Count
If testColRowCount > 1 Then
    msg_str = "This program cannot process multiple column header rows." & vbNewLine
    tmpResponse = MsgBox(msg_str & "Would you like to modify your input selection?", _
                            vbYesNo, _
                            "Modify Column Header Selection?")
    If tmpResponse = vbYes Then GoTo ModifyColumnHeaderSelection
Else:
    DT.Cells(1, j + 1) = ColHeaderTitle1
    j = DT.UsedRange.Columns.Count
End If
    
DT.Cells(1, j + 1) = "Value"
DT.Cells(1, j + 2) = "Value Comment"

'datacount = 2
'rCount = 0

Application.EnableEvents = True

data_cc = RawWS.Range(Data.Address).Columns.Count
data_rc = RawWS.Range(Data.Address).Rows.Count

val_metadata_range = RawWS.Range(RawWS.Range(RowHeader_add).Offset(1, 0), _
                        RawWS.Range(RowHeader_add).Offset(data_rc, 0)).Address
Data_ColNum = j + 1
DataCmt_ColNum = j + 2

For iColumn = 1 To data_cc
    tmp_add = RawWS.Range(Data.Address)(iColumn).Address
    tmp_irow = Range(tmp_add).Row
    tmp_icol = Range(tmp_add).Column
    val_data_range = RawWS.Range(RawWS.Cells(tmp_irow, tmp_icol), _
                                    RawWS.Cells(tmp_irow, tmp_icol).Offset(data_rc - 1, 0)).Address
    destination_rc = DT.UsedRange.Rows.Count
    start_row = destination_rc + 1
    
    destination_metadata_range = DT.Range(DT.Cells(start_row, 1), _
                                            DT.Cells(start_row, RH_num).Offset(data_rc - 1, 0)).Address
    DT.Range(destination_metadata_range) = RawWS.Range(val_metadata_range).Value
    
    destination_data_range = DT.Range(DT.Cells(start_row, Data_ColNum), _
                                        DT.Cells(start_row, Data_ColNum).Offset(data_rc - 1, 0)).Address
    DT.Range(destination_data_range) = RawWS.Range(val_data_range).Value

    destination_datacolumn_range = DT.Range(destination_data_range).Offset(0, -1).Address
    DT.Range(destination_datacolumn_range) = RawWS.Range(ColHeader_add)(iColumn).Value
    
    For iValue = 1 To data_rc
        iadd_tmp = RawWS.Range(val_data_range)(iValue).Address
        If Not RawWS.Range(iadd_tmp).Comment Is Nothing Then
            DataCmt = WorksheetFunction.Clean(Trim(RawWS.Range(iadd_tmp).Comment.Text))
            DT.Range(destination_data_range)(iValue).Offset(0, 1) = DataCmt
        End If
    Next
Next

DT.UsedRange.EntireColumn.AutoFit

MsgBox "Data Gather Complete"

QuitRoutine:
'MsgBox "The data gathering process experienced an issue"
End Sub

