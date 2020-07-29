Attribute VB_Name = "DataConditioningToolkit"
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

Public Sub CreateDistinct_lists()
Dim ParentWB As Workbook
Dim RawWS, DistinctWS As Worksheet

Application.ScreenUpdating = False
Set ParentWB = ActiveWorkbook
Set RawWS = ActiveSheet

UniqueString = Format(Now(), "yyyyMMdd_hhmmss")
ParentWB.Worksheets.Add.Name = "Distinct_" & UniqueString
Set DistinctWS = ParentWB.Worksheets("Distinct_" & UniqueString)

RawCC = RawWS.Range("A1").End(xlToRight).Column
RawRC = RawWS.UsedRange.Rows.Count
RawDR = RawWS.Range(RawWS.Range("A1"), RawWS.Cells(RawRC, RawCC)).Address

DistinctWS.UsedRange.Clear
DistinctWS.Range(RawDR).Value = RawWS.Range(RawDR).Value

DistinctWS.Cells(1, RawCC + 1) = DistinctWS.Cells(1, RawCC).Text & "_clean"

For j = 0 To (RawCC - 2)
    tmp_ccj = RawCC - j
    DistinctWS.Columns(tmp_ccj).RemoveDuplicates Columns:=1, Header:=xlYes
    DistinctWS.Sort.SortFields.Clear
    DistinctWS.Sort.SortFields.Add Key:=Columns(tmp_ccj), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With DistinctWS.Sort
        .SetRange Columns(tmp_ccj)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    DistinctWS.Columns(tmp_ccj).RemoveDuplicates Columns:=1, Header:=xlYes
    DistinctWS.Columns(tmp_ccj).Insert (xlToRight)
    DistinctWS.Cells(1, tmp_ccj) = DistinctWS.Cells(1, tmp_ccj - 1).Text & "_clean"
Next

DistinctWS.Columns(tmp_ccj).RemoveDuplicates Columns:=1, Header:=xlYes
DistinctCC = DistinctWS.Range("A1").End(xlToRight).Column

DistinctWS.Columns.AutoFit
DistinctWS.Activate

Application.ScreenUpdating = True

MsgBox "Distinct Value Lookup Tables are ready for cleaning."

End Sub


Public Sub Clean_Lookup()
Dim ParentWB As Workbook
Dim DataInputWS, CleanedLookupWS, MergedWS As Worksheet
Dim DataInput, CleanedLookup As Range


Set ParentWB = ActiveWorkbook
Set DataInput = Application.InputBox(Prompt:="Select the ""A1"" cell in the Data Input Worksheet which produced the Distinct Worksheet", _
    Title:="Select Data Input Worksheet", Type:=8)
Set CleanedLookup = Application.InputBox(Prompt:="Select the ""A1"" cell in the Distinct Lookup Worksheet which you just cleaned.", _
    Title:="Select Distinct Lookup Worksheet", Type:=8)

Application.ScreenUpdating = False

DataInput_temp = DataInput.Parent.Name
Set DataInputWS = ParentWB.Worksheets(DataInput_temp)

CleanedLookup_temp = CleanedLookup.Parent.Name
Set CleanedLookupWS = ParentWB.Worksheets(CleanedLookup_temp)

UniqueString = Format(Now(), "yyyyMMdd_hhmmss")
Worksheets.Add(After:=CleanedLookupWS).Name = "Merged_" & UniqueString
Set MergedWS = ParentWB.Worksheets("Merged_" & UniqueString)

DataCC = DataInputWS.Range("A1").End(xlToRight).Column
DataRC = DataInputWS.Range("A1").End(xlDown).Row
DataRange = DataInputWS.Range(DataInputWS.Range("A1"), DataInputWS.Cells(DataRC, DataCC)).Address

MergedWS.Activate
i = 2
For j = 1 To DataCC
    MergedWS.Range(MergedWS.Cells(i, j), MergedWS.Cells(DataRC, j)) = _
        "=IFERROR(IF(ISBLANK(" & DataInput_temp & "!RC)," & vbNullString & ",INDEX(" & CleanedLookup_temp & "!C[" & j & "]" & _
        ",MATCH(" & DataInput_temp & "!RC," & CleanedLookup_temp & "!C[" & j - 1 & "],0)))," & DataInput_temp & "!RC)"
    MergedWS.Range(MergedWS.Cells(i, j), MergedWS.Cells(DataRC, j)).Value = _
        MergedWS.Range(MergedWS.Cells(i, j), MergedWS.Cells(DataRC, j)).Value
Next

MergedWS.Range(DataRange).Value = MergedWS.Range(DataRange).Value
MergedWS.Rows(1).Value = DataInputWS.Rows(1).Value
MergedWS.Columns.AutoFit
MergedWS.Activate

Application.ScreenUpdating = True

MsgBox "Your Original Values have been Cleaned."

End Sub




