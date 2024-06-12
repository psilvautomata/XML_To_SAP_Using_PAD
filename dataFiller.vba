Sub Data_Filler ()

'Define variables for counter i and UL (Last Line)
Dim UL As Variant
Dim i As Variant

'Disable events and screen updating
Application.ScreenUpdating = False
Application.EnableEvents = False

'Refresh the table
Range("A2").Select
ThisWorkbook.Queries("info_files (2)").Refresh

'Count to the last row and add +1 for the header
UL = ActiveSheet.ListObjects("info_files__2").ListRows.Count + 1

'Sort date filter in ascending order
Range("info_files__2[[#Headers],[File_Date]]").Select
ActiveWorkbook.Worksheets("info_files (2)").ListObjects("info_files__2"). _
    Sort.SortFields.Clear
ActiveWorkbook.Worksheets("info_files (2)").ListObjects("info_files__2"). _
    Sort.SortFields.Add2 Key:=Range("info_files__2[[#All],[File_Date]]"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
With ActiveWorkbook.Worksheets("info_files (2)").ListObjects( _
    "info_files__2").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Wait for 5 seconds
Application.Wait (Now + TimeValue("0:00:05"))

'Loop to add missing dates within the range
i = 3

Do Until i = UL - 3
    If (Range("A" & i).Value - Range("A" & i - 1).Value) > 1 Then
        Range("A" & i).EntireRow.Insert xlShiftDown
        Range("A" & i).Value = Range("A" & i - 1).Value + 1
        UL = UL + 1
    End If
    i = i + 1
Loop

'AutoFill for 7 days ahead
Range("A" & (UL - 2)).Select '-2 because the query file is always generated on the current day
Selection.AutoFill Destination:=Range("A" & (UL - 2) & ":A" & (UL + 7)), Type:=xlFillDefault

'Re-enable events and screen updating
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
