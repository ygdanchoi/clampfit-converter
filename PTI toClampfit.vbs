Sub toClampfit()
    
    Dim cols As Integer
    Dim rows As Integer
    Dim lastCol As String
    Dim i As Integer
    Dim interval As Double
    interval = 4.5
    Dim minT As Double
    Dim maxT As Double
    Dim maxTrow As Integer

    ' Save number of rows/columns, then trim 340 sheet
    
    Sheets("340").Select
    cols = Range("A1").Value
    rows = Range("A2").Value
    Range("A3").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    For i = 3 To cols + 2
        Columns(i).Select
        Selection.Delete Shift:=xlToLeft
    Next i
    Range("1:2,4:4").Select
    Selection.Delete Shift:=xlUp
    
    ' Trim 380 sheet
    
    Sheets("380").Select
    Range("A3").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    For i = 3 To cols + 2
        Columns(i).Select
        Selection.Delete Shift:=xlToLeft
    Next i
    Range("1:2,4:4").Select
    Selection.Delete Shift:=xlUp
    
    ' Trim ratio sheet
    
    Sheets("ratio").Select
    Range("A3").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    For i = 3 To cols + 2
        Columns(i).Select
        Selection.Delete Shift:=xlToLeft
    Next i
    Range("1:2,4:4").Select
    Selection.Delete Shift:=xlUp
    
    ' Consolidate 340/380/ratios into new MROUND sheet

    Sheets("340").Select
    Sheets.Add.Name = "MROUND"
    Sheets("340").Select
    Columns(1).Select
    Selection.Copy
    Sheets("MROUND").Select
    Columns(2).Select
    ActiveSheet.Paste
    
    For i = 1 To cols
        Sheets("340").Select
        Columns(i + 1).Select
        Selection.Copy
        Sheets("MROUND").Select
        Columns((i - 1) * 3 + 3).Select
        ActiveSheet.Paste
        Sheets("380").Select
        Columns(i + 1).Select
        Selection.Copy
        Sheets("MROUND").Select
        Columns((i - 1) * 3 + 4).Select
        ActiveSheet.Paste
        Sheets("ratio").Select
        Columns(i + 1).Select
        Selection.Copy
        Sheets("MROUND").Select
        Columns((i - 1) * 3 + 5).Select
        ActiveSheet.Paste
    Next i
    
    ' Round times to the nearest interval
    
    Range("A1").Value = "MROUND"
    Range("A2").Select
    For i = 2 To rows + 1
        ActiveCell.Formula = "=MROUND(B" & i & "," & interval & ")"
        ActiveCell.Offset(1, 0).Select
    Next i
    
    ' Save min/max times & position of last row in MROUND
    
    ActiveCell.Offset(-1, 0).Select
    maxT = ActiveCell.Value
    maxTrow = ActiveCell.Row
    minT = Range("A2").Value
    
    ' Generate label row
    
    Range("B1").Value = "Time (s)"
    Range("C1").Select
    For i = 1 To cols
        ActiveCell.Value = "Trace #" & i & " (W1)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = "Trace #" & i & " (W2)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = "Trace #" & i & " (ratio)"
        ActiveCell.Offset(0, 1).Select
    Next i
    
    ' Save position of last column in MROUND
    
    ActiveCell.Offset(0, -1).Select
    lastCol = Split(ActiveCell.Address, "$")(1)
    
    ' Generate header of TXT file

    Sheets.Add.Name = "TXT"
    Range("A1").Value = "ATF"
    Range("B1").Value = "1"
    Range("A2").Value = "7"
    Range("B2").Value = cols * 3 + 1
    Range("A3").Value = "AcquisitionMode=Episodic Stimulation"
    Range("A4").Value = "Comment="
    Range("A5").Value = "YTop=500,500,2"
    Range("A6").Value = "YBottom=-500,-500,-1"
    Range("A7").Value = "SweepStartTimesMS=0.000"
    Range("A8").Value = "SignalsExported=W1,W2,ratio"
    Range("A9").Value = "Signals="
    Range("B9").Select
    For i = 1 To cols
        ActiveCell.Value = "W1"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = "W2"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = "ratio"
        ActiveCell.Offset(0, 1).Select
    Next i
    Range("A10").Value = "Time (s)"
    
    ' Generate column of sequential time values
    
    Range("A11").Select
    For i = 1 To (maxT - minT) / interval + 1
        ActiveCell.Value = minT + (i - 1) * interval
        ActiveCell.Offset(1, 0).Select
    Next i
    
    ' Cross-reference time values from MROUND to generate table
    Range("B10").Select
    For i = 1 To cols * 3
        ActiveCell.Formula = "=VLOOKUP($A10,MROUND!$A$1:$" & lastCol & "$" & maxTrow & "," & i + 2 & ",TRUE)"
        ActiveCell.Offset(0, 1).Select
    Next i
    
    Range("B10:" & lastCol & "10").Select
    Selection.AutoFill Destination:=Range("B10:" & lastCol & (maxT - minT) / interval + 11)

    ' Save as .txt file
    ' ActiveWorkbook.SaveAs Filename:=Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName)) & ".txt", _
    '     FileFormat:=xlText, CreateBackup:=False
    
End Sub