Sub toClampfit()
    
    Dim c As Integer
    Dim r As Integer
    Dim firstNumRow As Integer
    Dim lastCol As String
    Dim i As Integer
    Dim interval As Double
    Dim minT As Double
    Dim maxT As Double
    Dim maxTrow As Integer

    interval = 1.25

    ' Save number of rows/columns
    
    Sheets("data").Select
    c = 0
    r = 0
    firstNumRow = 1
    
    Range("A1").Select
    Do While Not (IsNumeric(ActiveCell) And IsNumeric(ActiveCell.Offset(0, 1)))
        firstNumRow = firstNumRow + 1
        ActiveCell.Offset(1, 0).Select
    Loop
    Do While Not IsEmpty(ActiveCell)
        r = r + 1
        ActiveCell.Offset(1, 0).Select
    Loop
    
    Range("A" & firstNumRow).Select
    Do While Not IsEmpty(ActiveCell)
        c = c + 1
        ActiveCell.Offset(0, 1).Select
    Loop
    r = r - 1
    c = (c - 1) / 3
    
    ' Trim data sheet
    
    For i = 1 To firstNumRow - 1
        rows(1).Delete Shift:=xlUp
    Next i
    
    ' Round times to the nearest interval, save max time & position of last row in MROUND
    
    maxT = 0
    maxTrow = 2
    
    Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "MROUND"
    Range("A2").Select
    For i = 2 To r + 1
        ActiveCell.Formula = "=MROUND(B" & i & "," & interval & ")"
        If ActiveCell.Value > maxT Then
            maxT = ActiveCell.Value
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    ' Save min time & position of last row in MROUND
    
    ActiveCell.Offset(-1, 0).Select
    maxTrow = ActiveCell.Row
    minT = Range("A2").Value
    
    ' Generate label row
    
    Range("B1").Value = "Time (s)"
    Range("C1").Select
    For i = 1 To c
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
    Range("B2").Value = c * 3 + 1
    Range("A3").Value = "AcquisitionMode=Episodic Stimulation"
    Range("A4").Value = "Comment="
    Range("A5").Value = "YTop=500,500,2"
    Range("A6").Value = "YBottom=-500,-500,-1"
    Range("A7").Value = "SweepStartTimesMS=0.000"
    Range("A8").Value = "SignalsExported=W1,W2,ratio"
    Range("A9").Value = "Signals="
    Range("B9").Select
    For i = 1 To c
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
    For i = 1 To c * 3
        ActiveCell.Formula = "=VLOOKUP($A10,data!$A$1:$" & lastCol & "$" & maxTrow & "," & i + 2 & ",TRUE)"
        ActiveCell.Offset(0, 1).Select
    Next i
    
    Range("B10:" & lastCol & "10").Select
    Selection.AutoFill Destination:=Range("B10:" & lastCol & (maxT - minT) / interval + 11)
    
    ' ' Save as .txt file
    ' ActiveWorkbook.SaveAs Filename:=Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - 5) & ".txt", _
    '     FileFormat:=xlText, CreateBackup:=False
    
End Sub