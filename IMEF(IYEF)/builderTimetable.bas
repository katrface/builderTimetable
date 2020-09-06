Attribute VB_Name = "Module2"
Option Explicit

Const DAYS As Integer = 6
Const LESSONS_PER_DAY As Integer = 7

Const LOWER_MARGIN = 3
Const LEFT_MARGIN = 2
Const FINISH_ZOOM = 31
Const ROW_HEIGHT = 200

Dim timesheet As String
Dim newSheet As String
Dim group As String

Public Sub showUserForm()
    UserForm1.Show
End Sub

Public Sub createTimesheet()
    timesheet = ActiveSheet.Name
    group = ActiveCell.Address
    newSheet = Sheets.Add.Name
    
    designNewSheet
    writeNewSheet
    
    ActiveWindow.ZOOM = FINISH_ZOOM
End Sub

Private Sub writeNewSheet()
    Dim colTimesheet As Integer, rowTimesheet As Integer, day As Integer
    
    With Worksheets(timesheet).Range(group)
        colTimesheet = .Column
        rowTimesheet = .Row
    End With
    
    day = 0
    rowTimesheet = rowTimesheet + 2
    For day = 0 To (DAYS - 1)
        writeDay colTimesheet, rowTimesheet, day
    Next day
End Sub

Private Sub writeDay(col As Integer, r As Integer, day As Integer)
    Dim l As Integer
    l = 0
    With Worksheets(timesheet)
        Do While l < LESSONS_PER_DAY
            Worksheets(newSheet).Cells(LOWER_MARGIN + l, LEFT_MARGIN + day).Value2 = .Cells(r, col).MergeArea.Value2
            r = r + .Cells(r, col).MergeArea.Rows.Count
            l = l + 1
        Loop
    End With
End Sub

Private Sub designNewSheet()
    With Worksheets(newSheet)
        With .Range("A1:G9")
            .Borders.LineStyle = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            With .Font
                .Name = "Times New Roman"
                .Size = 22
                '.Bold = True
            End With
        End With
        With .Range("B1:G2")
            .ColumnWidth = 100
            .RowHeight = 50
            .Font.Size = 28
            .Interior.Color = RGB(169, 208, 142)
        End With
        
        With .Range("B1:G1")
            .Merge
            .Value2 = Worksheets(timesheet).Range(group).Value2
            .Interior.Color = RGB(255, 204, 153)
        End With
        
        .Range("B2").Value2 = "Monday"
        .Range("C2").Value2 = "Tuesday"
        .Range("D2").Value2 = "Wednesday"
        .Range("E2").Value2 = "Thursday"
        .Range("F2").Value2 = "Friday"
        .Range("G2").Value2 = "Saturday"
        .Range("A1:A2").Merge
        
        With .Range("A3:A9")
            .RowHeight = ROW_HEIGHT
            .Interior.Color = RGB(169, 208, 142)
            .Orientation = 90
            .Font.Bold = True
        End With
        
        .Range("A3").Value2 = "8.30-10.00"
        .Range("A4").Value2 = "10.10-11.40"
        .Range("A5").Value2 = "11.50-13.20"
        .Range("A6").Value2 = "14.00-15.30"
        .Range("A7").Value2 = "15.40-17.10"
        .Range("A8").Value2 = "17.50-19.20"
        .Range("A9").Value2 = "19.30-21.00"
    End With
    
End Sub
