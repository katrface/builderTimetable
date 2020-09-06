Attribute VB_Name = "Module2"
Option Explicit

Const DAYS As Integer = 6
Const LESSONS_PER_DAY As Integer = 7

Const LOWER_MARGIN = 3
Const LEFT_MARGIN = 2
Const FINISH_ZOOM = 31
Const ROW_HEIGHT = 100

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
    cleanNewSheet
    
    ActiveWindow.ZOOM = FINISH_ZOOM
End Sub

Private Sub cleanNewSheet()
    Dim day As Integer
    For day = 0 To DAYS
        cleanDay day
    Next day
End Sub

Private Sub cleanDay(day As Integer)
    Dim l As Integer
    l = 0
    With Worksheets(newSheet)
        For l = 0 To LESSONS_PER_DAY
            If IsEmpty(.Cells(LOWER_MARGIN + 2 * l, LEFT_MARGIN + day)) Or IsEmpty(.Cells(LOWER_MARGIN + 2 * l + 1, LEFT_MARGIN + day)) Then
                .Range(.Cells(LOWER_MARGIN + 2 * l, LEFT_MARGIN + day), .Cells(LOWER_MARGIN + 2 * l + 1, LEFT_MARGIN + day)).Merge
            End If
        Next l
    End With
End Sub

Private Sub writeNewSheet()
    Dim colTimesheet As Integer, rowTimesheet As Integer, day As Integer
    
    With Worksheets(timesheet).Range(group)
        colTimesheet = .Column
        rowTimesheet = .Row
    End With
    
    day = 0
    For day = 0 To (DAYS - 1)
        rowTimesheet = rowTimesheet + 1
        writeDay colTimesheet, rowTimesheet, day
    Next day
End Sub

Private Sub writeDay(col As Integer, r As Integer, day As Integer)
    Dim l As Integer
    l = 0
    With Worksheets(timesheet)
        Do While l < LESSONS_PER_DAY * 2
            If .Cells(r + l, col).MergeArea.Rows.Count = 1 Then
                Worksheets(newSheet).Cells(LOWER_MARGIN + l, LEFT_MARGIN + day).Value2 = .Cells(r + l, col).MergeArea.Value2
                l = l + 1
            Else
                With Worksheets(newSheet)
                    .Range(.Cells(LOWER_MARGIN + l, LEFT_MARGIN + day), .Cells(LOWER_MARGIN + l + 1, LEFT_MARGIN + day)).Merge
                    .Cells(LOWER_MARGIN + l, LEFT_MARGIN + day).MergeArea.Value2 = Worksheets(timesheet).Cells(r + l, col).MergeArea.Value2
                End With
                l = l + 2
            End If
        Loop
    End With
    r = r + l
End Sub

Private Sub designNewSheet()
    With Worksheets(newSheet)
        With .Range("A1:G16")
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
        
        With .Range("A3:A16")
            .RowHeight = ROW_HEIGHT
            .Interior.Color = RGB(169, 208, 142)
            .Orientation = 90
            .Font.Bold = True
        End With
        
        With .Range("A3:A4")
            .Merge
            .Value2 = "8.30-10.00"
        End With
        With .Range("A5:A6")
            .Merge
            .Value2 = "10.10-11.40"
        End With
        With .Range("A7:A8")
            .Merge
            .Value2 = "11.50-13.20"
        End With
        With .Range("A9:A10")
            .Merge
            .Value2 = "14.00-15.30"
        End With
        With .Range("A11:A12")
            .Merge
            .Value2 = "15.40-17.10"
        End With
        With .Range("A13:A14")
            .Merge
            .Value2 = "17.50-19.20"
        End With
        With .Range("A15:A16")
            .Merge
            .Value2 = "19.30-21.00"
        End With
    End With
    
End Sub
