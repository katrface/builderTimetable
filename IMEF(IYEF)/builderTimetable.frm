VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "BuildTimetable"
   ClientHeight    =   3744
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3732
   OleObjectBlob   =   "builderTimetable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Function searchStartPosition() As String
    Dim col As Integer, r As Integer
    col = 3
    r = 1
    With ActiveSheet
        Do While InStr(.Cells(r, col).Value2, "14.5-") = 0
            r = r + 1
        Loop
    searchStartPosition = .Cells(r, col).Address
    End With
End Function

Private Sub CommandButton1_Click()
    searchGroup
End Sub



Private Sub searchGroup()
    Dim col As Integer, r As Integer, isSearch As Boolean, gr As String
    
    With ActiveSheet.Range(searchStartPosition)
        col = .Column
        r = .Row
    End With
        
    isSearch = False
    gr = UserForm1.TextBox1.Value
    Do While (Not isSearch) And (Not IsEmpty(ActiveSheet.Cells(r, col).MergeArea))
        isSearch = checkCell(r, col, gr)
        col = col + 1
    Loop
    col = col - 1
    If isSearch Then
        ActiveSheet.Cells(r, col).Select
        createTimesheet
        UserForm1.Hide
    Else
        Dim st As String
        st = "Not found on this sheet" + vbNewLine + vbNewLine
        st = st + "Common mistake:" + vbNewLine
        st = st + "1. Invalid input." + vbNewLine
        st = st + "2. General timetable sheet is not select." + vbNewLine
        MsgBox st
    End If
End Sub

Private Function checkCell(r As Integer, col As Integer, gr As String) As Boolean
    If Not (TypeName(ActiveSheet.Cells(r, col).MergeArea.Value2) = "Variant()") Then
        If (Trim(ActiveSheet.Cells(r, col).MergeArea.Value2) = gr) Then
            checkCell = True
            Exit Function
        End If
    End If
    checkCell = False
End Function
