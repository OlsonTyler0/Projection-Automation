Option Explicit

' ================================================================
'  DEBUG / RESET SCRIPT
'  DO NOT RUN THIS SCRIPT UNLESS YOU ARE TRYING TO TEST CAPIBILITY
'  THIS WILL NUKE ALL STUDENT DATA FROM ALL COURSE SHEETS AND THE DASHBOARD
' ================================================================

' ─────────────────────────────────────────────────────────────
'  MAIN ENTRY POINT
' ─────────────────────────────────────────────────────────────
Sub ClearAllProjections()

    Dim ws As Worksheet
    Dim sheetsCleared As Long
    Dim rowsCleared As Long
    Dim response As VbMsgBoxResult

    ' Confirmation dialog
    response = MsgBox("This will clear ALL student data from all course sheets." & vbCrLf & vbCrLf & _
                      "Headers (rows 1-2) will be preserved." & vbCrLf & _
                      "The Dashboard will also be cleared." & vbCrLf & vbCrLf & _
                      "Continue?", _
                      vbYesNo + vbExclamation, "CONFIRM: Clear All Projections")

    If response <> vbYes Then
        MsgBox "Cancelled.", vbInformation, "Reset Cancelled"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Clear all course sheets
    For Each ws In ThisWorkbook.Worksheets
        If IsCourseSheet(ws.Name) Then
            Dim deletedRows As Long
            deletedRows = ClearCourseSheet(ws)
            rowsCleared = rowsCleared + deletedRows
            sheetsCleared = sheetsCleared + 1
        End If
    Next ws

    ' Clear Dashboard if it exists
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Cells.ClearContents
    ThisWorkbook.Worksheets("Dashboard").Cells.ClearFormats
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Reset Complete!" & vbCrLf & vbCrLf & _
           "  Course sheets cleared:  " & sheetsCleared & vbCrLf & _
           "  Student rows removed:   " & rowsCleared & vbCrLf & vbCrLf & _
           "  You can now re-run 'Import Projections' to test again.", _
           vbInformation, "Reset Complete"

End Sub


' ─────────────────────────────────────────────────────────────
'  HELPER: Clear one course sheet (keeps headers in rows 1-2)
' ─────────────────────────────────────────────────────────────
Private Function ClearCourseSheet(ws As Worksheet) As Long

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow <= 2 Then
        ' Only headers exist, nothing to clear
        ClearCourseSheet = 0
        Exit Function
    End If

    ' Clear rows from 3 to lastRow (student data only)
    Dim rowsToClear As Long
    rowsToClear = lastRow - 2

    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 4)).ClearContents
    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 4)).ClearFormats

    ClearCourseSheet = rowsToClear

End Function


' ─────────────────────────────────────────────────────────────
'  HELPER: Check if sheet is a course sheet
'  (Matches pattern from main script)
' ─────────────────────────────────────────────────────────────
Private Function IsCourseSheet(sName As String) As Boolean
    If InStr(sName, " - ") = 0 Then Exit Function
    Dim parts() As String
    parts = Split(sName, " - ")
    Dim sem As String
    sem = Trim(parts(UBound(parts)))
    Dim pre As String : pre = UCase(Left(sem, 2))
    If pre <> "FA" And pre <> "SU" And pre <> "SP" Then Exit Function
    If Not IsNumeric(Right(sem, 2)) Then Exit Function
    IsCourseSheet = True
End Function


' ─────────────────────────────────────────────────────────────
'  HELPER: Check if string is numeric
' ─────────────────────────────────────────────────────────────
Private Function IsNumeric(s As String) As Boolean
    On Error Resume Next
    IsNumeric = Not IsNull(s + 0)
    On Error GoTo 0
End Function
