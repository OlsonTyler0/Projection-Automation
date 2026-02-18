Option Explicit

' ================================================================
'  MUST-HAVE CLASS PROJECTIONS - VBA Script
'  Created by Tyler Olson In Spring 2026 for 
'  Missouri State University Graduate Programs Office.
'
'  This script is intended to be used with an imported sheet about
'  student projections, upon running this macro, it will look for
'  all course sheets (named like "ACC 711 - FA26") and update their
'  student lists based on the imported data. 
' 
'  Since this is automated, there is a few warnings:
'  1) If the projections sheet has extra lines for a student for notes, it will likely think the student is projected for an additional semester. An attempt to mitigate this was made.
'  2) If this script errors I've done my best to make sure it fails gracefully without deleting any data, but it's always a good idea to make a backup before running any new script on important data.
'
'  New students are APPENDED at the bottom.
'  Dropped students are HIGHLIGHTED in yellow — never deleted —
'  so advisors can review their notes before removing manually.
'
'  Works for any semester automatically (FA26, SU27, SP28 ...)
'  No code changes needed when rolling to a new semester.
' ================================================================


' ----------------------------------------------------------------
'  PUBLIC ENTRY POINT — bound to the "Import Projections" button 
'  You can also run this directly with ALT + F8, select "ImportProjections", and click Run.
' ----------------------------------------------------------------
Sub ImportProjections()

    Dim wsImported  As Worksheet
    Dim ws          As Worksheet
    Dim startTime   As Double
    Dim totalSheets As Long
    Dim processed   As Long
    Dim skipped     As Long
    Dim added       As Long
    Dim flagged     As Long
    Dim errList     As String

    startTime = Timer

    ' ── Locate imported-data ─────────────────────────────────
    On Error Resume Next
    Set wsImported = ThisWorkbook.Worksheets("imported-data")
    On Error GoTo 0

    If wsImported Is Nothing Then
        MsgBox "ERROR: Cannot find a sheet named 'imported-data'." & vbCrLf & vbCrLf & _
               "The sheet must be named exactly:  imported-data" & vbCrLf & _
               "(all lowercase, with a hyphen)", vbCritical, "Sheet Not Found"
        Exit Sub
    End If

    Dim lastImportRow As Long
    lastImportRow = wsImported.Cells(wsImported.Rows.Count, 1).End(xlUp).Row

    If lastImportRow < 2 Then
        MsgBox "The 'imported-data' sheet is empty.", vbCritical, "No Data"
        Exit Sub
    End If

    ' ── Find essential columns in imported-data ──────────────
    Dim colMNum As Long : colMNum = FindCol(wsImported, "M#")
    Dim colName As Long : colName = FindCol(wsImported, "Name")

    If colMNum = 0 Or colName = 0 Then
        MsgBox "ERROR: imported-data must have columns named 'M#' and 'Name' in Row 1.", _
               vbCritical, "Columns Not Found"
        Exit Sub
    End If

    ' ── Build the ordered list of semester column names ──────
    '    We need this to determine which semesters come AFTER
    '    a given course-sheet's semester (for Last Semester logic)
    Dim semesterCols() As String
    Dim semColCount As Long
    semColCount = BuildSemesterColumnList(wsImported, semesterCols)

    ' ── Count course sheets ───────────────────────────────────
    For Each ws In ThisWorkbook.Worksheets
        If IsCourseSheet(ws.Name) Then totalSheets = totalSheets + 1
    Next ws

    If totalSheets = 0 Then
        MsgBox "No course sheets found." & vbCrLf & _
               "Sheets must be named like:  ACC 711 - FA26  or  FIN 780 - SU27", _
               vbExclamation, "No Course Sheets"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' ── Process every course sheet ────────────────────────────
    For Each ws In ThisWorkbook.Worksheets

        If Not IsCourseSheet(ws.Name) Then GoTo NextSheet

        processed = processed + 1
        Application.StatusBar = "Importing " & processed & " / " & totalSheets & ":  " & ws.Name

        Dim courseCode       As String
        Dim semCode          As String
        Dim fullSemName      As String
        courseCode  = ParseCourseCode(ws.Name)
        semCode     = ParseSemesterCode(ws.Name)
        fullSemName = SemCodeToFullName(semCode)

        If fullSemName = "" Then
            errList = errList & vbCrLf & "  " & ws.Name & " — unrecognised semester code '" & semCode & "'"
            skipped = skipped + 1
            GoTo NextSheet
        End If

        Dim colSem As Long
        colSem = FindCol(wsImported, fullSemName)

        If colSem = 0 Then
            errList = errList & vbCrLf & "  " & ws.Name & " — no '" & fullSemName & "' column in imported-data"
            skipped = skipped + 1
            GoTo NextSheet
        End If

        ' Figure out which semester columns come AFTER this one
        Dim futureCols() As Long
        Dim futureCount As Long
        futureCount = GetFutureColumns(wsImported, semesterCols, semColCount, fullSemName, futureCols)

        ' ── Ensure "Must Have (Yes/No)" header exists in Col E ───
        If Trim(CStr(ws.Cells(2, 5).Value)) = "" Then
            ws.Cells(2, 3).Value = "Must Have (Yes/No)"
            ws.Cells(2, 5).Font.Bold = True
        End If

        ' ── Build dictionary: M# -> row  for existing rows 3+ ─
        Dim existing As Object
        Set existing = CreateObject("Scripting.Dictionary")

        Dim lastDataRow As Long
        lastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastDataRow < 3 Then lastDataRow = 2  ' nothing there yet

        Dim r As Long
        For r = 3 To lastDataRow
            Dim exM As String
            exM = Trim(CStr(ws.Cells(r, 1).Value))
            If exM <> "" Then existing(exM) = r
        Next r

        ' ── Build dictionary: M# -> student info  from import ─
        Dim incoming As Object
        Set incoming = CreateObject("Scripting.Dictionary")

        Dim j As Long
        For j = 2 To lastImportRow
            Dim iM As String
            iM = Trim(CStr(wsImported.Cells(j, colMNum).Value))
            If iM = "" Or iM = "0" Then GoTo NextImportRow

            Dim iCourse As String
            iCourse = Trim(CStr(wsImported.Cells(j, colSem).Value))
            If Left(iCourse, Len(courseCode)) <> courseCode Then GoTo NextImportRow

            ' Determine last semester flag
            Dim isLast As String
            isLast = IsLastSemester(wsImported, j, futureCols, futureCount)

            ' Store: Name | Course | LastSemester
            incoming(iM) = wsImported.Cells(j, colName).Value & "|" & iCourse & "|" & isLast

NextImportRow:
        Next j

        ' ── STEP 1: Update existing rows & flag dropped students
        Dim key As Variant
        For Each key In existing.Keys
            Dim existRow As Long
            existRow = existing(key)

            If incoming.Exists(key) Then
                ' Student still projected — refresh cols A, B, E only
                Dim parts() As String
                parts = Split(incoming(key), "|")
                ws.Cells(existRow, 1).Value = key        ' M#
                ws.Cells(existRow, 2).Value = parts(0)   ' Name
                ws.Cells(existRow, 3).Value = parts(2)   ' Must Have (Yes/No)

                ' Style the Last Semester cell
                StyleLastSemCell ws.Cells(existRow, 5)

                ' Clear any dropped-student highlight on this row
                ws.Range(ws.Cells(existRow, 1), ws.Cells(existRow, 5)).Interior.ColorIndex = xlNone

            Else
                ' Student no longer projected — highlight row, don't delete
                ws.Range(ws.Cells(existRow, 1), ws.Cells(existRow, 5)).Interior.Color = RGB(255, 235, 156)
                ws.Cells(existRow, 5).Value = "⚠ No longer projected"
                ws.Cells(existRow, 5).Font.Color = RGB(156, 87, 0)
                flagged = flagged + 1
            End If
        Next key

        ' ── STEP 2: Append brand-new students ────────────────
        Dim nextRow As Long
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If nextRow < 3 Then nextRow = 3
        
        Dim newStart As Long
        newStart = nextRow

        For Each key In incoming.Keys
            If Not existing.Exists(key) Then
                Dim np() As String
                np = Split(incoming(key), "|")
                ws.Cells(nextRow, 1).Value = key     ' M#
                ws.Cells(nextRow, 2).Value = np(0)   ' Name
                ws.Cells(nextRow, 3).Value = np(2)   ' Must Have (Yes/No)
                StyleLastSemCell ws.Cells(nextRow, 5)
                nextRow = nextRow + 1
                added = added + 1
            End If
        Next key

        If added > 0 Then
            ApplyOutlineBorders ws, newStart, nextRow - 1, 1, 4
        End If

NextSheet:
    Next ws

    ' ── Refresh the Dashboard summary ────────────────────────
    RefreshDashboard wsImported, colMNum, colName

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    ' ── Completion message ────────────────────────────────────
    Dim elapsed As String
    elapsed = Format(Timer - startTime, "0.0")

    Dim msg As String
    msg = "Import Complete!" & vbCrLf & vbCrLf & _
          "  Sheets processed:       " & processed & vbCrLf & _
          "  New students added:     " & added & vbCrLf & _
          "  Students flagged*:      " & flagged & vbCrLf & _
          "  Time:                   " & elapsed & " sec"

    If flagged > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "* Flagged students are highlighted yellow." & vbCrLf & _
              "  They are no longer projected for this course." & vbCrLf & _
              "  Review their Notes before deleting manually."
    End If

    If errList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "SKIPPED sheets (check names):" & errList
    End If

    MsgBox msg, vbInformation, "Import Projections"

End Sub


' ================================================================
'  DASHBOARD BUILDER / REFRESHER
'  Creates (or refreshes) the Dashboard sheet with a summary
'  table of all courses, projected enrollment, and last-semester
'  student counts.
' ================================================================
Sub RefreshDashboard(wsImported As Worksheet, colMNum As Long, colName As Long)

    Dim wsDash As Worksheet

    ' Create or clear Dashboard sheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0

    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        wsDash.Name = "Dashboard"
    Else
        wsDash.Cells.ClearContents
        wsDash.Cells.ClearFormats
    End If

    ' ── Title block ──────────────────────────────────────────
    With wsDash.Range("A1:H1")
        .Merge
        .Value = "Must-Have Class Projections — Dashboard"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlCenter
        .RowHeight = 36
    End With

    With wsDash.Range("A2:H2")
        .Merge
        .Value = "Last updated:  " & Format(Now(), "mmmm d, yyyy  hh:mm AM/PM")
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
        .Interior.Color = RGB(217, 225, 242)
        .HorizontalAlignment = xlCenter
        .RowHeight = 22
    End With

    ' ── Legend ───────────────────────────────────────────────
    Dim legRow As Long : legRow = 3
    wsDash.Cells(legRow, 1).Value = "Enrollment Colors:"
    wsDash.Cells(legRow, 1).Font.Bold = True

    wsDash.Cells(legRow, 2).Interior.Color = RGB(198, 239, 206)
    wsDash.Cells(legRow, 2).Value = "< 15  Low"

    wsDash.Cells(legRow, 3).Interior.Color = RGB(255, 235, 156)
    wsDash.Cells(legRow, 3).Value = "15–29  Medium"

    wsDash.Cells(legRow, 4).Interior.Color = RGB(255, 199, 206)
    wsDash.Cells(legRow, 4).Value = "30+  High"

    wsDash.Cells(legRow, 6).Value = "⚠ = No longer projected"
    wsDash.Cells(legRow, 7).Value = "★ = Last semester student"
    wsDash.Rows(legRow).Font.Size = 9

    ' ── Table header ─────────────────────────────────────────
    Dim hdrRow As Long : hdrRow = 5
    Dim headers As Variant
    headers = Array("Course Sheet", "Semester", "Projected Students", _
                    "Last Semester Students", "Flagged (Dropped)", "% Last Semester")

    Dim h As Integer
    For h = 0 To UBound(headers)
        With wsDash.Cells(hdrRow, h + 1)
            .Value = headers(h)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
    Next h
    wsDash.Rows(hdrRow).RowHeight = 30

    ' ── Populate rows for each course sheet ──────────────────
    Dim dataRow As Long : dataRow = hdrRow + 1
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets

        If Not IsCourseSheet(ws.Name) Then GoTo NextDashSheet

        Dim semCode  As String : semCode  = ParseSemesterCode(ws.Name)
        Dim semFull  As String : semFull  = SemCodeToFullName(semCode)
        Dim lastRow  As Long   : lastRow  = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim projected As Long
        Dim lastSem   As Long
        Dim dropped   As Long

        Dim r As Long
        For r = 3 To lastRow
            Dim mVal As String : mVal = Trim(CStr(ws.Cells(r, 1).Value))
            If mVal = "" Then GoTo NextDashRow

            Dim eVal As String : eVal = Trim(CStr(ws.Cells(r, 5).Value))

            If InStr(eVal, "No longer") > 0 Then
                dropped = dropped + 1
            Else
                projected = projected + 1
                If UCase(Left(eVal, 1)) = "Y" Then lastSem = lastSem + 1
            End If
NextDashRow:
        Next r

        ' Write row
        wsDash.Cells(dataRow, 1).Value = ws.Name
        wsDash.Cells(dataRow, 2).Value = semFull
        wsDash.Cells(dataRow, 3).Value = projected
        wsDash.Cells(dataRow, 4).Value = lastSem
        wsDash.Cells(dataRow, 5).Value = dropped
        If projected > 0 Then
            wsDash.Cells(dataRow, 6).Value = Format(lastSem / projected, "0%")
        Else
            wsDash.Cells(dataRow, 6).Value = "—"
        End If

        ' Enrollment color on projected count
        With wsDash.Cells(dataRow, 3)
            .HorizontalAlignment = xlCenter
            If projected >= 30 Then
                .Interior.Color = RGB(255, 199, 206)
            ElseIf projected >= 15 Then
                .Interior.Color = RGB(255, 235, 156)
            Else
                .Interior.Color = RGB(198, 239, 206)
            End If
        End With

        ' Dropped flag color
        If dropped > 0 Then
            wsDash.Cells(dataRow, 5).Interior.Color = RGB(255, 235, 156)
        End If

        wsDash.Cells(dataRow, 4).HorizontalAlignment = xlCenter
        wsDash.Cells(dataRow, 5).HorizontalAlignment = xlCenter
        wsDash.Cells(dataRow, 6).HorizontalAlignment = xlCenter

        ' Alternate row shading
        If (dataRow - hdrRow) Mod 2 = 0 Then
            wsDash.Rows(dataRow).Interior.Color = RGB(242, 242, 242)
        End If

        dataRow = dataRow + 1
        projected = 0 : lastSem = 0 : dropped = 0

NextDashSheet:
    Next ws

    ' ── Totals row ───────────────────────────────────────────
    Dim totRow As Long : totRow = dataRow
    wsDash.Cells(totRow, 1).Value = "TOTAL"
    wsDash.Cells(totRow, 1).Font.Bold = True
    wsDash.Cells(totRow, 3).Value = "=SUM(C" & hdrRow + 1 & ":C" & totRow - 1 & ")"
    wsDash.Cells(totRow, 4).Value = "=SUM(D" & hdrRow + 1 & ":D" & totRow - 1 & ")"
    wsDash.Cells(totRow, 5).Value = "=SUM(E" & hdrRow + 1 & ":E" & totRow - 1 & ")"
    wsDash.Range(wsDash.Cells(totRow, 1), wsDash.Cells(totRow, 6)).Font.Bold = True
    wsDash.Range(wsDash.Cells(totRow, 1), wsDash.Cells(totRow, 6)).Interior.Color = RGB(217, 225, 242)

    ' ── Instructions / button area ───────────────────────────
    Dim instrRow As Long : instrRow = totRow + 3
    With wsDash.Cells(instrRow, 1)
        .Value = "HOW TO USE:  Update the 'imported-data' sheet with new projections, then click the 'Import Projections' button to refresh all course sheets and this dashboard."
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
        .Font.Size = 9
    End With
    wsDash.Range(wsDash.Cells(instrRow, 1), wsDash.Cells(instrRow, 6)).Merge

    ' ── Column widths ─────────────────────────────────────────
    wsDash.Columns("A").ColumnWidth = 28
    wsDash.Columns("B").ColumnWidth = 16
    wsDash.Columns("C").ColumnWidth = 18
    wsDash.Columns("D").ColumnWidth = 20
    wsDash.Columns("E").ColumnWidth = 18
    wsDash.Columns("F").ColumnWidth = 16

    wsDash.Activate

End Sub


' ================================================================
'  HELPER FUNCTIONS
'  These are functions to perform specific tasks that are used pretty often, mainly to help with processing course codes.
' ================================================================

' Returns True if the sheet name looks like a course sheet
' e.g. "ACC 711 - FA26", "FIN 780 - SU27", "MGT 767 - SP28"
Private Function IsCourseSheet(sName As String) As Boolean
    If InStr(sName, " - ") = 0 Then Exit Function
    Dim parts() As String
    parts = Split(sName, " - ")
    Dim sem As String
    sem = Trim(parts(UBound(parts)))
    Dim pre As String : pre = UCase(Left(sem, 2))
    If pre <> "FA" And pre <> "SU" And pre <> "SP" And pre <> "WI" Then Exit Function
    If Not IsNumeric(Right(sem, 2)) Then Exit Function
    IsCourseSheet = True
End Function

' "ACC 711 - FA26"  ->  "ACC 711"
Private Function ParseCourseCode(sName As String) As String
    Dim pos As Long : pos = InStr(sName, " - ")
    If pos > 0 Then ParseCourseCode = Trim(Left(sName, pos - 1))
End Function

' "ACC 711 - FA26"  ->  "FA26"
' Handles typos like "SU256"  ->  "SU26"
Private Function ParseSemesterCode(sName As String) As String
    Dim parts() As String
    parts = Split(sName, " - ")
    Dim raw As String : raw = Trim(parts(UBound(parts)))
    Dim pre As String : pre = UCase(Left(raw, 2))
    Dim rawSuffix As String : rawSuffix = Mid(raw, 3)
    Dim digits As String
    Dim k As Integer
    For k = 1 To Len(rawSuffix)
        If IsNumeric(Mid(rawSuffix, k, 1)) Then
            digits = digits & Mid(rawSuffix, k, 1)
            If Len(digits) = 2 Then Exit For
        End If
    Next k
    ParseSemesterCode = pre & digits
End Function

' "FA26" -> "Fall 2026"  /  "SU27" -> "Summer 2027"  etc.
Private Function SemCodeToFullName(semCode As String) As String
    If Len(semCode) < 4 Then Exit Function
    Dim pre As String : pre = Left(semCode, 2)
    Dim yr  As String : yr  = "20" & Right(semCode, 2)
    Select Case UCase(pre)
        Case "FA" : SemCodeToFullName = "Fall " & yr
        Case "SU" : SemCodeToFullName = "Summer " & yr
        Case "SP" : SemCodeToFullName = "Spring " & yr
        Case "WI" : SemCodeToFullName = "Winter " & yr
    End Select
End Function

' Finds a column number by scanning Row 1 for a header value
Private Function FindCol(ws As Worksheet, headerName As String) As Long
    Dim lastC As Long
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastC
        If Trim(CStr(ws.Cells(1, c).Value)) = headerName Then
            FindCol = c
            Exit Function
        End If
    Next c
End Function

' Builds an ordered array of semester column names from imported-data
' Returns count of semester columns found
Private Function BuildSemesterColumnList(ws As Worksheet, _
                                          ByRef cols() As String) As Long
    Dim lastC As Long
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ReDim cols(1 To lastC)
    Dim cnt As Long
    Dim c As Long
    For c = 1 To lastC
        Dim hdr As String : hdr = Trim(CStr(ws.Cells(1, c).Value))
        If IsSemesterHeader(hdr) Then
            cnt = cnt + 1
            cols(cnt) = hdr
        End If
    Next c
    BuildSemesterColumnList = cnt
End Function

' Returns True if a string looks like a semester header ("Fall 2026" etc.)
Private Function IsSemesterHeader(s As String) As Boolean
    IsSemesterHeader = (InStr(s, "Fall ") > 0 Or InStr(s, "Summer ") > 0 Or _
                        InStr(s, "Spring ") > 0 Or InStr(s, "Winter ") > 0)
End Function

' Fills futureCols() with column numbers that come AFTER thisSemName
' in the ordered semester list. Returns count.
Private Function GetFutureColumns(ws As Worksheet, _
                                   semCols() As String, semCount As Long, _
                                   thisSemName As String, _
                                   ByRef futureCols() As Long) As Long
    ReDim futureCols(1 To semCount)
    Dim found As Boolean
    Dim cnt As Long
    Dim i As Long
    For i = 1 To semCount
        If semCols(i) = thisSemName Then found = True : GoTo NextSemCol
        If found Then
            Dim c As Long : c = FindCol(ws, semCols(i))
            If c > 0 Then
                cnt = cnt + 1
                futureCols(cnt) = c
            End If
        End If
NextSemCol:
    Next i
    GetFutureColumns = cnt
End Function

' Returns "Yes" if the student in row iRow has NO real courses
' in any semester column after the current one; otherwise "No"
Private Function IsLastSemester(ws As Worksheet, iRow As Long, _
                                  futureCols() As Long, _
                                  futureCount As Long) As String
    Dim k As Long
    For k = 1 To futureCount
        Dim v As String
        v = Trim(CStr(ws.Cells(iRow, futureCols(k)).Value))
        If v <> "" And v <> "0" Then
            Select Case LCase(v)
                Case "not projected", "taking a semester off", _
                     "not continuing with program", "nan"
                    ' not a real course — keep checking
                Case Else
                    IsLastSemester = "No"
                    Exit Function
            End Select
        End If
    Next k
    IsLastSemester = "Yes"
End Function

' Applies formatting to a Last Semester cell (col E)
Private Sub StyleLastSemCell(cell As Range)
    If UCase(Trim(CStr(cell.Value))) = "YES" Then
        cell.Interior.Color = RGB(198, 239, 206)   ' green
        cell.Font.Color = RGB(0, 97, 0)
        cell.Font.Bold = True
    ElseIf UCase(Trim(CStr(cell.Value))) = "NO" Then
        cell.Interior.ColorIndex = xlNone
        cell.Font.Color = RGB(0, 0, 0)
        cell.Font.Bold = False
    End If
End Sub

' Applies an outline (all borders) to a rectangular range of cells
Private Sub ApplyOutlineBorders(ws As Worksheet, firstRow As Long, lastRow As Long, Optional firstCol As Long = 1, Optional lastCol As Long = 4)
    If lastRow < firstRow Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = xlAutomatic
    End With
End Sub


' ================================================================
'  SETUP HELPER — Intended to be an easy setup script for the
'  dashboard this script uses in code. Making it easy to import
'  accross other sheets. 
' ================================================================
Sub SetupDashboard()

    ' Build the dashboard (blank run just to create the sheet)
    Dim wsImported As Worksheet
    On Error Resume Next
    Set wsImported = ThisWorkbook.Worksheets("imported-data")
    On Error GoTo 0

    If wsImported Is Nothing Then
        MsgBox "Please make sure the 'imported-data' sheet exists first.", vbExclamation
        Exit Sub
    End If

    Dim colM As Long : colM = FindCol(wsImported, "M#")
    Dim colN As Long : colN = FindCol(wsImported, "Name")
    RefreshDashboard wsImported, colM, colN

    ' Add the Import Projections button on the Dashboard
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")

    ' Place button in a prominent spot below the title
    Dim btn As Object
    Set btn = wsDash.Buttons.Add(10, 60, 200, 36)
    With btn
        .OnAction = "ImportProjections"
        .Caption = "Import Projections"
        .Characters.Font.Size = 13
        .Characters.Font.Bold = True
    End With

    MsgBox "Dashboard created!" & vbCrLf & vbCrLf & _
           "Click the 'Import Projections' button on the Dashboard " & _
           "whenever you update the imported-data sheet.", _
           vbInformation, "Setup Complete"

End Sub
