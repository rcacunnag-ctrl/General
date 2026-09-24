Attribute VB_Name = "ProgressLogger"
Option Explicit

' Records daily and weekly progress snapshots of the Jupiter schedule.
'
' Source : row LIVE_ROW of sheet "Progress Log" (formulas that read the Dashboard).
' Daily  : one row per calendar day in "Progress Log"; recording again the same day overwrites it.
' Weekly : one row per week-ending Friday in "Weekly Log"; overwritten until the week closes
'          (Sat/Sun belong to the Friday just passed), then marked "Final".
' Triggers: Workbook_Open, Workbook_BeforeSave (see ThisWorkbook) and the Dashboard button.

Private Const LOG_SHEET As String = "Progress Log"
Private Const WEEK_SHEET As String = "Weekly Log"
Private Const DASH_SHEET As String = "Dashboard"
Private Const STATUS_CELL As String = "C4"
Private Const LIVE_ROW As Long = 6
Private Const DAILY_FIRST_ROW As Long = 10
Private Const WEEK_FIRST_ROW As Long = 6
Private Const NCOLS As Long = 36          ' columns A:AJ of the snapshot
Private Const WEEK_OFFSET As Long = 2     ' snapshot starts in column C of "Weekly Log"

Private mLastError As String

Public Sub RecordProgressNow()
    If RecordSnapshot() Then
        MsgBox "Progress recorded for " & Format(Date, "ddd mm/dd/yyyy") & _
               " (daily + week ending " & Format(WeekEndingFriday(Date), "mm/dd") & ").", _
               vbInformation, "Progress Log"
    Else
        MsgBox "Progress could not be recorded: " & mLastError, vbExclamation, "Progress Log"
    End If
End Sub

Public Sub AutoRecord()
    RecordSnapshot
End Sub

Private Function RecordSnapshot() As Boolean
    Dim wsDash As Worksheet, wsLog As Worksheet, wsWeek As Worksheet
    Dim statusCell As Range, oldFormula As String, oldValue As Variant, hadFormula As Boolean
    Dim vals As Variant, r As Long, asOf As Date, wkEnd As Date, restored As Boolean

    On Error GoTo Fail
    Set wsDash = ThisWorkbook.Worksheets(DASH_SHEET)
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET)
    Set wsWeek = ThisWorkbook.Worksheets(WEEK_SHEET)
    Set statusCell = wsDash.Range(STATUS_CELL)
    asOf = Date

    ' The snapshot is always "as of today": point the status date to today, read, then restore it.
    hadFormula = statusCell.HasFormula
    If hadFormula Then oldFormula = statusCell.Formula Else oldValue = statusCell.Value
    Application.ScreenUpdating = False
    statusCell.Value = asOf
    Application.Calculate
    vals = wsLog.Range(wsLog.Cells(LIVE_ROW, 1), wsLog.Cells(LIVE_ROW, NCOLS)).Value
    vals(1, 1) = Now
    vals(1, 2) = asOf
    vals(1, 3) = CurrentUser()
    RestoreStatus statusCell, hadFormula, oldFormula, oldValue
    restored = True
    Application.Calculate

    r = FindDateRow(wsLog, 2, DAILY_FIRST_ROW, asOf)
    If r = 0 Then r = NextEmptyRow(wsLog, 2, DAILY_FIRST_ROW)
    wsLog.Range(wsLog.Cells(r, 1), wsLog.Cells(r, NCOLS)).Value = vals

    wkEnd = WeekEndingFriday(asOf)
    r = FindDateRow(wsWeek, 1, WEEK_FIRST_ROW, wkEnd)
    If r = 0 Then
        r = NextEmptyRow(wsWeek, 1, WEEK_FIRST_ROW)
        wsWeek.Cells(r, 1).Value = wkEnd
    End If
    wsWeek.Range(wsWeek.Cells(r, 1 + WEEK_OFFSET), wsWeek.Cells(r, NCOLS + WEEK_OFFSET)).Value = vals
    If asOf >= wkEnd Then
        wsWeek.Cells(r, NCOLS + WEEK_OFFSET + 1).Value = "Final"
    Else
        wsWeek.Cells(r, NCOLS + WEEK_OFFSET + 1).Value = "Week-to-date"
    End If

    Application.ScreenUpdating = True
    RecordSnapshot = True
    Exit Function

Fail:
    mLastError = Err.Number & " - " & Err.Description
    If Not restored And Not statusCell Is Nothing Then RestoreStatus statusCell, hadFormula, oldFormula, oldValue
    Application.ScreenUpdating = True
    RecordSnapshot = False
End Function

Private Function CurrentUser() As String
    On Error Resume Next
    CurrentUser = Application.UserName
    If Len(CurrentUser) = 0 Then CurrentUser = Environ$("USERNAME")
End Function

Private Sub RestoreStatus(ByVal c As Range, ByVal hadFormula As Boolean, ByVal oldFormula As String, ByVal oldValue As Variant)
    If hadFormula Then c.Formula = oldFormula Else c.Value = oldValue
End Sub

Public Function WeekEndingFriday(ByVal d As Date) As Date
    Dim wd As Long
    wd = Weekday(d, vbMonday)             ' Mon = 1 ... Sun = 7
    If wd <= 5 Then
        WeekEndingFriday = d + (5 - wd)
    Else
        WeekEndingFriday = d - (wd - 5)
    End If
End Function

Private Function FindDateRow(ByVal ws As Worksheet, ByVal col As Long, ByVal firstRow As Long, ByVal d As Date) As Long
    Dim r As Long, v As Variant
    r = firstRow
    Do While Not IsEmpty(ws.Cells(r, col).Value)
        v = ws.Cells(r, col).Value
        If IsDate(v) Or IsNumeric(v) Then
            If CLng(CDate(v)) = CLng(d) Then
                FindDateRow = r
                Exit Function
            End If
        End If
        r = r + 1
    Loop
    FindDateRow = 0
End Function

Private Function NextEmptyRow(ByVal ws As Worksheet, ByVal col As Long, ByVal firstRow As Long) As Long
    Dim r As Long
    r = firstRow
    Do While Not IsEmpty(ws.Cells(r, col).Value)
        r = r + 1
    Loop
    NextEmptyRow = r
End Function
