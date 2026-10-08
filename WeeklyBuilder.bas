Attribute VB_Name = "WeeklyBuilder"
Option Explicit

' ============================================================
'  Weekly Schedule Builder
'  Combines a Microsoft Bookings export (Bookings_Data) with a
'  ServiceNow list export (SNOW_Data) into a clean, day-by-day
'  weekly sheet (Weekly).
'
'  Macros:
'    BuildWeeklySheet  - builds the Weekly tab
'    ClearInputData    - wipes the two paste tabs for next week
'    AddButtons        - run once to put buttons on "Start Here"
' ============================================================

Private Const SH_SETUP As String = "Start Here"
Private Const SH_BOOK As String = "Bookings_Data"
Private Const SH_SNOW As String = "SNOW_Data"
Private Const SH_OUT As String = "Weekly"

' Appointment record fields
Private Const F_KEY As Long = 0      ' date+time serial, used for sorting
Private Const F_TIME As Long = 1
Private Const F_CUST As Long = 2
Private Const F_SVC As Long = 3
Private Const F_STAFF As Long = 4
Private Const F_TKT As Long = 5
Private Const F_STATE As Long = 6
Private Const F_PRI As Long = 7
Private Const F_DESC As Long = 8
Private Const F_ASSIGN As Long = 9
Private Const F_NOTES As Long = 10

Private Const NCOLS As Long = 10    ' output columns A:J
Private Const HDR_ROW As Long = 5

' ------------------------------------------------------------
'  MAIN
' ------------------------------------------------------------
Public Sub BuildWeeklySheet()
    Dim wsB As Worksheet, wsS As Worksheet, wsO As Worksheet
    Dim weekStart As Date, dayCount As Long
    Dim tickets As Object, byName As Object
    Dim appts() As Variant, nAppt As Long
    Dim calcMode As Long

    On Error GoTo Fail
    Set wsB = ThisWorkbook.Worksheets(SH_BOOK)
    Set wsS = ThisWorkbook.Worksheets(SH_SNOW)
    Set wsO = ThisWorkbook.Worksheets(SH_OUT)

    weekStart = GetWeekStart()
    dayCount = IIf(IncludeWeekend(), 7, 5)

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set tickets = CreateObject("Scripting.Dictionary")
    tickets.CompareMode = vbTextCompare
    Set byName = CreateObject("Scripting.Dictionary")
    byName.CompareMode = vbTextCompare

    LoadTickets wsS, tickets, byName
    nAppt = LoadAppointments(wsB, weekStart, dayCount, tickets, byName, appts)
    If nAppt > 1 Then SortAppts appts, nAppt
    RenderWeekly wsO, weekStart, dayCount, appts, nAppt

    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    wsO.Activate
    wsO.Range("A1").Select
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    If calcMode <> 0 Then Application.Calculation = calcMode
    MsgBox "Couldn't build the weekly sheet:" & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Weekly Schedule Builder"
End Sub

' ------------------------------------------------------------
'  SERVICENOW
' ------------------------------------------------------------
Private Sub LoadTickets(ws As Worksheet, tickets As Object, byName As Object)
    Dim lastR As Long, lastC As Long, r As Long, data As Variant
    Dim cNum As Long, cDesc As Long, cState As Long, cPri As Long
    Dim cAssign As Long, cCaller As Long, cReqFor As Long
    Dim num As String, rec As Variant

    lastR = LastRow(ws)
    If lastR < 2 Then Exit Sub          ' no tickets pasted - schedule still builds

    cNum = FindCol(ws, "Number", "Ticket", "Ticket Number")
    If cNum = 0 Then Err.Raise vbObjectError + 1, , _
        "SNOW_Data: no 'Number' column found in row 1. Paste the export starting at cell A1, including headers."
    cDesc = FindCol(ws, "Short description", "Description", "Summary")
    cState = FindCol(ws, "State", "Status", "Incident state")
    cPri = FindCol(ws, "Priority")
    cAssign = FindCol(ws, "Assigned to", "Assignee")
    cCaller = FindCol(ws, "Caller", "Opened by", "Requested by")
    cReqFor = FindCol(ws, "Requested for")

    lastC = LastCol(ws)
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value

    For r = 2 To lastR
        num = UCase$(CellText(data, r, cNum))
        If Len(num) > 0 Then
            ' rec: 0 Number, 1 State, 2 Priority, 3 Short desc, 4 Assigned to
            rec = Array(num, CellText(data, r, cState), CellText(data, r, cPri), _
                        CellText(data, r, cDesc), CellText(data, r, cAssign))
            tickets(num) = rec
            AddName byName, CellText(data, r, cCaller), rec
            AddName byName, CellText(data, r, cReqFor), rec
        End If
    Next r
End Sub

' Name lookup: keep an open ticket over a closed one for the same person
Private Sub AddName(byName As Object, ByVal nm As String, rec As Variant)
    Dim k As String, cur As Variant
    k = NormName(nm)
    If Len(k) = 0 Then Exit Sub
    If Not byName.Exists(k) Then
        byName.Add k, rec
    Else
        cur = byName(k)
        If IsClosed(CStr(cur(1))) And Not IsClosed(CStr(rec(1))) Then byName(k) = rec
    End If
End Sub

Private Function IsClosed(ByVal state As String) As Boolean
    state = LCase$(state)
    IsClosed = (InStr(state, "closed") > 0 Or InStr(state, "resolved") > 0 _
                Or InStr(state, "cancel") > 0 Or InStr(state, "complete") > 0)
End Function

' ------------------------------------------------------------
'  BOOKINGS
' ------------------------------------------------------------
Private Function LoadAppointments(ws As Worksheet, ByVal weekStart As Date, ByVal dayCount As Long, _
                                  tickets As Object, byName As Object, appts() As Variant) As Long
    Dim lastR As Long, lastC As Long, r As Long, c As Long, n As Long, data As Variant
    Dim cDt As Long, cTime As Long, cCust As Long, cSvc As Long, cStaff As Long, cNotes As Long
    Dim dt As Double, tm As Double, dayOnly As Double, tmp As Double
    Dim rowText As String, tkt As String, cust As String, rec As Variant, hasTime As Boolean

    lastR = LastRow(ws)
    If lastR < 2 Then Err.Raise vbObjectError + 2, , _
        "Bookings_Data is empty. Paste the Bookings export starting at cell A1, including headers."

    cDt = FindCol(ws, "Date", "Start Date", "Appointment Date", "Start", "Date Time", "Start Date Time")
    If cDt = 0 Then Err.Raise vbObjectError + 3, , _
        "Bookings_Data: no date column found in row 1 (looked for Date / Start Date / Appointment Date / Start)."
    cTime = FindCol(ws, "Time", "Start Time", "Appointment Time")
    If cTime = cDt Then cTime = 0
    cCust = FindCol(ws, "Customer Name", "Customer", "Name", "Attendee", "Client Name")
    cSvc = FindCol(ws, "Service", "Service Name", "Service Type", "Appointment Type")
    cStaff = FindCol(ws, "Staff", "Staff Name", "Staff Member", "Staff Members", "Staff Name(s)", "Assigned Staff")
    cNotes = FindCol(ws, "Notes", "Customer Notes", "Internal Notes", "Additional Information", "Comments")

    lastC = LastCol(ws)
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim appts(1 To lastR)

    For r = 2 To lastR
        If ParseDate(data(r, cDt), dt) Then
            dayOnly = Int(dt)
            If dayOnly >= CDbl(weekStart) And dayOnly < CDbl(weekStart) + dayCount Then

                ' Time of day: Time column if present, else from the date cell
                hasTime = False
                If cTime > 0 Then
                    If ParseDate(data(r, cTime), tmp) Then tm = tmp - Int(tmp): hasTime = True
                End If
                If Not hasTime Then tm = dt - dayOnly

                ' Ticket: look for INC/RITM/REQ/etc anywhere in the row, else match by name
                rowText = ""
                For c = 1 To lastC
                    rowText = rowText & " " & CellText(data, r, c)
                Next c
                cust = CellText(data, r, cCust)
                tkt = ExtractTicket(rowText)

                If Len(tkt) > 0 Then
                    If tickets.Exists(tkt) Then
                        rec = tickets(tkt)
                    Else
                        rec = Array(tkt, "Not in SNOW export", "", "", "")
                    End If
                ElseIf byName.Exists(NormName(cust)) Then
                    rec = byName(NormName(cust))
                Else
                    rec = Array("", "", "", "", "")
                End If

                n = n + 1
                appts(n) = Array(dayOnly + tm, Format$(tm, "h:mm AM/PM"), cust, _
                                 CellText(data, r, cSvc), CellText(data, r, cStaff), _
                                 rec(0), rec(1), rec(2), rec(3), rec(4), CellText(data, r, cNotes))
            End If
        End If
    Next r

    LoadAppointments = n
End Function

' Finds the first ServiceNow-style number (e.g. INC0012345, RITM0045678)
Private Function ExtractTicket(ByVal s As String) As String
    Dim prefixes As Variant, p As Variant, pos As Long, j As Long, digits As String, ch As String
    s = UCase$(s)
    prefixes = Array("SCTASK", "RITM", "INC", "REQ", "CHG", "PRB", "TASK")
    For Each p In prefixes
        pos = InStr(1, s, p)
        Do While pos > 0
            j = pos + Len(p)
            digits = ""
            Do While j <= Len(s)
                ch = Mid$(s, j, 1)
                If ch Like "#" Then
                    digits = digits & ch
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            If Len(digits) >= 5 Then
                ExtractTicket = p & digits
                Exit Function
            End If
            pos = InStr(pos + 1, s, p)
        Loop
    Next p
End Function

' ------------------------------------------------------------
'  OUTPUT
' ------------------------------------------------------------
Private Sub RenderWeekly(ws As Worksheet, ByVal weekStart As Date, ByVal dayCount As Long, _
                         appts() As Variant, ByVal n As Long)
    Dim headers As Variant, widths As Variant
    Dim r As Long, d As Long, i As Long, k As Long, cnt As Long, linked As Long
    Dim dayDate As Date, a As Variant, zebra As Boolean
    Dim navy As Long, band As Long, stripe As Long, rule As Long, muted As Long

    navy = RGB(31, 56, 100)
    band = RGB(217, 225, 242)
    stripe = RGB(246, 248, 252)
    rule = RGB(210, 214, 222)
    muted = RGB(120, 120, 120)

    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = xlTop

    headers = Array("Time", "Customer", "Service", "Staff", "Ticket #", "State", "Priority", _
                    "Short Description", "Assigned To", "Notes")
    widths = Array(10, 22, 22, 18, 14, 16, 12, 42, 18, 32)
    For i = 0 To NCOLS - 1
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i

    For i = 1 To n
        If Len(appts(i)(F_TKT)) > 0 Then linked = linked + 1
    Next i

    ' Title block
    With ws.Range("A1")
        .Value = "Weekly Schedule"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = navy
    End With
    ws.Range("A2").Value = "Week of " & Format$(weekStart, "dddd, mmm d") & " to " & _
                           Format$(weekStart + dayCount - 1, "dddd, mmm d, yyyy")
    ws.Range("A2").Font.Size = 11
    ws.Range("A3").Value = n & " appointments   |   " & linked & " linked to a ticket   |   built " & _
                           Format$(Now, "mmm d, h:mm AM/PM")
    ws.Range("A3").Font.Color = muted

    ' Column headers
    For i = 0 To NCOLS - 1
        ws.Cells(HDR_ROW, i + 1).Value = headers(i)
    Next i
    With ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW, NCOLS))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = navy
        .RowHeight = 20
        .VerticalAlignment = xlCenter
    End With

    r = HDR_ROW + 1
    k = 1   ' appts are sorted, so walk them in order

    For d = 0 To dayCount - 1
        dayDate = weekStart + d

        cnt = 0
        For i = 1 To n
            If Int(appts(i)(F_KEY)) = CDbl(dayDate) Then cnt = cnt + 1
        Next i

        ' Day band
        ws.Cells(r, 1).Value = UCase$(Format$(dayDate, "dddd")) & "   " & Format$(dayDate, "mmm d") & _
                               "   (" & cnt & IIf(cnt = 1, " appointment)", " appointments)")
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS))
            .Interior.Color = band
            .Font.Bold = True
            .Font.Color = navy
            .RowHeight = 20
            .VerticalAlignment = xlCenter
        End With
        r = r + 1

        If cnt = 0 Then
            ws.Cells(r, 1).Value = "No appointments"
            ws.Cells(r, 1).Font.Italic = True
            ws.Cells(r, 1).Font.Color = muted
            r = r + 1
        Else
            zebra = False
            For i = 1 To n
                a = appts(i)
                If Int(a(F_KEY)) = CDbl(dayDate) Then
                    ws.Cells(r, 1).Value = a(F_TIME)
                    ws.Cells(r, 2).Value = a(F_CUST)
                    ws.Cells(r, 3).Value = a(F_SVC)
                    ws.Cells(r, 4).Value = a(F_STAFF)
                    ws.Cells(r, 5).Value = IIf(Len(a(F_TKT)) > 0, a(F_TKT), "-")
                    ws.Cells(r, 6).Value = a(F_STATE)
                    ws.Cells(r, 7).Value = a(F_PRI)
                    ws.Cells(r, 8).Value = a(F_DESC)
                    ws.Cells(r, 9).Value = a(F_ASSIGN)
                    ws.Cells(r, 10).Value = a(F_NOTES)

                    If zebra Then ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS)).Interior.Color = stripe
                    zebra = Not zebra

                    ' Highlights
                    If Left$(Trim$(a(F_PRI)), 1) = "1" Or Left$(Trim$(a(F_PRI)), 1) = "2" Then
                        ws.Cells(r, 7).Font.Bold = True
                        ws.Cells(r, 7).Font.Color = RGB(192, 0, 0)
                    End If
                    If IsClosed(a(F_STATE)) Or a(F_STATE) = "Not in SNOW export" Then
                        ws.Cells(r, 6).Font.Color = muted
                    End If
                    If Len(a(F_TKT)) = 0 Then ws.Cells(r, 5).Font.Color = muted
                    ws.Cells(r, 5).Font.Bold = (Len(a(F_TKT)) > 0)

                    With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS)).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Color = rule
                        .Weight = xlThin
                    End With
                    r = r + 1
                End If
            Next i
        End If

        ws.Rows(r).RowHeight = 8   ' spacer between days
        r = r + 1
    Next d

    ' Wrap long text columns
    ws.Range(ws.Cells(HDR_ROW + 1, 8), ws.Cells(r, 8)).WrapText = True
    ws.Range(ws.Cells(HDR_ROW + 1, 10), ws.Cells(r, 10)).WrapText = True
    ws.Range(ws.Cells(HDR_ROW + 1, 1), ws.Cells(r, 1)).HorizontalAlignment = xlLeft

    ' View: no gridlines, freeze header
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A" & HDR_ROW + 1).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False

    ' Print: landscape, one page wide, header row on every page
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, NCOLS)).Address
        .PrintTitleRows = "$" & HDR_ROW & ":$" & HDR_ROW
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterFooter = "Page &P of &N"
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

Private Sub SortAppts(appts() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, tmp As Variant
    For i = 2 To n
        tmp = appts(i)
        j = i - 1
        Do While j >= 1
            If appts(j)(F_KEY) <= tmp(F_KEY) Then Exit Do
            appts(j + 1) = appts(j)
            j = j - 1
        Loop
        appts(j + 1) = tmp
    Next i
End Sub

' ------------------------------------------------------------
'  UTILITIES
' ------------------------------------------------------------
Public Sub ClearInputData()
    If MsgBox("Clear everything pasted on Bookings_Data and SNOW_Data?", _
              vbYesNo + vbQuestion, "Weekly Schedule Builder") = vbNo Then Exit Sub
    ThisWorkbook.Worksheets(SH_BOOK).Cells.Clear
    ThisWorkbook.Worksheets(SH_SNOW).Cells.Clear
End Sub

Public Sub AddButtons()
    Dim ws As Worksheet, b As Object, x As Double, y As Double
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    x = ws.Range("E3").Left
    y = ws.Range("E3").Top
    Set b = ws.Buttons.Add(x, y, 170, 32)
    b.Caption = "Build Weekly Sheet"
    b.OnAction = "BuildWeeklySheet"
    b.Font.Bold = True
    Set b = ws.Buttons.Add(x, y + 40, 170, 26)
    b.Caption = "Clear Pasted Data"
    b.OnAction = "ClearInputData"
End Sub

Private Function GetWeekStart() As Date
    Dim v As Variant, d As Date
    On Error Resume Next
    v = ThisWorkbook.Names("WeekStart").RefersToRange.Value
    On Error GoTo 0
    If IsDate(v) Then d = CDate(v) Else d = Date
    d = DateSerial(Year(d), Month(d), Day(d))
    GetWeekStart = d - (Weekday(d, vbMonday) - 1)     ' snap to Monday
End Function

Private Function IncludeWeekend() As Boolean
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Names("IncludeWeekend").RefersToRange.Value
    On Error GoTo 0
    IncludeWeekend = (LCase$(Left$(Trim$(CStr(v)), 1)) = "y")
End Function

' Header lookup in row 1 (case-insensitive, first matching name wins)
Private Function FindCol(ws As Worksheet, ParamArray names() As Variant) As Long
    Dim lastC As Long, c As Long, i As Long
    lastC = LastCol(ws)
    For i = LBound(names) To UBound(names)
        For c = 1 To lastC
            If LCase$(Trim$(CStr(ws.Cells(1, c).Value))) = LCase$(names(i)) Then
                FindCol = c
                Exit Function
            End If
        Next c
    Next i
End Function

Private Function CellText(data As Variant, ByVal r As Long, ByVal c As Long) As String
    If c = 0 Then Exit Function
    If IsError(data(r, c)) Then Exit Function
    CellText = Trim$(CStr(data(r, c)))
End Function

' Accepts real Excel dates/times or text like "10/7/2026 9:00 AM",
' "Wed, Oct 7, 2026", "2026-10-07T14:00:00", "9:00 AM - 9:30 AM"
Private Function ParseDate(ByVal v As Variant, ByRef outD As Double) As Boolean
    Dim s As String, p As Long
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate, vbDouble, vbSingle, vbLong, vbInteger, vbCurrency
            outD = CDbl(v)
            ParseDate = True
            Exit Function
    End Select
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If Len(s) >= 16 And Mid$(s, 11, 1) = "T" Then s = Left$(s, 10) & " " & Mid$(s, 12, 8)
    p = InStr(s, " - ")
    If p > 0 Then s = Trim$(Left$(s, p - 1))
    If Not IsDate(s) Then
        p = InStr(s, ",")
        If p > 0 And p <= 10 Then s = Trim$(Mid$(s, p + 1))   ' drop leading weekday
    End If
    If IsDate(s) Then
        outD = CDbl(CDate(s))
        ParseDate = True
    End If
End Function

' "Smith, Jane" -> "jane smith"
Private Function NormName(ByVal s As String) As String
    Dim p As Long
    s = LCase$(Trim$(s))
    p = InStr(s, ",")
    If p > 0 Then s = Trim$(Mid$(s, p + 1)) & " " & Trim$(Left$(s, p - 1))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormName = s
End Function

Private Function LastRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then LastRow = f.Row
End Function

Private Function LastCol(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then LastCol = f.Column
End Function
