Attribute VB_Name = "WeeklyBuilder"
Option Explicit

' ============================================================
'  Weekly Schedule Builder
'
'  Bookings_Data  (Bookings report)  uses: Date, Customer, Custom Fields
'  SNOW_Data      (ServiceNow export) uses: Reserved for, Serial number, Model
'
'  Weekly output columns:
'    Date | Ticket | Old | New | Name of User | Device New | Device Old
'
'    Ticket       = RITM####### or INC####### found in Custom Fields
'    Name of User = Bookings Customer (matched to SNOW "Reserved for",
'                   so the name is shown once)
'    New          = SNOW Serial number
'    Device New   = SNOW Model
'    Old / Device Old are left blank to fill in at the appointment
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

Private Const NCOLS As Long = 7
Private Const HDR_ROW As Long = 5

' Appointment record fields
Private Const F_KEY As Long = 0     ' date+time serial (sort)
Private Const F_HASTIME As Long = 1
Private Const F_TKT As Long = 2
Private Const F_NAME As Long = 3
Private Const F_SERIAL As Long = 4
Private Const F_MODEL As Long = 5

' ------------------------------------------------------------
'  MAIN
' ------------------------------------------------------------
Public Sub BuildWeeklySheet()
    Dim wsB As Worksheet, wsS As Worksheet, wsO As Worksheet
    Dim weekStart As Date, dayCount As Long
    Dim serials As Object, models As Object
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

    Set serials = CreateObject("Scripting.Dictionary")
    serials.CompareMode = vbTextCompare
    Set models = CreateObject("Scripting.Dictionary")
    models.CompareMode = vbTextCompare

    LoadReservations wsS, serials, models
    nAppt = LoadAppointments(wsB, weekStart, dayCount, serials, models, appts)
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
    MsgBox "Couldn't build the weekly sheet:" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Weekly Schedule Builder"
End Sub

' ------------------------------------------------------------
'  SERVICENOW: Reserved for -> Serial number / Model
' ------------------------------------------------------------
Private Sub LoadReservations(ws As Worksheet, serials As Object, models As Object)
    Dim lastR As Long, lastC As Long, r As Long, data As Variant
    Dim cRes As Long, cSer As Long, cMod As Long
    Dim k As String, ser As String, mdl As String

    lastR = LastRow(ws)
    If lastR < 2 Then Exit Sub          ' nothing pasted - schedule still builds

    cRes = FindCol(ws, "Reserved for")
    If cRes = 0 Then Err.Raise vbObjectError + 1, , _
        "SNOW_Data: no 'Reserved for' column found in row 1. Paste the export starting at A1, headers included."
    cSer = FindCol(ws, "Serial number", "Serial")
    cMod = FindCol(ws, "Model", "Model ID")
    If cSer = 0 And cMod = 0 Then Err.Raise vbObjectError + 2, , _
        "SNOW_Data: no 'Serial number' or 'Model' column found in row 1."

    lastC = LastCol(ws)
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value

    For r = 2 To lastR
        k = NormName(CellText(data, r, cRes))
        If Len(k) > 0 Then
            ser = CellText(data, r, cSer)
            mdl = CellText(data, r, cMod)
            ' Same person with more than one device: stack them in one cell
            If serials.Exists(k) Then
                serials(k) = serials(k) & vbLf & ser
                models(k) = models(k) & vbLf & mdl
            Else
                serials.Add k, ser
                models.Add k, mdl
            End If
        End If
    Next r
End Sub

' ------------------------------------------------------------
'  BOOKINGS: Date, Customer, Custom Fields
' ------------------------------------------------------------
Private Function LoadAppointments(ws As Worksheet, ByVal weekStart As Date, ByVal dayCount As Long, _
                                  serials As Object, models As Object, appts() As Variant) As Long
    Dim lastR As Long, lastC As Long, r As Long, n As Long, data As Variant
    Dim cDt As Long, cCust As Long, cCustom As Long
    Dim dt As Double, dayOnly As Double
    Dim cust As String, k As String, ser As String, mdl As String

    lastR = LastRow(ws)
    If lastR < 2 Then Err.Raise vbObjectError + 3, , _
        "Bookings_Data is empty. Paste the Bookings report starting at A1, headers included."

    cDt = FindCol(ws, "Date", "Start Date", "Appointment Date", "Start")
    If cDt = 0 Then Err.Raise vbObjectError + 4, , "Bookings_Data: no 'Date' column found in row 1."
    cCust = FindCol(ws, "Customer", "Customer Name")
    If cCust = 0 Then Err.Raise vbObjectError + 5, , "Bookings_Data: no 'Customer' column found in row 1."
    cCustom = FindCol(ws, "Custom Fields", "Custom Field", "Custom Questions")

    lastC = LastCol(ws)
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim appts(1 To lastR)

    For r = 2 To lastR
        If ParseDate(data(r, cDt), dt) Then
            dayOnly = Int(dt)
            If dayOnly >= CDbl(weekStart) And dayOnly < CDbl(weekStart) + dayCount Then
                cust = CellText(data, r, cCust)
                k = NormName(cust)
                ser = ""
                mdl = ""
                If Len(k) > 0 Then
                    If serials.Exists(k) Then
                        ser = serials(k)
                        mdl = models(k)
                    End If
                End If
                n = n + 1
                appts(n) = Array(dt, (dt - dayOnly) > 0.0001, _
                                 ExtractTickets(CellText(data, r, cCustom)), cust, ser, mdl)
            End If
        End If
    Next r

    LoadAppointments = n
End Function

' Returns every RITM+7 digits / INC+7 digits found (exact length), comma separated
Private Function ExtractTickets(ByVal s As String) As String
    Dim prefixes As Variant, p As Variant, pos As Long, j As Long
    Dim digits As String, before As String, found As String, tkt As String

    s = UCase$(s)
    prefixes = Array("RITM", "INC")
    For Each p In prefixes
        pos = InStr(1, s, p)
        Do While pos > 0
            ' prefix must not be glued to another letter (e.g. "XINC")
            If pos > 1 Then before = Mid$(s, pos - 1, 1) Else before = " "
            If Not (before Like "[A-Z]") Then
                j = pos + Len(p)
                digits = ""
                Do While j <= Len(s)
                    If Mid$(s, j, 1) Like "#" Then
                        digits = digits & Mid$(s, j, 1)
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop
                If Len(digits) = 7 Then
                    tkt = p & digits
                    If InStr(1, "," & found & ",", "," & tkt & ",") = 0 Then
                        If Len(found) > 0 Then found = found & ","
                        found = found & tkt
                    End If
                End If
            End If
            pos = InStr(pos + 1, s, p)
        Loop
    Next p
    ExtractTickets = Replace(found, ",", ", ")
End Function

' ------------------------------------------------------------
'  OUTPUT
' ------------------------------------------------------------
Private Sub RenderWeekly(ws As Worksheet, ByVal weekStart As Date, ByVal dayCount As Long, _
                         appts() As Variant, ByVal n As Long)
    Dim headers As Variant, widths As Variant
    Dim r As Long, i As Long, matched As Long, a As Variant
    Dim curDay As Double, shade As Boolean, firstRow As Long
    Dim navy As Long, dayFill As Long, rule As Long, muted As Long, fillIn As Long

    navy = RGB(31, 56, 100)
    dayFill = RGB(238, 242, 249)
    rule = RGB(210, 214, 222)
    muted = RGB(120, 120, 120)
    fillIn = RGB(255, 251, 235)

    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = xlTop
    ws.Columns(4).NumberFormat = "@"   ' keep serial numbers as text

    headers = Array("Date", "Ticket", "Old", "New", "Name of User", "Device New", "Device Old")
    widths = Array(20, 16, 18, 20, 26, 30, 26)
    For i = 0 To NCOLS - 1
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i

    For i = 1 To n
        If Len(appts(i)(F_SERIAL)) > 0 Or Len(appts(i)(F_MODEL)) > 0 Then matched = matched + 1
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
    ws.Range("A3").Value = n & " appointments   |   " & matched & " with a reserved device   |   built " & _
                           Format$(Now, "mmm d, h:mm AM/PM")
    ws.Range("A3").Font.Color = muted

    ' Headers
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
    firstRow = r
    curDay = -1

    If n = 0 Then
        ws.Cells(r, 1).Value = "No appointments this week"
        ws.Cells(r, 1).Font.Italic = True
        ws.Cells(r, 1).Font.Color = muted
        r = r + 1
    End If

    For i = 1 To n
        a = appts(i)

        ' New day: flip shading and draw a divider
        If Int(a(F_KEY)) <> curDay Then
            If curDay <> -1 Then
                With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS)).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Color = navy
                    .Weight = xlMedium
                End With
            End If
            curDay = Int(a(F_KEY))
            shade = Not shade
        End If

        If a(F_HASTIME) Then
            ws.Cells(r, 1).Value = Format$(a(F_KEY), "ddd m/d  h:mm AM/PM")
        Else
            ws.Cells(r, 1).Value = Format$(a(F_KEY), "ddd m/d")
        End If
        ws.Cells(r, 2).Value = IIf(Len(a(F_TKT)) > 0, a(F_TKT), "-")
        ws.Cells(r, 4).Value = a(F_SERIAL)
        ws.Cells(r, 5).Value = a(F_NAME)
        ws.Cells(r, 6).Value = a(F_MODEL)

        If shade Then ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS)).Interior.Color = dayFill
        ws.Cells(r, 3).Interior.Color = fillIn       ' Old        - fill in by hand
        ws.Cells(r, 7).Interior.Color = fillIn       ' Device Old - fill in by hand

        If Len(a(F_TKT)) = 0 Then
            ws.Cells(r, 2).Font.Color = muted
        Else
            ws.Cells(r, 2).Font.Bold = True
        End If
        If Len(a(F_SERIAL)) = 0 And Len(a(F_MODEL)) = 0 Then
            ws.Cells(r, 6).Value = "No reservation found"
            ws.Cells(r, 6).Font.Italic = True
            ws.Cells(r, 6).Font.Color = muted
        End If
        ws.Cells(r, 1).Font.Bold = True

        With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = rule
            .Weight = xlThin
        End With
        r = r + 1
    Next i

    With ws.Range(ws.Cells(firstRow, 1), ws.Cells(r, NCOLS))
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With

    ' View
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A" & HDR_ROW + 1).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False

    ' Print: landscape, one page wide, header on every page
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

' Accepts real Excel dates or text like "10/7/2026 9:00 AM",
' "Wed, Oct 7, 2026", "2026-10-07T14:00:00"
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

' "Smith, Jane" -> "jane smith"   (so Bookings and SNOW names line up)
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
