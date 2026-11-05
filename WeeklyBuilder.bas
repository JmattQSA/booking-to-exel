Attribute VB_Name = "WeeklyBuilder"
Option Explicit

' ============================================================
'  Weekly Schedule Builder
'
'  Bookings_Data  (Bookings report)  uses: Date, Customer, Custom Fields
'  SNOW_Data      (ServiceNow export) uses: Reserved for, Serial number, Model
'  Business_Areas (permanent list) uses: Name, Email, Group_Descr (= Business
'                 Area on Weekly; "02B" and a trailing "Group" are removed)
'  Inventory      (permanent list) uses: Serial number, Model (fills in a
'                 missing Device New / Device Old model by serial)
'  SNOW_Old       (ServiceNow export) uses: Current serial number, Number,
'                 Model, and the customer/user name
'                 "Current serial number" looks like  5CG1188AAA - Dell Latitude 7420
'
'  Weekly output columns:
'    Date | Ticket | Old | New | Name of User | Email | Device New | Device Old
'
'    Ticket       = RITM####### or INC####### found in Custom Fields
'    Name of User = Bookings Customer (matched to SNOW "Reserved for",
'                   so the name is shown once)
'    Email        = Email from Business_Areas (matched by name); if not found,
'                   firstname.lastname@ny.frb.org built from the customer name
'    Business Area= Business_Areas Group_Descr with "02B" and a trailing
'                   "Group" removed
'    Confirm      = per-user pre-filled Microsoft Forms link (column hidden
'                   until a link is pasted on Start Here)
'    New          = SNOW Serial number
'    Device New   = SNOW Model
'    Excluded     = any booking row containing "NY EUS Mobile device Deployment"
'    Layout       = each day is split into Morning (before 12:00 PM),
'                   Afternoon (12:00-1:59 PM) and Evening (2:00 PM on).
'                   Core windows: 8:00-11:59 / 12:00-1:59 / 2:00-6:00.
'                   Anything earlier than 8 AM goes in Morning, later than
'                   6 PM in Evening; no time at all -> "Time not listed".
'    Old          = the 10-character serial BEFORE the "-"
'    Device Old   = SNOW_Old "Model" column (if none, the name AFTER the "-")
'                   Old match priority: ticket + name > ticket > name.
'                   Left blank (shaded) if no match, to fill in by hand.
'
'  Macros:
'    BuildWeeklySheet  - builds the Weekly tab
'    ClearInputData    - wipes the paste tabs for next week
'    AddButtons        - run once to put buttons on "Start Here"
' ============================================================

Private Const SH_SETUP As String = "Start Here"
Private Const SH_BOOK As String = "Bookings_Data"
Private Const SH_SNOW As String = "SNOW_Data"
Private Const SH_OUT As String = "Weekly"
Private Const SH_OLD As String = "SNOW_Old"
Private Const SH_AREAS As String = "Business_Areas"   ' permanent reference list
Private Const SH_INV As String = "Inventory"          ' permanent reference list

' Any Bookings row containing this text (in any column) is left off the schedule
Private Const EXCLUDE_TEXT As String = "NY EUS Mobile device Deployment"

Private Const NCOLS As Long = 10
Private Const EMAIL_DOMAIN As String = "ny.frb.org"
Private Const HDR_ROW As Long = 5

' Appointment record fields
Private Const F_KEY As Long = 0     ' date+time serial (sort)
Private Const F_HASTIME As Long = 1
Private Const F_TKT As Long = 2
Private Const F_NAME As Long = 3
Private Const F_SERIAL As Long = 4
Private Const F_MODEL As Long = 5
Private Const F_OLDSER As Long = 6
Private Const F_OLDDEV As Long = 7
Private Const F_AREA As Long = 8
Private Const F_EMAIL As Long = 9

' Old-device rows from SNOW_Old
Private mAreaByName As Object, mAreaByEmail As Object, mEmailByName As Object, mInvModel As Object
Private mOldCount As Long
Private mOldTkts() As String, mOldName() As String, mOldSer() As String, mOldDev() As String

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
    LoadOldDevices
    LoadReference
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
    Dim tkts As String, oSer As String, oDev As String, area As String, em As String
    Dim c As Long, rowAll As String, excl As String

    excl = LCase$(EXCLUDE_TEXT)

    lastR = LastRow(ws)
    If lastR < 2 Then Err.Raise vbObjectError + 3, , _
        "Bookings_Data is empty. Paste the Bookings report starting at A1, headers included."

    cDt = FindCol(ws, "Date", "Date time", "Datetime", "Date/Time", "Start Date", "Start Date Time", _
                  "Appointment Date", "Start")
    If cDt = 0 Then Err.Raise vbObjectError + 4, , "Bookings_Data: no 'Date' column found in row 1."
    cCust = FindCol(ws, "Customer", "Customer Name")
    If cCust = 0 Then Err.Raise vbObjectError + 5, , "Bookings_Data: no 'Customer' column found in row 1."
    cCustom = FindCol(ws, "Custom Fields", "Custom Field", "Custom Questions")

    lastC = LastCol(ws)
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim appts(1 To lastR)

    For r = 2 To lastR
        ' Skip excluded booking types (checks every column, ignores case / extra spaces)
        rowAll = ""
        For c = 1 To lastC
            rowAll = rowAll & " " & CellText(data, r, c)
        Next c
        rowAll = LCase$(rowAll)
        Do While InStr(rowAll, "  ") > 0
            rowAll = Replace(rowAll, "  ", " ")
        Loop

        If InStr(rowAll, excl) = 0 And ParseDate(data(r, cDt), dt) Then
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
                tkts = ExtractTickets(CellText(data, r, cCustom))
                MatchOld tkts, k, oSer, oDev
                ' Missing models are looked up by serial in the Inventory list
                If Len(oDev) = 0 Then oDev = InvModelFor(oSer)
                If Len(mdl) = 0 Then mdl = InvModelFor(ser)
                area = AreaFor(cust)
                em = EmailFor(cust)
                n = n + 1
                appts(n) = Array(dt, (dt - dayOnly) > 0.0001, tkts, cust, ser, mdl, oSer, oDev, area, em)
            End If
        End If
    Next r

    LoadAppointments = n
End Function

' ------------------------------------------------------------
'  PERMANENT REFERENCE LISTS (updated about twice a month)
'    Business_Areas : Name and/or Email + Business area
'    Inventory      : Serial number + Model (other columns are just stored)
'  Clear Pasted Data never touches these two tabs.
' ------------------------------------------------------------
Private Function GetOrMakeSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrMakeSheet = ws
End Function

Private Sub LoadReference()
    Dim ws As Worksheet, lastR As Long, lastC As Long, r As Long, data As Variant
    Dim cName As Long, cMail As Long, cArea As Long, cSer As Long, cMod As Long
    Dim nm As String, ml As String, ar As String, ser As String

    Set mAreaByName = CreateObject("Scripting.Dictionary")
    mAreaByName.CompareMode = vbTextCompare
    Set mAreaByEmail = CreateObject("Scripting.Dictionary")
    mAreaByEmail.CompareMode = vbTextCompare
    Set mEmailByName = CreateObject("Scripting.Dictionary")
    mEmailByName.CompareMode = vbTextCompare
    Set mInvModel = CreateObject("Scripting.Dictionary")
    mInvModel.CompareMode = vbTextCompare

    ' ---- Business_Areas ----
    Set ws = GetOrMakeSheet(SH_AREAS)
    lastR = LastRow(ws)
    If lastR >= 2 Then
        cArea = FindCol(ws, "Group_Descr", "Group Descr", "Group Description", "Business area", _
                        "Business unit", "Department", "Division", "Group")
        cName = FindCol(ws, "Name", "Name_Display", "Display name", "Full name", "Employee name", _
                        "User", "Customer name", "Customer", _
                        "Requested for", "Caller", "Assigned to")
        cMail = FindCol(ws, "Email", "E-mail", "Email address", "Email_Addr", "Email Addr", "Email ID")
        If cArea = 0 Then Err.Raise vbObjectError + 7, , _
            SH_AREAS & ": no 'Group_Descr' (business area) column found in row 1."
        If cName = 0 And cMail = 0 Then Err.Raise vbObjectError + 8, , _
            SH_AREAS & ": needs a 'Name' and/or 'Email' column in row 1."
        lastC = LastCol(ws)
        data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
        For r = 2 To lastR
            ar = CleanArea(CellText(data, r, cArea))
            ml = LCase$(CellText(data, r, cMail))
            nm = NormName(CellText(data, r, cName))
            If Len(ar) > 0 Then
                If Len(ml) > 0 Then
                    If Not mAreaByEmail.Exists(ml) Then mAreaByEmail.Add ml, ar
                End If
                If Len(nm) > 0 Then
                    If Not mAreaByName.Exists(nm) Then mAreaByName.Add nm, ar
                End If
            End If
            If Len(ml) > 0 And Len(nm) > 0 Then
                If Not mEmailByName.Exists(nm) Then mEmailByName.Add nm, ml
            End If
        Next r
    End If

    ' ---- Inventory (optional lookup: serial -> model) ----
    Set ws = GetOrMakeSheet(SH_INV)
    lastR = LastRow(ws)
    If lastR >= 2 Then
        cSer = FindCol(ws, "Serial number", "Serial", "Serial no", "Serial #")
        cMod = FindCol(ws, "Model", "Model ID", "Model name")
        If cSer > 0 And cMod > 0 Then
            lastC = LastCol(ws)
            data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
            For r = 2 To lastR
                ser = UCase$(CellText(data, r, cSer))
                If Len(ser) > 0 Then
                    If Not mInvModel.Exists(ser) Then mInvModel.Add ser, CellText(data, r, cMod)
                End If
            Next r
        End If
    End If
End Sub

' Business area text with the "02B" code removed (and any dangling separators)
Private Function CleanArea(ByVal s As String) As String
    Dim seps As String
    seps = " -:|/_.," & ChrW(8211) & ChrW(8212)
    s = Replace(s, "02B", "", , , vbTextCompare)
    s = Trim$(s)
    Do While Len(s) > 0
        If InStr(seps, Left$(s, 1)) = 0 Then Exit Do
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0
        If InStr(seps, Right$(s, 1)) = 0 Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    ' Drop a trailing "Group" word (e.g. "Supervision Group" -> "Supervision")
    If Len(s) >= 6 Then
        If LCase$(Right$(s, 5)) = "group" And InStr(seps, Mid$(s, Len(s) - 5, 1)) > 0 Then
            s = Left$(s, Len(s) - 5)
            Do While Len(s) > 0
                If InStr(seps, Right$(s, 1)) = 0 Then Exit Do
                s = Left$(s, Len(s) - 1)
            Loop
        End If
    End If
    CleanArea = s
End Function

' Email from the Business_Areas list; falls back to firstname.lastname@ny.frb.org
Private Function EmailFor(ByVal cust As String) As String
    Dim k As String
    k = NormName(cust)
    If Len(k) > 0 Then
        If mEmailByName.Exists(k) Then
            EmailFor = mEmailByName(k)
            Exit Function
        End If
    End If
    EmailFor = NameToEmail(cust)
End Function

Private Function AreaFor(ByVal cust As String) As String
    Dim em As String, k As String
    em = NameToEmail(cust)
    If Len(em) > 0 Then
        If mAreaByEmail.Exists(em) Then
            AreaFor = mAreaByEmail(em)
            Exit Function
        End If
    End If
    k = NormName(cust)
    If Len(k) > 0 Then
        If mAreaByName.Exists(k) Then AreaFor = mAreaByName(k)
    End If
End Function

Private Function InvModelFor(ByVal ser As String) As String
    ser = UCase$(Trim$(ser))
    If Len(ser) = 0 Then Exit Function
    If InStr(ser, vbLf) > 0 Then Exit Function      ' several serials stacked - skip
    If mInvModel.Exists(ser) Then InvModelFor = mInvModel(ser)
End Function

' ------------------------------------------------------------
'  SNOW_Old: current (old) device serial + device name
' ------------------------------------------------------------
Private Sub LoadOldDevices()
    Dim ws As Worksheet, lastR As Long, lastC As Long, r As Long, c As Long, data As Variant
    Dim cName As Long, cSer As Long, cDev As Long, cNum As Long, cMod As Long
    Dim rowText As String, raw As String, ser As String, dev As String, p As Long

    mOldCount = 0
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_OLD)
    On Error GoTo 0
    If ws Is Nothing Then
        ' First run after the update: create the paste tab so it's there next week
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(SH_SNOW))
        ws.Name = SH_OLD
        Exit Sub
    End If

    lastR = LastRow(ws)
    If lastR < 2 Then Exit Sub          ' nothing pasted - Old columns stay blank

    cSer = FindCol(ws, "Current serial number", "Current device serial number", _
                   "Old serial number", "Serial number", "Serial")
    If cSer = 0 Then Err.Raise vbObjectError + 6, , _
        SH_OLD & ": no 'Current serial number' column found in row 1."
    cNum = FindCol(ws, "Number", "Ticket", "Ticket number")
    cName = FindCol(ws, "Customer name", "Customer", "Name of user", "Requested for", _
                    "Reserved for", "User", "Assigned to", "Caller", "Name")
    cMod = FindCol(ws, "Model", "Model ID", "Device model", "Current model")
    cDev = FindCol(ws, "Current device", "Device name", "Device")             ' fallback only

    lastC = LastCol(ws)
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    ReDim mOldTkts(1 To lastR)
    ReDim mOldName(1 To lastR)
    ReDim mOldSer(1 To lastR)
    ReDim mOldDev(1 To lastR)

    For r = 2 To lastR
        raw = CellText(data, r, cSer)
        If Len(raw) > 0 Then
            ' Split "5CG1188AAA - Dell Latitude 7420" at the first dash
            p = InStr(raw, "-")
            If p > 0 Then
                ser = Trim$(Left$(raw, p - 1))
                dev = Trim$(Mid$(raw, p + 1))
            Else
                ser = raw
                dev = CellText(data, r, cDev)
            End If
            ' Device Old = the Model column when SNOW_Old has one (else the name after the dash)
            If cMod > 0 Then
                If Len(CellText(data, r, cMod)) > 0 Then dev = CellText(data, r, cMod)
            End If

            ' Ticket: Number column first, else look anywhere in the row
            rowText = CellText(data, r, cNum)
            If Len(ExtractTickets(rowText)) = 0 Then
                rowText = ""
                For c = 1 To lastC
                    rowText = rowText & " " & CellText(data, r, c)
                Next c
            End If

            mOldCount = mOldCount + 1
            mOldTkts(mOldCount) = Replace(ExtractTickets(rowText), ", ", ",")
            mOldName(mOldCount) = NormName(CellText(data, r, cName))
            mOldSer(mOldCount) = ser
            mOldDev(mOldCount) = dev
        End If
    Next r
End Sub

' Best SNOW_Old match: ticket + name (3) > ticket only (2) > name only (1)
Private Sub MatchOld(ByVal tkts As String, ByVal nameKey As String, _
                     ByRef ser As String, ByRef dev As String)
    Dim i As Long, score As Long, best As Long, t As Variant

    ser = ""
    dev = ""
    If mOldCount = 0 Then Exit Sub

    For i = 1 To mOldCount
        score = 0
        If Len(tkts) > 0 And Len(mOldTkts(i)) > 0 Then
            For Each t In Split(tkts, ", ")
                If InStr(1, "," & mOldTkts(i) & ",", "," & t & ",", vbTextCompare) > 0 Then
                    score = 2
                    Exit For
                End If
            Next t
        End If
        If Len(nameKey) > 0 Then
            If mOldName(i) = nameKey Then score = score + 1
        End If

        If score > best Then
            best = score
            ser = mOldSer(i)
            dev = mOldDev(i)
        ElseIf score = best And score > 0 Then
            ' Same quality match (e.g. two old devices on one ticket): stack them
            If InStr(1, vbLf & ser & vbLf, vbLf & mOldSer(i) & vbLf, vbTextCompare) = 0 Then
                ser = ser & vbLf & mOldSer(i)
                dev = dev & vbLf & mOldDev(i)
            End If
        End If
    Next i
End Sub

' Device name only: removes the serial and stray separators
Private Function CleanDevice(ByVal dev As String, ByVal ser As String) As String
    If Len(ser) > 0 Then dev = Replace(dev, ser, "", , , vbTextCompare)
    dev = Trim$(dev)
    Do While Len(dev) > 0
        If InStr(" -|,:;()[]/", Right$(dev, 1)) = 0 Then Exit Do
        dev = Left$(dev, Len(dev) - 1)
    Loop
    Do While Len(dev) > 0
        If InStr(" -|,:;()[]/", Left$(dev, 1)) = 0 Then Exit Do
        dev = Mid$(dev, 2)
    Loop
    CleanDevice = dev
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
    Dim r As Long, i As Long, matched As Long, oldMatched As Long, a As Variant
    Dim firstRow As Long, d As Long, p As Long, dayCnt As Long, pCnt As Long
    Dim dayDate As Date, pNames As Variant, pRanges As Variant, tmpl As String
    Dim navy As Long, periodFill As Long, rule As Long, muted As Long, fillIn As Long

    navy = RGB(31, 56, 100)
    periodFill = RGB(217, 225, 242)
    rule = RGB(210, 214, 222)
    muted = RGB(120, 120, 120)
    fillIn = RGB(255, 251, 235)

    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = xlTop
    ws.Columns(3).NumberFormat = "@"   ' keep serial numbers as text
    ws.Columns(4).NumberFormat = "@"

    headers = Array("Date", "Ticket", "Old", "New", "Name of User", "Email", "Business Area", "Device New", "Device Old", "Confirm")
    widths = Array(25, 16, 18, 20, 24, 32, 24, 30, 26, 14)
    For i = 0 To NCOLS - 1
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i

    For i = 1 To n
        If Len(appts(i)(F_SERIAL)) > 0 Or Len(appts(i)(F_MODEL)) > 0 Then matched = matched + 1
        If Len(appts(i)(F_OLDSER)) > 0 Then oldMatched = oldMatched + 1
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
    ws.Range("A3").Value = n & " appointments   |   " & matched & " new devices found   |   " & _
                           oldMatched & " old devices found   |   built " & _
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

    ' Pre-filled Microsoft Forms link (Start Here > Pre-filled form link)
    tmpl = GetFormTemplate()
    If Len(tmpl) > 0 Then
        If Not HasTokens(tmpl) Then
            MsgBox "The pre-filled form link on Start Here doesn't contain any placeholders (like EMAILHERE)." & _
                   vbCrLf & vbCrLf & "Confirm links were skipped. See the Start Here tab for the placeholder list.", _
                   vbExclamation, "Weekly Schedule Builder"
            tmpl = ""
        End If
    End If
    If Len(tmpl) = 0 Then ws.Columns(10).Hidden = True

    pNames = Array("TIME NOT LISTED", "MORNING", "AFTERNOON", "EVENING")
    pRanges = Array("", "8:00 AM - 11:59 AM", "12:00 PM - 1:59 PM", "2:00 PM - 6:00 PM")

    r = HDR_ROW + 1
    firstRow = r

    For d = 0 To dayCount - 1
        dayDate = weekStart + d

        dayCnt = 0
        For i = 1 To n
            If Int(appts(i)(F_KEY)) = CDbl(dayDate) Then dayCnt = dayCnt + 1
        Next i

        ' Day band
        ws.Cells(r, 1).Value = UCase$(Format$(dayDate, "dddd")) & "   " & Format$(dayDate, "mmm d") & _
                               "   (" & dayCnt & IIf(dayCnt = 1, " appointment)", " appointments)")
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS))
            .Interior.Color = navy
            .Font.Bold = True
            .Font.Color = vbWhite
            .Font.Size = 11
            .RowHeight = 22
            .VerticalAlignment = xlCenter
        End With
        r = r + 1

        If dayCnt = 0 Then
            ws.Cells(r, 1).Value = "No appointments"
            ws.Cells(r, 1).Font.Italic = True
            ws.Cells(r, 1).Font.Color = muted
            r = r + 1
        Else
            For p = 0 To 3
                pCnt = 0
                For i = 1 To n
                    If Int(appts(i)(F_KEY)) = CDbl(dayDate) And PeriodOf(appts(i)) = p Then pCnt = pCnt + 1
                Next i

                ' "Time not listed" only appears when it is actually used
                If p > 0 Or pCnt > 0 Then
                    ws.Cells(r, 1).Value = pNames(p) & IIf(p > 0, "   " & pRanges(p), "") & "   (" & pCnt & ")"
                    With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS))
                        .Interior.Color = periodFill
                        .Font.Bold = True
                        .Font.Color = navy
                        .RowHeight = 17
                        .VerticalAlignment = xlCenter
                    End With
                    ws.Cells(r, 1).IndentLevel = 1
                    r = r + 1

                    If pCnt = 0 Then
                        ws.Cells(r, 1).Value = "No appointments"
                        ws.Cells(r, 1).Font.Italic = True
                        ws.Cells(r, 1).Font.Color = muted
                        ws.Cells(r, 1).IndentLevel = 1
                        r = r + 1
                    Else
                        For i = 1 To n
                            If Int(appts(i)(F_KEY)) = CDbl(dayDate) And PeriodOf(appts(i)) = p Then
                                WriteApptRow ws, r, appts(i), muted, fillIn, rule, tmpl
                                r = r + 1
                            End If
                        Next i
                    End If
                End If
            Next p
        End If

        ws.Rows(r).RowHeight = 10    ' spacer between days
        r = r + 1
    Next d

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

' 0 = no time, 1 = Morning, 2 = Afternoon, 3 = Evening
Private Function PeriodOf(a As Variant) As Long
    Dim mins As Long
    If Not a(F_HASTIME) Then
        PeriodOf = 0
        Exit Function
    End If
    mins = CLng((a(F_KEY) - Int(a(F_KEY))) * 1440#)
    If mins < 720 Then                 ' before 12:00 PM (8:00-11:59, plus anything earlier)
        PeriodOf = 1
    ElseIf mins < 840 Then             ' 12:00 PM - 1:59 PM
        PeriodOf = 2
    Else                               ' 2:00 PM on (core 2:00-6:00, plus anything later)
        PeriodOf = 3
    End If
End Function

Private Sub WriteApptRow(ws As Worksheet, ByVal r As Long, a As Variant, _
                         ByVal muted As Long, ByVal fillIn As Long, ByVal rule As Long, _
                         ByVal tmpl As String)
    Dim url As String
    If a(F_HASTIME) Then
        ws.Cells(r, 1).Value = Format$(a(F_KEY), "ddd m/d  h:mm AM/PM")
    Else
        ws.Cells(r, 1).Value = Format$(a(F_KEY), "ddd m/d")
    End If
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).IndentLevel = 1
    ws.Cells(r, 2).Value = IIf(Len(a(F_TKT)) > 0, a(F_TKT), "-")
    ws.Cells(r, 3).Value = a(F_OLDSER)
    ws.Cells(r, 4).Value = a(F_SERIAL)
    ws.Cells(r, 5).Value = a(F_NAME)
    ws.Cells(r, 6).Value = a(F_EMAIL)
    ws.Cells(r, 7).Value = a(F_AREA)
    ws.Cells(r, 8).Value = a(F_MODEL)
    ws.Cells(r, 9).Value = a(F_OLDDEV)

    ' Confirm: per-user pre-filled form link (right-click > Copy Hyperlink to send in Teams)
    If Len(tmpl) > 0 Then
        url = BuildFormLink(tmpl, a)
        If Len(url) > 2000 Then
            ws.Cells(r, 10).Value = "Link too long"
            ws.Cells(r, 10).Font.Color = muted
        Else
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 10), Address:=url, TextToDisplay:="Open form"
            ws.Cells(r, 10).Font.Name = "Arial"
            ws.Cells(r, 10).Font.Size = 10
        End If
    End If

    ' No old-device match: shade Old / Device Old to fill in by hand
    If Len(a(F_OLDSER)) = 0 Then ws.Cells(r, 3).Interior.Color = fillIn
    If Len(a(F_OLDDEV)) = 0 Then ws.Cells(r, 9).Interior.Color = fillIn

    If Len(a(F_TKT)) = 0 Then
        ws.Cells(r, 2).Font.Color = muted
    Else
        ws.Cells(r, 2).Font.Bold = True
    End If
    If Len(a(F_SERIAL)) = 0 And Len(a(F_MODEL)) = 0 Then
        ws.Cells(r, 8).Value = "No reservation found"
        ws.Cells(r, 8).Font.Italic = True
        ws.Cells(r, 8).Font.Color = muted
    End If

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = rule
        .Weight = xlThin
    End With
End Sub

' "Smith, Jane" or "Jane Smith" -> jane.smith@ny.frb.org
' ------------------------------------------------------------
'  PRE-FILLED MICROSOFT FORMS LINK
'  Start Here holds a link copied from Forms ("Get pre-filled URL")
'  filled with placeholders. Each placeholder is swapped per user:
'    EMAILHERE      firstname.lastname@ny.frb.org
'    TICKETHERE     RITM / INC number
'    BUSINESSAREAHERE  from the Business_Areas tab (omit to let users pick)
'    RETURNINGHERE  old serial + old model   (one box)
'    RECEIVINGHERE  new serial + new model   (one box)
'    OLDSERIALHERE  OLDMODELHERE  NEWSERIALHERE  NEWMODELHERE   (separate boxes)
' ------------------------------------------------------------
Private Function GetFormTemplate() As String
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Names("FormTemplateURL").RefersToRange.Value
    On Error GoTo 0
    GetFormTemplate = Trim$(CStr(v))
    If Left$(LCase$(GetFormTemplate), 4) <> "http" Then GetFormTemplate = ""
End Function

Private Function HasTokens(ByVal t As String) As Boolean
    Dim tok As Variant
    For Each tok In Array("EMAILHERE", "TICKETHERE", "BUSINESSAREAHERE", "RETURNINGHERE", "RECEIVINGHERE", _
                          "OLDSERIALHERE", "OLDMODELHERE", "NEWSERIALHERE", "NEWMODELHERE", _
                          "OLDSERIAL%20OLDMODEL", "NEWSERIAL%20NEWMODEL", _
                          "OLDSERIAL+OLDMODEL", "NEWSERIAL+NEWMODEL")
        If InStr(1, t, tok, vbTextCompare) > 0 Then
            HasTokens = True
            Exit Function
        End If
    Next tok
End Function

Private Function BuildFormLink(ByVal tmpl As String, a As Variant) As String
    Dim email As String, tkt As String, oSer As String, oMod As String
    Dim nSer As String, nMod As String, oBoth As String, nBoth As String, u As String
    Dim area As String

    email = CStr(a(F_EMAIL))
    area = CStr(a(F_AREA))
    tkt = CStr(a(F_TKT))
    oSer = Replace(CStr(a(F_OLDSER)), vbLf, ", ")
    oMod = Replace(CStr(a(F_OLDDEV)), vbLf, ", ")
    nSer = Replace(CStr(a(F_SERIAL)), vbLf, ", ")
    nMod = Replace(CStr(a(F_MODEL)), vbLf, ", ")
    oBoth = Trim$(oSer & " " & oMod)
    nBoth = Trim$(nSer & " " & nMod)

    u = tmpl
    ' combined boxes first (also accepts the older "OLDSERIAL OLDMODEL" style placeholder)
    u = Replace(u, "OLDSERIAL%20OLDMODEL", UrlEnc(oBoth), , , vbTextCompare)
    u = Replace(u, "OLDSERIAL+OLDMODEL", UrlEnc(oBoth), , , vbTextCompare)
    u = Replace(u, "NEWSERIAL%20NEWMODEL", UrlEnc(nBoth), , , vbTextCompare)
    u = Replace(u, "NEWSERIAL+NEWMODEL", UrlEnc(nBoth), , , vbTextCompare)
    u = Replace(u, "RETURNINGHERE", UrlEnc(oBoth), , , vbTextCompare)
    u = Replace(u, "RECEIVINGHERE", UrlEnc(nBoth), , , vbTextCompare)
    u = Replace(u, "OLDSERIALHERE", UrlEnc(oSer), , , vbTextCompare)
    u = Replace(u, "OLDMODELHERE", UrlEnc(oMod), , , vbTextCompare)
    u = Replace(u, "NEWSERIALHERE", UrlEnc(nSer), , , vbTextCompare)
    u = Replace(u, "NEWMODELHERE", UrlEnc(nMod), , , vbTextCompare)
    u = Replace(u, "EMAILHERE", UrlEnc(email), , , vbTextCompare)
    u = Replace(u, "TICKETHERE", UrlEnc(tkt), , , vbTextCompare)
    u = Replace(u, "BUSINESSAREAHERE", UrlEnc(area), , , vbTextCompare)
    BuildFormLink = u
End Function

Private Function UrlEnc(ByVal s As String) As String
    Dim i As Long, ch As String, c As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        If ch Like "[A-Za-z0-9]" Or ch = "-" Or ch = "_" Or ch = "." Or ch = "~" Then
            out = out & ch
        ElseIf c >= 0 And c < 256 Then
            out = out & "%" & Right$("0" & Hex$(c), 2)
        End If
    Next i
    UrlEnc = out
End Function

Private Function NameToEmail(ByVal nm As String) As String
    Dim parts() As String, first As String, last As String, n As Long
    nm = NormName(nm)
    If Len(nm) = 0 Then Exit Function
    parts = Split(nm, " ")
    n = UBound(parts)
    ' drop a trailing suffix (Jr, Sr, II, III, IV)
    If n >= 2 Then
        Select Case parts(n)
            Case "jr", "jr.", "sr", "sr.", "ii", "iii", "iv": n = n - 1
        End Select
    End If
    If n < 1 Then Exit Function          ' single name - can't build first.last
    first = CleanNamePart(parts(0))
    last = CleanNamePart(parts(n))
    If Len(first) = 0 Or Len(last) = 0 Then Exit Function
    NameToEmail = first & "." & last & "@" & EMAIL_DOMAIN
End Function

' keep letters and hyphens only (drops apostrophes and periods)
Private Function CleanNamePart(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or ch = "-" Then out = out & ch
    Next i
    CleanNamePart = out
End Function

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
' Clears only the weekly paste tabs. Business_Areas and Inventory are permanent.
Public Sub ClearInputData()
    If MsgBox("Clear everything pasted on Bookings_Data, SNOW_Data and SNOW_Old?", _
              vbYesNo + vbQuestion, "Weekly Schedule Builder") = vbNo Then Exit Sub
    ThisWorkbook.Worksheets(SH_BOOK).Cells.Clear
    ThisWorkbook.Worksheets(SH_SNOW).Cells.Clear
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_OLD).Cells.Clear
    On Error GoTo 0
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
