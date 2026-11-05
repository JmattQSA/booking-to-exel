Attribute VB_Name = "InventoryReconcile"
Option Explicit

'====================== SETTINGS (edit if needed) ======================
Private Const SN_SHEET As String = "ServiceNow"       ' sheet holding the ServiceNow export
Private Const SCAN_SHEET As String = "Scan"           ' sheet holding scanned serials (column A)
Private Const OUT_SHEET As String = "Reconciliation"  ' results sheet (rebuilt every run)
Private Const SHELF_BY_LAST_NAME As Boolean = False   ' True = shelve by last name instead
Private Const SHELVES As String = "ABC,DEF,GHI,JKL,MNO,PQR,STU,VWX,YZ"
'=======================================================================

Public Sub ReconcileInventory()
    Dim wsSN As Worksheet, wsScan As Worksheet, wsOut As Worksheet
    Dim colSerial As Long, colUser As Long, colModel As Long, colAsset As Long
    Dim scans As Object, seen As Object
    Dim r As Long, n As Long, lastScan As Long, lastSN As Long, maxRows As Long
    Dim k As String, user As String, key As Variant, color As Long
    Dim nFound As Long, nMissing As Long, nExtra As Long, nDupes As Long
    Dim out() As Variant

    On Error GoTo Fail

    Set wsSN = GetSheet(SN_SHEET)
    Set wsScan = GetSheet(SCAN_SHEET)
    If wsSN Is Nothing Or wsScan Is Nothing Then
        MsgBox "This workbook needs a sheet named """ & SN_SHEET & """ and one named """ & _
               SCAN_SHEET & """.", vbExclamation, "Inventory Reconcile"
        Exit Sub
    End If

    ' --- Find columns in the ServiceNow export by header name (row 1) ---
    colSerial = FindCol(wsSN, Array("serial number", "serial_number", "serial", "s/n"))
    colUser = FindCol(wsSN, Array("assigned to", "assigned_to", "owned by", "owned_by", "user", "name"))
    colModel = FindCol(wsSN, Array("model", "model id", "model_id", "display name"))
    colAsset = FindCol(wsSN, Array("asset tag", "asset_tag", "asset"))
    If colModel = colUser Then colModel = 0

    If colSerial = 0 Then
        MsgBox "Couldn't find a Serial Number column in row 1 of """ & SN_SHEET & """.", vbExclamation
        Exit Sub
    End If
    If colUser = 0 Then
        MsgBox "Couldn't find an Assigned To / User column in row 1 of """ & SN_SHEET & """.", vbExclamation
        Exit Sub
    End If

    ' --- Read scans (column A, any row; header row skipped automatically) ---
    Set scans = CreateObject("Scripting.Dictionary")
    scans.CompareMode = vbTextCompare
    lastScan = wsScan.Cells(wsScan.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastScan
        k = CleanSerial(CellText(wsScan.Cells(r, 1)))
        If Len(k) > 0 And InStr(1, k, "SERIAL", vbTextCompare) = 0 Then
            scans(k) = scans(k) + 1
        End If
    Next r

    ' --- Compare ServiceNow records against scans ---
    lastSN = wsSN.Cells(wsSN.Rows.Count, colSerial).End(xlUp).Row
    maxRows = Application.WorksheetFunction.Max(1, (lastSN - 1) + scans.Count)
    ReDim out(1 To maxRows, 1 To 7)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For r = 2 To lastSN
        k = CleanSerial(CellText(wsSN.Cells(r, colSerial)))
        If Len(k) > 0 And Not seen.Exists(k) Then
            seen.Add k, True
            n = n + 1
            user = Trim$(CellText(wsSN.Cells(r, colUser)))
            out(n, 1) = ShelfFor(user)
            out(n, 2) = IIf(user = "", "(no user)", user)
            out(n, 3) = k
            If colModel > 0 Then out(n, 4) = CellText(wsSN.Cells(r, colModel))
            If colAsset > 0 Then out(n, 5) = CellText(wsSN.Cells(r, colAsset))
            If scans.Exists(k) Then
                out(n, 6) = "Found"
                out(n, 7) = scans(k)
                nFound = nFound + 1
                If scans(k) > 1 Then nDupes = nDupes + 1
            Else
                out(n, 6) = "Missing - not scanned"
                out(n, 7) = 0
                nMissing = nMissing + 1
            End If
        End If
    Next r

    ' --- Anything scanned that ServiceNow doesn't know about ---
    For Each key In scans.Keys
        If Not seen.Exists(key) Then
            n = n + 1
            out(n, 1) = "Unknown"
            out(n, 2) = "(not in ServiceNow)"
            out(n, 3) = key
            out(n, 6) = "Not in ServiceNow"
            out(n, 7) = scans(key)
            nExtra = nExtra + 1
        End If
    Next key

    ' --- Build the results sheet ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wsOut = GetSheet(OUT_SHEET)
    If Not wsOut Is Nothing Then wsOut.Delete
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = OUT_SHEET
    wsOut.Range("A1:G1").Value = Array("Shelf", "User", "Serial Number", "Model", "Asset Tag", "Status", "Times Scanned")

    If n > 0 Then
        wsOut.Range("C2").Resize(n, 1).NumberFormat = "@"
        wsOut.Range("A2").Resize(n, 7).Value = out

        With wsOut.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsOut.Range("A2").Resize(n, 1), Order:=xlAscending
            .SortFields.Add Key:=wsOut.Range("B2").Resize(n, 1), Order:=xlAscending
            .SetRange wsOut.Range("A1").Resize(n + 1, 7)
            .Header = xlYes
            .Apply
        End With

        For r = 2 To n + 1
            Select Case wsOut.Cells(r, 6).Value
                Case "Found": color = RGB(198, 239, 206)            ' green
                Case "Not in ServiceNow": color = RGB(255, 235, 156) ' yellow
                Case Else: color = RGB(255, 199, 206)                ' red
            End Select
            wsOut.Cells(r, 1).Resize(1, 7).Interior.Color = color
            If wsOut.Cells(r, 7).Value > 1 Then wsOut.Cells(r, 7).Font.Bold = True
        Next r
    End If

    With wsOut.Range("A1:G1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 56, 100)
    End With
    wsOut.Range("A1").Resize(n + 1, 7).AutoFilter

    ' --- Summary block ---
    wsOut.Range("I1:J1").Value = Array("Summary", "Count")
    wsOut.Range("I2:J2").Value = Array("Found on shelf", nFound)
    wsOut.Range("I3:J3").Value = Array("Missing (not scanned)", nMissing)
    wsOut.Range("I4:J4").Value = Array("Scanned, not in ServiceNow", nExtra)
    wsOut.Range("I5:J5").Value = Array("Scanned more than once", nDupes)
    wsOut.Range("I6:J6").Value = Array("Run on", Format(Now, "yyyy-mm-dd hh:nn"))
    With wsOut.Range("I1:J1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 56, 100)
    End With
    wsOut.Range("I2:I3").Offset(0, 0).Font.Bold = True

    wsOut.Columns("A:J").AutoFit
    wsOut.Activate
    ActiveWindow.FreezePanes = False
    wsOut.Range("A2").Select
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True

    MsgBox "Done." & vbCrLf & vbCrLf & _
           "Found: " & nFound & vbCrLf & _
           "Missing: " & nMissing & vbCrLf & _
           "Not in ServiceNow: " & nExtra & vbCrLf & _
           "Double-scanned: " & nDupes, vbInformation, "Inventory Reconcile"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Inventory Reconcile"
End Sub

'============================ HELPERS ============================

Private Function GetSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

' Finds a header in row 1: exact match first, then "contains" for longer keywords.
Private Function FindCol(ws As Worksheet, names As Variant) As Long
    Dim lastC As Long, c As Long, i As Long, h As String
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = LBound(names) To UBound(names)
        For c = 1 To lastC
            h = LCase$(Trim$(CellText(ws.Cells(1, c))))
            If h = names(i) Then FindCol = c: Exit Function
        Next c
    Next i
    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 3 Then
            For c = 1 To lastC
                h = LCase$(Trim$(CellText(ws.Cells(1, c))))
                If InStr(1, h, names(i)) > 0 Then FindCol = c: Exit Function
            Next c
        End If
    Next i
End Function

' Returns a cell as text without scientific notation on numeric serials.
Private Function CellText(cell As Range) As String
    Dim v As Variant
    v = cell.Value2
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Then
        If v = Int(v) Then CellText = Format(v, "0") Else CellText = CStr(v)
    Else
        CellText = CStr(v)
    End If
End Function

' Strips spaces, tabs, line breaks and hidden characters scanners add; uppercases.
Private Function CleanSerial(ByVal s As String) As String
    Dim i As Long, code As Long, res As String
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code >= 33 And code <= 126 Then res = res & Mid$(s, i, 1)
    Next i
    CleanSerial = UCase$(res)
End Function

' Picks last (or first) name. Handles "First Last", "Last, First", and Jr/Sr/II/III.
Private Function PickName(ByVal nm As String) As String
    Dim parts() As String, i As Long, w As String, p As Long
    nm = Application.WorksheetFunction.Trim(nm)
    p = InStr(nm, ",")
    If p > 0 Then
        If SHELF_BY_LAST_NAME Then
            PickName = Trim$(Left$(nm, p - 1))
        Else
            PickName = Trim$(Mid$(nm, p + 1))
        End If
        Exit Function
    End If
    parts = Split(nm, " ")
    If Not SHELF_BY_LAST_NAME Then PickName = parts(0): Exit Function
    For i = UBound(parts) To 0 Step -1
        w = UCase$(Replace(parts(i), ".", ""))
        If w <> "JR" And w <> "SR" And w <> "II" And w <> "III" And w <> "IV" Then
            PickName = parts(i)
            Exit Function
        End If
    Next i
    PickName = parts(UBound(parts))
End Function

' Maps a user's name to a shelf label like "A-C".
Private Function ShelfFor(ByVal fullName As String) As String
    Dim nm As String, letter As String, groups() As String, i As Long
    If Len(Trim$(fullName)) = 0 Then ShelfFor = "Unassigned": Exit Function
    nm = PickName(fullName)
    letter = UCase$(Left$(nm, 1))
    If letter < "A" Or letter > "Z" Then ShelfFor = "Other": Exit Function
    groups = Split(SHELVES, ",")
    For i = 0 To UBound(groups)
        If InStr(1, groups(i), letter, vbBinaryCompare) > 0 Then
            ShelfFor = Left$(groups(i), 1) & "-" & Right$(groups(i), 1)
            Exit Function
        End If
    Next i
    ShelfFor = "Other"
End Function
