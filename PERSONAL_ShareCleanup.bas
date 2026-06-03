Attribute VB_Name = "PERSONAL_ShareCleanup"
Option Explicit

' ============================================================
'  Cascade Energy - Share Cleanup
'
'  Converts every cell containing a Cascade custom formula into
'  its evaluated value, flags converted cells with a fill color,
'  and logs each converted block to an audit sheet with a
'  representative formula. Operates on a SAVED COPY only; the
'  original workbook on disk is never modified.
'
'  Blocks are grouped by which custom function(s) a cell uses
'  (a "signature"), so:
'    - a uniform filled-down column collapses to one audit row
'      even when its R1C1 text varies row to row;
'    - a column whose formula changes partway down splits into
'      separate audit rows, each with a correct representative.
'
'  Does NOT require any Trust Center changes.
' ============================================================

Private Const AUDIT_SHEET_NAME As String = "Conversion Audit"

' Fill color applied to converted cells (RGB).
Private Const FLAG_COLOR_R As Long = 255
Private Const FLAG_COLOR_G As Long = 242
Private Const FLAG_COLOR_B As Long = 204   ' pale amber

' Returns the list of custom function names.
' This is the ONLY thing that changes between overhauls.
' It is populated from an external manifest (see Python tooling).
Private Function CustomFunctionNames() As Variant
    ' --- BEGIN GENERATED LIST ---
    CustomFunctionNames = Array( _
        "CascadeKWH", _
        "CascadeDegreeDays", _
        "CascadeBaseload", _
        "CascadeNormalize" _
    )
    ' --- END GENERATED LIST ---
End Function

' ============================================================
'  Entry point
' ============================================================

Public Sub ConvertCustomFormulasToValues()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim names As Variant
    Dim convertedCount As Long
    Dim resp As VbMsgBoxResult
    Dim auditWs As Worksheet
    Dim auditData As Collection
    Dim newPath As Variant

    ' Saved Application states, restored in both success and failure paths.
    Dim savedScreen As Boolean, savedEvents As Boolean
    Dim savedAlerts As Boolean, savedCalc As XlCalculation

    resp = MsgBox("This will save a COPY of the workbook, then convert all " & _
                  "Cascade custom-formula cells to values in that copy." & vbCrLf & vbCrLf & _
                  "The conversion cannot be undone. Continue?", _
                  vbExclamation + vbYesNo, "Confirm conversion")
    If resp <> vbYes Then Exit Sub

    ' --- Forced Save As before any destructive action ---
    newPath = Application.GetSaveAsFilename( _
                  InitialFileName:=SuggestedCopyName(ActiveWorkbook), _
                  FileFilter:="Excel Workbook (*.xlsx), *.xlsx," & _
                              "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
                  Title:="Save a shareable COPY before converting")
    If newPath = False Then Exit Sub   ' user cancelled

    ' Capture current Application state.
    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedAlerts = Application.DisplayAlerts
    savedCalc = Application.Calculation

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Save under the new name; the saved copy becomes the active workbook,
    ' so the original on disk is untouched.
    ActiveWorkbook.SaveAs Filename:=CStr(newPath), _
                          FileFormat:=FileFormatFromPath(CStr(newPath))
    Set wb = ActiveWorkbook

    names = CustomFunctionNames()
    convertedCount = 0
    Set auditData = New Collection

    For Each ws In wb.Worksheets
        If ws.Name <> AUDIT_SHEET_NAME Then
            ResetUsedRange ws
            convertedCount = convertedCount + _
                ProcessSheet(ws, names, auditData)
        End If
    Next ws

    ' Build / refresh the audit sheet from accumulated entries.
    Set auditWs = CreateAuditSheet(wb)
    WriteAuditSheet auditWs, auditData
    FinalizeAuditSheet auditWs

    Application.Calculation = savedCalc
    Application.DisplayAlerts = savedAlerts
    Application.EnableEvents = savedEvents
    wb.Save
    Application.ScreenUpdating = savedScreen

    MsgBox convertedCount & " cell(s) converted and flagged, across " & _
           auditData.Count & " block(s) logged to '" & AUDIT_SHEET_NAME & "'." & _
           vbCrLf & vbCrLf & "Saved copy: " & wb.FullName, _
           vbInformation, "Done"
    Exit Sub

CleanFail:
    Application.Calculation = savedCalc
    Application.DisplayAlerts = savedAlerts
    Application.EnableEvents = savedEvents
    Application.ScreenUpdating = savedScreen
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ============================================================
'  Per-sheet / per-area processing
' ============================================================

Private Function ProcessSheet(ByVal ws As Worksheet, ByVal names As Variant, _
                              ByVal auditData As Collection) As Long
    Dim formulaRng As Range
    Dim area As Range
    Dim count As Long

    On Error Resume Next
    Set formulaRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaRng Is Nothing Then Exit Function

    ' A SpecialCells result may be several non-contiguous areas.
    For Each area In formulaRng.Areas
        count = count + ProcessArea(ws, area, names, auditData)
    Next area

    ProcessSheet = count
End Function

Private Function ProcessArea(ByVal ws As Worksheet, ByVal area As Range, _
                             ByVal names As Variant, _
                             ByVal auditData As Collection) As Long
    Dim fArr As Variant            ' A1-style formulas, bulk-read
    Dim r As Long, c As Long
    Dim nRows As Long, nCols As Long
    Dim firstRowAbs As Long, firstColAbs As Long
    Dim count As Long
    Dim sig As String
    Dim flagColor As Long

    ' groups:     signature -> Range (union of cells sharing that signature)
    ' repFormula: signature -> representative A1 formula (first cell seen)
    Dim groups As Object, repFormula As Object
    Set groups = CreateObject("Scripting.Dictionary")
    Set repFormula = CreateObject("Scripting.Dictionary")

    flagColor = RGB(FLAG_COLOR_R, FLAG_COLOR_G, FLAG_COLOR_B)

    ' Bulk-read formulas into an in-memory array (one boundary crossing).
    fArr = area.Formula
    firstRowAbs = area.Row
    firstColAbs = area.Column

    If Not IsArray(fArr) Then
        ' Single-cell area: .Formula returns a scalar, not a 2-D array.
        sig = FormulaSignature(CStr(fArr), names)
        If Len(sig) > 0 Then
            AddToGroup groups, repFormula, sig, CStr(fArr), area
            count = 1
        End If
    Else
        nRows = UBound(fArr, 1)
        nCols = UBound(fArr, 2)
        For r = 1 To nRows
            For c = 1 To nCols
                sig = FormulaSignature(CStr(fArr(r, c)), names)
                If Len(sig) > 0 Then
                    AddToGroup groups, repFormula, sig, CStr(fArr(r, c)), _
                               ws.Cells(firstRowAbs + r - 1, firstColAbs + c - 1)
                    count = count + 1
                End If
            Next c
        Next r
    End If

    If count = 0 Then Exit Function

    ' Convert, flag, and log -- per signature group, per contiguous block.
    Dim sigKey As Variant
    Dim grpRange As Range, blk As Range
    For Each sigKey In groups.Keys
        Set grpRange = groups(sigKey)

        ' Collapse to values. Per-block so each cell keeps its own result.
        For Each blk In grpRange.Areas
            blk.Value = blk.Value
        Next blk

        ' Flag the whole signature group at once.
        grpRange.Interior.Color = flagColor

        ' One audit row per contiguous block within the group.
        For Each blk In grpRange.Areas
            auditData.Add AuditEntry(ws.Name, blk.Address(False, False), _
                                     blk.Cells.count, CStr(repFormula(sigKey)))
        Next blk
    Next sigKey

    ProcessArea = count
End Function

' Adds a cell to the union for its signature, recording the first-seen
' A1 formula as that signature's representative.
Private Sub AddToGroup(ByVal groups As Object, ByVal repFormula As Object, _
                       ByVal sig As String, ByVal a1Formula As String, _
                       ByVal cell As Range)
    If groups.Exists(sig) Then
        Set groups(sig) = Union(groups(sig), cell)
    Else
        groups.Add sig, cell
        repFormula.Add sig, a1Formula
    End If
End Sub

' ============================================================
'  Formula scanning (single pass: match + signature together)
' ============================================================

' Returns a signature built from the set of custom-function names the
' formula uses, sorted and joined with "+". Returns "" when the formula
' uses no custom function -- so the caller treats "" as "no match" and a
' non-empty result as both "matched" and "its grouping key", in one pass.
Private Function FormulaSignature(ByVal f As String, ByVal names As Variant) As String
    Dim i As Long
    Dim upperF As String
    Dim hits As Collection

    If Len(f) = 0 Then Exit Function

    upperF = UCase$(f)
    Set hits = New Collection
    For i = LBound(names) To UBound(names)
        If ContainsFunctionCall(upperF, UCase$(CStr(names(i)))) Then
            hits.Add UCase$(CStr(names(i)))
        End If
    Next i

    FormulaSignature = JoinSortedCollection(hits)
End Function

Private Function JoinSortedCollection(ByVal c As Collection) As String
    Dim arr() As String
    Dim i As Long, j As Long, tmp As String
    If c.count = 0 Then Exit Function

    ReDim arr(1 To c.count)
    For i = 1 To c.count
        arr(i) = c(i)
    Next i
    ' Tiny set (distinct custom funcs within one formula); simple sort.
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i) : arr(i) = arr(j) : arr(j) = tmp
            End If
        Next j
    Next i
    JoinSortedCollection = Join(arr, "+")
End Function

' A function call is the name immediately followed by "(", preceded by a
' non-identifier character (or the start of the string). Avoids matching
' a name as a substring of a longer name or inside another token.
Private Function ContainsFunctionCall(ByVal hay As String, _
                                      ByVal needle As String) As Boolean
    Dim pos As Long, startAt As Long
    Dim before As String
    startAt = 1
    Do
        pos = InStr(startAt, hay, needle & "(")
        If pos = 0 Then Exit Do
        If pos = 1 Then
            before = ""
        Else
            before = Mid$(hay, pos - 1, 1)
        End If
        If Not IsIdentifierChar(before) Then
            ContainsFunctionCall = True
            Exit Function
        End If
        startAt = pos + 1
    Loop
End Function

Private Function IsIdentifierChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then
        IsIdentifierChar = False
    Else
        IsIdentifierChar = (ch Like "[A-Za-z0-9_.]")
    End If
End Function

' ============================================================
'  Audit sheet
' ============================================================

Private Function AuditEntry(ByVal sheetName As String, ByVal addr As String, _
                            ByVal cellCount As Long, _
                            ByVal repFormula As String) As Variant
    AuditEntry = Array(sheetName, addr, cellCount, repFormula, Now)
End Function

Private Function CreateAuditSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(AUDIT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = AUDIT_SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1:E1").Value = Array("Sheet", "Range (copied as values)", _
                                    "Cell Count", _
                                    "Representative Formula (first cell of block)", _
                                    "Converted At")
    ws.Range("A1:E1").Font.Bold = True
    Set CreateAuditSheet = ws
End Function

Private Sub WriteAuditSheet(ByVal auditWs As Worksheet, ByVal auditData As Collection)
    Dim n As Long, i As Long
    Dim outArr() As Variant
    Dim entry As Variant

    n = auditData.count
    If n = 0 Then Exit Sub

    ReDim outArr(1 To n, 1 To 5)
    For i = 1 To n
        entry = auditData(i)
        outArr(i, 1) = entry(0)
        outArr(i, 2) = entry(1)
        outArr(i, 3) = entry(2)
        outArr(i, 4) = "'" & entry(3)   ' leading apostrophe: store formula as inert text
        outArr(i, 5) = entry(4)
    Next i

    ' Single bulk write of the whole audit body.
    auditWs.Range("A2").Resize(n, 5).Value = outArr
End Sub

Private Sub FinalizeAuditSheet(ByVal auditWs As Worksheet)
    auditWs.Columns("A:E").AutoFit
    auditWs.Columns("E").NumberFormat = "yyyy-mm-dd hh:mm"
    ' Cap the formula column so a long formula does not produce an unusable width.
    If auditWs.Columns("D").ColumnWidth > 80 Then
        auditWs.Columns("D").ColumnWidth = 80
    End If
End Sub

' ============================================================
'  Used-range reset
' ============================================================

' Referencing UsedRange.Address forces Excel to recompute the true used
' range, discarding phantom extent from stray formatting in distant cells.
' This is the safe (non-destructive) form. If files have severely bloated
' used ranges that this does not fix, a destructive variant that deletes
' trailing empty rows/columns is possible -- and acceptable here because
' the macro only ever operates on a saved copy.
Private Sub ResetUsedRange(ByVal ws As Worksheet)
    Dim dummy As String
    dummy = ws.UsedRange.Address
End Sub

' ============================================================
'  Save-As helpers
' ============================================================

Private Function SuggestedCopyName(ByVal wb As Workbook) As String
    Dim base As String, dotPos As Long
    base = wb.Name
    dotPos = InStrRev(base, ".")
    If dotPos > 0 Then base = Left$(base, dotPos - 1)
    SuggestedCopyName = base & "_SHARE_" & Format(Now, "yyyymmdd_hhnn")
End Function

Private Function FileFormatFromPath(ByVal p As String) As XlFileFormat
    If LCase$(Right$(p, 5)) = ".xlsm" Then
        FileFormatFromPath = xlOpenXMLWorkbookMacroEnabled
    Else
        FileFormatFromPath = xlOpenXMLWorkbook
    End If
End Function
