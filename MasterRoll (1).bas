Option Explicit

'================================================================================
' MASTER ROLL MACRO
' SheetList columns:
'   A: Sheet Name
'   B: Source column letter(s) or range spec  (e.g. "C", "D:F", "C2:D50")
'   C: Target spec  => one of:
'        - explicit column letter(s) or range  => Case 1  (F1 freeze)
'        - "REST"                              => Case 3  (F2 freeze on source)
'        - "WEST"                              => Case 4  (F2 freeze on source)
'        - "NEST"                              => Case 5  (F2 freeze on source)
'        - blank                              => Case 2  (F2 freeze on source)
'   D: Column letter(s) or range to DELETE after rolling.  Blank = no delete.
'
' PRIORITY ORDER:
'   1. Process all Case-1 rows (Col C is explicit range/column)
'   2. Process all non-explicit-range rows (Cases 2-5)
'   3. Delete columns listed in Col D (if any)
'================================================================================

' ── F2 freeze modes ──────────────────────────────────────────────────────────
' F1 : freeze direct-external formulas only   (used when target is explicit)
' F2 : freeze direct-external AND indirect-external (used when source stays as OB)
' ─────────────────────────────────────────────────────────────────────────────

Public Sub MasterRoll_Run()

    Const LIST_SHEET As String = "SheetList"

    Dim wsList      As Worksheet
    Dim lastRow     As Long
    Dim i           As Long
    Dim startTime   As Double
    Dim elapsed     As Double

    Dim oldCalc     As XlCalculation
    Dim oldScreen   As Boolean
    Dim oldEvents   As Boolean
    Dim oldStatus   As Variant

    On Error GoTo FailSafe

    startTime = Timer

    ' ── Grab SheetList ──────────────────────────────────────────────────────
    On Error Resume Next
    Set wsList = ThisWorkbook.Worksheets(LIST_SHEET)
    On Error GoTo FailSafe
    If wsList Is Nothing Then
        MsgBox "Sheet '" & LIST_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If

    lastRow = LastUsedRow(wsList)
    If lastRow < 2 Then
        MsgBox "SheetList has no data rows.", vbInformation
        Exit Sub
    End If

    ' ── Performance settings ────────────────────────────────────────────────
    oldCalc   = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldStatus = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual
    Application.CutCopyMode    = False

    ' ══════════════════════════════════════════════════════════════════════════
    ' PASS 1: Case 1 – Col C is an explicit column/range
    ' ══════════════════════════════════════════════════════════════════════════
    Application.StatusBar = "Pass 1: explicit-target rows..."

    For i = 2 To lastRow
        Dim sName1   As String
        Dim srcSpec1 As String
        Dim tgtSpec1 As String

        sName1   = Trim$(CStr(wsList.Cells(i, "A").Value2))
        srcSpec1 = Trim$(CStr(wsList.Cells(i, "B").Value2))
        tgtSpec1 = Trim$(CStr(wsList.Cells(i, "C").Value2))

        If sName1 = "" Or srcSpec1 = "" Then GoTo NextRow1
        If Not IsExplicitRangeOrColumn(tgtSpec1) Then GoTo NextRow1

        Application.StatusBar = "Pass 1 – " & sName1 & " row " & i

        Dim ws1 As Worksheet
        Set ws1 = Nothing
        On Error Resume Next
        Set ws1 = ThisWorkbook.Worksheets(sName1)
        On Error GoTo FailSafe
        If ws1 Is Nothing Then GoTo NextRow1

        ProcessCase1_ExplicitTarget ws1, srcSpec1, tgtSpec1

NextRow1:
        Set ws1 = Nothing
    Next i

    ' ══════════════════════════════════════════════════════════════════════════
    ' PASS 2: Cases 2-5 – Col C is blank / NEST / WEST / REST
    ' ══════════════════════════════════════════════════════════════════════════
    Application.StatusBar = "Pass 2: structural-roll rows..."

    For i = 2 To lastRow
        Dim sName2   As String
        Dim srcSpec2 As String
        Dim tgtSpec2 As String

        sName2   = Trim$(CStr(wsList.Cells(i, "A").Value2))
        srcSpec2 = Trim$(CStr(wsList.Cells(i, "B").Value2))
        tgtSpec2 = UCase$(Trim$(CStr(wsList.Cells(i, "C").Value2)))

        If sName2 = "" Or srcSpec2 = "" Then GoTo NextRow2
        If IsExplicitRangeOrColumn(wsList.Cells(i, "C").Value2) Then GoTo NextRow2

        Application.StatusBar = "Pass 2 – " & sName2 & " row " & i

        Dim ws2 As Worksheet
        Set ws2 = Nothing
        On Error Resume Next
        Set ws2 = ThisWorkbook.Worksheets(sName2)
        On Error GoTo FailSafe
        If ws2 Is Nothing Then GoTo NextRow2

        Select Case tgtSpec2
            Case ""      : ProcessCase2_BlankTarget  ws2, srcSpec2   ' insert right
            Case "REST"  : ProcessCase3_REST          ws2, srcSpec2   ' use existing right
            Case "WEST"  : ProcessCase4_WEST          ws2, srcSpec2   ' insert left
            Case "NEST"  : ProcessCase5_NEST          ws2, srcSpec2   ' use existing left
            Case Else    ' unknown keyword – skip silently
        End Select

NextRow2:
        Set ws2 = Nothing
    Next i

    ' ══════════════════════════════════════════════════════════════════════════
    ' PASS 3: Delete columns listed in Col D
    ' ══════════════════════════════════════════════════════════════════════════
    Application.StatusBar = "Pass 3: deleting columns..."

    For i = 2 To lastRow
        Dim sName3  As String
        Dim delSpec As String

        sName3  = Trim$(CStr(wsList.Cells(i, "A").Value2))
        delSpec = Trim$(CStr(wsList.Cells(i, "D").Value2))

        If sName3 = "" Or delSpec = "" Then GoTo NextRow3

        Application.StatusBar = "Pass 3 – delete " & delSpec & " on " & sName3

        Dim ws3 As Worksheet
        Set ws3 = Nothing
        On Error Resume Next
        Set ws3 = ThisWorkbook.Worksheets(sName3)
        On Error GoTo FailSafe
        If ws3 Is Nothing Then GoTo NextRow3

        DeleteColumnSpec ws3, delSpec

NextRow3:
        Set ws3 = Nothing
    Next i

    ' ── Wrap-up ─────────────────────────────────────────────────────────────
CleanExit:
    Application.CutCopyMode    = False
    Application.StatusBar      = oldStatus
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents   = oldEvents
    Application.Calculation    = oldCalc

    elapsed = ElapsedSec(startTime)
    MsgBox "MasterRoll completed in " & Format$(elapsed, "0.00") & " seconds.", vbInformation
    Exit Sub

FailSafe:
    Application.CutCopyMode    = False
    Application.StatusBar      = oldStatus
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents   = oldEvents
    Application.Calculation    = oldCalc
    MsgBox "MasterRoll stopped: " & Err.Description, vbExclamation

End Sub


'================================================================================
' CASE 1 – Explicit target  (F1 logic: freeze direct-external only)
'
' Source range → Target range
' Target becomes the current-period opening balance.
' After copy:
'   • internal formulas in source  → kept as formula in target
'   • direct-external formulas     → frozen to value in target
'   • internal formulas already in target → preserved (not overwritten)
'================================================================================
Private Sub ProcessCase1_ExplicitTarget( _
        ByVal ws As Worksheet, _
        ByVal srcSpec As String, _
        ByVal tgtSpec As String)

    Dim srcRng As Range
    Dim tgtRng As Range

    Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    Set tgtRng = ResolveAndTrimRange(ws, tgtSpec)
    If srcRng Is Nothing Or tgtRng Is Nothing Then Exit Sub

    ' Dimension guard
    If srcRng.Rows.Count <> tgtRng.Rows.Count Or _
       srcRng.Columns.Count <> tgtRng.Columns.Count Then Exit Sub

    ws.DisplayPageBreaks = False

    ' Build level map for F1: only level-1 (direct external) cells get frozen
    Dim levelMap As Object
    Set levelMap = BuildFormulaLevels_F1(ws)   ' returns dict: ADDR -> 1 if direct-external

    ' Step A: Remember internal formulas already in target (preserve them)
    Dim keepFormula As Object
    Dim keepTarget  As Object
    Set keepFormula = CreateObject("Scripting.Dictionary")
    Set keepTarget  = CreateObject("Scripting.Dictionary")

    Dim fc As Range, c As Range
    On Error Resume Next
    Set fc = tgtRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not fc Is Nothing Then
        For Each c In fc.Cells
            Dim fKey1 As String
            fKey1 = c.Formula
            If Not IsDirectExternalFormula(fKey1, ws.Name) Then
                ' internal formula in target – preserve it
                Dim rKey1 As String
                rKey1 = RelKey(c, tgtRng)
                keepTarget(rKey1)  = True
                keepFormula(rKey1) = c.FormulaR1C1
            End If
        Next c
    End If

    ' Step B: Copy values en-masse
    tgtRng.Value2 = srcRng.Value2

    ' Step C: Process source formulas
    Dim sf As Range
    On Error Resume Next
    Set sf = srcRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not sf Is Nothing Then
        Dim tgtCell As Range
        For Each c In sf.Cells
            Dim rKey2 As String
            rKey2 = RelKey(c, srcRng)

            ' Skip if target has its own internal formula to preserve
            If keepTarget.Exists(rKey2) Then GoTo NextSrcCell

            Set tgtCell = tgtRng.Cells( _
                c.Row - srcRng.Row + 1, _
                c.Column - srcRng.Column + 1)

            Dim srcAddr As String
            srcAddr = UCase$(c.Address(False, False))

            ' F1: direct-external → freeze to value; internal → keep formula
            If levelMap.Exists(srcAddr) Then
                ' direct-external: freeze
                tgtCell.Value2 = c.Value2
            Else
                ' internal: transplant formula (shift R1C1 offsets)
                tgtCell.FormulaR1C1 = c.FormulaR1C1
            End If

NextSrcCell:
        Next c
    End If

    ' Step D: Restore preserved target-internal formulas
    Dim k As Variant
    Dim parts() As String
    For Each k In keepFormula.Keys
        parts = Split(CStr(k), "|")
        tgtRng.Cells(CLng(parts(0)), CLng(parts(1))).FormulaR1C1 = keepFormula(k)
    Next k

    ' Step E: Clear comments/notes from target (clean opening balance)
    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' CASE 2 – Blank Col C  (insert new columns to the RIGHT of source)
'
' Step-by-step:
'   1. Capture source formulas/values BEFORE the column insert
'      (insert shifts column indices, so we must read first)
'   2. Insert blank columns immediately to the right of source
'   3. Copy formats from (now-shifted) source → new target
'   4. Write captured data into target:
'        • ALL formulas are kept as-is (internal AND external)
'          – external formulas on the new target are correct: they point to
'            the same external sheet and the same column-absolute refs
'          – internal formulas use R1C1 so they self-adjust to the new column
'   5. Freeze the source (OB current) → only direct-external formulas → value
'      (pure-internal formulas like SUM(N10,N11…) remain as formulas on source)
'================================================================================
Private Sub ProcessCase2_BlankTarget( _
        ByVal ws As Worksheet, _
        ByVal srcSpec As String)

    Dim srcRng As Range
    Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    If srcRng Is Nothing Then Exit Sub

    ws.DisplayPageBreaks = False

    Dim colCount    As Long : colCount    = srcRng.Columns.Count
    Dim srcFirstCol As Long : srcFirstCol = srcRng.Column
    Dim srcLastCol  As Long : srcLastCol  = srcRng.Column + colCount - 1
    Dim srcFirstRow As Long : srcFirstRow = srcRng.Row
    Dim srcRowCount As Long : srcRowCount = srcRng.Rows.Count

    ' ── Step 1: Capture source content BEFORE insert ────────────────────────
    ' We need Value2 for non-formula cells, and FormulaR1C1 for formula cells.
    ' R1C1 is used so that relative refs auto-adjust when written to the new col.
    Dim arrVal   As Variant : arrVal   = srcRng.Value2
    Dim arrFmlR  As Variant : arrFmlR  = srcRng.FormulaR1C1   ' R1C1 text
    Dim arrFmlA1 As Variant : arrFmlA1 = srcRng.Formula        ' A1 text (for classification)

    ' ── Step 2: Insert columns immediately to the right of source ────────────
    Dim insertAt As Long : insertAt = srcLastCol + 1
    Dim ci As Long
    For ci = 1 To colCount
        ws.Columns(insertAt).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next ci
    ' After insert, source has shifted by colCount if insertAt <= srcFirstCol,
    ' but since we insert to the RIGHT of source, source indices are unchanged.

    ' ── Step 3: Define target range and copy formats ─────────────────────────
    Dim tgtRng As Range
    Set tgtRng = ws.Range( _
        ws.Cells(srcFirstRow, insertAt), _
        ws.Cells(srcFirstRow + srcRowCount - 1, insertAt + colCount - 1))

    CopyRangeFormats srcRng, tgtRng   ' srcRng still valid (source did not shift)

    ' ── Step 4: Write captured data to target ────────────────────────────────
    ' Rule: keep ALL formulas (internal + external) on the new target.
    ' The target is the "live current period" column – it must have full formulas.
    ' R1C1 transplant handles relative-ref adjustment automatically.
    WriteArrayToRange_AllFormulas arrVal, arrFmlR, arrFmlA1, tgtRng

    ' ── Step 5: Freeze source (it is now the OB-current frozen column) ───────
    ' Only direct-external formulas are frozen to values.
    ' Pure-internal formulas (e.g. SUM of internal cells) stay as formulas.
    FreezeRange_DirectExtOnly ws, srcRng

    ' ── Step 6: Clear comments/notes from new target ─────────────────────────
    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' CASE 3 – REST  (use existing columns to the RIGHT as target)
' Source stays as opening-balance current  →  freeze direct-external on source
'================================================================================
Private Sub ProcessCase3_REST( _
        ByVal ws As Worksheet, _
        ByVal srcSpec As String)

    Dim srcRng As Range
    Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    If srcRng Is Nothing Then Exit Sub

    ws.DisplayPageBreaks = False

    Dim colCount   As Long : colCount   = srcRng.Columns.Count
    Dim srcLastCol As Long : srcLastCol = srcRng.Column + colCount - 1

    ' Target = same row extent, same width, immediately to the right
    Dim tgtFirstCol As Long : tgtFirstCol = srcLastCol + 1
    Dim tgtLastCol  As Long : tgtLastCol  = tgtFirstCol + colCount - 1

    If tgtLastCol > ws.Columns.Count Then Exit Sub

    Dim tgtRng As Range
    Set tgtRng = ws.Range( _
        ws.Cells(srcRng.Row, tgtFirstCol), _
        ws.Cells(srcRng.Row + srcRng.Rows.Count - 1, tgtLastCol))

    ' Capture source data before any changes
    Dim arrVal   As Variant : arrVal   = srcRng.Value2
    Dim arrFmlR  As Variant : arrFmlR  = srcRng.FormulaR1C1
    Dim arrFmlA1 As Variant : arrFmlA1 = srcRng.Formula

    ' Unhide / ungroup existing target columns
    UnhideUngroup ws, tgtFirstCol, tgtLastCol

    ' Write all formulas (internal + external) to target
    WriteArrayToRange_AllFormulas arrVal, arrFmlR, arrFmlA1, tgtRng

    ' Freeze source: direct-external only → value
    FreezeRange_DirectExtOnly ws, srcRng

    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' CASE 4 – WEST  (insert new columns to the LEFT of source)
' After insert, source shifts right; inserted block becomes target.
' Source stays as OB current  →  freeze direct-external on (shifted) source
'================================================================================
Private Sub ProcessCase4_WEST( _
        ByVal ws As Worksheet, _
        ByVal srcSpec As String)

    Dim srcRng As Range
    Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    If srcRng Is Nothing Then Exit Sub

    ws.DisplayPageBreaks = False

    Dim colCount    As Long : colCount    = srcRng.Columns.Count
    Dim srcFirstCol As Long : srcFirstCol = srcRng.Column
    Dim srcLastRow  As Long : srcLastRow  = srcRng.Row + srcRng.Rows.Count - 1

    ' Capture source data BEFORE insert (insert will shift srcRng reference)
    Dim arrVal   As Variant : arrVal   = srcRng.Value2
    Dim arrFmlR  As Variant : arrFmlR  = srcRng.FormulaR1C1
    Dim arrFmlA1 As Variant : arrFmlA1 = srcRng.Formula

    ' Insert columns to the left of source
    Dim ci As Long
    For ci = 1 To colCount
        ws.Columns(srcFirstCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next ci

    ' After insert: target = srcFirstCol..(srcFirstCol+colCount-1)
    '               source  = (srcFirstCol+colCount)..(srcFirstCol+2*colCount-1)
    Dim tgtRng As Range
    Set tgtRng = ws.Range( _
        ws.Cells(srcRng.Row, srcFirstCol), _
        ws.Cells(srcLastRow, srcFirstCol + colCount - 1))

    ' Re-resolve shifted source range
    Dim newSrcRng As Range
    Set newSrcRng = ws.Range( _
        ws.Cells(srcRng.Row, srcFirstCol + colCount), _
        ws.Cells(srcLastRow, srcFirstCol + 2 * colCount - 1))

    ' Copy formats from (now-shifted) source → target
    CopyRangeFormats newSrcRng, tgtRng

    ' Write all formulas (internal + external) to target
    WriteArrayToRange_AllFormulas arrVal, arrFmlR, arrFmlA1, tgtRng

    ' Freeze direct-external on (shifted) source
    FreezeRange_DirectExtOnly ws, newSrcRng

    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' CASE 5 – NEST  (use existing columns to the LEFT as target)
' Source stays as OB current  →  freeze direct-external on source
'================================================================================
Private Sub ProcessCase5_NEST( _
        ByVal ws As Worksheet, _
        ByVal srcSpec As String)

    Dim srcRng As Range
    Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    If srcRng Is Nothing Then Exit Sub

    ws.DisplayPageBreaks = False

    Dim colCount    As Long : colCount    = srcRng.Columns.Count
    Dim srcFirstCol As Long : srcFirstCol = srcRng.Column

    Dim tgtFirstCol As Long : tgtFirstCol = srcFirstCol - colCount
    Dim tgtLastCol  As Long : tgtLastCol  = srcFirstCol - 1

    If tgtFirstCol < 1 Then Exit Sub

    Dim tgtRng As Range
    Set tgtRng = ws.Range( _
        ws.Cells(srcRng.Row, tgtFirstCol), _
        ws.Cells(srcRng.Row + srcRng.Rows.Count - 1, tgtLastCol))

    ' Capture source data
    Dim arrVal   As Variant : arrVal   = srcRng.Value2
    Dim arrFmlR  As Variant : arrFmlR  = srcRng.FormulaR1C1
    Dim arrFmlA1 As Variant : arrFmlA1 = srcRng.Formula

    ' Unhide / ungroup existing left columns
    UnhideUngroup ws, tgtFirstCol, tgtLastCol

    ' Write all formulas (internal + external) to target
    WriteArrayToRange_AllFormulas arrVal, arrFmlR, arrFmlA1, tgtRng

    ' Freeze direct-external on source
    FreezeRange_DirectExtOnly ws, srcRng

    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' DELETE – Col D column/range deletion
'================================================================================
Private Sub DeleteColumnSpec(ByVal ws As Worksheet, ByVal delSpec As String)
    ' delSpec can be a single column letter ("C"), column range ("C:E"),
    ' or a comma-separated list ("C,F,H:J")

    Dim parts()  As String
    Dim p        As Long
    Dim spec     As String
    Dim rng      As Range

    parts = Split(delSpec, ",")

    ' Delete in reverse order to avoid index shift issues
    Dim colNums() As Long
    Dim cnt       As Long
    cnt = 0

    For p = 0 To UBound(parts)
        spec = Trim$(parts(p))
        If spec = "" Then GoTo NextPart

        On Error Resume Next
        Set rng = ResolveColSpec(ws, spec)
        On Error GoTo 0

        If Not rng Is Nothing Then
            ' Collect all column indices
            Dim col As Long
            For col = rng.Column To rng.Column + rng.Columns.Count - 1
                cnt = cnt + 1
                ReDim Preserve colNums(1 To cnt)
                colNums(cnt) = col
            Next col
        End If

NextPart:
    Next p

    If cnt = 0 Then Exit Sub

    ' Sort descending (bubble) so we delete from right to left
    Dim a As Long, b As Long, tmp As Long
    For a = 1 To cnt - 1
        For b = a + 1 To cnt
            If colNums(b) > colNums(a) Then
                tmp = colNums(a) : colNums(a) = colNums(b) : colNums(b) = tmp
            End If
        Next b
    Next a

    ' Delete each column, deduplicated
    Dim lastDel As Long : lastDel = -1
    For a = 1 To cnt
        If colNums(a) <> lastDel Then
            On Error Resume Next
            ws.Columns(colNums(a)).Delete
            On Error GoTo 0
            lastDel = colNums(a)
        End If
    Next a

End Sub


'================================================================================
' F1 FREEZE HELPER – builds a set of cell addresses that are direct-external
'   level 1 = formula references another sheet/workbook directly
'   Returns Scripting.Dictionary: ADDR (no $ , uppercase) -> True
'================================================================================
Private Function BuildFormulaLevels_F1(ByVal ws As Worksheet) As Object
    ' We only need level-1 for F1: cells whose formula directly references outside.
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim fc As Range
    On Error Resume Next
    Set fc = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If fc Is Nothing Then
        Set BuildFormulaLevels_F1 = result
        Exit Function
    End If

    Dim c As Range
    For Each c In fc.Cells
        If IsDirectExternalFormula(c.Formula, ws.Name) Then
            result(UCase$(c.Address(False, False))) = 1
        End If
    Next c

    Set BuildFormulaLevels_F1 = result
End Function


'================================================================================
' FreezeRange_F2  (kept as alias for backward compat – same as DirectExtOnly now)
'================================================================================
Private Sub FreezeRange_F2(ByVal ws As Worksheet, ByVal rng As Range)
    FreezeRange_DirectExtOnly ws, rng
End Sub


'================================================================================
' CopyRangeDataF1 – copy source → target using F1 rules
'   • values always copied
'   • internal formulas transplanted (FormulaR1C1)
'   • direct-external formulas → frozen (value only)
'   • target-internal formulas preserved
'================================================================================
Private Sub CopyRangeDataF1( _
        ByVal ws As Worksheet, _
        ByVal srcRng As Range, _
        ByVal tgtRng As Range)

    ' Build F1 level map for source cells
    Dim levelMap As Object
    Set levelMap = BuildFormulaLevels_F1(ws)

    ' Preserve existing internal formulas in target
    Dim keepFormula As Object : Set keepFormula = CreateObject("Scripting.Dictionary")
    Dim keepTarget  As Object : Set keepTarget  = CreateObject("Scripting.Dictionary")

    Dim fc As Range, c As Range, tgtCell As Range
    On Error Resume Next
    Set fc = tgtRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not fc Is Nothing Then
        For Each c In fc.Cells
            If Not IsDirectExternalFormula(c.Formula, ws.Name) Then
                Dim rk As String : rk = RelKey(c, tgtRng)
                keepTarget(rk)  = True
                keepFormula(rk) = c.FormulaR1C1
            End If
        Next c
    End If

    ' Mass-copy values
    tgtRng.Value2 = srcRng.Value2

    ' Process source formulas
    On Error Resume Next
    Set fc = srcRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not fc Is Nothing Then
        For Each c In fc.Cells
            Dim rk2 As String : rk2 = RelKey(c, srcRng)
            If keepTarget.Exists(rk2) Then GoTo SkipSrc
            Set tgtCell = tgtRng.Cells( _
                c.Row - srcRng.Row + 1, _
                c.Column - srcRng.Column + 1)

            If levelMap.Exists(UCase$(c.Address(False, False))) Then
                tgtCell.Value2 = c.Value2          ' F1: freeze external
            Else
                tgtCell.FormulaR1C1 = c.FormulaR1C1 ' internal: keep formula
            End If
SkipSrc:
        Next c
    End If

    ' Restore preserved target-internal formulas
    Dim k As Variant
    Dim parts() As String
    For Each k In keepFormula.Keys
        parts = Split(CStr(k), "|")
        tgtRng.Cells(CLng(parts(0)), CLng(parts(1))).FormulaR1C1 = keepFormula(k)
    Next k

End Sub


'================================================================================
' WriteArrayToRange_AllFormulas
'   Writes pre-captured arrays into a target range.
'   ALL formulas are kept (both internal AND external).
'   R1C1 is used for relative-ref transplant so refs auto-adjust to new column.
'   Values are written for non-formula cells.
'================================================================================
Private Sub WriteArrayToRange_AllFormulas( _
        ByVal arrVal   As Variant, _
        ByVal arrFmlR  As Variant, _
        ByVal arrFmlA1 As Variant, _
        ByVal tgtRng   As Range)

    Dim r   As Long, co  As Long
    Dim fR  As String    ' R1C1 formula text
    Dim fA1 As String    ' A1 formula text (for is-formula test)

    For r = 1 To tgtRng.Rows.Count
        For co = 1 To tgtRng.Columns.Count

            fR  = ""
            fA1 = ""
            If IsArray(arrFmlR) Then
                If VarType(arrFmlR(r, co)) = vbString Then fR = CStr(arrFmlR(r, co))
            End If
            If IsArray(arrFmlA1) Then
                If VarType(arrFmlA1(r, co)) = vbString Then fA1 = CStr(arrFmlA1(r, co))
            End If

            If Len(fA1) > 1 And Left$(fA1, 1) = "=" Then
                ' Keep formula as-is using R1C1 (handles relative-col adjustment)
                On Error Resume Next
                tgtRng.Cells(r, co).FormulaR1C1 = fR
                On Error GoTo 0
            Else
                ' Plain value
                If IsArray(arrVal) Then
                    On Error Resume Next
                    tgtRng.Cells(r, co).Value2 = arrVal(r, co)
                    On Error GoTo 0
                End If
            End If

        Next co
    Next r

End Sub


'================================================================================
' FreezeRange_DirectExtOnly
'   Replaces FreezeRange_F2.
'   Freezes ONLY cells whose formula directly references another sheet/workbook.
'   Pure-internal formulas (SUM of same-sheet cells, ROUND of same-sheet cells,
'   OFFSET-based totals, etc.) are left untouched even if those cells happen
'   to depend on external data transitively.
'
'   Rationale (from user spec):
'     N21 = ROUND(SUM(N10,N11,...),2)  →  all refs are same-sheet → F0 → no freeze
'     Q12 = ROUND(SUMIFS('Aggregate TB'!...),2)  →  direct external → freeze
'================================================================================
Private Sub FreezeRange_DirectExtOnly( _
        ByVal ws As Worksheet, _
        ByVal rng As Range)

    Dim fc  As Range
    Dim c   As Range

    On Error Resume Next
    Set fc = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If fc Is Nothing Then Exit Sub

    For Each c In fc.Cells
        If IsDirectExternalFormula(c.Formula, ws.Name) Then
            ' Freeze: replace formula with its current computed value
            Dim v As Variant
            v = c.Value2
            c.Value2 = v
        End If
        ' internal formula (regardless of transitive dependencies) → leave as formula
    Next c

End Sub


'================================================================================
' SUBTOTAL / FORMAT UTILITIES
'================================================================================

' Copy column widths and cell formats from source → target range
Private Sub CopyRangeFormats( _
        ByVal srcRng As Range, _
        ByVal tgtRng As Range)

    On Error Resume Next
    srcRng.Copy
    tgtRng.PasteSpecial xlPasteFormats
    tgtRng.PasteSpecial xlPasteValidation
    Application.CutCopyMode = False

    ' Match column widths
    Dim c As Long
    For c = 1 To srcRng.Columns.Count
        tgtRng.Columns(c).ColumnWidth = srcRng.Columns(c).ColumnWidth
    Next c
    On Error GoTo 0

End Sub

' Clear comments and notes from a range
Private Sub ClearCommentsNotes(ByVal rng As Range)
    On Error Resume Next
    rng.ClearComments
    rng.ClearNotes
    On Error GoTo 0
End Sub

' Unhide and ungroup a column range
Private Sub UnhideUngroup(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long)
    On Error Resume Next
    ws.Range(ws.Columns(firstCol), ws.Columns(lastCol)).Hidden = False
    ws.Range(ws.Columns(firstCol), ws.Columns(lastCol)).Ungroup
    On Error GoTo 0
End Sub


'================================================================================
' RANGE RESOLUTION UTILITIES
'================================================================================

' Resolve a range spec and trim whole-column refs to used rows
Private Function ResolveAndTrimRange( _
        ByVal ws As Worksheet, _
        ByVal spec As String) As Range

    Dim rng As Range
    On Error Resume Next
    Set rng = ResolveColOrRange(ws, spec)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function

    ' Trim whole-column references to actual used rows
    If rng.Rows.Count = ws.Rows.Count Then
        Dim ur As Range
        Set ur = ws.UsedRange
        If ur Is Nothing Then
            Set ResolveAndTrimRange = Nothing
            Exit Function
        End If
        Dim fr As Long : fr = ur.Row
        Dim lr As Long : lr = ur.Row + ur.Rows.Count - 1
        Dim fc As Long : fc = rng.Column
        Dim lc As Long : lc = rng.Column + rng.Columns.Count - 1
        Set ResolveAndTrimRange = ws.Range(ws.Cells(fr, fc), ws.Cells(lr, lc))
    Else
        Set ResolveAndTrimRange = rng
    End If

End Function

' Resolve "C", "C:E", "C2:D50", "$C$2:$D$50" → Range
Private Function ResolveColOrRange(ByVal ws As Worksheet, ByVal spec As String) As Range
    Dim s As String
    s = Trim$(Replace(spec, "$", ""))
    If s = "" Then Exit Function
    On Error Resume Next
    If IsColumnLettersOnly(s) Or IsColumnRangeLettersOnly(s) Then
        Set ResolveColOrRange = ws.Columns(s)
    Else
        Set ResolveColOrRange = ws.Range(s)
    End If
    On Error GoTo 0
End Function

' Resolve a delete-spec that may be "C", "C:E" (columns only)
Private Function ResolveColSpec(ByVal ws As Worksheet, ByVal spec As String) As Range
    Dim s As String
    s = Trim$(Replace(spec, "$", ""))
    If s = "" Then Exit Function
    On Error Resume Next
    Set ResolveColSpec = ws.Columns(s)
    On Error GoTo 0
End Function

' True if spec is purely column letters (A, B, AC …) with optional colon for range
Private Function IsExplicitRangeOrColumn(ByVal spec As Variant) As Boolean
    Dim s As String
    s = Trim$(CStr(spec))
    If s = "" Then Exit Function

    Dim u As String : u = UCase$(s)
    If u = "REST" Or u = "WEST" Or u = "NEST" Then Exit Function

    ' If it contains a digit or colon or $ it's a range/column spec
    ' Accept: "C", "AB", "C:F", "C2:D50", "$C$2:$D$50"
    On Error Resume Next
    Dim rng As Range
    Set rng = ThisWorkbook.Worksheets(1).Range(Replace(s, "$", ""))
    On Error GoTo 0
    If Not rng Is Nothing Then
        IsExplicitRangeOrColumn = True
        Exit Function
    End If

    ' Fallback: treat as column if only letters (possibly with colon)
    IsExplicitRangeOrColumn = IsColumnLettersOnly(s) Or IsColumnRangeLettersOnly(s)
End Function

Private Function IsColumnLettersOnly(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    s = Trim$(s)
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then
            Exit Function
        End If
    Next i
    IsColumnLettersOnly = True
End Function

Private Function IsColumnRangeLettersOnly(ByVal s As String) As Boolean
    ' Matches patterns like "C:F" or "AB:CD"
    Dim colonPos As Long
    colonPos = InStr(s, ":")
    If colonPos < 2 Then Exit Function
    Dim left1  As String : left1  = Left$(s, colonPos - 1)
    Dim right1 As String : right1 = Mid$(s, colonPos + 1)
    IsColumnRangeLettersOnly = IsColumnLettersOnly(left1) And IsColumnLettersOnly(right1)
End Function

' Relative position key for a cell within a base range: "row|col"
Private Function RelKey(ByVal c As Range, ByVal baseRng As Range) As String
    RelKey = CStr(c.Row - baseRng.Row + 1) & "|" & CStr(c.Column - baseRng.Column + 1)
End Function


'================================================================================
' FORMULA ANALYSIS UTILITIES
'================================================================================

' Returns True if the formula directly references a different sheet or workbook.
' Uses a fast Regex approach (cached Static).
Private Function IsDirectExternalFormula( _
        ByVal formulaText As String, _
        ByVal currentSheetName As String) As Boolean

    Static re As Object

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        With re
            .Global     = True
            .IgnoreCase = False
            .Pattern    = "((?:'[^']*(?:''[^']*)*'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+))!"
        End With
    End If

    If Not re.Test(formulaText) Then Exit Function   ' no sheet qualifier → internal

    Dim matches As Object
    Set matches = re.Execute(formulaText)

    Dim m         As Object
    Dim qualifier As String
    Dim cleaned   As String

    For Each m In matches
        qualifier = m.SubMatches(0)

        ' External workbook reference
        If InStr(qualifier, "[") > 0 Then
            IsDirectExternalFormula = True
            Exit Function
        End If

        ' Strip quotes
        cleaned = qualifier
        If Left$(cleaned, 1) = "'" And Right$(cleaned, 1) = "'" Then
            cleaned = Mid$(cleaned, 2, Len(cleaned) - 2)
            cleaned = Replace(cleaned, "''", "'")
        End If

        If StrComp(cleaned, currentSheetName, vbTextCompare) <> 0 Then
            IsDirectExternalFormula = True
            Exit Function
        End If
    Next m

End Function



'================================================================================
' GENERAL HELPERS
'================================================================================

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                          MatchCase:=False)
    If f Is Nothing Then LastUsedRow = 0 Else LastUsedRow = f.Row
End Function

Private Function ElapsedSec(ByVal t0 As Double) As Double
    Dim t1 As Double : t1 = Timer
    If t1 < t0 Then t1 = t1 + 86400#
    ElapsedSec = t1 - t0
End Function
