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
    WriteArrayToRange_AllFormulas arrVal, arrFmlR, arrFmlA1, tgtRng, srcFirstCol, ws.Name

    ' ── Step 5: Freeze source (F1 direct-external + F2 indirect-external) ────
    ' O12=ROUND(Q12-...) where Q12 is SUMIFS→external → O12 is F2 → freeze ✓
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

    ' Write all formulas to target: internal refs shifted, external refs verbatim
    WriteArrayToRange_AllFormulas arrVal, arrFmlR, arrFmlA1, tgtRng, srcRng.Column, ws.Name

    ' Freeze source: F1 + F2 BFS
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

    ' Write all formulas to target: internal refs shifted, external refs verbatim
    ' srcFirstCol is the ORIGINAL source position (before insert shifted it)
    WriteArrayToRange_AllFormulas arrVal, arrFmlR, arrFmlA1, tgtRng, srcFirstCol, ws.Name

    ' Freeze F1+F2 on (shifted) source
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

    ' Write all formulas to target: internal refs shifted, external refs verbatim
    WriteArrayToRange_AllFormulas arrVal, arrFmlR, arrFmlA1, tgtRng, srcFirstCol, ws.Name

    ' Freeze F1+F2 on source
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
' FreezeRange_F2  (alias – delegates to the full F1+F2 BFS freeze)
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
'
'   KEY RULE for column-shift:
'     • Internal (same-sheet) relative column refs  → shifted by colDelta
'     • External sheet refs (Sheet!CellAddr)        → kept EXACTLY as captured
'       e.g. ='Cover page'!E19  stays ='Cover page'!E19 in every new column
'       e.g. =G12+'Cover page'!E19  →  G shifts to H, Cover page E19 stays E19
'
'   We do NOT use FormulaR1C1 because that shifts ALL relative refs including
'   external ones.  Instead we use the A1 formula and surgically shift only
'   internal column references.
'
'   colDelta = (target first column) - (source first column)
'================================================================================
Private Sub WriteArrayToRange_AllFormulas( _
        ByVal arrVal   As Variant, _
        ByVal arrFmlR  As Variant, _
        ByVal arrFmlA1 As Variant, _
        ByVal tgtRng   As Range, _
        ByVal srcFirstCol As Long, _
        ByVal sheetName   As String)

    Dim r        As Long
    Dim co       As Long
    Dim fA1      As String
    Dim fShifted As String
    Dim colDelta As Long

    colDelta = tgtRng.Column - srcFirstCol   ' how many columns to the right target is

    For r = 1 To tgtRng.Rows.Count
        For co = 1 To tgtRng.Columns.Count

            fA1 = ""
            If IsArray(arrFmlA1) Then
                If VarType(arrFmlA1(r, co)) = vbString Then fA1 = CStr(arrFmlA1(r, co))
            End If

            If Len(fA1) > 1 And Left$(fA1, 1) = "=" Then
                ' Shift only internal relative column refs by colDelta
                ' External sheet refs are left verbatim
                If colDelta = 0 Then
                    fShifted = fA1
                Else
                    fShifted = ShiftInternalColRefs(fA1, sheetName, colDelta)
                End If
                On Error Resume Next
                tgtRng.Cells(r, co).Formula = fShifted
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
' ShiftInternalColRefs
'   Rewrites an A1-notation formula string, shifting ONLY internal (same-sheet)
'   relative column references by colDelta columns.
'   External sheet qualifiers (Sheet!Addr or 'Sheet name'!Addr) are left intact.
'
'   Algorithm:
'     Walk the formula character-by-character.
'     Track whether we are inside a string literal ("...").
'     When we find a sheet!  qualifier followed by a cell address, skip it
'       (leave the entire SheetName!CellAddr token untouched).
'     When we find a bare cell address (no sheet qualifier), check if the
'       column part is relative (no leading $).  If so, shift it by colDelta.
'     Row part is never touched.
'================================================================================
Private Function ShiftInternalColRefs( _
        ByVal formula    As String, _
        ByVal sheetName  As String, _
        ByVal colDelta   As Long) As String

    Dim n       As Long : n = Len(formula)
    Dim i       As Long : i = 1
    Dim out     As String : out = ""
    Dim ch      As String

    Do While i <= n
        ch = Mid$(formula, i, 1)

        ' ── String literal: copy verbatim until closing quote ────────────────
        If ch = """" Then
            out = out & ch
            i = i + 1
            Do While i <= n
                ch = Mid$(formula, i, 1)
                out = out & ch
                i = i + 1
                If ch = """" Then
                    ' doubled quote = escaped quote, not end of string
                    If i <= n And Mid$(formula, i, 1) = """" Then
                        out = out & Mid$(formula, i, 1)
                        i = i + 1
                    Else
                        Exit Do
                    End If
                End If
            Loop
            GoTo ContinueLoop
        End If

        ' ── Sheet-qualified reference: Sheet!Addr or 'Sheet name'!Addr ───────
        ' We detect a sheet qualifier by looking for the pattern:
        '   identifier! or 'quoted name'!
        ' When found we copy the qualifier + address verbatim (no shift).
        Dim qualStart As Long
        Dim qualEnd   As Long
        Dim qualName  As String

        If ch = "'" Then
            ' Try to read a quoted sheet name
            qualStart = i
            qualEnd = i + 1
            Do While qualEnd <= n
                If Mid$(formula, qualEnd, 1) = "'" Then
                    ' Check for escaped ''
                    If qualEnd + 1 <= n And Mid$(formula, qualEnd + 1, 1) = "'" Then
                        qualEnd = qualEnd + 2
                    Else
                        Exit Do
                    End If
                Else
                    qualEnd = qualEnd + 1
                End If
            Loop
            ' qualEnd now points to the closing '
            ' Check if followed by !
            If qualEnd <= n And Mid$(formula, qualEnd, 1) = "'" Then
                If qualEnd + 1 <= n And Mid$(formula, qualEnd + 1, 1) = "!" Then
                    ' This is a sheet qualifier – copy qualifier + ! + cell addr verbatim
                    Dim qualToken As String
                    qualToken = Mid$(formula, qualStart, qualEnd - qualStart + 2) ' includes '...'!
                    out = out & qualToken
                    i = qualEnd + 2
                    ' Now copy the cell address verbatim (it belongs to external sheet)
                    i = CopyAddressVerbatim(formula, i, out)
                    GoTo ContinueLoop
                End If
            End If
            ' Not a sheet qualifier – just copy the ' and move on
            out = out & ch
            i = i + 1
            GoTo ContinueLoop
        End If

        ' Unquoted identifier followed by !  → external sheet qualifier
        If IsColStartChar(ch) Then
            Dim idStart As Long : idStart = i
            Dim j       As Long : j = i + 1
            Do While j <= n And IsSheetIdentChar(Mid$(formula, j, 1))
                j = j + 1
            Loop
            ' j now points to char after identifier
            If j <= n And Mid$(formula, j, 1) = "!" Then
                ' external sheet qualifier – copy verbatim: identifier + ! + cell addr
                Dim idToken As String : idToken = Mid$(formula, idStart, j - idStart + 1)
                out = out & idToken
                i = j + 1
                i = CopyAddressVerbatim(formula, i, out)
                GoTo ContinueLoop
            End If
            ' Not followed by ! – could be start of a bare cell address
            ' Fall through to cell-address detection below (reset i to idStart)
            i = idStart
        End If

        ' ── Bare cell address (no sheet qualifier): try to shift col ─────────
        If IsColStartChar(ch) Or ch = "$" Then
            Dim addrStart As Long : addrStart = i
            Dim colAbs    As Boolean : colAbs = False
            Dim colLetters As String : colLetters = ""
            Dim rowAbs    As Boolean : rowAbs = False
            Dim rowDigits As String  : rowDigits = ""
            Dim k         As Long    : k = i

            If Mid$(formula, k, 1) = "$" Then
                colAbs = True
                k = k + 1
            End If

            Do While k <= n And IsLetterAZ(Mid$(formula, k, 1))
                colLetters = colLetters & Mid$(formula, k, 1)
                k = k + 1
            Loop

            If Len(colLetters) > 0 And Len(colLetters) <= 3 Then
                If k <= n And Mid$(formula, k, 1) = "$" Then
                    rowAbs = True
                    k = k + 1
                End If
                Do While k <= n And IsDigit09(Mid$(formula, k, 1))
                    rowDigits = rowDigits & Mid$(formula, k, 1)
                    k = k + 1
                Loop

                If Len(rowDigits) > 0 Then
                    ' Valid cell address – check prev char is not ! (already handled)
                    Dim prevCh As String
                    prevCh = ""
                    If addrStart > 1 Then prevCh = Mid$(formula, addrStart - 1, 1)

                    If prevCh <> "!" Then
                        ' Internal cell ref – shift column if relative
                        If Not colAbs Then
                            Dim colNum As Long
                            colNum = ColLettersToNum(UCase$(colLetters)) + colDelta
                            If colNum >= 1 And colNum <= 16384 Then
                                colLetters = NumToColLetters(colNum)
                            End If
                        End If
                        ' Reconstruct token
                        Dim tok As String
                        tok = IIf(colAbs, "$", "") & colLetters
                        tok = tok & IIf(rowAbs, "$", "") & rowDigits
                        out = out & tok
                        i = k
                        GoTo ContinueLoop
                    End If
                End If
            End If
        End If

        ' Default: copy character as-is
        out = out & ch
        i = i + 1

ContinueLoop:
    Loop

    ShiftInternalColRefs = out
End Function


' Copy a cell address token verbatim from formula at position i.
' Returns new i after address.  Appends to out.
Private Function CopyAddressVerbatim( _
        ByVal formula As String, _
        ByVal i       As Long, _
        ByRef out     As String) As Long

    Dim n  As Long : n = Len(formula)
    Dim k  As Long : k = i

    ' Optional $ before column
    If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1

    ' Column letters
    Do While k <= n And IsLetterAZ(Mid$(formula, k, 1))
        k = k + 1
    Loop

    ' Optional $ before row
    If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1

    ' Row digits
    Do While k <= n And IsDigit09(Mid$(formula, k, 1))
        k = k + 1
    Loop

    ' Also handle column-only range like "A:B" after sheet qualifier
    If k <= n And Mid$(formula, k, 1) = ":" Then
        k = k + 1
        If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1
        Do While k <= n And IsLetterAZ(Mid$(formula, k, 1))
            k = k + 1
        Loop
        If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1
        Do While k <= n And IsDigit09(Mid$(formula, k, 1))
            k = k + 1
        Loop
    End If

    out = out & Mid$(formula, i, k - i)
    CopyAddressVerbatim = k
End Function


'================================================================================
' FreezeRange_DirectExtOnly  (source-column freeze, Cases 2-5)
'
' The rule, derived from user examples:
'
'   FREEZE a cell on source if:
'     (a) its formula directly references an external sheet (F1), OR
'     (b) its formula references a same-sheet cell that was ITSELF a live
'         formula at the time of processing AND that cell is F1 (F2 rule)
'
'   DO NOT FREEZE a cell if all same-sheet cells it references have
'   already been frozen to values by the time we reach it.
'
' This is implemented as a two-pass, top-to-bottom scan:
'
'   Pass 1 – collect all F1 cell addresses (direct-external formulas).
'             These are candidates for propagation, but NOT yet frozen.
'
'   Pass 2 – iterate the source range row by row (top to bottom).
'     For each formula cell C:
'       • If C is F1 → freeze it; remove it from the "live F1" set.
'       • Else if C references any address still in the "live F1" set
'         (i.e. a cell that is still a live external formula, not yet
'         frozen to a value) → freeze C too (it is F2).
'       • Else → leave C as a formula (pure internal or sum of already-
'         frozen values like the subtotal row).
'
' Why this works for the two problem cases:
'
'   Q12 = ROUND(SUMIFS('Aggregate TB'!...),2)
'     → F1. Pass 1 adds Q12 to liveF1. Pass 2 freezes Q12, removes from liveF1.
'
'   O12 = ROUND(Q12 - SUM(F12:OFFSET(N12,0,-1)),2)
'     → not F1. Pass 2: does O12 reference anything in liveF1? Q12 was in
'       liveF1 BEFORE it was processed. But O12 comes AFTER Q12 (lower row),
'       so by the time we reach O12, Q12 has already been removed from liveF1
'       (frozen to value). Therefore O12 sees no live F1 refs → stays formula.
'     WAIT – this is wrong direction. O12 IS in a lower row than Q12 so it
'     would be processed after Q12. After Q12 is frozen, liveF1 no longer
'     contains Q12. So O12 would stay as formula. But user says O12 must freeze!
'
' ── REVISED UNDERSTANDING ────────────────────────────────────────────────────
'
' Re-reading the user: source column is e.g. column O (or N).
' Q12 is NOT in the source column – it is in a different column (Q).
' The source column contains O12 = ROUND(Q12-SUM(F12:OFFSET(N12,0,-1)),2).
' Q12 is external to the source column and is NEVER frozen by this routine
' (only the source column cells are frozen).
' So the liveF1 check must look at ALL formula cells on the sheet, not just
' the source column.
'
' Correct algorithm:
'   Step 1: Build a set of ALL F1 cell addresses on the WHOLE SHEET (not just
'           source column). These cells directly reference external sheets.
'           They remain as formulas (we don't freeze them here – they are in
'           OTHER columns). They are the "live F1 anchors".
'
'   Step 2: Iterate source-column cells top-to-bottom.
'     For each formula cell C in source:
'       • If C is itself F1 → freeze it.
'       • Else if C's formula references any address in the whole-sheet F1 set
'         → freeze C (it is F2: depends on a live external-data cell).
'       • Else if C's formula references any source-column cell that was
'         already frozen in this pass → freeze C too (propagation within source).
'       • Else → leave as formula.
'
' This correctly handles:
'   O12 = ROUND(Q12-SUM(F12:OFFSET(N12,0,-1)),2)
'     Q12 is in whole-sheet F1 set → O12 is F2 → FREEZE ✓
'
'   SUM cell = ROUND(SUM(O10,O11,O12,...),2)  (below O12 in source column)
'     O10,O11 are also F2 in source and get frozen.
'     O12 gets frozen. But SUM cell references O10,O11,O12 which are same-sheet
'     and NOT in the whole-sheet F1 set (they are internal formulas, not direct
'     external). So SUM cell is NOT F2 by the whole-sheet rule.
'     And O10..O12 are in the SOURCE column and were frozen earlier in this pass,
'     but the propagation-within-source rule would also freeze SUM cell.
'     HOWEVER user says SUM cell must stay formula. So the third bullet
'     (propagation within source) must NOT apply.
'
' FINAL RULE (confirmed by all user examples):
'   Freeze source cell C if and only if:
'     (a) C is itself F1 (direct external), OR
'     (b) C's formula text contains a reference to a cell in the WHOLE-SHEET
'         F1 set (a cell that directly pulls from an external sheet).
'   No further propagation beyond one hop.
'================================================================================
Private Sub FreezeRange_DirectExtOnly( _
        ByVal ws As Worksheet, _
        ByVal rng As Range)

    ' ── Step 1: Build whole-sheet F1 set ────────────────────────────────────
    ' All formula cells on the sheet that directly reference an external sheet.
    Dim f1Set As Object
    Set f1Set = CreateObject("Scripting.Dictionary")
    f1Set.CompareMode = vbTextCompare

    Dim allFC As Range
    On Error Resume Next
    Set allFC = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not allFC Is Nothing Then
        Dim wc As Range
        For Each wc In allFC.Cells
            If IsDirectExternalFormula(wc.Formula, ws.Name) Then
                f1Set(UCase$(wc.Address(False, False))) = True
            End If
        Next wc
    End If

    ' ── Step 2: Iterate source range, freeze F1 and F2 cells ─────────────────
    ' F1: cell's own formula is direct-external.
    ' F2: cell's formula references at least one address in f1Set
    '     (one-hop only – no further propagation).
    ' SUM/aggregation cells that reference other SOURCE cells (not f1Set) → kept.
    Dim fc As Range
    On Error Resume Next
    Set fc = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If fc Is Nothing Then Exit Sub

    Dim c As Range
    For Each c In fc.Cells
        Dim fml As String : fml = c.Formula

        ' F1: direct external in this cell's own formula
        If IsDirectExternalFormula(fml, ws.Name) Then
            Dim v1 As Variant : v1 = c.Value2
            c.Value2 = v1
            GoTo NextCell
        End If

        ' F2: formula text contains a reference to a whole-sheet F1 cell
        If FormulaReferencesAnyInSet(fml, ws.Name, f1Set) Then
            Dim v2 As Variant : v2 = c.Value2
            c.Value2 = v2
        End If
        ' Else: pure internal (SUM of same-sheet cells not in f1Set) → leave as formula

NextCell:
    Next c

End Sub


'================================================================================
' FormulaReferencesAnyInSet
'   Returns True if the formula text contains a bare (same-sheet) cell reference
'   whose address exists in the given address set (dictionary).
'   External sheet refs (Sheet!Addr) are excluded from matching.
'   Uses the same character-walk as ShiftInternalColRefs.
'================================================================================
Private Function FormulaReferencesAnyInSet( _
        ByVal formula  As String, _
        ByVal sheetName As String, _
        ByVal addrSet  As Object) As Boolean

    Dim n  As Long : n = Len(formula)
    Dim i  As Long : i = 1
    Dim ch As String

    Do While i <= n
        ch = Mid$(formula, i, 1)

        ' Skip string literals
        If ch = """" Then
            i = i + 1
            Do While i <= n
                Dim qch As String : qch = Mid$(formula, i, 1)
                i = i + 1
                If qch = """" Then
                    If i <= n And Mid$(formula, i, 1) = """" Then
                        i = i + 1
                    Else
                        Exit Do
                    End If
                End If
            Loop
            GoTo NextCh
        End If

        ' Skip quoted sheet name → external ref → skip qualifier + address
        If ch = "'" Then
            i = i + 1
            Do While i <= n
                Dim qc2 As String : qc2 = Mid$(formula, i, 1)
                i = i + 1
                If qc2 = "'" Then
                    If i <= n And Mid$(formula, i, 1) = "'" Then
                        i = i + 1
                    Else
                        Exit Do
                    End If
                End If
            Loop
            If i <= n And Mid$(formula, i, 1) = "!" Then i = i + 1
            i = SkipCellAddr(formula, i)   ' skip the cell address after Sheet!
            GoTo NextCh
        End If

        ' Unquoted identifier possibly followed by ! → external qualifier
        If IsLetterAZ(ch) Then
            Dim idS2 As Long : idS2 = i
            Dim j2   As Long : j2 = i + 1
            Do While j2 <= n And IsSheetIdentChar(Mid$(formula, j2, 1))
                j2 = j2 + 1
            Loop
            If j2 <= n And Mid$(formula, j2, 1) = "!" Then
                ' External: skip identifier + ! + address
                i = j2 + 1
                i = SkipCellAddr(formula, i)
                GoTo NextCh
            End If
            ' Not external – try to parse as cell address starting at idS2
            i = idS2
        End If

        ' Try to parse bare cell address
        If IsLetterAZ(ch) Or ch = "$" Then
            Dim colAbs3  As Boolean : colAbs3 = False
            Dim colL3    As String  : colL3 = ""
            Dim rowAbs3  As Boolean : rowAbs3 = False
            Dim rowD3    As String  : rowD3 = ""
            Dim k3       As Long    : k3 = i

            If Mid$(formula, k3, 1) = "$" Then colAbs3 = True : k3 = k3 + 1
            Do While k3 <= n And IsLetterAZ(Mid$(formula, k3, 1))
                colL3 = colL3 & Mid$(formula, k3, 1) : k3 = k3 + 1
            Loop
            If Len(colL3) >= 1 And Len(colL3) <= 3 Then
                If k3 <= n And Mid$(formula, k3, 1) = "$" Then
                    rowAbs3 = True : k3 = k3 + 1
                End If
                Do While k3 <= n And IsDigit09(Mid$(formula, k3, 1))
                    rowD3 = rowD3 & Mid$(formula, k3, 1) : k3 = k3 + 1
                Loop
                If Len(rowD3) > 0 Then
                    Dim bareA As String
                    bareA = UCase$(colL3) & rowD3
                    If addrSet.Exists(bareA) Then
                        FormulaReferencesAnyInSet = True
                        Exit Function
                    End If
                    i = k3
                    GoTo NextCh
                End If
            End If
        End If

        i = i + 1
NextCh:
    Loop

End Function


' Skip over a cell address (or range like A1:B2 or A:B) at position i.
Private Function SkipCellAddr(ByVal formula As String, ByVal i As Long) As Long
    Dim n As Long : n = Len(formula)
    If i <= n And Mid$(formula, i, 1) = "$" Then i = i + 1
    Do While i <= n And IsLetterAZ(Mid$(formula, i, 1)) : i = i + 1 : Loop
    If i <= n And Mid$(formula, i, 1) = "$" Then i = i + 1
    Do While i <= n And IsDigit09(Mid$(formula, i, 1)) : i = i + 1 : Loop
    If i <= n And Mid$(formula, i, 1) = ":" Then
        i = i + 1
        If i <= n And Mid$(formula, i, 1) = "$" Then i = i + 1
        Do While i <= n And IsLetterAZ(Mid$(formula, i, 1)) : i = i + 1 : Loop
        If i <= n And Mid$(formula, i, 1) = "$" Then i = i + 1
        Do While i <= n And IsDigit09(Mid$(formula, i, 1)) : i = i + 1 : Loop
    End If
    SkipCellAddr = i
End Function




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
' CHARACTER / COLUMN HELPERS
'================================================================================

Private Function IsColStartChar(ByVal ch As String) As Boolean
    ' Valid first char of a column identifier or cell address
    IsColStartChar = IsLetterAZ(ch)
End Function

Private Function IsSheetIdentChar(ByVal ch As String) As Boolean
    ' Valid char inside an unquoted sheet name identifier (A-Z, 0-9, _, .)
    If Len(ch) = 0 Then Exit Function
    IsSheetIdentChar = IsLetterAZ(ch) Or IsDigit09(ch) Or ch = "_" Or ch = "."
End Function

Private Function IsLetterAZ(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim a As Integer : a = Asc(UCase$(ch))
    IsLetterAZ = (a >= 65 And a <= 90)
End Function

Private Function IsDigit09(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim a As Integer : a = Asc(ch)
    IsDigit09 = (a >= 48 And a <= 57)
End Function

Private Function ColLettersToNum(ByVal letters As String) As Long
    Dim i As Long, v As Long
    For i = 1 To Len(letters)
        v = v * 26 + (Asc(UCase$(Mid$(letters, i, 1))) - 64)
    Next i
    ColLettersToNum = v
End Function

Private Function NumToColLetters(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop
    NumToColLetters = s
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
