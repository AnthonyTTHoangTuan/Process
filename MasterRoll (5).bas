Option Explicit

'================================================================================
' MasterRoll – Period-Roll Macro for Multi-Sheet Financial Models
'
' PURPOSE
'   Advances one reporting period across all sheets configured in SheetList.
'   Depending on the instruction in Col C, the macro inserts or reuses a target
'   column block, transplants formulas and values intelligently, and optionally
'   deletes a stale column (Col D).
'
' SHEETLIST LAYOUT  (header row 1, data from row 2 onward)
'   Col A  Sheet name.
'   Col B  Source column letter(s) or range  (e.g. "O", "O:P", "O2:P50").
'   Col C  Target specification – one of:
'            <column/range>   Explicit target (Case 1).  F1 freeze on target.
'            REST             Reuse the columns immediately to the right (Case 3).
'            WEST             Insert new columns immediately to the left  (Case 4).
'            NEST             Reuse the columns immediately to the left   (Case 5).
'            <blank>          Insert new columns immediately to the right (Case 2).
'   Col D  Column(s) to delete after rolling.  Leave blank to skip.
'
' FREEZE CLASSIFICATION
'   F1  A formula cell whose own text directly references another sheet or an
'       external workbook (e.g. =SUMIFS('Aggregate TB'!$G:$G,...)).
'   F2  A formula cell that contains a bare same-sheet reference to an F1 cell.
'       Detected by a one-hop lookup against the whole-sheet F1 address set.
'       Example: =ROUND(Q12-SUM(F12:OFFSET(N12,0,-1)),2) where Q12 is F1.
'   F0  A formula whose own text and whose direct same-sheet references are all
'       internal (e.g. =ROUND(SUM(O10,O11,O12),2)).  Stays as a live formula.
'
' FREEZE RULES BY CASE
'   Case 1 (explicit target):
'     F1 source cells → written as values into target (external data locked in).
'     F0 source cells → formula transplanted to target via FormulaR1C1.
'     Source column is not modified.
'   Cases 2-5 (structural rolls):
'     Target receives ALL formulas (F0, F1, F2) – it is the new live column.
'     Internal column refs are shifted to the target position; external refs
'     are left verbatim (see WriteAllFormulasToRange).
'     Source is frozen: F1 and F2 cells → values; F0 cells stay as formulas.
'
' EXECUTION ORDER  (three passes to honour stated priority)
'   Pass 1  All Case-1 rows (explicit target) – processed first so target
'           columns are populated before any structural insert or delete.
'   Pass 2  All Cases 2-5 rows (structural rolls).
'   Pass 3  All Col D deletions – run last to avoid invalidating column indices
'           referenced in Passes 1 and 2.
'
' PERFORMANCE DESIGN
'   All SheetList cells are read in a single Value2 array snapshot before the
'   loops, eliminating repeated COM round-trips to the worksheet.
'   Source range content (Value2, Formula) is captured once per roll call.
'   The whole-sheet F1 address set is a Scripting.Dictionary for O(1) lookup.
'   The RegExp object in IsDirectExternalFormula is Static (compiled once).
'   The formula rewriter (ShiftInternalColRefs) is a single-pass character scan.
'   Application.Calculation, ScreenUpdating, and EnableEvents are suspended for
'   the full macro run and restored unconditionally in both normal and error paths.
'================================================================================

Private Const SHOW_ELAPSED As Boolean = True   ' set False to suppress timing MsgBox


'================================================================================
' PUBLIC ENTRY POINT
'================================================================================

Public Sub MasterRoll_Run()

    Const LIST_SHEET As String = "SheetList"

    Dim wsList    As Worksheet
    Dim lastRow   As Long
    Dim i         As Long
    Dim startTime As Double
    Dim elapsed   As Double
    Dim oldCalc   As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldStatus As Variant

    On Error GoTo FailSafe

    startTime = Timer

    On Error Resume Next
    Set wsList = ThisWorkbook.Worksheets(LIST_SHEET)
    On Error GoTo FailSafe
    If wsList Is Nothing Then
        MsgBox "Sheet '" & LIST_SHEET & "' was not found in this workbook.", vbExclamation
        Exit Sub
    End If

    lastRow = LastUsedRow(wsList)
    If lastRow < 2 Then
        MsgBox "SheetList contains no data rows (header only or empty).", vbInformation
        Exit Sub
    End If

    ' Suspend recalculation and screen refresh for the duration of the macro.
    oldCalc   = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldStatus = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual
    Application.CutCopyMode    = False

    ' Read the entire SheetList instruction block in one array assignment.
    ' Accessing a 2-D Variant array in VBA is orders of magnitude faster than
    ' calling .Cells(r, c).Value2 inside a loop.
    Dim listData  As Variant
    listData = wsList.Range(wsList.Cells(2, 1), wsList.Cells(lastRow, 4)).Value2
    Dim rowCount  As Long : rowCount = lastRow - 1   ' number of data rows

    ' ══════════════════════════════════════════════════════════════════════════
    ' PASS 1 – Explicit-target rows  (Col C contains a real column/range spec)
    '
    ' Processed first so that explicit prior-period columns are populated before
    ' any structural roll could insert columns nearby and shift their indices.
    ' ══════════════════════════════════════════════════════════════════════════
    Application.StatusBar = "MasterRoll Pass 1 of 3: explicit-target rows..."

    For i = 1 To rowCount
        Dim sName1   As String : sName1   = Trim$(CStr(listData(i, 1)))
        Dim srcSpec1 As String : srcSpec1 = Trim$(CStr(listData(i, 2)))
        Dim tgtSpec1 As String : tgtSpec1 = Trim$(CStr(listData(i, 3)))

        If sName1 = "" Or srcSpec1 = "" Then GoTo NextRow1
        If Not IsExplicitRangeOrColumn(tgtSpec1) Then GoTo NextRow1

        Application.StatusBar = "MasterRoll Pass 1: " & sName1 & " (row " & i + 1 & ")"

        Dim ws1 As Worksheet : Set ws1 = Nothing
        On Error Resume Next
        Set ws1 = ThisWorkbook.Worksheets(sName1)
        On Error GoTo FailSafe
        If ws1 Is Nothing Then GoTo NextRow1

        ProcessCase1_ExplicitTarget ws1, srcSpec1, tgtSpec1

NextRow1:
        Set ws1 = Nothing
    Next i

    ' ══════════════════════════════════════════════════════════════════════════
    ' PASS 2 – Structural-roll rows  (Col C is blank, REST, WEST, or NEST)
    ' ══════════════════════════════════════════════════════════════════════════
    Application.StatusBar = "MasterRoll Pass 2 of 3: structural-roll rows..."

    For i = 1 To rowCount
        Dim sName2   As String : sName2   = Trim$(CStr(listData(i, 1)))
        Dim srcSpec2 As String : srcSpec2 = Trim$(CStr(listData(i, 2)))
        Dim tgtSpec2 As String : tgtSpec2 = UCase$(Trim$(CStr(listData(i, 3))))

        If sName2 = "" Or srcSpec2 = "" Then GoTo NextRow2
        If IsExplicitRangeOrColumn(listData(i, 3)) Then GoTo NextRow2

        Application.StatusBar = "MasterRoll Pass 2: " & sName2 & " – " & _
                                 IIf(tgtSpec2 = "", "INSERT RIGHT", tgtSpec2) & _
                                 " (row " & i + 1 & ")"

        Dim ws2 As Worksheet : Set ws2 = Nothing
        On Error Resume Next
        Set ws2 = ThisWorkbook.Worksheets(sName2)
        On Error GoTo FailSafe
        If ws2 Is Nothing Then GoTo NextRow2

        Select Case tgtSpec2
            Case ""     : ProcessCase2_InsertRight ws2, srcSpec2
            Case "REST" : ProcessCase3_UseRight    ws2, srcSpec2
            Case "WEST" : ProcessCase4_InsertLeft  ws2, srcSpec2
            Case "NEST" : ProcessCase5_UseLeft     ws2, srcSpec2
            ' Any unrecognised keyword is silently skipped.
        End Select

NextRow2:
        Set ws2 = Nothing
    Next i

    ' ══════════════════════════════════════════════════════════════════════════
    ' PASS 3 – Column deletions  (Col D)
    '
    ' Deletions run after all roll passes to avoid shifting column indices
    ' that are still needed in earlier passes.
    ' ══════════════════════════════════════════════════════════════════════════
    Application.StatusBar = "MasterRoll Pass 3 of 3: column deletions..."

    For i = 1 To rowCount
        Dim sName3  As String : sName3  = Trim$(CStr(listData(i, 1)))
        Dim delSpec As String : delSpec = Trim$(CStr(listData(i, 4)))

        If sName3 = "" Or delSpec = "" Then GoTo NextRow3

        Application.StatusBar = "MasterRoll Pass 3: delete '" & delSpec & "' on " & sName3

        Dim ws3 As Worksheet : Set ws3 = Nothing
        On Error Resume Next
        Set ws3 = ThisWorkbook.Worksheets(sName3)
        On Error GoTo FailSafe
        If ws3 Is Nothing Then GoTo NextRow3

        DeleteColumnSpec ws3, delSpec

NextRow3:
        Set ws3 = Nothing
    Next i

CleanExit:
    Application.CutCopyMode    = False
    Application.StatusBar      = oldStatus
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents   = oldEvents
    Application.Calculation    = oldCalc

    elapsed = ElapsedSec(startTime)
    If SHOW_ELAPSED Then
        MsgBox "MasterRoll completed in " & Format$(elapsed, "0.00") & " seconds.", vbInformation
    End If
    Exit Sub

FailSafe:
    Application.CutCopyMode    = False
    Application.StatusBar      = oldStatus
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents   = oldEvents
    Application.Calculation    = oldCalc
    MsgBox "MasterRoll encountered an error: " & Err.Description, vbExclamation

End Sub


'================================================================================
' CASE 1 – Explicit Target Column / Range  (F1 freeze rule applied to target)
'
' The source range is copied to an explicitly specified target range.
' The target becomes the prior-period opening-balance column.
'
' Freeze rule on the TARGET:
'   F1 source cells  Written as static values; external data is locked in.
'   F0 source cells  Formula transplanted via FormulaR1C1 so relative refs
'                    auto-adjust to the target column position.
'   Pre-existing internal formulas already in target  Preserved unchanged;
'   they belong to a previous structural layout and must not be lost.
'   The source column itself is never modified in Case 1.
'================================================================================
Private Sub ProcessCase1_ExplicitTarget( _
        ByVal ws      As Worksheet, _
        ByVal srcSpec As String, _
        ByVal tgtSpec As String)

    Dim srcRng As Range : Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    Dim tgtRng As Range : Set tgtRng = ResolveAndTrimRange(ws, tgtSpec)
    If srcRng Is Nothing Or tgtRng Is Nothing Then Exit Sub

    ' Source and target must be identical in shape.
    If srcRng.Rows.Count    <> tgtRng.Rows.Count Or _
       srcRng.Columns.Count <> tgtRng.Columns.Count Then Exit Sub

    ws.DisplayPageBreaks = False

    ' Build the whole-sheet F1 address set once; queried for every source formula.
    Dim f1Set As Object : Set f1Set = BuildF1Set(ws)

    ' Snapshot any internal formulas already present in the target so they can
    ' be restored after the mass-value overwrite in step B below.
    Dim keepFormula As Object : Set keepFormula = CreateObject("Scripting.Dictionary")
    Dim keepTarget  As Object : Set keepTarget  = CreateObject("Scripting.Dictionary")

    Dim tgtFC As Range
    On Error Resume Next : Set tgtFC = tgtRng.SpecialCells(xlCellTypeFormulas) : On Error GoTo 0

    If Not tgtFC Is Nothing Then
        Dim tc As Range
        For Each tc In tgtFC.Cells
            If Not IsDirectExternalFormula(tc.Formula, ws.Name) Then
                Dim rk1 As String : rk1 = RelKey(tc, tgtRng)
                keepTarget(rk1)  = True
                keepFormula(rk1) = tc.FormulaR1C1
            End If
        Next tc
    End If

    ' Mass-copy all source values to target in a single array write.
    ' Formula cells will be corrected individually in the next step.
    tgtRng.Value2 = srcRng.Value2

    ' Correct formula cells: F1 cells are frozen to values; F0 cells keep the formula.
    Dim srcFC As Range
    On Error Resume Next : Set srcFC = srcRng.SpecialCells(xlCellTypeFormulas) : On Error GoTo 0

    If Not srcFC Is Nothing Then
        Dim sc As Range, tgtCell As Range
        For Each sc In srcFC.Cells
            Dim rk2 As String : rk2 = RelKey(sc, srcRng)
            If keepTarget.Exists(rk2) Then GoTo SkipSrcCell

            Set tgtCell = tgtRng.Cells(sc.Row - srcRng.Row + 1, sc.Column - srcRng.Column + 1)

            If f1Set.Exists(UCase$(sc.Address(False, False))) Then
                tgtCell.Value2 = sc.Value2              ' F1: lock in the computed value
            Else
                tgtCell.FormulaR1C1 = sc.FormulaR1C1   ' F0: transplant; R1C1 auto-adjusts
            End If
SkipSrcCell:
        Next sc
    End If

    ' Restore any pre-existing internal formulas that were saved above.
    Dim k As Variant, pts() As String
    For Each k In keepFormula.Keys
        pts = Split(CStr(k), "|")
        tgtRng.Cells(CLng(pts(0)), CLng(pts(1))).FormulaR1C1 = keepFormula(k)
    Next k

    ClearCommentsNotes tgtRng   ' prior-period target does not need carry-forward notes

End Sub


'================================================================================
' CASE 2 – Insert New Columns to the RIGHT  (Col C is blank)
'
' A new column block equal in width to the source is inserted immediately to
' its right.  The inserted block becomes the new live current-period column
' (target).  The source becomes the prior-period OB column and is frozen.
'
' Critical sequencing detail:
'   Source content is captured in arrays BEFORE the insert.  Although the
'   source column indices do not shift (insert is to the right of source),
'   OFFSET-based formulas recalculate relative to their position and could
'   return incorrect formula text if read after nearby columns are modified.
'   Capturing pre-insert eliminates this risk entirely.
'================================================================================
Private Sub ProcessCase2_InsertRight( _
        ByVal ws      As Worksheet, _
        ByVal srcSpec As String)

    Dim srcRng As Range : Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    If srcRng Is Nothing Then Exit Sub

    ws.DisplayPageBreaks = False

    Dim colCount    As Long : colCount    = srcRng.Columns.Count
    Dim srcFirstCol As Long : srcFirstCol = srcRng.Column
    Dim srcLastCol  As Long : srcLastCol  = srcFirstCol + colCount - 1
    Dim srcFirstRow As Long : srcFirstRow = srcRng.Row
    Dim srcRowCount As Long : srcRowCount = srcRng.Rows.Count

    ' Capture source content before any structural changes.
    Dim arrVal  As Variant : arrVal  = srcRng.Value2
    Dim arrFmlA As Variant : arrFmlA = srcRng.Formula   ' A1 text used by column shifter

    ' Insert colCount blank columns immediately to the right of source.
    Dim insertAt As Long : insertAt = srcLastCol + 1
    Dim ci       As Long
    For ci = 1 To colCount
        ws.Columns(insertAt).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next ci

    ' Construct the target range in the newly created columns.
    Dim tgtRng As Range
    Set tgtRng = ws.Range( _
        ws.Cells(srcFirstRow, insertAt), _
        ws.Cells(srcFirstRow + srcRowCount - 1, insertAt + colCount - 1))

    ' Copy formats and column widths so the new column inherits the visual layout.
    CopyRangeFormats srcRng, tgtRng

    ' Write all formulas into the target with internal column refs shifted by colDelta.
    ' External sheet refs (e.g. 'Cover page'!E19) are left verbatim.
    WriteAllFormulasToRange arrVal, arrFmlA, tgtRng, srcFirstCol, ws.Name

    ' Freeze the source (now the prior-period OB column): F1 and F2 → values, F0 → formula.
    FreezSourceRange ws, srcRng

    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' CASE 3 – Reuse Existing Columns to the RIGHT  (Col C = "REST")
'
' The colCount columns immediately to the right of the source are used as the
' target.  They are unhidden/ungrouped if necessary and overwritten with the
' source content.  No column is inserted.
' Source is frozen (F1 + F2) identically to Case 2.
'================================================================================
Private Sub ProcessCase3_UseRight( _
        ByVal ws      As Worksheet, _
        ByVal srcSpec As String)

    Dim srcRng As Range : Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    If srcRng Is Nothing Then Exit Sub

    ws.DisplayPageBreaks = False

    Dim colCount    As Long : colCount    = srcRng.Columns.Count
    Dim tgtFirstCol As Long : tgtFirstCol = srcRng.Column + colCount
    Dim tgtLastCol  As Long : tgtLastCol  = tgtFirstCol + colCount - 1

    If tgtLastCol > ws.Columns.Count Then Exit Sub

    Dim tgtRng As Range
    Set tgtRng = ws.Range( _
        ws.Cells(srcRng.Row, tgtFirstCol), _
        ws.Cells(srcRng.Row + srcRng.Rows.Count - 1, tgtLastCol))

    Dim arrVal  As Variant : arrVal  = srcRng.Value2
    Dim arrFmlA As Variant : arrFmlA = srcRng.Formula

    UnhideUngroup ws, tgtFirstCol, tgtLastCol
    WriteAllFormulasToRange arrVal, arrFmlA, tgtRng, srcRng.Column, ws.Name
    FreezSourceRange ws, srcRng
    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' CASE 4 – Insert New Columns to the LEFT  (Col C = "WEST")
'
' A new column block is inserted immediately to the LEFT of the source.
' After the insert, the source shifts right by colCount positions; the newly
' inserted block (at the original source position) becomes the target.
'
' Because the insert shifts the source rightward, all source content must be
' captured BEFORE the insert.  After insertion the shifted source is frozen.
'================================================================================
Private Sub ProcessCase4_InsertLeft( _
        ByVal ws      As Worksheet, _
        ByVal srcSpec As String)

    Dim srcRng As Range : Set srcRng = ResolveAndTrimRange(ws, srcSpec)
    If srcRng Is Nothing Then Exit Sub

    ws.DisplayPageBreaks = False

    Dim colCount    As Long : colCount    = srcRng.Columns.Count
    Dim srcFirstCol As Long : srcFirstCol = srcRng.Column
    Dim srcLastRow  As Long : srcLastRow  = srcRng.Row + srcRng.Rows.Count - 1

    ' Capture source content before the insert shifts all column indices.
    Dim arrVal  As Variant : arrVal  = srcRng.Value2
    Dim arrFmlA As Variant : arrFmlA = srcRng.Formula

    ' Insert colCount blank columns to the left of source (repeated single inserts
    ' so that each successive column inherits format from the column to its left).
    Dim ci As Long
    For ci = 1 To colCount
        ws.Columns(srcFirstCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next ci

    ' After insertion layout:
    '   Target (new)  = srcFirstCol … srcFirstCol + colCount - 1
    '   Source (old)  = srcFirstCol + colCount … srcFirstCol + 2*colCount - 1
    Dim tgtRng As Range
    Set tgtRng = ws.Range( _
        ws.Cells(srcRng.Row, srcFirstCol), _
        ws.Cells(srcLastRow, srcFirstCol + colCount - 1))

    Dim newSrcRng As Range
    Set newSrcRng = ws.Range( _
        ws.Cells(srcRng.Row, srcFirstCol + colCount), _
        ws.Cells(srcLastRow, srcFirstCol + 2 * colCount - 1))

    ' Copy formats from the now-shifted source to the new target block.
    CopyRangeFormats newSrcRng, tgtRng

    ' Write formulas: srcFirstCol is the ORIGINAL source position so colDelta = 0
    ' (target lands exactly where source was; column refs need no shift).
    WriteAllFormulasToRange arrVal, arrFmlA, tgtRng, srcFirstCol, ws.Name

    FreezSourceRange ws, newSrcRng
    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' CASE 5 – Reuse Existing Columns to the LEFT  (Col C = "NEST")
'
' The colCount columns immediately to the LEFT of the source are used as the
' target.  They are unhidden/ungrouped and overwritten with source content.
' Source is frozen (F1 + F2) identically to Cases 2-4.
'================================================================================
Private Sub ProcessCase5_UseLeft( _
        ByVal ws      As Worksheet, _
        ByVal srcSpec As String)

    Dim srcRng As Range : Set srcRng = ResolveAndTrimRange(ws, srcSpec)
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

    Dim arrVal  As Variant : arrVal  = srcRng.Value2
    Dim arrFmlA As Variant : arrFmlA = srcRng.Formula

    UnhideUngroup ws, tgtFirstCol, tgtLastCol
    WriteAllFormulasToRange arrVal, arrFmlA, tgtRng, srcFirstCol, ws.Name
    FreezSourceRange ws, srcRng
    ClearCommentsNotes tgtRng

End Sub


'================================================================================
' COLUMN DELETION  (Col D)
'
' Deletes one or more columns from the specified worksheet.
' delSpec formats accepted:
'   Single column letter    "C"
'   Column range            "C:E"
'   Comma-separated list    "C,F,H:J"
'
' Columns are collected into an array, sorted descending, and deleted
' right-to-left so that removing one column does not renumber those
' yet to be deleted.  Duplicate indices are suppressed automatically.
'================================================================================
Private Sub DeleteColumnSpec(ByVal ws As Worksheet, ByVal delSpec As String)

    Dim parts()  As String : parts = Split(delSpec, ",")
    Dim colNums() As Long
    Dim cnt      As Long : cnt = 0
    Dim p        As Long
    Dim rng      As Range
    Dim col      As Long

    For p = 0 To UBound(parts)
        Dim spec As String : spec = Trim$(parts(p))
        If spec = "" Then GoTo NextPart

        Set rng = Nothing
        On Error Resume Next : Set rng = ResolveColSpec(ws, spec) : On Error GoTo 0

        If Not rng Is Nothing Then
            For col = rng.Column To rng.Column + rng.Columns.Count - 1
                cnt = cnt + 1
                ReDim Preserve colNums(1 To cnt)
                colNums(cnt) = col
            Next col
        End If
NextPart:
    Next p

    If cnt = 0 Then Exit Sub

    ' Insertion-sort descending (column lists are typically very short).
    Dim a As Long, b As Long, tmp As Long
    For a = 1 To cnt - 1
        For b = a + 1 To cnt
            If colNums(b) > colNums(a) Then
                tmp = colNums(a) : colNums(a) = colNums(b) : colNums(b) = tmp
            End If
        Next b
    Next a

    Dim lastDel As Long : lastDel = -1
    For a = 1 To cnt
        If colNums(a) <> lastDel Then
            On Error Resume Next : ws.Columns(colNums(a)).Delete : On Error GoTo 0
            lastDel = colNums(a)
        End If
    Next a

End Sub


'================================================================================
' FREEZE HELPERS
'================================================================================

'--------------------------------------------------------------------------------
' BuildF1Set
'   Scans every formula cell in the worksheet's used range and returns a
'   Scripting.Dictionary keyed by ADDR (uppercase, no $) → True for every cell
'   whose formula text directly references another sheet or external workbook.
'   This is the "F1 anchor set" used for both Case-1 target freeze decisions
'   and the one-hop F2 check in FreezSourceRange.
'   SpecialCells(xlCellTypeFormulas) limits the scan to formula cells only,
'   avoiding iteration over blank or value-only areas.
'--------------------------------------------------------------------------------
Private Function BuildF1Set(ByVal ws As Worksheet) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim fc As Range
    On Error Resume Next : Set fc = ws.UsedRange.SpecialCells(xlCellTypeFormulas) : On Error GoTo 0
    If fc Is Nothing Then Set BuildF1Set = result : Exit Function

    Dim c As Range
    For Each c In fc.Cells
        If IsDirectExternalFormula(c.Formula, ws.Name) Then
            result(UCase$(c.Address(False, False))) = True
        End If
    Next c

    Set BuildF1Set = result
End Function


'--------------------------------------------------------------------------------
' FreezSourceRange  (source freeze for Cases 2-5)
'
' Converts formula cells in the source range to static values where those
' formulas pull live external data (F1) or reference a cell that does (F2).
' Pure internal aggregation formulas (F0) are left as live formulas.
'
' Step 1  Build the whole-sheet F1 address set via BuildF1Set.
'         The set contains every formula cell on the sheet – not just in the
'         source column – that directly references an external sheet.  It is
'         built once per call and queried for every source formula cell.
'
' Step 2  Iterate every formula cell in the source range:
'
'   F1 (direct external):
'     IsDirectExternalFormula returns True for this cell's own formula text.
'     The cell is immediately converted to its current computed value.
'
'   F2 (one-hop indirect):
'     FormulaReferencesAnyInSet finds at least one bare same-sheet cell address
'     in the formula that exists in the whole-sheet F1 set.
'     Example: =ROUND(Q12-SUM(F12:OFFSET(N12,0,-1)),2) where Q12 is F1.
'     Q12 appears as a bare address in the formula text → the cell is F2 → frozen.
'     Only a single hop is checked; no further propagation is performed.
'
'   F0 (pure internal):
'     The formula references only same-sheet cells that are not in the F1 set.
'     Example: =ROUND(SUM(O10,O11,O12),2).
'     O10, O11, O12 are internal formulas, not direct-external → not in F1 set.
'     The cell is left as a live formula because its result can always be
'     recomputed correctly from whatever state the referenced cells now hold.
'--------------------------------------------------------------------------------
Private Sub FreezSourceRange( _
        ByVal ws  As Worksheet, _
        ByVal rng As Range)

    Dim f1Set As Object : Set f1Set = BuildF1Set(ws)

    Dim fc As Range
    On Error Resume Next : Set fc = rng.SpecialCells(xlCellTypeFormulas) : On Error GoTo 0
    If fc Is Nothing Then Exit Sub

    Dim c   As Range
    Dim fml As String
    For Each c In fc.Cells
        fml = c.Formula

        If IsDirectExternalFormula(fml, ws.Name) Then
            Dim v1 As Variant : v1 = c.Value2 : c.Value2 = v1   ' F1: lock value
            GoTo NextFreezeCell
        End If

        If FormulaReferencesAnyInSet(fml, ws.Name, f1Set) Then
            Dim v2 As Variant : v2 = c.Value2 : c.Value2 = v2   ' F2: lock value
        End If
        ' F0: no action – formula stays live

NextFreezeCell:
    Next c

End Sub


'--------------------------------------------------------------------------------
' FreezeRange_F2  (backward-compatibility alias)
'--------------------------------------------------------------------------------
Private Sub FreezeRange_F2(ByVal ws As Worksheet, ByVal rng As Range)
    FreezSourceRange ws, rng
End Sub


'================================================================================
' TARGET WRITE HELPERS
'================================================================================

'--------------------------------------------------------------------------------
' WriteAllFormulasToRange
'   Writes pre-captured source arrays into a target range, transplanting ALL
'   formulas (F0, F1, and F2) as live formulas so the target becomes the new
'   current-period column with full recalculation capability.
'
'   Column-reference shifting rule:
'     Internal (same-sheet) relative column refs shifted by colDelta.
'       =G12 in source col N, target col O (colDelta=1) → =H12 in target.
'       =$G12 (absolute column) → =$G12 unchanged.
'     External sheet refs left EXACTLY as written in source.
'       ='Cover page'!E19 → stays ='Cover page'!E19 regardless of colDelta.
'       =G12+'Cover page'!E19 → =H12+'Cover page'!E19 (G shifts, E19 does not).
'
'   Why not FormulaR1C1?
'     R1C1 encodes every relative ref as a row/column offset from the cell.
'     Writing to a new column using FormulaR1C1 correctly shifts internal refs,
'     BUT it also shifts external refs like 'Cover page'!E19 (stored internally
'     as a relative offset) to 'Cover page'!F19, which is wrong.
'     The A1 + custom column-shifter approach avoids this by skipping all
'     sheet-qualified address tokens during the shift pass.
'
'   Parameters:
'     arrVal      Value2 array captured from source range.
'     arrFmlA1    Formula (A1) array captured from source range.
'     tgtRng      Destination range (already inserted or pre-existing).
'     srcFirstCol Column index of the left edge of the original source range.
'     sheetName   Name of the host worksheet (for external-ref detection).
'--------------------------------------------------------------------------------
Private Sub WriteAllFormulasToRange( _
        ByVal arrVal      As Variant, _
        ByVal arrFmlA1    As Variant, _
        ByVal tgtRng      As Range, _
        ByVal srcFirstCol As Long, _
        ByVal sheetName   As String)

    ' colDelta: how many columns to the right the target is from the original source.
    ' Positive for Cases 2, 3; zero for Case 4 (target replaces source position);
    ' negative for Case 5 (target is to the left).
    Dim colDelta As Long : colDelta = tgtRng.Column - srcFirstCol

    Dim r  As Long, co As Long
    Dim fA As String, fS As String

    For r = 1 To tgtRng.Rows.Count
        For co = 1 To tgtRng.Columns.Count

            fA = ""
            If IsArray(arrFmlA1) Then
                If VarType(arrFmlA1(r, co)) = vbString Then fA = CStr(arrFmlA1(r, co))
            End If

            If Len(fA) > 1 And Left$(fA, 1) = "=" Then
                ' Shift only internal relative column refs; external refs are verbatim.
                fS = IIf(colDelta = 0, fA, ShiftInternalColRefs(fA, sheetName, colDelta))
                On Error Resume Next
                tgtRng.Cells(r, co).Formula = fS
                On Error GoTo 0
            Else
                If IsArray(arrVal) Then
                    On Error Resume Next
                    tgtRng.Cells(r, co).Value2 = arrVal(r, co)
                    On Error GoTo 0
                End If
            End If

        Next co
    Next r

End Sub


'--------------------------------------------------------------------------------
' ShiftInternalColRefs
'   Rewrites an A1-notation formula string, advancing ONLY internal (same-sheet)
'   relative column references by colDelta positions.
'   This is a single left-to-right character scan; no RegExp is used here
'   because this function is called once per formula cell in WriteAllFormulasToRange
'   and raw string performance matters more than pattern flexibility.
'
'   Tokens handled:
'     String literals ("...")
'       Copied verbatim; inner double-quotes (escaped as "") are preserved.
'     Quoted sheet qualifier ('SheetName'!addr)
'       Entire token (qualifier + "!" + address) copied verbatim.
'     Unquoted sheet qualifier (Sheet1!addr)
'       Entire token copied verbatim.
'     Bare cell address (no sheet qualifier)
'       Column part shifted by colDelta if relative (no leading "$").
'       Absolute columns ($G) are not shifted.
'     All other characters
'       Copied as-is.
'--------------------------------------------------------------------------------
Private Function ShiftInternalColRefs( _
        ByVal formula   As String, _
        ByVal sheetName As String, _
        ByVal colDelta  As Long) As String

    Dim n   As Long  : n   = Len(formula)
    Dim i   As Long  : i   = 1
    Dim out As String : out = ""
    Dim ch  As String

    Do While i <= n
        ch = Mid$(formula, i, 1)

        ' ── String literal: copy everything inside quotes verbatim ────────────
        If ch = """" Then
            out = out & ch : i = i + 1
            Do While i <= n
                ch = Mid$(formula, i, 1)
                out = out & ch : i = i + 1
                If ch = """" Then
                    ' A doubled quote ("") is an escaped literal; keep both and continue.
                    If i <= n And Mid$(formula, i, 1) = """" Then
                        out = out & Mid$(formula, i, 1) : i = i + 1
                    Else
                        Exit Do   ' true closing quote
                    End If
                End If
            Loop
            GoTo NextChar
        End If

        ' ── Quoted sheet qualifier: 'SheetName'!addr ─────────────────────────
        ' Scan for the matching closing single-quote, respecting '' escapes,
        ' then confirm a "!" follows to distinguish it from a stray apostrophe.
        If ch = "'" Then
            Dim qStart As Long : qStart = i
            Dim qEnd   As Long : qEnd   = i + 1
            Do While qEnd <= n
                If Mid$(formula, qEnd, 1) = "'" Then
                    If qEnd + 1 <= n And Mid$(formula, qEnd + 1, 1) = "'" Then
                        qEnd = qEnd + 2   ' skip '' escape
                    Else
                        Exit Do           ' found closing '
                    End If
                Else
                    qEnd = qEnd + 1
                End If
            Loop
            If qEnd <= n And Mid$(formula, qEnd, 1) = "'" Then
                If qEnd + 1 <= n And Mid$(formula, qEnd + 1, 1) = "!" Then
                    ' Confirmed sheet qualifier – emit qualifier + "!" + address verbatim.
                    out = out & Mid$(formula, qStart, qEnd - qStart + 2)
                    i = qEnd + 2
                    i = CopyAddressVerbatim(formula, i, out)
                    GoTo NextChar
                End If
            End If
            out = out & ch : i = i + 1   ' not a qualifier; emit bare apostrophe
            GoTo NextChar
        End If

        ' ── Unquoted identifier possibly followed by "!" (external ref) ───────
        If IsLetterAZ(ch) Then
            Dim idStart As Long : idStart = i
            Dim j       As Long : j       = i + 1
            Do While j <= n And IsSheetIdentChar(Mid$(formula, j, 1)) : j = j + 1 : Loop
            If j <= n And Mid$(formula, j, 1) = "!" Then
                ' External sheet ref: emit identifier + "!" + address verbatim.
                out = out & Mid$(formula, idStart, j - idStart + 1)
                i = j + 1
                i = CopyAddressVerbatim(formula, i, out)
                GoTo NextChar
            End If
            i = idStart   ' fall through to bare-address parser below
        End If

        ' ── Bare cell address: shift relative column by colDelta ──────────────
        If IsLetterAZ(ch) Or ch = "$" Then
            Dim aStart     As Long    : aStart     = i
            Dim colAbs     As Boolean : colAbs     = False
            Dim colLetters As String  : colLetters = ""
            Dim rowAbs     As Boolean : rowAbs     = False
            Dim rowDigits  As String  : rowDigits  = ""
            Dim k          As Long    : k          = i

            If Mid$(formula, k, 1) = "$" Then colAbs = True : k = k + 1

            Do While k <= n And IsLetterAZ(Mid$(formula, k, 1))
                colLetters = colLetters & Mid$(formula, k, 1) : k = k + 1
            Loop

            If Len(colLetters) >= 1 And Len(colLetters) <= 3 Then
                If k <= n And Mid$(formula, k, 1) = "$" Then rowAbs = True : k = k + 1

                Do While k <= n And IsDigit09(Mid$(formula, k, 1))
                    rowDigits = rowDigits & Mid$(formula, k, 1) : k = k + 1
                Loop

                If Len(rowDigits) > 0 Then
                    ' Belt-and-braces: reject if preceded by "!" (already handled above).
                    Dim prevCh As String
                    prevCh = IIf(aStart > 1, Mid$(formula, aStart - 1, 1), "")
                    If prevCh <> "!" Then
                        ' Shift only relative columns; leave absolute columns unchanged.
                        If Not colAbs Then
                            Dim newCol As Long
                            newCol = ColLettersToNum(UCase$(colLetters)) + colDelta
                            If newCol >= 1 And newCol <= 16384 Then
                                colLetters = NumToColLetters(newCol)
                            End If
                        End If
                        out = out & IIf(colAbs, "$", "") & colLetters & _
                                    IIf(rowAbs, "$", "") & rowDigits
                        i = k
                        GoTo NextChar
                    End If
                End If
            End If
        End If

        ' ── Default: emit the character unchanged ─────────────────────────────
        out = out & ch : i = i + 1

NextChar:
    Loop

    ShiftInternalColRefs = out
End Function


'--------------------------------------------------------------------------------
' CopyAddressVerbatim
'   After a sheet qualifier has been written to the output buffer, this helper
'   consumes and emits the following cell address (or column-range token such
'   as $A:$B) verbatim – no column shifting is performed.
'   Returns the new scan position immediately after the consumed token.
'--------------------------------------------------------------------------------
Private Function CopyAddressVerbatim( _
        ByVal formula As String, _
        ByVal i       As Long, _
        ByRef out     As String) As Long

    Dim n As Long : n = Len(formula)
    Dim k As Long : k = i

    If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1
    Do While k <= n And IsLetterAZ(Mid$(formula, k, 1))  : k = k + 1 : Loop
    If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1
    Do While k <= n And IsDigit09(Mid$(formula, k, 1))   : k = k + 1 : Loop

    ' Handle column-only ranges like $A:$B after a sheet qualifier.
    If k <= n And Mid$(formula, k, 1) = ":" Then
        k = k + 1
        If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1
        Do While k <= n And IsLetterAZ(Mid$(formula, k, 1)) : k = k + 1 : Loop
        If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1
        Do While k <= n And IsDigit09(Mid$(formula, k, 1))  : k = k + 1 : Loop
    End If

    out = out & Mid$(formula, i, k - i)
    CopyAddressVerbatim = k
End Function


'--------------------------------------------------------------------------------
' FormulaReferencesAnyInSet
'   Returns True if the formula text contains at least one bare (same-sheet)
'   cell address whose normalised form (uppercase, no $) exists as a key in
'   addrSet.  External sheet refs and string literals are excluded from matching.
'
'   Called by FreezSourceRange to perform the one-hop F2 classification:
'     addrSet = whole-sheet F1 address dictionary.
'     A hit means the source formula depends on a cell that pulls external data.
'
'   The scan logic mirrors ShiftInternalColRefs to ensure consistent handling
'   of string literals, quoted sheet qualifiers, and unquoted identifiers.
'--------------------------------------------------------------------------------
Private Function FormulaReferencesAnyInSet( _
        ByVal formula   As String, _
        ByVal sheetName As String, _
        ByVal addrSet   As Object) As Boolean

    Dim n  As Long : n = Len(formula)
    Dim i  As Long : i = 1
    Dim ch As String

    Do While i <= n
        ch = Mid$(formula, i, 1)

        ' Skip string literals.
        If ch = """" Then
            i = i + 1
            Do While i <= n
                Dim qch As String : qch = Mid$(formula, i, 1) : i = i + 1
                If qch = """" Then
                    If i <= n And Mid$(formula, i, 1) = """" Then
                        i = i + 1   ' "" escape
                    Else
                        Exit Do
                    End If
                End If
            Loop
            GoTo SkipCh
        End If

        ' Skip quoted sheet qualifier and the address following it.
        If ch = "'" Then
            i = i + 1
            Do While i <= n
                Dim qc As String : qc = Mid$(formula, i, 1) : i = i + 1
                If qc = "'" Then
                    If i <= n And Mid$(formula, i, 1) = "'" Then
                        i = i + 1   ' '' escape
                    Else
                        Exit Do
                    End If
                End If
            Loop
            If i <= n And Mid$(formula, i, 1) = "!" Then i = i + 1
            i = SkipCellAddr(formula, i)
            GoTo SkipCh
        End If

        ' Skip unquoted sheet qualifier (Identifier!) and its address.
        If IsLetterAZ(ch) Then
            Dim idS As Long : idS = i
            Dim j   As Long : j   = i + 1
            Do While j <= n And IsSheetIdentChar(Mid$(formula, j, 1)) : j = j + 1 : Loop
            If j <= n And Mid$(formula, j, 1) = "!" Then
                i = j + 1 : i = SkipCellAddr(formula, i)
                GoTo SkipCh
            End If
            i = idS   ' not external – fall through to bare-address parser
        End If

        ' Test bare cell address against addrSet.
        If IsLetterAZ(ch) Or ch = "$" Then
            Dim colL As String : colL = ""
            Dim rowD As String : rowD = ""
            Dim k    As Long   : k    = i

            If Mid$(formula, k, 1) = "$" Then k = k + 1
            Do While k <= n And IsLetterAZ(Mid$(formula, k, 1))
                colL = colL & Mid$(formula, k, 1) : k = k + 1
            Loop
            If Len(colL) >= 1 And Len(colL) <= 3 Then
                If k <= n And Mid$(formula, k, 1) = "$" Then k = k + 1
                Do While k <= n And IsDigit09(Mid$(formula, k, 1))
                    rowD = rowD & Mid$(formula, k, 1) : k = k + 1
                Loop
                If Len(rowD) > 0 Then
                    If addrSet.Exists(UCase$(colL) & rowD) Then
                        FormulaReferencesAnyInSet = True : Exit Function
                    End If
                    i = k
                    GoTo SkipCh
                End If
            End If
        End If

        i = i + 1
SkipCh:
    Loop

End Function


'--------------------------------------------------------------------------------
' SkipCellAddr
'   Advances the scan position past a cell address token (or A:B column range)
'   without emitting output.  Called after a sheet qualifier has been identified
'   to skip the external address without any processing.
'--------------------------------------------------------------------------------
Private Function SkipCellAddr(ByVal formula As String, ByVal i As Long) As Long
    Dim n As Long : n = Len(formula)
    If i <= n And Mid$(formula, i, 1) = "$" Then i = i + 1
    Do While i <= n And IsLetterAZ(Mid$(formula, i, 1))  : i = i + 1 : Loop
    If i <= n And Mid$(formula, i, 1) = "$" Then i = i + 1
    Do While i <= n And IsDigit09(Mid$(formula, i, 1))   : i = i + 1 : Loop
    If i <= n And Mid$(formula, i, 1) = ":" Then
        i = i + 1
        If i <= n And Mid$(formula, i, 1) = "$" Then i = i + 1
        Do While i <= n And IsLetterAZ(Mid$(formula, i, 1)) : i = i + 1 : Loop
        If i <= n And Mid$(formula, i, 1) = "$" Then i = i + 1
        Do While i <= n And IsDigit09(Mid$(formula, i, 1))  : i = i + 1 : Loop
    End If
    SkipCellAddr = i
End Function


'================================================================================
' FORMAT / STRUCTURE UTILITIES
'================================================================================

'--------------------------------------------------------------------------------
' CopyRangeFormats
'   Transfers number formats, borders, fill, font, conditional formatting,
'   data validation, and column widths from srcRng to tgtRng using Paste Special.
'   Values and formulas are not copied – only visual and validation properties.
'--------------------------------------------------------------------------------
Private Sub CopyRangeFormats(ByVal srcRng As Range, ByVal tgtRng As Range)
    On Error Resume Next
    srcRng.Copy
    tgtRng.PasteSpecial xlPasteFormats
    tgtRng.PasteSpecial xlPasteValidation
    Application.CutCopyMode = False
    Dim c As Long
    For c = 1 To srcRng.Columns.Count
        tgtRng.Columns(c).ColumnWidth = srcRng.Columns(c).ColumnWidth
    Next c
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------------------
' ClearCommentsNotes
'   Removes all threaded comments and legacy notes from the range.
'   Called on the target range after a roll so prior-period annotations do not
'   appear in the new current-period column.
'--------------------------------------------------------------------------------
Private Sub ClearCommentsNotes(ByVal rng As Range)
    On Error Resume Next : rng.ClearComments : rng.ClearNotes : On Error GoTo 0
End Sub


'--------------------------------------------------------------------------------
' UnhideUngroup
'   Makes a column range visible and removes any outline grouping.
'   Called before writing to REST or NEST targets that may have been grouped
'   or hidden at the end of the previous period.
'--------------------------------------------------------------------------------
Private Sub UnhideUngroup(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long)
    On Error Resume Next
    ws.Range(ws.Columns(firstCol), ws.Columns(lastCol)).Hidden = False
    ws.Range(ws.Columns(firstCol), ws.Columns(lastCol)).Ungroup
    On Error GoTo 0
End Sub


'================================================================================
' RANGE RESOLUTION UTILITIES
'================================================================================

'--------------------------------------------------------------------------------
' ResolveAndTrimRange
'   Parses spec into a Range on ws.  If the spec resolves to a full column
'   (e.g. "O"), the result is trimmed to the sheet's used row extent to prevent
'   Value2 array reads from spanning the full 1,048,576-row sheet height.
'--------------------------------------------------------------------------------
Private Function ResolveAndTrimRange(ByVal ws As Worksheet, ByVal spec As String) As Range
    Dim rng As Range
    On Error Resume Next : Set rng = ResolveColOrRange(ws, spec) : On Error GoTo 0
    If rng Is Nothing Then Exit Function

    If rng.Rows.Count = ws.Rows.Count Then
        Dim ur As Range : Set ur = ws.UsedRange
        If ur Is Nothing Then Exit Function
        Set ResolveAndTrimRange = ws.Range( _
            ws.Cells(ur.Row, rng.Column), _
            ws.Cells(ur.Row + ur.Rows.Count - 1, rng.Column + rng.Columns.Count - 1))
    Else
        Set ResolveAndTrimRange = rng
    End If
End Function


'--------------------------------------------------------------------------------
' ResolveColOrRange
'   Converts a spec string ("O", "O:P", "O2:P50", "$O$2:$P$50") to a Range.
'   Dollar signs are stripped to normalise absolute and relative notations.
'--------------------------------------------------------------------------------
Private Function ResolveColOrRange(ByVal ws As Worksheet, ByVal spec As String) As Range
    Dim s As String : s = Trim$(Replace(spec, "$", ""))
    If s = "" Then Exit Function
    On Error Resume Next
    If IsColumnLettersOnly(s) Or IsColumnRangeLettersOnly(s) Then
        Set ResolveColOrRange = ws.Columns(s)
    Else
        Set ResolveColOrRange = ws.Range(s)
    End If
    On Error GoTo 0
End Function


'--------------------------------------------------------------------------------
' ResolveColSpec
'   Converts a deletion spec ("C", "C:E") to a column Range.
'   Used exclusively by DeleteColumnSpec.
'--------------------------------------------------------------------------------
Private Function ResolveColSpec(ByVal ws As Worksheet, ByVal spec As String) As Range
    Dim s As String : s = Trim$(Replace(spec, "$", ""))
    If s = "" Then Exit Function
    On Error Resume Next : Set ResolveColSpec = ws.Columns(s) : On Error GoTo 0
End Function


'--------------------------------------------------------------------------------
' IsExplicitRangeOrColumn
'   Returns True if spec is a genuine column letter or range address (Case 1)
'   rather than a keyword (REST/WEST/NEST) or a blank value.
'   Logic: blank → False; keyword → False; valid Range() call → True;
'   fallback: pure letter string ("O") or letter:letter ("O:P") → True.
'--------------------------------------------------------------------------------
Private Function IsExplicitRangeOrColumn(ByVal spec As Variant) As Boolean
    Dim s As String : s = Trim$(CStr(spec))
    If s = "" Then Exit Function

    Dim u As String : u = UCase$(s)
    If u = "REST" Or u = "WEST" Or u = "NEST" Then Exit Function

    On Error Resume Next
    Dim rng As Range
    Set rng = ThisWorkbook.Worksheets(1).Range(Replace(s, "$", ""))
    On Error GoTo 0
    If Not rng Is Nothing Then IsExplicitRangeOrColumn = True : Exit Function

    IsExplicitRangeOrColumn = IsColumnLettersOnly(s) Or IsColumnRangeLettersOnly(s)
End Function


'--------------------------------------------------------------------------------
' IsColumnLettersOnly / IsColumnRangeLettersOnly
'   Classify a string as a pure column-letter token ("O") or a column-range
'   token ("O:P").  Used to route Columns() vs Range() in ResolveColOrRange.
'--------------------------------------------------------------------------------
Private Function IsColumnLettersOnly(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    s = Trim$(s)
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then Exit Function
    Next i
    IsColumnLettersOnly = True
End Function

Private Function IsColumnRangeLettersOnly(ByVal s As String) As Boolean
    Dim p As Long : p = InStr(s, ":")
    If p < 2 Then Exit Function
    IsColumnRangeLettersOnly = IsColumnLettersOnly(Left$(s, p - 1)) And _
                               IsColumnLettersOnly(Mid$(s, p + 1))
End Function


'--------------------------------------------------------------------------------
' RelKey
'   Returns a "row|col" string (1-based offset within baseRng) for use as a
'   dictionary key when tracking which target cells hold internal formulas that
'   must be preserved across the Case-1 mass-value overwrite.
'--------------------------------------------------------------------------------
Private Function RelKey(ByVal c As Range, ByVal baseRng As Range) As String
    RelKey = CStr(c.Row - baseRng.Row + 1) & "|" & CStr(c.Column - baseRng.Column + 1)
End Function


'================================================================================
' FORMULA CLASSIFICATION
'================================================================================

'--------------------------------------------------------------------------------
' IsDirectExternalFormula
'   Returns True if formulaText contains a sheet qualifier that names a sheet
'   other than currentSheetName, or names an external workbook ([Book.xlsx]).
'
'   A cached Static RegExp is used so the pattern is compiled only once per
'   macro run rather than on every call.  The pattern matches any token of the
'   form  qualifier!  where qualifier is a quoted sheet name, an external
'   workbook reference, or a plain identifier.
'
'   After matching, each qualifier is:
'     1. Checked for "[" → external workbook → True.
'     2. Stripped of enclosing single-quotes (with '' → ' unescaping).
'     3. Compared case-insensitively to currentSheetName.
'        Mismatch → True (references a different sheet on the same workbook).
'--------------------------------------------------------------------------------
Private Function IsDirectExternalFormula( _
        ByVal formulaText      As String, _
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

    If Not re.Test(formulaText) Then Exit Function   ' no "!" → no sheet qualifier

    Dim matches   As Object : Set matches = re.Execute(formulaText)
    Dim m         As Object
    Dim qualifier As String
    Dim cleaned   As String

    For Each m In matches
        qualifier = m.SubMatches(0)

        If InStr(qualifier, "[") > 0 Then
            IsDirectExternalFormula = True : Exit Function   ' external workbook
        End If

        cleaned = qualifier
        If Left$(cleaned, 1) = "'" And Right$(cleaned, 1) = "'" Then
            cleaned = Mid$(cleaned, 2, Len(cleaned) - 2)
            cleaned = Replace(cleaned, "''", "'")
        End If

        If StrComp(cleaned, currentSheetName, vbTextCompare) <> 0 Then
            IsDirectExternalFormula = True : Exit Function   ' different sheet
        End If
    Next m

End Function


'================================================================================
' CHARACTER AND COLUMN HELPERS
'================================================================================

' IsLetterAZ  True for ASCII letters A-Z / a-z.
Private Function IsLetterAZ(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim a As Integer : a = Asc(UCase$(ch))
    IsLetterAZ = (a >= 65 And a <= 90)
End Function

' IsDigit09  True for ASCII digits 0-9.
Private Function IsDigit09(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim a As Integer : a = Asc(ch)
    IsDigit09 = (a >= 48 And a <= 57)
End Function

' IsSheetIdentChar  True for characters valid inside an unquoted sheet name.
Private Function IsSheetIdentChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsSheetIdentChar = IsLetterAZ(ch) Or IsDigit09(ch) Or ch = "_" Or ch = "."
End Function

' IsColStartChar  True if ch can start a column letter token (same as IsLetterAZ).
Private Function IsColStartChar(ByVal ch As String) As Boolean
    IsColStartChar = IsLetterAZ(ch)
End Function

' ColLettersToNum  Converts column letters to a 1-based index ("A"=1, "AA"=27).
Private Function ColLettersToNum(ByVal letters As String) As Long
    Dim i As Long, v As Long
    For i = 1 To Len(letters)
        v = v * 26 + (Asc(UCase$(Mid$(letters, i, 1))) - 64)
    Next i
    ColLettersToNum = v
End Function

' NumToColLetters  Inverse of ColLettersToNum (1→"A", 27→"AA").
Private Function NumToColLetters(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop
    NumToColLetters = s
End Function


'================================================================================
' GENERAL UTILITIES
'================================================================================

'--------------------------------------------------------------------------------
' LastUsedRow
'   Returns the row index of the last cell containing data or a formula.
'   Returns 0 for a completely empty sheet.
'   Cells.Find with xlByRows / xlPrevious is faster than inspecting UsedRange
'   for sheets that have many trailing blank rows.
'--------------------------------------------------------------------------------
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find( _
        What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False)
    If f Is Nothing Then LastUsedRow = 0 Else LastUsedRow = f.Row
End Function


'--------------------------------------------------------------------------------
' ElapsedSec
'   Wall-clock elapsed seconds since t0 (from Timer).
'   Handles midnight rollover (Timer resets to 0 at 00:00:00).
'--------------------------------------------------------------------------------
Private Function ElapsedSec(ByVal t0 As Double) As Double
    Dim t1 As Double : t1 = Timer
    If t1 < t0 Then t1 = t1 + 86400#
    ElapsedSec = t1 - t0
End Function
