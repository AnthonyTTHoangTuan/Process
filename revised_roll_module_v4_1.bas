Option Explicit

' ╔═══════════════════════════════════════════════════════════════════════════╗
'  ROLL MODULE  v4.1  –  All v4 optimisations + critical correctness fixes
'
'  Cumulative optimisations (v1 → v4):
'   [OPT-1..OPT-12]  See v3 header for full list.
'   [OPT-13] Phase 3A formula writes batched into contiguous blocks.
'   [OPT-14] Pre-size formula cell arrays from UsedRange dimensions.
'   [OPT-15] Dictionary pooling in BuildFormulaLevelsForSheet.
'   [OPT-16] Fast IsFormulaVariant via AscW.
'   [OPT-17] ColumnLettersToNumberRaw for internal callers.
'
'  v4.1 NEW fixes (this file):
'   [BUG-4]  STALE LEVEL MAP: g_LevelMapCache REMOVED. Each roll inserts
'            or shifts columns, invalidating the dependency graph for any
'            subsequent roll on the same sheet. The level map is now rebuilt
'            fresh per RollOneSheetOneColumn call. Correctness > speed.
'   [BUG-5]  SINGLE-QUOTED SHEET NAMES: Added SkipSingleQuotedSheetQualifierByte
'            for the byte-array fast paths. BuildFrozenTargetFormula and
'            FillExplicitRefsInternal now correctly skip ='Sheet Name'!
'            qualifiers instead of mis-parsing their contents as references.
'   [BUG-6]  vbError CRASH: BuildFormulaLevelsForSheet now guards against
'            error-value cells (#N/A, #REF!, etc.) in the UsedRange.Formula
'            array. CStr() on a vbError variant caused runtime errors.
'   [BUG-7]  SINGLE-CELL SCALAR: Added Ensure2DColumnArray() helper.
'            srcRng.Formula/.FormulaR1C1/.Value2 return scalars when
'            lastRow = 1. Code now wraps them into 2-D arrays consistently.
'   [BUG-8]  ProcessOneControlRow now raises explicit errors for missing
'            worksheets and invalid column specs instead of silently exiting.
'   [MAINT]  Large source arrays released after Phase 3C to free memory.
'
'  Prior bug fixes (v4):
'   [BUG-1]  CleanFail captures Err before On Error Resume Next.
'   [BUG-2]  funcStack enlarged to 128 with overflow warning.
'   [BUG-3]  ExtractExplicitRefsInternal recursive dict pooling.
'
'  Risk mitigations (from v2, unchanged):
'   • g_SafeMode flag  →  EnableSafeMode True/False
'   • _Safe fallback copies of all byte-optimised functions
'   • RunSelfTests_RollParser() — 35-case self-test suite
' ╚═══════════════════════════════════════════════════════════════════════════╝


' ── Level constants ──────────────────────────────────────────────────────────
Private Const LVL_INTERNAL               As Long = 0
Private Const LVL_DIRECT_OTHER           As Long = 1
Private Const LVL_PARENT_OF_DIRECT_OTHER As Long = 2
Private Const LVL_HARDCODED              As Long = 3

' ── Debug toggle (False = silent in production) ──────────────────────────────
Private Const ROLL_DBG_PRINT As Boolean = False   ' [minor] guard Debug.Print

' ── [OPT-1] Column-letter cache ──────────────────────────────────────────────
Private g_colCache(1 To 16384) As String

' ── [OPT-6] Shared reference dictionary (reused with RemoveAll) ─────────────
Private g_RefDict As Object

' ── [BUG-4] g_LevelMapCache REMOVED in v4.1 — stale after column inserts ──

' ── [OPT-15] Pooled dictionaries for BuildFormulaLevelsForSheet ─────────────
Private g_PoolDictA As Object   ' used as "result" (level map output)
Private g_PoolDictB As Object   ' used as "dictIndex" (addr→index lookup)

' ── Safety toggle ────────────────────────────────────────────────────────────
Private g_SafeMode As Boolean

' ── ASCII byte constants ──────────────────────────────────────────────────────
Private Const BC_DQUOTE As Byte = 34    '"
Private Const BC_SQUOTE As Byte = 39    ''
Private Const BC_BANG   As Byte = 33    '!
Private Const BC_DOLLAR As Byte = 36    '$
Private Const BC_COLON  As Byte = 58    ':
Private Const BC_COMMA  As Byte = 44    ',
Private Const BC_SEMI   As Byte = 59    ';
Private Const BC_LPAREN As Byte = 40    '(
Private Const BC_RPAREN As Byte = 41    ')
Private Const BC_LBRACK As Byte = 91    '[
Private Const BC_RBRACK As Byte = 93    ']
Private Const BC_SPACE  As Byte = 32    ' space
Private Const BC_EQUALS As Byte = 61    '=
Private Const BC_DOT    As Byte = 46    '.
Private Const BC_UNDER  As Byte = 95    '_
Private Const BC_PLUS   As Byte = 43    '+
Private Const BC_MINUS  As Byte = 45    '-
Private Const BC_CARET  As Byte = 94    '^
Private Const BC_AMP    As Byte = 38    '&
Private Const BC_STAR   As Byte = 42    '*
Private Const BC_SLASH  As Byte = 47    '/
Private Const BC_LBRACE As Byte = 123   '{
Private Const BC_GT     As Byte = 62    '>
Private Const BC_LT     As Byte = 60    '<

' ── [BUG-2] Increased funcStack size ─────────────────────────────────────────
Private Const FUNC_STACK_SIZE As Long = 128


' ═══════════════════════════════════════════════════════════════════════════════
'  PUBLIC CONTROL SUBS
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub EnableSafeMode(ByVal enable As Boolean)
    g_SafeMode = enable
    Debug.Print "RollModule: " & IIf(enable, "SAFE MODE (string)", "FAST MODE (byte-array)") & " active."
End Sub

Public Sub RollReportingPacks_FromControl()

    Dim ctlWs           As Worksheet
    Dim lastCtlRow      As Long
    Dim r               As Long
    Dim startTick       As Double
    Dim secs            As Double

    Dim oldCalc         As XlCalculation
    Dim oldScreen       As Boolean
    Dim oldEvents       As Boolean
    Dim oldStatusBar    As Variant
    Dim activeSheetName As String
    Dim ctlData         As Variant   ' [OPT-2]

    On Error GoTo CleanFail

    startTick = Timer
    Set ctlWs = ActiveSheet
    activeSheetName = ctlWs.Name

    oldCalc      = Application.Calculation
    oldScreen    = Application.ScreenUpdating
    oldEvents    = Application.EnableEvents
    oldStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual

    ' [BUG-4] Level map cache removed in v4.1. Each roll rebuilds the graph
    ' because column inserts invalidate previously cached dependency maps.

    lastCtlRow = LastUsedRowAny(ctlWs)
    If lastCtlRow < 2 Then GoTo CleanExit

    ' [OPT-2] One COM call reads all control rows into a 2-D array.
    ctlData = ctlWs.Range(ctlWs.Cells(2, 1), ctlWs.Cells(lastCtlRow, 5)).Value2

    For r = 2 To lastCtlRow
        Application.StatusBar = "Rolling control row " & r & " of " & lastCtlRow & _
                                " [" & Trim$(CStr(ctlData(r - 1, 1))) & "]..."
        ProcessOneControlRow _
            targetSheetName := Trim$(CStr(ctlData(r - 1, 1))), _
            colSpec         := Trim$(CStr(ctlData(r - 1, 2))), _
            methodText      := Trim$(CStr(ctlData(r - 1, 3))), _
            freezeMaxText   := Trim$(CStr(ctlData(r - 1, 4))), _
            directionText   := Trim$(CStr(ctlData(r - 1, 5)))
    Next r

CleanExit:
    secs = ElapsedSeconds(startTick)
    RestoreAppState oldCalc, oldScreen, oldEvents, oldStatusBar
    MsgBox "Completed on control sheet [" & activeSheetName & "]." & vbCrLf & _
           "Elapsed time: " & Format(secs, "0.00") & " seconds", vbInformation
    Exit Sub

CleanFail:
    ' [BUG-1] Capture error BEFORE On Error Resume Next clears it.
    Dim errNum  As Long:   errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description

    secs = ElapsedSeconds(startTick)
    RestoreAppState oldCalc, oldScreen, oldEvents, oldStatusBar

    Dim errName As String
    On Error Resume Next
    If IsArray(ctlData) Then
        Dim si As Long: si = r - 1
        If si >= 1 And si <= UBound(ctlData, 1) Then errName = Trim$(CStr(ctlData(si, 1)))
    End If
    If Len(errName) = 0 Then errName = Trim$(CStr(ctlWs.Cells(r, 1).Value))
    On Error GoTo 0
    MsgBox "Error on control row " & r & " (" & errName & ")." & vbCrLf & _
           errDesc & vbCrLf & _
           "Elapsed time: " & Format(secs, "0.00") & " seconds", vbExclamation
End Sub


' ═══════════════════════════════════════════════════════════════════════════════
'  SELF-TEST SUITE  (unchanged from v2)
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub RunSelfTests_RollParser()

    Dim pass As Long, fail As Long
    Dim prevSafe As Boolean: prevSafe = g_SafeMode

    Debug.Print String(60, "=")
    Debug.Print "RunSelfTests_RollParser  " & Now()
    Debug.Print String(60, "=")

    Dim tests As Variant
    tests = Array( _
        Array("=A1+B2", "Sheet1", 1, "=B1+C2"), _
        Array("=$A1", "Sheet1", 1, "=$A1"), _
        Array("=A$1", "Sheet1", 1, "=B$1"), _
        Array("=$A$1", "Sheet1", 1, "=$A$1"), _
        Array("=SUM(A1:G1)", "Sheet1", 1, "=SUM(B1:H1)"), _
        Array("=SUM($A1:$G1)", "Sheet1", 1, "=SUM($A1:$G1)"), _
        Array("=Sheet1!A1+B2", "Sheet1", 1, "=Sheet1!B1+C2"), _
        Array("='Sheet1'!A1", "Sheet1", 1, "='Sheet1'!B1"), _
        Array("='Jan 2026'!C3", "Sheet1", 1, "='Jan 2026'!C3"), _
        Array("=OtherSheet!C3+A1", "Sheet1", 1, "=OtherSheet!C3+B1"), _
        Array("=IF(A1=""Yes"",1,0)", "Sheet1", 1, "=IF(B1=""Yes"",1,0)"), _
        Array("=IF(A1="""",0,A1)", "Sheet1", 1, "=IF(B1="""",0,B1)"), _
        Array("=A1+100", "Sheet1", 1, "=B1+100"), _
        Array("=ROUND(A1-0.2,2)", "Sheet1", 1, "=ROUND(B1-0.2,2)"), _
        Array("=SUM(A:A)", "Sheet1", 1, "=SUM(B:B)"), _
        Array("=SUM($A:$A)", "Sheet1", 1, "=SUM($A:$A)"), _
        Array("=A1", "Sheet1", -1, "=Z1"), _
        Array("=B1", "Sheet1", -1, "=A1"), _
        Array("=$B1", "Sheet1", -1, "=$B1"), _
        Array("=INDIRECT(""A1"")", "Sheet1", 1, "=INDIRECT(""A1"")"), _
        Array("='Sheet1'!A1+'Sheet1'!B2", "Sheet1", 1, "='Sheet1'!B1+'Sheet1'!C2"), _
        Array("=A1&"" text ""&B1", "Sheet1", 1, "=B1&"" text ""&C1"), _
        Array("=Z1", "Sheet1", 1, "=AA1"), _
        Array("=AA1", "Sheet1", 1, "=AB1"), _
        Array("=XFD1", "Sheet1", -1, "=XFC1"), _
        Array("='Sheet A1'!C3+A1", "Sheet1", 1, "='Sheet A1'!C3+B1"), _
        Array("='B2 Data'!D4+B2", "Sheet1", 1, "='B2 Data'!D4+C2"), _
        Array("='It''s A1'!E5+A1", "Sheet1", 1, "='It''s A1'!E5+B1"), _
        Array("='Sheet1'!A1+100", "Sheet1", 1, "='Sheet1'!B1+100"), _
        Array("='Sheet1'!$A1+B1", "Sheet1", 1, "='Sheet1'!$A1+C1"), _
        Array("='Multi Word Sheet'!A1+'Multi Word Sheet'!B1", "Sheet1", 1, _
              "='Multi Word Sheet'!A1+'Multi Word Sheet'!B1"), _
        Array("='Sheet1'!A1+'Sheet1'!B1+C1", "Sheet1", 1, _
              "='Sheet1'!B1+'Sheet1'!C1+D1") _
    )

    Dim i As Long
    For i = 0 To UBound(tests)
        Dim formula As String:  formula  = CStr(tests(i)(0))
        Dim host    As String:  host     = CStr(tests(i)(1))
        Dim delta   As Long:    delta    = CLng(tests(i)(2))
        Dim expected As String: expected = CStr(tests(i)(3))

        g_SafeMode = False
        Dim rFast As String: rFast = BuildFrozenTargetFormula(formula, host, delta)
        g_SafeMode = True
        Dim rSafe As String: rSafe = BuildFrozenTargetFormula(formula, host, delta)

        If rFast = expected And rSafe = expected Then
            pass = pass + 1
            Debug.Print "  PASS [" & i & "] " & formula
        Else
            fail = fail + 1
            Debug.Print "  FAIL [" & i & "] " & formula
            Debug.Print "       Expected : " & expected
            Debug.Print "       Fast got : " & rFast
            Debug.Print "       Safe got : " & rSafe
        End If
    Next i

    Debug.Print String(40, "-")
    Dim hcTests As Variant
    hcTests = Array( _
        Array("=A1+100", True), _
        Array("=A1+B2", False), _
        Array("=SUM(A1:G1)", False), _
        Array("=A1-0.2", True), _
        Array("=IF(A1=0,1,A2)", True), _
        Array("=IF(A1<5,A1,A2)", True), _
        Array("=SUM(A1,B2)", False), _
        Array("=A1*B1+C1", False), _
        Array("=100", True), _
        Array("=""hello""", False), _
        Array("='Data Sheet'!A1+100", True), _
        Array("='A1 B2'!C3", False), _
        Array("='Sheet 99'!A1+B1", False), _
        Array("='It''s A1'!E5-0.5", True) _
    )

    Dim j As Long
    For j = 0 To UBound(hcTests)
        Dim hcF As String:  hcF = CStr(hcTests(j)(0))
        Dim hcEx As Boolean: hcEx = CBool(hcTests(j)(1))

        g_SafeMode = False
        Dim hcFast As Boolean: hcFast = FormulaHasHardCodedNumbers(hcF)
        g_SafeMode = True
        Dim hcSafe As Boolean: hcSafe = FormulaHasHardCodedNumbers(hcF)

        If hcFast = hcEx And hcSafe = hcEx Then
            pass = pass + 1
            Debug.Print "  PASS [HC" & j & "] " & hcF
        Else
            fail = fail + 1
            Debug.Print "  FAIL [HC" & j & "] " & hcF
            Debug.Print "       Expected : " & hcEx
            Debug.Print "       Fast got : " & hcFast
            Debug.Print "       Safe got : " & hcSafe
        End If
    Next j

    g_SafeMode = prevSafe
    Debug.Print String(60, "=")
    Debug.Print "RESULT: " & pass & " PASSED, " & fail & " FAILED"
    If fail = 0 Then
        Debug.Print ">>> ALL TESTS PASSED <<<"
    Else
        Debug.Print ">>> FAILURES – run EnableSafeMode(True) until fixed <<<"
    End If
    Debug.Print String(60, "=")
End Sub


' ═══════════════════════════════════════════════════════════════════════════════
'  PROCESS ONE CONTROL ROW  (unchanged)
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub ProcessOneControlRow( _
    ByVal targetSheetName As String, _
    ByVal colSpec         As String, _
    ByVal methodText      As String, _
    ByVal freezeMaxText   As String, _
    ByVal directionText   As String)

    Dim ws             As Worksheet
    Dim srcCol         As Long
    Dim freezeMaxLevel As Long
    Dim isReverse      As Boolean

    On Error GoTo SafeExit

    If Len(targetSheetName) = 0 Then Exit Sub
    If Len(colSpec) = 0 Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(targetSheetName)
    On Error GoTo SafeExit
    ' [BUG-8] Raise an explicit error instead of silently skipping.
    If ws Is Nothing Then
        Err.Raise vbObjectError + 7001, "ProcessOneControlRow", _
                  "Worksheet not found: [" & targetSheetName & "]"
    End If

    srcCol = ParseColumnSpec(colSpec)
    ' [BUG-8] Raise an explicit error instead of silently skipping.
    If srcCol < 1 Or srcCol > 16384 Then
        Err.Raise vbObjectError + 7002, "ProcessOneControlRow", _
                  "Invalid column spec: [" & colSpec & "]"
    End If

    freezeMaxLevel = ParseFreezeMaxLevel(freezeMaxText)
    isReverse = (UCase$(Trim$(directionText)) = "REVERSE")

    RollOneSheetOneColumn ws, srcCol, methodText, freezeMaxLevel, isReverse
    Exit Sub

SafeExit:
    Debug.Print "ProcessOneControlRow failed. Sheet=[" & targetSheetName & _
                "], ColSpec=[" & colSpec & "], Method=[" & methodText & _
                "], Direction=[" & directionText & "]. " & _
                Err.Number & " - " & Err.Description
End Sub


' ═══════════════════════════════════════════════════════════════════════════════
'  ROLL ONE SHEET / ONE COLUMN
'  [OPT-3]  Three-phase write (unchanged structure)
'  [BUG-4]  Level map rebuilt fresh per call (cache removed in v4.1)
'  [OPT-8]  Phase 3 uses WriteColumnContiguousBlocks for value/source writes
'  [OPT-9]  .Value2 throughout
'  [OPT-13] Phase 3A formula writes now batched via WriteColumnContiguousFormulas
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub RollOneSheetOneColumn( _
    ByVal ws             As Worksheet, _
    ByVal srcCol         As Long, _
    ByVal methodText     As String, _
    ByVal freezeMaxLevel As Long, _
    ByVal isReverse      As Boolean)

    Const RT_R1C1     As Integer = 0
    Const RT_VALUE    As Integer = 1
    Const RT_FROZEN   As Integer = 2
    Const RT_HARDCODE As Integer = 3

    Dim levelMap        As Object
    Dim lastRow         As Long
    Dim tgtCol          As Long
    Dim workSrcCol      As Long
    Dim methodUpper     As String
    Dim colDelta        As Long

    Dim srcRng          As Range
    Dim arrFormula      As Variant
    Dim arrFormulaR1C1  As Variant
    Dim arrValue2       As Variant    ' [OPT-9] Value2 instead of Value

    Dim i               As Long
    Dim addrOriginal    As String
    Dim lvl             As Long
    Dim isFormula       As Boolean
    Dim srcColLetters   As String
    Dim wsNameUpper     As String    ' [minor] cached once

    Dim rowClass()       As Integer
    Dim frozenFormulas() As String
    Dim flagValue()      As Boolean  ' [OPT-8]  marks RT_VALUE rows for bulk write
    Dim flagSrc()        As Boolean  ' [OPT-8]  marks frozen-source rows for bulk write
    Dim flagFormula()    As Boolean  ' [OPT-13] marks RT_FROZEN/RT_HARDCODE rows for bulk formula write
    Dim overrideFormulas() As String ' [OPT-13] stores the formula to write for each flagged row

    methodUpper  = UCase$(Trim$(methodText))
    If Len(methodUpper) = 0 Then methodUpper = "INSERT"
    wsNameUpper  = UCase$(ws.Name)   ' [minor] cache

    If isReverse Then
        If srcCol = 1 Then Exit Sub
        colDelta = -1
    Else
        If srcCol = 16384 Then Exit Sub
        colDelta = 1
    End If

    ' [BUG-4] Always rebuild the level map. Column inserts by prior control
    ' rows invalidate cached dependency graphs for the same sheet.
    Set levelMap = BuildFormulaLevelsForSheet(ws)

    lastRow = LastUsedRowAny(ws)
    If lastRow < 1 Then lastRow = 1

    ' Read source column into arrays BEFORE any structural column change.
    ' [BUG-7] Ensure2DColumnArray guards against single-cell scalar returns.
    Set srcRng = ws.Range(ws.Cells(1, srcCol), ws.Cells(lastRow, srcCol))
    arrFormula      = Ensure2DColumnArray(srcRng.Formula)
    arrFormulaR1C1  = Ensure2DColumnArray(srcRng.FormulaR1C1)
    arrValue2       = Ensure2DColumnArray(srcRng.Value2)
    srcColLetters   = ColumnNumberToLetters(srcCol)

    ' ── Structural: insert / ungroup ──────────────────────────────────────────
    If Not isReverse Then
        workSrcCol = srcCol
        tgtCol     = srcCol + 1
        Select Case methodUpper
            Case "INSERT", "I"
                ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Case "UNGROUP", "U"
                PrepareExistingTargetColumn ws, tgtCol
            Case Else
                Debug.Print "Unknown method [" & methodText & "] sheet [" & ws.Name & "]. Defaulted INSERT."
                ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End Select
    Else
        If methodUpper = "INSERT" Or methodUpper = "I" Then
            tgtCol     = srcCol
            ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            workSrcCol = srcCol + 1
        ElseIf methodUpper = "UNGROUP" Or methodUpper = "U" Then
            workSrcCol = srcCol
            tgtCol     = srcCol - 1
            PrepareExistingTargetColumn ws, tgtCol
        Else
            Debug.Print "Unknown method [" & methodText & "] sheet [" & ws.Name & "]. Defaulted INSERT."
            tgtCol     = srcCol
            ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            workSrcCol = srcCol + 1
        End If
    End If

    ' [OPT-10] Copy layout for used rows only (not entire 1M-row column).
    CopyColumnLayoutNoComments ws, workSrcCol, tgtCol, lastRow

    ' ── PHASE 1: classify rows (zero COM calls) ───────────────────────────────
    ReDim rowClass(1 To lastRow)
    ReDim frozenFormulas(1 To lastRow)
    ReDim flagValue(1 To lastRow)
    ReDim flagSrc(1 To lastRow)
    ReDim flagFormula(1 To lastRow)        ' [OPT-13]
    ReDim overrideFormulas(1 To lastRow)   ' [OPT-13]

    For i = 1 To lastRow
        isFormula    = IsFormulaVariant(arrFormula(i, 1))   ' [OPT-16] fast AscW check
        addrOriginal = srcColLetters & CStr(i)

        If isFormula Then
            If levelMap.Exists(addrOriginal) Then
                lvl = CLng(levelMap(addrOriginal))
            Else
                lvl = LVL_INTERNAL
            End If

            If lvl = LVL_HARDCODED And freezeMaxLevel >= LVL_HARDCODED Then
                rowClass(i) = RT_HARDCODE
                ' [OPT-13] Queue original formula for batch write
                flagFormula(i) = True
                overrideFormulas(i) = CStr(arrFormula(i, 1))
            ElseIf ShouldFreezeLevel(lvl, freezeMaxLevel) Then
                rowClass(i)       = RT_FROZEN
                frozenFormulas(i) = BuildFrozenTargetFormula( _
                    CStr(arrFormula(i, 1)), ws.Name, colDelta)
                flagSrc(i) = True   ' [OPT-8] flag for bulk source write
                ' [OPT-13] Queue frozen formula for batch write
                flagFormula(i) = True
                overrideFormulas(i) = frozenFormulas(i)
            Else
                rowClass(i) = RT_R1C1
            End If
        Else
            rowClass(i) = RT_VALUE
            If VarType(arrValue2(i, 1)) <> vbEmpty Then
                flagValue(i) = True   ' [OPT-8] flag for bulk value write
            End If
        End If
    Next i

    ' ── PHASE 2: single bulk FormulaR1C1 write (all rows, 1 COM call) ─────────
    Application.StatusBar = "Rolling [" & ws.Name & "] — bulk-writing " & lastRow & " rows..."
    ws.Range(ws.Cells(1, tgtCol), ws.Cells(lastRow, tgtCol)).FormulaR1C1 = arrFormulaR1C1

    ' ── PHASE 3A: [OPT-13] batch formula overrides for RT_FROZEN & RT_HARDCODE ─
    ' Instead of per-cell writes, contiguous flagged rows are flushed as array blocks.
    WriteColumnContiguousFormulas ws, tgtCol, overrideFormulas, flagFormula, lastRow

    ' ── PHASE 3B: [OPT-8] bulk value writes for RT_VALUE rows ────────────────
    WriteColumnContiguousBlocks ws, tgtCol, arrValue2, flagValue, lastRow

    ' ── PHASE 3C: [OPT-8] bulk source value updates for frozen rows ──────────
    WriteColumnContiguousBlocks ws, workSrcCol, arrValue2, flagSrc, lastRow

    ' [MAINT] Release large arrays to free memory.
    Erase arrFormula
    Erase arrFormulaR1C1
    Erase arrValue2
End Sub


' ═══════════════════════════════════════════════════════════════════════════════
'  [OPT-8] WRITE COLUMN IN CONTIGUOUS BLOCKS  (values)
'  Writes srcArr(i, 1) to ws.Cells(i, col) only where flagArr(i) = True.
'  Consecutive flagged rows are batched into a single Range.Value2 call.
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub WriteColumnContiguousBlocks( _
    ByVal ws       As Worksheet, _
    ByVal col      As Long, _
    ByRef srcArr   As Variant, _
    ByRef flagArr() As Boolean, _
    ByVal lastRow  As Long)

    Dim runStart As Long: runStart = 0
    Dim i        As Long
    Dim k        As Long
    Dim runLen   As Long
    Dim blk      As Variant

    For i = 1 To lastRow + 1
        Dim inRun As Boolean
        If i <= lastRow Then inRun = flagArr(i) Else inRun = False

        If inRun Then
            If runStart = 0 Then runStart = i
        ElseIf runStart > 0 Then
            ' Flush contiguous block [runStart, i-1]
            runLen = i - runStart
            ReDim blk(1 To runLen, 1 To 1)
            For k = 1 To runLen
                blk(k, 1) = srcArr(runStart + k - 1, 1)
            Next k
            ws.Range(ws.Cells(runStart, col), _
                     ws.Cells(runStart + runLen - 1, col)).Value2 = blk   ' [OPT-9]
            runStart = 0
        End If
    Next i
End Sub


' ═══════════════════════════════════════════════════════════════════════════════
'  [OPT-13] WRITE COLUMN IN CONTIGUOUS BLOCKS  (formulas)
'  Writes fmlArr(i) as .Formula to ws.Cells(i, col) where flagArr(i) = True.
'  Consecutive flagged rows are batched into a single Range.Formula = array call.
'  This replaces the per-cell Phase 3A loop from v3.
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub WriteColumnContiguousFormulas( _
    ByVal ws       As Worksheet, _
    ByVal col      As Long, _
    ByRef fmlArr() As String, _
    ByRef flagArr() As Boolean, _
    ByVal lastRow  As Long)

    Dim runStart As Long: runStart = 0
    Dim i        As Long
    Dim k        As Long
    Dim runLen   As Long
    Dim blk      As Variant
    Dim inRun    As Boolean

    For i = 1 To lastRow + 1
        If i <= lastRow Then inRun = flagArr(i) Else inRun = False

        If inRun Then
            If runStart = 0 Then runStart = i
        ElseIf runStart > 0 Then
            ' Flush contiguous block [runStart, i-1]
            runLen = i - runStart
            If runLen = 1 Then
                ' Single cell: direct write avoids array overhead
                ws.Cells(runStart, col).Formula = fmlArr(runStart)
            Else
                ' Multi-cell: batch write via 2-D array
                ReDim blk(1 To runLen, 1 To 1)
                For k = 1 To runLen
                    blk(k, 1) = fmlArr(runStart + k - 1)
                Next k
                ws.Range(ws.Cells(runStart, col), _
                         ws.Cells(runStart + runLen - 1, col)).Formula = blk
            End If
            runStart = 0
        End If
    Next i
End Sub


' ═══════════════════════════════════════════════════════════════════════════════
'  LAYOUT / APP STATE HELPERS
' ═══════════════════════════════════════════════════════════════════════════════

' [OPT-10] Copy only rows 1:lastRow instead of the entire 1M-row column.
Private Sub CopyColumnLayoutNoComments( _
    ByVal ws      As Worksheet, _
    ByVal srcCol  As Long, _
    ByVal tgtCol  As Long, _
    ByVal lastRow As Long)

    On Error Resume Next
    ' [OPT-10] Limit to used rows
    ws.Range(ws.Cells(1, srcCol), ws.Cells(lastRow, srcCol)).Copy
    ws.Range(ws.Cells(1, tgtCol), ws.Cells(lastRow, tgtCol)).PasteSpecial xlPasteFormats
    ws.Range(ws.Cells(1, tgtCol), ws.Cells(lastRow, tgtCol)).PasteSpecial xlPasteValidation
    Application.CutCopyMode = False
    ws.Columns(tgtCol).ColumnWidth = ws.Columns(srcCol).ColumnWidth
    On Error GoTo 0

    ClearCommentsAndNotes ws.Range(ws.Cells(1, tgtCol), ws.Cells(lastRow, tgtCol))
End Sub

Private Sub ClearCommentsAndNotes(ByVal rng As Range)
    On Error Resume Next
    rng.ClearComments
    rng.ClearNotes
    On Error GoTo 0
End Sub

Private Sub RestoreAppState( _
    ByVal oldCalc      As XlCalculation, _
    ByVal oldScreen    As Boolean, _
    ByVal oldEvents    As Boolean, _
    ByVal oldStatusBar As Variant)
    Application.StatusBar      = oldStatusBar
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents   = oldEvents
    Application.Calculation    = oldCalc
End Sub


' ═══════════════════════════════════════════════════════════════════════════════
'  BYTE-ARRAY FROZEN FORMULA BUILDER  [OPT-5, unchanged from v2]
' ═══════════════════════════════════════════════════════════════════════════════
Private Function BuildFrozenTargetFormula( _
    ByVal formulaText   As String, _
    ByVal hostSheetName As String, _
    ByVal colDelta      As Long) As String

    If g_SafeMode Then
        BuildFrozenTargetFormula = BuildFrozenTargetFormula_Safe(formulaText, hostSheetName, colDelta)
        Exit Function
    End If

    Dim b() As Byte
    b = StrConv(formulaText, vbFromUnicode)

    Dim n         As Long: n = Len(formulaText)
    Dim i         As Long: i = 1
    Dim lastEmit  As Long: lastEmit = 1
    Dim tokenStart As Long, tokenEnd As Long
    Dim addrNorm  As String, rawTok As String, shName As String
    Dim nextPos   As Long, normText As String
    Dim parts()   As String, partCount As Long, partCap As Long
    Dim bCh       As Byte

    Do While i <= n
        bCh = b(i - 1)

        If bCh = BC_DQUOTE Then
            i = SkipDoubleQuotedString(formulaText, i): GoTo BftCont
        End If

        ' [BUG-5] Skip single-quoted sheet qualifiers in byte-array path.
        If bCh = BC_SQUOTE Then
            Dim bftSkipTo As Long
            bftSkipTo = SkipSingleQuotedSheetQualifierByte(b, i, n)
            If bftSkipTo > i Then i = bftSkipTo: GoTo BftCont
        End If

        If bCh = BC_DOLLAR Or BIsLetter(bCh) Then
            If TryParseColumnRangeToken(formulaText, i, tokenStart, tokenEnd, normText) Then
                shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
                AppendStringPart parts, partCount, partCap, Mid$(formulaText, lastEmit, i - lastEmit)
                rawTok = Mid$(formulaText, tokenStart, tokenEnd - tokenStart + 1)
                If StrComp(shName, UCase$(hostSheetName), vbTextCompare) = 0 Then
                    AppendStringPart parts, partCount, partCap, _
                        ShiftColumnRangeTokenHorizontallyPreserveDollar(rawTok, colDelta)
                Else
                    AppendStringPart parts, partCount, partCap, rawTok
                End If
                i = tokenEnd + 1: lastEmit = i: GoTo BftCont
            End If

            If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then
                nextPos = tokenEnd + 1
                Do While nextPos <= n And b(nextPos - 1) = BC_SPACE: nextPos = nextPos + 1: Loop
                If nextPos <= n And b(nextPos - 1) = BC_BANG Then
                    i = i + 1: GoTo BftCont
                End If
                shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
                AppendStringPart parts, partCount, partCap, Mid$(formulaText, lastEmit, i - lastEmit)
                rawTok = Mid$(formulaText, tokenStart, tokenEnd - tokenStart + 1)
                If StrComp(shName, UCase$(hostSheetName), vbTextCompare) = 0 Then
                    AppendStringPart parts, partCount, partCap, _
                        ShiftA1TokenHorizontallyPreserveDollar(rawTok, colDelta)
                Else
                    AppendStringPart parts, partCount, partCap, rawTok
                End If
                i = tokenEnd + 1: lastEmit = i: GoTo BftCont
            End If
        End If

        i = i + 1
BftCont:
    Loop

    AppendStringPart parts, partCount, partCap, Mid$(formulaText, lastEmit)
    If partCount = 0 Then
        BuildFrozenTargetFormula = vbNullString
    Else
        ReDim Preserve parts(1 To partCount)
        BuildFrozenTargetFormula = Join(parts, vbNullString)
    End If
End Function

Private Function BuildFrozenTargetFormula_Safe( _
    ByVal formulaText   As String, _
    ByVal hostSheetName As String, _
    ByVal colDelta      As Long) As String

    Dim n         As Long: n = Len(formulaText)
    Dim i         As Long: i = 1
    Dim lastEmit  As Long: lastEmit = 1
    Dim tokenStart As Long, tokenEnd As Long
    Dim addrNorm  As String, rawTok As String, shName As String
    Dim nextPos   As Long, ch As String, normText As String
    Dim parts()   As String, partCount As Long, partCap As Long

    Do While i <= n
        If TryParseColumnRangeToken(formulaText, i, tokenStart, tokenEnd, normText) Then
            shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
            AppendStringPart parts, partCount, partCap, Mid$(formulaText, lastEmit, i - lastEmit)
            rawTok = Mid$(formulaText, tokenStart, tokenEnd - tokenStart + 1)
            If StrComp(shName, UCase$(hostSheetName), vbTextCompare) = 0 Then
                AppendStringPart parts, partCount, partCap, _
                    ShiftColumnRangeTokenHorizontallyPreserveDollar(rawTok, colDelta)
            Else
                AppendStringPart parts, partCount, partCap, rawTok
            End If
            i = tokenEnd + 1: lastEmit = i: GoTo SfBftCont
        End If

        If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then
            nextPos = tokenEnd + 1
            Do While nextPos <= n And Mid$(formulaText, nextPos, 1) = " ": nextPos = nextPos + 1: Loop
            If nextPos <= n And Mid$(formulaText, nextPos, 1) = "!" Then
                i = i + 1: GoTo SfBftCont
            End If
            shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
            AppendStringPart parts, partCount, partCap, Mid$(formulaText, lastEmit, i - lastEmit)
            rawTok = Mid$(formulaText, tokenStart, tokenEnd - tokenStart + 1)
            If StrComp(shName, UCase$(hostSheetName), vbTextCompare) = 0 Then
                AppendStringPart parts, partCount, partCap, _
                    ShiftA1TokenHorizontallyPreserveDollar(rawTok, colDelta)
            Else
                AppendStringPart parts, partCount, partCap, rawTok
            End If
            i = tokenEnd + 1: lastEmit = i: GoTo SfBftCont
        End If

        ch = Mid$(formulaText, i, 1)
        If ch = """" Then i = SkipDoubleQuotedString(formulaText, i) Else i = i + 1
SfBftCont:
    Loop

    AppendStringPart parts, partCount, partCap, Mid$(formulaText, lastEmit)
    If partCount = 0 Then
        BuildFrozenTargetFormula_Safe = vbNullString
    Else
        ReDim Preserve parts(1 To partCount)
        BuildFrozenTargetFormula_Safe = Join(parts, vbNullString)
    End If
End Function

Private Sub AppendStringPart( _
    ByRef parts()   As String, _
    ByRef partCount As Long, _
    ByRef partCap   As Long, _
    ByVal text      As String)
    If Len(text) = 0 Then Exit Sub
    If partCap = 0 Then
        partCap = 16: ReDim parts(1 To partCap)
    ElseIf partCount >= partCap Then
        partCap = partCap * 2: ReDim Preserve parts(1 To partCap)
    End If
    partCount = partCount + 1
    parts(partCount) = text
End Sub


' ═══════════════════════════════════════════════════════════════════════════════
'  TOKEN SHIFTERS  (unchanged)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function ShiftA1TokenHorizontallyPreserveDollar( _
    ByVal rawToken As String, ByVal colDelta As Long) As String

    Dim p As Long: p = 1
    Dim n As Long: n = Len(rawToken)
    Dim colAbs As Boolean, rowAbs As Boolean
    Dim colLetters As String, rowDigits As String, newCol As Long

    If p <= n And Mid$(rawToken, p, 1) = "$" Then colAbs = True: p = p + 1
    Do While p <= n And IsLetterAZ(Mid$(rawToken, p, 1))
        colLetters = colLetters & Mid$(rawToken, p, 1): p = p + 1
    Loop
    If p <= n And Mid$(rawToken, p, 1) = "$" Then rowAbs = True: p = p + 1
    Do While p <= n And IsDigit09(Mid$(rawToken, p, 1))
        rowDigits = rowDigits & Mid$(rawToken, p, 1): p = p + 1
    Loop

    If Len(colLetters) = 0 Or Len(rowDigits) = 0 Or p <= n Then
        ShiftA1TokenHorizontallyPreserveDollar = rawToken: Exit Function
    End If
    If colAbs Then ShiftA1TokenHorizontallyPreserveDollar = rawToken: Exit Function

    newCol = ColLettersToNumber2(UCase$(colLetters)) + colDelta
    If newCol < 1 Or newCol > 16384 Then
        ShiftA1TokenHorizontallyPreserveDollar = rawToken: Exit Function
    End If

    ShiftA1TokenHorizontallyPreserveDollar = ColumnNumberToLetters(newCol)
    If rowAbs Then ShiftA1TokenHorizontallyPreserveDollar = _
        ShiftA1TokenHorizontallyPreserveDollar & "$"
    ShiftA1TokenHorizontallyPreserveDollar = _
        ShiftA1TokenHorizontallyPreserveDollar & rowDigits
End Function

Private Function ShiftColumnRangeTokenHorizontallyPreserveDollar( _
    ByVal rawToken As String, ByVal colDelta As Long) As String
    Dim parts() As String
    If InStr(1, rawToken, ":", vbBinaryCompare) = 0 Then
        ShiftColumnRangeTokenHorizontallyPreserveDollar = rawToken: Exit Function
    End If
    parts = Split(rawToken, ":")
    If UBound(parts) <> 1 Then
        ShiftColumnRangeTokenHorizontallyPreserveDollar = rawToken: Exit Function
    End If
    ShiftColumnRangeTokenHorizontallyPreserveDollar = _
        ShiftSingleColumnTokenPreserveDollar(parts(0), colDelta) & ":" & _
        ShiftSingleColumnTokenPreserveDollar(parts(1), colDelta)
End Function

Private Function ShiftSingleColumnTokenPreserveDollar( _
    ByVal rawToken As String, ByVal colDelta As Long) As String
    Dim s As String: s = rawToken
    Dim p As Long: p = 1
    Dim absCol As Boolean, colLetters As String, newCol As Long

    If Len(s) = 0 Then ShiftSingleColumnTokenPreserveDollar = s: Exit Function
    If Mid$(s, p, 1) = "$" Then absCol = True: p = p + 1
    Do While p <= Len(s) And IsLetterAZ(Mid$(s, p, 1))
        colLetters = colLetters & Mid$(s, p, 1): p = p + 1
    Loop
    If Len(colLetters) = 0 Or p <= Len(s) Then
        ShiftSingleColumnTokenPreserveDollar = s: Exit Function
    End If
    If absCol Then ShiftSingleColumnTokenPreserveDollar = s: Exit Function

    newCol = ColLettersToNumber2(UCase$(colLetters)) + colDelta
    If newCol < 1 Or newCol > 16384 Then
        ShiftSingleColumnTokenPreserveDollar = s: Exit Function
    End If
    ShiftSingleColumnTokenPreserveDollar = ColumnNumberToLetters(newCol)
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  LEVEL / FREEZE HELPERS  (unchanged)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function ShouldFreezeLevel( _
    ByVal lvl As Long, ByVal freezeMaxLevel As Long) As Boolean
    If freezeMaxLevel <= 0 Then Exit Function
    If lvl = LVL_HARDCODED Then Exit Function
    If lvl < LVL_DIRECT_OTHER Or lvl > LVL_PARENT_OF_DIRECT_OTHER Then Exit Function
    If freezeMaxLevel >= LVL_PARENT_OF_DIRECT_OTHER Then
        ShouldFreezeLevel = (lvl = LVL_DIRECT_OTHER Or lvl = LVL_PARENT_OF_DIRECT_OTHER)
    Else
        ShouldFreezeLevel = (lvl = LVL_DIRECT_OTHER)
    End If
End Function

Private Sub PrepareExistingTargetColumn(ByVal ws As Worksheet, ByVal tgtCol As Long)
    On Error Resume Next
    ws.Columns(tgtCol).Hidden = False
    ws.Columns(tgtCol).Ungroup
    On Error GoTo 0
End Sub

Private Function ParseFreezeMaxLevel(ByVal txt As String) As Long
    Dim s As String: s = Trim$(txt)
    If Len(s) = 0 Then
        ParseFreezeMaxLevel = LVL_PARENT_OF_DIRECT_OTHER
    ElseIf IsNumeric(s) Then
        ParseFreezeMaxLevel = CLng(s)
        If ParseFreezeMaxLevel < LVL_INTERNAL  Then ParseFreezeMaxLevel = LVL_INTERNAL
        If ParseFreezeMaxLevel > LVL_HARDCODED Then ParseFreezeMaxLevel = LVL_HARDCODED
    Else
        ParseFreezeMaxLevel = LVL_PARENT_OF_DIRECT_OTHER
    End If
End Function

Private Function ParseColumnSpec(ByVal colSpec As String) As Long
    Dim s As String: s = Trim$(colSpec)
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Then ParseColumnSpec = CLng(s) Else ParseColumnSpec = ColumnLettersToNumber(UCase$(s))
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  COLUMN LETTER / NUMBER CONVERTERS  [OPT-1 cache, OPT-17 raw variant]
' ═══════════════════════════════════════════════════════════════════════════════
Private Function ColumnLettersToNumber(ByVal letters As String) As Long
    Dim i As Long, v As Long, ch As Integer
    letters = UCase$(Trim$(letters))
    If Len(letters) = 0 Then Exit Function
    For i = 1 To Len(letters)
        ch = Asc(Mid$(letters, i, 1))
        If ch < 65 Or ch > 90 Then ColumnLettersToNumber = 0: Exit Function
        v = v * 26 + (ch - 64)
    Next i
    ColumnLettersToNumber = v
End Function

' [OPT-17] Fast path for callers that already pass clean UPPER-CASE input.
' Avoids redundant UCase$ + Trim$ overhead in hot loops.
Private Function ColumnLettersToNumberRaw(ByVal letters As String) As Long
    Dim i As Long, v As Long, ch As Integer
    If Len(letters) = 0 Then Exit Function
    For i = 1 To Len(letters)
        ch = Asc(Mid$(letters, i, 1))
        If ch < 65 Or ch > 90 Then ColumnLettersToNumberRaw = 0: Exit Function
        v = v * 26 + (ch - 64)
    Next i
    ColumnLettersToNumberRaw = v
End Function

Private Function ColumnNumberToLetters(ByVal colNum As Long) As String
    If colNum < 1 Or colNum > 16384 Then Exit Function
    If Len(g_colCache(colNum)) > 0 Then ColumnNumberToLetters = g_colCache(colNum): Exit Function
    Dim n As Long: n = colNum
    Dim s As String
    Do While n > 0
        s = Chr$(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop
    g_colCache(colNum) = s
    ColumnNumberToLetters = s
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  MISCELLANEOUS HELPERS
' ═══════════════════════════════════════════════════════════════════════════════

' [OPT-16] Fast check using AscW — avoids Left$() string allocation.
Private Function IsFormulaVariant(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then
        Dim s As String: s = CStr(v)
        If Len(s) > 0 Then IsFormulaVariant = (AscW(s) = 61)  ' 61 = "="
    End If
End Function

' [BUG-7] Excel returns a scalar (not an array) when a range is exactly 1 cell.
' This helper wraps it into a 2-D (1 To 1, 1 To 1) array so callers can
' always use arr(i, 1) syntax without subscript errors.
Private Function Ensure2DColumnArray(ByVal v As Variant) As Variant
    If IsArray(v) Then
        Ensure2DColumnArray = v
    Else
        Dim tmp(1 To 1, 1 To 1) As Variant
        tmp(1, 1) = v
        Ensure2DColumnArray = tmp
    End If
End Function

Private Function LastUsedRowAny(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then LastUsedRowAny = 1 Else LastUsedRowAny = f.Row
End Function

Private Function ElapsedSeconds(ByVal startTick As Double) As Double
    If Timer >= startTick Then
        ElapsedSeconds = Timer - startTick
    Else
        ElapsedSeconds = (86400# - startTick) + Timer
    End If
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  [OPT-7 / OPT-11 / OPT-12 / OPT-14 / OPT-15]  FORMULA LEVEL CLASSIFIER
'
'  v4 changes vs. v3:
'   [OPT-14] Pre-size formula cell arrays from urRows * urCols * 0.3,
'            reducing ReDim Preserve calls on formula-dense sheets.
'   [OPT-15] result and dictIndex are module-level pooled dictionaries
'            (g_PoolDictA, g_PoolDictB) — avoids 2× CreateObject per sheet.
'   [OPT-17] ColLettersToNumber2 uses ColumnLettersToNumberRaw for internal use.
' ═══════════════════════════════════════════════════════════════════════════════
Private Function BuildFormulaLevelsForSheet(ByVal ws As Worksheet) As Object

    ' [OPT-15] Pool result dictionary.
    Dim result As Object
    If g_PoolDictA Is Nothing Then
        Set g_PoolDictA = CreateObject("Scripting.Dictionary")
        g_PoolDictA.CompareMode = vbTextCompare
    Else
        g_PoolDictA.RemoveAll
    End If
    Set result = g_PoolDictA

    ' [OPT-11] Read entire UsedRange formula array in one COM call.
    Dim usedRng As Range
    On Error Resume Next
    Set usedRng = ws.UsedRange
    On Error GoTo 0
    If usedRng Is Nothing Then Set BuildFormulaLevelsForSheet = result: Exit Function

    Dim arrF As Variant
    arrF = usedRng.Formula   ' 1 COM call; returns 2-D variant (1-based)

    ' Guard: if UsedRange is exactly 1 cell, arrF is a scalar, not an array.
    If Not IsArray(arrF) Then
        Dim scl As Variant: scl = arrF
        ReDim arrF(1 To 1, 1 To 1): arrF(1, 1) = scl
    End If

    Dim urRows     As Long: urRows     = UBound(arrF, 1)
    Dim urCols     As Long: urCols     = UBound(arrF, 2)
    Dim urFirstRow As Long: urFirstRow = usedRng.Row
    Dim urFirstCol As Long: urFirstCol = usedRng.Column

    ' [minor] Cache sheet name upper-case once.
    Dim wsNameUpper As String: wsNameUpper = UCase$(ws.Name)

    ' ── [OPT-11/OPT-14] Collect formula cells by scanning the array ──────────
    ' [OPT-14] Pre-size with estimate: ~30% of cells are typically formulas.
    Dim cap As Long
    Dim totalCells As Long: totalCells = urRows * urCols
    If totalCells > 800 Then cap = CLng(totalCells * 0.3) Else cap = 256
    If cap < 256 Then cap = 256
    Dim n   As Long: n = 0

    Dim addrArr()     As String    ' A1 address (used as dict key and in output)
    Dim fmlArr()      As String    ' formula text from array
    Dim rowArr()      As Long      ' absolute row number
    Dim colArr()      As Long      ' absolute column number
    ReDim addrArr(1 To cap)
    ReDim fmlArr(1 To cap)
    ReDim rowArr(1 To cap)
    ReDim colArr(1 To cap)

    Dim ri As Long, ci As Long
    Dim absR As Long, absC As Long
    Dim fml  As String, nodeAddr As String

    ' [OPT-11] Pure VBA scan – zero COM calls inside this loop.
    For ri = 1 To urRows
        For ci = 1 To urCols
            ' [BUG-6] Guard against error-value cells (#N/A, #REF!, etc.).
            ' usedRng.Formula returns vbError variants for these; CStr() would crash.
            If VarType(arrF(ri, ci)) = vbError Then GoTo NextFormulaCell
            fml = CStr(arrF(ri, ci))
            If Len(fml) > 0 Then
                ' [OPT-16] Inline AscW check instead of Left$()
                If AscW(fml) = 61 Then
                    n = n + 1
                    If n > cap Then
                        cap = cap * 2
                        ReDim Preserve addrArr(1 To cap)
                        ReDim Preserve fmlArr(1 To cap)
                        ReDim Preserve rowArr(1 To cap)
                        ReDim Preserve colArr(1 To cap)
                    End If
                    absR = urFirstRow + ri - 1
                    absC = urFirstCol + ci - 1
                    ' [OPT-12] Build A1 address without Range.Address COM call.
                    nodeAddr = ColumnNumberToLetters(absC) & CStr(absR)
                    addrArr(n) = nodeAddr
                    fmlArr(n)  = fml
                    rowArr(n)  = absR
                    colArr(n)  = absC
                End If
            End If
NextFormulaCell:
        Next ci
    Next ri

    If n = 0 Then Set BuildFormulaLevelsForSheet = result: Exit Function

    ReDim Preserve addrArr(1 To n)
    ReDim Preserve fmlArr(1 To n)
    ReDim Preserve rowArr(1 To n)
    ReDim Preserve colArr(1 To n)

    ' [OPT-15] Pool dictIndex dictionary.
    Dim dictIndex As Object
    If g_PoolDictB Is Nothing Then
        Set g_PoolDictB = CreateObject("Scripting.Dictionary")
        g_PoolDictB.CompareMode = vbTextCompare
    Else
        g_PoolDictB.RemoveAll
    End If
    Set dictIndex = g_PoolDictB

    Dim i As Long
    For i = 1 To n
        If Not dictIndex.Exists(addrArr(i)) Then dictIndex.Add addrArr(i), i
    Next i

    ' Allocate classification and parent-link arrays.
    Dim parentColls()    As Collection
    Dim hasDirectOther() As Boolean
    Dim levelCap()       As Long
    ReDim parentColls(1 To n)
    ReDim hasDirectOther(1 To n)
    ReDim levelCap(1 To n)

    ' [OPT-6] Initialise shared ref dict once.
    If g_RefDict Is Nothing Then
        Set g_RefDict = CreateObject("Scripting.Dictionary")
        g_RefDict.CompareMode = vbTextCompare
    End If

    ' ── [OPT-11] Second pass: use stored formula text and row/col arrays ───────
    Dim key      As Variant, refKey As String
    Dim barPos   As Long, shName As String, addr As String, childIdx As Long

    For i = 1 To n
        If FormulaHasHardCodedNumbers(fmlArr(i)) Then levelCap(i) = LVL_HARDCODED

        ' [OPT-6] Reuse g_RefDict with RemoveAll.
        g_RefDict.RemoveAll
        FillExplicitRefs fmlArr(i), ws.Name, rowArr(i), colArr(i), g_RefDict

        If g_RefDict.Count > 0 Then
            For Each key In g_RefDict.Keys
                refKey = CStr(key)
                barPos = InStr(1, refKey, "|", vbBinaryCompare)
                shName = Left$(refKey, barPos - 1)
                addr   = Mid$(refKey, barPos + 1)

                If StrComp(shName, wsNameUpper, vbTextCompare) = 0 Then
                    If Left$(addr, 1) <> "#" Then
                        If dictIndex.Exists(addr) Then
                            childIdx = CLng(dictIndex(addr))
                            If parentColls(childIdx) Is Nothing Then
                                Set parentColls(childIdx) = New Collection
                            End If
                            parentColls(childIdx).Add i
                        End If
                    End If
                Else
                    hasDirectOther(i) = True
                End If
            Next key
        End If
    Next i

    ' ── BFS propagation (unchanged) ──────────────────────────────────────────
    Dim q()       As Long
    Dim head      As Long: head = 1
    Dim tail      As Long: tail = 0
    Dim parentIdx As Long, childIdx2 As Long
    Dim part      As Variant
    ReDim q(1 To n)

    For i = 1 To n
        If hasDirectOther(i) Then
            If levelCap(i) <> LVL_HARDCODED Then levelCap(i) = LVL_DIRECT_OTHER
            tail = tail + 1: q(tail) = i
        End If
    Next i

    Do While head <= tail
        childIdx2 = q(head): head = head + 1
        If Not parentColls(childIdx2) Is Nothing Then
            For Each part In parentColls(childIdx2)
                parentIdx = CLng(part)
                If levelCap(parentIdx) = LVL_INTERNAL Then
                    levelCap(parentIdx) = LVL_PARENT_OF_DIRECT_OTHER
                    tail = tail + 1: q(tail) = parentIdx
                End If
            Next part
        End If
    Loop

    For i = 1 To n
        result(addrArr(i)) = levelCap(i)
    Next i

    ' [BUG-4] With the level-map cache removed in v4.1, the caller uses the
    ' returned dictionary locally and discards it when RollOneSheetOneColumn exits.
    ' g_PoolDictA is safe to reuse on the next call — no detachment needed.

    Set BuildFormulaLevelsForSheet = result
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  [OPT-5] FORMULA HAS HARD-CODED NUMBERS  (byte-array + safe fallback)
'  [BUG-2] funcStack enlarged to FUNC_STACK_SIZE (128)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function FormulaHasHardCodedNumbers(ByVal formulaText As String) As Boolean

    If g_SafeMode Then
        FormulaHasHardCodedNumbers = FormulaHasHardCodedNumbers_Safe(formulaText)
        Exit Function
    End If

    Dim b() As Byte
    b = StrConv(formulaText, vbFromUnicode)

    Dim n            As Long: n = Len(formulaText)
    Dim i            As Long: i = 1
    Dim bCh          As Byte
    Dim tokenStart   As Long, tokenEnd As Long
    Dim addrNorm     As String, normText As String
    Dim funcStack(1 To FUNC_STACK_SIZE) As String   ' [BUG-2] was 64
    Dim stackTop     As Long
    Dim pendingFunc  As String
    Dim identEnd     As Long, identText As String

    Do While i <= n
        bCh = b(i - 1)

        If bCh = BC_DQUOTE Then i = SkipDoubleQuotedString(formulaText, i): GoTo HcCont
        ' [BUG-5] Use byte-array version for consistency with other fast paths.
        If bCh = BC_SQUOTE Then
            Dim hcSkipTo As Long
            hcSkipTo = SkipSingleQuotedSheetQualifierByte(b, i, n)
            If hcSkipTo > i Then i = hcSkipTo: GoTo HcCont
        End If

        If bCh = BC_DOLLAR Or BIsLetter(bCh) Then
            If TryParseColumnRangeToken(formulaText, i, tokenStart, tokenEnd, normText) Then
                i = tokenEnd + 1: pendingFunc = vbNullString: GoTo HcCont
            End If
            If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then
                i = tokenEnd + 1: pendingFunc = vbNullString: GoTo HcCont
            End If
            If TryParseBareIdentifier(formulaText, i, identEnd, identText) Then
                pendingFunc = vbNullString
                If NextSignificantChar(formulaText, identEnd + 1) = "(" Then
                    pendingFunc = UCase$(identText)
                End If
                i = identEnd + 1: GoTo HcCont
            End If
        End If

        If bCh = BC_LPAREN Then
            PushFunctionName funcStack, stackTop, pendingFunc
            pendingFunc = vbNullString: i = i + 1: GoTo HcCont
        End If
        If bCh = BC_RPAREN Then
            If stackTop > 0 Then stackTop = stackTop - 1
            pendingFunc = vbNullString: i = i + 1: GoTo HcCont
        End If

        If IsNumericLiteralStart(formulaText, i) Then
            tokenEnd = NumericLiteralTokenEnd(formulaText, i)
            If NumericLiteralMeansHardCodedLevel3(formulaText, i, funcStack, stackTop) Then
                FormulaHasHardCodedNumbers = True: Exit Function
            End If
            If tokenEnd < i Then tokenEnd = i
            i = tokenEnd + 1: pendingFunc = vbNullString: GoTo HcCont
        End If

        If bCh <> BC_SPACE Then pendingFunc = vbNullString
        i = i + 1
HcCont:
    Loop
End Function

Private Function FormulaHasHardCodedNumbers_Safe(ByVal formulaText As String) As Boolean
    Dim n As Long: n = Len(formulaText)
    Dim i As Long: i = 1
    Dim tokenStart As Long, tokenEnd As Long
    Dim addrNorm As String, normText As String, ch As String
    Dim funcStack(1 To FUNC_STACK_SIZE) As String   ' [BUG-2] was 64
    Dim stackTop As Long, pendingFunc As String
    Dim identEnd As Long, identText As String

    Do While i <= n
        ch = Mid$(formulaText, i, 1)
        If ch = """" Then i = SkipDoubleQuotedString(formulaText, i): GoTo SfHcCont
        If ch = "'" Then
            tokenEnd = SkipSingleQuotedSheetQualifier(formulaText, i)
            If tokenEnd > i Then i = tokenEnd: GoTo SfHcCont
        End If
        If TryParseColumnRangeToken(formulaText, i, tokenStart, tokenEnd, normText) Then
            i = tokenEnd + 1: pendingFunc = vbNullString: GoTo SfHcCont
        End If
        If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then
            i = tokenEnd + 1: pendingFunc = vbNullString: GoTo SfHcCont
        End If
        If TryParseBareIdentifier(formulaText, i, identEnd, identText) Then
            pendingFunc = vbNullString
            If NextSignificantChar(formulaText, identEnd + 1) = "(" Then pendingFunc = UCase$(identText)
            i = identEnd + 1: GoTo SfHcCont
        End If
        If ch = "(" Then
            PushFunctionName funcStack, stackTop, pendingFunc
            pendingFunc = vbNullString: i = i + 1: GoTo SfHcCont
        End If
        If ch = ")" Then
            If stackTop > 0 Then stackTop = stackTop - 1
            pendingFunc = vbNullString: i = i + 1: GoTo SfHcCont
        End If
        If IsNumericLiteralStart(formulaText, i) Then
            tokenEnd = NumericLiteralTokenEnd(formulaText, i)
            If NumericLiteralMeansHardCodedLevel3(formulaText, i, funcStack, stackTop) Then
                FormulaHasHardCodedNumbers_Safe = True: Exit Function
            End If
            If tokenEnd < i Then tokenEnd = i
            i = tokenEnd + 1: pendingFunc = vbNullString: GoTo SfHcCont
        End If
        If ch <> " " Then pendingFunc = vbNullString
        i = i + 1
SfHcCont:
    Loop
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  REFERENCE EXTRACTION  [OPT-5 byte-array + OPT-6 dict reuse]
'  [BUG-3] ExtractExplicitRefsInternal now avoids CreateObject per recursion
' ═══════════════════════════════════════════════════════════════════════════════
Private Function ExtractExplicitRefs( _
    ByVal formulaText   As String, _
    ByVal hostSheetName As String, _
    Optional ByVal hostRow As Long = 1, _
    Optional ByVal hostCol As Long = 1) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim ok As Boolean
    FillExplicitRefsInternal formulaText, hostSheetName, hostRow, hostCol, False, ok, d
    Set ExtractExplicitRefs = d
End Function

Private Sub FillExplicitRefs( _
    ByVal formulaText   As String, _
    ByVal hostSheetName As String, _
    ByVal hostRow       As Long, _
    ByVal hostCol       As Long, _
    ByRef d             As Object)
    Dim ok As Boolean
    FillExplicitRefsInternal formulaText, hostSheetName, hostRow, hostCol, False, ok, d
End Sub

Private Sub FillExplicitRefsInternal( _
    ByVal formulaText             As String, _
    ByVal hostSheetName           As String, _
    ByVal hostRow                 As Long, _
    ByVal hostCol                 As Long, _
    ByVal failOnUnresolvedDynamic As Boolean, _
    ByRef parseOk                 As Boolean, _
    ByRef d                       As Object)

    parseOk = True

    If g_SafeMode Then
        FillExplicitRefsInternal_Safe formulaText, hostSheetName, hostRow, hostCol, _
            failOnUnresolvedDynamic, parseOk, d
        Exit Sub
    End If

    Dim b() As Byte
    b = StrConv(formulaText, vbFromUnicode)

    Dim n          As Long: n = Len(formulaText)
    Dim i          As Long: i = 1
    Dim bCh        As Byte
    Dim tokenStart As Long, tokenEnd As Long
    Dim addrNorm   As String, shName As String, key As String
    Dim nextPos    As Long, consumeRes As Long

    Do While i <= n
        bCh = b(i - 1)

        consumeRes = TryConsumeIndirectOrIgnore( _
            formulaText, i, hostSheetName, hostRow, hostCol, d, tokenEnd)
        If consumeRes <> 0 Then
            If consumeRes = 2 And failOnUnresolvedDynamic Then parseOk = False: Exit Do
            i = tokenEnd + 1: GoTo FriCont
        End If

        consumeRes = TryConsumeOffsetOrIgnore(formulaText, i, hostSheetName, d, tokenEnd)
        If consumeRes <> 0 Then
            If consumeRes = 2 And failOnUnresolvedDynamic Then parseOk = False: Exit Do
            i = tokenEnd + 1: GoTo FriCont
        End If

        If TryConsumeFullColumnOrRowRef(formulaText, i, hostSheetName, d, tokenEnd) Then
            i = tokenEnd + 1: GoTo FriCont
        End If

        If bCh = BC_DQUOTE Then i = SkipDoubleQuotedString(formulaText, i): GoTo FriCont

        ' [BUG-5] Skip single-quoted sheet qualifiers in byte-array path.
        If bCh = BC_SQUOTE Then
            Dim friSkipTo As Long
            friSkipTo = SkipSingleQuotedSheetQualifierByte(b, i, n)
            If friSkipTo > i Then i = friSkipTo: GoTo FriCont
        End If

        If bCh = BC_DOLLAR Or BIsLetter(bCh) Then
            If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then
                nextPos = tokenEnd + 1
                Do While nextPos <= n And b(nextPos - 1) = BC_SPACE: nextPos = nextPos + 1: Loop
                If nextPos <= n And b(nextPos - 1) = BC_BANG Then
                    i = tokenEnd + 1
                Else
                    shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
                    key = UCase$(shName) & "|" & addrNorm
                    If Not d.Exists(key) Then d.Add key, True
                    i = tokenEnd + 1
                End If
                GoTo FriCont
            End If
        End If

        i = i + 1
FriCont:
    Loop
End Sub

Private Sub FillExplicitRefsInternal_Safe( _
    ByVal formulaText             As String, _
    ByVal hostSheetName           As String, _
    ByVal hostRow                 As Long, _
    ByVal hostCol                 As Long, _
    ByVal failOnUnresolvedDynamic As Boolean, _
    ByRef parseOk                 As Boolean, _
    ByRef d                       As Object)

    Dim n As Long: n = Len(formulaText)
    Dim i As Long: i = 1
    Dim tokenStart As Long, tokenEnd As Long
    Dim addrNorm As String, shName As String, key As String, ch As String
    Dim nextPos As Long, consumeRes As Long

    Do While i <= n
        consumeRes = TryConsumeIndirectOrIgnore( _
            formulaText, i, hostSheetName, hostRow, hostCol, d, tokenEnd)
        If consumeRes <> 0 Then
            If consumeRes = 2 And failOnUnresolvedDynamic Then parseOk = False: Exit Do
            i = tokenEnd + 1: GoTo SfFriCont
        End If
        consumeRes = TryConsumeOffsetOrIgnore(formulaText, i, hostSheetName, d, tokenEnd)
        If consumeRes <> 0 Then
            If consumeRes = 2 And failOnUnresolvedDynamic Then parseOk = False: Exit Do
            i = tokenEnd + 1: GoTo SfFriCont
        End If
        If TryConsumeFullColumnOrRowRef(formulaText, i, hostSheetName, d, tokenEnd) Then
            i = tokenEnd + 1: GoTo SfFriCont
        End If
        ch = Mid$(formulaText, i, 1)
        If ch = """" Then i = SkipDoubleQuotedString(formulaText, i): GoTo SfFriCont
        If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then
            nextPos = tokenEnd + 1
            Do While nextPos <= n And Mid$(formulaText, nextPos, 1) = " ": nextPos = nextPos + 1: Loop
            If nextPos <= n And Mid$(formulaText, nextPos, 1) = "!" Then
                i = tokenEnd + 1
            Else
                shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
                key = UCase$(shName) & "|" & addrNorm
                If Not d.Exists(key) Then d.Add key, True
                i = tokenEnd + 1
            End If
            GoTo SfFriCont
        End If
        i = i + 1
SfFriCont:
    Loop
End Sub

' [BUG-3] Recursive variant that reuses the caller's dictionary instead of
' creating a new Scripting.Dictionary per INDIRECT() recursion.
Private Function ExtractExplicitRefsInternal( _
    ByVal formulaText             As String, _
    ByVal hostSheetName           As String, _
    ByVal hostRow                 As Long, _
    ByVal hostCol                 As Long, _
    ByVal failOnUnresolvedDynamic As Boolean, _
    ByRef parseOk                 As Boolean) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    FillExplicitRefsInternal formulaText, hostSheetName, hostRow, hostCol, _
        failOnUnresolvedDynamic, parseOk, d
    Set ExtractExplicitRefsInternal = d
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  FUNCTION / SPECIAL-REF CONSUMERS  (unchanged)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function TryConsumeIndirectOrIgnore( _
    ByVal s As String, ByVal pos As Long, ByVal hostSheetName As String, _
    ByVal hostRow As Long, ByVal hostCol As Long, _
    ByRef outDict As Object, ByRef endPos As Long) As Long

    Dim openParenPos As Long, closePos As Long
    Dim args() As String, lit As String, litOk As Boolean
    Dim subRefs As Object, k As Variant, isA1Mode As Boolean
    Dim shName As String, r1 As Long, c1 As Long, r2 As Long, c2 As Long

    TryConsumeIndirectOrIgnore = 0
    endPos = pos
    If Not MatchFunctionNameAt(s, pos, "INDIRECT", openParenPos) Then Exit Function
    If Not TryReadFunctionArgs(s, openParenPos, args, closePos) Then
        endPos = openParenPos: TryConsumeIndirectOrIgnore = 2: Exit Function
    End If
    endPos = closePos
    If UBound(args) < 1 Then TryConsumeIndirectOrIgnore = 2: Exit Function
    If Not TryGetStandaloneQuotedLiteral(args(1), lit) Then
        TryConsumeIndirectOrIgnore = 2: Exit Function
    End If
    isA1Mode = True
    If UBound(args) >= 2 Then
        If Not TryParseIndirectA1Mode(args(2), isA1Mode) Then
            TryConsumeIndirectOrIgnore = 2: Exit Function
        End If
    End If
    If isA1Mode Then
        Set subRefs = ExtractExplicitRefsInternal(lit, hostSheetName, hostRow, hostCol, True, litOk)
        If (Not litOk) Or subRefs.Count = 0 Then TryConsumeIndirectOrIgnore = 2: Exit Function
        For Each k In subRefs.Keys
            If Not outDict.Exists(CStr(k)) Then outDict.Add CStr(k), True
        Next k
        TryConsumeIndirectOrIgnore = 1: Exit Function
    End If
    If Not TryParseR1C1RefOrRangeText(lit, hostSheetName, hostRow, hostCol, shName, r1, c1, r2, c2) Then
        TryConsumeIndirectOrIgnore = 2: Exit Function
    End If
    AddRefKey outDict, shName, RowColToA1(r1, c1)
    AddRefKey outDict, shName, RowColToA1(r2, c2)
    TryConsumeIndirectOrIgnore = 1
End Function

Private Function TryConsumeOffsetOrIgnore( _
    ByVal s As String, ByVal pos As Long, ByVal hostSheetName As String, _
    ByRef outDict As Object, ByRef endPos As Long) As Long

    Dim openParenPos As Long, closePos As Long
    Dim args() As String, argCount As Long
    Dim shName As String, r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Dim rowsOff As Long, colsOff As Long, h As Long, w As Long
    Dim topRow As Long, botRow As Long, leftCol As Long, rightCol As Long
    Dim startRow As Long, startCol As Long, endRow As Long, endCol As Long

    TryConsumeOffsetOrIgnore = 0
    endPos = pos
    If Not MatchFunctionNameAt(s, pos, "OFFSET", openParenPos) Then Exit Function
    If Not TryReadFunctionArgs(s, openParenPos, args, closePos) Then
        endPos = openParenPos: TryConsumeOffsetOrIgnore = 2: Exit Function
    End If
    endPos = closePos: argCount = UBound(args)
    If argCount < 3 Or argCount > 5 Then TryConsumeOffsetOrIgnore = 2: Exit Function
    If Not TryParseSimpleRefOrRangeArg(args(1), hostSheetName, shName, r1, c1, r2, c2) Then
        TryConsumeOffsetOrIgnore = 2: Exit Function
    End If
    If Not TryParseLongLiteral(args(2), rowsOff, False, True) Then
        TryConsumeOffsetOrIgnore = 2: Exit Function
    End If
    If Not TryParseLongLiteral(args(3), colsOff, False, True) Then
        TryConsumeOffsetOrIgnore = 2: Exit Function
    End If
    If r1 <= r2 Then topRow = r1: botRow = r2 Else topRow = r2: botRow = r1
    If c1 <= c2 Then leftCol = c1: rightCol = c2 Else leftCol = c2: rightCol = c1
    startRow = topRow + rowsOff: startCol = leftCol + colsOff
    If argCount = 3 Then
        h = botRow - topRow + 1: w = rightCol - leftCol + 1
    ElseIf argCount = 4 Then
        If Not TryParseLongLiteral(args(4), h, True, False) Then
            TryConsumeOffsetOrIgnore = 2: Exit Function
        End If
        w = rightCol - leftCol + 1
    Else
        If Not TryParseLongLiteral(args(4), h, True, False) Then
            TryConsumeOffsetOrIgnore = 2: Exit Function
        End If
        If Not TryParseLongLiteral(args(5), w, True, False) Then
            TryConsumeOffsetOrIgnore = 2: Exit Function
        End If
    End If
    endRow = startRow + h - 1: endCol = startCol + w - 1
    If startRow < 1 Or startRow > 1048576 Then TryConsumeOffsetOrIgnore = 2: Exit Function
    If startCol < 1 Or startCol > 16384   Then TryConsumeOffsetOrIgnore = 2: Exit Function
    If endRow < 1 Or endRow > 1048576     Then TryConsumeOffsetOrIgnore = 2: Exit Function
    If endCol < 1 Or endCol > 16384       Then TryConsumeOffsetOrIgnore = 2: Exit Function
    AddRefKey outDict, shName, RowColToA1(startRow, startCol)
    AddRefKey outDict, shName, RowColToA1(endRow, endCol)
    TryConsumeOffsetOrIgnore = 1
End Function

Private Function TryConsumeFullColumnOrRowRef( _
    ByVal s As String, ByVal pos As Long, ByVal hostSheetName As String, _
    ByRef outDict As Object, ByRef endPos As Long) As Boolean

    Dim tokenStart As Long, tokenEnd As Long, normText As String, shName As String
    TryConsumeFullColumnOrRowRef = False
    endPos = pos
    If TryParseColumnRangeToken(s, pos, tokenStart, tokenEnd, normText) Then
        shName = GetQualifierSheetName(s, tokenStart, hostSheetName)
        AddRefKey outDict, shName, "#COL#" & normText
        endPos = tokenEnd: TryConsumeFullColumnOrRowRef = True: Exit Function
    End If
    If TryParseRowRangeToken(s, pos, tokenStart, tokenEnd, normText) Then
        shName = GetQualifierSheetName(s, tokenStart, hostSheetName)
        AddRefKey outDict, shName, "#ROW#" & normText
        endPos = tokenEnd: TryConsumeFullColumnOrRowRef = True: Exit Function
    End If
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  TOKEN PARSERS  (unchanged)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function TryParseColumnRangeToken( _
    ByVal s As String, ByVal pos As Long, _
    ByRef tokenStart As Long, ByRef tokenEnd As Long, ByRef normText As String) As Boolean

    Dim n As Long: n = Len(s)
    Dim p As Long: p = pos
    Dim ch As String, start1 As Long, start2 As Long, col1 As String, col2 As String

    TryParseColumnRangeToken = False: normText = vbNullString
    If pos < 1 Or pos > n Then Exit Function
    If pos > 1 Then ch = Mid$(s, pos - 1, 1): If IsAlphaNumUnderscoreDot(ch) Or ch = "[" Then Exit Function

    If Mid$(s, p, 1) = "$" Then p = p + 1
    start1 = p
    Do While p <= n And IsLetterAZ(Mid$(s, p, 1)): p = p + 1: Loop
    If p = start1 Or (p - start1) > 3 Then Exit Function
    col1 = UCase$(Mid$(s, start1, p - start1))
    If p > n Or Mid$(s, p, 1) <> ":" Then Exit Function
    p = p + 1
    If p <= n And Mid$(s, p, 1) = "$" Then p = p + 1
    start2 = p
    Do While p <= n And IsLetterAZ(Mid$(s, p, 1)): p = p + 1: Loop
    If p = start2 Or (p - start2) > 3 Then Exit Function
    col2 = UCase$(Mid$(s, start2, p - start2))
    If ColLettersToNumber2(col1) < 1 Or ColLettersToNumber2(col1) > 16384 Then Exit Function
    If ColLettersToNumber2(col2) < 1 Or ColLettersToNumber2(col2) > 16384 Then Exit Function
    If p <= n Then ch = Mid$(s, p, 1): If IsAlphaNumUnderscoreDot(ch) Then Exit Function

    tokenStart = pos: tokenEnd = p - 1
    normText = col1 & ":" & col2: TryParseColumnRangeToken = True
End Function

Private Function TryParseRowRangeToken( _
    ByVal s As String, ByVal pos As Long, _
    ByRef tokenStart As Long, ByRef tokenEnd As Long, ByRef normText As String) As Boolean

    Dim n As Long: n = Len(s)
    Dim p As Long: p = pos
    Dim ch As String, start1 As Long, start2 As Long
    Dim row1Text As String, row2Text As String, row1 As Double, row2 As Double

    TryParseRowRangeToken = False: normText = vbNullString
    If pos < 1 Or pos > n Then Exit Function
    If pos > 1 Then ch = Mid$(s, pos - 1, 1): If IsAlphaNumUnderscoreDot(ch) Or ch = "[" Then Exit Function

    If Mid$(s, p, 1) = "$" Then p = p + 1
    start1 = p
    Do While p <= n And IsDigit09(Mid$(s, p, 1)): p = p + 1: Loop
    If p = start1 Then Exit Function
    row1Text = Mid$(s, start1, p - start1)
    If p > n Or Mid$(s, p, 1) <> ":" Then Exit Function
    p = p + 1
    If p <= n And Mid$(s, p, 1) = "$" Then p = p + 1
    start2 = p
    Do While p <= n And IsDigit09(Mid$(s, p, 1)): p = p + 1: Loop
    If p = start2 Then Exit Function
    row2Text = Mid$(s, start2, p - start2)
    row1 = CDbl(row1Text): row2 = CDbl(row2Text)
    If row1 < 1 Or row1 > 1048576 Then Exit Function
    If row2 < 1 Or row2 > 1048576 Then Exit Function
    If p <= n Then ch = Mid$(s, p, 1): If IsAlphaNumUnderscoreDot(ch) Then Exit Function

    tokenStart = pos: tokenEnd = p - 1
    normText = CLng(row1) & ":" & CLng(row2): TryParseRowRangeToken = True
End Function

Private Function TryParseIndirectA1Mode(ByVal argText As String, ByRef isA1Mode As Boolean) As Boolean
    Dim s As String: s = UCase$(Replace(Trim$(argText), " ", ""))
    Select Case s
        Case "TRUE", "1":  isA1Mode = True:  TryParseIndirectA1Mode = True
        Case "FALSE", "0": isA1Mode = False: TryParseIndirectA1Mode = True
    End Select
End Function

Private Function TryParseR1C1RefOrRangeText( _
    ByVal expr As String, ByVal defaultSheetName As String, _
    ByVal hostRow As Long, ByVal hostCol As Long, _
    ByRef outSheet As String, _
    ByRef r1 As Long, ByRef c1 As Long, ByRef r2 As Long, ByRef c2 As Long) As Boolean

    Dim s As String: s = Trim$(expr)
    If Len(s) = 0 Then Exit Function
    Dim colonPos As Long: colonPos = FindTopLevelColon(s)
    Dim sh1 As String, sh2 As String, hasSh1 As Boolean, hasSh2 As Boolean

    If colonPos = 0 Then
        If Not TryParseQualifiedR1C1Ref(s, defaultSheetName, hostRow, hostCol, sh1, r1, c1, hasSh1) Then Exit Function
        outSheet = sh1: r2 = r1: c2 = c1: TryParseR1C1RefOrRangeText = True: Exit Function
    End If
    If Not TryParseQualifiedR1C1Ref(Trim$(Left$(s, colonPos - 1)), defaultSheetName, hostRow, hostCol, sh1, r1, c1, hasSh1) Then Exit Function
    If Not TryParseQualifiedR1C1Ref(Trim$(Mid$(s, colonPos + 1)), IIf(hasSh1, sh1, defaultSheetName), hostRow, hostCol, sh2, r2, c2, hasSh2) Then Exit Function
    If StrComp(sh1, sh2, vbTextCompare) <> 0 Then Exit Function
    outSheet = sh1: TryParseR1C1RefOrRangeText = True
End Function

Private Function TryParseQualifiedR1C1Ref( _
    ByVal s As String, ByVal defaultSheetName As String, _
    ByVal hostRow As Long, ByVal hostCol As Long, _
    ByRef outSheet As String, ByRef outRow As Long, ByRef outCol As Long, _
    ByRef hasExplicitSheet As Boolean) As Boolean

    Dim t As String: t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    hasExplicitSheet = False: outSheet = UCase$(defaultSheetName)
    Dim bangPos As Long: bangPos = FindLastBangOutsideQuotes(t)
    If bangPos > 0 Then
        outSheet = UCase$(CleanSheetQualifier(Trim$(Left$(t, bangPos - 1))))
        t = Trim$(Mid$(t, bangPos + 1)): hasExplicitSheet = True
    End If
    If Not TryParseR1C1Single(t, hostRow, hostCol, outRow, outCol) Then Exit Function
    TryParseQualifiedR1C1Ref = True
End Function

Private Function TryParseR1C1Single( _
    ByVal s As String, ByVal hostRow As Long, ByVal hostCol As Long, _
    ByRef outRow As Long, ByRef outCol As Long) As Boolean

    Dim t As String: t = UCase$(Replace(Trim$(s), " ", ""))
    Dim n As Long: n = Len(t)
    If n = 0 Then Exit Function

    Dim p As Long: p = 1
    Dim startPos As Long, ch As String, v As Long

    If Mid$(t, p, 1) <> "R" Then Exit Function
    p = p + 1

    If p <= n And Mid$(t, p, 1) = "[" Then
        p = p + 1: startPos = p
        If p <= n Then ch = Mid$(t, p, 1): If ch = "+" Or ch = "-" Then p = p + 1
        Do While p <= n And IsDigit09(Mid$(t, p, 1)): p = p + 1: Loop
        If p = startPos Then Exit Function
        If p = startPos + 1 Then ch = Mid$(t, startPos, 1): If ch = "+" Or ch = "-" Then Exit Function
        If p > n Or Mid$(t, p, 1) <> "]" Then Exit Function
        outRow = hostRow + CLng(Mid$(t, startPos, p - startPos)): p = p + 1
    Else
        startPos = p
        Do While p <= n And IsDigit09(Mid$(t, p, 1)): p = p + 1: Loop
        If p > startPos Then outRow = CLng(Mid$(t, startPos, p - startPos)) Else outRow = hostRow
    End If

    If p > n Or Mid$(t, p, 1) <> "C" Then Exit Function
    p = p + 1

    If p <= n And Mid$(t, p, 1) = "[" Then
        p = p + 1: startPos = p
        If p <= n Then ch = Mid$(t, p, 1): If ch = "+" Or ch = "-" Then p = p + 1
        Do While p <= n And IsDigit09(Mid$(t, p, 1)): p = p + 1: Loop
        If p = startPos Then Exit Function
        If p = startPos + 1 Then ch = Mid$(t, startPos, 1): If ch = "+" Or ch = "-" Then Exit Function
        If p > n Or Mid$(t, p, 1) <> "]" Then Exit Function
        outCol = hostCol + CLng(Mid$(t, startPos, p - startPos)): p = p + 1
    Else
        startPos = p
        Do While p <= n And IsDigit09(Mid$(t, p, 1)): p = p + 1: Loop
        If p > startPos Then outCol = CLng(Mid$(t, startPos, p - startPos)) Else outCol = hostCol
    End If

    If p <= n Then Exit Function
    If outRow < 1 Or outRow > 1048576 Then Exit Function
    If outCol < 1 Or outCol > 16384  Then Exit Function
    TryParseR1C1Single = True
End Function

Private Function MatchFunctionNameAt( _
    ByVal s As String, ByVal pos As Long, ByVal funcName As String, _
    ByRef openParenPos As Long) As Boolean

    Dim n As Long: n = Len(s)
    Dim L As Long: L = Len(funcName)
    Dim p As Long, ch As String

    MatchFunctionNameAt = False: openParenPos = 0
    If pos < 1 Or pos + L - 1 > n Then Exit Function
    If UCase$(Mid$(s, pos, 1)) <> Left$(UCase$(funcName), 1) Then Exit Function
    If UCase$(Mid$(s, pos, L)) <> UCase$(funcName) Then Exit Function
    If pos > 1 Then ch = Mid$(s, pos - 1, 1): If IsAlphaNumUnderscoreDot(ch) Or ch = "[" Then Exit Function
    If pos + L <= n Then ch = Mid$(s, pos + L, 1): If IsAlphaNumUnderscoreDot(ch) Then Exit Function
    p = pos + L
    Do While p <= n And Mid$(s, p, 1) = " ": p = p + 1: Loop
    If p > n Or Mid$(s, p, 1) <> "(" Then Exit Function
    openParenPos = p: MatchFunctionNameAt = True
End Function

Private Function TryReadFunctionArgs( _
    ByVal s As String, ByVal openParenPos As Long, _
    ByRef args() As String, ByRef closeParenPos As Long) As Boolean

    If openParenPos < 1 Or openParenPos > Len(s) Then Exit Function
    If Mid$(s, openParenPos, 1) <> "(" Then Exit Function

    Dim n As Long: n = Len(s)
    Dim p As Long: p = openParenPos + 1
    Dim depth As Long: depth = 1
    Dim argStart As Long: argStart = p
    Dim argCount As Long, capacity As Long: capacity = 8
    Dim ch As String, finalArg As String
    ReDim args(1 To capacity)

    Do While p <= n
        ch = Mid$(s, p, 1)
        If ch = """" Then
            p = SkipDoubleQuotedString(s, p)
        ElseIf ch = "(" Then
            depth = depth + 1: p = p + 1
        ElseIf ch = ")" Then
            depth = depth - 1
            If depth = 0 Then
                finalArg = Trim$(Mid$(s, argStart, p - argStart))
                If argCount = 0 And Len(finalArg) = 0 Then
                    ReDim args(0 To 0): args(0) = vbNullString
                Else
                    argCount = argCount + 1
                    If argCount > capacity Then capacity = capacity * 2: ReDim Preserve args(1 To capacity)
                    args(argCount) = finalArg
                    ReDim Preserve args(1 To argCount)
                End If
                closeParenPos = p: TryReadFunctionArgs = True: Exit Function
            Else
                p = p + 1
            End If
        ElseIf ch = "," And depth = 1 Then
            argCount = argCount + 1
            If argCount > capacity Then capacity = capacity * 2: ReDim Preserve args(1 To capacity)
            args(argCount) = Trim$(Mid$(s, argStart, p - argStart))
            argStart = p + 1: p = p + 1
        Else
            p = p + 1
        End If
    Loop
    ReDim args(0 To 0): args(0) = vbNullString
End Function

Private Function TryGetStandaloneQuotedLiteral(ByVal expr As String, ByRef outText As String) As Boolean
    Dim s As String: s = Trim$(expr)
    Dim qEnd As Long
    TryGetStandaloneQuotedLiteral = False: outText = vbNullString
    If Len(s) < 2 Or Left$(s, 1) <> """" Then Exit Function
    outText = ReadDoubleQuotedLiteral(s, 1, qEnd)
    If qEnd = 0 Or qEnd <> Len(s) Then Exit Function
    TryGetStandaloneQuotedLiteral = True
End Function

Private Function ReadDoubleQuotedLiteral( _
    ByVal s As String, ByVal startQuotePos As Long, ByRef endQuotePos As Long) As String
    endQuotePos = 0
    If startQuotePos < 1 Or startQuotePos > Len(s) Then Exit Function
    If Mid$(s, startQuotePos, 1) <> """" Then Exit Function
    Dim n As Long: n = Len(s)
    Dim p As Long: p = startQuotePos + 1
    Do While p <= n
        If Mid$(s, p, 1) = """" Then
            If p < n And Mid$(s, p + 1, 1) = """" Then p = p + 2 Else endQuotePos = p: ReadDoubleQuotedLiteral = Replace(Mid$(s, startQuotePos + 1, p - startQuotePos - 1), """""", """"): Exit Function
        Else
            p = p + 1
        End If
    Loop
End Function

Private Function TryParseSimpleRefOrRangeArg( _
    ByVal argText As String, ByVal hostSheetName As String, _
    ByRef outSheet As String, ByRef r1 As Long, ByRef c1 As Long, _
    ByRef r2 As Long, ByRef c2 As Long) As Boolean

    Dim s As String: s = Trim$(argText)
    If Len(s) = 0 Then Exit Function
    Dim colonPos As Long: colonPos = FindTopLevelColon(s)
    Dim addr1 As String, addr2 As String, sh1 As String, sh2 As String
    Dim hasSh1 As Boolean, hasSh2 As Boolean

    If colonPos = 0 Then
        If Not TryParseQualifiedCellRef(s, hostSheetName, sh1, addr1, hasSh1) Then Exit Function
        outSheet = sh1: A1ToRowCol addr1, r1, c1: r2 = r1: c2 = c1
        TryParseSimpleRefOrRangeArg = True: Exit Function
    End If
    If Not TryParseQualifiedCellRef(Trim$(Left$(s, colonPos - 1)), hostSheetName, sh1, addr1, hasSh1) Then Exit Function
    If Not TryParseQualifiedCellRef(Trim$(Mid$(s, colonPos + 1)), IIf(hasSh1, sh1, hostSheetName), sh2, addr2, hasSh2) Then Exit Function
    If StrComp(sh1, sh2, vbTextCompare) <> 0 Then Exit Function
    outSheet = sh1: A1ToRowCol addr1, r1, c1: A1ToRowCol addr2, r2, c2
    TryParseSimpleRefOrRangeArg = True
End Function

Private Function TryParseQualifiedCellRef( _
    ByVal s As String, ByVal defaultSheetName As String, _
    ByRef outSheet As String, ByRef outAddr As String, _
    ByRef hasExplicitSheet As Boolean) As Boolean

    Dim t As String: t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    hasExplicitSheet = False: outSheet = UCase$(defaultSheetName): outAddr = vbNullString
    Dim bangPos As Long: bangPos = FindLastBangOutsideQuotes(t)
    If bangPos > 0 Then
        outSheet = UCase$(CleanSheetQualifier(Trim$(Left$(t, bangPos - 1))))
        t = Trim$(Mid$(t, bangPos + 1)): hasExplicitSheet = True
    End If
    outAddr = NormalizeA1(t)
    If LenB(outAddr) = 0 Then Exit Function
    TryParseQualifiedCellRef = True
End Function

Private Function TryParseCellToken( _
    ByVal s As String, ByVal pos As Long, _
    ByRef tokenStart As Long, ByRef tokenEnd As Long, _
    ByRef addrNorm As String) As Boolean

    Dim n As Long: n = Len(s)
    Dim p As Long: p = pos
    Dim ch As String, colStart As Long, rowStart As Long

    TryParseCellToken = False: addrNorm = vbNullString
    If pos < 1 Or pos > n Then Exit Function
    If pos > 1 Then ch = Mid$(s, pos - 1, 1): If IsAlphaNumUnderscoreDot(ch) Or ch = "[" Then Exit Function

    If Mid$(s, p, 1) = "$" Then p = p + 1
    If p > n Then Exit Function
    colStart = p
    Do While p <= n And IsLetterAZ(Mid$(s, p, 1)): p = p + 1: Loop
    If p = colStart Or (p - colStart) > 3 Then Exit Function
    If p <= n Then If Mid$(s, p, 1) = "$" Then p = p + 1
    rowStart = p
    Do While p <= n And IsDigit09(Mid$(s, p, 1)): p = p + 1: Loop
    If p = rowStart Then Exit Function
    If p <= n Then ch = Mid$(s, p, 1): If IsAlphaNumUnderscoreDot(ch) Then Exit Function

    addrNorm = NormalizeA1(Mid$(s, pos, p - pos))
    If LenB(addrNorm) = 0 Then Exit Function
    tokenStart = pos: tokenEnd = p - 1: TryParseCellToken = True
End Function

Private Function GetQualifierSheetName( _
    ByVal formulaText As String, ByVal tokenStart As Long, _
    ByVal hostSheetName As String) As String

    Dim exPos As Long: exPos = tokenStart - 1
    Do While exPos >= 1 And Mid$(formulaText, exPos, 1) = " ": exPos = exPos - 1: Loop
    If exPos < 1 Or Mid$(formulaText, exPos, 1) <> "!" Then
        GetQualifierSheetName = UCase$(hostSheetName): Exit Function
    End If

    Dim endPos As Long: endPos = exPos - 1
    If endPos < 1 Then GetQualifierSheetName = UCase$(hostSheetName): Exit Function

    Dim startPos As Long, raw As String, ch As String
    If Mid$(formulaText, endPos, 1) = "'" Then
        startPos = endPos - 1
        Do While startPos >= 1
            If Mid$(formulaText, startPos, 1) = "'" Then
                If startPos > 1 And Mid$(formulaText, startPos - 1, 1) = "'" Then
                    startPos = startPos - 2
                Else
                    Exit Do
                End If
            Else
                startPos = startPos - 1
            End If
        Loop
        If startPos < 1 Then GetQualifierSheetName = UCase$(hostSheetName): Exit Function
        raw = Mid$(formulaText, startPos, endPos - startPos + 1)
    Else
        startPos = endPos
        Do While startPos >= 1
            ch = Mid$(formulaText, startPos, 1)
            If IsSheetQualifierChar(ch) Then startPos = startPos - 1 Else Exit Do
        Loop
        raw = Mid$(formulaText, startPos + 1, endPos - startPos)
    End If

    raw = CleanSheetQualifier(raw)
    If LenB(raw) = 0 Then raw = hostSheetName
    GetQualifierSheetName = UCase$(raw)
End Function

Private Function CleanSheetQualifier(ByVal raw As String) As String
    raw = Trim$(raw)
    If Len(raw) = 0 Then CleanSheetQualifier = vbNullString: Exit Function
    If Left$(raw, 1) = "'" And Right$(raw, 1) = "'" Then
        raw = Mid$(raw, 2, Len(raw) - 2): raw = Replace(raw, "''", "'")
    End If
    If InStrRev(raw, "]") > 0 Then raw = Mid$(raw, InStrRev(raw, "]") + 1)
    CleanSheetQualifier = raw
End Function

Private Function NormalizeA1(ByVal raw As String) As String
    Dim s As String: s = UCase$(Replace(raw, "$", ""))
    If Len(s) = 0 Then Exit Function

    Dim i As Long: i = 1
    Dim colPart As String, rowPart As String
    Do While i <= Len(s) And IsLetterAZ(Mid$(s, i, 1)): colPart = colPart & Mid$(s, i, 1): i = i + 1: Loop
    Do While i <= Len(s) And IsDigit09(Mid$(s, i, 1)): rowPart = rowPart & Mid$(s, i, 1): i = i + 1: Loop

    If i <= Len(s) Or Len(colPart) = 0 Or Len(rowPart) = 0 Or Len(colPart) > 3 Then Exit Function
    Dim colNum As Long: colNum = ColumnLettersToNumberRaw(colPart)  ' [OPT-17]
    If colNum < 1 Or colNum > 16384 Then Exit Function
    Dim rowNum As Double: rowNum = CDbl(rowPart)
    If rowNum < 1 Or rowNum > 1048576 Then Exit Function
    NormalizeA1 = colPart & CLng(rowNum)
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  NUMERIC LITERAL HELPERS  (unchanged except BUG-2 stack size)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function NumericLiteralMeansHardCodedLevel3( _
    ByVal formulaText As String, ByVal tokenStart As Long, _
    ByRef funcStack() As String, ByVal stackTop As Long) As Boolean

    Dim prevPos      As Long: prevPos = PrevNonSpacePos(formulaText, tokenStart - 1)
    Dim prevCh       As String: If prevPos > 0 Then prevCh = Mid$(formulaText, prevPos, 1)
    Dim tokenFirstCh As String: tokenFirstCh = Mid$(formulaText, tokenStart, 1)
    Dim activeFunc   As String: activeFunc = CurrentFunctionName(funcStack, stackTop)

    If prevPos = 1 And Mid$(formulaText, 1, 1) = "=" Then
        NumericLiteralMeansHardCodedLevel3 = True: Exit Function
    End If

    If tokenFirstCh = "+" Or tokenFirstCh = "-" Then
        Select Case prevCh
            Case ",", ";":      NumericLiteralMeansHardCodedLevel3 = (activeFunc = "SUM")
            Case "(":           NumericLiteralMeansHardCodedLevel3 = (activeFunc = "SUM")
            Case "<", ">", "=": NumericLiteralMeansHardCodedLevel3 = False
            Case Else:          NumericLiteralMeansHardCodedLevel3 = True
        End Select
        Exit Function
    End If

    Select Case prevCh
        Case "+", "-":          NumericLiteralMeansHardCodedLevel3 = True
        Case ",", ";", "(":     NumericLiteralMeansHardCodedLevel3 = (activeFunc = "SUM")
        Case Else:              NumericLiteralMeansHardCodedLevel3 = False
    End Select
End Function

' [BUG-2] Now warns on overflow instead of silently dropping.
Private Sub PushFunctionName( _
    ByRef funcStack() As String, ByRef stackTop As Long, ByVal funcName As String)
    If stackTop >= UBound(funcStack) Then
        If ROLL_DBG_PRINT Then Debug.Print "WARNING: funcStack overflow at depth " & stackTop & ". Function [" & funcName & "] dropped."
        Exit Sub
    End If
    stackTop = stackTop + 1: funcStack(stackTop) = UCase$(Trim$(funcName))
End Sub

Private Function CurrentFunctionName(ByRef funcStack() As String, ByVal stackTop As Long) As String
    Dim i As Long
    For i = stackTop To 1 Step -1
        If Len(funcStack(i)) > 0 Then CurrentFunctionName = funcStack(i): Exit Function
    Next i
End Function

Private Function TryParseBareIdentifier( _
    ByVal s As String, ByVal pos As Long, _
    ByRef tokenEnd As Long, ByRef identText As String) As Boolean

    Dim n As Long: n = Len(s)
    If pos < 1 Or pos > n Then Exit Function
    Dim ch As String: ch = Mid$(s, pos, 1)
    If Not IsLetterAZ(ch) And ch <> "_" Then Exit Function
    Dim p As Long: p = pos + 1
    Do While p <= n
        ch = Mid$(s, p, 1)
        If IsLetterAZ(ch) Or IsDigit09(ch) Or ch = "_" Or ch = "." Then p = p + 1 Else Exit Do
    Loop
    tokenEnd = p - 1: identText = Mid$(s, pos, tokenEnd - pos + 1)
    TryParseBareIdentifier = True
End Function

Private Function NextSignificantChar(ByVal s As String, ByVal pos As Long) As String
    Dim n As Long: n = Len(s): Dim p As Long: p = pos
    Do While p <= n
        If Mid$(s, p, 1) <> " " Then NextSignificantChar = Mid$(s, p, 1): Exit Function
        p = p + 1
    Loop
End Function

Private Function PrevNonSpacePos(ByVal s As String, ByVal pos As Long) As Long
    Dim p As Long: p = pos
    Do While p >= 1
        If Mid$(s, p, 1) <> " " Then PrevNonSpacePos = p: Exit Function
        p = p - 1
    Loop
End Function

Private Function NumericLiteralTokenEnd(ByVal s As String, ByVal pos As Long) As Long
    Dim n As Long: n = Len(s)
    Dim p As Long: p = pos
    If pos < 1 Or pos > n Then Exit Function
    Dim ch As String: ch = Mid$(s, p, 1)
    If ch = "+" Or ch = "-" Then p = p + 1
    Do While p <= n And IsDigit09(Mid$(s, p, 1)): p = p + 1: Loop
    If p <= n And Mid$(s, p, 1) = "." Then
        p = p + 1
        Do While p <= n And IsDigit09(Mid$(s, p, 1)): p = p + 1: Loop
    End If
    If p <= n Then
        ch = Mid$(s, p, 1)
        If ch = "E" Or ch = "e" Then
            p = p + 1
            If p <= n Then ch = Mid$(s, p, 1): If ch = "+" Or ch = "-" Then p = p + 1
            Do While p <= n And IsDigit09(Mid$(s, p, 1)): p = p + 1: Loop
        End If
    End If
    NumericLiteralTokenEnd = p - 1
End Function

Private Function IsNumericLiteralStart(ByVal s As String, ByVal pos As Long) As Boolean
    Dim n As Long: n = Len(s)
    If pos < 1 Or pos > n Then Exit Function
    Dim ch As String: ch = Mid$(s, pos, 1)
    Dim p  As Long: p = pos
    Dim prevCh As String, sawDigit As Boolean

    If ch = "+" Or ch = "-" Then
        If pos > 1 Then
            prevCh = Mid$(s, pos - 1, 1)
            If prevCh <> "(" And prevCh <> "," And prevCh <> ";" And prevCh <> "=" And _
               prevCh <> "+" And prevCh <> "-" And prevCh <> "*" And prevCh <> "/" And _
               prevCh <> "^" And prevCh <> "&" And prevCh <> "{" Then Exit Function
        End If
        If p >= n Then Exit Function
        ch = Mid$(s, p + 1, 1)
        If Not IsDigit09(ch) And ch <> "." Then Exit Function
        p = p + 1
    End If

    ch = Mid$(s, p, 1)
    If ch = "." Then
        If p >= n Or Not IsDigit09(Mid$(s, p + 1, 1)) Then Exit Function
        p = p + 1
        Do While p <= n And IsDigit09(Mid$(s, p, 1)): sawDigit = True: p = p + 1: Loop
    ElseIf IsDigit09(ch) Then
        Do While p <= n And IsDigit09(Mid$(s, p, 1)): sawDigit = True: p = p + 1: Loop
        If p <= n And Mid$(s, p, 1) = "." Then
            p = p + 1
            Do While p <= n And IsDigit09(Mid$(s, p, 1)): sawDigit = True: p = p + 1: Loop
        End If
    Else
        Exit Function
    End If

    If Not sawDigit Then Exit Function

    If p <= n Then
        ch = Mid$(s, p, 1)
        If ch = "E" Or ch = "e" Then
            p = p + 1
            If p <= n Then ch = Mid$(s, p, 1): If ch = "+" Or ch = "-" Then p = p + 1
            If p > n Or Not IsDigit09(Mid$(s, p, 1)) Then Exit Function
            Do While p <= n And IsDigit09(Mid$(s, p, 1)): p = p + 1: Loop
        End If
    End If

    If pos > 1 Then
        prevCh = Mid$(s, pos - 1, 1)
        If IsAlphaNumUnderscoreDot(prevCh) Or prevCh = "$" Or prevCh = "]" Then Exit Function
    End If
    If p <= n Then
        ch = Mid$(s, p, 1)
        If IsLetterAZ(ch) Or IsDigit09(ch) Or ch = "_" Or ch = "." Then Exit Function
    End If
    IsNumericLiteralStart = True
End Function

Private Function TryParseLongLiteral( _
    ByVal expr As String, ByRef outValue As Long, _
    ByVal mustBePositive As Boolean, ByVal allowZero As Boolean) As Boolean

    Dim s As String: s = Replace(expr, " ", "")
    If Len(s) = 0 Then Exit Function
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If i = 1 And (ch = "+" Or ch = "-") Then
        ElseIf Not IsDigit09(ch) Then Exit Function
        End If
    Next i
    outValue = CLng(s)
    If mustBePositive Then
        If allowZero Then
            If outValue < 0 Then Exit Function
        Else
            If outValue <= 0 Then Exit Function
        End If
    Else
        If Not allowZero And outValue = 0 Then Exit Function
    End If
    TryParseLongLiteral = True
End Function

Private Function FindTopLevelColon(ByVal s As String) As Long
    Dim i As Long: i = 1: Dim depth As Long: Dim ch As String
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then i = SkipDoubleQuotedString(s, i)
        ElseIf ch = "(" Then depth = depth + 1: i = i + 1
        ElseIf ch = ")" Then If depth > 0 Then depth = depth - 1: i = i + 1
        ElseIf ch = ":" And depth = 0 Then FindTopLevelColon = i: Exit Function
        Else i = i + 1
        End If
    Loop
End Function

Private Function FindLastBangOutsideQuotes(ByVal s As String) As Long
    Dim i As Long, inQuote As Boolean, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "'" Then
            If i < Len(s) And Mid$(s, i + 1, 1) = "'" Then i = i + 1 Else inQuote = Not inQuote
        ElseIf ch = "!" And Not inQuote Then
            FindLastBangOutsideQuotes = i
        End If
    Next i
End Function

Private Sub A1ToRowCol(ByVal addr As String, ByRef outRow As Long, ByRef outCol As Long)
    Dim i As Long: i = 1
    Dim colPart As String, rowPart As String
    Do While i <= Len(addr) And IsLetterAZ(Mid$(addr, i, 1)): colPart = colPart & Mid$(addr, i, 1): i = i + 1: Loop
    Do While i <= Len(addr) And IsDigit09(Mid$(addr, i, 1)): rowPart = rowPart & Mid$(addr, i, 1): i = i + 1: Loop
    outCol = ColLettersToNumber2(colPart): outRow = CLng(rowPart)
End Sub

Private Function RowColToA1(ByVal rowNum As Long, ByVal colNum As Long) As String
    RowColToA1 = ColumnNumberToLetters(colNum) & CStr(rowNum)
End Function

Private Sub AddRefKey(ByRef d As Object, ByVal sheetName As String, ByVal addr As String)
    Dim k As String: k = UCase$(sheetName) & "|" & UCase$(addr)
    If Not d.Exists(k) Then d.Add k, True
End Sub

' [OPT-17] Uses ColumnLettersToNumberRaw for pre-cleaned input.
Private Function ColLettersToNumber2(ByVal colLetters As String) As Long
    ColLettersToNumber2 = ColumnLettersToNumberRaw(UCase$(colLetters))
End Function

Private Function SkipSingleQuotedSheetQualifier(ByVal s As String, ByVal pos As Long) As Long
    Dim n As Long: n = Len(s)
    If pos < 1 Or pos > n Or Mid$(s, pos, 1) <> "'" Then Exit Function
    Dim p As Long: p = pos + 1
    Do While p <= n
        If Mid$(s, p, 1) = "'" Then
            If p < n And Mid$(s, p + 1, 1) = "'" Then
                p = p + 2
            Else
                p = p + 1
                Do While p <= n And Mid$(s, p, 1) = " ": p = p + 1: Loop
                If p <= n And Mid$(s, p, 1) = "!" Then SkipSingleQuotedSheetQualifier = p + 1
                Exit Function
            End If
        Else
            p = p + 1
        End If
    Loop
End Function

Private Function SkipDoubleQuotedString(ByVal s As String, ByVal pos As Long) As Long
    Dim n As Long: n = Len(s)
    pos = pos + 1
    Do While pos <= n
        If Mid$(s, pos, 1) = """" Then
            If pos < n And Mid$(s, pos + 1, 1) = """" Then pos = pos + 2 Else SkipDoubleQuotedString = pos + 1: Exit Function
        Else
            pos = pos + 1
        End If
    Loop
    SkipDoubleQuotedString = pos
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  BYTE-LEVEL CLASSIFIERS  [OPT-5, unchanged]
' ═══════════════════════════════════════════════════════════════════════════════
Private Function BIsLetter(ByVal b As Byte) As Boolean
    BIsLetter = ((b >= 65 And b <= 90) Or (b >= 97 And b <= 122))
End Function

Private Function BIsDigit(ByVal b As Byte) As Boolean
    BIsDigit = (b >= 48 And b <= 57)
End Function

Private Function BIsAlphaNumDotUnder(ByVal b As Byte) As Boolean
    BIsAlphaNumDotUnder = BIsLetter(b) Or BIsDigit(b) Or b = 95 Or b = 46
End Function

' [BUG-5] Byte-array version of SkipSingleQuotedSheetQualifier.
' Skips a 'Sheet Name'! qualifier in the byte array.
' startPos is 1-based (points to the opening single quote).
' n is the total string length.
' Returns the 1-based position AFTER the '!' (i.e. the next char to parse),
' or startPos if the pattern does not match.
Private Function SkipSingleQuotedSheetQualifierByte( _
    ByRef b() As Byte, ByVal startPos As Long, ByVal n As Long) As Long

    Dim p As Long: p = startPos + 1   ' skip opening quote
    Do While p <= n
        If b(p - 1) = BC_SQUOTE Then
            If p < n And b(p) = BC_SQUOTE Then
                p = p + 2   ' escaped '' inside sheet name
            Else
                ' Closing quote found
                p = p + 1
                ' Skip optional spaces before '!'
                Do While p <= n And b(p - 1) = BC_SPACE: p = p + 1: Loop
                If p <= n And b(p - 1) = BC_BANG Then
                    SkipSingleQuotedSheetQualifierByte = p + 1
                    Exit Function
                End If
                ' No '!' after closing quote — not a valid sheet qualifier
                SkipSingleQuotedSheetQualifierByte = startPos
                Exit Function
            End If
        Else
            p = p + 1
        End If
    Loop
    ' Unterminated quote — return startPos (no skip)
    SkipSingleQuotedSheetQualifierByte = startPos
End Function


' ═══════════════════════════════════════════════════════════════════════════════
'  STRING-BASED CHARACTER CLASSIFIERS  (unchanged)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function IsLetterAZ(ByVal ch As String) As Boolean
    Dim a As Integer
    If Len(ch) = 0 Then Exit Function
    a = Asc(UCase$(ch))
    IsLetterAZ = (a >= 65 And a <= 90)
End Function

Private Function IsDigit09(ByVal ch As String) As Boolean
    Dim a As Integer
    If Len(ch) = 0 Then Exit Function
    a = Asc(ch)
    IsDigit09 = (a >= 48 And a <= 57)
End Function

Private Function IsAlphaNumUnderscoreDot(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsAlphaNumUnderscoreDot = IsLetterAZ(ch) Or IsDigit09(ch) Or ch = "_" Or ch = "."
End Function

Private Function IsSheetQualifierChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsSheetQualifierChar = IsLetterAZ(ch) Or IsDigit09(ch) Or _
        ch = "_" Or ch = "." Or ch = "[" Or ch = "]"
End Function
