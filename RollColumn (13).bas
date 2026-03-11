Option Explicit

'=============================================================================
' ROLL COLUMN MACRO  v3
'
' Techniques from Wrap_and_Roll reference (incorporated):
'   - Array-based R1C1 reads; Value2-first + cell-by-cell formula writes
'   - IsExternalFormulaFast_R1C1: InStr-based, no RegExp
'   - HasSameRowRightRefR1C1_FAST + ReadLeadingSignedInt/Int: R1C1 right-ref detection
'   - ReplaceColRefs_SkipStrings_FAST + ReplaceColRefs_RAW_FAST:
'       safe A1 absolute column ref replacement (skips strings, skips Sheet!ref)
'   - Absolute ref fixes after both LEFT and RIGHT inserts
'   - Direct ws.Comments iteration for comment copy (simpler than pre-capture)
'   - Find-based GetLastUsedRow/Col
'   - FixSameRowOffset_Column: RC[n]→RC[m] drift correction (utility, not auto-called)
'   - Proper CleanFail / CleanExit error handling
'
' Our own additions (not in reference):
'   - Layout: Normal (Prev|Curr) / Reverse (Curr|Prev) / Ungrouped
'   - IsSandwichFormulaA1: precise "result between referenced cols" detection
'   - ProcessFigures: shape delete/free-float handling
'
' SheetList structure (Col A–D, data from row 2):
'   A = Sheet name  (blank → skip)
'   B = Column to roll (letter or number)
'   C = Direction: LEFT / RIGHT
'   D = Layout:  blank = Normal (Prev|Curr)
'                "Reverse"   = Reverse (Curr|Prev)
'                "Ungrouped" = no insert; copy formulas B→neighbour then ungroup
'=============================================================================

Private Const LOG_STEP_TIMES    As Boolean = False
Private Const SHOW_TOTAL_MSGBOX As Boolean = True

'========================
' ENTRY
'========================
Public Sub RollColumns()
    Dim t0 As Double: t0 = Timer

    Dim oldCalc   As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldStatus As Variant

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc   = Application.Calculation
    oldStatus = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual
    Application.StatusBar      = "RollColumns: Running..."

    On Error GoTo CleanFail

    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("SheetList")
    On Error GoTo CleanFail
    If wsConfig Is Nothing Then
        MsgBox "Cannot find sheet named 'SheetList'.", vbCritical
        GoTo CleanExit
    End If

    Dim cfgLastRow As Long
    cfgLastRow = GetLastUsedRow(wsConfig)
    If cfgLastRow < 2 Then
        MsgBox "No entries found in SheetList.", vbInformation
        GoTo CleanExit
    End If

    Dim errLog As String: errLog = ""
    Dim r As Long

    For r = 2 To cfgLastRow
        Dim sName  As String: sName  = Trim$(CStr(wsConfig.Cells(r, 1).Value2))
        Dim colRef As String: colRef = Trim$(CStr(wsConfig.Cells(r, 2).Value2))
        Dim dir    As String: dir    = UCase$(Trim$(CStr(wsConfig.Cells(r, 3).Value2)))
        Dim layout As String: layout = UCase$(Trim$(CStr(wsConfig.Cells(r, 4).Value2)))

        ' Skip blank rows and the SheetList sheet itself
        If sName = "" And colRef = "" Then GoTo NextRow
        If UCase$(sName) = "SHEETLIST" Then GoTo NextRow

        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sName)
        On Error GoTo CleanFail
        If ws Is Nothing Then
            errLog = errLog & "Sheet not found: " & sName & vbNewLine
            GoTo NextRow
        End If

        Dim targetCol As Long
        targetCol = ColumnToNumber(colRef)
        If targetCol < 1 Then
            errLog = errLog & "Invalid column '" & colRef & "' on '" & sName & "'" & vbNewLine
            GoTo NextRow
        End If

        '--- UNGROUPED: no insert; copy formulas into neighbour, then ungroup ---
        If layout = "UNGROUPED" Then
            If dir <> "LEFT" And dir <> "RIGHT" Then
                errLog = errLog & "Ungrouped needs direction on '" & sName & "'" & vbNewLine
                GoTo NextRow
            End If
            Dim nbr As Long: nbr = IIf(dir = "RIGHT", targetCol + 1, targetCol - 1)
            If nbr < 1 Or nbr > ws.Columns.Count Then
                errLog = errLog & "Ungrouped neighbour out of range on '" & sName & "'" & vbNewLine
                GoTo NextRow
            End If
            Dim ugLR As Long: ugLR = GetLastUsedRow(ws)
            If ugLR > 0 Then CopyColumnFormulasAsIs ws, targetCol, nbr, ugLR
            On Error Resume Next
            ws.Columns(nbr).Ungroup
            On Error GoTo CleanFail
            GoTo NextRow
        End If

        '--- NORMAL ROLL ---
        If dir <> "LEFT" And dir <> "RIGHT" Then
            errLog = errLog & "Invalid direction '" & dir & "' on '" & sName & "'" & vbNewLine
            GoTo NextRow
        End If

        RollOneColumn ws, targetCol, dir, layout, errLog

NextRow:
        Set ws = Nothing
    Next r

CleanExit:
    Dim secs As Double: secs = ElapsedSeconds(t0)
    Application.StatusBar = "RollColumns: Done in " & Format$(secs, "0.00") & "s"
    Debug.Print "RollColumns total: " & Format$(secs, "0.00") & " seconds"
    If SHOW_TOTAL_MSGBOX Then
        If errLog <> "" Then
            MsgBox "Roll completed with warnings:" & vbNewLine & errLog, vbExclamation, "Roll Warnings"
        Else
            MsgBox "Roll completed in " & Format$(secs, "0.00") & "s", vbInformation, "Done"
        End If
    End If
    Application.CutCopyMode    = False
    Application.Calculation    = oldCalc
    Application.EnableEvents   = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar      = oldStatus
    Exit Sub

CleanFail:
    errLog = errLog & "Error: " & Err.Description & vbNewLine
    Resume CleanExit
End Sub

Private Function ElapsedSeconds(ByVal t0 As Double) As Double
    Dim t1 As Double: t1 = Timer
    If t1 < t0 Then t1 = t1 + 86400#
    ElapsedSeconds = t1 - t0
End Function

'=============================================================================
' CORE ROLL
'=============================================================================
Private Sub RollOneColumn(ByVal ws As Worksheet, ByVal targetCol As Long, _
                          ByVal insertDir As String, ByVal layout As String, _
                          ByRef errLog As String)

    Dim tStep   As Double: tStep = Timer
    Dim lastRow As Long:   lastRow = GetLastUsedRow(ws)
    Dim lastCol As Long:   lastCol = GetLastUsedCol(ws)
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    '--- Insert and establish physical col indices ---
    Dim newColIdx As Long   ' the blank inserted column
    Dim oldColIdx As Long   ' the column that already held data

    If insertDir = "RIGHT" Then
        ws.Columns(targetCol + 1).Insert Shift:=xlToRight
        oldColIdx = targetCol
        newColIdx = targetCol + 1
    Else ' LEFT
        ws.Columns(targetCol).Insert Shift:=xlToRight
        newColIdx = targetCol
        oldColIdx = targetCol + 1
    End If
    lastCol = lastCol + 1
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " insert: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- Determine Current / Previous by layout ---
    ' Normal  (Prev|Curr): RIGHT→new=Curr,  LEFT→new=Prev
    ' Reverse (Curr|Prev): RIGHT→new=Prev,  LEFT→new=Curr
    Dim isReverse    As Boolean: isReverse    = (layout = "REVERSE")
    Dim newIsCurrent As Boolean
    If isReverse Then
        newIsCurrent = (insertDir = "LEFT")
    Else
        newIsCurrent = (insertDir = "RIGHT")
    End If

    Dim currentCol  As Long
    Dim previousCol As Long
    If newIsCurrent Then
        currentCol  = newColIdx
        previousCol = oldColIdx
    Else
        currentCol  = oldColIdx
        previousCol = newColIdx
    End If

    Dim colOffset As Long: colOffset = newColIdx - oldColIdx  ' +1 RIGHT, -1 LEFT

    '--- Copy formats from existing col into new blank col ---
    tStep = Timer
    CopyColumnFormats ws, oldColIdx, newColIdx
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " formats: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- COMMENTS ---
    ' Previous col always gets the comments; Current col is always blank.
    ' When newIsCurrent=True:  previousCol=oldColIdx keeps its comments naturally,
    '                          just clear currentCol=newColIdx (blank anyway).
    ' When newIsCurrent=False: copy comments old→new (previousCol=newColIdx),
    '                          then clear currentCol=oldColIdx.
    tStep = Timer
    If newIsCurrent Then
        ClearCommentsFast ws.Range(ws.Cells(1, currentCol), ws.Cells(lastRow, currentCol))
    Else
        CopyNotesWithFormatting_Column ws, oldColIdx, newColIdx
        ClearCommentsFast ws.Range(ws.Cells(1, currentCol), ws.Cells(lastRow, currentCol))
    End If
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " comments: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- FORMULAS: Current column ---
    ' Reads from oldColIdx in R1C1; writes to currentCol.
    ' R1C1 relative refs (RC[n]) are position-independent — no manual shift needed.
    ' All formulas kept (including external refs).
    tStep = Timer
    BuildCurrentColumn ws, ws.Name, oldColIdx, currentCol, lastRow
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " current formulas: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- FORMULAS: Previous column ---
    ' Rules applied to previousCol:
    '   1. External sheet ref  → freeze to value
    '   2. Same-row sandwich   → freeze to value
    '   3. Otherwise           → keep formula; shift relative refs if needed
    '
    ' Auto-shift analysis:
    '   INSERT RIGHT: oldColIdx stays in place → NOT auto-shifted by Excel
    '   INSERT LEFT:  oldColIdx shifts right   → IS auto-shifted by Excel
    '   newColIdx (blank insert):              → never auto-shifted
    '
    ' prevNeedsShift: does the previousCol need an explicit relative-ref shift?
    tStep = Timer
    Dim prevNeedsShift As Boolean
    If newIsCurrent Then
        ' previousCol = oldColIdx. Auto-shifted only on INSERT LEFT.
        prevNeedsShift = (colOffset > 0)   ' True only for INSERT RIGHT
    Else
        ' previousCol = newColIdx (blank). Never auto-shifted.
        prevNeedsShift = True
    End If
    BuildPreviousColumn ws, ws.Name, previousCol, lastRow, colOffset, prevNeedsShift
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " previous formulas: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- FIX ABSOLUTE A1 COLUMN REFS ---
    ' Relative R1C1 refs are already correct. Absolute A1 refs in copied formulas
    ' still point at old column numbers and need updating.
    tStep = Timer
    If insertDir = "RIGHT" Then
        ' Formulas in newCol (currentCol or previousCol) copied from oldCol:
        ' absolute refs to targetCol-1 (old Prev) → targetCol (new Prev)
        ' absolute refs to targetCol-2 (two-back)  → targetCol-1
        FixAbsoluteRefs_RightInsert ws, targetCol, newColIdx, lastRow, lastCol
    Else ' LEFT
        ' In inserted col (newColIdx): refs targetCol-2 → targetCol-1
        ' In cols to the right (oldColIdx..lastCol): refs targetCol-1 → targetCol
        FixAbsoluteRefs_LeftInsert ws, newColIdx, lastRow, lastCol
    End If
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " abs ref fix: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- FIGURES ---
    tStep = Timer
    ProcessFigures ws, currentCol, previousCol
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " figures: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

End Sub

'=============================================================================
' BUILD CURRENT COLUMN
' Reads from oldColIdx in R1C1 form, writes to currentCol.
' R1C1 relative refs are position-independent → no manual shift needed.
' All formulas kept (including external). Value cells copied as values.
' Strategy: write all values first via array (fast), then set formula cells
' individually — avoids the mixed-type array write that silently fails in Excel.
'=============================================================================
Private Sub BuildCurrentColumn(ByVal ws As Worksheet, ByVal sheetName As String, _
                                ByVal oldColIdx As Long, ByVal currentCol As Long, _
                                ByVal lastRow As Long)

    Dim rngOld As Range: Set rngOld = ws.Range(ws.Cells(1, oldColIdx), ws.Cells(lastRow, oldColIdx))
    Dim rngNew As Range: Set rngNew = ws.Range(ws.Cells(1, currentCol), ws.Cells(lastRow, currentCol))

    Dim arrVal As Variant: arrVal = rngOld.Value2
    Dim arrF   As Variant: arrF   = rngOld.FormulaR1C1

    rngNew.Value2 = arrVal  ' write values first (array write, fast)

    Dim i As Long, f As String
    For i = 1 To UBound(arrF, 1)
        If VarType(arrF(i, 1)) = vbString Then
            f = CStr(arrF(i, 1))
            If Len(f) > 0 And Left$(f, 1) = "=" Then
                rngNew.Cells(i, 1).FormulaR1C1 = f   ' keep as-is: R1C1 is position-independent
            End If
        End If
    Next i
End Sub

'=============================================================================
' BUILD PREVIOUS COLUMN
' Rules:
'   1. External sheet ref  → leave as value (already written by Value2 step)
'   2. Same-row sandwich   → leave as value
'   3. Needs shift         → ShiftR1C1RelativeByOffset
'   4. Otherwise           → restore formula as-is
'
' prevNeedsShift: True when previousCol was NOT auto-shifted by Excel's insert.
'=============================================================================
Private Sub BuildPreviousColumn(ByVal ws As Worksheet, ByVal sheetName As String, _
                                 ByVal previousCol As Long, ByVal lastRow As Long, _
                                 ByVal colOffset As Long, ByVal prevNeedsShift As Boolean)

    Dim rng As Range: Set rng = ws.Range(ws.Cells(1, previousCol), ws.Cells(lastRow, previousCol))

    Dim arrVal As Variant: arrVal = rng.Value2
    Dim arrF   As Variant: arrF   = rng.FormulaR1C1
    Dim arrA1  As Variant: arrA1  = rng.Formula       ' A1 form for sandwich check

    rng.Value2 = arrVal  ' write values first; rules 1 and 2 leave cells as-is

    Dim i As Long, f As String, fA1 As String
    For i = 1 To UBound(arrF, 1)
        If VarType(arrF(i, 1)) = vbString Then
            f = CStr(arrF(i, 1))
            If Len(f) > 0 And Left$(f, 1) = "=" Then
                fA1 = CStr(arrA1(i, 1))

                If IsExternalFormulaFast_R1C1(f, sheetName) Then
                    ' Rule 1: external → value already written, skip

                ElseIf IsSandwichFormulaA1(fA1, previousCol, i) Then
                    ' Rule 2: sandwich → value already written, skip

                ElseIf prevNeedsShift And colOffset <> 0 Then
                    ' Rule 3: shift relative R1C1 col refs by colOffset
                    rng.Cells(i, 1).FormulaR1C1 = ShiftR1C1RelativeByOffset(f, colOffset)

                Else
                    ' Rule 4: formula already correct (auto-shifted by Excel or same col)
                    rng.Cells(i, 1).FormulaR1C1 = f
                End If
            End If
        End If
    Next i
End Sub

'=============================================================================
' SHIFT R1C1 RELATIVE COL REFS
' Scans the R1C1 formula string and rewrites every C[n] to C[n+colOffset].
' Absolute column refs (C17, not C[n]) are left unchanged.
' External sheet refs are left unchanged.
'=============================================================================
Private Function ShiftR1C1RelativeByOffset(ByVal fR1C1 As String, ByVal colOffset As Long) As String
    Dim out As String: out = ""
    Dim i   As Long:  i   = 1
    Dim n   As Long:  n   = Len(fR1C1)

    Do While i <= n
        If Mid$(fR1C1, i, 2) = "C[" Then
            Dim bracketEnd As Long: bracketEnd = InStr(i + 2, fR1C1, "]")
            If bracketEnd > 0 Then
                Dim numStr As String: numStr = Mid$(fR1C1, i + 2, bracketEnd - (i + 2))
                If IsNumeric(numStr) Then
                    out = out & "C[" & CStr(CLng(numStr) + colOffset) & "]"
                    i = bracketEnd + 1
                    GoTo NextChar
                End If
            End If
        End If
        out = out & Mid$(fR1C1, i, 1)
        i = i + 1
NextChar:
    Loop
    ShiftR1C1RelativeByOffset = out
End Function

'=============================================================================
' FIX ABSOLUTE A1 COLUMN REFS — RIGHT INSERT
' After inserting a new col at targetCol+1:
'   Formulas in newColIdx (copied from oldCol) have stale absolute refs:
'     col(targetCol-1) → col(targetCol)    [old-Prev becomes new-Prev]
'     col(targetCol-2) → col(targetCol-1)  [two-back becomes old-Prev]
'   Same fixes applied to all cols to the right of newColIdx.
'   External sheet refs (Sheet!Q) are NOT touched.
'=============================================================================
Private Sub FixAbsoluteRefs_RightInsert(ByVal ws As Worksheet, _
                                        ByVal targetCol As Long, ByVal newColIdx As Long, _
                                        ByVal lastRow As Long, ByVal lastCol As Long)
    If targetCol - 2 < 1 Then Exit Sub

    Dim from1 As String: from1 = NumberToColumn(targetCol - 1)
    Dim to1   As String: to1   = NumberToColumn(targetCol)
    Dim from2 As String: from2 = NumberToColumn(targetCol - 2)
    Dim to2   As String: to2   = NumberToColumn(targetCol - 1)

    ' Fix in the new column (formulas copied from old)
    FixAbsoluteRefs_InColumn ws, newColIdx, lastRow, ws.Name, from1, to1, from2, to2

    ' Fix in all columns to the right of newColIdx
    If newColIdx < lastCol Then
        FixAbsoluteRefs_InRange ws, newColIdx + 1, lastCol, lastRow, ws.Name, from1, to1, from2, to2
    End If
End Sub

'=============================================================================
' FIX ABSOLUTE A1 COLUMN REFS — LEFT INSERT
' After inserting a new col at targetCol (= newColIdx):
'   In inserted col itself:
'     col(newColIdx-2) → col(newColIdx-1)
'   In cols to the right of inserted col (oldColIdx onwards):
'     col(newColIdx-1) → col(newColIdx)
'=============================================================================
Private Sub FixAbsoluteRefs_LeftInsert(ByVal ws As Worksheet, _
                                       ByVal newColIdx As Long, _
                                       ByVal lastRow As Long, ByVal lastCol As Long)
    ' Fix in the inserted col (two-back → one-back)
    If newColIdx - 2 >= 1 Then
        Dim fromA As String: fromA = NumberToColumn(newColIdx - 2)
        Dim toA   As String: toA   = NumberToColumn(newColIdx - 1)
        FixAbsoluteRefs_InColumn ws, newColIdx, lastRow, ws.Name, fromA, toA, "", ""
    End If

    ' Fix in cols to the right (one-back → inserted)
    If newColIdx - 1 >= 1 And newColIdx + 1 <= lastCol Then
        Dim fromB As String: fromB = NumberToColumn(newColIdx - 1)
        Dim toB   As String: toB   = NumberToColumn(newColIdx)
        FixAbsoluteRefs_InRange ws, newColIdx + 1, lastCol, lastRow, ws.Name, fromB, toB, "", ""
    End If
End Sub

' Apply up to two from→to substitutions on a single column (array-based)
Private Sub FixAbsoluteRefs_InColumn(ByVal ws As Worksheet, ByVal colIdx As Long, _
                                     ByVal lastRow As Long, ByVal sheetName As String, _
                                     ByVal from1 As String, ByVal to1 As String, _
                                     ByVal from2 As String, ByVal to2 As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, colIdx), ws.Cells(lastRow, colIdx))
    Dim arr As Variant: arr = rng.Formula

    Dim i As Long, f As String
    For i = 1 To UBound(arr, 1)
        If VarType(arr(i, 1)) = vbString Then
            f = CStr(arr(i, 1))
            If Len(f) > 0 And Left$(f, 1) = "=" Then
                If Not IsExternalFormulaFast_R1C1(rng.Cells(i, 1).FormulaR1C1, sheetName) Then
                    If from1 <> "" Then f = ReplaceColRefs_SkipStrings_FAST(f, from1, to1)
                    If from2 <> "" Then f = ReplaceColRefs_SkipStrings_FAST(f, from2, to2)
                    arr(i, 1) = f
                End If
            End If
        End If
    Next i
    rng.Formula = arr
End Sub

' Apply up to two from→to substitutions across a column range, formula cells only
Private Sub FixAbsoluteRefs_InRange(ByVal ws As Worksheet, _
                                    ByVal fromColIdx As Long, ByVal toColIdx As Long, _
                                    ByVal lastRow As Long, ByVal sheetName As String, _
                                    ByVal from1 As String, ByVal to1 As String, _
                                    ByVal from2 As String, ByVal to2 As String)
    Dim rng  As Range
    Dim rngF As Range
    Set rng = ws.Range(ws.Cells(1, fromColIdx), ws.Cells(lastRow, toColIdx))

    On Error Resume Next
    Set rngF = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rngF Is Nothing Then Exit Sub

    Dim c As Range, f As String
    For Each c In rngF.Cells
        If Not IsExternalFormulaFast_R1C1(c.FormulaR1C1, sheetName) Then
            f = c.Formula
            If from1 <> "" Then f = ReplaceColRefs_SkipStrings_FAST(f, from1, to1)
            If from2 <> "" Then f = ReplaceColRefs_SkipStrings_FAST(f, from2, to2)
            If f <> c.Formula Then c.Formula = f
        End If
    Next c
End Sub

'=============================================================================
' SANDWICH DETECTION (A1 formula space)
' Returns True when ALL of:
'   (a) formula contains at least one colon-range (not just comma-separated refs)
'   (b) every A1 cell ref is on formulaRow (no cross-row refs)
'   (c) resultCol is strictly BETWEEN min and max referenced column
'       (including INDIRECT("RC[n]",FALSE) endpoints)
'=============================================================================
Private Function IsSandwichFormulaA1(ByVal fA1 As String, _
                                     ByVal resultCol As Long, _
                                     ByVal formulaRow As Long) As Boolean
    IsSandwichFormulaA1 = False
    If Len(fA1) = 0 Or Left$(fA1, 1) <> "=" Then Exit Function
    If InStr(fA1, ":") = 0 Then Exit Function

    Dim hasDirectRange   As Boolean: hasDirectRange   = HasColonRangeA1(fA1)
    Dim hasIndirectRange As Boolean
    hasIndirectRange = (InStr(1, fA1, "INDIRECT", vbTextCompare) > 0)

    If Not hasDirectRange And Not hasIndirectRange Then Exit Function

    Dim minCol As Long: minCol = 2147483647
    Dim maxCol As Long: maxCol = 0
    Dim foundAny As Boolean: foundAny = False

    Dim fUp  As String: fUp  = UCase$(fA1)
    Dim fLen As Long:   fLen = Len(fA1)
    Dim pos  As Long:   pos  = 1

    Do While pos <= fLen
        ' Skip string literals
        If Mid$(fA1, pos, 1) = """" Then
            Dim qEnd As Long: qEnd = InStr(pos + 1, fA1, """")
            If qEnd = 0 Then Exit Do
            pos = qEnd + 1
            GoTo NextChar
        End If

        ' Try to match $?[A-Z]{1,3}$?[0-9]+ at pos
        Dim cs As Long: cs = pos
        If Mid$(fUp, pos, 1) = "$" Then cs = pos + 1

        Dim cLetters As String: cLetters = ""
        Dim k As Long: k = cs
        Do While k <= fLen
            Dim ch As String: ch = Mid$(fUp, k, 1)
            If ch >= "A" And ch <= "Z" Then
                cLetters = cLetters & ch: k = k + 1
            Else: Exit Do
            End If
        Loop

        If Len(cLetters) >= 1 And Len(cLetters) <= 3 Then
            Dim rs As Long: rs = k
            If rs <= fLen And Mid$(fA1, rs, 1) = "$" Then rs = rs + 1
            Dim rDigits As String: rDigits = ""
            Dim rk As Long: rk = rs
            Do While rk <= fLen
                Dim dc As String: dc = Mid$(fA1, rk, 1)
                If dc >= "0" And dc <= "9" Then
                    rDigits = rDigits & dc: rk = rk + 1
                Else: Exit Do
                End If
            Loop
            If Len(rDigits) > 0 Then
                If CLng(rDigits) <> formulaRow Then Exit Function  ' cross-row → not sandwich
                Dim rc As Long: rc = ColumnToNumber(cLetters)
                If rc < minCol Then minCol = rc
                If rc > maxCol Then maxCol = rc
                foundAny = True
                pos = rk
                GoTo NextChar
            End If
        End If

        pos = pos + 1
NextChar:
    Loop

    ' Resolve INDIRECT("RC[n]",FALSE) endpoints
    If hasIndirectRange Then
        Dim iPos As Long: iPos = InStr(1, fA1, "INDIRECT", vbTextCompare)
        Do While iPos > 0
            Dim qo As Long: qo = InStr(iPos, fA1, """")
            If qo = 0 Then Exit Do
            Dim qc As Long: qc = InStr(qo + 1, fA1, """")
            If qc = 0 Then Exit Do
            Dim r1c1s As String: r1c1s = UCase$(Mid$(fA1, qo + 1, qc - qo - 1))
            If Left$(r1c1s, 2) = "RC" Then
                Dim cp As String: cp = Mid$(r1c1s, 3)
                Dim icn As Long: icn = 0
                If Len(cp) = 0 Then
                    icn = resultCol
                ElseIf Left$(cp, 1) = "[" Then
                    Dim cbe As Long: cbe = InStr(cp, "]")
                    If cbe > 0 Then
                        Dim os As String: os = Mid$(cp, 2, cbe - 2)
                        If IsNumeric(os) Then icn = resultCol + CLng(os)
                    End If
                ElseIf IsNumeric(cp) Then
                    icn = CLng(cp)
                End If
                If icn > 0 Then
                    foundAny = True
                    If icn < minCol Then minCol = icn
                    If icn > maxCol Then maxCol = icn
                End If
            Else
                Exit Function  ' R offset in INDIRECT → cross-row → not sandwich
            End If
            iPos = InStr(qc + 1, fA1, "INDIRECT", vbTextCompare)
        Loop
    End If

    If Not foundAny Then Exit Function
    IsSandwichFormulaA1 = (resultCol > minCol And resultCol < maxCol)
End Function

Private Function HasColonRangeA1(ByVal fA1 As String) As Boolean
    Dim p As Long: p = InStr(fA1, ":")
    Do While p > 0
        If p > 1 Then
            If Mid$(fA1, p - 1, 1) >= "0" And Mid$(fA1, p - 1, 1) <= "9" Then
                HasColonRangeA1 = True: Exit Function
            End If
        End If
        p = InStr(p + 1, fA1, ":")
    Loop
End Function

'=============================================================================
' FAST EXTERNAL FORMULA DETECTION (no RegExp) — from reference
' Handles: 'Sheet Name'!, SheetName!, [Workbook.xlsx]Sheet!
'=============================================================================
Private Function IsExternalFormulaFast_R1C1(ByVal f As String, _
                                            ByVal thisSheetName As String) As Boolean
    Dim p As Long, q1 As Long, q0 As Long
    Dim token As String, nm As String
    Dim sn As String:      sn      = LCase$(thisSheetName)
    Dim scanPos As Long:   scanPos = 1

    Do
        p = InStr(scanPos, f, "!", vbTextCompare)
        If p = 0 Then Exit Do

        ' [Workbook]! pattern → external workbook
        If InStr(1, Left$(f, p), "[") > 0 And InStr(1, Left$(f, p), "]") > 0 Then
            IsExternalFormulaFast_R1C1 = True: Exit Function
        End If

        If p > 1 And Mid$(f, p - 1, 1) = "'" Then
            ' 'Sheet Name'!
            q1 = p - 1
            q0 = InStrRev(f, "'", q1 - 1)
            If q0 > 0 Then
                token = Mid$(f, q0 + 1, q1 - q0 - 1)
                nm = LCase$(Trim$(token))
                If nm <> sn Then IsExternalFormulaFast_R1C1 = True: Exit Function
            End If
        Else
            ' SheetName!
            Dim k As Long: k = p - 1
            Do While k > 0
                Dim ch2 As String: ch2 = Mid$(f, k, 1)
                If (ch2 Like "[A-Za-z0-9_]") Or ch2 = "." Or ch2 = "-" Then
                    k = k - 1
                ElseIf ch2 = "]" Then
                    IsExternalFormulaFast_R1C1 = True: Exit Function
                Else
                    Exit Do
                End If
            Loop
            token = Mid$(f, k + 1, p - (k + 1))
            nm = LCase$(Trim$(token))
            If nm <> "" And nm <> sn Then IsExternalFormulaFast_R1C1 = True: Exit Function
        End If

        scanPos = p + 1
    Loop
    IsExternalFormulaFast_R1C1 = False
End Function

'=============================================================================
' FAST SAME-ROW RIGHT-REF DETECTION (R1C1) — from reference
' Returns True if the formula references any column to the RIGHT of srcCol.
' Checks RC[+n] (positive relative offset) and absolute R{srcRow}C{>srcCol}.
'=============================================================================
Private Function HasSameRowRightRefR1C1_FAST(ByVal fR1C1 As String, _
                                             ByVal srcRow As Long, _
                                             ByVal srcCol As Long) As Boolean
    Dim pos As Long, n As Long, s As String

    ' Check for positive RC[+n] (column to the right)
    pos = 1
    Do
        pos = InStr(pos, fR1C1, "C[", vbTextCompare)
        If pos = 0 Then Exit Do
        s = Mid$(fR1C1, pos + 2)
        n = ReadLeadingSignedInt(s)
        If n > 0 Then HasSameRowRightRefR1C1_FAST = True: Exit Function
        pos = pos + 2
    Loop

    ' Check for absolute same-row ref R{srcRow}C{>srcCol}
    Dim key As String: key = "R" & CStr(srcRow) & "C"
    pos = InStr(1, fR1C1, key, vbTextCompare)
    Do While pos > 0
        s = Mid$(fR1C1, pos + Len(key))
        n = ReadLeadingInt(s)
        If n > srcCol Then HasSameRowRightRefR1C1_FAST = True: Exit Function
        pos = InStr(pos + Len(key), fR1C1, key, vbTextCompare)
    Loop

    HasSameRowRightRefR1C1_FAST = False
End Function

Private Function ReadLeadingSignedInt(ByVal s As String) As Long
    Dim i As Long, ch As String, sign As Long
    sign = 1: i = 1
    If Len(s) = 0 Then Exit Function
    ch = Mid$(s, 1, 1)
    If ch = "+" Then sign = 1:  i = 2
    If ch = "-" Then sign = -1: i = 2
    Dim v As Long: v = 0
    For i = i To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
        v = v * 10 + (Asc(ch) - 48)
    Next i
    ReadLeadingSignedInt = sign * v
End Function

Private Function ReadLeadingInt(ByVal s As String) As Long
    Dim i As Long, ch As String, v As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
        v = v * 10 + (Asc(ch) - 48)
    Next i
    ReadLeadingInt = v
End Function

'=============================================================================
' REPLACE ABSOLUTE A1 COL REFS — SAFE (from reference)
' Outer function skips content inside string literals "..."
' Inner function does char-by-char replacement with boundary checks:
'   - preceded by "!" → skip (Sheet!ColRef, don't touch)
'   - preceded by letter/digit/_ → skip (part of a name)
'   - followed by digit, ":", "$", or end → valid col ref → replace
'   - handles "$Q" (absolute) and "Q" (bare column letter) forms
'=============================================================================
Private Function ReplaceColRefs_SkipStrings_FAST(ByVal formulaText As String, _
                                                  ByVal fromCol As String, _
                                                  ByVal toCol As String) As String
    Dim i As Long, j As Long
    Dim inQ As Boolean: inQ = False
    Dim seg As String, out As String

    i = 1
    Do While i <= Len(formulaText)
        j = InStr(i, formulaText, """")
        If j = 0 Then
            seg = Mid$(formulaText, i)
            If Not inQ Then seg = ReplaceColRefs_RAW_FAST(seg, fromCol, toCol)
            out = out & seg
            Exit Do
        Else
            seg = Mid$(formulaText, i, j - i)
            If Not inQ Then seg = ReplaceColRefs_RAW_FAST(seg, fromCol, toCol)
            out = out & seg & """"
            inQ = Not inQ
            i = j + 1
        End If
    Loop
    ReplaceColRefs_SkipStrings_FAST = out
End Function

Private Function ReplaceColRefs_RAW_FAST(ByVal s As String, _
                                          ByVal fromCol As String, _
                                          ByVal toCol As String) As String
    Dim i As Long, n As Long
    Dim uFrom As String: uFrom = UCase$(fromCol)
    Dim uTo   As String: uTo   = UCase$(toCol)

    n = Len(s)
    If n = 0 Then ReplaceColRefs_RAW_FAST = s: Exit Function

    Dim out As String: out = ""
    i = 1
    Do While i <= n
        Dim ch As String: ch = Mid$(s, i, 1)

        If ch = "$" Or (UCase$(ch) >= "A" And UCase$(ch) <= "Z") Then
            Dim startI    As Long:   startI    = i
            Dim hasDollar As Boolean: hasDollar = (ch = "$")
            Dim k         As Long:   k          = i + IIf(hasDollar, 1, 0)

            If k + Len(uFrom) - 1 <= n Then
                Dim cand As String: cand = UCase$(Mid$(s, k, Len(uFrom)))
                If cand = uFrom Then
                    Dim prev As String: prev = IIf(startI = 1, "", Mid$(s, startI - 1, 1))
                    If prev <> "!" And Not (prev Like "[A-Za-z0-9_]") Then
                        Dim nextPos As Long: nextPos = k + Len(uFrom)
                        Dim nxt     As String: nxt = IIf(nextPos <= n, Mid$(s, nextPos, 1), "")
                        If (nxt Like "[0-9]") Or nxt = ":" Or nxt = "$" Or nxt = "" Then
                            out = out & IIf(hasDollar, "$", "") & uTo
                            i = k + Len(uFrom)
                            GoTo ContinueLoop
                        End If
                    End If
                End If
            End If
        End If

        out = out & ch
        i = i + 1
ContinueLoop:
    Loop
    ReplaceColRefs_RAW_FAST = out
End Function

'=============================================================================
' FIX SAME-ROW RC OFFSET DRIFT (utility — call manually if needed)
' In a column, replaces RC[fromOffset] with RC[toOffset] (and R[0]C[n] form).
' Useful when a column roll causes formulas that said "2 steps away" to now
' correctly say "1 step away". Array-based, fast.
' Example: FixSameRowOffset_Column ws, oldCol, 1, lastRow, 2, 1
'=============================================================================
Private Sub FixSameRowOffset_Column(ByVal ws As Worksheet, ByVal colIdx As Long, _
                                    ByVal startRow As Long, ByVal endRow As Long, _
                                    ByVal fromOffset As Long, ByVal toOffset As Long)
    Dim rng As Range: Set rng = ws.Range(ws.Cells(startRow, colIdx), ws.Cells(endRow, colIdx))
    Dim arr As Variant: arr = rng.FormulaR1C1

    Dim a As String: a = "RC["    & fromOffset & "]"
    Dim b As String: b = "RC["    & toOffset   & "]"
    Dim c As String: c = "R[0]C[" & fromOffset & "]"
    Dim d As String: d = "R[0]C[" & toOffset   & "]"

    Dim i As Long, f As String
    For i = 1 To UBound(arr, 1)
        If VarType(arr(i, 1)) = vbString Then
            f = CStr(arr(i, 1))
            If Len(f) > 0 And Left$(f, 1) = "=" Then
                If InStr(1, f, a, vbTextCompare) > 0 Or InStr(1, f, c, vbTextCompare) > 0 Then
                    f = Replace(f, c, d, , , vbTextCompare)
                    f = Replace(f, a, b, , , vbTextCompare)
                    arr(i, 1) = f
                End If
            End If
        End If
    Next i
    rng.FormulaR1C1 = arr
End Sub

'=============================================================================
' UNGROUPED: copy formulas as-is from srcCol to dstCol (no roll, no value rules)
'=============================================================================
Private Sub CopyColumnFormulasAsIs(ByVal ws As Worksheet, _
                                    ByVal srcCol As Long, ByVal dstCol As Long, _
                                    ByVal lastRow As Long)
    Dim rngSrc As Range: Set rngSrc = ws.Range(ws.Cells(1, srcCol), ws.Cells(lastRow, srcCol))
    Dim rngDst As Range: Set rngDst = ws.Range(ws.Cells(1, dstCol), ws.Cells(lastRow, dstCol))

    Dim arrVal As Variant: arrVal = rngSrc.Value2
    Dim arrF   As Variant: arrF   = rngSrc.FormulaR1C1

    rngDst.Value2 = arrVal  ' write values first

    Dim i As Long, f As String
    For i = 1 To UBound(arrF, 1)
        If VarType(arrF(i, 1)) = vbString Then
            f = CStr(arrF(i, 1))
            If Len(f) > 0 And Left$(f, 1) = "=" Then
                rngDst.Cells(i, 1).FormulaR1C1 = f
            End If
        End If
    Next i

    CopyColumnFormats ws, srcCol, dstCol
End Sub

'=============================================================================
' COMMENTS — direct copy using ws.Comments (from reference pattern)
' Iterates existing comment objects only (not all rows).
' Preserves rich-text formatting (bold/italic/underline runs).
'=============================================================================
Private Sub CopyNotesWithFormatting_Column(ByVal ws As Worksheet, _
                                           ByVal oldColIdx As Long, _
                                           ByVal newColIdx As Long)
    Dim cmt     As Comment
    Dim srcCell As Range
    Dim dstCell As Range

    On Error Resume Next
    For Each cmt In ws.Comments
        Set srcCell = cmt.Parent
        If Not srcCell Is Nothing Then
            If srcCell.Column = oldColIdx Then
                Set dstCell = srcCell.Offset(0, newColIdx - oldColIdx)
                CopyLegacyNoteWithFormatting cmt, dstCell
            End If
        End If
    Next cmt
    On Error GoTo 0
End Sub

Private Sub CopyLegacyNoteWithFormatting(ByVal srcCmt As Comment, ByVal dstCell As Range)
    On Error Resume Next
    If Not dstCell.Comment Is Nothing Then dstCell.Comment.Delete
    On Error GoTo 0

    Dim srcTF As Object: Set srcTF = srcCmt.Shape.TextFrame
    Dim txt   As String: txt = srcTF.Characters.Text

    On Error Resume Next
    dstCell.AddComment txt
    On Error GoTo 0
    If dstCell.Comment Is Nothing Then Exit Sub

    On Error Resume Next
    dstCell.Comment.Shape.Width  = srcCmt.Shape.Width
    dstCell.Comment.Shape.Height = srcCmt.Shape.Height
    On Error GoTo 0

    Dim dstTF As Object: Set dstTF = dstCell.Comment.Shape.TextFrame

    On Error Resume Next
    With dstTF.Characters.Font
        .Name  = srcTF.Characters.Font.Name
        .Size  = srcTF.Characters.Font.Size
        .Color = srcTF.Characters.Font.Color
    End With
    On Error GoTo 0

    CopyCommentRichRuns srcTF, dstTF
End Sub

Private Sub CopyCommentRichRuns(ByVal srcTF As Object, ByVal dstTF As Object)
    Dim s As String: s = srcTF.Characters.Text
    Dim n As Long:   n = Len(s)
    If n <= 0 Then Exit Sub

    Dim i As Long, runStart As Long, runLen As Long
    Dim pb As Variant, pit As Variant, pu As Variant
    Dim b  As Variant, it  As Variant, u   As Variant

    On Error Resume Next
    pb = srcTF.Characters(1, 1).Font.Bold
    pit = srcTF.Characters(1, 1).Font.Italic
    pu = srcTF.Characters(1, 1).Font.Underline
    runStart = 1

    For i = 2 To n
        b  = srcTF.Characters(i, 1).Font.Bold
        it = srcTF.Characters(i, 1).Font.Italic
        u  = srcTF.Characters(i, 1).Font.Underline
        If (b <> pb) Or (it <> pit) Or (u <> pu) Then
            runLen = i - runStart
            With dstTF.Characters(runStart, runLen).Font
                .Bold = pb: .Italic = pit: .Underline = pu
            End With
            runStart = i: pb = b: pit = it: pu = u
        End If
    Next i
    runLen = n - runStart + 1
    With dstTF.Characters(runStart, runLen).Font
        .Bold = pb: .Italic = pit: .Underline = pu
    End With
    On Error GoTo 0
End Sub

Private Sub ClearCommentsFast(ByVal rng As Range)
    On Error Resume Next
    rng.ClearNotes
    rng.ClearComments
    On Error GoTo 0
End Sub

'=============================================================================
' FIGURES (Shapes)
' Current col  → delete shapes anchored to it
' Previous col → set FreeFloating (detach from column)
'=============================================================================
Private Sub ProcessFigures(ByVal ws As Worksheet, _
                            ByVal currentCol As Long, ByVal previousCol As Long)
    Dim shp      As Shape
    Dim toDelete As Collection
    Set toDelete = New Collection

    On Error Resume Next
    For Each shp In ws.Shapes
        Dim shpL As Long: shpL = GetShapeColumn(ws, shp, True)
        Dim shpR As Long: shpR = GetShapeColumn(ws, shp, False)
        If shpL = currentCol Or shpR = currentCol Then
            toDelete.Add shp
        ElseIf shpL = previousCol Or shpR = previousCol Then
            shp.Placement = xlFreeFloating
        End If
    Next shp
    On Error GoTo 0

    Dim s As Shape
    For Each s In toDelete
        On Error Resume Next
        s.Delete
        On Error GoTo 0
    Next s
End Sub

Private Function GetShapeColumn(ByVal ws As Worksheet, ByVal shp As Shape, _
                                 ByVal leftEdge As Boolean) As Long
    Dim pos    As Double: pos    = IIf(leftEdge, shp.Left, shp.Left + shp.Width)
    Dim maxC   As Long:   maxC   = ws.UsedRange.Column + ws.UsedRange.Columns.Count
    Dim c      As Long
    For c = 1 To maxC
        If ws.Cells(1, c).Left + ws.Cells(1, c).Width >= pos Then
            GetShapeColumn = c: Exit Function
        End If
    Next c
    GetShapeColumn = 0
End Function

'=============================================================================
' COLUMN FORMATS
'=============================================================================
Private Sub CopyColumnFormats(ByVal ws As Worksheet, ByVal fromCol As Long, ByVal toCol As Long)
    ws.Columns(fromCol).Copy
    ws.Columns(toCol).PasteSpecial Paste:=xlPasteFormats
    ws.Columns(toCol).ColumnWidth = ws.Columns(fromCol).ColumnWidth
    Application.CutCopyMode = False
End Sub

'=============================================================================
' UTILITIES
'=============================================================================
Private Function ColumnToNumber(ByVal colRef As Variant) As Long
    If IsNumeric(colRef) Then ColumnToNumber = CLng(colRef): Exit Function
    Dim s As String: s = UCase$(Trim$(Replace(CStr(colRef), "$", "")))
    Dim i As Long, ch As Integer, n As Long
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 65 Or ch > 90 Then Exit Function
        n = n * 26 + (ch - 64)
    Next i
    ColumnToNumber = n
End Function

Private Function NumberToColumn(ByVal colNum As Long) As String
    Dim n As Long: n = colNum
    NumberToColumn = ""
    Do While n > 0
        Dim r As Long: r = (n - 1) Mod 26
        NumberToColumn = Chr$(65 + r) & NumberToColumn
        n = (n - 1) \ 26
    Loop
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then GetLastUsedRow = 0 Else GetLastUsedRow = f.Row
End Function

Private Function GetLastUsedCol(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByColumns, _
                          SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then GetLastUsedCol = 0 Else GetLastUsedCol = f.Column
End Function
