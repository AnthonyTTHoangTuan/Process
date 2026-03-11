Option Explicit

'=============================================================================
' ROLL COLUMN MACRO
' Techniques learned from Wrap_and_Roll reference:
'   - Array-based R1C1 reads/writes (no cell-by-cell loops)
'   - R1C1 formula space for all formula logic (no regex)
'   - Fast InStr-based external sheet detection
'   - Fast same-row right-ref detection via RC[+n]
'   - ReplaceColRefs skipping string literals
'   - Rich-text comment copying
'   - Find-based lastRow/lastCol
'   - Proper error handling with cleanup
'
' SheetList structure:
'   Col A = Sheet Name
'   Col B = Column to roll (letter or number)
'   Col C = Insert Direction: "Left" or "Right"
'   Col D = Layout:
'           blank     = Normal   (Previous | Current)
'           "Reverse" = Reverse  (Current | Previous)
'           "Ungrouped" = no insert; copy formulas from B into neighbour
'                         per Col C direction, then ungroup that neighbour
'=============================================================================

Private Const LOG_STEP_TIMES   As Boolean = False
Private Const SHOW_TOTAL_MSGBOX As Boolean = True

'========================
' ENTRY
'========================
Public Sub RollColumns()
    Dim t0 As Double: t0 = Timer

    Dim oldCalc    As XlCalculation
    Dim oldEvents  As Boolean
    Dim oldScreen  As Boolean
    Dim oldStatus  As Variant

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
    Dim i As Long

    For i = 2 To cfgLastRow
        Dim sheetName As String
        Dim colRef    As String
        Dim insertDir As String
        Dim layout    As String

        sheetName = Trim$(CStr(wsConfig.Cells(i, 1).Value2))
        colRef    = Trim$(CStr(wsConfig.Cells(i, 2).Value2))
        insertDir = UCase$(Trim$(CStr(wsConfig.Cells(i, 3).Value2)))
        layout    = UCase$(Trim$(CStr(wsConfig.Cells(i, 4).Value2)))

        If sheetName = "" And colRef = "" Then GoTo NextRow

        Dim wsTarget As Worksheet
        Set wsTarget = Nothing
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        On Error GoTo CleanFail
        If wsTarget Is Nothing Then
            errLog = errLog & "Sheet not found: " & sheetName & vbNewLine
            GoTo NextRow
        End If

        Dim targetCol As Long
        targetCol = ColumnToNumber(colRef)
        If targetCol < 1 Then
            errLog = errLog & "Invalid column '" & colRef & "' on sheet '" & sheetName & "'" & vbNewLine
            GoTo NextRow
        End If

        '--- UNGROUPED: no insert, copy formulas to neighbour then ungroup ---
        If layout = "UNGROUPED" Then
            If insertDir <> "LEFT" And insertDir <> "RIGHT" Then
                errLog = errLog & "Ungrouped needs direction on '" & sheetName & "'" & vbNewLine
                GoTo NextRow
            End If
            Dim ungroupNeighbour As Long
            ungroupNeighbour = IIf(insertDir = "RIGHT", targetCol + 1, targetCol - 1)
            If ungroupNeighbour < 1 Or ungroupNeighbour > wsTarget.Columns.Count Then
                errLog = errLog & "Ungrouped neighbour out of range on '" & sheetName & "'" & vbNewLine
                GoTo NextRow
            End If
            Dim ugLastRow As Long
            ugLastRow = GetLastUsedRow(wsTarget)
            If ugLastRow > 0 Then
                Call CopyColumnFormulasAsIs(wsTarget, targetCol, ungroupNeighbour, ugLastRow)
            End If
            On Error Resume Next
            wsTarget.Columns(ungroupNeighbour).Ungroup
            On Error GoTo CleanFail
            GoTo NextRow
        End If

        '--- NORMAL ROLL ---
        If insertDir <> "LEFT" And insertDir <> "RIGHT" Then
            errLog = errLog & "Invalid direction '" & insertDir & "' on '" & sheetName & "'" & vbNewLine
            GoTo NextRow
        End If

        Call RollOneColumn(wsTarget, targetCol, insertDir, layout, errLog)

NextRow:
        Set wsTarget = Nothing
    Next i

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
    Application.CutCopyMode   = False
    Application.Calculation   = oldCalc
    Application.EnableEvents  = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar     = oldStatus
    Exit Sub

CleanFail:
    errLog = errLog & "Unexpected error: " & Err.Description & vbNewLine
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

    Dim tStep    As Double: tStep = Timer
    Dim lastRow  As Long:   lastRow = GetLastUsedRow(ws)
    Dim lastCol  As Long:   lastCol = GetLastUsedCol(ws)
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    '--- Determine new/existing col indices ---
    Dim newColIdx  As Long
    Dim oldColIdx  As Long   ' existingCol = was current before roll

    If insertDir = "RIGHT" Then
        ws.Columns(targetCol + 1).Insert Shift:=xlToRight
        newColIdx = targetCol + 1
        oldColIdx = targetCol
        lastCol   = lastCol + 1
    Else ' LEFT
        ws.Columns(targetCol).Insert Shift:=xlToRight
        newColIdx = targetCol
        oldColIdx = targetCol + 1
        lastCol   = lastCol + 1
    End If

    If LOG_STEP_TIMES Then Debug.Print ws.Name & " insert " & insertDir & ": " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- Determine which inserted col is Current, which is Previous ---
    ' Normal  (Prev|Curr): Left=>new=Prev, Right=>new=Curr
    ' Reverse (Curr|Prev): Left=>new=Curr, Right=>new=Prev
    Dim isReverse    As Boolean
    isReverse = (layout = "REVERSE" Or layout = "PREVRIGHT")

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
        previousCol = newColIdx
        currentCol  = oldColIdx
    End If

    Dim colOffset As Long
    colOffset = newColIdx - oldColIdx   ' +1 for RIGHT, -1 for LEFT

    '--- Copy column formats ---
    tStep = Timer
    CopyColumnFormats ws, oldColIdx, newColIdx
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " formats: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- COMMENTS (capture BEFORE any changes) ---
    ' Previous col = inherits comments from pre-roll current (oldColIdx)
    ' Current col  = blank
    tStep = Timer
    Dim savedComments As Object
    Set savedComments = CaptureComments(ws, oldColIdx)
    ' Clear current col comments
    ClearCommentsFast ws.Range(ws.Cells(1, currentCol), ws.Cells(lastRow, currentCol))
    ' Restore onto previous col
    ClearCommentsFast ws.Range(ws.Cells(1, previousCol), ws.Cells(lastRow, previousCol))
    RestoreCommentsWithFormatting ws, previousCol, savedComments
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " comments: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- FORMULAS: Current column ---
    ' Copy from oldColIdx, shifted by colOffset.
    ' Always keeps all formulas (including external).
    tStep = Timer
    BuildCurrentColumn_FAST ws, oldColIdx, newColIdx, currentCol, lastRow, colOffset
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " current formulas: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- FORMULAS: Previous column ---
    ' Excel auto-shift behaviour:
    '   INSERT RIGHT: oldColIdx stays => NOT auto-shifted => needs explicit shift
    '   INSERT LEFT:  oldColIdx shifts right => IS auto-shifted => no extra shift
    ' Additionally apply value-paste rules:
    '   - External sheet ref      => paste as value
    '   - Same-row sandwich range => paste as value
    tStep = Timer
    Dim prevNeedsShift As Boolean
    If newIsCurrent Then
        ' Previous = oldColIdx. Auto-shifted only on INSERT LEFT (colOffset=-1).
        prevNeedsShift = (colOffset > 0)   ' True for INSERT RIGHT
    Else
        ' Previous = newColIdx (blank insert). Never auto-shifted.
        prevNeedsShift = True
    End If
    BuildPreviousColumn_FAST ws, previousCol, lastRow, ws.Name, colOffset, prevNeedsShift
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " previous formulas: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

    '--- FIGURES ---
    tStep = Timer
    ProcessFigures ws, currentCol, previousCol
    If LOG_STEP_TIMES Then Debug.Print ws.Name & " figures: " & Format$(ElapsedSeconds(tStep), "0.00") & "s"

End Sub

'=============================================================================
' BUILD CURRENT COLUMN (array-based, R1C1)
' Copies formulas from oldColIdx into newColIdx (currentCol), all formulas kept.
' Column shift is handled by writing FormulaR1C1 directly — R1C1 relative refs
' (RC[n]) are position-independent, so no manual shifting needed.
' For the Current col we just copy arrF from old to new as-is in R1C1 space.
'=============================================================================
Private Sub BuildCurrentColumn_FAST(ByVal ws As Worksheet, _
                                    ByVal oldColIdx As Long, ByVal newColIdx As Long, _
                                    ByVal currentCol As Long, ByVal lastRow As Long, _
                                    ByVal colOffset As Long)

    Dim rngOld As Range, rngNew As Range
    Dim arrVal As Variant, arrF As Variant, arrOut As Variant
    Dim i As Long, f As String

    Set rngOld = ws.Range(ws.Cells(1, oldColIdx), ws.Cells(lastRow, oldColIdx))
    Set rngNew = ws.Range(ws.Cells(1, currentCol), ws.Cells(lastRow, currentCol))

    arrVal = rngOld.Value2
    arrF   = rngOld.FormulaR1C1
    arrOut = arrVal   ' default: values

    For i = 1 To UBound(arrF, 1)
        If VarType(arrF(i, 1)) = vbString Then
            f = CStr(arrF(i, 1))
            If Len(f) > 0 And Left$(f, 1) = "=" Then
                arrOut(i, 1) = f   ' keep formula as-is in R1C1 (relative refs auto-correct)
            End If
        End If
    Next i

    rngNew.FormulaR1C1 = arrOut
End Sub

'=============================================================================
' BUILD PREVIOUS COLUMN (array-based, R1C1)
' Applies value-paste rules to the previous column:
'   Rule 1: External sheet ref         => paste as value
'   Rule 2: Same-row sandwich range    => paste as value
'   Otherwise (and prevNeedsShift):    => formula already in place (auto-shifted
'             by Excel or copied), no change needed
' colOffset and prevNeedsShift control whether an explicit A1 column shift
' is applied (only needed when oldColIdx was NOT auto-shifted by Excel).
'=============================================================================
Private Sub BuildPreviousColumn_FAST(ByVal ws As Worksheet, _
                                     ByVal previousCol As Long, ByVal lastRow As Long, _
                                     ByVal thisSheetName As String, _
                                     ByVal colOffset As Long, _
                                     ByVal prevNeedsShift As Boolean)

    Dim rng    As Range
    Dim arrVal As Variant, arrF As Variant, arrOut As Variant
    Dim i As Long, f As String, fA1 As String

    Set rng  = ws.Range(ws.Cells(1, previousCol), ws.Cells(lastRow, previousCol))
    arrVal   = rng.Value2
    arrF     = rng.FormulaR1C1
    arrOut   = arrF   ' default: keep formula

    Dim arrA1 As Variant
    arrA1 = rng.Formula   ' for sandwich check (needs A1 cell ref text)

    For i = 1 To UBound(arrF, 1)
        If VarType(arrF(i, 1)) = vbString Then
            f = CStr(arrF(i, 1))
            If Len(f) > 0 And Left$(f, 1) = "=" Then
                fA1 = CStr(arrA1(i, 1))

                ' Rule 1: external ref => value
                If IsExternalFormulaFast_R1C1(f, thisSheetName) Then
                    arrOut(i, 1) = arrVal(i, 1)

                ' Rule 2: same-row sandwich => value
                ElseIf IsSandwichFormulaA1(fA1, previousCol, i) Then
                    arrOut(i, 1) = arrVal(i, 1)

                ' Rule 3: needs explicit col shift (prevNeedsShift)
                ElseIf prevNeedsShift And colOffset <> 0 Then
                    ' Shift relative A1 column refs by colOffset, skip strings/external
                    arrOut(i, 1) = ShiftR1C1RelativeByOffset(f, colOffset)
                End If
                ' else: formula already correct, leave as-is
            End If
        End If
    Next i

    rng.FormulaR1C1 = arrOut
End Sub

'=============================================================================
' SHIFT R1C1 FORMULA: adjust RC[n] offsets by colOffset
' Used when Previous col was NOT auto-shifted by Excel (INSERT RIGHT case).
' RC[n] => RC[n+colOffset], RC => RC[colOffset] (if colOffset<>0), RC[-1] etc.
' Absolute column refs (RC5) and row refs are left unchanged.
' External sheet refs are left unchanged.
'=============================================================================
Private Function ShiftR1C1RelativeByOffset(ByVal fR1C1 As String, ByVal colOffset As Long) As String
    ' Strategy: find all C[n] tokens and replace with C[n+colOffset]
    ' Also handle plain "C" (no bracket) = RC[0] relative => becomes RC[colOffset]
    ' We do this via string scanning (no regex)
    Dim result As String: result = fR1C1
    Dim i As Long, n As Long
    Dim ch As String

    ' Process in reverse so positions stay valid
    ' Build output left-to-right scanning
    Dim out As String: out = ""
    i = 1
    Do While i <= Len(result)
        ' Look for "C[" pattern
        If Mid$(result, i, 2) = "C[" Then
            ' Read the signed integer inside brackets
            Dim bracketEnd As Long
            bracketEnd = InStr(i + 2, result, "]")
            If bracketEnd > 0 Then
                Dim numStr As String
                numStr = Mid$(result, i + 2, bracketEnd - (i + 2))
                If IsNumeric(numStr) Then
                    Dim oldOff As Long: oldOff = CLng(numStr)
                    Dim newOff As Long: newOff = oldOff + colOffset
                    out = out & "C[" & CStr(newOff) & "]"
                    i = bracketEnd + 1
                    GoTo ContinueLoop
                End If
            End If
        End If
        ' Handle plain "C" not followed by "[" or digit (relative current-col ref)
        ' In R1C1, bare "C" in "RC" means current column (offset 0)
        ' We leave absolute Cn (digit follows) untouched
        out = out & Mid$(result, i, 1)
        i = i + 1
ContinueLoop:
    Loop
    ShiftR1C1RelativeByOffset = out
End Function

'=============================================================================
' SANDWICH DETECTION (A1 formula space)
' Returns True when:
'   (a) Formula contains a colon-range (not just individual refs)
'   (b) All A1 cell refs are on formulaRow (no cross-row refs)
'   (c) resultCol is strictly BETWEEN min and max referenced col
'   AND the range includes at least one INDIRECT("RC[n]") endpoint
'      or a direct A1 range like H14:Q14
'=============================================================================
Private Function IsSandwichFormulaA1(ByVal fA1 As String, _
                                     ByVal resultCol As Long, _
                                     ByVal formulaRow As Long) As Boolean
    IsSandwichFormulaA1 = False
    If Len(fA1) = 0 Or Left$(fA1, 1) <> "=" Then Exit Function

    ' Gate: must contain a colon range (A1:B1 or A1:INDIRECT(...))
    If InStr(fA1, ":") = 0 Then Exit Function

    ' Check for A1:A1 style range
    Dim hasDirectRange As Boolean
    hasDirectRange = HasColonRangeA1(fA1)
    ' Check for INDIRECT range endpoint
    Dim hasIndirectRange As Boolean
    hasIndirectRange = (InStr(1, fA1, "INDIRECT", vbTextCompare) > 0 And InStr(fA1, ":") > 0)

    If Not hasDirectRange And Not hasIndirectRange Then Exit Function

    ' Collect all A1 cell refs and check they're all on formulaRow
    Dim minCol As Long: minCol = 2147483647
    Dim maxCol As Long: maxCol = 0
    Dim foundAny As Boolean: foundAny = False

    ' Parse all A1-style refs: optional $, col letters, optional $, row digits
    Dim pos As Long: pos = 1
    Dim fUpper As String: fUpper = UCase$(fA1)
    Dim fLen As Long: fLen = Len(fA1)

    Do While pos <= fLen
        ' Skip string literals
        If Mid$(fA1, pos, 1) = """" Then
            pos = InStr(pos + 1, fA1, """")
            If pos = 0 Then Exit Do
            pos = pos + 1
            GoTo NextCharSandwich
        End If

        ' Try to match column letters at pos
        Dim colStart As Long: colStart = pos
        If Mid$(fUpper, pos, 1) = "$" Then colStart = pos + 1

        Dim colLetters As String: colLetters = ""
        Dim k As Long: k = colStart
        Do While k <= fLen
            Dim ch2 As String: ch2 = Mid$(fUpper, k, 1)
            If ch2 >= "A" And ch2 <= "Z" Then
                colLetters = colLetters & ch2
                k = k + 1
            Else
                Exit Do
            End If
        Loop

        If Len(colLetters) >= 1 And Len(colLetters) <= 3 Then
            ' Skip $ before row
            Dim rowStart As Long: rowStart = k
            If rowStart <= fLen And Mid$(fA1, rowStart, 1) = "$" Then rowStart = rowStart + 1
            ' Read row digits
            Dim rowDigits As String: rowDigits = ""
            Dim rk As Long: rk = rowStart
            Do While rk <= fLen
                Dim cd As String: cd = Mid$(fA1, rk, 1)
                If cd >= "0" And cd <= "9" Then
                    rowDigits = rowDigits & cd
                    rk = rk + 1
                Else
                    Exit Do
                End If
            Loop

            If Len(rowDigits) > 0 And IsNumeric(rowDigits) Then
                Dim refRow As Long: refRow = CLng(rowDigits)
                If refRow <> formulaRow Then
                    Exit Function   ' cross-row ref => not sandwich
                End If
                Dim refCol As Long: refCol = ColumnToNumber(colLetters)
                If refCol < minCol Then minCol = refCol
                If refCol > maxCol Then maxCol = refCol
                foundAny = True
                pos = rk
                GoTo NextCharSandwich
            End If
        End If

        pos = pos + 1
NextCharSandwich:
    Loop

    ' Resolve INDIRECT("RC[n]") endpoint if present
    If hasIndirectRange Then
        Dim indirPos As Long
        indirPos = InStr(1, fA1, "INDIRECT", vbTextCompare)
        Do While indirPos > 0
            Dim qOpen As Long: qOpen = InStr(indirPos, fA1, """")
            If qOpen = 0 Then Exit Do
            Dim qClose As Long: qClose = InStr(qOpen + 1, fA1, """")
            If qClose = 0 Then Exit Do
            Dim r1c1Str As String
            r1c1Str = UCase$(Mid$(fA1, qOpen + 1, qClose - qOpen - 1))
            ' Must be same-row: starts with "R" then "C" (no R offset)
            If Left$(r1c1Str, 2) = "RC" Then
                Dim cPart As String: cPart = Mid$(r1c1Str, 3)
                Dim indirColNum As Long
                If Len(cPart) = 0 Then
                    indirColNum = resultCol
                ElseIf Left$(cPart, 1) = "[" Then
                    Dim cbEnd As Long: cbEnd = InStr(cPart, "]")
                    If cbEnd > 0 Then
                        Dim offStr As String: offStr = Mid$(cPart, 2, cbEnd - 2)
                        If IsNumeric(offStr) Then
                            indirColNum = resultCol + CLng(offStr)
                        End If
                    End If
                ElseIf IsNumeric(cPart) Then
                    indirColNum = CLng(cPart)
                End If
                If indirColNum > 0 Then
                    foundAny = True
                    If indirColNum < minCol Then minCol = indirColNum
                    If indirColNum > maxCol Then maxCol = indirColNum
                End If
            Else
                Exit Function   ' R offset in INDIRECT => cross-row => not sandwich
            End If
            indirPos = InStr(qClose + 1, fA1, "INDIRECT", vbTextCompare)
        Loop
    End If

    If Not foundAny Then Exit Function
    IsSandwichFormulaA1 = (resultCol > minCol And resultCol < maxCol)
End Function

Private Function HasColonRangeA1(ByVal fA1 As String) As Boolean
    ' Quick check: does formula contain XN:XN pattern (direct range)?
    Dim p As Long: p = InStr(fA1, ":")
    Do While p > 0
        ' Check char before ":" is digit or $ (end of A1 ref)
        If p > 1 Then
            Dim cb As String: cb = Mid$(fA1, p - 1, 1)
            If cb >= "0" And cb <= "9" Then
                HasColonRangeA1 = True
                Exit Function
            End If
        End If
        p = InStr(p + 1, fA1, ":")
    Loop
    HasColonRangeA1 = False
End Function

'=============================================================================
' FAST external formula detection (no RegExp) — from reference implementation
'=============================================================================
Private Function IsExternalFormulaFast_R1C1(ByVal f As String, ByVal thisSheetName As String) As Boolean
    Dim p As Long, q1 As Long, q0 As Long
    Dim token As String, nm As String
    Dim sn As String: sn = LCase$(thisSheetName)
    Dim scanPos As Long: scanPos = 1

    Do
        p = InStr(scanPos, f, "!", vbTextCompare)
        If p = 0 Then Exit Do

        ' Workbook prefix [..]..! => external workbook
        If InStr(1, Left$(f, p), "[", vbTextCompare) > 0 And _
           InStr(1, Left$(f, p), "]", vbTextCompare) > 0 Then
            IsExternalFormulaFast_R1C1 = True
            Exit Function
        End If

        If p > 1 And Mid$(f, p - 1, 1) = "'" Then
            ' Quoted sheet: 'Sheet Name'!
            q1 = p - 1
            q0 = InStrRev(f, "'", q1 - 1)
            If q0 > 0 Then
                token = Mid$(f, q0 + 1, q1 - q0 - 1)
                nm = LCase$(Trim$(token))
                If nm <> sn Then
                    IsExternalFormulaFast_R1C1 = True
                    Exit Function
                End If
            End If
        Else
            ' Unquoted: SheetName!
            Dim k As Long: k = p - 1
            Do While k > 0
                Dim ch As String: ch = Mid$(f, k, 1)
                If (ch Like "[A-Za-z0-9_]") Or ch = "." Or ch = "-" Then
                    k = k - 1
                ElseIf ch = "]" Then
                    IsExternalFormulaFast_R1C1 = True
                    Exit Function
                Else
                    Exit Do
                End If
            Loop
            token = Mid$(f, k + 1, p - (k + 1))
            nm = LCase$(Trim$(token))
            If nm <> "" And nm <> sn Then
                IsExternalFormulaFast_R1C1 = True
                Exit Function
            End If
        End If

        scanPos = p + 1
    Loop

    IsExternalFormulaFast_R1C1 = False
End Function

'=============================================================================
' UNGROUPED: copy formulas as-is from srcCol to dstCol (array-based)
'=============================================================================
Private Sub CopyColumnFormulasAsIs(ByVal ws As Worksheet, _
                                   ByVal srcCol As Long, ByVal dstCol As Long, _
                                   ByVal lastRow As Long)
    Dim rngSrc As Range, rngDst As Range
    Dim arrVal As Variant, arrF As Variant, arrOut As Variant
    Dim i As Long, f As String

    Set rngSrc = ws.Range(ws.Cells(1, srcCol), ws.Cells(lastRow, srcCol))
    Set rngDst = ws.Range(ws.Cells(1, dstCol), ws.Cells(lastRow, dstCol))

    arrVal = rngSrc.Value2
    arrF   = rngSrc.FormulaR1C1
    arrOut = arrVal

    For i = 1 To UBound(arrF, 1)
        If VarType(arrF(i, 1)) = vbString Then
            f = CStr(arrF(i, 1))
            If Len(f) > 0 And Left$(f, 1) = "=" Then
                arrOut(i, 1) = f
            End If
        End If
    Next i

    rngDst.FormulaR1C1 = arrOut
    ' Copy number formats
    CopyColumnFormats ws, srcCol, dstCol
End Sub

'=============================================================================
' COMMENTS — rich-text preserved (from reference implementation)
'=============================================================================
Private Function CaptureComments(ByVal ws As Worksheet, ByVal colIdx As Long) As Object
    Dim col As Collection
    Set col = New Collection
    Dim cmt As Comment
    On Error Resume Next
    For Each cmt In ws.Comments
        If Not cmt Is Nothing Then
            If cmt.Parent.Column = colIdx Then
                Dim entry(2) As Variant
                entry(0) = cmt.Parent.Row
                entry(1) = cmt.Shape.TextFrame.Characters.Text
                entry(2) = cmt   ' keep reference for rich-run copy
                col.Add entry
            End If
        End If
    Next cmt
    On Error GoTo 0
    Set CaptureComments = col
End Function

Private Sub RestoreCommentsWithFormatting(ByVal ws As Worksheet, _
                                          ByVal colIdx As Long, _
                                          ByVal comments As Object)
    Dim entry As Variant
    Dim dstCell As Range
    On Error Resume Next
    For Each entry In comments
        Set dstCell = ws.Cells(entry(0), colIdx)
        If Not dstCell Is Nothing Then
            If Not dstCell.Comment Is Nothing Then dstCell.Comment.Delete
            CopyLegacyNoteWithFormatting entry(2), dstCell
        End If
    Next entry
    On Error GoTo 0
End Sub

Private Sub CopyLegacyNoteWithFormatting(ByVal srcComment As Comment, ByVal dstCell As Range)
    On Error Resume Next
    If Not dstCell.Comment Is Nothing Then dstCell.Comment.Delete
    On Error GoTo 0

    Dim srcTF As Object
    Set srcTF = srcComment.Shape.TextFrame
    Dim txt As String: txt = srcTF.Characters.Text

    On Error Resume Next
    dstCell.AddComment txt
    On Error GoTo 0
    If dstCell.Comment Is Nothing Then Exit Sub

    On Error Resume Next
    dstCell.Comment.Shape.Width  = srcComment.Shape.Width
    dstCell.Comment.Shape.Height = srcComment.Shape.Height
    On Error GoTo 0

    Dim dstTF As Object
    Set dstTF = dstCell.Comment.Shape.TextFrame

    ' Copy base font
    On Error Resume Next
    With dstTF.Characters.Font
        .Name  = srcTF.Characters.Font.Name
        .Size  = srcTF.Characters.Font.Size
        .Color = srcTF.Characters.Font.Color
    End With
    On Error GoTo 0

    ' Copy rich runs (bold/italic/underline) per character
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
                .Bold      = pb
                .Italic    = pit
                .Underline = pu
            End With
            runStart = i
            pb = b: pit = it: pu = u
        End If
    Next i

    runLen = n - runStart + 1
    With dstTF.Characters(runStart, runLen).Font
        .Bold      = pb
        .Italic    = pit
        .Underline = pu
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
' Current col  : delete shapes anchored to it
' Previous col : set FreeFloating (hard-coded position)
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
    Dim pos    As Double
    Dim c      As Long
    Dim maxCol As Long
    pos    = IIf(leftEdge, shp.Left, shp.Left + shp.Width)
    maxCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count
    For c = 1 To maxCol
        If ws.Cells(1, c).Left + ws.Cells(1, c).Width >= pos Then
            GetShapeColumn = c
            Exit Function
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
    If IsNumeric(colRef) Then
        ColumnToNumber = CLng(colRef)
        Exit Function
    End If
    Dim s As String: s = UCase$(Trim$(Replace(CStr(colRef), "$", "")))
    Dim i As Long, ch As Integer, n As Long
    n = 0
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
