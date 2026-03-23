' ============================================================
' REPLACEMENT FUNCTIONS — drop-in replacements for roll_1__2_.bas
' Three changes:
'   1. RollOneSheetOneColumn      — hoist ColumnNumberToLetters out of loop
'   2. BuildFormulaLevelsForSheet — read .Formula via array (one COM call)
'   3. ExtractExplicitRefsInternal — simplified: explicit cell refs only,
'                                    INDIRECT / OFFSET / full-col-row ignored
'
' DEAD CODE you can delete after applying these (no longer called):
'   TryConsumeIndirectOrIgnore, TryConsumeOffsetOrIgnore,
'   TryConsumeFullColumnOrRowRef, TryParseColumnRangeToken (ExtractRefs path only;
'   keep for BuildFrozenTargetFormula), TryParseRowRangeToken,
'   TryReadFunctionArgs, TryGetStandaloneQuotedLiteral, ReadDoubleQuotedLiteral,
'   TryParseIndirectA1Mode, TryParseR1C1RefOrRangeText, TryParseQualifiedR1C1Ref,
'   TryParseR1C1Single, TryParseSimpleRefOrRangeArg, TryParseQualifiedCellRef,
'   TryParseLongLiteral, FindTopLevelColon, FindLastBangOutsideQuotes,
'   MatchFunctionNameAt, A1ToRowCol, AddRefKey
' ============================================================


' ------------------------------------------------------------
' REPLACEMENT 1 of 3
' RollOneSheetOneColumn
' Change: srcColLetters computed ONCE before the loop instead
'         of calling ColumnNumberToLetters(srcCol) on every row.
' Replace the entire existing Private Sub RollOneSheetOneColumn with this.
' ------------------------------------------------------------
Private Sub RollOneSheetOneColumn( _
    ByVal ws As Worksheet, _
    ByVal srcCol As Long, _
    ByVal methodText As String, _
    ByVal freezeMaxLevel As Long, _
    ByVal isReverse As Boolean)

    Dim levelMap As Object
    Dim lastRow As Long
    Dim tgtCol As Long
    Dim workSrcCol As Long
    Dim methodUpper As String
    Dim colDelta As Long

    Dim srcRng As Range
    Dim arrFormula As Variant
    Dim arrFormulaR1C1 As Variant
    Dim arrValue As Variant

    Dim i As Long
    Dim addrOriginal As String
    Dim lvl As Long
    Dim isFormula As Boolean

    ' *** NEW: compute once, reuse for every row ***
    Dim srcColLetters As String

    methodUpper = UCase$(Trim$(methodText))
    If Len(methodUpper) = 0 Then methodUpper = "INSERT"

    If isReverse Then
        If srcCol = 1 Then Exit Sub
        colDelta = -1
    Else
        If srcCol = 16384 Then Exit Sub
        colDelta = 1
    End If

    Set levelMap = BuildFormulaLevelsForSheet(ws)

    lastRow = LastUsedRowAny(ws)
    If lastRow < 1 Then lastRow = 1

    Set srcRng = ws.Range(ws.Cells(1, srcCol), ws.Cells(lastRow, srcCol))

    arrFormula = srcRng.Formula
    arrFormulaR1C1 = srcRng.FormulaR1C1
    arrValue = srcRng.Value

    If Not isReverse Then
        workSrcCol = srcCol
        tgtCol = srcCol + 1

        Select Case methodUpper
            Case "INSERT", "I"
                ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Case "UNGROUP", "U"
                PrepareExistingTargetColumn ws, tgtCol
            Case Else
                ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End Select
    Else
        If methodUpper = "INSERT" Or methodUpper = "I" Then
            tgtCol = srcCol
            ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            workSrcCol = srcCol + 1
        Else
            workSrcCol = srcCol
            tgtCol = srcCol - 1
            PrepareExistingTargetColumn ws, tgtCol
        End If
    End If

    CopyColumnLayoutNoComments ws, workSrcCol, tgtCol, lastRow

    ' *** Hoist: compute column letters once before the loop ***
    srcColLetters = UCase$(ColumnNumberToLetters(srcCol))

    For i = 1 To lastRow
        isFormula = IsFormulaVariant(arrFormula(i, 1))

        ' Was: UCase$(ColumnNumberToLetters(srcCol) & CStr(i))
        addrOriginal = srcColLetters & CStr(i)

        If isFormula Then
            If levelMap.Exists(addrOriginal) Then
                lvl = CLng(levelMap(addrOriginal))
            Else
                lvl = 0
            End If

            If ShouldFreezeLevel(lvl, freezeMaxLevel) Then
                ws.Cells(i, workSrcCol).Value = arrValue(i, 1)
                ws.Cells(i, tgtCol).Formula = BuildFrozenTargetFormula(CStr(arrFormula(i, 1)), ws.Name, colDelta)
            Else
                ws.Cells(i, tgtCol).FormulaR1C1 = CStr(arrFormulaR1C1(i, 1))
            End If
        Else
            ws.Cells(i, tgtCol).Value = arrValue(i, 1)
        End If
    Next i

    ClearCommentsAndNotes ws.Range(ws.Cells(1, tgtCol), ws.Cells(lastRow, tgtCol))
End Sub


' ------------------------------------------------------------
' REPLACEMENT 2 of 3
' BuildFormulaLevelsForSheet
' Change: UsedRange.Formula read into a 2-D array ONCE instead
'         of calling c.Formula on every formula cell individually.
'         For a sheet with 500 formula cells that is 500 → 1 COM call.
' Replace the entire existing Private Function BuildFormulaLevelsForSheet.
' ------------------------------------------------------------
Private Function BuildFormulaLevelsForSheet(ByVal ws As Worksheet) As Object

    Dim result As Object
    Dim formulaCells As Range
    Dim c As Range

    Dim dictIndex As Object
    Dim parents() As Collection
    Dim addrArr() As String
    Dim hasDirectOther() As Boolean
    Dim levelCap() As Long

    Dim refs As Object
    Dim n As Long
    Dim i As Long
    Dim idx As Long
    Dim nodeAddr As String

    Dim key As Variant
    Dim refKey As String
    Dim barPos As Long
    Dim shName As String
    Dim addr As String
    Dim childIdx As Long

    Dim q() As Long
    Dim head As Long
    Dim tail As Long
    Dim parentIdx As Long

    ' *** NEW: bulk formula array ***
    Dim usedRng As Range
    Dim formulaBase As Variant
    Dim urRow As Long
    Dim urCol As Long
    Dim cellFormula As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaCells Is Nothing Then
        Set BuildFormulaLevelsForSheet = result
        Exit Function
    End If

    n = CLng(formulaCells.CountLarge)
    If n = 0 Then
        Set BuildFormulaLevelsForSheet = result
        Exit Function
    End If

    ' *** Read ALL formulas from the sheet in ONE call ***
    Set usedRng = ws.UsedRange
    urRow = usedRng.Row
    urCol = usedRng.Column

    If usedRng.Cells.Count = 1 Then
        ' Single-cell UsedRange returns a scalar — normalise to 2-D array
        ReDim formulaBase(1 To 1, 1 To 1)
        formulaBase(1, 1) = usedRng.Formula
    Else
        formulaBase = usedRng.Formula   ' 2-D variant array, 1-based
    End If

    Set dictIndex = CreateObject("Scripting.Dictionary")
    dictIndex.CompareMode = vbTextCompare

    ReDim parents(1 To n)
    ReDim addrArr(1 To n)
    ReDim hasDirectOther(1 To n)
    ReDim levelCap(1 To n)

    For i = 1 To n
        Set parents(i) = New Collection
    Next i

    ' Pass 1: index every formula cell by its address
    i = 0
    For Each c In formulaCells.Cells
        i = i + 1
        nodeAddr = UCase$(c.Address(False, False))
        addrArr(i) = nodeAddr
        If Not dictIndex.Exists(nodeAddr) Then
            dictIndex.Add nodeAddr, i
        End If
    Next c

    ' Pass 2: build dependency edges using the pre-read array
    For Each c In formulaCells.Cells
        idx = CLng(dictIndex(UCase$(c.Address(False, False))))

        ' *** Read from array — no COM call ***
        cellFormula = CStr(formulaBase(c.Row - urRow + 1, c.Column - urCol + 1))

        Set refs = ExtractExplicitRefs(cellFormula, ws.Name, c.Row, c.Column)

        If refs.Count > 0 Then
            For Each key In refs.Keys
                refKey = CStr(key)
                barPos = InStr(1, refKey, "|", vbBinaryCompare)
                shName = Left$(refKey, barPos - 1)
                addr = Mid$(refKey, barPos + 1)

                If StrComp(shName, UCase$(ws.Name), vbTextCompare) = 0 Then
                    If dictIndex.Exists(addr) Then
                        childIdx = CLng(dictIndex(addr))
                        parents(childIdx).Add idx
                    End If
                Else
                    hasDirectOther(idx) = True
                End If
            Next key
        End If
    Next c

    ' BFS to propagate levels
    ReDim q(1 To n)
    head = 1
    tail = 0

    For i = 1 To n
        If hasDirectOther(i) Then
            levelCap(i) = 1
            tail = tail + 1
            q(tail) = i
        End If
    Next i

    Do While head <= tail
        childIdx = q(head)
        head = head + 1

        For Each key In parents(childIdx)
            parentIdx = CLng(key)
            If levelCap(parentIdx) = 0 Then
                levelCap(parentIdx) = 2
                tail = tail + 1
                q(tail) = parentIdx
            End If
        Next key
    Loop

    For i = 1 To n
        result(addrArr(i)) = levelCap(i)
    Next i

    Set BuildFormulaLevelsForSheet = result
End Function


' ------------------------------------------------------------
' REPLACEMENT 3 of 3
' ExtractExplicitRefsInternal  (and its public wrapper ExtractExplicitRefs)
'
' Change: INDIRECT, OFFSET, and full-column/row-range tokens are all
'         ignored entirely. The scanner only collects explicit A1 cell
'         addresses that appear literally in the formula text.
'
' Example: =ROUND(Q75-SUM(G75:OFFSET(O75,0,-1)),2)
'   Collected:  Q75, G75, O75
'   Ignored:    OFFSET() call and its computed result
'
' The failOnUnresolvedDynamic parameter is kept for signature compatibility
' but is now a no-op (there are no dynamic references to fail on).
'
' Replace both existing functions with these two.
' ------------------------------------------------------------
Private Function ExtractExplicitRefs( _
    ByVal formulaText As String, _
    ByVal hostSheetName As String, _
    Optional ByVal hostRow As Long = 1, _
    Optional ByVal hostCol As Long = 1) As Object

    Dim ok As Boolean
    Set ExtractExplicitRefs = ExtractExplicitRefsInternal( _
        formulaText, hostSheetName, hostRow, hostCol, False, ok)
End Function

Private Function ExtractExplicitRefsInternal( _
    ByVal formulaText As String, _
    ByVal hostSheetName As String, _
    ByVal hostRow As Long, _
    ByVal hostCol As Long, _
    ByVal failOnUnresolvedDynamic As Boolean, _
    ByRef parseOk As Boolean) As Object

    ' Scans formulaText for explicit A1 cell references only.
    ' INDIRECT(), OFFSET(), and full column/row range tokens (e.g. A:A, 1:1)
    ' are all skipped — only literal addresses such as Q75 or $G$75 are recorded.

    Dim d As Object
    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim tokenStart As Long
    Dim tokenEnd As Long
    Dim addrNorm As String
    Dim shName As String
    Dim key As String
    Dim nextPos As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    parseOk = True

    n = Len(formulaText)
    i = 1

    Do While i <= n
        ch = Mid$(formulaText, i, 1)

        ' Skip quoted string literals so "A1" inside a string isn't treated as a ref
        If ch = """" Then
            i = SkipDoubleQuotedString(formulaText, i)

        ' Attempt to parse an A1-style cell token at this position
        ElseIf TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then

            ' If the token is immediately followed by "!" it is a sheet-name
            ' qualifier, not an address — skip it and let the next iteration
            ' pick up the real cell address that follows.
            nextPos = tokenEnd + 1
            Do While nextPos <= n And Mid$(formulaText, nextPos, 1) = " "
                nextPos = nextPos + 1
            Loop

            If nextPos <= n And Mid$(formulaText, nextPos, 1) = "!" Then
                i = tokenEnd + 1        ' skip the qualifier token
            Else
                shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
                key = UCase$(shName) & "|" & addrNorm
                If Not d.Exists(key) Then d.Add key, True
                i = tokenEnd + 1
            End If

        Else
            i = i + 1
        End If
    Loop

    Set ExtractExplicitRefsInternal = d
End Function


' ============================================================
' FUNCTIONS THAT ARE NOW DEAD CODE — safe to delete
' ============================================================
'
'  These were only reachable through ExtractExplicitRefsInternal's
'  INDIRECT / OFFSET / column-range analysis paths, which are removed above.
'  BuildFrozenTargetFormula still uses TryParseColumnRangeToken,
'  TryParseCellToken, GetQualifierSheetName — keep those.
'
'  DELETE list:
'    Private Function TryConsumeIndirectOrIgnore(...)
'    Private Function TryConsumeOffsetOrIgnore(...)
'    Private Function TryConsumeFullColumnOrRowRef(...)
'    Private Function TryParseRowRangeToken(...)
'    Private Function TryReadFunctionArgs(...)
'    Private Function TryGetStandaloneQuotedLiteral(...)
'    Private Function ReadDoubleQuotedLiteral(...)
'    Private Function TryParseIndirectA1Mode(...)
'    Private Function TryParseR1C1RefOrRangeText(...)
'    Private Function TryParseQualifiedR1C1Ref(...)
'    Private Function TryParseR1C1Single(...)
'    Private Function TryParseSimpleRefOrRangeArg(...)
'    Private Function TryParseQualifiedCellRef(...)
'    Private Function TryParseLongLiteral(...)
'    Private Function FindTopLevelColon(...)
'    Private Function FindLastBangOutsideQuotes(...)
'    Private Function MatchFunctionNameAt(...)
'    Private Sub      A1ToRowCol(...)
'    Private Sub      AddRefKey(...)
'
'  KEEP (still used by BuildFrozenTargetFormula or other live paths):
'    TryParseColumnRangeToken, TryParseCellToken, GetQualifierSheetName,
'    CleanSheetQualifier, NormalizeA1, ShiftA1TokenHorizontallyPreserveDollar,
'    ShiftColumnRangeTokenHorizontallyPreserveDollar,
'    ShiftSingleColumnTokenPreserveDollar, SkipDoubleQuotedString,
'    ColLettersToNumber2, ColumnLettersToNumber, ColumnNumberToLetters,
'    IsLetterAZ, IsDigit09, IsAlphaNumUnderscoreDot, IsSheetQualifierChar,
'    IsFormulaVariant, LastUsedRowAny, ElapsedSeconds, ShouldFreezeLevel,
'    CopyColumnLayoutNoComments, ClearCommentsAndNotes, PrepareExistingTargetColumn,
'    ParseFreezeMaxLevel, ParseColumnSpec, BuildFrozenTargetFormula,
'    RowColToA1 (used in BuildFormulaLevelsForSheet indirectly via ExtractExplicitRefs
'               but actually no longer needed — safe to delete too)
' ============================================================
