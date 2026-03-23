Option Explicit

' =============================================================================
' roll_v3_optimised.bas
' Optimisations applied vs original:
'   1. RollOneSheetOneColumn     - ColumnNumberToLetters hoisted out of row loop
'   2. BuildFormulaLevelsForSheet - UsedRange.Formula read in ONE array call
'   3. ExtractExplicitRefsInternal - INDIRECT / OFFSET / full col-row paths
'                                    removed; explicit A1 refs only collected
'   4. ~20 helper functions deleted (only reachable via removed INDIRECT/OFFSET
'      paths): TryConsumeIndirectOrIgnore, TryConsumeOffsetOrIgnore,
'      TryConsumeFullColumnOrRowRef, TryParseRowRangeToken, TryReadFunctionArgs,
'      TryGetStandaloneQuotedLiteral, ReadDoubleQuotedLiteral,
'      TryParseIndirectA1Mode, TryParseR1C1RefOrRangeText,
'      TryParseQualifiedR1C1Ref, TryParseR1C1Single, TryParseSimpleRefOrRangeArg,
'      TryParseQualifiedCellRef, TryParseLongLiteral, FindTopLevelColon,
'      FindLastBangOutsideQuotes, MatchFunctionNameAt, A1ToRowCol, AddRefKey,
'      RowColToA1
' =============================================================================


' -----------------------------------------------------------------------------
' PUBLIC ENTRY POINT
' -----------------------------------------------------------------------------
Public Sub RollReportingPacks_FromControl()

    Dim ctlWs As Worksheet
    Dim lastCtlRow As Long
    Dim r As Long
    Dim startTick As Double
    Dim secs As Double

    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldStatusBar As Variant

    On Error GoTo CleanFail

    startTick = Timer
    Set ctlWs = ActiveSheet

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    lastCtlRow = LastUsedRowAny(ctlWs)
    If lastCtlRow < 2 Then GoTo CleanExit

    For r = 2 To lastCtlRow
        Application.StatusBar = "Rolling row " & r & " of " & lastCtlRow & "..."
        ProcessOneControlRow _
            targetSheetName:=Trim$(CStr(ctlWs.Cells(r, 1).Value)), _
            colSpec:=Trim$(CStr(ctlWs.Cells(r, 2).Value)), _
            methodText:=Trim$(CStr(ctlWs.Cells(r, 3).Value)), _
            freezeMaxText:=Trim$(CStr(ctlWs.Cells(r, 4).Value)), _
            directionText:=Trim$(CStr(ctlWs.Cells(r, 5).Value))
    Next r

CleanExit:
    secs = ElapsedSeconds(startTick)

    Application.StatusBar = oldStatusBar
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc

    MsgBox "Elapsed time: " & Format(secs, "0.00") & " seconds", vbInformation
    Exit Sub

CleanFail:
    secs = ElapsedSeconds(startTick)

    Application.StatusBar = oldStatusBar
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc

    MsgBox "Elapsed time: " & Format(secs, "0.00") & " seconds", vbExclamation
End Sub


' -----------------------------------------------------------------------------
' CONTROL ROW DISPATCH
' -----------------------------------------------------------------------------
Private Sub ProcessOneControlRow( _
    ByVal targetSheetName As String, _
    ByVal colSpec As String, _
    ByVal methodText As String, _
    ByVal freezeMaxText As String, _
    ByVal directionText As String)

    Dim ws As Worksheet
    Dim srcCol As Long
    Dim freezeMaxLevel As Long
    Dim isReverse As Boolean

    On Error GoTo SafeExit

    If Len(targetSheetName) = 0 Then Exit Sub
    If Len(colSpec) = 0 Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(targetSheetName)
    On Error GoTo SafeExit
    If ws Is Nothing Then Exit Sub

    srcCol = ParseColumnSpec(colSpec)
    If srcCol < 1 Or srcCol > 16384 Then Exit Sub

    freezeMaxLevel = ParseFreezeMaxLevel(freezeMaxText)
    isReverse = (UCase$(Trim$(directionText)) = "REVERSE")

    RollOneSheetOneColumn ws, srcCol, methodText, freezeMaxLevel, isReverse

SafeExit:
End Sub


' -----------------------------------------------------------------------------
' CORE ROLL — OPTIMISED
' Change vs original: srcColLetters computed ONCE before the For i loop
' instead of calling ColumnNumberToLetters(srcCol) on every row iteration.
' -----------------------------------------------------------------------------
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
    Dim srcColLetters As String          ' *** hoisted: computed once below ***

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

    ' *** Compute column letters once — reused on every row iteration ***
    srcColLetters = UCase$(ColumnNumberToLetters(srcCol))

    For i = 1 To lastRow
        isFormula = IsFormulaVariant(arrFormula(i, 1))
        addrOriginal = srcColLetters & CStr(i)   ' no function call inside loop

        If isFormula Then
            If levelMap.Exists(addrOriginal) Then
                lvl = CLng(levelMap(addrOriginal))
            Else
                lvl = 0
            End If

            If ShouldFreezeLevel(lvl, freezeMaxLevel) Then
                ws.Cells(i, workSrcCol).Value = arrValue(i, 1)
                ws.Cells(i, tgtCol).Formula = BuildFrozenTargetFormula( _
                    CStr(arrFormula(i, 1)), ws.Name, colDelta)
            Else
                ws.Cells(i, tgtCol).FormulaR1C1 = CStr(arrFormulaR1C1(i, 1))
            End If
        Else
            ws.Cells(i, tgtCol).Value = arrValue(i, 1)
        End If
    Next i

    ClearCommentsAndNotes ws.Range(ws.Cells(1, tgtCol), ws.Cells(lastRow, tgtCol))
End Sub


' -----------------------------------------------------------------------------
' COLUMN LAYOUT COPY
' -----------------------------------------------------------------------------
Private Sub CopyColumnLayoutNoComments( _
    ByVal ws As Worksheet, _
    ByVal srcCol As Long, _
    ByVal tgtCol As Long, _
    ByVal lastRow As Long)

    On Error Resume Next

    ws.Columns(srcCol).Copy
    ws.Columns(tgtCol).PasteSpecial xlPasteFormats
    ws.Columns(tgtCol).PasteSpecial xlPasteValidation
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


' -----------------------------------------------------------------------------
' FROZEN FORMULA BUILDER
' (unchanged — still needs TryParseColumnRangeToken, TryParseCellToken,
'  GetQualifierSheetName, Shift* helpers, SkipDoubleQuotedString)
' -----------------------------------------------------------------------------
Private Function BuildFrozenTargetFormula( _
    ByVal formulaText As String, _
    ByVal hostSheetName As String, _
    ByVal colDelta As Long) As String

    Dim n As Long
    Dim i As Long
    Dim lastEmit As Long
    Dim outText As String
    Dim tokenStart As Long
    Dim tokenEnd As Long
    Dim addrNorm As String
    Dim rawTok As String
    Dim shName As String
    Dim nextPos As Long
    Dim ch As String
    Dim normText As String

    n = Len(formulaText)
    i = 1
    lastEmit = 1

    Do While i <= n

        If TryParseColumnRangeToken(formulaText, i, tokenStart, tokenEnd, normText) Then
            shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
            outText = outText & Mid$(formulaText, lastEmit, i - lastEmit)
            rawTok = Mid$(formulaText, tokenStart, tokenEnd - tokenStart + 1)

            If StrComp(shName, UCase$(hostSheetName), vbTextCompare) = 0 Then
                outText = outText & ShiftColumnRangeTokenHorizontallyPreserveDollar(rawTok, colDelta)
            Else
                outText = outText & rawTok
            End If

            i = tokenEnd + 1
            lastEmit = i
            GoTo ContinueLoop
        End If

        If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then
            nextPos = tokenEnd + 1
            Do While nextPos <= n And Mid$(formulaText, nextPos, 1) = " "
                nextPos = nextPos + 1
            Loop

            If nextPos <= n And Mid$(formulaText, nextPos, 1) = "!" Then
                i = i + 1
                GoTo ContinueLoop
            End If

            shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
            outText = outText & Mid$(formulaText, lastEmit, i - lastEmit)
            rawTok = Mid$(formulaText, tokenStart, tokenEnd - tokenStart + 1)

            If StrComp(shName, UCase$(hostSheetName), vbTextCompare) = 0 Then
                outText = outText & ShiftA1TokenHorizontallyPreserveDollar(rawTok, colDelta)
            Else
                outText = outText & rawTok
            End If

            i = tokenEnd + 1
            lastEmit = i
            GoTo ContinueLoop
        End If

        ch = Mid$(formulaText, i, 1)
        If ch = """" Then
            i = SkipDoubleQuotedString(formulaText, i)
        Else
            i = i + 1
        End If

ContinueLoop:
    Loop

    outText = outText & Mid$(formulaText, lastEmit)
    BuildFrozenTargetFormula = outText
End Function


' -----------------------------------------------------------------------------
' TOKEN SHIFT HELPERS
' -----------------------------------------------------------------------------
Private Function ShiftA1TokenHorizontallyPreserveDollar( _
    ByVal rawToken As String, _
    ByVal colDelta As Long) As String

    Dim p As Long
    Dim n As Long
    Dim colAbs As Boolean
    Dim rowAbs As Boolean
    Dim colLetters As String
    Dim rowDigits As String
    Dim newCol As Long

    n = Len(rawToken)
    p = 1

    If p <= n And Mid$(rawToken, p, 1) = "$" Then
        colAbs = True
        p = p + 1
    End If

    Do While p <= n And IsLetterAZ(Mid$(rawToken, p, 1))
        colLetters = colLetters & Mid$(rawToken, p, 1)
        p = p + 1
    Loop

    If p <= n And Mid$(rawToken, p, 1) = "$" Then
        rowAbs = True
        p = p + 1
    End If

    Do While p <= n And IsDigit09(Mid$(rawToken, p, 1))
        rowDigits = rowDigits & Mid$(rawToken, p, 1)
        p = p + 1
    Loop

    If Len(colLetters) = 0 Or Len(rowDigits) = 0 Or p <= n Then
        ShiftA1TokenHorizontallyPreserveDollar = rawToken
        Exit Function
    End If

    If colAbs Then
        ShiftA1TokenHorizontallyPreserveDollar = rawToken
        Exit Function
    End If

    newCol = ColLettersToNumber2(UCase$(colLetters)) + colDelta
    If newCol < 1 Or newCol > 16384 Then
        ShiftA1TokenHorizontallyPreserveDollar = rawToken
        Exit Function
    End If

    ShiftA1TokenHorizontallyPreserveDollar = ColumnNumberToLetters(newCol)
    If rowAbs Then ShiftA1TokenHorizontallyPreserveDollar = _
        ShiftA1TokenHorizontallyPreserveDollar & "$"
    ShiftA1TokenHorizontallyPreserveDollar = _
        ShiftA1TokenHorizontallyPreserveDollar & rowDigits
End Function


Private Function ShiftColumnRangeTokenHorizontallyPreserveDollar( _
    ByVal rawToken As String, _
    ByVal colDelta As Long) As String

    Dim parts() As String

    If InStr(1, rawToken, ":", vbBinaryCompare) = 0 Then
        ShiftColumnRangeTokenHorizontallyPreserveDollar = rawToken
        Exit Function
    End If

    parts = Split(rawToken, ":")
    If UBound(parts) <> 1 Then
        ShiftColumnRangeTokenHorizontallyPreserveDollar = rawToken
        Exit Function
    End If

    ShiftColumnRangeTokenHorizontallyPreserveDollar = _
        ShiftSingleColumnTokenPreserveDollar(parts(0), colDelta) & ":" & _
        ShiftSingleColumnTokenPreserveDollar(parts(1), colDelta)
End Function


Private Function ShiftSingleColumnTokenPreserveDollar( _
    ByVal rawToken As String, _
    ByVal colDelta As Long) As String

    Dim s As String
    Dim absCol As Boolean
    Dim colLetters As String
    Dim p As Long
    Dim newCol As Long

    s = rawToken
    p = 1

    If Len(s) = 0 Then
        ShiftSingleColumnTokenPreserveDollar = s
        Exit Function
    End If

    If Mid$(s, p, 1) = "$" Then
        absCol = True
        p = p + 1
    End If

    Do While p <= Len(s) And IsLetterAZ(Mid$(s, p, 1))
        colLetters = colLetters & Mid$(s, p, 1)
        p = p + 1
    Loop

    If Len(colLetters) = 0 Or p <= Len(s) Then
        ShiftSingleColumnTokenPreserveDollar = s
        Exit Function
    End If

    If absCol Then
        ShiftSingleColumnTokenPreserveDollar = s
        Exit Function
    End If

    newCol = ColLettersToNumber2(UCase$(colLetters)) + colDelta
    If newCol < 1 Or newCol > 16384 Then
        ShiftSingleColumnTokenPreserveDollar = s
        Exit Function
    End If

    ShiftSingleColumnTokenPreserveDollar = ColumnNumberToLetters(newCol)
End Function


' -----------------------------------------------------------------------------
' FREEZE LEVEL LOGIC
' -----------------------------------------------------------------------------
Private Function ShouldFreezeLevel( _
    ByVal lvl As Long, _
    ByVal freezeMaxLevel As Long) As Boolean

    If freezeMaxLevel <= 0 Then Exit Function
    If lvl <= 0 Then Exit Function

    If freezeMaxLevel >= 2 Then
        ShouldFreezeLevel = (lvl = 1 Or lvl = 2)
    Else
        ShouldFreezeLevel = (lvl = 1)
    End If
End Function


' -----------------------------------------------------------------------------
' COLUMN / ROW UTILITIES
' -----------------------------------------------------------------------------
Private Sub PrepareExistingTargetColumn( _
    ByVal ws As Worksheet, _
    ByVal tgtCol As Long)

    On Error Resume Next
    ws.Columns(tgtCol).Hidden = False
    ws.Columns(tgtCol).Ungroup
    On Error GoTo 0
End Sub


Private Function ParseFreezeMaxLevel(ByVal txt As String) As Long
    Dim s As String

    s = Trim$(txt)

    If Len(s) = 0 Then
        ParseFreezeMaxLevel = 2
    ElseIf IsNumeric(s) Then
        ParseFreezeMaxLevel = CLng(s)
        If ParseFreezeMaxLevel < 0 Then ParseFreezeMaxLevel = 0
        If ParseFreezeMaxLevel > 2 Then ParseFreezeMaxLevel = 2
    Else
        ParseFreezeMaxLevel = 2
    End If
End Function


Private Function ParseColumnSpec(ByVal colSpec As String) As Long
    Dim s As String
    s = Trim$(colSpec)

    If Len(s) = 0 Then
        ParseColumnSpec = 0
        Exit Function
    End If

    If IsNumeric(s) Then
        ParseColumnSpec = CLng(s)
    Else
        ParseColumnSpec = ColumnLettersToNumber(UCase$(s))
    End If
End Function


Private Function ColumnLettersToNumber(ByVal letters As String) As Long
    Dim i As Long
    Dim v As Long
    Dim ch As Integer

    letters = UCase$(Trim$(letters))
    If Len(letters) = 0 Then Exit Function

    For i = 1 To Len(letters)
        ch = Asc(Mid$(letters, i, 1))
        If ch < 65 Or ch > 90 Then
            ColumnLettersToNumber = 0
            Exit Function
        End If
        v = v * 26 + (ch - 64)
    Next i

    ColumnLettersToNumber = v
End Function


Private Function ColumnNumberToLetters(ByVal colNum As Long) As String
    Dim n As Long
    Dim s As String

    n = colNum
    Do While n > 0
        s = Chr$(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop

    ColumnNumberToLetters = s
End Function


' -----------------------------------------------------------------------------
' FORMULA LEVEL MAP — OPTIMISED
' Change vs original: entire UsedRange.Formula read in ONE array call instead
' of one c.Formula COM call per formula cell in the second For Each pass.
' For a sheet with N formula cells this reduces COM calls N+1 -> 2.
' -----------------------------------------------------------------------------
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

    ' *** bulk formula read ***
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

    ' *** Read all formulas from the sheet in ONE COM call ***
    Set usedRng = ws.UsedRange
    urRow = usedRng.Row
    urCol = usedRng.Column

    If usedRng.Cells.Count = 1 Then
        ' Single-cell UsedRange returns a scalar — normalise to 2-D array
        ReDim formulaBase(1 To 1, 1 To 1)
        formulaBase(1, 1) = usedRng.Formula
    Else
        formulaBase = usedRng.Formula   ' 2-D variant array, 1-based rows/cols
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

    ' Pass 1: index every formula cell address
    i = 0
    For Each c In formulaCells.Cells
        i = i + 1
        nodeAddr = UCase$(c.Address(False, False))
        addrArr(i) = nodeAddr
        If Not dictIndex.Exists(nodeAddr) Then
            dictIndex.Add nodeAddr, i
        End If
    Next c

    ' Pass 2: build dependency edges — formula read from pre-loaded array
    For Each c In formulaCells.Cells
        idx = CLng(dictIndex(UCase$(c.Address(False, False))))

        ' *** Array index replaces c.Formula COM call ***
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

    ' BFS propagation
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


' -----------------------------------------------------------------------------
' EXPLICIT REF EXTRACTION — OPTIMISED (explicit A1 refs only)
'
' INDIRECT(), OFFSET(), and full column/row range tokens (e.g. A:A, 1:1)
' are completely ignored. Only literal cell addresses present in the formula
' text are collected.
'
' Example:  =ROUND(Q75-SUM(G75:OFFSET(O75,0,-1)),2)
'   Collected : Q75, G75, O75
'   Ignored   : OFFSET() call and its computed range result
' -----------------------------------------------------------------------------
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

    ' failOnUnresolvedDynamic kept for signature compatibility; now a no-op.

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

        ' Skip string literals — prevents "A1" inside text being treated as a ref
        If ch = """" Then
            i = SkipDoubleQuotedString(formulaText, i)

        ' Try to match an A1 cell token at position i
        ElseIf TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then

            ' If followed by "!" this token is a sheet-name qualifier, not a ref.
            ' Skip it and let the scanner advance to the actual address after "!".
            nextPos = tokenEnd + 1
            Do While nextPos <= n And Mid$(formulaText, nextPos, 1) = " "
                nextPos = nextPos + 1
            Loop

            If nextPos <= n And Mid$(formulaText, nextPos, 1) = "!" Then
                i = tokenEnd + 1    ' skip qualifier, real address comes next
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


' -----------------------------------------------------------------------------
' TOKEN PARSERS (used by BuildFrozenTargetFormula and ExtractExplicitRefsInternal)
' -----------------------------------------------------------------------------
Private Function TryParseColumnRangeToken( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByRef tokenStart As Long, _
    ByRef tokenEnd As Long, _
    ByRef normText As String) As Boolean

    Dim n As Long
    Dim p As Long
    Dim ch As String
    Dim start1 As Long, start2 As Long
    Dim col1 As String, col2 As String

    TryParseColumnRangeToken = False
    normText = vbNullString

    n = Len(s)
    If pos < 1 Or pos > n Then Exit Function

    If pos > 1 Then
        ch = Mid$(s, pos - 1, 1)
        If IsAlphaNumUnderscoreDot(ch) Or ch = "[" Then Exit Function
    End If

    p = pos
    If Mid$(s, p, 1) = "$" Then p = p + 1

    start1 = p
    Do While p <= n And IsLetterAZ(Mid$(s, p, 1))
        p = p + 1
    Loop
    If p = start1 Then Exit Function
    If (p - start1) > 3 Then Exit Function
    col1 = UCase$(Mid$(s, start1, p - start1))

    If p > n Or Mid$(s, p, 1) <> ":" Then Exit Function
    p = p + 1

    If p <= n And Mid$(s, p, 1) = "$" Then p = p + 1

    start2 = p
    Do While p <= n And IsLetterAZ(Mid$(s, p, 1))
        p = p + 1
    Loop
    If p = start2 Then Exit Function
    If (p - start2) > 3 Then Exit Function
    col2 = UCase$(Mid$(s, start2, p - start2))

    If ColLettersToNumber2(col1) < 1 Or ColLettersToNumber2(col1) > 16384 Then Exit Function
    If ColLettersToNumber2(col2) < 1 Or ColLettersToNumber2(col2) > 16384 Then Exit Function

    If p <= n Then
        ch = Mid$(s, p, 1)
        If IsAlphaNumUnderscoreDot(ch) Then Exit Function
    End If

    tokenStart = pos
    tokenEnd = p - 1
    normText = col1 & ":" & col2
    TryParseColumnRangeToken = True
End Function


Private Function TryParseCellToken( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByRef tokenStart As Long, _
    ByRef tokenEnd As Long, _
    ByRef addrNorm As String) As Boolean

    Dim n As Long
    Dim p As Long
    Dim ch As String
    Dim colStart As Long
    Dim rowStart As Long

    TryParseCellToken = False
    addrNorm = vbNullString

    n = Len(s)
    If pos < 1 Or pos > n Then Exit Function

    If pos > 1 Then
        ch = Mid$(s, pos - 1, 1)
        If IsAlphaNumUnderscoreDot(ch) Or ch = "[" Then Exit Function
    End If

    p = pos

    If Mid$(s, p, 1) = "$" Then p = p + 1
    If p > n Then Exit Function

    colStart = p
    Do While p <= n And IsLetterAZ(Mid$(s, p, 1))
        p = p + 1
    Loop
    If p = colStart Then Exit Function
    If (p - colStart) > 3 Then Exit Function

    If p <= n Then
        If Mid$(s, p, 1) = "$" Then p = p + 1
    End If

    rowStart = p
    Do While p <= n And IsDigit09(Mid$(s, p, 1))
        p = p + 1
    Loop
    If p = rowStart Then Exit Function

    If p <= n Then
        ch = Mid$(s, p, 1)
        If IsAlphaNumUnderscoreDot(ch) Then Exit Function
    End If

    addrNorm = NormalizeA1(Mid$(s, pos, p - pos))
    If LenB(addrNorm) = 0 Then Exit Function

    tokenStart = pos
    tokenEnd = p - 1
    TryParseCellToken = True
End Function


' -----------------------------------------------------------------------------
' SHEET QUALIFIER HELPERS
' -----------------------------------------------------------------------------
Private Function GetQualifierSheetName( _
    ByVal formulaText As String, _
    ByVal tokenStart As Long, _
    ByVal hostSheetName As String) As String

    Dim exPos As Long
    Dim endPos As Long
    Dim startPos As Long
    Dim raw As String
    Dim ch As String

    exPos = tokenStart - 1
    Do While exPos >= 1 And Mid$(formulaText, exPos, 1) = " "
        exPos = exPos - 1
    Loop

    If exPos < 1 Or Mid$(formulaText, exPos, 1) <> "!" Then
        GetQualifierSheetName = UCase$(hostSheetName)
        Exit Function
    End If

    endPos = exPos - 1
    If endPos < 1 Then
        GetQualifierSheetName = UCase$(hostSheetName)
        Exit Function
    End If

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

        If startPos < 1 Then
            GetQualifierSheetName = UCase$(hostSheetName)
            Exit Function
        End If

        raw = Mid$(formulaText, startPos, endPos - startPos + 1)
    Else
        startPos = endPos
        Do While startPos >= 1
            ch = Mid$(formulaText, startPos, 1)
            If IsSheetQualifierChar(ch) Then
                startPos = startPos - 1
            Else
                Exit Do
            End If
        Loop
        raw = Mid$(formulaText, startPos + 1, endPos - startPos)
    End If

    raw = CleanSheetQualifier(raw)
    If LenB(raw) = 0 Then raw = hostSheetName

    GetQualifierSheetName = UCase$(raw)
End Function


Private Function CleanSheetQualifier(ByVal raw As String) As String
    raw = Trim$(raw)

    If Len(raw) = 0 Then
        CleanSheetQualifier = vbNullString
        Exit Function
    End If

    If Left$(raw, 1) = "'" And Right$(raw, 1) = "'" Then
        raw = Mid$(raw, 2, Len(raw) - 2)
        raw = Replace(raw, "''", "'")
    End If

    If InStrRev(raw, "]") > 0 Then
        raw = Mid$(raw, InStrRev(raw, "]") + 1)
    End If

    CleanSheetQualifier = raw
End Function


' -----------------------------------------------------------------------------
' ADDRESS NORMALISATION
' -----------------------------------------------------------------------------
Private Function NormalizeA1(ByVal raw As String) As String
    Dim s As String
    Dim i As Long
    Dim colPart As String
    Dim rowPart As String
    Dim colNum As Long
    Dim rowNum As Double

    NormalizeA1 = vbNullString

    s = UCase$(Replace(raw, "$", ""))
    If Len(s) = 0 Then Exit Function

    i = 1
    Do While i <= Len(s) And IsLetterAZ(Mid$(s, i, 1))
        colPart = colPart & Mid$(s, i, 1)
        i = i + 1
    Loop

    Do While i <= Len(s) And IsDigit09(Mid$(s, i, 1))
        rowPart = rowPart & Mid$(s, i, 1)
        i = i + 1
    Loop

    If i <= Len(s) Then Exit Function
    If Len(colPart) = 0 Or Len(rowPart) = 0 Then Exit Function
    If Len(colPart) > 3 Then Exit Function

    colNum = ColLettersToNumber2(colPart)
    If colNum < 1 Or colNum > 16384 Then Exit Function

    rowNum = CDbl(rowPart)
    If rowNum < 1 Or rowNum > 1048576 Then Exit Function

    NormalizeA1 = colPart & CLng(rowNum)
End Function


' -----------------------------------------------------------------------------
' STRING / QUOTE HELPERS
' -----------------------------------------------------------------------------
Private Function SkipDoubleQuotedString( _
    ByVal s As String, _
    ByVal pos As Long) As Long

    Dim n As Long
    n = Len(s)
    pos = pos + 1

    Do While pos <= n
        If Mid$(s, pos, 1) = """" Then
            If pos < n And Mid$(s, pos + 1, 1) = """" Then
                pos = pos + 2
            Else
                SkipDoubleQuotedString = pos + 1
                Exit Function
            End If
        Else
            pos = pos + 1
        End If
    Loop

    SkipDoubleQuotedString = pos
End Function


' -----------------------------------------------------------------------------
' COLUMN NUMBER HELPERS
' -----------------------------------------------------------------------------
Private Function ColLettersToNumber2(ByVal colLetters As String) As Long
    Dim i As Long
    Dim v As Long

    For i = 1 To Len(colLetters)
        v = v * 26 + (Asc(Mid$(colLetters, i, 1)) - 64)
    Next i

    ColLettersToNumber2 = v
End Function


' -----------------------------------------------------------------------------
' CHARACTER CLASS HELPERS
' -----------------------------------------------------------------------------
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
    IsAlphaNumUnderscoreDot = _
        IsLetterAZ(ch) Or _
        IsDigit09(ch) Or _
        (ch = "_") Or _
        (ch = ".")
End Function


Private Function IsSheetQualifierChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsSheetQualifierChar = _
        IsLetterAZ(ch) Or _
        IsDigit09(ch) Or _
        (ch = "_") Or _
        (ch = ".") Or _
        (ch = "[") Or _
        (ch = "]")
End Function


' -----------------------------------------------------------------------------
' FORMULA / CELL DETECTION
' -----------------------------------------------------------------------------
Private Function IsFormulaVariant(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then
        If Len(v) > 0 Then
            IsFormulaVariant = (Left$(CStr(v), 1) = "=")
        End If
    End If
End Function


' -----------------------------------------------------------------------------
' WORKSHEET UTILITIES
' -----------------------------------------------------------------------------
Private Function LastUsedRowAny(ByVal ws As Worksheet) As Long
    Dim f As Range

    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious)
    On Error GoTo 0

    If f Is Nothing Then
        LastUsedRowAny = 1
    Else
        LastUsedRowAny = f.Row
    End If
End Function


Private Function ElapsedSeconds(ByVal startTick As Double) As Double
    If Timer >= startTick Then
        ElapsedSeconds = Timer - startTick
    Else
        ElapsedSeconds = (86400# - startTick) + Timer
    End If
End Function
