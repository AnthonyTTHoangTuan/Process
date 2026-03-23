Option Explicit

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
    Dim newFormula As String

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

    For i = 1 To lastRow
        isFormula = IsFormulaVariant(arrFormula(i, 1))
        addrOriginal = UCase$(ColumnNumberToLetters(srcCol) & CStr(i))

        If isFormula Then
            If levelMap.Exists(addrOriginal) Then
                lvl = CLng(levelMap(addrOriginal))
            Else
                lvl = 0
            End If

            If ShouldFreezeLevel(lvl, freezeMaxLevel) Then
                ws.Cells(i, workSrcCol).Value = arrValue(i, 1)

                If lvl = 3 And freezeMaxLevel >= 3 Then
                    newFormula = BuildLevel3TargetFormula(CStr(arrFormula(i, 1)), ws.Name, colDelta)
                Else
                    newFormula = BuildFrozenTargetFormula(CStr(arrFormula(i, 1)), ws.Name, colDelta)
                End If

                If Len(newFormula) = 0 Then
                    ws.Cells(i, tgtCol).ClearContents
                ElseIf Left$(newFormula, 1) = "=" Then
                    ws.Cells(i, tgtCol).Formula = newFormula
                Else
                    ws.Cells(i, tgtCol).Value = newFormula
                End If
            Else
                ws.Cells(i, tgtCol).FormulaR1C1 = CStr(arrFormulaR1C1(i, 1))
            End If
        Else
            ws.Cells(i, tgtCol).Value = arrValue(i, 1)
        End If
    Next i

    ClearCommentsAndNotes ws.Range(ws.Cells(1, tgtCol), ws.Cells(lastRow, tgtCol))
End Sub

Private Sub CopyColumnLayoutNoComments(ByVal ws As Worksheet, ByVal srcCol As Long, ByVal tgtCol As Long, ByVal lastRow As Long)
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

Private Function BuildFrozenTargetFormula(ByVal formulaText As String, ByVal hostSheetName As String, ByVal colDelta As Long) As String

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

Private Function ShiftA1TokenHorizontallyPreserveDollar(ByVal rawToken As String, ByVal colDelta As Long) As String

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
    If rowAbs Then ShiftA1TokenHorizontallyPreserveDollar = ShiftA1TokenHorizontallyPreserveDollar & "$"
    ShiftA1TokenHorizontallyPreserveDollar = ShiftA1TokenHorizontallyPreserveDollar & rowDigits
End Function

Private Function ShiftColumnRangeTokenHorizontallyPreserveDollar(ByVal rawToken As String, ByVal colDelta As Long) As String
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

Private Function ShiftSingleColumnTokenPreserveDollar(ByVal rawToken As String, ByVal colDelta As Long) As String
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

Private Function ShouldFreezeLevel(ByVal lvl As Long, ByVal freezeMaxLevel As Long) As Boolean
    If freezeMaxLevel <= 0 Then Exit Function
    If lvl <= 0 Then Exit Function

    Select Case freezeMaxLevel
        Case Is >= 3
            ShouldFreezeLevel = (lvl = 1 Or lvl = 2 Or lvl = 3)
        Case 2
            ShouldFreezeLevel = (lvl = 1 Or lvl = 2)
        Case Else
            ShouldFreezeLevel = (lvl = 1)
    End Select
End Function

Private Sub PrepareExistingTargetColumn(ByVal ws As Worksheet, ByVal tgtCol As Long)
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
        If ParseFreezeMaxLevel > 3 Then ParseFreezeMaxLevel = 3
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

Private Function IsFormulaVariant(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then
        If Len(v) > 0 Then
            IsFormulaVariant = (Left$(CStr(v), 1) = "=")
        End If
    End If
End Function

Private Function LastUsedRowAny(ByVal ws As Worksheet) As Long
    Dim f As Range

    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
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

    Set dictIndex = CreateObject("Scripting.Dictionary")
    dictIndex.CompareMode = vbTextCompare

    ReDim parents(1 To n)
    ReDim addrArr(1 To n)
    ReDim hasDirectOther(1 To n)
    ReDim levelCap(1 To n)

    For i = 1 To n
        Set parents(i) = New Collection
    Next i

    i = 0
    For Each c In formulaCells.Cells
        i = i + 1
        nodeAddr = UCase$(c.Address(False, False))
        addrArr(i) = nodeAddr
        If Not dictIndex.Exists(nodeAddr) Then
            dictIndex.Add nodeAddr, i
        End If
    Next c

    For Each c In formulaCells.Cells
        idx = CLng(dictIndex(UCase$(c.Address(False, False))))
        Set refs = ExtractExplicitRefs(c.Formula, ws.Name, c.Row, c.Column)

        If refs.Count > 0 Then
            For Each key In refs.Keys
                refKey = CStr(key)
                barPos = InStr(1, refKey, "|", vbBinaryCompare)
                shName = Left$(refKey, barPos - 1)
                addr = Mid$(refKey, barPos + 1)

                If StrComp(shName, UCase$(ws.Name), vbTextCompare) = 0 Then
                    If Left$(addr, 1) <> "#" Then
                        If dictIndex.Exists(addr) Then
                            childIdx = CLng(dictIndex(addr))
                            parents(childIdx).Add idx
                        End If
                    End If
                Else
                    hasDirectOther(idx) = True
                End If
            Next key
        End If
    Next c

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

    i = 0
    For Each c In formulaCells.Cells
        i = i + 1
        If levelCap(i) = 0 Then
            If HasHardcodedNumberLevel3(CStr(c.Formula)) Then
                levelCap(i) = 3
            End If
        End If
    Next c

    For i = 1 To n
        result(addrArr(i)) = levelCap(i)
    Next i

    Set BuildFormulaLevelsForSheet = result
End Function

Private Function ExtractExplicitRefs( _
    ByVal formulaText As String, _
    ByVal hostSheetName As String, _
    Optional ByVal hostRow As Long = 1, _
    Optional ByVal hostCol As Long = 1) As Object

    Dim ok As Boolean
    Set ExtractExplicitRefs = ExtractExplicitRefsInternal(formulaText, hostSheetName, hostRow, hostCol, False, ok)
End Function

Private Function ExtractExplicitRefsInternal( _
    ByVal formulaText As String, _
    ByVal hostSheetName As String, _
    ByVal hostRow As Long, _
    ByVal hostCol As Long, _
    ByVal failOnUnresolvedDynamic As Boolean, _
    ByRef parseOk As Boolean) As Object

    Dim d As Object
    Dim i As Long
    Dim tokenStart As Long
    Dim tokenEnd As Long
    Dim addrNorm As String
    Dim shName As String
    Dim key As String
    Dim n As Long
    Dim ch As String
    Dim nextPos As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    parseOk = True
    n = Len(formulaText)
    i = 1

    Do While i <= n

        If TryConsumeIndirectVisibleLiteralRef(formulaText, i, hostSheetName, d, tokenEnd) Then
            i = tokenEnd + 1
            GoTo ContinueLoop
        End If

        If TryConsumeOffsetVisibleBaseRef(formulaText, i, hostSheetName, d, tokenEnd) Then
            i = tokenEnd + 1
            GoTo ContinueLoop
        End If

        If TryConsumeFullColumnOrRowRef(formulaText, i, hostSheetName, d, tokenEnd) Then
            i = tokenEnd + 1
            GoTo ContinueLoop
        End If

        ch = Mid$(formulaText, i, 1)

        If ch = """" Then
            i = SkipDoubleQuotedString(formulaText, i)
        Else
            If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then

                nextPos = tokenEnd + 1
                Do While nextPos <= n And Mid$(formulaText, nextPos, 1) = " "
                    nextPos = nextPos + 1
                Loop

                If nextPos <= n And Mid$(formulaText, nextPos, 1) = "!" Then
                    i = tokenEnd + 1
                Else
                    shName = GetQualifierSheetName(formulaText, tokenStart, hostSheetName)
                    key = UCase$(shName) & "|" & addrNorm
                    If Not d.Exists(key) Then d.Add key, True
                    i = tokenEnd + 1
                End If

            Else
                i = i + 1
            End If
        End If

ContinueLoop:
    Loop

    Set ExtractExplicitRefsInternal = d
End Function

Private Function TryConsumeIndirectOrIgnore( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByVal hostSheetName As String, _
    ByVal hostRow As Long, _
    ByVal hostCol As Long, _
    ByRef outDict As Object, _
    ByRef endPos As Long) As Long

    Dim openParenPos As Long
    Dim closePos As Long
    Dim args() As String
    Dim lit As String
    Dim subRefs As Object
    Dim k As Variant
    Dim litOk As Boolean
    Dim isA1Mode As Boolean

    Dim shName As String
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long

    TryConsumeIndirectOrIgnore = 0
    endPos = pos

    If Not MatchFunctionNameAt(s, pos, "INDIRECT", openParenPos) Then Exit Function

    If Not TryReadFunctionArgs(s, openParenPos, args, closePos) Then
        endPos = openParenPos
        TryConsumeIndirectOrIgnore = 2
        Exit Function
    End If

    endPos = closePos

    If UBound(args) < 1 Then
        TryConsumeIndirectOrIgnore = 2
        Exit Function
    End If

    If Not TryGetStandaloneQuotedLiteral(args(1), lit) Then
        TryConsumeIndirectOrIgnore = 2
        Exit Function
    End If

    isA1Mode = True
    If UBound(args) >= 2 Then
        If Not TryParseIndirectA1Mode(args(2), isA1Mode) Then
            TryConsumeIndirectOrIgnore = 2
            Exit Function
        End If
    End If

    If isA1Mode Then
        Set subRefs = ExtractExplicitRefsInternal(lit, hostSheetName, hostRow, hostCol, True, litOk)

        If (Not litOk) Or subRefs.Count = 0 Then
            TryConsumeIndirectOrIgnore = 2
            Exit Function
        End If

        For Each k In subRefs.Keys
            If Not outDict.Exists(CStr(k)) Then outDict.Add CStr(k), True
        Next k

        TryConsumeIndirectOrIgnore = 1
        Exit Function
    End If

    If Not TryParseR1C1RefOrRangeText(lit, hostSheetName, hostRow, hostCol, shName, r1, c1, r2, c2) Then
        TryConsumeIndirectOrIgnore = 2
        Exit Function
    End If

    AddRefKey outDict, shName, RowColToA1(r1, c1)
    AddRefKey outDict, shName, RowColToA1(r2, c2)

    TryConsumeIndirectOrIgnore = 1
End Function

Private Function TryConsumeOffsetOrIgnore( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByVal hostSheetName As String, _
    ByRef outDict As Object, _
    ByRef endPos As Long) As Long

    Dim openParenPos As Long
    Dim closePos As Long
    Dim args() As String
    Dim argCount As Long

    Dim shName As String
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Dim rowsOff As Long, colsOff As Long
    Dim h As Long, w As Long

    Dim topRow As Long, botRow As Long
    Dim leftCol As Long, rightCol As Long
    Dim startRow As Long, startCol As Long
    Dim endRow As Long, endCol As Long

    TryConsumeOffsetOrIgnore = 0
    endPos = pos

    If Not MatchFunctionNameAt(s, pos, "OFFSET", openParenPos) Then Exit Function

    If Not TryReadFunctionArgs(s, openParenPos, args, closePos) Then
        endPos = openParenPos
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If

    endPos = closePos
    argCount = UBound(args)

    If argCount < 3 Or argCount > 5 Then
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If

    If Not TryParseSimpleRefOrRangeArg(args(1), hostSheetName, shName, r1, c1, r2, c2) Then
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If

    If Not TryParseLongLiteral(args(2), rowsOff, False, True) Then
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If

    If Not TryParseLongLiteral(args(3), colsOff, False, True) Then
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If

    If r1 <= r2 Then
        topRow = r1
        botRow = r2
    Else
        topRow = r2
        botRow = r1
    End If

    If c1 <= c2 Then
        leftCol = c1
        rightCol = c2
    Else
        leftCol = c2
        rightCol = c1
    End If

    startRow = topRow + rowsOff
    startCol = leftCol + colsOff

    If argCount = 3 Then
        h = botRow - topRow + 1
        w = rightCol - leftCol + 1
    ElseIf argCount = 4 Then
        If Not TryParseLongLiteral(args(4), h, True, False) Then
            TryConsumeOffsetOrIgnore = 2
            Exit Function
        End If
        w = rightCol - leftCol + 1
    Else
        If Not TryParseLongLiteral(args(4), h, True, False) Then
            TryConsumeOffsetOrIgnore = 2
            Exit Function
        End If
        If Not TryParseLongLiteral(args(5), w, True, False) Then
            TryConsumeOffsetOrIgnore = 2
            Exit Function
        End If
    End If

    endRow = startRow + h - 1
    endCol = startCol + w - 1

    If startRow < 1 Or startRow > 1048576 Then
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If
    If startCol < 1 Or startCol > 16384 Then
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If
    If endRow < 1 Or endRow > 1048576 Then
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If
    If endCol < 1 Or endCol > 16384 Then
        TryConsumeOffsetOrIgnore = 2
        Exit Function
    End If

    AddRefKey outDict, shName, RowColToA1(startRow, startCol)
    AddRefKey outDict, shName, RowColToA1(endRow, endCol)

    TryConsumeOffsetOrIgnore = 1
End Function

Private Function TryConsumeFullColumnOrRowRef( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByVal hostSheetName As String, _
    ByRef outDict As Object, _
    ByRef endPos As Long) As Boolean

    Dim tokenStart As Long
    Dim tokenEnd As Long
    Dim normText As String
    Dim shName As String

    TryConsumeFullColumnOrRowRef = False
    endPos = pos

    If TryParseColumnRangeToken(s, pos, tokenStart, tokenEnd, normText) Then
        shName = GetQualifierSheetName(s, tokenStart, hostSheetName)
        AddRefKey outDict, shName, "#COL#" & normText
        endPos = tokenEnd
        TryConsumeFullColumnOrRowRef = True
        Exit Function
    End If

    If TryParseRowRangeToken(s, pos, tokenStart, tokenEnd, normText) Then
        shName = GetQualifierSheetName(s, tokenStart, hostSheetName)
        AddRefKey outDict, shName, "#ROW#" & normText
        endPos = tokenEnd
        TryConsumeFullColumnOrRowRef = True
        Exit Function
    End If
End Function

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

Private Function TryParseRowRangeToken( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByRef tokenStart As Long, _
    ByRef tokenEnd As Long, _
    ByRef normText As String) As Boolean

    Dim n As Long
    Dim p As Long
    Dim ch As String
    Dim start1 As Long, start2 As Long
    Dim row1Text As String, row2Text As String
    Dim row1 As Double, row2 As Double

    TryParseRowRangeToken = False
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
    Do While p <= n And IsDigit09(Mid$(s, p, 1))
        p = p + 1
    Loop
    If p = start1 Then Exit Function
    row1Text = Mid$(s, start1, p - start1)

    If p > n Or Mid$(s, p, 1) <> ":" Then Exit Function
    p = p + 1

    If p <= n And Mid$(s, p, 1) = "$" Then p = p + 1

    start2 = p
    Do While p <= n And IsDigit09(Mid$(s, p, 1))
        p = p + 1
    Loop
    If p = start2 Then Exit Function
    row2Text = Mid$(s, start2, p - start2)

    row1 = CDbl(row1Text)
    row2 = CDbl(row2Text)
    If row1 < 1 Or row1 > 1048576 Then Exit Function
    If row2 < 1 Or row2 > 1048576 Then Exit Function

    If p <= n Then
        ch = Mid$(s, p, 1)
        If IsAlphaNumUnderscoreDot(ch) Then Exit Function
    End If

    tokenStart = pos
    tokenEnd = p - 1
    normText = CLng(row1) & ":" & CLng(row2)
    TryParseRowRangeToken = True
End Function

Private Function TryParseIndirectA1Mode(ByVal argText As String, ByRef isA1Mode As Boolean) As Boolean
    Dim s As String

    TryParseIndirectA1Mode = False
    s = UCase$(Replace(Trim$(argText), " ", ""))

    Select Case s
        Case "TRUE", "1"
            isA1Mode = True
            TryParseIndirectA1Mode = True
        Case "FALSE", "0"
            isA1Mode = False
            TryParseIndirectA1Mode = True
    End Select
End Function

Private Function TryParseR1C1RefOrRangeText( _
    ByVal expr As String, _
    ByVal defaultSheetName As String, _
    ByVal hostRow As Long, _
    ByVal hostCol As Long, _
    ByRef outSheet As String, _
    ByRef r1 As Long, ByRef c1 As Long, _
    ByRef r2 As Long, ByRef c2 As Long) As Boolean

    Dim s As String
    Dim colonPos As Long
    Dim leftPart As String, rightPart As String
    Dim sh1 As String, sh2 As String
    Dim hasSh1 As Boolean, hasSh2 As Boolean

    TryParseR1C1RefOrRangeText = False

    s = Trim$(expr)
    If Len(s) = 0 Then Exit Function

    colonPos = FindTopLevelColon(s)

    If colonPos = 0 Then
        If Not TryParseQualifiedR1C1Ref(s, defaultSheetName, hostRow, hostCol, sh1, r1, c1, hasSh1) Then Exit Function
        outSheet = sh1
        r2 = r1
        c2 = c1
        TryParseR1C1RefOrRangeText = True
        Exit Function
    End If

    leftPart = Trim$(Left$(s, colonPos - 1))
    rightPart = Trim$(Mid$(s, colonPos + 1))

    If Not TryParseQualifiedR1C1Ref(leftPart, defaultSheetName, hostRow, hostCol, sh1, r1, c1, hasSh1) Then Exit Function
    If Not TryParseQualifiedR1C1Ref(rightPart, IIf(hasSh1, sh1, defaultSheetName), hostRow, hostCol, sh2, r2, c2, hasSh2) Then Exit Function
    If StrComp(sh1, sh2, vbTextCompare) <> 0 Then Exit Function

    outSheet = sh1
    TryParseR1C1RefOrRangeText = True
End Function

Private Function TryParseQualifiedR1C1Ref( _
    ByVal s As String, _
    ByVal defaultSheetName As String, _
    ByVal hostRow As Long, _
    ByVal hostCol As Long, _
    ByRef outSheet As String, _
    ByRef outRow As Long, _
    ByRef outCol As Long, _
    ByRef hasExplicitSheet As Boolean) As Boolean

    Dim t As String
    Dim bangPos As Long
    Dim qual As String
    Dim refText As String

    TryParseQualifiedR1C1Ref = False
    hasExplicitSheet = False
    outSheet = UCase$(defaultSheetName)

    t = Trim$(s)
    If Len(t) = 0 Then Exit Function

    bangPos = FindLastBangOutsideQuotes(t)

    If bangPos > 0 Then
        qual = Trim$(Left$(t, bangPos - 1))
        refText = Trim$(Mid$(t, bangPos + 1))
        outSheet = UCase$(CleanSheetQualifier(qual))
        hasExplicitSheet = True
    Else
        refText = t
    End If

    If Not TryParseR1C1Single(refText, hostRow, hostCol, outRow, outCol) Then Exit Function
    TryParseQualifiedR1C1Ref = True
End Function

Private Function TryParseR1C1Single( _
    ByVal s As String, _
    ByVal hostRow As Long, _
    ByVal hostCol As Long, _
    ByRef outRow As Long, _
    ByRef outCol As Long) As Boolean

    Dim t As String
    Dim p As Long
    Dim n As Long
    Dim startPos As Long
    Dim txt As String
    Dim ch As String
    Dim v As Long

    TryParseR1C1Single = False

    t = UCase$(Replace(Trim$(s), " ", ""))
    n = Len(t)
    If n = 0 Then Exit Function

    p = 1
    If Mid$(t, p, 1) <> "R" Then Exit Function
    p = p + 1

    If p <= n And Mid$(t, p, 1) = "[" Then
        p = p + 1
        startPos = p

        If p <= n Then
            ch = Mid$(t, p, 1)
            If ch = "+" Or ch = "-" Then p = p + 1
        End If

        Do While p <= n And IsDigit09(Mid$(t, p, 1))
            p = p + 1
        Loop

        If p = startPos Then Exit Function
        If p = startPos + 1 Then
            ch = Mid$(t, startPos, 1)
            If ch = "+" Or ch = "-" Then Exit Function
        End If

        If p > n Or Mid$(t, p, 1) <> "]" Then Exit Function

        txt = Mid$(t, startPos, p - startPos)
        v = CLng(txt)
        outRow = hostRow + v
        p = p + 1
    Else
        startPos = p
        Do While p <= n And IsDigit09(Mid$(t, p, 1))
            p = p + 1
        Loop

        If p > startPos Then
            outRow = CLng(Mid$(t, startPos, p - startPos))
        Else
            outRow = hostRow
        End If
    End If

    If p > n Or Mid$(t, p, 1) <> "C" Then Exit Function
    p = p + 1

    If p <= n And Mid$(t, p, 1) = "[" Then
        p = p + 1
        startPos = p

        If p <= n Then
            ch = Mid$(t, p, 1)
            If ch = "+" Or ch = "-" Then p = p + 1
        End If

        Do While p <= n And IsDigit09(Mid$(t, p, 1))
            p = p + 1
        Loop

        If p = startPos Then Exit Function
        If p = startPos + 1 Then
            ch = Mid$(t, startPos, 1)
            If ch = "+" Or ch = "-" Then Exit Function
        End If

        If p > n Or Mid$(t, p, 1) <> "]" Then Exit Function

        txt = Mid$(t, startPos, p - startPos)
        v = CLng(txt)
        outCol = hostCol + v
        p = p + 1
    Else
        startPos = p
        Do While p <= n And IsDigit09(Mid$(t, p, 1))
            p = p + 1
        Loop

        If p > startPos Then
            outCol = CLng(Mid$(t, startPos, p - startPos))
        Else
            outCol = hostCol
        End If
    End If

    If p <= n Then Exit Function
    If outRow < 1 Or outRow > 1048576 Then Exit Function
    If outCol < 1 Or outCol > 16384 Then Exit Function

    TryParseR1C1Single = True
End Function

Private Function MatchFunctionNameAt( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByVal funcName As String, _
    ByRef openParenPos As Long) As Boolean

    Dim n As Long
    Dim p As Long
    Dim ch As String
    Dim L As Long

    MatchFunctionNameAt = False
    openParenPos = 0

    n = Len(s)
    L = Len(funcName)

    If pos < 1 Or pos + L - 1 > n Then Exit Function
    If UCase$(Mid$(s, pos, L)) <> UCase$(funcName) Then Exit Function

    If pos > 1 Then
        ch = Mid$(s, pos - 1, 1)
        If IsAlphaNumUnderscoreDot(ch) Or ch = "[" Then Exit Function
    End If

    If pos + L <= n Then
        ch = Mid$(s, pos + L, 1)
        If IsAlphaNumUnderscoreDot(ch) Then Exit Function
    End If

    p = pos + L
    Do While p <= n And Mid$(s, p, 1) = " "
        p = p + 1
    Loop

    If p > n Or Mid$(s, p, 1) <> "(" Then Exit Function

    openParenPos = p
    MatchFunctionNameAt = True
End Function

Private Function TryReadFunctionArgs( _
    ByVal s As String, _
    ByVal openParenPos As Long, _
    ByRef args() As String, _
    ByRef closeParenPos As Long) As Boolean

    Dim p As Long
    Dim n As Long
    Dim depth As Long
    Dim argStart As Long
    Dim argCount As Long
    Dim ch As String

    TryReadFunctionArgs = False
    closeParenPos = 0

    If openParenPos < 1 Or openParenPos > Len(s) Then Exit Function
    If Mid$(s, openParenPos, 1) <> "(" Then Exit Function

    n = Len(s)
    depth = 1
    argStart = openParenPos + 1
    p = openParenPos + 1
    argCount = 0

    Do While p <= n
        ch = Mid$(s, p, 1)

        If ch = """" Then
            p = SkipDoubleQuotedString(s, p)
        ElseIf ch = "(" Then
            depth = depth + 1
            p = p + 1
        ElseIf ch = ")" Then
            depth = depth - 1
            If depth = 0 Then
                argCount = argCount + 1
                ReDim Preserve args(1 To argCount)
                args(argCount) = Trim$(Mid$(s, argStart, p - argStart))
                closeParenPos = p
                TryReadFunctionArgs = True
                Exit Function
            Else
                p = p + 1
            End If
        ElseIf ch = "," And depth = 1 Then
            argCount = argCount + 1
            ReDim Preserve args(1 To argCount)
            args(argCount) = Trim$(Mid$(s, argStart, p - argStart))
            argStart = p + 1
            p = p + 1
        Else
            p = p + 1
        End If
    Loop
End Function

Private Function TryGetStandaloneQuotedLiteral(ByVal expr As String, ByRef outText As String) As Boolean
    Dim s As String
    Dim qEnd As Long

    TryGetStandaloneQuotedLiteral = False
    outText = vbNullString

    s = Trim$(expr)
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> """" Then Exit Function

    outText = ReadDoubleQuotedLiteral(s, 1, qEnd)
    If qEnd = 0 Then Exit Function
    If qEnd <> Len(s) Then Exit Function

    TryGetStandaloneQuotedLiteral = True
End Function

Private Function ReadDoubleQuotedLiteral( _
    ByVal s As String, _
    ByVal startQuotePos As Long, _
    ByRef endQuotePos As Long) As String

    Dim p As Long
    Dim n As Long
    Dim outText As String

    endQuotePos = 0
    ReadDoubleQuotedLiteral = vbNullString

    If startQuotePos < 1 Or startQuotePos > Len(s) Then Exit Function
    If Mid$(s, startQuotePos, 1) <> """" Then Exit Function

    n = Len(s)
    p = startQuotePos + 1

    Do While p <= n
        If Mid$(s, p, 1) = """" Then
            If p < n And Mid$(s, p + 1, 1) = """" Then
                outText = outText & """"
                p = p + 2
            Else
                endQuotePos = p
                ReadDoubleQuotedLiteral = outText
                Exit Function
            End If
        Else
            outText = outText & Mid$(s, p, 1)
            p = p + 1
        End If
    Loop
End Function

Private Function SkipDoubleQuotedString(ByVal s As String, ByVal pos As Long) As Long
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

Private Function TryParseSimpleRefOrRangeArg( _
    ByVal argText As String, _
    ByVal hostSheetName As String, _
    ByRef outSheet As String, _
    ByRef r1 As Long, ByRef c1 As Long, _
    ByRef r2 As Long, ByRef c2 As Long) As Boolean

    Dim s As String
    Dim colonPos As Long
    Dim leftPart As String
    Dim rightPart As String
    Dim addr1 As String
    Dim addr2 As String
    Dim sh1 As String
    Dim sh2 As String
    Dim hasSh1 As Boolean
    Dim hasSh2 As Boolean

    TryParseSimpleRefOrRangeArg = False

    s = Trim$(argText)
    If Len(s) = 0 Then Exit Function

    colonPos = FindTopLevelColon(s)

    If colonPos = 0 Then
        If Not TryParseQualifiedCellRef(s, hostSheetName, sh1, addr1, hasSh1) Then Exit Function
        outSheet = sh1
        A1ToRowCol addr1, r1, c1
        r2 = r1
        c2 = c1
        TryParseSimpleRefOrRangeArg = True
        Exit Function
    End If

    leftPart = Trim$(Left$(s, colonPos - 1))
    rightPart = Trim$(Mid$(s, colonPos + 1))

    If Not TryParseQualifiedCellRef(leftPart, hostSheetName, sh1, addr1, hasSh1) Then Exit Function
    If Not TryParseQualifiedCellRef(rightPart, IIf(hasSh1, sh1, hostSheetName), sh2, addr2, hasSh2) Then Exit Function
    If StrComp(sh1, sh2, vbTextCompare) <> 0 Then Exit Function

    outSheet = sh1
    A1ToRowCol addr1, r1, c1
    A1ToRowCol addr2, r2, c2

    TryParseSimpleRefOrRangeArg = True
End Function

Private Function TryParseQualifiedCellRef( _
    ByVal s As String, _
    ByVal defaultSheetName As String, _
    ByRef outSheet As String, _
    ByRef outAddr As String, _
    ByRef hasExplicitSheet As Boolean) As Boolean

    Dim t As String
    Dim bangPos As Long
    Dim qual As String
    Dim addrText As String

    TryParseQualifiedCellRef = False
    hasExplicitSheet = False
    outSheet = UCase$(defaultSheetName)
    outAddr = vbNullString

    t = Trim$(s)
    If Len(t) = 0 Then Exit Function

    bangPos = FindLastBangOutsideQuotes(t)

    If bangPos > 0 Then
        qual = Trim$(Left$(t, bangPos - 1))
        addrText = Trim$(Mid$(t, bangPos + 1))
        outSheet = UCase$(CleanSheetQualifier(qual))
        hasExplicitSheet = True
    Else
        addrText = t
    End If

    outAddr = NormalizeA1(addrText)
    If LenB(outAddr) = 0 Then Exit Function

    TryParseQualifiedCellRef = True
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

Private Function GetQualifierSheetName(ByVal formulaText As String, ByVal tokenStart As Long, ByVal hostSheetName As String) As String
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

Private Function TryParseLongLiteral( _
    ByVal expr As String, _
    ByRef outValue As Long, _
    ByVal mustBePositive As Boolean, _
    ByVal allowZero As Boolean) As Boolean

    Dim s As String
    Dim i As Long
    Dim ch As String

    TryParseLongLiteral = False
    outValue = 0

    s = Replace(expr, " ", "")
    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If i = 1 And (ch = "+" Or ch = "-") Then
        ElseIf Not IsDigit09(ch) Then
            Exit Function
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
    Dim i As Long
    Dim depth As Long
    Dim ch As String

    FindTopLevelColon = 0
    depth = 0
    i = 1

    Do While i <= Len(s)
        ch = Mid$(s, i, 1)

        If ch = """" Then
            i = SkipDoubleQuotedString(s, i)
        ElseIf ch = "(" Then
            depth = depth + 1
            i = i + 1
        ElseIf ch = ")" Then
            If depth > 0 Then depth = depth - 1
            i = i + 1
        ElseIf ch = ":" And depth = 0 Then
            FindTopLevelColon = i
            Exit Function
        Else
            i = i + 1
        End If
    Loop
End Function

Private Function FindLastBangOutsideQuotes(ByVal s As String) As Long
    Dim i As Long
    Dim inQuote As Boolean
    Dim ch As String

    FindLastBangOutsideQuotes = 0
    inQuote = False

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch = "'" Then
            If i < Len(s) And Mid$(s, i + 1, 1) = "'" Then
                i = i + 1
            Else
                inQuote = Not inQuote
            End If
        ElseIf ch = "!" And Not inQuote Then
            FindLastBangOutsideQuotes = i
        End If
    Next i
End Function

Private Sub A1ToRowCol(ByVal addr As String, ByRef outRow As Long, ByRef outCol As Long)
    Dim i As Long
    Dim colPart As String
    Dim rowPart As String

    i = 1
    Do While i <= Len(addr) And IsLetterAZ(Mid$(addr, i, 1))
        colPart = colPart & Mid$(addr, i, 1)
        i = i + 1
    Loop

    Do While i <= Len(addr) And IsDigit09(Mid$(addr, i, 1))
        rowPart = rowPart & Mid$(addr, i, 1)
        i = i + 1
    Loop

    outCol = ColLettersToNumber2(colPart)
    outRow = CLng(rowPart)
End Sub

Private Function RowColToA1(ByVal rowNum As Long, ByVal colNum As Long) As String
    Dim n As Long
    Dim s As String

    n = colNum
    Do While n > 0
        s = Chr$(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop

    RowColToA1 = s & CStr(rowNum)
End Function

Private Function BuildLevel3TargetFormula(ByVal formulaText As String, ByVal hostSheetName As String, ByVal colDelta As Long) As String
    Dim shifted As String
    Dim cleaned As String

    shifted = BuildFrozenTargetFormula(formulaText, hostSheetName, colDelta)
    cleaned = StripHardcodedNumericAdjustments(shifted)

    If Len(cleaned) = 0 Or cleaned = "=" Then
        BuildLevel3TargetFormula = vbNullString
    ElseIf IsPureNumericExpression(RemoveLeadingEquals(cleaned)) Then
        BuildLevel3TargetFormula = vbNullString
    Else
        BuildLevel3TargetFormula = cleaned
    End If
End Function

Private Function HasHardcodedNumberLevel3(ByVal formulaText As String) As Boolean
    Dim body As String

    body = RemoveLeadingEquals(formulaText)
    body = Trim$(body)
    If Len(body) = 0 Then Exit Function

    If IsPureNumericExpression(body) Then
        HasHardcodedNumberLevel3 = True
        Exit Function
    End If

    If FormulaHasSignedNumericTerm(body) Then
        HasHardcodedNumberLevel3 = True
        Exit Function
    End If

    If FormulaHasNumericArgInFunction(body, "SUM") Then
        HasHardcodedNumberLevel3 = True
        Exit Function
    End If
End Function

Private Function StripHardcodedNumericAdjustments(ByVal formulaText As String) As String
    Dim body As String

    body = RemoveLeadingEquals(formulaText)
    body = Trim$(body)
    If Len(body) = 0 Then Exit Function

    If IsPureNumericExpression(body) Then
        StripHardcodedNumericAdjustments = vbNullString
        Exit Function
    End If

    body = RemoveNumericArgumentsFromNamedFunctions(body, "SUM")
    body = RemoveSignedNumericTerms(body)
    body = CleanupFormulaBody(body)

    If Len(body) = 0 Then
        StripHardcodedNumericAdjustments = vbNullString
    ElseIf IsPureNumericExpression(body) Then
        StripHardcodedNumericAdjustments = vbNullString
    Else
        StripHardcodedNumericAdjustments = "=" & body
    End If
End Function

Private Function RemoveLeadingEquals(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) > 0 Then
        If Left$(s, 1) = "=" Then
            RemoveLeadingEquals = Mid$(s, 2)
            Exit Function
        End If
    End If
    RemoveLeadingEquals = s
End Function

Private Function FormulaHasSignedNumericTerm(ByVal body As String) As Boolean
    Dim i As Long
    Dim prevCh As String
    Dim nextPos As Long
    Dim numEnd As Long

    i = 1
    Do While i <= Len(body)
        Select Case Mid$(body, i, 1)
            Case """"
                i = SkipDoubleQuotedString(body, i)
            Case "+", "-"
                prevCh = PreviousSignificantChar(body, i - 1)
                If prevCh <> "E" And prevCh <> "e" Then
                    nextPos = SkipSpacesLocal(body, i + 1)
                    If ReadNumberLiteral(body, nextPos, numEnd) Then
                        FormulaHasSignedNumericTerm = True
                        Exit Function
                    End If
                End If
                i = i + 1
            Case Else
                i = i + 1
        End Select
    Loop
End Function

Private Function FormulaHasNumericArgInFunction(ByVal body As String, ByVal fnName As String) As Boolean
    Dim i As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim args() As String
    Dim j As Long

    i = 1
    Do While i <= Len(body)
        If MatchFunctionNameAt(body, i, fnName, openPos) Then
            If TryReadFunctionArgs(body, openPos, args, closePos) Then
                For j = LBound(args) To UBound(args)
                    If IsPureNumericExpression(args(j)) Then
                        FormulaHasNumericArgInFunction = True
                        Exit Function
                    End If
                    If FormulaHasNumericArgInFunction(args(j), fnName) Then
                        FormulaHasNumericArgInFunction = True
                        Exit Function
                    End If
                Next j
                i = closePos + 1
            Else
                i = i + 1
            End If
        ElseIf Mid$(body, i, 1) = """" Then
            i = SkipDoubleQuotedString(body, i)
        Else
            i = i + 1
        End If
    Loop
End Function

Private Function RemoveNumericArgumentsFromNamedFunctions(ByVal body As String, ByVal fnName As String) As String
    Dim outText As String
    Dim i As Long
    Dim lastEmit As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim args() As String
    Dim kept As Collection
    Dim processed As String
    Dim j As Long

    i = 1
    lastEmit = 1

    Do While i <= Len(body)
        If MatchFunctionNameAt(body, i, fnName, openPos) Then
            If TryReadFunctionArgs(body, openPos, args, closePos) Then
                outText = outText & Mid$(body, lastEmit, i - lastEmit)

                Set kept = New Collection
                For j = LBound(args) To UBound(args)
                    processed = RemoveNumericArgumentsFromNamedFunctions(args(j), fnName)
                    processed = CleanupFormulaBody(processed)
                    If Not IsPureNumericExpression(processed) Then
                        If Len(processed) > 0 Then kept.Add processed
                    End If
                Next j

                outText = outText & Mid$(body, i, openPos - i + 1)
                If kept.Count > 0 Then
                    For j = 1 To kept.Count
                        If j > 1 Then outText = outText & ","
                        outText = outText & CStr(kept(j))
                    Next j
                End If
                outText = outText & ")"

                i = closePos + 1
                lastEmit = i
                GoTo ContinueLoop
            End If
        End If

        If Mid$(body, i, 1) = """" Then
            i = SkipDoubleQuotedString(body, i)
        Else
            i = i + 1
        End If

ContinueLoop:
    Loop

    outText = outText & Mid$(body, lastEmit)
    RemoveNumericArgumentsFromNamedFunctions = outText
End Function

Private Function RemoveSignedNumericTerms(ByVal body As String) As String
    Dim outText As String
    Dim i As Long
    Dim prevSig As String
    Dim nextPos As Long
    Dim numEnd As Long

    i = 1
    Do While i <= Len(body)
        If Mid$(body, i, 1) = """" Then
            Dim strEnd As Long
            strEnd = SkipDoubleQuotedString(body, i)
            outText = outText & Mid$(body, i, strEnd - i + 1)
            i = strEnd + 1
        ElseIf (Mid$(body, i, 1) = "+" Or Mid$(body, i, 1) = "-") Then
            prevSig = PreviousSignificantChar(outText, Len(outText))
            If prevSig <> "E" And prevSig <> "e" Then
                nextPos = SkipSpacesLocal(body, i + 1)
                If ReadNumberLiteral(body, nextPos, numEnd) Then
                    i = numEnd + 1
                Else
                    outText = outText & Mid$(body, i, 1)
                    i = i + 1
                End If
            Else
                outText = outText & Mid$(body, i, 1)
                i = i + 1
            End If
        Else
            outText = outText & Mid$(body, i, 1)
            i = i + 1
        End If
    Loop

    RemoveSignedNumericTerms = outText
End Function

Private Function CleanupFormulaBody(ByVal body As String) As String
    Dim s As String
    Dim oldS As String

    s = Trim$(body)
    If Len(s) = 0 Then Exit Function

    Do
        oldS = s
        s = Replace(s, ",,", ",")
        s = Replace(s, "(,", "(")
        s = Replace(s, ",)", ")")
        s = Replace(s, "+,", ",")
        s = Replace(s, "-,", ",")
        s = Replace(s, "++", "+")
        s = Replace(s, "+-", "-")
        s = Replace(s, "-+", "-")
        s = Replace(s, "--", "+")
        s = Trim$(s)

        Do While Len(s) > 0 And (Left$(s, 1) = "+" Or Left$(s, 1) = ",")
            s = Mid$(s, 2)
            s = Trim$(s)
        Loop

        Do While Len(s) > 0 And (Right$(s, 1) = "+" Or Right$(s, 1) = "-" Or Right$(s, 1) = ",")
            s = Left$(s, Len(s) - 1)
            s = Trim$(s)
        Loop
    Loop While s <> oldS

    If s = "()" Then s = vbNullString
    CleanupFormulaBody = s
End Function

Private Function IsPureNumericExpression(ByVal s As String) As Boolean
    Dim p As Long
    Dim endPos As Long

    s = Trim$(s)
    If Len(s) = 0 Then Exit Function

    Do While Len(s) >= 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" And ParenthesesWrapWholeExpression(s)
        s = Trim$(Mid$(s, 2, Len(s) - 2))
    Loop

    p = 1
    If ReadSignedNumberLiteral(s, p, endPos) Then
        If SkipSpacesLocal(s, endPos + 1) > Len(s) Then
            IsPureNumericExpression = True
        End If
    End If
End Function

Private Function ParenthesesWrapWholeExpression(ByVal s As String) As Boolean
    Dim i As Long
    Dim depth As Long

    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> "(" Or Right$(s, 1) <> ")" Then Exit Function

    For i = 1 To Len(s)
        Select Case Mid$(s, i, 1)
            Case """"
                i = SkipDoubleQuotedString(s, i)
            Case "("
                depth = depth + 1
            Case ")"
                depth = depth - 1
                If depth = 0 And i < Len(s) Then Exit Function
        End Select
    Next i

    ParenthesesWrapWholeExpression = (depth = 0)
End Function

Private Function PreviousSignificantChar(ByVal s As String, ByVal pos As Long) As String
    Do While pos >= 1
        If Mid$(s, pos, 1) <> " " Then
            PreviousSignificantChar = Mid$(s, pos, 1)
            Exit Function
        End If
        pos = pos - 1
    Loop
End Function

Private Function SkipSpacesLocal(ByVal s As String, ByVal pos As Long) As Long
    Do While pos <= Len(s) And Mid$(s, pos, 1) = " "
        pos = pos + 1
    Loop
    SkipSpacesLocal = pos
End Function

Private Function ReadSignedNumberLiteral(ByVal s As String, ByVal pos As Long, ByRef endPos As Long) As Boolean
    Dim p As Long

    p = SkipSpacesLocal(s, pos)
    If p <= Len(s) Then
        If Mid$(s, p, 1) = "+" Or Mid$(s, p, 1) = "-" Then
            p = p + 1
        End If
    End If

    ReadSignedNumberLiteral = ReadNumberLiteral(s, p, endPos)
End Function

Private Function ReadNumberLiteral(ByVal s As String, ByVal pos As Long, ByRef endPos As Long) As Boolean
    Dim p As Long
    Dim hasDigit As Boolean
    Dim ch As String

    p = pos
    If p < 1 Or p > Len(s) Then Exit Function

    Do While p <= Len(s) And Mid$(s, p, 1) = " "
        p = p + 1
    Loop
    If p > Len(s) Then Exit Function

    Do While p <= Len(s) And IsDigit09(Mid$(s, p, 1))
        hasDigit = True
        p = p + 1
    Loop

    If p <= Len(s) And Mid$(s, p, 1) = "." Then
        p = p + 1
        Do While p <= Len(s) And IsDigit09(Mid$(s, p, 1))
            hasDigit = True
            p = p + 1
        Loop
    End If

    If Not hasDigit Then Exit Function

    If p <= Len(s) Then
        ch = Mid$(s, p, 1)
        If ch = "E" Or ch = "e" Then
            p = p + 1
            If p <= Len(s) Then
                ch = Mid$(s, p, 1)
                If ch = "+" Or ch = "-" Then p = p + 1
            End If

            hasDigit = False
            Do While p <= Len(s) And IsDigit09(Mid$(s, p, 1))
                hasDigit = True
                p = p + 1
            Loop
            If Not hasDigit Then Exit Function
        End If
    End If

    If p <= Len(s) Then
        ch = Mid$(s, p, 1)
        If IsLetterAZ(ch) Or IsDigit09(ch) Or ch = "_" Or ch = "." Then Exit Function
    End If

    endPos = p - 1
    ReadNumberLiteral = True
End Function

Private Function TryConsumeIndirectVisibleLiteralRef( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByVal hostSheetName As String, _
    ByRef outDict As Object, _
    ByRef endPos As Long) As Boolean

    Dim openParenPos As Long
    Dim closePos As Long
    Dim args() As String
    Dim lit As String
    Dim subRefs As Object
    Dim k As Variant

    endPos = pos
    If Not MatchFunctionNameAt(s, pos, "INDIRECT", openParenPos) Then Exit Function
    If Not TryReadFunctionArgs(s, openParenPos, args, closePos) Then Exit Function
    If UBound(args) < 1 Then Exit Function
    If Not TryGetStandaloneQuotedLiteral(args(1), lit) Then Exit Function

    Set subRefs = ExtractWholeReferenceText(lit, hostSheetName)
    If subRefs Is Nothing Then
        endPos = closePos
        TryConsumeIndirectVisibleLiteralRef = True
        Exit Function
    End If

    For Each k In subRefs.Keys
        If Not outDict.Exists(CStr(k)) Then outDict.Add CStr(k), True
    Next k

    endPos = closePos
    TryConsumeIndirectVisibleLiteralRef = True
End Function

Private Function TryConsumeOffsetVisibleBaseRef( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByVal hostSheetName As String, _
    ByRef outDict As Object, _
    ByRef endPos As Long) As Boolean

    Dim openParenPos As Long
    Dim closePos As Long
    Dim args() As String
    Dim subRefs As Object
    Dim k As Variant

    endPos = pos
    If Not MatchFunctionNameAt(s, pos, "OFFSET", openParenPos) Then Exit Function
    If Not TryReadFunctionArgs(s, openParenPos, args, closePos) Then Exit Function
    If UBound(args) < 1 Then Exit Function

    Set subRefs = ExtractExplicitRefs(args(1), hostSheetName)
    For Each k In subRefs.Keys
        If Not outDict.Exists(CStr(k)) Then outDict.Add CStr(k), True
    Next k

    endPos = closePos
    TryConsumeOffsetVisibleBaseRef = True
End Function

Private Function ExtractWholeReferenceText(ByVal s As String, ByVal hostSheetName As String) As Object
    Dim d As Object
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long
    Dim bangPos As Long
    Dim shName As String
    Dim leftPart As String, rightPart As String
    Dim addr1 As String, addr2 As String
    Dim key As String

    s = Trim$(s)
    If Len(s) = 0 Then Exit Function

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    shName = hostSheetName
    bangPos = InStrRev(s, "!")
    If bangPos > 0 Then
        leftPart = Trim$(Left$(s, bangPos - 1))
        rightPart = Trim$(Mid$(s, bangPos + 1))
        If Len(leftPart) = 0 Or Len(rightPart) = 0 Then Exit Function

        If Left$(leftPart, 1) = "'" And Right$(leftPart, 1) = "'" And Len(leftPart) >= 2 Then
            shName = Replace(Mid$(leftPart, 2, Len(leftPart) - 2), "''", "'")
        Else
            shName = leftPart
        End If

        s = rightPart
    End If

    p1 = 1
    If TryParseCellToken(s, p1, p1, p2, addr1) Then
        If p2 = Len(s) Then
            key = UCase$(shName) & "|" & addr1
            d(key) = True
            Set ExtractWholeReferenceText = d
            Exit Function
        End If

        If Mid$(s, p2 + 1, 1) = ":" Then
            p3 = p2 + 2
            If TryParseCellToken(s, p3, p3, p4, addr2) Then
                If p4 = Len(s) Then
                    d(UCase$(shName) & "|" & addr1) = True
                    d(UCase$(shName) & "|" & addr2) = True
                    Set ExtractWholeReferenceText = d
                    Exit Function
                End If
            End If
        End If
    End If
End Function

Private Sub AddRefKey(ByRef d As Object, ByVal sheetName As String, ByVal addr As String)
    Dim k As String
    k = UCase$(sheetName) & "|" & UCase$(addr)
    If Not d.Exists(k) Then d.Add k, True
End Sub

Private Function ColLettersToNumber2(ByVal colLetters As String) As Long
    Dim i As Long
    Dim v As Long

    For i = 1 To Len(colLetters)
        v = v * 26 + (Asc(Mid$(colLetters, i, 1)) - 64)
    Next i

    ColLettersToNumber2 = v
End Function

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

