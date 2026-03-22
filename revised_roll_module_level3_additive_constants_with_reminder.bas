Option Explicit

Private Const LVL_INTERNAL As Long = 0
Private Const LVL_DIRECT_OTHER As Long = 1
Private Const LVL_PARENT_OF_DIRECT_OTHER As Long = 2
Private Const LVL_HARDCODED As Long = 3


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
    Dim activeSheetName As String

    On Error GoTo CleanFail

    startTick = Timer
    Set ctlWs = ActiveSheet
    activeSheetName = ctlWs.Name

    If Not ConfirmSheetListReminder(ctlWs) Then Exit Sub

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
        Application.StatusBar = "Rolling control row " & r & " of " & lastCtlRow & " [" & _
                                Trim$(CStr(ctlWs.Cells(r, 1).Value)) & "]..."
        ProcessOneControlRow _
            targetSheetName:=Trim$(CStr(ctlWs.Cells(r, 1).Value)), _
            colSpec:=Trim$(CStr(ctlWs.Cells(r, 2).Value)), _
            methodText:=Trim$(CStr(ctlWs.Cells(r, 3).Value)), _
            freezeMaxText:=Trim$(CStr(ctlWs.Cells(r, 4).Value)), _
            directionText:=Trim$(CStr(ctlWs.Cells(r, 5).Value))
    Next r

CleanExit:
    secs = ElapsedSeconds(startTick)
    RestoreAppState oldCalc, oldScreen, oldEvents, oldStatusBar
    MsgBox "Completed on control sheet [" & activeSheetName & "]." & vbCrLf & _
           "Elapsed time: " & Format(secs, "0.00") & " seconds", vbInformation
    Exit Sub

CleanFail:
    secs = ElapsedSeconds(startTick)
    RestoreAppState oldCalc, oldScreen, oldEvents, oldStatusBar
    MsgBox "Error on control row " & r & " (" & Trim$(CStr(ctlWs.Cells(r, 1).Value)) & ")." & vbCrLf & _
           Err.Description & vbCrLf & _
           "Elapsed time: " & Format(secs, "0.00") & " seconds", vbExclamation
End Sub

Private Function ConfirmSheetListReminder(ByVal ctlWs As Worksheet) As Boolean

    Dim msg As String
    Dim sheetLabel As String
    Dim resp As VbMsgBoxResult

    sheetLabel = ctlWs.Name

    msg = "Please confirm the meaning of each column on control sheet [" & sheetLabel & "] before running:" & vbCrLf & vbCrLf & _
          "Column A = Target sheet name to process" & vbCrLf & _
          "Column B = Source month column to roll (number or letters, e.g. 3 or C)" & vbCrLf & _
          "Column C = Method" & vbCrLf & _
          "    INSERT / I = insert next column and roll into it" & vbCrLf & _
          "    UNGROUP / U = only ungroup the next column" & vbCrLf & _
          "Column D = Freeze max level" & vbCrLf & _
          "    0 = normal internal formulas only" & vbCrLf & _
          "    1 = also freeze direct other-sheet formulas" & vbCrLf & _
          "    2 = also freeze parents of level 1" & vbCrLf & _
          "    3 = also freeze formulas with hard-coded +/- adjustments" & vbCrLf & _
          "        (e.g. A1+100, +100, SUM(A1:G1,100), SUM(A1:G1)-0.2)" & vbCrLf & _
          "Column E = Direction" & vbCrLf & _
          "    blank = normal forward roll" & vbCrLf & _
          "    REVERSE = copy back to previous column" & vbCrLf & vbCrLf & _
          "Default behavior if blank:" & vbCrLf & _
          "    Method = INSERT, Freeze level = 2, Direction = forward" & vbCrLf & vbCrLf & _
          "Press OK to continue, or Cancel to stop and review SheetList."

    resp = MsgBox(msg, vbOKCancel + vbInformation, "Confirm SheetList columns")
    ConfirmSheetListReminder = (resp = vbOK)

End Function

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
    Exit Sub

SafeExit:
    Debug.Print "ProcessOneControlRow failed. Sheet=[" & targetSheetName & _
                "], ColSpec=[" & colSpec & "], Method=[" & methodText & _
                "], Direction=[" & directionText & "]. " & Err.Number & " - " & Err.Description
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
    srcColLetters = ColumnNumberToLetters(srcCol)

    If Not isReverse Then
        workSrcCol = srcCol
        tgtCol = srcCol + 1

        Select Case methodUpper
            Case "INSERT", "I"
                ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Case "UNGROUP", "U"
                PrepareExistingTargetColumn ws, tgtCol
            Case Else
                Debug.Print "Unknown method [" & methodText & "] on sheet [" & ws.Name & "]. Defaulted to INSERT."
                ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End Select
    Else
        If methodUpper = "INSERT" Or methodUpper = "I" Then
            tgtCol = srcCol
            ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            workSrcCol = srcCol + 1
        ElseIf methodUpper = "UNGROUP" Or methodUpper = "U" Then
            workSrcCol = srcCol
            tgtCol = srcCol - 1
            PrepareExistingTargetColumn ws, tgtCol
        Else
            Debug.Print "Unknown method [" & methodText & "] on sheet [" & ws.Name & "]. Defaulted to INSERT."
            tgtCol = srcCol
            ws.Columns(tgtCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            workSrcCol = srcCol + 1
        End If
    End If

    CopyColumnLayoutNoComments ws, workSrcCol, tgtCol, lastRow

    For i = 1 To lastRow
        If (i Mod 200) = 0 Then
            Application.StatusBar = "Rolling [" & ws.Name & "] row " & i & " of " & lastRow & "..."
        End If

        isFormula = IsFormulaVariant(arrFormula(i, 1))
        addrOriginal = srcColLetters & CStr(i)

        If isFormula Then
            If levelMap.Exists(addrOriginal) Then
                lvl = CLng(levelMap(addrOriginal))
            Else
                lvl = LVL_INTERNAL
            End If

            If lvl = LVL_HARDCODED And freezeMaxLevel >= LVL_HARDCODED Then
                ws.Cells(i, tgtCol).Formula = CStr(arrFormula(i, 1))
                Debug.Print "Level 3 hard-coded formula preserved without rolling: [" & ws.Name & "!" & addrOriginal & "] -> [" & ws.Name & "!" & ColumnNumberToLetters(tgtCol) & CStr(i) & "]"
            ElseIf ShouldFreezeLevel(lvl, freezeMaxLevel) Then
                ws.Cells(i, workSrcCol).Value = arrValue(i, 1)
                ws.Cells(i, tgtCol).Formula = BuildFrozenTargetFormula(CStr(arrFormula(i, 1)), ws.Name, colDelta)
            Else
                ws.Cells(i, tgtCol).FormulaR1C1 = CStr(arrFormulaR1C1(i, 1))
            End If
        Else
            ws.Cells(i, tgtCol).Value = arrValue(i, 1)
        End If
    Next i
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

Private Sub RestoreAppState( _
    ByVal oldCalc As XlCalculation, _
    ByVal oldScreen As Boolean, _
    ByVal oldEvents As Boolean, _
    ByVal oldStatusBar As Variant)

    Application.StatusBar = oldStatusBar
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
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
    Dim s As String

    s = Trim$(txt)

    If Len(s) = 0 Then
        ParseFreezeMaxLevel = LVL_PARENT_OF_DIRECT_OTHER
    ElseIf IsNumeric(s) Then
        ParseFreezeMaxLevel = CLng(s)
        If ParseFreezeMaxLevel < LVL_INTERNAL Then ParseFreezeMaxLevel = LVL_INTERNAL
        If ParseFreezeMaxLevel > LVL_HARDCODED Then ParseFreezeMaxLevel = LVL_HARDCODED
    Else
        ParseFreezeMaxLevel = LVL_PARENT_OF_DIRECT_OTHER
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
    Dim parentList() As String
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
    Dim parentParts() As String
    Dim part As Variant

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

    ReDim parentList(1 To n)
    ReDim addrArr(1 To n)
    ReDim hasDirectOther(1 To n)
    ReDim levelCap(1 To n)

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

        If FormulaHasHardCodedNumbers(CStr(c.Formula)) Then
            levelCap(idx) = LVL_HARDCODED
        End If

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
                            parentList(childIdx) = parentList(childIdx) & CStr(idx) & "|"
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
            If levelCap(i) <> LVL_HARDCODED Then levelCap(i) = LVL_DIRECT_OTHER
            tail = tail + 1
            q(tail) = i
        End If
    Next i

    Do While head <= tail
        childIdx = q(head)
        head = head + 1

        If Len(parentList(childIdx)) > 0 Then
            parentParts = Split(parentList(childIdx), "|")
            For Each part In parentParts
                If Len(CStr(part)) > 0 Then
                    parentIdx = CLng(part)
                    If levelCap(parentIdx) = LVL_INTERNAL Then
                        levelCap(parentIdx) = LVL_PARENT_OF_DIRECT_OTHER
                        tail = tail + 1
                        q(tail) = parentIdx
                    End If
                End If
            Next part
        End If
    Loop

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
    Dim consumeRes As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    parseOk = True
    n = Len(formulaText)
    i = 1

    Do While i <= n

        consumeRes = TryConsumeIndirectOrIgnore(formulaText, i, hostSheetName, hostRow, hostCol, d, tokenEnd)
        If consumeRes <> 0 Then
            If consumeRes = 2 And failOnUnresolvedDynamic Then
                parseOk = False
                Exit Do
            End If
            i = tokenEnd + 1
            GoTo ContinueLoop
        End If

        consumeRes = TryConsumeOffsetOrIgnore(formulaText, i, hostSheetName, d, tokenEnd)
        If consumeRes <> 0 Then
            If consumeRes = 2 And failOnUnresolvedDynamic Then
                parseOk = False
                Exit Do
            End If
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
    Dim capacity As Long
    Dim finalArg As String

    TryReadFunctionArgs = False
    closeParenPos = 0

    If openParenPos < 1 Or openParenPos > Len(s) Then Exit Function
    If Mid$(s, openParenPos, 1) <> "(" Then Exit Function

    n = Len(s)
    depth = 1
    argStart = openParenPos + 1
    p = openParenPos + 1
    argCount = 0
    capacity = 8
    ReDim args(1 To capacity)

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
                finalArg = Trim$(Mid$(s, argStart, p - argStart))
                If argCount = 0 And Len(finalArg) = 0 Then
                    ReDim args(0 To 0)
                    args(0) = vbNullString
                Else
                    argCount = argCount + 1
                    If argCount > capacity Then
                        capacity = capacity * 2
                        ReDim Preserve args(1 To capacity)
                    End If
                    args(argCount) = finalArg
                    ReDim Preserve args(1 To argCount)
                End If

                closeParenPos = p
                TryReadFunctionArgs = True
                Exit Function
            Else
                p = p + 1
            End If
        ElseIf ch = "," And depth = 1 Then
            argCount = argCount + 1
            If argCount > capacity Then
                capacity = capacity * 2
                ReDim Preserve args(1 To capacity)
            End If
            args(argCount) = Trim$(Mid$(s, argStart, p - argStart))
            argStart = p + 1
            p = p + 1
        Else
            p = p + 1
        End If
    Loop

    ReDim args(0 To 0)
    args(0) = vbNullString
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
    Dim rawText As String

    endQuotePos = 0
    ReadDoubleQuotedLiteral = vbNullString

    If startQuotePos < 1 Or startQuotePos > Len(s) Then Exit Function
    If Mid$(s, startQuotePos, 1) <> """" Then Exit Function

    n = Len(s)
    p = startQuotePos + 1

    Do While p <= n
        If Mid$(s, p, 1) = """" Then
            If p < n And Mid$(s, p + 1, 1) = """" Then
                p = p + 2
            Else
                endQuotePos = p
                rawText = Mid$(s, startQuotePos + 1, endQuotePos - startQuotePos - 1)
                ReadDoubleQuotedLiteral = Replace(rawText, """""", """")
                Exit Function
            End If
        Else
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

Private Function FormulaHasHardCodedNumbers(ByVal formulaText As String) As Boolean
    Dim n As Long
    Dim i As Long
    Dim tokenStart As Long
    Dim tokenEnd As Long
    Dim addrNorm As String
    Dim normText As String
    Dim ch As String
    Dim funcStack() As String
    Dim stackTop As Long
    Dim pendingFuncName As String
    Dim identEnd As Long
    Dim identText As String

    n = Len(formulaText)
    i = 1

    Do While i <= n
        ch = Mid$(formulaText, i, 1)

        If ch = """ Then
            i = SkipDoubleQuotedString(formulaText, i)
            GoTo ContinueLoop
        End If

        If ch = "'" Then
            tokenEnd = SkipSingleQuotedSheetQualifier(formulaText, i)
            If tokenEnd > i Then
                i = tokenEnd
                GoTo ContinueLoop
            End If
        End If

        If TryParseColumnRangeToken(formulaText, i, tokenStart, tokenEnd, normText) Then
            i = tokenEnd + 1
            pendingFuncName = vbNullString
            GoTo ContinueLoop
        End If

        If TryParseCellToken(formulaText, i, tokenStart, tokenEnd, addrNorm) Then
            i = tokenEnd + 1
            pendingFuncName = vbNullString
            GoTo ContinueLoop
        End If

        If TryParseBareIdentifier(formulaText, i, identEnd, identText) Then
            pendingFuncName = vbNullString
            If NextSignificantChar(formulaText, identEnd + 1) = "(" Then
                pendingFuncName = UCase$(identText)
            End If
            i = identEnd + 1
            GoTo ContinueLoop
        End If

        If ch = "(" Then
            PushFunctionName funcStack, stackTop, pendingFuncName
            pendingFuncName = vbNullString
            i = i + 1
            GoTo ContinueLoop
        End If

        If ch = ")" Then
            If stackTop > 0 Then stackTop = stackTop - 1
            pendingFuncName = vbNullString
            i = i + 1
            GoTo ContinueLoop
        End If

        If IsNumericLiteralStart(formulaText, i) Then
            tokenEnd = NumericLiteralTokenEnd(formulaText, i)
            If NumericLiteralMeansHardCodedLevel3(formulaText, i, funcStack, stackTop) Then
                FormulaHasHardCodedNumbers = True
                Exit Function
            End If
            If tokenEnd < i Then tokenEnd = i
            i = tokenEnd + 1
            pendingFuncName = vbNullString
            GoTo ContinueLoop
        End If

        If ch <> " " Then pendingFuncName = vbNullString
        i = i + 1
ContinueLoop:
    Loop
End Function

Private Function NumericLiteralMeansHardCodedLevel3( _
    ByVal formulaText As String, _
    ByVal tokenStart As Long, _
    ByRef funcStack() As String, _
    ByVal stackTop As Long) As Boolean

    Dim prevPos As Long
    Dim prevCh As String
    Dim tokenFirstCh As String
    Dim activeFunc As String

    prevPos = PrevNonSpacePos(formulaText, tokenStart - 1)
    If prevPos > 0 Then prevCh = Mid$(formulaText, prevPos, 1)
    tokenFirstCh = Mid$(formulaText, tokenStart, 1)
    activeFunc = CurrentFunctionName(funcStack, stackTop)

    If prevPos = 1 And Mid$(formulaText, 1, 1) = "=" Then
        NumericLiteralMeansHardCodedLevel3 = True
        Exit Function
    End If

    If tokenFirstCh = "+" Or tokenFirstCh = "-" Then
        Select Case prevCh
            Case ",", ";"
                NumericLiteralMeansHardCodedLevel3 = (activeFunc = "SUM")
            Case "("
                NumericLiteralMeansHardCodedLevel3 = (activeFunc = "SUM")
            Case "<", ">", "="
                NumericLiteralMeansHardCodedLevel3 = False
            Case Else
                NumericLiteralMeansHardCodedLevel3 = True
        End Select
        Exit Function
    End If

    Select Case prevCh
        Case "+", "-"
            NumericLiteralMeansHardCodedLevel3 = True
        Case ",", ";", "("
            NumericLiteralMeansHardCodedLevel3 = (activeFunc = "SUM")
        Case Else
            NumericLiteralMeansHardCodedLevel3 = False
    End Select
End Function

Private Sub PushFunctionName(ByRef funcStack() As String, ByRef stackTop As Long, ByVal funcName As String)
    stackTop = stackTop + 1
    ReDim Preserve funcStack(1 To stackTop)
    funcStack(stackTop) = UCase$(Trim$(funcName))
End Sub

Private Function CurrentFunctionName(ByRef funcStack() As String, ByVal stackTop As Long) As String
    Dim i As Long

    For i = stackTop To 1 Step -1
        If Len(funcStack(i)) > 0 Then
            CurrentFunctionName = funcStack(i)
            Exit Function
        End If
    Next i
End Function

Private Function TryParseBareIdentifier( _
    ByVal s As String, _
    ByVal pos As Long, _
    ByRef tokenEnd As Long, _
    ByRef identText As String) As Boolean

    Dim n As Long
    Dim p As Long
    Dim ch As String

    n = Len(s)
    If pos < 1 Or pos > n Then Exit Function

    ch = Mid$(s, pos, 1)
    If Not IsLetterAZ(ch) And ch <> "_" Then Exit Function

    p = pos + 1
    Do While p <= n
        ch = Mid$(s, p, 1)
        If IsLetterAZ(ch) Or IsDigit09(ch) Or ch = "_" Or ch = "." Then
            p = p + 1
        Else
            Exit Do
        End If
    Loop

    tokenEnd = p - 1
    identText = Mid$(s, pos, tokenEnd - pos + 1)
    TryParseBareIdentifier = True
End Function

Private Function NextSignificantChar(ByVal s As String, ByVal pos As Long) As String
    Dim n As Long
    Dim p As Long

    n = Len(s)
    p = pos

    Do While p <= n
        If Mid$(s, p, 1) <> " " Then
            NextSignificantChar = Mid$(s, p, 1)
            Exit Function
        End If
        p = p + 1
    Loop
End Function

Private Function PrevNonSpacePos(ByVal s As String, ByVal pos As Long) As Long
    Dim p As Long

    p = pos
    Do While p >= 1
        If Mid$(s, p, 1) <> " " Then
            PrevNonSpacePos = p
            Exit Function
        End If
        p = p - 1
    Loop
End Function

Private Function NumericLiteralTokenEnd(ByVal s As String, ByVal pos As Long) As Long
    Dim n As Long
    Dim p As Long
    Dim ch As String

    n = Len(s)
    p = pos
    If pos < 1 Or pos > n Then Exit Function

    ch = Mid$(s, p, 1)
    If ch = "+" Or ch = "-" Then p = p + 1

    Do While p <= n And IsDigit09(Mid$(s, p, 1))
        p = p + 1
    Loop

    If p <= n And Mid$(s, p, 1) = "." Then
        p = p + 1
        Do While p <= n And IsDigit09(Mid$(s, p, 1))
            p = p + 1
        Loop
    End If

    If p <= n Then
        ch = Mid$(s, p, 1)
        If ch = "E" Or ch = "e" Then
            p = p + 1
            If p <= n Then
                ch = Mid$(s, p, 1)
                If ch = "+" Or ch = "-" Then p = p + 1
            End If
            Do While p <= n And IsDigit09(Mid$(s, p, 1))
                p = p + 1
            Loop
        End If
    End If

    NumericLiteralTokenEnd = p - 1
End Function

Private Function SkipSingleQuotedSheetQualifier(ByVal s As String, ByVal pos As Long) As Long
    Dim n As Long
    Dim p As Long

    n = Len(s)
    If pos < 1 Or pos > n Then Exit Function
    If Mid$(s, pos, 1) <> "'" Then Exit Function

    p = pos + 1
    Do While p <= n
        If Mid$(s, p, 1) = "'" Then
            If p < n And Mid$(s, p + 1, 1) = "'" Then
                p = p + 2
            Else
                p = p + 1
                Do While p <= n And Mid$(s, p, 1) = " "
                    p = p + 1
                Loop
                If p <= n And Mid$(s, p, 1) = "!" Then
                    SkipSingleQuotedSheetQualifier = p + 1
                End If
                Exit Function
            End If
        Else
            p = p + 1
        End If
    Loop
End Function

Private Function IsNumericLiteralStart(ByVal s As String, ByVal pos As Long) As Boolean
    Dim n As Long
    Dim p As Long
    Dim prevCh As String
    Dim ch As String
    Dim sawDigit As Boolean

    n = Len(s)
    If pos < 1 Or pos > n Then Exit Function

    ch = Mid$(s, pos, 1)
    p = pos

    If ch = "+" Or ch = "-" Then
        If pos > 1 Then
            prevCh = Mid$(s, pos - 1, 1)
            If prevCh <> "(" And prevCh <> "," And prevCh <> ";" And prevCh <> "=" And _
               prevCh <> "+" And prevCh <> "-" And prevCh <> "*" And prevCh <> "/" And _
               prevCh <> "^" And prevCh <> "&" And prevCh <> "{" Then
                Exit Function
            End If
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
        Do While p <= n And IsDigit09(Mid$(s, p, 1))
            sawDigit = True
            p = p + 1
        Loop
    ElseIf IsDigit09(ch) Then
        Do While p <= n And IsDigit09(Mid$(s, p, 1))
            sawDigit = True
            p = p + 1
        Loop

        If p <= n And Mid$(s, p, 1) = "." Then
            p = p + 1
            Do While p <= n And IsDigit09(Mid$(s, p, 1))
                sawDigit = True
                p = p + 1
            Loop
        End If
    Else
        Exit Function
    End If

    If Not sawDigit Then Exit Function

    If p <= n Then
        ch = Mid$(s, p, 1)
        If ch = "E" Or ch = "e" Then
            p = p + 1
            If p <= n Then
                ch = Mid$(s, p, 1)
                If ch = "+" Or ch = "-" Then p = p + 1
            End If
            If p > n Or Not IsDigit09(Mid$(s, p, 1)) Then Exit Function
            Do While p <= n And IsDigit09(Mid$(s, p, 1))
                p = p + 1
            Loop
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

