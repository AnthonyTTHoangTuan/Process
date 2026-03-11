Option Explicit

'=============================================================================
' ROLL COLUMN MACRO
' SheetList structure:
'   Col A = Sheet Name
'   Col B = Column to roll (letter or number)
'   Col C = Insert Direction: "Left" or "Right"
'   Col D = Layout:
'           blank        = Normal layout (Previous | Current)
'           "Reverse"    = Reverse layout (Current | Previous)
'           "Ungrouped"  = Same full roll applies PLUS ungroup the outer
'                          neighbour of the newly inserted column
'                          (Insert Left  => ungroup col to LEFT  of new col)
'                          (Insert Right => ungroup col to RIGHT of new col)
'=============================================================================

Sub RollColumns()
    Dim wsConfig  As Worksheet
    Dim wsTarget  As Worksheet
    Dim lastRow   As Long
    Dim i         As Long
    Dim sheetName As String
    Dim colRef    As String
    Dim insertDir As String
    Dim layout    As String
    Dim targetCol As Long

    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("SheetList")
    On Error GoTo 0
    If wsConfig Is Nothing Then
        MsgBox "Cannot find sheet named 'SheetList'.", vbCritical
        Exit Sub
    End If

    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No entries found in SheetList.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim errLog As String
    errLog = ""

    For i = 2 To lastRow
        sheetName = Trim(wsConfig.Cells(i, 1).Value)
        colRef    = Trim(wsConfig.Cells(i, 2).Value)
        insertDir = Trim(wsConfig.Cells(i, 3).Value)
        layout    = Trim(wsConfig.Cells(i, 4).Value)

        If sheetName = "" Or colRef = "" Then GoTo NextRow

        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        If wsTarget Is Nothing Then
            errLog = errLog & "Sheet not found: " & sheetName & vbNewLine
            GoTo NextRow
        End If

        If IsNumeric(colRef) Then
            targetCol = CLng(colRef)
        Else
            targetCol = ColLetterToNumber(colRef)
        End If
        If targetCol < 1 Then
            errLog = errLog & "Invalid column '" & colRef & "' on sheet '" & sheetName & "'" & vbNewLine
            GoTo NextRow
        End If

        ' Ungrouped layout: no insert. Copy formulas from targetCol into the
        ' neighbour col (per Col C direction), apply Previous-column formula
        ' rules to that neighbour, then ungroup it.
        If UCase(layout) = "UNGROUPED" Then
            If UCase(insertDir) <> "LEFT" And UCase(insertDir) <> "RIGHT" Then
                errLog = errLog & "Ungrouped requires direction on sheet '" & sheetName & "'" & vbNewLine
                GoTo NextRow
            End If
            Dim ungroupNeighbour As Long
            If UCase(insertDir) = "RIGHT" Then
                ungroupNeighbour = targetCol + 1
            Else
                ungroupNeighbour = targetCol - 1
            End If
            If ungroupNeighbour < 1 Or ungroupNeighbour > wsTarget.Columns.Count Then
                errLog = errLog & "Ungrouped neighbour out of range on sheet '" & sheetName & "'" & vbNewLine
                GoTo NextRow
            End If
            Dim ungroupLastRow As Long
            ungroupLastRow = wsTarget.Cells(wsTarget.Rows.Count, targetCol).End(xlUp).Row
            If ungroupLastRow < 1 Then ungroupLastRow = wsTarget.UsedRange.Rows.Count
            Dim ungroupOffset As Long
            ungroupOffset = ungroupNeighbour - targetCol
            Call CopyFormulasAsIs(wsTarget, targetCol, ungroupNeighbour, ungroupOffset, ungroupLastRow)
            On Error Resume Next
            wsTarget.Columns(ungroupNeighbour).Ungroup
            On Error GoTo 0
            GoTo NextRow
        End If

        If UCase(insertDir) <> "LEFT" And UCase(insertDir) <> "RIGHT" Then
            errLog = errLog & "Invalid direction '" & insertDir & "' on sheet '" & sheetName & "'" & vbNewLine
            GoTo NextRow
        End If

        Dim isReverse As Boolean
        isReverse = (UCase(layout) = "REVERSE" Or UCase(layout) = "PREVRIGHT")

        '----------------------------------------------------------------------
        ' Normal layout  (Previous | Current):
        '   Insert Left  => new col = Previous,  existing stays Current
        '   Insert Right => new col = Current,   existing becomes Previous
        '
        ' Reverse layout (Current | Previous):
        '   Insert Left  => new col = Current,   existing becomes Previous
        '   Insert Right => new col = Previous,  existing stays Current
        '----------------------------------------------------------------------
        Dim newIsCurrent As Boolean
        If isReverse Then
            newIsCurrent = (UCase(insertDir) = "LEFT")
        Else
            newIsCurrent = (UCase(insertDir) = "RIGHT")
        End If

        ' Capture comments from targetCol BEFORE insert in case Excel
        ' shifts/loses them during column insertion
        Dim preRollComments As Object
        Set preRollComments = CaptureComments(wsTarget, targetCol)

        Dim newCol      As Long
        Dim existingCol As Long
        If UCase(insertDir) = "LEFT" Then
            wsTarget.Columns(targetCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            newCol      = targetCol
            existingCol = targetCol + 1
        Else
            wsTarget.Columns(targetCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            newCol      = targetCol + 1
            existingCol = targetCol
        End If

        Dim currentCol  As Long
        Dim previousCol As Long
        If newIsCurrent Then
            currentCol  = newCol
            previousCol = existingCol
        Else
            previousCol = newCol
            currentCol  = existingCol
        End If

        Call ProcessRoll(wsTarget, currentCol, previousCol, existingCol, newCol, newIsCurrent, preRollComments)

NextRow:
        Set wsTarget = Nothing
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    If errLog <> "" Then
        MsgBox "Roll completed with warnings:" & vbNewLine & errLog, vbExclamation, "Roll Warnings"
    Else
        MsgBox "Roll completed successfully.", vbInformation, "Done"
    End If
End Sub

'=============================================================================
' Core roll: comments, formulas, figures
' preRollComments: Collection of {row, text} captured from targetCol BEFORE
' the insert, so comments are never lost due to Excel's column shift behaviour.
'=============================================================================
Private Sub ProcessRoll(ws As Worksheet, _
                        currentCol As Long, previousCol As Long, _
                        existingCol As Long, newCol As Long, _
                        newIsCurrent As Boolean, _
                        preRollComments As Object)

    Dim lastRow As Long
    Dim lastRowNew As Long
    lastRow    = ws.Cells(ws.Rows.Count, existingCol).End(xlUp).Row
    lastRowNew = ws.Cells(ws.Rows.Count, newCol).End(xlUp).Row
    If lastRowNew > lastRow Then lastRow = lastRowNew
    If lastRow < 1 Then lastRow = ws.UsedRange.Rows.Count

    ' --- COMMENTS ---
    ' Rule: Previous col = inherits comments from pre-roll Current (targetCol)
    '       Current col  = always blank
    '
    ' We use preRollComments (captured before insert) as the authoritative
    ' source so Excel's column shift cannot lose them.
    '
    ' Case A: newIsCurrent = True  (newCol=Current, existingCol=Previous)
    '   previousCol = existingCol: restore pre-roll comments there.
    '   currentCol  = newCol:      clear.
    '
    ' Case B: newIsCurrent = False (newCol=Previous, existingCol=Current)
    '   previousCol = newCol:      restore pre-roll comments there.
    '   currentCol  = existingCol: clear.

    ' Clear current col first
    Call ClearColumnComments(ws, currentCol, lastRow)
    ' Restore pre-roll comments onto previous col
    Call ClearColumnComments(ws, previousCol, lastRow)
    Call RestoreComments(ws, previousCol, preRollComments)

    ' --- FORMULAS / VALUES ---
    Call ProcessFormulas(ws, currentCol, previousCol, existingCol, newCol, newIsCurrent, lastRow)

    ' --- FIGURES ---
    Call ProcessFigures(ws, currentCol, previousCol)

End Sub

'=============================================================================
' CopyFormulasAsIs
' Used by the Ungrouped path: copies formulas from srcCol into dstCol
' exactly as-is — no sandwich/external checks, no value conversion.
' Relative column references are shifted by colOffset so the formula stays
' correct relative to dstCol's position (same rolling behaviour as Current col).
' Source col formulas are left untouched.
' Number format is also copied.
'=============================================================================
Private Sub CopyFormulasAsIs(ws As Worksheet, _
                             srcCol    As Long, _
                             dstCol    As Long, _
                             colOffset As Long, _
                             lastRow   As Long)
    Dim r       As Long
    Dim srcCell As Range
    Dim dstCell As Range

    For r = 1 To lastRow
        Set srcCell = ws.Cells(r, srcCol)
        Set dstCell = ws.Cells(r, dstCol)

        If srcCell.HasFormula Then
            Dim shifted As String
            shifted = ShiftRelativeColumns(srcCell.Formula, ws.Name, srcCell.Row, colOffset)
            dstCell.Formula = shifted
        ElseIf Not IsEmpty(srcCell) Then
            dstCell.Value = srcCell.Value
        End If

        dstCell.NumberFormat = srcCell.NumberFormat
    Next r
End Sub

'=============================================================================
' FORMULA PROCESSING
'
' New column inherits formulas from existingCol. Because we assign .Formula
' to a cell at a different column, Excel auto-adjusts relative references —
' this handles the "current col inherits rolled formula" requirement for free.
'
' Column becoming Previous: decide per-cell:
'   Rule 1 – external sheet reference            => paste as value
'   Rule 2 – same-row sandwich                   => paste as value
'             (result cell column falls between the min and max column of all
'              same-row references in the formula, e.g. O13 in H13:Q13)
'   Otherwise                                    => keep formula, but shift
'             every relative column reference by colOffset so the formula
'             stays correct relative to the previous column's new position.
'=============================================================================
Private Sub ProcessFormulas(ws As Worksheet, _
                            currentCol As Long, previousCol As Long, _
                            existingCol As Long, newCol As Long, _
                            newIsCurrent As Boolean, lastRow As Long)

    Dim r        As Long
    Dim srcCell  As Range
    Dim dstCell  As Range
    Dim prevCell As Range
    ' Offset from existingCol to newCol (used to adjust previous-col formulas)
    Dim colOffset As Long
    colOffset = newCol - existingCol   ' e.g. +1 if inserted right, -1 if inserted left

    For r = 1 To lastRow
        Set srcCell = ws.Cells(r, existingCol)
        Set dstCell = ws.Cells(r, newCol)

        ' --- Copy formula/value to new column (Current col) ---
        ' Use ShiftRelativeColumns so relative refs are adjusted by colOffset.
        ' (.Formula assignment does NOT auto-adjust; only Copy/Paste does.)
        If srcCell.HasFormula Then
            Dim currShifted As String
            currShifted = ShiftRelativeColumns(srcCell.formula, ws.Name, srcCell.Row, colOffset)
            dstCell.formula = currShifted
        ElseIf Not IsEmpty(srcCell) Then
            dstCell.Value = srcCell.Value
        End If
        dstCell.NumberFormat = srcCell.NumberFormat

        ' --- Determine the Previous cell for this row ---
        If newIsCurrent Then
            Set prevCell = ws.Cells(r, existingCol)
        Else
            Set prevCell = ws.Cells(r, newCol)
        End If

        If prevCell.HasFormula Then
            Dim fml          As String
            Dim pasteAsValue As Boolean
            fml          = prevCell.formula
            pasteAsValue = False

            ' Rule 1: external sheet reference
            If FormulaLinksExternalSheet(fml, ws.Name) Then
                pasteAsValue = True
            End If

            ' Rule 2: same-row sandwich
            ' Result column (prevCell.Column) lies BETWEEN min and max
            ' of all same-row referenced columns in the formula
            If Not pasteAsValue Then
                If IsSandwichFormula(fml, prevCell.Column, prevCell.Row) Then
                    pasteAsValue = True
                End If
            End If

            If pasteAsValue Then
                prevCell.Value = prevCell.Value
            Else
                ' Keep formula but shift relative column references by colOffset
                ' so the formula is correct from the previous column's position
                Dim shifted As String
                shifted = ShiftRelativeColumns(fml, ws.Name, prevCell.Row, colOffset)
                If shifted <> fml Then
                    prevCell.formula = shifted
                End If
            End If
        End If

    Next r

End Sub

'=============================================================================
' IsSandwichFormula
' Returns True ONLY when ALL of the following hold:
'   (a) The formula contains at least one RANGE using the colon operator.
'       Accepted range forms:
'         - Normal A1 range:   H13:Q13
'         - INDIRECT R1C1 range endpoint: H14:INDIRECT("RC[-1]",FALSE)
'           where "RC[n]" means same row, relative column offset.
'       Plain individual cell refs like SUM(H13,Q13) do NOT qualify.
'   (b) Every explicit A1 cell reference in the formula is on the same row
'       as formulaRow. Any cross-row ref disqualifies entirely.
'   (c) resultCol falls STRICTLY between the minimum and maximum column
'       numbers found (including any resolved INDIRECT RC offset columns).
'
' INDIRECT("RC[-1]",FALSE) resolution:
'   "RC[n]"  => same row, resultCol + n  (e.g. RC[-1] => resultCol - 1)
'   "RC"     => same row, resultCol      (current column)
'   "RCn"    => same row, absolute col n (1-based)
'   Row part must be plain "R" (same row); "R[n]" or "Rn" = cross-row => disqualify.
'
' Examples:
'   TRUE  – O14: =ROUND(Q14-SUM(H14:INDIRECT("RC[-1]",FALSE)),2)
'            Range H14:INDIRECT("RC[-1]") => RC[-1] resolves to col O-1 = N(14)
'            Explicit refs: Q14(col17), H14(col8). Indirect endpoint: N(14).
'            min=8(H), max=17(Q), resultCol O=15 => 8<15<17 => TRUE
'
'   TRUE  – O13: =ROUND(R13-SUM(H13:Q13),2)
'            Range H13:Q13, all row 13, min=8 max=18, O=15 => TRUE
'
'   FALSE – O13: =SUM(H13,Q13,R13)
'            No colon range => FALSE
'
'   FALSE – O13: =R13-SUM(H12:Q13)
'            H12 is row 12 (cross-row) => disqualified => FALSE
'=============================================================================
Private Function IsSandwichFormula(formula As String, _
                                   resultCol  As Long, _
                                   formulaRow As Long) As Boolean

    IsSandwichFormula = False

    Dim regRange  As Object
    Dim regAll    As Object
    Dim regIndir  As Object
    On Error Resume Next
    Set regRange = CreateObject("VBScript.RegExp")
    Set regAll   = CreateObject("VBScript.RegExp")
    Set regIndir = CreateObject("VBScript.RegExp")
    On Error GoTo 0
    If regRange Is Nothing Then Exit Function

    '--------------------------------------------------------------------------
    ' (a) Gate: formula must contain a colon-range.
    '     Accept both normal ranges AND ranges where one endpoint is INDIRECT.
    '     Pattern matches:
    '       A1:A1  style  OR  A1:INDIRECT(...)  OR  INDIRECT(...):A1
    '--------------------------------------------------------------------------
    regRange.Pattern    = "(\$?[A-Za-z]{1,3}\$?\d+\s*:\s*\$?[A-Za-z]{1,3}\$?\d+)" & _
                          "|(\$?[A-Za-z]{1,3}\$?\d+\s*:\s*INDIRECT\s*\()" & _
                          "|(INDIRECT\s*\([^)]*\)\s*:\s*\$?[A-Za-z]{1,3}\$?\d+)"
    regRange.Global     = True
    regRange.IgnoreCase = True
    If Not regRange.Test(formula) Then Exit Function

    '--------------------------------------------------------------------------
    ' (b) Collect min/max col from all explicit A1 refs; cross-row disqualifies
    '--------------------------------------------------------------------------
    regAll.Pattern    = "(?:'[^']+'!|[A-Za-z_]\w*!)?\$?([A-Za-z]{1,3})\$?(\d+)"
    regAll.Global     = True
    regAll.IgnoreCase = True

    Dim minCol   As Long
    Dim maxCol   As Long
    Dim refCol   As Long
    Dim refRow   As Long
    Dim foundAny As Boolean
    minCol   = 2147483647
    maxCol   = 0
    foundAny = False

    If regAll.Test(formula) Then
        Dim mAll As Object
        Dim ma   As Object
        Set mAll = regAll.Execute(formula)
        For Each ma In mAll
            If IsNumeric(ma.SubMatches(1)) Then
                refRow = CLng(ma.SubMatches(1))
                If refRow <> formulaRow Then
                    Exit Function          ' cross-row ref — disqualify
                End If
                refCol   = ColLetterToNumber(ma.SubMatches(0))
                foundAny = True
                If refCol < minCol Then minCol = refCol
                If refCol > maxCol Then maxCol = refCol
            End If
        Next ma
    End If

    '--------------------------------------------------------------------------
    ' (c) Resolve any INDIRECT("RC[n]", FALSE) / INDIRECT("RCn") endpoints
    '     and fold them into min/max.
    '     Supported patterns inside the string argument:
    '       RC[n]   => same row, resultCol + n
    '       RC[-n]  => same row, resultCol - n
    '       RC      => same row, resultCol  (n=0)
    '       RCn     => same row, absolute column n (1-based)
    '     Row part must be plain "R" (no offset/absolute) — otherwise cross-row.
    '--------------------------------------------------------------------------
    ' Match INDIRECT("...") or INDIRECT('...')
    regIndir.Pattern    = "INDIRECT\s*\(\s*[""']([^""']*)[""']\s*(?:,[^)]+)?\)"
    regIndir.Global     = True
    regIndir.IgnoreCase = True

    If regIndir.Test(formula) Then
        Dim mIndir  As Object
        Dim mi      As Object
        Dim r1c1    As String
        Dim regRC   As Object
        Set regRC = CreateObject("VBScript.RegExp")
        ' Match R1C1 same-row pattern: R then C then optional [offset] or absolute
        ' Row must be plain "R" (no bracket/digit after R before C)
        regRC.Pattern    = "^R(C(\[(-?\d+)\]|(\d+))?)?$"
        regRC.IgnoreCase = True

        Set mIndir = regIndir.Execute(formula)
        For Each mi In mIndir
            r1c1 = Trim(mi.SubMatches(0))   ' content inside the quotes

            ' Check row part: must start with "R" then immediately "C" or end
            ' i.e. no R[n] or Rn before the C
            If Not regRC.Test(r1c1) Then
                ' Row offset present => cross-row INDIRECT => disqualify
                Exit Function
            End If

            ' Resolve column
            Dim rc As Object
            Set rc = regRC.Execute(r1c1)
            If rc.Count = 0 Then Exit Function

            Dim indirCol As Long
            Dim cPart    As String
            cPart = rc(0).SubMatches(1)   ' everything after "C"

            If cPart = "" Then
                ' Plain "RC" => same column as result
                indirCol = resultCol
            ElseIf rc(0).SubMatches(2) <> "" Then
                ' RC[n] or RC[-n] — relative offset
                indirCol = resultCol + CLng(rc(0).SubMatches(2))
            ElseIf rc(0).SubMatches(3) <> "" Then
                ' RCn — absolute column (1-based)
                indirCol = CLng(rc(0).SubMatches(3))
            Else
                indirCol = resultCol
            End If

            If indirCol < 1 Then Exit Function   ' resolved out of range

            foundAny = True
            If indirCol < minCol Then minCol = indirCol
            If indirCol > maxCol Then maxCol = indirCol
        Next mi
    End If

    If Not foundAny Then Exit Function

    ' Sandwich: result column strictly between min and max
    IsSandwichFormula = (resultCol > minCol And resultCol < maxCol)
End Function

'=============================================================================
' ShiftRelativeColumns
' Rewrites all RELATIVE (non-$-locked) column references in an A1 formula
' by adding colOffset to their column number.
' Absolute references ($A1, $A$1) are left unchanged.
' References to other sheets are left unchanged.
'
' Strategy: parse every A1 cell reference token, check if column is relative
' (no $ before the column letters), then replace col letters with the shifted
' column letters. Row relativity is preserved as-is.
'=============================================================================
Private Function ShiftRelativeColumns(formula As String, _
                                      wsName   As String, _
                                      formulaRow As Long, _
                                      colOffset  As Long) As String
    If colOffset = 0 Then
        ShiftRelativeColumns = formula
        Exit Function
    End If

    Dim regex   As Object
    Dim matches As Object
    Dim m       As Object

    On Error Resume Next
    Set regex = CreateObject("VBScript.RegExp")
    On Error GoTo 0
    If regex Is Nothing Then
        ShiftRelativeColumns = formula
        Exit Function
    End If

    ' Capture groups:
    '   (1) optional sheet prefix including "!"  e.g. "'Sheet1'!" or "Sheet1!"
    '   (2) optional "$" before column  (absolute column marker)
    '   (3) column letters
    '   (4) optional "$" before row     (absolute row marker)
    '   (5) row digits
    regex.Pattern    = "((?:'[^']+'|[A-Za-z_]\w*)!)?" & _
                       "(\$?)([A-Za-z]{1,3})(\$?)(\d+)"
    regex.Global     = True
    regex.IgnoreCase = False   ' preserve case of sheet names

    If Not regex.Test(formula) Then
        ShiftRelativeColumns = formula
        Exit Function
    End If

    Set matches = regex.Execute(formula)

    ' Build result by walking through matches in reverse order
    ' (reverse so string positions stay valid as we replace)
    Dim result As String
    result = formula

    ' Collect replacements: index, length, replacement string
    Dim idxArr()   As Long
    Dim lenArr()   As Long
    Dim repArr()   As String
    Dim cnt        As Long
    cnt = matches.Count
    ReDim idxArr(cnt - 1)
    ReDim lenArr(cnt - 1)
    ReDim repArr(cnt - 1)

    Dim idx As Long
    idx = 0
    For Each m In matches
        Dim sheetPrefix  As String
        Dim dollarCol    As String
        Dim colLetters   As String
        Dim dollarRow    As String
        Dim rowDigits    As String

        sheetPrefix = m.SubMatches(0)  ' may be ""
        dollarCol   = m.SubMatches(1)  ' "$" or ""
        colLetters  = m.SubMatches(2)
        dollarRow   = m.SubMatches(3)  ' "$" or ""
        rowDigits   = m.SubMatches(4)

        ' Only shift if:
        '   - no sheet prefix (same-sheet reference)
        '   - column is RELATIVE (dollarCol = "")
        Dim newToken As String
        If sheetPrefix = "" And dollarCol = "" Then
            Dim origColNum As Long
            Dim newColNum  As Long
            origColNum = ColLetterToNumber(colLetters)
            newColNum  = origColNum + colOffset
            If newColNum >= 1 Then
                Dim newColLetter As String
                newColLetter = ColNumberToLetter(newColNum)
                newToken = newColLetter & dollarRow & rowDigits
            Else
                ' Would go out of range — leave unchanged
                newToken = m.Value
            End If
        Else
            newToken = m.Value
        End If

        idxArr(idx) = m.FirstIndex   ' 0-based
        lenArr(idx) = m.Length
        repArr(idx) = newToken
        idx = idx + 1
    Next m

    ' Apply replacements in reverse order (highest index first)
    Dim j As Long
    For j = cnt - 1 To 0 Step -1
        result = Left(result, idxArr(j)) & repArr(j) & Mid(result, idxArr(j) + lenArr(j) + 1)
    Next j

    ShiftRelativeColumns = result
End Function

'=============================================================================
' FormulaLinksExternalSheet
' Returns True if the formula references any sheet other than wsName
'=============================================================================
Private Function FormulaLinksExternalSheet(formula As String, wsName As String) As Boolean
    Dim regex   As Object
    Dim matches As Object
    Dim m       As Object

    On Error Resume Next
    Set regex = CreateObject("VBScript.RegExp")
    On Error GoTo 0

    If regex Is Nothing Then
        ' Fallback: look for "!" in formula
        If InStr(formula, "!") = 0 Then
            FormulaLinksExternalSheet = False
            Exit Function
        End If
        Dim s1 As String: s1 = "'" & wsName & "'!"
        Dim s2 As String: s2 = wsName & "!"
        Dim hasSelf As Boolean
        hasSelf = (InStr(1, formula, s1, vbTextCompare) > 0) Or _
                  (InStr(1, formula, s2, vbTextCompare) > 0)
        FormulaLinksExternalSheet = Not hasSelf
        Exit Function
    End If

    ' Match 'Sheet Name'! or SheetName!
    regex.Pattern    = "'([^']+)'!|([A-Za-z_][A-Za-z0-9_\. ]*)!"
    regex.Global     = True
    regex.IgnoreCase = True

    If Not regex.Test(formula) Then
        FormulaLinksExternalSheet = False
        Exit Function
    End If

    Set matches = regex.Execute(formula)
    For Each m In matches
        Dim sName As String
        If m.SubMatches(0) <> "" Then
            sName = m.SubMatches(0)
        Else
            sName = m.SubMatches(1)
        End If
        If LCase(Trim(sName)) <> LCase(Trim(wsName)) Then
            FormulaLinksExternalSheet = True
            Exit Function
        End If
    Next m

    FormulaLinksExternalSheet = False
End Function

'=============================================================================
' COMMENT HANDLING
'=============================================================================

' Capture all comments from a column into a Collection of arrays {row, text}
' Call this BEFORE inserting a column so Excel's shift cannot lose them.
Private Function CaptureComments(ws As Worksheet, colNum As Long) As Object
    Dim col As Collection
    Set col = New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    Dim r As Long
    For r = 1 To lastRow
        If Not ws.Cells(r, colNum).Comment Is Nothing Then
            Dim entry(1) As Variant
            entry(0) = r
            entry(1) = ws.Cells(r, colNum).Comment.Text
            col.Add entry
        End If
    Next r
    Set CaptureComments = col
End Function

' Restore previously captured comments onto a column.
Private Sub RestoreComments(ws As Worksheet, colNum As Long, comments As Object)
    Dim entry As Variant
    Dim dstCell As Range
    For Each entry In comments
        Set dstCell = ws.Cells(entry(0), colNum)
        If Not dstCell.Comment Is Nothing Then dstCell.Comment.Delete
        dstCell.AddComment entry(1)
    Next entry
End Sub

Private Sub ClearColumnComments(ws As Worksheet, colNum As Long, lastRow As Long)
    Dim r As Long
    For r = 1 To lastRow
        If Not ws.Cells(r, colNum).Comment Is Nothing Then
            ws.Cells(r, colNum).Comment.Delete
        End If
    Next r
End Sub

'=============================================================================
' FIGURE (Shape) HANDLING
' Current col  : delete any shape whose left or right edge sits in this column
' Previous col : lock shapes to FreeFloating (hard-coded position)
'=============================================================================
Private Sub ProcessFigures(ws As Worksheet, currentCol As Long, previousCol As Long)
    Dim shp            As Shape
    Dim toDelete       As Collection
    Set toDelete = New Collection

    For Each shp In ws.Shapes
        Dim shpL As Long, shpR As Long
        shpL = GetShapeColumn(ws, shp, "left")
        shpR = GetShapeColumn(ws, shp, "right")

        If shpL = currentCol Or shpR = currentCol Then
            toDelete.Add shp
        ElseIf shpL = previousCol Or shpR = previousCol Then
            shp.Placement = xlFreeFloating
        End If
    Next shp

    Dim s As Shape
    For Each s In toDelete
        s.Delete
    Next s
End Sub

Private Function GetShapeColumn(ws As Worksheet, shp As Shape, side As String) As Long
    Dim pos    As Double
    Dim c      As Long
    Dim maxCol As Long

    pos = IIf(LCase(side) = "left", shp.Left, shp.Left + shp.Width)
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
' UTILITY
'=============================================================================
Private Function ColLetterToNumber(colLetter As String) As Long
    On Error Resume Next
    ColLetterToNumber = Range(colLetter & "1").Column
    On Error GoTo 0
End Function

Private Function ColNumberToLetter(colNum As Long) As String
    Dim result    As String
    Dim n         As Long
    Dim remainder As Long
    n = colNum
    result = ""
    Do While n > 0
        remainder = (n - 1) Mod 26
        result = Chr(65 + remainder) & result
        n = (n - 1) \ 26
    Loop
    ColNumberToLetter = result
End Function
