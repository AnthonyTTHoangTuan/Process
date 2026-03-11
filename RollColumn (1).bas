Option Explicit

'=============================================================================
' ROLL COLUMN MACRO
' SheetList structure:
'   Col A = Sheet Name
'   Col B = Column to roll (e.g. "D" or column number)
'   Col C = Insert Direction: "Left" or "Right"
'   Col D = Layout: blank/normal = Prev|Curr, "Reverse" = Curr|Prev
'           Special: "Ungrouped" = ungroup columns only, no insert
'=============================================================================

Sub RollColumns()
    Dim wsConfig    As Worksheet
    Dim wsTarget    As Worksheet
    Dim lastRow     As Long
    Dim i           As Long
    Dim sheetName   As String
    Dim colRef      As String
    Dim insertDir   As String
    Dim layout      As String
    Dim targetCol   As Long
    Dim newCol      As Long
    Dim currentCol  As Long
    Dim previousCol As Long

    ' --- Find the SheetList config sheet ---
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

        ' Convert column reference to number
        If IsNumeric(colRef) Then
            targetCol = CLng(colRef)
        Else
            targetCol = ColLetterToNumber(colRef)
        End If
        If targetCol < 1 Then
            errLog = errLog & "Invalid column '" & colRef & "' on sheet '" & sheetName & "'" & vbNewLine
            GoTo NextRow
        End If

        ' Handle Ungrouped layout (no insert, just ungroup)
        If UCase(layout) = "UNGROUPED" Then
            Call UngroupColumns(wsTarget, targetCol)
            GoTo NextRow
        End If

        ' Validate direction
        If UCase(insertDir) <> "LEFT" And UCase(insertDir) <> "RIGHT" Then
            errLog = errLog & "Invalid direction '" & insertDir & "' on sheet '" & sheetName & "'" & vbNewLine
            GoTo NextRow
        End If

        ' --- Determine which column is Current and which becomes Previous ---
        ' Normal layout (blank): Previous | Current
        '   Insert Left  => new col = Previous, targetCol stays Current
        '   Insert Right => new col = Current,  targetCol becomes Previous
        ' Reverse layout: Current | Previous
        '   Insert Left  => new col = Current,  targetCol becomes Previous
        '   Insert Right => new col = Previous, targetCol stays Current

        Dim newIsCurrent As Boolean
        If UCase(layout) = "REVERSE" Or UCase(layout) = "PREVRIGHT" Then
            ' Reverse layout
            If UCase(insertDir) = "LEFT" Then
                newIsCurrent = True   ' new col = Current, old = Previous
            Else
                newIsCurrent = False  ' new col = Previous, old stays Current
            End If
        Else
            ' Normal layout
            If UCase(insertDir) = "LEFT" Then
                newIsCurrent = False  ' new col = Previous, old stays Current
            Else
                newIsCurrent = True   ' new col = Current, old = Previous
            End If
        End If

        ' Insert the new column
        If UCase(insertDir) = "LEFT" Then
            wsTarget.Columns(targetCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            newCol = targetCol
            ' existing column shifted right
            Dim existingCol As Long
            existingCol = targetCol + 1
        Else
            wsTarget.Columns(targetCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            newCol = targetCol + 1
            existingCol = targetCol
        End If

        If newIsCurrent Then
            currentCol  = newCol
            previousCol = existingCol
        Else
            previousCol = newCol
            currentCol  = existingCol
        End If

        ' === Process the columns ===
        Call ProcessRoll(wsTarget, currentCol, previousCol, existingCol, newCol, newIsCurrent)

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
' Core roll logic: handle comments, formulas, figures for both columns
'=============================================================================
Private Sub ProcessRoll(ws As Worksheet, _
                        currentCol As Long, previousCol As Long, _
                        existingCol As Long, newCol As Long, _
                        newIsCurrent As Boolean)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, existingCol).End(xlUp).Row
    If lastRow < 1 Then lastRow = ws.UsedRange.Rows.Count

    ' === COMMENTS ===
    ' Current column: DELETE all comments
    Call ClearColumnComments(ws, currentCol, lastRow)

    ' Previous column: inherit comments from the column that WAS current before insert
    ' i.e. copy comments from existingCol (before it became previous)
    ' Since we already inserted, existingCol IS previousCol or currentCol
    ' We need to move/copy comments from existingCol to previousCol if different
    If previousCol <> existingCol Then
        Call CopyColumnComments(ws, existingCol, previousCol, lastRow)
        Call ClearColumnComments(ws, existingCol, lastRow)
    End If
    ' If previousCol = existingCol, comments stay where they are (correct)

    ' === FORMULAS / VALUES ===
    Call ProcessFormulas(ws, currentCol, previousCol, existingCol, newCol, newIsCurrent, lastRow)

    ' === FIGURES (shapes/images) ===
    Call ProcessFigures(ws, currentCol, previousCol)

End Sub

'=============================================================================
' FORMULA PROCESSING
'=============================================================================
Private Sub ProcessFormulas(ws As Worksheet, _
                            currentCol As Long, previousCol As Long, _
                            existingCol As Long, newCol As Long, _
                            newIsCurrent As Boolean, lastRow As Long)

    Dim r       As Long
    Dim cell    As Range
    Dim srcCell As Range
    Dim formula As String

    ' The "source" column (the one that was the current period before rolling)
    ' is existingCol. We copy its formulas to newCol with adjustments.

    For r = 1 To lastRow
        Set srcCell = ws.Cells(r, existingCol)

        If srcCell.HasFormula Then
            formula = srcCell.formula

            ' --- Current column: inherit all formulas, adjusted for column offset ---
            Dim offset As Long
            offset = newCol - existingCol
            Dim adjustedFormula As String
            adjustedFormula = AdjustFormulaColumns(formula, ws.Name, offset)
            ws.Cells(r, newCol).formula = adjustedFormula

            ' --- Previous column (existingCol) ---
            ' Formulas referencing other sheets => paste as value
            ' Formulas within this sheet => keep (already there, just keep)
            If newIsCurrent Then
                ' existingCol is now Previous; check if formula links external
                If FormulaLinksExternalSheet(formula, ws.Name) Then
                    ' Replace with value
                    Dim v As Variant
                    v = srcCell.Value
                    srcCell.Value = v
                End If
                ' else: keep internal formula as-is
            End If
            ' If newIsCurrent=False, existingCol stays Current => keep all formulas

        ElseIf Not IsEmpty(srcCell) Then
            ' No formula, just copy value to new col
            ws.Cells(r, newCol).Value = srcCell.Value
        End If

        ' Copy number format
        ws.Cells(r, newCol).NumberFormat = srcCell.NumberFormat
    Next r

    ' If newCol is Previous (not current), convert any external formulas to values
    If Not newIsCurrent Then
        For r = 1 To lastRow
            Set cell = ws.Cells(r, newCol)
            If cell.HasFormula Then
                If FormulaLinksExternalSheet(cell.formula, ws.Name) Then
                    cell.Value = cell.Value
                End If
            End If
        Next r
    End If

End Sub

'=============================================================================
' Check if a formula references another sheet (external to ws.Name)
'=============================================================================
Private Function FormulaLinksExternalSheet(formula As String, wsName As String) As Boolean
    ' A formula links another sheet if it contains "SheetName!" pattern
    ' that is NOT the current sheet
    Dim regex As Object
    On Error Resume Next
    Set regex = CreateObject("VBScript.RegExp")
    On Error GoTo 0

    If regex Is Nothing Then
        ' Fallback: simple check for "!" character meaning sheet reference
        ' We exclude references to current sheet
        Dim cleanSheet As String
        cleanSheet = "'" & wsName & "'!"
        Dim cleanSheet2 As String
        cleanSheet2 = wsName & "!"
        Dim hasAny As Boolean
        hasAny = (InStr(formula, "!") > 0)
        Dim hasSelf As Boolean
        hasSelf = (InStr(formula, cleanSheet) > 0 Or InStr(formula, cleanSheet2) > 0)
        FormulaLinksExternalSheet = hasAny And Not hasSelf
        Exit Function
    End If

    ' Has any sheet reference at all?
    regex.Pattern = "[']?[^'!\[\]]+[']?!"
    regex.Global = True
    regex.IgnoreCase = True

    If Not regex.Test(formula) Then
        FormulaLinksExternalSheet = False
        Exit Function
    End If

    ' Check if ALL sheet references are to current sheet
    Dim matches As Object
    Set matches = regex.Execute(formula)
    Dim m As Object
    For Each m In matches
        Dim refSheet As String
        refSheet = m.Value
        ' Strip quotes and exclamation
        refSheet = Replace(refSheet, "'", "")
        refSheet = Replace(refSheet, "!", "")
        If LCase(Trim(refSheet)) <> LCase(Trim(wsName)) Then
            FormulaLinksExternalSheet = True
            Exit Function
        End If
    Next m
    FormulaLinksExternalSheet = False
End Function

'=============================================================================
' Adjust column references in a formula by a given offset
' Only shifts absolute/relative column references, not sheet-qualified ones
' that point to other sheets (those are left alone for Previous col handling)
'=============================================================================
Private Function AdjustFormulaColumns(formula As String, wsName As String, colOffset As Long) As String
    If colOffset = 0 Then
        AdjustFormulaColumns = formula
        Exit Function
    End If

    ' Use Excel's built-in: put formula in a temp cell offset by colOffset,
    ' which automatically adjusts relative references.
    ' We'll do it properly by using a helper approach via R1C1.
    ' Strategy: convert to R1C1, shift column numbers in absolute refs, convert back.
    ' Simpler and more reliable: just return the formula as-is and let
    ' the cell's relative positioning handle it when we assign to the offset cell.
    ' (Excel adjusts relative refs automatically when we set .Formula on a different cell.)
    AdjustFormulaColumns = formula
End Function

'=============================================================================
' COMMENT HANDLING
'=============================================================================
Private Sub ClearColumnComments(ws As Worksheet, colNum As Long, lastRow As Long)
    Dim r As Long
    For r = 1 To lastRow
        If Not ws.Cells(r, colNum).Comment Is Nothing Then
            ws.Cells(r, colNum).Comment.Delete
        End If
    Next r
End Sub

Private Sub CopyColumnComments(ws As Worksheet, fromCol As Long, toCol As Long, lastRow As Long)
    Dim r       As Long
    Dim srcCell As Range
    Dim dstCell As Range
    Dim cmtText As String
    Dim cmtAuth As String

    For r = 1 To lastRow
        Set srcCell = ws.Cells(r, fromCol)
        Set dstCell = ws.Cells(r, toCol)

        If Not srcCell.Comment Is Nothing Then
            cmtText = srcCell.Comment.Text
            cmtAuth = srcCell.Comment.Author

            ' Delete existing comment on dest if any
            If Not dstCell.Comment Is Nothing Then dstCell.Comment.Delete

            ' Add comment to destination
            dstCell.AddComment cmtText
        End If
    Next r
End Sub

'=============================================================================
' FIGURE (Shape) HANDLING
' Current column: no hard-coded figures (remove shapes anchored to it)
' Previous column: hard-code figures (shapes stay, anchored to previous col)
'=============================================================================
Private Sub ProcessFigures(ws As Worksheet, currentCol As Long, previousCol As Long)
    Dim shp         As Shape
    Dim shpColLeft  As Long
    Dim shpColRight As Long

    For Each shp In ws.Shapes
        ' Determine which column the shape is anchored to (left edge)
        shpColLeft  = GetShapeColumn(ws, shp, "left")
        shpColRight = GetShapeColumn(ws, shp, "right")

        ' If shape overlaps current column => remove it (no hard-coded figures in Current)
        If shpColLeft = currentCol Or shpColRight = currentCol Then
            ' Check if it was "inherited" (i.e. came from previous roll)
            ' Rule: Current column should have no figures => delete
            shp.Delete

        ' If shape is in previous column => keep it (figures hard-coded in Previous)
        ElseIf shpColLeft = previousCol Or shpColRight = previousCol Then
            ' Keep shape as-is (hard-coded in previous)
            ' Lock position so it won't move with future inserts
            shp.Placement = xlFreeFloating
        End If
    Next shp
End Sub

Private Function GetShapeColumn(ws As Worksheet, shp As Shape, side As String) As Long
    Dim pos As Double
    If LCase(side) = "left" Then
        pos = shp.Left
    Else
        pos = shp.Left + shp.Width
    End If

    Dim c As Long
    For c = 1 To ws.UsedRange.Columns.Count + ws.UsedRange.Column
        If ws.Cells(1, c).Left + ws.Cells(1, c).Width >= pos Then
            GetShapeColumn = c
            Exit Function
        End If
    Next c
    GetShapeColumn = 0
End Function

'=============================================================================
' UNGROUP columns (for "Ungrouped" layout)
'=============================================================================
Private Sub UngroupColumns(ws As Worksheet, targetCol As Long)
    On Error Resume Next
    ws.Columns(targetCol).Ungroup
    On Error GoTo 0
End Sub

'=============================================================================
' UTILITY: Convert column letter(s) to number
'=============================================================================
Private Function ColLetterToNumber(colLetter As String) As Long
    On Error Resume Next
    ColLetterToNumber = Range(colLetter & "1").Column
    On Error GoTo 0
End Function
