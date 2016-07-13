VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} errorPropForm 
   Caption         =   "Error Propogation Tool"
   ClientHeight    =   7200
   ClientLeft      =   -81825
   ClientTop       =   -32235
   ClientWidth     =   7005
   OleObjectBlob   =   "errorPropForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "errorPropForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Const SizeCoefForMac = 1.2

Public errorOccurred As Boolean, ignoreErrors As Boolean, userCancelled As Boolean, programRunning As Boolean
Public doApprox As Boolean, doSensitivity As Boolean
Public multiRow As Boolean, averageError As Boolean
Public varForm As Boolean
Public errorValue As Variant
Public oldVar As Double, oldErr As Double
Public singleCol As Integer, doubleCol As Integer
Public screenUpdatingBehavior As Boolean, calculationBehavior As Integer, fullCalculate As Boolean

Private Function errorString(Val As Variant) As String
    Select Case Val
        Case CVErr(xlErrDiv0)
            errorString = "#DIV/0!"
        Case CVErr(xlErrNA)
            errorString = "#N/A"
        Case CVErr(xlErrName)
            errorString = "#NAME?"
        Case CVErr(xlErrNull)
            errorString = "#NULL!"
        Case CVErr(xlErrNum)
            errorString = "#NUM!"
        Case CVErr(xlErrRef)
            errorString = "#REF!"
        Case CVErr(xlErrValue)
            errorString = "#VALUE!"
    End Select
End Function

Private Sub prepareProgress()
    Dim blank As String
    
    
    blank = "                                     "
    With errorPropForm
'        If twoDetailLines = True Then
        #If Mac Then
            .Height = 380
        #Else
            .Height = 395 / SizeCoefForMac
        #End If
            .message2.Visible = True
'        Else
'                .Height = 115
'                .message2.Visible = False
'        End If
        .labelBlack.Caption = blank & "0%"
        With .labelWhite
            .Caption = blank & "0%"
            .Tag = .Width
            .Width = 1
        End With
        With .progressBar
            .Tag = .Width
            .Width = 1
            .Visible = True
        End With
    End With
    
    Call doErrorProp
    
End Sub

Private Sub progressUpdate(percent As Double, detail1 As String, detail2 As String)
    Dim blank As String, temp As Integer
    
    blank = "                                     "
    With errorPropForm
        .message1.Caption = detail1
        With .message2
            If .Visible = True Then .Caption = detail2
        End With
        .labelBlack.Caption = blank & Format(percent, "0%")
        With .labelWhite
            temp = percent * .Tag
            .Caption = blank & Format(percent, "0%")
            .Width = temp
        End With
        DoEvents
        With .progressBar
            .Width = percent * .Tag
        End With
    End With
    
    DoEvents

End Sub

Private Sub errorPropSub(targetCell As Range, functionCell As Range, varRange As Range, errorRange As Range)
    Dim ddx() As Double, varList() As Double, errorList() As Double
    Dim funcTemp(-1 To 1)  As Double, funcOld As Double, deltaX As Double, _
        sumBig As Double, sumSmall As Double, _
        tempSmall As Double, tempBig As Double, _
        tempPercentDiff As Double, sAnalysis(-1 To 1) As Double
    Dim numVars As Integer, numRows As Integer, rowLength As Integer, i As Integer, j As Integer, k As Integer, l As Integer
    Dim tmpRow As Integer, tmpCol As Integer, valSign As Integer
    Dim dConverge As Double, isConverged As Boolean, dOld As Double, dNew As Double, numIter As Integer
    Dim magicNumber As Double
    Dim thePrecedents() As Range
    Dim hasForm As Boolean, formStr As String

    magicNumber = 65535
       
    numVars = varRange.Columns.Count
    numRows = varRange.Rows.Count
    rowLength = numVars
    numVars = numRows * rowLength
    
    If Not functionCell.Cells(1, 1).HasFormula Then
        If doApprox Then
             If doSensitivity Then
                For k = 1 To numVars + singleCol
                    targetCell.Cells(1, k).Value = "No Function"
                Next k
            Else
                For k = 1 To singleCol
                    targetCell.Cells(1, k).Value = "No Function"
                Next k
            End If
        Else
            If doSensitivity Then
                For k = 1 To numVars + doubleCol
                    targetCell.Cells(1, k).Value = "No Function"
                Next k
            Else
                For k = 1 To doubleCol
                    targetCell.Cells(1, k).Value = "No Function"
                Next k
            End If
        End If
        Exit Sub
    End If
    
    If IsError(functionCell.Cells(1, 1).Value) Then
        If doApprox Then
             If doSensitivity Then
                For k = 1 To numVars + singleCol
                    targetCell.Cells(1, k).Value = errorString(functionCell.Cells(1, 1).Value)
                Next k
            Else
                For k = 1 To singleCol
                    targetCell.Cells(1, k).Value = errorString(functionCell.Cells(1, 1).Value)
                Next k
            End If
        Else
            If doSensitivity Then
                For k = 1 To numVars + doubleCol
                    targetCell.Cells(1, k).Value = errorString(functionCell.Cells(1, 1).Value)
                Next k
            Else
                For k = 1 To doubleCol
                    targetCell.Cells(1, k).Value = errorString(functionCell.Cells(1, 1).Value)
                Next k
            End If
        End If
        Exit Sub
    End If
    
    ReDim ddx(1 To numVars), varList(1 To numVars), errorList(1 To numVars)
    
    For i = 1 To numRows
        For j = 1 To rowLength
            On Error GoTo cleanUp
            varList((i - 1) * rowLength + j) = varRange.Cells(i, j).Value
            If errorRange.Rows.Count > 1 Then
                errorList((i - 1) * rowLength + j) = errorRange.Cells(i, j).Value
            Else
                errorList((i - 1) * rowLength + j) = errorRange.Cells(1, j).Value
            End If
            On Error GoTo 0
        Next j
    Next i
    
    If Not fullCalculate Then
        If functionCell.Worksheet.ProtectContents Or varRange.Worksheet.ProtectContents _
            Or errorRange.Worksheet.ProtectContents Or targetCell.Worksheet.ProtectContents Then
                fullCalculate = True
        Else
            fullCalculate = False
            thePrecedents = ArrangePrecedents(GetAllPrecedents(functionCell))
        End If
    End If
    
    sumSmall = 0
    sumBig = 0
    dConverge = 1E-20
    numIter = 20
    
    funcOld = functionCell.Cells(1, 1).Value
    
    If funcOld > 0 Then
        valSign = 1
    Else
        valSign = -1
    End If
    
    For i = 1 To numVars
        If multiRow Then
            If userCancelled Then
                Call helpText
                Application.Calculation = calculationBehavior
                Application.ScreenUpdating = screenUpdatingBehavior
                Exit Sub
            End If
            Call progressUpdate(i / numVars, "Variable " & i & " of " & numVars, "")
        End If
        tmpRow = (i - 1) _ rowLength + 1
        tmpCol = (i - 1) Mod rowLength + 1
        With varRange.Cells(tmpRow, tmpCol)
            hasForm = .HasFormula
            If hasForm Then formStr = .Formula
        End With
        If Not hasForm Or varForm Then
            If (doSensitivity Or doApprox) Then
                On Error Resume Next
                    dConverge = Abs(functionCell.Cells(1, 1).Value * 0.0000000001)
                On Error GoTo 0
                isConverged = False
                dOld = 1E+200
                dNew = 0
                On Error Resume Next
                    deltaX = varList(i) * 0.1
                On Error GoTo 0
                j = 1
                Do While (Not isConverged) And j < numIter
                    errorOccurred = False
                    For k = -1 To 1 Step 2
                        varRange.Cells(tmpRow, tmpCol).Value = varList(i) + k * deltaX
                        If fullCalculate Then
                            Application.Calculate
                        Else
                            Call RecalculateRanges(thePrecedents)
                            functionCell.Calculate
                        End If
                        If (IsError(functionCell.Cells(1, 1).Value)) Then
                            errorValue = functionCell.Cells(1, 1).Value
                            errorOccurred = True
                            If ignoreErrors Then
                                If hasForm And varForm Then
                                    varRange.Cells(tmpRow, tmpCol).Formula = formStr
                                Else
                                    varRange.Cells(tmpRow, tmpCol).Value = varList(i)
                                End If
                                If fullCalculate Then
                                    Application.Calculate
                                Else
                                    Call RecalculateRanges(thePrecedents)
                                    functionCell.Calculate
                                End If
                            Else
                                oldVar = varList(i)
                                oldErr = errorList(i)
                                varRange.Cells(tmpRow, tmpCol).Value = "IT WAS ME!"
                                If Not doApprox Then
                                    errorRange.Cells(tmpRow, tmpCol).Value = "AND I HELPED!"
                                End If
                                If fullCalculate Then
                                    Application.Calculate
                                Else
                                    Call RecalculateRanges(thePrecedents)
                                    functionCell.Calculate
                                End If
                                Exit Sub
                            End If
                        End If
                        funcTemp(k) = functionCell.Cells(1, 1).Value
                        If j = 1 Then
                            sAnalysis(k) = (funcTemp(k) - funcOld) / funcOld
                        End If
                    Next k
                    varRange.Cells(tmpRow, tmpCol).Value = varList(i)
                    If Not doApprox Then Exit Do
                    On Error Resume Next
                        dNew = (funcTemp(1) - funcTemp(-1)) * 0.5 / deltaX
                    On Error GoTo 0
                    If (dNew = 0) Or (Abs(dNew - dOld) <= dConverge) Then
                        isConverged = True
                    Else
                        On Error Resume Next
                            deltaX = deltaX * 0.1
                        On Error GoTo 0
                        dOld = dNew
                    End If
                    j = j + 1
                Loop
                If fullCalculate Then
                    Application.Calculate
                Else
                    Call RecalculateRanges(thePrecedents)
                    functionCell.Calculate
                End If
                If doSensitivity Then
'                    tempPercentDiff = 0
'                    On Error Resume Next
'                        tempPercentDiff = ((functionCell.Cells(1, 1).value + dNew) - functionCell.Cells(1, 1).value) / functionCell.Cells(1, 1).value
'                    On Error GoTo 0
                    If doApprox Or averageError Then
                        If errorOccurred Then
                            If hasForm And varForm Then
                                varRange.Cells(tmpRow, tmpCol).Formula = formStr
                            Else
                                varRange.Cells(tmpRow, tmpCol).Value = varList(i)
                            End If
                            Exit Sub
                        Else
                        With targetCell.Cells(tmpRow, tmpCol * 2 - 1 + singleCol)
                                .NumberFormat = "0.00000%"
                                .Value = sAnalysis(-1)
                        End With
                        With targetCell.Cells(tmpRow, tmpCol * 2 + singleCol)
                            .NumberFormat = "0.00000%"
                            .Value = sAnalysis(1)
                        End With
                        If targetCell.Cells(tmpRow, tmpCol * 2 - 1 + singleCol).Value = magicNumber Then
                            targetCell.Cells(tmpRow, tmpCol * 2 - 1 + singleCol).Value = 0
                            targetCell.Cells(tmpRow, tmpCol * 2 + singleCol).Value = 0
                        End If
                        ddx(i) = dNew
                        End If
                    Else
                        If errorOccurred And ignoreErrors Then
                            targetCell.Cells(tmpRow, tmpCol * 2 - 1 + doubleCol).Value = errorString(errorValue)
                            targetCell.Cells(tmpRow, tmpCol * 2 + doubleCol).Value = errorString(errorValue)
                        Else
                            With targetCell.Cells(tmpRow, tmpCol * 2 - 1 + doubleCol)
                                .NumberFormat = "0.00000%"
                                .Value = sAnalysis(-1)
                            End With
                            With targetCell.Cells(tmpRow, tmpCol * 2 + doubleCol)
                                .NumberFormat = "0.00000%"
                                .Value = sAnalysis(1)
                            End With
                            If targetCell.Cells(tmpRow, tmpCol * 2 - 1 + doubleCol).Value = magicNumber Then
                                targetCell.Cells(tmpRow, tmpCol * 2 - 1 + doubleCol).Value = 0
                                targetCell.Cells(tmpRow, tmpCol * 2 + doubleCol).Value = 0
                            End If
                        End If
                    End If
                ElseIf doApprox Then
                    ddx(i) = dNew
                End If
            End If
            If Not doApprox Then
                For k = -1 To 1 Step 2
                    On Error Resume Next
                        varRange.Cells(tmpRow, tmpCol).Value = varList(i) + k * errorList(i)
                    On Error GoTo 0
                    If fullCalculate Then
                        Application.Calculate
                    Else
                        Call RecalculateRanges(thePrecedents)
                        functionCell.Calculate
                    End If
                    If (IsError(functionCell.Cells(1, 1).Value)) Then
                        errorValue = functionCell.Cells(1, 1).Value
                        errorOccurred = True
                        If ignoreErrors Then
                            For l = 1 To doubleCol
                                targetCell.Cells(1, l).Value = errorString(functionCell.Cells(1, 1).Value)
                            Next l
                            If hasForm And varForm Then
                                varRange.Cells(tmpRow, tmpCol).Formula = formStr
                            Else
                                varRange.Cells(tmpRow, tmpCol).Value = varList(i)
                            End If
                        Else
                            oldVar = varList(i)
                            oldErr = errorList(i)
                            varRange.Cells(tmpRow, tmpCol).Value = "IT WAS ME!"
                            errorRange.Cells(tmpRow, tmpCol).Value = "AND I HELPED!"
                        End If
                        If fullCalculate Then
                            Application.Calculate
                        Else
                            Call RecalculateRanges(thePrecedents)
                            functionCell.Calculate
                        End If
                        Exit Sub
                    End If
                    If ((k = -1) And (functionCell.Cells(1, 1).Value > funcOld) _
                        Or ((k = 1) And (functionCell.Cells(1, 1).Value < funcOld))) Then
                            varRange.Cells(tmpRow, tmpCol).Value = varList(i) - k * errorList(i)
                            If fullCalculate Then
                                Application.Calculate
                            Else
                                Call RecalculateRanges(thePrecedents)
                                functionCell.Calculate
                            End If
                    End If
                    funcTemp(k) = functionCell.Cells(1, 1).Value
                    varRange.Cells(tmpRow, tmpCol).Value = varList(i)
                    If fullCalculate Then
                        Application.Calculate
                    Else
                        Call RecalculateRanges(thePrecedents)
                        functionCell.Calculate
                    End If
                Next k
                sumSmall = sumSmall + (funcTemp(-1) - funcOld)
                sumBig = sumBig + (funcTemp(1) - funcOld)
            End If
        ElseIf doSensitivity Then
            If doApprox Then
                targetCell.Cells(tmpRow, tmpCol * 2 - 1 + singleCol).Value = "Formula"
                targetCell.Cells(tmpRow, tmpCol * 2 + singleCol).Value = "Formula"
            Else
                targetCell.Cells(tmpRow, tmpCol * 2 - 1 + doubleCol).Value = "Formula"
                targetCell.Cells(tmpRow, tmpCol * 2 + doubleCol).Value = "Formula"
            End If
        End If
        If hasForm And varForm Then varRange.Cells(tmpRow, tmpCol).Formula = formStr
    Next i
    
    If doApprox Then
        sumSmall = 0
        For i = 1 To numVars
            On Error Resume Next
                sumSmall = sumSmall + (ddx(i) * errorList(i)) * (ddx(i) * errorList(i))
            On Error GoTo 0
        Next i
        'On Error Resume Next
            targetCell.Cells(1, 1).Value = functionCell.Cells(1, 1).Value
            targetCell.Cells(1, 2).Value = Sqr(sumSmall)
            tempPercentDiff = ((functionCell.Cells(1, 1).Value + targetCell.Cells(1, 2).Value) - functionCell.Cells(1, 1).Value) / functionCell.Cells(1, 1).Value
            targetCell.Cells(1, 3).NumberFormat = "0.00000%"
            targetCell.Cells(1, 3).Value = tempPercentDiff
            If targetCell.Cells(1, 2).Value = magicNumber Then
                targetCell.Cells(1, 2).Value = "#OVFL Use Exact Method"
                targetCell.Cells(1, 3).Value = "#OVFL Use Exact Method"
            End If
        'On Error GoTo 0
    Else
        If averageError Then
            targetCell.Cells(1, 1).Value = functionCell.Cells(1, 1).Value
            tempPercentDiff = (Abs(sumSmall) + Abs(sumBig)) * 0.5
            targetCell.Cells(1, 2).Value = tempPercentDiff
            tempPercentDiff = ((functionCell.Cells(1, 1).Value + tempPercentDiff) - functionCell.Cells(1, 1).Value) / functionCell.Cells(1, 1).Value
            targetCell.Cells(1, 3).NumberFormat = "0.00000%"
            targetCell.Cells(1, 3).Value = valSign * tempPercentDiff
        Else
            targetCell.Cells(1, 1).Value = functionCell.Cells(1, 1).Value
            targetCell.Cells(1, 2).Value = Abs(sumSmall)
            targetCell.Cells(1, 3).Value = Abs(sumBig)
            tempPercentDiff = ((functionCell.Cells(1, 1).Value + targetCell.Cells(1, 2).Value) - functionCell.Cells(1, 1).Value) / functionCell.Cells(1, 1).Value
            targetCell.Cells(1, 4).NumberFormat = "0.00000%"
            targetCell.Cells(1, 4).Value = valSign * tempPercentDiff
            tempPercentDiff = ((functionCell.Cells(1, 1).Value + targetCell.Cells(1, 3).Value) - functionCell.Cells(1, 1).Value) / functionCell.Cells(1, 1).Value
            targetCell.Cells(1, 5).NumberFormat = "0.00000%"
            targetCell.Cells(1, 5).Value = valSign * tempPercentDiff
        End If
    End If
    
    Exit Sub
    
cleanUp:
    Exit Sub

End Sub

Private Sub doErrorProp()
    Dim functionRange As Range, varRange As Range, errorRange As Range, targetRange As Range
    Dim varRangeTemp As Range, errorRangeTemp As Range
    Dim rowCount As Integer, varCount As Integer, funcCount As Integer
    Dim colOffset As Integer, rowOffset As Integer, varCol As Integer, varColNum As Integer
    Dim errJunk As Variant, errStr As String
    Dim numHeaderRows As Integer, numData As Double
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    
    singleCol = 3
    doubleCol = 5
    
    
    If functionRef.Value = "" Or varRef.Value = "" Or errorRef.Value = "" Or targetRef.Value = "" Then
        MsgBox "Empty selection"
        errorOccurred = True
        Call helpText
        Exit Sub
    End If
    
    Set functionRange = Range(functionRef.Value)
    Set varRange = Range(varRef.Value)
    Set errorRange = Range(errorRef.Value)
    Set targetRange = Range(targetRef.Value)
    
    doSensitivity = sensitivityCheckbox.Value
    doApprox = approxCheckbox.Value
    ignoreErrors = ignoreErrorCheck.Value
    averageError = averageErrorCheckbox.Value
    varForm = varFormCheckbox.Value
    
    rowCount = functionRange.Rows.Count
    funcCount = functionRange.Columns.Count
    varCount = varRange.Columns.Count
    multiRow = multiRowCheckbox.Value
    
    numData = rowCount * funcCount
    
    If functionRange.Areas.Count > 1 Or varRange.Areas.Count > 1 Or errorRange.Areas.Count > 1 Then
        MsgBox "Must select continuous ranges."
        errorOccurred = True
        Call helpText
        Exit Sub
    End If
    
    If (multiRow And (rowCount > 1 Or funcCount > 1)) Then
        MsgBox "Only one function may be selected for multi-row functions."
        errorOccurred = True
        Call helpText
        Exit Sub
    End If
    
    If multiRow Then
        rowCount = varRange.Rows.Count
    End If
    
    If (Not multiRow And rowCount <> varRange.Rows.Count) Then
        MsgBox "Function and variable row counts don't match."
        errorOccurred = True
        Call helpText
        Exit Sub
    ElseIf (varCount <> errorRange.Columns.Count) Then
        MsgBox "Error and variable column counts don't match."
        errorOccurred = True
        Call helpText
        Exit Sub
    ElseIf (errorRange.Rows.Count <> 1 And errorRange.Rows.Count <> rowCount) Then
        MsgBox "Error row count must be 1 or equal to variable row count."
        errorOccurred = True
        Call helpText
        Exit Sub
    ElseIf (multiRow And (funcCount > 1)) Then
        MsgBox "Only one function may be selected for multi-row functions."
        errorOccurred = True
        Call helpText
        Exit Sub
    End If
    
    screenUpdatingBehavior = Application.ScreenUpdating
    calculationBehavior = Application.Calculation
    fullCalculate = False
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    errorOccurred = False
    
    numHeaderRows = CInt(numHeaderRowsLabel.Caption)
    rowOffset = 1
    If numHeaderRows > rowOffset Then
        rowOffset = numHeaderRows
    End If
    
    varColNum = 2
    varCol = varCount * varColNum
    
    Call enableForm(False)
    userCancelled = False
    programRunning = True
    
    For j = 1 To funcCount
        For i = 1 To rowCount
            If userCancelled Then
                Call helpText
                Application.Calculation = calculationBehavior
                Application.ScreenUpdating = screenUpdatingBehavior
                Exit Sub
            End If
            Call progressUpdate(((j - 1) * rowCount + (i - 1)) / numData, "Function " & j, "Data Point " & i)
            If multiRow Then
                Set varRangeTemp = Range(varRange.Cells(i, 1), varRange.Cells(rowCount, varCount))
                If (rowCount = errorRange.Rows.Count) Then
                    Set errorRangeTemp = Range(errorRange.Cells(i, 1), errorRange.Cells(rowCount, varCount))
                Else
                    Set errorRangeTemp = Range(errorRange.Cells(1, 1), errorRange.Cells(1, varCount))
                End If
            Else
                Set varRangeTemp = Range(varRange.Cells(i, 1), varRange.Cells(i, varCount))
                If (rowCount = errorRange.Rows.Count) Then
                    Set errorRangeTemp = Range(errorRange.Cells(i, 1), errorRange.Cells(i, varCount))
                Else
                    Set errorRangeTemp = Range(errorRange.Cells(1, 1), errorRange.Cells(1, varCount))
                End If
            End If
            If doApprox Or averageError Then
                 If doSensitivity Then
                    colOffset = 1 + (j - 1) * (singleCol + varCol)
                    If i = 1 Then
                        For k = 1 To varCount
                            If numHeaderRows > 0 Then
                                For l = 1 To numHeaderRows
                                    On Error Resume Next
                                        targetRange.Cells(i + l - 1, colOffset + k * 2 + singleCol - 2).Value = varRange.Cells(i - l, k).Value & ": -10%"
                                        targetRange.Cells(i + l - 1, colOffset + k * 2 + singleCol - 1).Value = varRange.Cells(i - l, k).Value & ": +10%"
                                    On Error GoTo 0
                                Next l
                            Else
                                targetRange.Cells(i, colOffset + k * 2 + singleCol - 2).Value = "Variable " & k & ": -10%"
                                targetRange.Cells(i, colOffset + k * 2 + singleCol - 1).Value = "Variable " & k & ": +10%"
                            End If
                        Next k
                        With Range(targetRange.Cells(1, colOffset), targetRange.Cells(rowCount + rowOffset, colOffset + singleCol - 1))
                            .Font.Bold = True
                            .Font.Italic = False
                            .NumberFormat = "General"
                        End With
                        With Range(targetRange.Cells(1, colOffset + singleCol), targetRange.Cells(rowCount + rowOffset, 1 + (j) * (singleCol + varCol)))
                            .Font.Bold = False
                            .Font.Italic = True
                            .NumberFormat = "General"
                        End With
                    End If
                Else
                    colOffset = 1 + (j - 1) * singleCol
                End If
                If i = 1 Then
                    If numHeaderRows > 0 Then
                        For l = 1 To numHeaderRows
                            On Error Resume Next
                                targetRange.Cells(i + numHeaderRows - l, colOffset).Value = functionRange.Cells(i - l, j).Value & ": Value"
                                targetRange.Cells(i + numHeaderRows - l, colOffset + 1).Value = functionRange.Cells(i - l, j).Value & ": +/- Error"
                                targetRange.Cells(i + numHeaderRows - l, colOffset + 2).Value = functionRange.Cells(i - l, j).Value & ": % Diff"
                            On Error GoTo 0
                        Next l
                    Else
                        targetRange.Cells(i, colOffset).Value = "Function " & j & ": Value"
                        targetRange.Cells(i, colOffset + 1).Value = "Function " & j & ": +/- Error"
                        targetRange.Cells(i, colOffset + 2).Value = "Function " & j & ": % Diff"
                    End If
                    With Range(targetRange.Cells(1, colOffset), targetRange.Cells(rowCount + rowOffset, colOffset + singleCol - 1))
                        .Font.Bold = True
                        .Font.Italic = False
                        .NumberFormat = "General"
                    End With
                End If
            Else
                If doSensitivity Then
                    colOffset = 1 + (j - 1) * (doubleCol + varCol)
                    If i = 1 Then
                        For k = 1 To varCount
                            If numHeaderRows > 0 Then
                                For l = 1 To numHeaderRows
                                    On Error Resume Next
                                        targetRange.Cells(i + numHeaderRows - l, colOffset + k * 2 + doubleCol - 2).Value = varRange.Cells(i - l, k).Value & ": -10%"
                                        targetRange.Cells(i + numHeaderRows - l, colOffset + k * 2 + doubleCol - 1).Value = varRange.Cells(i - l, k).Value & ": +10%"
                                    On Error GoTo 0
                                Next l
                            Else
                                targetRange.Cells(i, colOffset + k * 2 + doubleCol - 2).Value = "Variable " & k & ": -10%"
                                targetRange.Cells(i, colOffset + k * 2 + doubleCol - 1).Value = "Variable " & k & ": +10%"
                            End If
                        Next k
                        With Range(targetRange.Cells(1, colOffset), targetRange.Cells(rowCount + rowOffset, colOffset + doubleCol - 1))
                            .Font.Bold = True
                            .Font.Italic = False
                            .NumberFormat = "General"
                        End With
                        With Range(targetRange.Cells(1, colOffset + doubleCol), targetRange.Cells(rowCount + rowOffset, 1 + (j) * (doubleCol + varCol)))
                            .Font.Bold = False
                            .Font.Italic = True
                            .NumberFormat = "General"
                        End With
                    End If
                Else
                    colOffset = 1 + (j - 1) * doubleCol
                End If
                If i = 1 Then
                    If numHeaderRows > 0 Then
                        For l = 1 To numHeaderRows
                            On Error Resume Next
                                targetRange.Cells(i + numHeaderRows - l, colOffset).Value = functionRange.Cells(i - l, j).Value & ": Value"
                                targetRange.Cells(i + numHeaderRows - l, colOffset + 1).Value = functionRange.Cells(i - l, j).Value & ": - Error"
                                targetRange.Cells(i + numHeaderRows - l, colOffset + 2).Value = functionRange.Cells(i - l, j).Value & ": + Error"
                                targetRange.Cells(i + numHeaderRows - l, colOffset + 3).Value = functionRange.Cells(i - l, j).Value & ": - % Diff"
                                targetRange.Cells(i + numHeaderRows - l, colOffset + 4).Value = functionRange.Cells(i - l, j).Value & ": + % Diff"
                            On Error GoTo 0
                        Next l
                    Else
                        targetRange.Cells(i, colOffset).Value = "Function " & j & ": Value"
                        targetRange.Cells(i, colOffset + 1).Value = "Function " & j & ": - Error"
                        targetRange.Cells(i, colOffset + 2).Value = "Function " & j & ": + Error"
                        targetRange.Cells(i, colOffset + 3).Value = "Function " & j & ": - % Diff"
                        targetRange.Cells(i, colOffset + 4).Value = "Function " & j & ": + % Diff"
                    End If
                    With Range(targetRange.Cells(1, colOffset), targetRange.Cells(rowCount + rowOffset, colOffset + doubleCol - 1))
                        .Font.Bold = True
                        .Font.Italic = False
                        .NumberFormat = "General"
                    End With
                End If
            End If
            Application.Calculate
            Call errorPropSub(targetRange.Cells(i + rowOffset, colOffset), functionRange.Cells(i, j), varRangeTemp, errorRangeTemp)
            If errorOccurred And Not ignoreErrors Then
                targetRange.Cells(i + rowOffset, colOffset).Value = oldVar
                If Not doApprox Then
                    targetRange.Cells(i + rowOffset, colOffset + 1).Value = oldErr
                End If
                errJunk = MsgBox("A " & errorString(errorValue) & " error occurred! Fix the cell taking responsibility and try again." & Chr(10) & Chr(10) & "This happened because a variable has been changed in order to check its influence on a function, and has left the domain of the function (or an intermediate function), resulted in division by zero, etc." & Chr(10) & Chr(10) & "Find and prevent these issues or run again ignoring errors.", vbCritical, "Error")
                
                Call enableForm(True)
                errorPropForm.Hide
                Exit Sub
            End If
            If multiRow Then
                Exit For
            End If
        Next i
        If multiRow Then
            Exit For
        End If
    Next j
    
    If doApprox Then
         If doSensitivity Then
            colOffset = funcCount * (singleCol + varCount)
        Else
            colOffset = funcCount * singleCol
        End If
    Else
        If doSensitivity Then
            colOffset = funcCount * (doubleCol + varCount)
        Else
            colOffset = doubleCol * funcCount
        End If
    End If
    
    Worksheets(targetRange.Parent.Name).Range(targetRange.Cells(2 - targetRange.Cells(1, 1).Row, 1), targetRange.Cells(rowOffset + rowCount, colOffset)).Columns.AutoFit
    
    Application.Calculation = calculationBehavior
    Application.ScreenUpdating = screenUpdatingBehavior
    
    Application.Goto (targetRange.Cells(1, 1))
    
    Call helpText
    errorPropForm.Hide
End Sub


Private Sub clearForm()
    functionRef.Value = ""
    varRef.Value = ""
    errorRef.Value = ""
    targetRef.Value = ""
    approxCheckbox.Value = False
    sensitivityCheckbox.Value = False
    ignoreErrorCheck.Value = True
    multiRowCheckbox.Value = False
    averageErrorCheckbox.Value = False
    varFormCheckbox.Value = False
    numHeaderRowsLabel.Caption = "0"
End Sub

Private Sub helpText()
    fLabel.ControlTipText = "Select a range of one or more functions for one or more data points to calculate error for. Each column is a different function, and each row is for a different data point."
    functionRef.ControlTipText = "Select a range of one or more functions for one or more data points to calculate error for. Each column is a different function, and each row is for a different data point."
    vLabel.ControlTipText = "Select a range of variables that influence the selected function(s). Each row in the range is for a different data point."
    varRef.ControlTipText = "Select a range of variables that influence the selected function(s). Each row in the range is for a different data point."
    eLabel.ControlTipText = "Select the range of error values for selected variables. # of columns must match variable range. If unique error values, # of rows matches variables, 1 row otherwise."
    errorRef.ControlTipText = "Select the range of error values for selected variables. # of columns must match variable range. If unique error values, # of rows matches variables, 1 row otherwise."
    oLabel.ControlTipText = "Select the single cell to be the upper-left corner of all output data. This utility will overwrite anything that gets in itfs way, so be careful"
    targetRef.ControlTipText = "Select the single cell to be the upper-left corner of all output data. This utility will overwrite anything that gets in itfs way, so be careful"
    sensitivityCheckbox.ControlTipText = "Perform a sensitivity analysis for all selected variables for all selected functions. The values produced indicate the function change with the increase of the variable by 1."
    approxCheckbox.ControlTipText = "Use the approximate error propagation method to obtain a single value for +/- error. (Sum of product of squares of variable error and partial derivative)"
    ignoreErrorCheck.ControlTipText = "Ignore numerical errors like #DIV/0!, #NUM!, #VALUE!, or any other errors that appear in the spreadsheet while calculations are running."
    headingLabel.ControlTipText = "Select the numbers of heading rows that exist above the first row of selected functions/variables. If zero, then default headings will be produced for output."
    numHeaderRowsButton.ControlTipText = "Select the numbers of heading rows that exist above the first row of selected functions/variables. If zero, then default headings will be produced for output."
    numHeaderRowsLabel.ControlTipText = "Select the numbers of heading rows that exist above the first row of selected functions/variables. If zero, then default headings will be produced for output."
    multiRowCheckbox.ControlTipText = "With this option selected, you can evaluate the error associated with types of functions like INTERCEPT(), SUMX2PY2(), STDEV(), etc. that take many arguments. "
    averageErrorCheckbox.ControlTipText = "Return a single error value for each function; the average of the upper and lower error from the exact method."
    varFormCheckbox.ControlTipText = "Variable cells with formulas are ignored; only constant values are used. Check here to include error from ALL variable cells."

    #If Mac Then
        errorPropForm.Height = 295
        labelWhite.Width = 314
        progressBar.Width = 314
    #Else
        errorPropForm.Height = 310 / SizeCoefForMac
        labelWhite.Width = 314 / SizeCoefForMac
        progressBar.Width = 314 / SizeCoefForMac
    #End If
    programRunning = False
    userCancelled = False
    Call enableForm(True)
End Sub

Private Sub enableForm(bool As Boolean)
    Dim controlItem As Variant
    For Each controlItem In errorPropForm.Controls
        controlItem.Enabled = bool
    Next
    message1.Enabled = True
    message2.Enabled = True
    progressBar.Enabled = True
    labelWhite.Enabled = True
    labelBlack.Enabled = True
    CommandButton2.Enabled = True
End Sub

Private Sub demoButton_Click()
    Dim demoWS As Worksheet, newWS As Worksheet
    Dim activeWB As Workbook, addOnWB As Workbook
    
    Set activeWB = ActiveWorkbook
    Set addOnWB = Workbooks("TimsTools.xlam")
    Set demoWS = addOnWB.Worksheets("ErrPropDemo")
    
    demoWS.Copy after:=activeWB.Worksheets(activeWB.Worksheets.Count)
    
    Call CommandButton2_Click
    
End Sub

Private Sub labelBlack_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        userCancelled = True
        errorPropForm.Hide
    End If
End Sub

Private Sub approxCheckbox_Click()
    If approxCheckbox.Value = True Then
        averageErrorCheckbox.Value = False
    End If
End Sub

Private Sub averageErrorCheckbox_Click()
    If averageErrorCheckbox.Value = True Then
        approxCheckbox.Value = False
    End If
End Sub

Private Sub clearButton_Click()
    Call clearForm
    Call helpText
End Sub

 

Private Sub CommandButton1_Click()
    Call prepareProgress
    Call helpText
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    If Not errorOccurred Then errorPropForm.Hide
End Sub

Private Sub CommandButton2_Click()
    userCancelled = True
    errorPropForm.Hide
    Unload Me
End Sub

Private Sub errorRef_Change()
    errorPropForm.Caption = "Error Propogation Tool"
End Sub

Private Sub errorRef_DropButtonClick()
    errorPropForm.Caption = "Select Variable Errors Range"
End Sub

Private Sub functionRef_Change()
    errorPropForm.Caption = "Error Propogation Tool"
End Sub

Private Sub functionRef_DropButtonClick()
    errorPropForm.Caption = "Select Function(s) Range"
End Sub


Private Sub numHeaderRowsButton_SpinDown()
    With numHeaderRowsLabel
        If CInt(.Caption) > 0 Then
            .Caption = CInt(.Caption) - 1
        End If
    End With
End Sub

Private Sub numHeaderRowsButton_SpinUp()
    With numHeaderRowsLabel
        .Caption = CInt(.Caption) + 1
    End With
End Sub

Private Sub numHeaderRowsLabel_Click()

End Sub

Private Sub UserForm_Initialize()
    Call AdjustSizeForWin(Me)
    Call helpText
End Sub

Private Sub varRef_Change()
    errorPropForm.Caption = "Error Propogation Tool"
End Sub

Private Sub varRef_DropButtonClick()
    errorPropForm.Caption = "Select Variable(s) Range"
End Sub

Private Sub targetRef_Change()
    errorPropForm.Caption = "Error Propogation Tool"
End Sub

Private Sub targetRef_DropButtonClick()
    errorPropForm.Caption = "Select Output Cell"
End Sub


' This module contains functions for identifying and recalculating a cell's precedent cells, even if those precedents are on several different worksheets.


' Instead of letting Excel recalculate the whole sheet, we only recalculate the cells we know we need.
' This technique is useful for speeding up scripts that run in large, slow spreadsheets.
Sub TestArrangePrecedents()
  Dim G3Precedents() As Range
  
  ' Restore A3 to a value of 3 before beginning the test.
  Call ResetSpreadsheetFormulasForDemo
  
  ' Let's pretend we're working with a very complicated spreadsheet that takes a long time to recalculate.
  ' For now we'll prevent automatic recalculation.
  Application.Calculation = xlCalculationManual
  
  ' This provides an array of precedents to G3 in the correct calculation order.
  ' The advantage is that we can use and reuse this to avoid recalculating the ENTIRE workbook if we only need to to update G3.
  G3Precedents = ArrangePrecedents(GetAllPrecedents(Worksheets("Sheet1").Range("G3")))
  
  ' This is where the script would do the heavy lifting because the spreadsheet won't stop and recalculate when anything changes.
  
  ' To demonstrate, change the value of one cell.
  Worksheets("Sheet2").Range("A2").Value = 2
  Debug.Print vbCrLf & "We just changed A3's value to 4."
  ' Try recalculating G3 and show the result.
  Call Worksheets("Sheet1").Range("G3").Calculate
  Debug.Print "  G3 = " & Worksheets("Sheet1").Range("G3").Value & ", but the correct answer is 5."
  
  ' Try again, but this time recalculate G3's precedents first.
  Debug.Print "Recalculate G3's precedents and try again..."
  Call RecalculateRanges(G3Precedents) '<-- This makes the difference.
  Call Worksheets("Sheet1").Range("G3").Calculate
  Debug.Print "  ...G3 = " & Worksheets("Sheet1").Range("G3").Value & ", which is correct!"
  
  ' Again, change the value of A3 and recalculate G3 along with its precedents.
  ' Note that we're reusing G3Precedents that we calculated earlier; no need to rediscover electricity, uh, or something.
  Worksheets("Sheet2").Range("A2").Value = 8
  Debug.Print "Now change A3's value to 10.  Does this change propagate to G3?"
  Call RecalculateRanges(G3Precedents)
  Call Worksheets("Sheet1").Range("G3").Calculate
  Debug.Print "  Yes, G3 = " & Worksheets("Sheet1").Range("G3").Value & "."
  
  ' Restore A3 to a value of 3 now that the demonstration is over.
  Worksheets("Sheet2").Range("A2").Value = 1
  ' Now that the script is done, restore automatic calculation behavior.  Allow things to run slowly henceforth.
  Application.Calculation = xlCalculationAutomatic
End Sub


' Set up the demo on Sheet1 and Sheet2.
Sub ResetSpreadsheetFormulasForDemo()
  ' Set up the spreadsheets.
  Worksheets("Sheet2").Range("A2").Value = 1
  Worksheets("Sheet1").Range("A3").Value = "=2+Sheet2!A2"
  Worksheets("Sheet1").Range("B3").Formula = "=A3-1"
  Worksheets("Sheet1").Range("C3").Formula = "=B3-1"
  ' The only difference is that B3 and C3 are transposed.
  Worksheets("Sheet1").Range("F3").Formula = "=B3+C3"
  Worksheets("Sheet1").Range("G3").Formula = "=C3+B3"
End Sub


'won't navigate through precedents in closed workbooks
'won't navigate through precedents in protected worksheets
'won't identify precedents on hidden sheets
Public Function GetAllPrecedents(ByRef rngToCheck As Range) As Object
  Dim dicAllPrecedents As Object
  Dim strKey As String
  Dim screenUpdatingBehavior As Boolean
  
  #If Mac Then
        Set dicAllPrecedents = New Dictionary
  #Else
        Set dicAllPrecedents = CreateObject("Scripting.Dictionary")
  #End If
  ' Is the application updating the screen right now?
  'ScreenUpdatingBehavior = Application.ScreenUpdating
  ' The application should not update the screen because this script is busy.
  'Application.ScreenUpdating = False
  ' Initiate the search for precedent cells.  The result is stored in dicAllPrecedents and transferred to the function return value.
  Call GetPrecedents(rngToCheck, dicAllPrecedents, 1)
  Set GetAllPrecedents = dicAllPrecedents
  ' Restore screen updating behavior.
  'Application.ScreenUpdating = ScreenUpdatingBehavior
End Function


' This converts a range (of possibly more than one cell) to a series of calls to GetCellPrecedents.
Private Sub GetPrecedents(ByRef rngToCheck As Range, ByRef dicAllPrecedents As Object, ByVal lngLevel As Long)
  Dim rngCell As Range
  Dim rngFormulas As Range
  ' Don't check further if the cell's worksheet is protected.
  ' Note the misnamed property ProtectContents that should be named something like ContentsAreProtected.
  If Not rngToCheck.Worksheet.ProtectContents Then
    ' Is there more than one cell in this range?
    If rngToCheck.Cells.CountLarge > 1 Then
      On Error Resume Next
      ' Only check the cells that have formulas in them.
      Set rngFormulas = rngToCheck.SpecialCells(xlCellTypeFormulas)
      On Error GoTo 0
    Else
      ' This must have been only one cell (not a range of many cells).
      ' Does this cell contain a formula?
      If rngToCheck.HasFormula Then
        ' This has a formula, so we want to check for its precedents.
        Set rngFormulas = rngToCheck
      End If
    End If
    ' At this point rngFormulas either contains nothing or contains one or more cells with formulas.
    ' Did we find anything?
    If Not rngFormulas Is Nothing Then
      ' Iterate once for each cell with a formula.
      For Each rngCell In rngFormulas.Cells


'        ' Start the whole process for this cell.  The Colin version does not evaluate the levels correctly.
'        Call GetCellPrecedentsColin(rngCell, dicAllPrecedents, lngLevel)
        ' Start the whole process for this cell.  The MMM version arrives at the correct result.
        If rngCell.HasFormula Then
            Call GetCellPrecedentsMMM(rngCell, dicAllPrecedents, lngLevel)
        End If
      
      If fullCalculate Then Exit For
      Next rngCell
      ' We're done with these cells (though we may come across this worksheet again).
      rngFormulas.Worksheet.ClearArrows
    End If
  End If
End Sub


' Compiles a list of precedents to a single cell.  The result is stored in dicAllPrecedents.
Private Sub GetCellPrecedentsMMM(ByRef rngCell As Range, ByRef dicAllPrecedents As Object, ByVal lngLevel As Long)
  Dim lngArrow As Long
  Dim lngLink As Long
  Dim ContinueLookingForArrows As Boolean
  Dim strPrecedentAddress As String
  Dim rngPrecedentRange As Range
  Dim OldSelection As Variant
  
  ' The NavigateArrow method takes numerical "arrow" and "link" parameters.
  ' Loop through the arrows.  Then loop through the links.  Each valid arrow must have at least one valid link.
  ' The Excel object model doesn't provide a function that returns the number of valid arrows.  That would be too easy.
  ' It doesn't provide a function to return the number of valid links for a given arrow.  That also would be too easy.
  ' When you exceed the number of valid links for a given arrow, you get an error message. (And there may be more valid arrows to look through.)
  ' When you exceed the number of valid arrows, you get a reference back to the cell you are searching. (At which point there are no more valid arrows or links.)
  
  ' We haven't checked any arrows yet.
  lngArrow = 0
  ' Outer Do loop - Loop for arrows
  Do
    ' We're checking another arrow...
    lngArrow = lngArrow + 1
    ' Assume that we should not move to the next arrow after this loop.
    ' (This will change later if at least one valid result is found.)
    ContinueLookingForArrows = False
    ' ...but we haven't checked any links yet.
    lngLink = 0
    ' Inner Do loop - Loop for links
    Do
      ' We're checking the next link for this particular arrow.
      lngLink = lngLink + 1
      ' For some reason Excel skips some of the precedents if we don't do this again and again in each iteration.
      rngCell.ShowPrecedents
      ' Rather than tell us how far to search, Excel expects us to generate an error message and catch it when it arrives.
      On Error Resume Next
      
      
      
      Set OldSelection = Application.Selection
      ' Attempt to find a precedent with this arrow and link number.
      Set rngPrecedentRange = rngCell.NavigateArrow(True, lngArrow, lngLink)
      If rngPrecedentRange.Precedents.Count > 50 Then
        fullCalculate = True
        Exit Sub
      End If
      
      If Not (OldSelection Is Nothing) Then
        If IsObject(OldSelection) Then
          If TypeName(OldSelection) = "Range" Then
            'Call OldSelection.Parent.Activate
            Call OldSelection.Select
          End If
        End If
      End If
      
    ' If this generated an error, the arrow/link combination must not refer to a valid precedent cell.
    ' In that case, stop looking for more links to go with this arrow.
    If Err.Number <> 0 Then
      On Error GoTo 0
      Exit Do
    End If
    On Error GoTo 0
    ' If we got this far, there was no error, which means that the arrow/link combination produces a valid cell reference
    ' (but not necessarily a reference to a precedent).
    strPrecedentAddress = rngPrecedentRange.Address(False, False, xlA1, True)
    ' The only other thing that can go wrong is that the arrow/link combination references rngCell itself.
    ' This reveals an inconsistency in the NavigateArrow function's behavior.
    ' If one thing goes wrong, you get an error.  If something else goes wrong, you get a result that doesn't tell you anything.
    ' If this "precedent" references the cell we were searching, then we've exhausted the possible arrow values.
    If strPrecedentAddress = rngCell.Address(False, False, xlA1, True) Then
      ' This exits the inner Do loop.
      Exit Do
    Else
      ' The arrow/link combination produced a useful result, so there may be more valid arrows after this one.
      ' When this inner Do loop (for links) finishes, continue iterating in the outer Do loop (for more arrows).
      ContinueLookingForArrows = True
      ' If this is already in the list of precedents, its level (and its precedents' levels) may need to be updated.
      If dicAllPrecedents.Exists(strPrecedentAddress) Then
        ' Does the dictionary list a shallower level?  (If so, update it.  If not, leave it alone.)
        If dicAllPrecedents.Item(strPrecedentAddress) < lngLevel Then
          ' Replace the existing level with the updated, deeper level.
          dicAllPrecedents.Item(strPrecedentAddress) = lngLevel
          ' The precedent cell's own precedent cells also need to be updated.
          Call GetPrecedents(rngPrecedentRange, dicAllPrecedents, lngLevel + 1)
        End If
      ElseIf rngPrecedentRange.HasFormula Then
        ' This item must not be in the dictionary.
        ' Add this item and its precedents as usual.
        Call dicAllPrecedents.Add(strPrecedentAddress, lngLevel)
        Call GetPrecedents(rngPrecedentRange, dicAllPrecedents, lngLevel + 1)
      End If
    End If
    Loop
  ' If ContinueLookingForArrows is False, that marks the end of this branch of recursive calls.
  Loop While ContinueLookingForArrows
End Sub


' Recalculates an array of Range objects in order from first to last.
Sub RecalculateRanges(RangeArray() As Range)
  Dim i As Long
  For i = LBound(RangeArray) To UBound(RangeArray)
    Call RangeArray(i).Calculate
  Next
End Sub


' Returns an array of Range objects that are precedents to another Range.
' Each Range is a single cell.  Proper recalculation of the parent Range (which these were calculated from) proceeds from first to last.
Function ArrangePrecedents(PrecedentsDictionary As Object) As Range()
  Dim i As Long
  Dim j As Long
  Dim ItemsArray() As Variant
  Dim KeysArray() As String 'Variant
  Dim MaxLevel As Long
  Dim ThisIndex As Long
  Dim Result() As Range
  
  If fullCalculate Then
    ArrangePrecedents = Result
    Exit Function
  End If
  
  ' Retrieve the dictionary data as arrays.
  ItemsArray = PrecedentsDictionary.Items()
  KeysArray = PrecedentsDictionary.Keys()
  ' Find the maximum level of any item in the array.
  MaxLevel = 0
  ' Loop once for each item and search for the highest value.  That will be MaxLevel.
  
  For i = LBound(ItemsArray) To UBound(ItemsArray)
    If ItemsArray(i) > MaxLevel Then
      MaxLevel = ItemsArray(i)
    End If
  Next
  ' The result must have the same size as the dictionary keys.
  ReDim Result(LBound(KeysArray) To UBound(KeysArray))
  ' We haven't chosen an index, but this will be incremented before it is first used.
  ' As we keep adding ranges to Result, we need to keep track of which element we used last.
  ThisIndex = LBound(Result) - 1
  ' Start with the deepest level MaxLevel and work down to 1.
  For i = MaxLevel To 1 Step -1
    ' Loop once for each item that we might add to the array.
    For j = LBound(ItemsArray) To UBound(ItemsArray)
      ' If this item has the level we're looking for, we need to add it to the result.
      If ItemsArray(j) = i Then
        ' Use the next location in Result and add the Range corresponding to the address string.
        ThisIndex = ThisIndex + 1
        Set Result(ThisIndex) = Application.Evaluate(KeysArray(j))
      End If
    Next
  Next
  ArrangePrecedents = Result
End Function


'Private Sub CommandButton1_Click()
'    Dim functionRange As Range, varRange As Range, errorRange As Range, targetRange As Range
'    Dim varRangeTemp As Range, errorRangeTemp As Range
'    Dim rowCount As Integer, varCount As Integer, funcCount As Integer
'    Dim colOffset As Integer, rowOffset As Integer, varCol As Integer, varColNum As Integer
'    Dim errJunk As Variant, errStr As String
'    Dim numHeaderRows As Integer
'    Dim i As Integer, j As Integer, k As Integer, l As Integer
'
'    singleCol = 2
'    doubleCol = 3
'
'
'    If functionRef.Value = "" Or varRef.Value = "" Or errorRef.Value = "" Or targetRef.Value = "" Then
'        MsgBox "Empty selection"
'        Exit Sub
'    End If
'
'    Set functionRange = Range(functionRef.Value)
'    Set varRange = Range(varRef.Value)
'    Set errorRange = Range(errorRef.Value)
'    Set targetRange = Range(targetRef.Value)
'
'    doSensitivity = sensitivityCheckbox.Value
'    doApprox = approxCheckbox.Value
'    ignoreErrors = ignoreErrorCheck.Value
'    averageError = averageErrorCheckbox.Value
'
'    rowCount = functionRange.Rows.Count
'    funcCount = functionRange.Columns.Count
'    varCount = varRange.Columns.Count
'    multiRow = multiRowCheckbox.Value
'
'    If (multiRow And (rowCount > 1 Or funcCount > 1)) Then
'        MsgBox "Only one function may be selected for multi-row functions."
'        Exit Sub
'    End If
'
'    If multiRow Then
'        rowCount = varRange.Rows.Count
'    End If
'
'    If (Not multiRow And rowCount <> varRange.Rows.Count) Then
'        MsgBox "Function and variable row counts don't match."
'        Exit Sub
'    ElseIf (varCount <> errorRange.Columns.Count) Then
'        MsgBox "Error and variable column counts don't match."
'        Exit Sub
'    ElseIf (errorRange.Rows.Count <> 1 And errorRange.Rows.Count <> rowCount) Then
'        MsgBox "Error row count must be 1 or equal to variable row count."
'        Exit Sub
'    ElseIf (multiRow And (funcCount > 1)) Then
'        MsgBox "Only one function may be selected for multi-row functions."
'        Exit Sub
'    End If
'
'    Application.ScreenUpdating = False
'
'    errorOccurred = False
'
'    numHeaderRows = CInt(numHeaderRowsLabel.Caption)
'    rowOffset = 1
'    If numHeaderRows > rowOffset Then
'        rowOffset = numHeaderRows
'    End If
'
'    varColNum = 2
'    varCol = varCount * varColNum
'
'    For j = 1 To funcCount
'        For i = 1 To rowCount
'            If multiRow Then
'                Set varRangeTemp = Range(varRange.Cells(i, 1), varRange.Cells(rowCount, varCount))
'                If (rowCount = errorRange.Rows.Count) Then
'                    Set errorRangeTemp = Range(errorRange.Cells(i, 1), errorRange.Cells(rowCount, varCount))
'                Else
'                    Set errorRangeTemp = Range(errorRange.Cells(1, 1), errorRange.Cells(1, varCount))
'                End If
'            Else
'                Set varRangeTemp = Range(varRange.Cells(i, 1), varRange.Cells(i, varCount))
'                If (rowCount = errorRange.Rows.Count) Then
'                    Set errorRangeTemp = Range(errorRange.Cells(i, 1), errorRange.Cells(i, varCount))
'                Else
'                    Set errorRangeTemp = Range(errorRange.Cells(1, 1), errorRange.Cells(1, varCount))
'                End If
'            End If
'            If doApprox Or averageError Then
'                 If doSensitivity Then
'                    colOffset = 1 + (j - 1) * (singleCol + varCol)
'                    If i = 1 Then
'                        For k = 1 To varCount
'                            If numHeaderRows > 0 Then
'                                For l = 1 To numHeaderRows
'                                    On Error Resume Next
'                                        targetRange.Cells(i + l - 1, colOffset + k * 2).Value = varRange.Cells(i - l, k).Value
'                                        targetRange.Cells(i + l - 1, colOffset + k * 2 + 1).Value = varRange.Cells(i - l, k).Value & ": % Diff"
'                                    On Error GoTo 0
'                                Next l
'                            Else
'                                targetRange.Cells(i, colOffset + k * 2).Value = "Variable " & k
'                                targetRange.Cells(i, colOffset + k * 2 + 1).Value = "Variable " & k & ": % Diff"
'                            End If
'                        Next k
'                        With Range(targetRange.Cells(1, colOffset), targetRange.Cells(rowCount + rowOffset, colOffset + 1))
'                            .Font.Bold = True
'                            .Font.Italic = False
'                            .NumberFormat = "General"
'                        End With
'                        With Range(targetRange.Cells(1, colOffset + singleCol), targetRange.Cells(rowCount + rowOffset, 1 + (j) * (singleCol + varCol)))
'                            .Font.Bold = False
'                            .Font.Italic = True
'                            .NumberFormat = "General"
'                        End With
'                    End If
'                Else
'                    colOffset = 1 + (j - 1) * singleCol
'                End If
'                If i = 1 Then
'                    If numHeaderRows > 0 Then
'                        For l = 1 To numHeaderRows
'                            On Error Resume Next
'                                targetRange.Cells(i + numHeaderRows - l, colOffset).Value = functionRange.Cells(i - l, j).Value & ": Value"
'                                targetRange.Cells(i + numHeaderRows - l, colOffset + 1).Value = functionRange.Cells(i - l, j).Value & ": +/- Error"
'                            On Error GoTo 0
'                        Next l
'                    Else
'                        targetRange.Cells(i, colOffset).Value = "Function " & j & ": Value"
'                        targetRange.Cells(i, colOffset + 1).Value = "Function " & j & ": +/- Error"
'                    End If
'                    With Range(targetRange.Cells(1, colOffset), targetRange.Cells(rowCount + rowOffset, colOffset + 1))
'                        .Font.Bold = True
'                        .Font.Italic = False
'                        .NumberFormat = "General"
'                    End With
'                End If
'            Else
'                If doSensitivity Then
'                    colOffset = 1 + (j - 1) * (doubleCol + varCol)
'                    If i = 1 Then
'                        For k = 1 To varCount
'                            If numHeaderRows > 0 Then
'                                For l = 1 To numHeaderRows
'                                    On Error Resume Next
'                                        targetRange.Cells(i + numHeaderRows - l, colOffset + k * 2 + 1).Value = varRange.Cells(i - l, k).Value
'                                        targetRange.Cells(i + numHeaderRows - l, colOffset + k * 2 + 2).Value = varRange.Cells(i - l, k).Value & ": % Diff"
'                                    On Error GoTo 0
'                                Next l
'                            Else
'                                targetRange.Cells(i, colOffset + k * 2 + 1).Value = "Variable " & k
'                                targetRange.Cells(i, colOffset + k * 2 + 2).Value = "Variable " & k & ": % Diff"
'                            End If
'                        Next k
'                        With Range(targetRange.Cells(1, colOffset), targetRange.Cells(rowCount + rowOffset, colOffset + 2))
'                            .Font.Bold = True
'                            .Font.Italic = False
'                            .NumberFormat = "General"
'                        End With
'                        With Range(targetRange.Cells(1, colOffset + doubleCol), targetRange.Cells(rowCount + rowOffset, 1 + (j) * (doubleCol + varCol)))
'                            .Font.Bold = False
'                            .Font.Italic = True
'                            .NumberFormat = "General"
'                        End With
'                    End If
'                Else
'                    colOffset = 1 + (j - 1) * doubleCol
'                End If
'                If i = 1 Then
'                    If numHeaderRows > 0 Then
'                        For l = 1 To numHeaderRows
'                            On Error Resume Next
'                                targetRange.Cells(i + numHeaderRows - l, colOffset).Value = functionRange.Cells(i - l, j).Value & ": Value"
'                                targetRange.Cells(i + numHeaderRows - l, colOffset + 1).Value = functionRange.Cells(i - l, j).Value & ": - Error"
'                                targetRange.Cells(i + numHeaderRows - l, colOffset + 2).Value = functionRange.Cells(i - l, j).Value & ": + Error"
'                            On Error GoTo 0
'                        Next l
'                    Else
'                        targetRange.Cells(i, colOffset).Value = "Function " & j & ": Value"
'                        targetRange.Cells(i, colOffset + 1).Value = "Function " & j & ": Low"
'                        targetRange.Cells(i, colOffset + 2).Value = "Function " & j & ": High"
'                    End If
'                    With Range(targetRange.Cells(1, colOffset), targetRange.Cells(rowCount + rowOffset, colOffset + 2))
'                        .Font.Bold = True
'                        .Font.Italic = False
'                        .NumberFormat = "General"
'                    End With
'                End If
'            End If
'            Application.Calculate
'            Call errorPropSub(targetRange.Cells(i + rowOffset, colOffset), functionRange.Cells(i, j), varRangeTemp, errorRangeTemp)
'            If errorOccurred And Not ignoreErrors Then
'                targetRange.Cells(i + rowOffset, colOffset).Value = oldVar
'                If Not doApprox Then
'                    targetRange.Cells(i + rowOffset, colOffset + 1).Value = oldErr
'                End If
'                errJunk = MsgBox("A " & errorString(errorValue) & " error occurred! Fix the cell taking responsibility and try again." & Chr(10) & Chr(10) & "This happened because a variable has been changed in order to check its influence on a function, and has left the domain of the function (or an intermediate function), resulted in division by zero, etc." & Chr(10) & Chr(10) & "Find and prevent these issues or run again ignoring errors.", vbCritical, "Error")
'                errorPropForm.Hide
'                Exit Sub
'            End If
'            If multiRow Then
'                Exit For
'            End If
'        Next i
'        If multiRow Then
'            Exit For
'        End If
'    Next j
'
'    If doApprox Then
'         If doSensitivity Then
'            colOffset = funcCount * (singleCol + varCount)
'        Else
'            colOffset = funcCount * singleCol
'        End If
'    Else
'        If doSensitivity Then
'            colOffset = funcCount * (doubleCol + varCount)
'        Else
'            colOffset = doubleCol * funcCount
'        End If
'    End If
'
'    Worksheets(targetRange.Parent.Name).Range(targetRange.Cells(1, 1), targetRange.Cells(rowOffset + rowCount, colOffset)).Columns.AutoFit
'
'    Application.ScreenUpdating = True
'
'    Application.GoTo (targetRange.Cells(1, 1))
'
'    errorPropForm.Hide
'End Sub
