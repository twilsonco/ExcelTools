VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelToLaTeX 
   Caption         =   "Excel to LaTeX tables"
   ClientHeight    =   11415
   ClientLeft      =   -75
   ClientTop       =   -48765
   ClientWidth     =   11190
   OleObjectBlob   =   "ExcelToLaTeX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelToLaTeX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' VBA bug: backslahes get turned into angstrom symbols, so they are avoided by using
' the QUOTIENT excel funtion or by specifying Chr(92) instead
                        
Option Explicit
Option Base 1

Dim RunOnChange As Boolean
Dim TB As String, CR As String, IND As String
Dim obj As New DataObject
    
Dim sBS As String, BS As String, dBS As String

Private Enum CellStatus_e
    IncompleteCell = 0
    IgnoreCell = 1
    BlankCell = 2
    NormalCell = 3
End Enum

Private Type Cell_s
    Status As CellStatus_e
    MCols As Integer
    MRows As Integer
    IsMerged As Boolean
    Alignment As String
    stringValue As String
    LRBorder(1 To 2) As Integer
    TBBorder(1 To 2) As Integer
    LeaderRC(1 To 2) As Integer
    '   Width and height are in inches
    Width As Double
    Height As Double
End Type

Private Sub B_Cancel_Click()
    ExcelToLaTeX.Hide
    Unload Me
End Sub

Private Sub B_Run_Click()
    Call copyToClipboard
End Sub

Private Sub BTN_Update_Click()
    Call main
End Sub

Private Sub CB_CompressWhiteSpace_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub CB_ConvertCode_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub CB_PreserveCellSizes_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub CB_PreserveFontSizes_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub CB_PreserveFormatting_Click()
    If (CB_PreserveFormatting) Then
        CB_ScaleTable.Enabled = True
        CB_UseBooktabs.Enabled = True
        CB_CompressWhiteSpace.Enabled = True
        CB_PreserveCellSizes.Enabled = True
    Else
        CB_ScaleTable.Enabled = False
        CB_UseBooktabs.Enabled = False
        CB_CompressWhiteSpace.Enabled = False
        CB_PreserveCellSizes.Enabled = False
    End If
    
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub CB_ScaleTable_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub CB_TabSpaces_Click()
    CB_TabTabs.Value = Not CB_TabSpaces.Value
    If CB_TabSpaces Then
        TB_TabSpaces.Enabled = True
        SB_TabSpaces.Enabled = True
    Else
        TB_TabSpaces.Enabled = False
        SB_TabSpaces.Enabled = False
    End If
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub CB_TabTabs_Click()
    CB_TabSpaces.Value = Not CB_TabTabs.Value
    
    Call CB_TabSpaces_Click
End Sub

Private Sub CB_UseBooktabs_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub LB_Indent_Click()
    Call SB_Indent_SpinUp
End Sub

Private Sub OB_CellsAligned_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub OB_CellsPerLine_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub OB_CellsShortened_Click()
    If RunOnChange Then
        Call main
    End If
End Sub

'Private Sub RE_Selection_Change()
'    If RunOnChange And Range(RE_Selection.value).Cells.Count > 1 Then
'        Call main
'    End If
'End Sub

'Private Sub RE_Selection_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER)
'    If RunOnChange And Range(RE_Selection.value).Cells.Count > 1 Then
'        Call main
'    End If
'End Sub

Private Sub SB_Indent_SpinDown()
    Dim Val As Integer
    
    With TB_Indent
        Val = .Value
        If (Val >= 1) Then
            Val = Val - 1
            .Locked = False
            .Value = Val
            .Locked = True
        End If
    End With
    
    If RunOnChange Then
        Call main
    End If
    
End Sub

Private Sub SB_Indent_SpinUp()
    Dim Val As Integer
    
    With TB_Indent
        Val = .Value
        Val = Val + 1
        .Locked = False
        .Value = Val
        .Locked = True
    End With
    
    Call main
End Sub

Private Sub SB_TabSpaces_SpinDown()
    Dim Val As Integer
    
    With TB_TabSpaces
        Val = .Value
        If (Val >= 2) Then
            Val = Val - 1
            .Locked = False
            .Value = Val
            .Locked = True
        End If
    End With
    
    If RunOnChange Then
        Call main
    End If
End Sub

Private Sub SB_TabSpaces_SpinUp()
    Dim Val As Integer
    
    With TB_TabSpaces
        Val = .Value
        Val = Val + 1
        .Locked = False
        .Value = Val
        .Locked = True
    End With
    
    If RunOnChange Then
        Call main
    End If
End Sub

Public Sub UserForm_Initialize()
    Dim theSelection As Range
    Dim theAddress As String

    RunOnChange = False

    CB_PreserveFormatting.Value = True
    
    CB_ScaleTable.Enabled = True
    CB_UseBooktabs.Enabled = True
    CB_CompressWhiteSpace.Enabled = True
    CB_ConvertCode.Enabled = True
    
    CB_ScaleTable.Value = False
    CB_UseBooktabs.Value = False
    CB_CompressWhiteSpace.Value = False
    CB_ConvertCode.Value = False
    
    With TB_Indent
        .Locked = False
        .Value = 0
        .Locked = True
    End With
    
    OB_CellsAligned.Value = True
    
    CB_TabTabs.Value = True
    
    CB_TabSpaces.Value = False
    TB_TabSpaces.Enabled = True
    SB_TabSpaces.Enabled = True
    TB_TabSpaces.Locked = False
    TB_TabSpaces.Value = 2
    TB_TabSpaces.Locked = True
    TB_TabSpaces.Enabled = False
    SB_TabSpaces.Enabled = False
    
    TB_Output.Enabled = True
    
    Set theSelection = Application.Selection
    theAddress = "'" & ActiveWorkbook.ActiveSheet.Name & "'!" & theSelection.Address
    RE_Selection.Value = theAddress
    
    RunOnChange = True
    
    Call main
    
    'Call ExcelToLaTeX.Show
End Sub

Private Function padString(theString As String, newLen As Integer) As String
    Dim i As Integer, numSpaces As Integer
    
    numSpaces = newLen - Len(theString)
    
    padString = theString
    
    For i = 1 To numSpaces
        padString = padString & " "
    Next i
End Function

Private Function getFontSizeString(fontSize As Double) As String
    ' based on the equivalent font sizes at https://en.wikibooks.org/wiki/LaTeX/Fonts
    Dim outStr As String
    outStr = ""
    
    If fontSize <= 6 Then
        outStr = sBS & "tiny"
    ElseIf fontSize <= 8 Then
        outStr = sBS & "scriptsize"
    ElseIf fontSize <= 10 Then
        outStr = sBS & "footnotesize"
    ElseIf fontSize <= 11 Then
        outStr = sBS & "small"
    ElseIf fontSize <= 12 Then
        'outStr = outStr
    ElseIf fontSize <= 14.4 Then
        outStr = sBS & "large"
    ElseIf fontSize <= 17.28 Then
        outStr = sBS & "Large"
    ElseIf fontSize <= 20.74 Then
        outStr = sBS & "LARGE"
    ElseIf fontSize <= 24.88 Then
        outStr = sBS & "huge"
    Else
        outStr = sBS & "Huge"
    End If
    
    getFontSizeString = outStr
End Function

Private Function getAlignment(tmpRange As Range) As String
    Dim isFound As Boolean, tmpDbl As Double
    
    isFound = False
    
    Select Case tmpRange.HorizontalAlignment
        Case xlLeft
            getAlignment = "l"
            isFound = True
        Case xlCenter
            getAlignment = "c"
            isFound = True
        Case xlRight
            getAlignment = "r"
            isFound = True
        Case xlJustify
            getAlignment = "l"
            isFound = True
        Case xlDistributed
            getAlignment = "l"
            isFound = True
        Case xlCenterAcrossSelection
            getAlignment = "c"
            isFound = True
    End Select
    
    If Not isFound Then
        If IsNumeric(tmpRange.Text) Then
            getAlignment = "r"
        Else
            getAlignment = "l"
        End If
    End If
End Function


Private Sub main()
    '   Variables for both methods
    Dim laTeXStr As String
    Dim stringValue As String
    Dim numCols As Integer, numRows As Integer
    Dim tableRange As Range, tmpRange As Range
    Dim isMac As Boolean, isRange As Boolean
    Dim DEL As String
    
    Dim EscapeChars, AlwaysEscapeChars, colorNames, colorRGBVals
    
    '   Variables used only when preserving formatting
    Dim useXColor As Boolean, useMultiRow As Boolean, useHyperRef As Boolean, useUlem As Boolean, useBooktabs As Boolean, useGraphicx As Boolean, isNamedColor As Boolean
    Dim hLineStrs, cLineStr As String, LCRs(), LRLines() As Integer, hLineCols() As Integer, tmpCLineStr As String, cNums(2) As Integer
    Dim tmpVar, tmpRGB(3) As Integer, LCR(3) As String
    Dim tmpAlignment As String, colorStr As String, tmpLineCounts() As Integer, lcrCounts() As Integer
    Dim tmpMax As Integer, tmpInd As Integer, alignmentStr As String, maxStringLengths() As Integer, tableContentsStr As String, tmpLineStr As String
    Dim cellMatrix() As Cell_s
    Dim i As Integer, j As Integer, r As Integer, c As Integer
    Dim uniqueCell As Boolean
    
    #If Mac Then
        isMac = True
    #Else
        isMac = False
    #End If
    
    sBS = Chr(92)
    BS = sBS & sBS
    dBS = BS & BS
    
    EscapeChars = Array( _
    Array(sBS, BS & "textbackslash "), _
    Array("$", BS & "$ "), _
    Array("^", BS & "textasciicircum "), _
    Array("_", BS & "_ "), _
    Array("{", BS & "{ "), _
    Array("}", BS & "} "), _
    Array("~", BS & "textasciitilda "), _
    Array("_", "---"), _
    Array("_", "--"))
    
    AlwaysEscapeChars = Array( _
    Array("#", BS & "# "), _
    Array("&", BS & "& "), _
    Array("%", BS & "% "))
    
    colorNames = Array( _
    "white", _
    "black", _
    "red", _
    "green", _
    "blue", _
    "cyan", _
    "magenta", _
    "yellow")
    
    colorRGBVals = Array( _
    Array(255, 255, 255), _
    Array(0, 0, 0), _
    Array(255, 0, 0), _
    Array(0, 255, 0), _
    Array(0, 0, 255), _
    Array(0, 255, 255), _
    Array(255, 0, 255), _
    Array(255, 255, 0))

    TB = Chr$(9)
    If CB_TabSpaces Then
        TB = ""
        For i = 1 To TB_TabSpaces.Value
            TB = TB & " "
        Next i
    End If
    IND = ""
    If (TB_Indent.Value > 0) Then
        IND = ""
        For i = 1 To TB_Indent.Value
            IND = IND & TB
        Next i
    End If
    
    CR = vbCrLf & IND 'vbNewLine 'Chr$(13)
    #If MAC_OFFICE_VERSION >= 15 Then
        CR = vbNewLine & IND 'vbNewLine 'Chr$(13)
    #End If
    
    On Error Resume Next
        isRange = IsObject(Range(RE_Selection.Value))
    On Error GoTo 0
    
    If isRange Then
        Set tableRange = Range(RE_Selection.Value)
        
        tableRange.Select
    
        With tableRange
            numCols = .Columns.Count
            numRows = .Rows.Count
        End With
    End If
    
    If Not isRange Or (numCols <= 1 And numRows <= 1) Then
        With TB_Output
            .Locked = False
            .ScrollBars = fmScrollBarsNone
            .Value = "Must select at least 2 cells"
            .ScrollBars = fmScrollBarsBoth
            .Locked = True
        End With
        Exit Sub
    End If
    
    ReDim LCRs(numCols)
    ReDim LRLines(numCols, 2)
    ReDim cellMatrix(numRows, numCols)
    
    #If Mac Then
        DEL = ":"
    #Else
        DEL = sBS
    #End If
    
    laTeXStr = IND & "% Copied from Microsoft Excel" & CR & _
        "% Workbook: " & ActiveWorkbook.Path & DEL & ActiveWorkbook.Name & CR & _
        "% Worksheet: " & ActiveWorkbook.ActiveSheet.Name & CR & _
        "% Range: " & tableRange.Address & CR
    
    If Not CB_PreserveFormatting Then
        ' If the user doesn't want to preserve formatting, then the code is much simpler and
        ' faster, so it's kept separate.
        laTeXStr = laTeXStr & BS & "begin{table}[h]" & CR & BS & "begin{tabular}{"
        For i = 1 To numCols
            laTeXStr = laTeXStr & "l"
        Next i
        laTeXStr = laTeXStr & "}" & CR
        For r = 1 To numRows
            For c = 1 To numCols
                With tableRange.Cells(r, c)
                    stringValue = .Text
                    If CB_ConvertCode Then
                        For Each tmpVar In EscapeChars
                            cellMatrix(r, c).stringValue = Replace(cellMatrix(r, c).stringValue, tmpVar(1), tmpVar(2))
                        Next tmpVar
                    Else
                        cellMatrix(r, c).stringValue = Replace(cellMatrix(r, c).stringValue, sBS, BS)
                    End If
                    laTeXStr = laTeXStr & stringValue
                    If (c < numCols) Then laTeXStr = laTeXStr & " & "
                End With
            Next c
            laTeXStr = laTeXStr & " " & dBS & CR
        Next r
        laTeXStr = laTeXStr & BS & "end{tabular}" & CR & BS & "label{table:label}" & CR & BS & "end{table}"
    Else
        useGraphicx = CB_ScaleTable
        useBooktabs = CB_UseBooktabs
        useXColor = False
        useMultiRow = False
        useHyperRef = False
        useUlem = False
        
        If useBooktabs Then
            hLineStrs = Array(BS & "toprule", BS & "midrule", BS & "bottomrule")
            cLineStr = BS & "cmidrule(lr)"
        Else
            hLineStrs = Array(BS & "hline", BS & "hline", BS & "hline")
            cLineStr = BS & "cline"
        End If

        LCR(1) = "l"
        LCR(2) = "c"
        LCR(3) = "r"
        
        '   Get all information from selected cells
        For c = 1 To numCols
            For r = 1 To numRows
                With tableRange.Cells(r, c)
                    cellMatrix(r, c).stringValue = .Text
                    
                    ' Get cell width and height in inches (72 points per inch)
                    cellMatrix(r, c).Height = Range(.Address).Height / 72#
                    cellMatrix(r, c).Width = Range(.Address).Width / 72#
                    
                    For Each tmpVar In AlwaysEscapeChars
                        cellMatrix(r, c).stringValue = Replace(cellMatrix(r, c).stringValue, tmpVar(1), tmpVar(2))
                    Next tmpVar
                    
                    If CB_ConvertCode Then
                        For Each tmpVar In EscapeChars
                            cellMatrix(r, c).stringValue = Replace(cellMatrix(r, c).stringValue, tmpVar(1), tmpVar(2))
                        Next tmpVar
                    Else
                        cellMatrix(r, c).stringValue = Replace(cellMatrix(r, c).stringValue, sBS, BS)
                    End If
                    
                    With .Font
                        If .Bold Then cellMatrix(r, c).stringValue = BS & "textbf{" & cellMatrix(r, c).stringValue & "}"
                        If .Italic Then cellMatrix(r, c).stringValue = BS & "textit{" & cellMatrix(r, c).stringValue & "}"
                        If .Underline <> xlUnderlineStyleNone Then cellMatrix(r, c).stringValue = BS & "underline{" & cellMatrix(r, c).stringValue & "}"
                        If .StrikeThrough Then
                            useUlem = True
                            cellMatrix(r, c).stringValue = BS & "sout{" & cellMatrix(r, c).stringValue & "}"
                        End If
                        If CB_PreserveFontSizes Then
                            On Error Resume Next
                                tmpVar = getFontSizeString(CDbl(.Size))
                                If tmpVar <> "" Then
                                    cellMatrix(r, c).stringValue = "{" & tmpVar & " " & cellMatrix(r, c).stringValue & "}"
                                End If
                            On Error GoTo 0
                        End If
                        tmpRGB(1) = .Color Mod 256
                        'tmpRGB(2) = .Color \ 256 Mod 256
                        'tmpRGB(3) = .Color \ 65536 Mod 256
                        tmpRGB(2) = Application.Quotient(.Color, 256) Mod 256
                        tmpRGB(3) = Application.Quotient(.Color, 65536) Mod 256
                    End With
                    
                    For Each tmpVar In tmpRGB
                        If CInt(tmpVar) <> 0 Then
                            useXColor = True
                            isNamedColor = False
                            For i = LBound(colorNames) To UBound(colorNames)
                                isNamedColor = True
                                For j = 1 To 3
                                    If (colorRGBVals(i)(j) <> tmpRGB(j)) Then
                                        isNamedColor = False
                                        Exit For
                                    End If
                                Next j
                                If isNamedColor Then
                                    colorStr = colorNames(i)
                                    Exit For
                                End If
                            Next i
                            
                            If Not isNamedColor Then
                                colorStr = "{" & BS & "color[rgb]{"
                                For i = 1 To 3
                                    colorStr = colorStr & CStr(CDbl(tmpRGB(i) / 255#))
                                    If i < 3 Then colorStr = colorStr & ","
                                Next i
                                cellMatrix(r, c).stringValue = colorStr & "}" & cellMatrix(r, c).stringValue & "}"
                                Exit For
                            Else
                                cellMatrix(r, c).stringValue = "{" & BS & "color{" & colorStr & "}" & cellMatrix(r, c).stringValue & "}"
                            End If
                        End If
                    Next tmpVar
                    
                    With .Hyperlinks
                        If .Count > 0 Then
                            useHyperRef = True
                            cellMatrix(r, c).stringValue = BS & "href{" & .Item(1).Address & "}{" & cellMatrix(r, c).stringValue & "}"
                        End If
                    End With
                    
                    With .Interior
                        tmpRGB(1) = .Color Mod 256
                        'tmpRGB(2) = .Color \ 256 Mod 256
                        'tmpRGB(3) = .Color \ 65536 Mod 256
                        tmpRGB(2) = Application.Quotient(.Color, 256) Mod 256
                        tmpRGB(3) = Application.Quotient(.Color, 65536) Mod 256
                    End With
                    
                    For Each tmpVar In tmpRGB
                        If CInt(tmpVar) <> 255 Then
                            useXColor = True
                            isNamedColor = False
                            For i = LBound(colorNames) To UBound(colorNames)
                                isNamedColor = True
                                For j = 1 To 3
                                    If (colorRGBVals(i)(j) <> tmpRGB(j)) Then
                                        isNamedColor = False
                                        Exit For
                                    End If
                                Next j
                                If isNamedColor Then
                                    colorStr = colorNames(i)
                                    Exit For
                                End If
                            Next i
                            
                            If Not isNamedColor Then
                                colorStr = BS & "cellcolor[rgb]{"
                                For i = 1 To 3
                                    colorStr = colorStr & CStr(CDbl(tmpRGB(i) / 255#))
                                    If i < 3 Then colorStr = colorStr & ","
                                Next i
                                cellMatrix(r, c).stringValue = colorStr & "}" & cellMatrix(r, c).stringValue
                                Exit For
                            Else
                                cellMatrix(r, c).stringValue = "{" & BS & "cellcolor{" & colorStr & "}" & cellMatrix(r, c).stringValue & "}"
                            End If
                        End If
                    Next tmpVar
                    
                    With .MergeArea
                        cellMatrix(r, c).MCols = .Columns.Count
                        cellMatrix(r, c).MRows = .Rows.Count
                    End With
                    
                    cellMatrix(r, c).Alignment = getAlignment(tableRange.Cells(r, c))
                    
                    If .Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone Then cellMatrix(r, c).LRBorder(1) = 1
                    If .Borders(xlEdgeRight).LineStyle <> xlLineStyleNone Then cellMatrix(r, c).LRBorder(2) = 1
                    If .Borders(xlEdgeTop).LineStyle <> xlLineStyleNone Then cellMatrix(r, c).TBBorder(1) = 1
                    If .Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then cellMatrix(r, c).TBBorder(2) = 1
                End With
                With cellMatrix(r, c)
                    .LeaderRC(1) = r
                    .LeaderRC(2) = c
                    .Status = IncompleteCell
                End With
            Next r
        Next c
        
        '   Now all information is known from Excel about tableRange
        '   Next, adjust cellMatrix for merged areas and get alignment string
        
        For c = 1 To numCols
            ReDim tmpLineCounts(2)
            ReDim lcrCounts(3)
'            For i = 1 To 3
'                If i < 3 Then tmpLineCounts(i) = 0
'                lcrCounts(i) = 0
'            Next i
            For r = 1 To numRows
                With cellMatrix(r, c)
                    If .MCols > 1 Then
                        cellMatrix(.LeaderRC(1), .LeaderRC(2)).IsMerged = True
                        .LRBorder(2) = .LRBorder(2) + cellMatrix(r, c + .MCols - 1).LRBorder(2)
                        For j = c + 1 To c + .MCols - 1
                            cellMatrix(r, j).Alignment = ""
                            cellMatrix(r, j).LeaderRC(1) = .LeaderRC(1)
                            cellMatrix(r, j).LeaderRC(2) = .LeaderRC(2)
                            cellMatrix(r, j).MCols = 0
                        Next j
                    End If
                    If .MRows > 1 Then
                        cellMatrix(.LeaderRC(1), .LeaderRC(2)).IsMerged = True
                        For i = r + 1 To r + .MRows - 1
                            cellMatrix(i, c).LeaderRC(1) = .LeaderRC(1)
                            cellMatrix(i, c).LeaderRC(2) = .LeaderRC(2)
                            cellMatrix(i, c).MRows = 0
                            cellMatrix(i - 1, c).TBBorder(2) = 0
                        Next i
                    End If
                    
                    For i = 1 To 3
                        If StrComp(.Alignment, LCR(i)) = 0 Then
                            lcrCounts(i) = lcrCounts(i) + 1
                            Exit For
                        End If
                    Next i
                    
                    If .MCols <= 0 Then
                        tmpLineCounts(1) = tmpLineCounts(1) + 1
                        tmpLineCounts(2) = tmpLineCounts(2) + 1
                    End If
                    
                    For i = 1 To 2
                        tmpLineCounts(i) = tmpLineCounts(i) + .LRBorder(i)
                        If c < numCols Then Exit For
                    Next i
                End With
            Next r
            
            For i = 1 To 2
                If tmpLineCounts(i) >= numRows Then LRLines(c, i) = 1
            Next i
            
            tmpMax = 0
            tmpInd = 1
            For i = 1 To 3
                If lcrCounts(i) > tmpMax Then
                    tmpMax = lcrCounts(i)
                    tmpInd = i
                End If
            Next i
            LCRs(c) = LCR(tmpInd)
            
        Next c
        
        alignmentStr = ""
        If useBooktabs Then alignmentStr = alignmentStr & "@{}"
        For c = 1 To numCols
            For i = 1 To LRLines(c, 1)
                alignmentStr = alignmentStr & "|"
            Next i
            alignmentStr = alignmentStr & LCRs(c)
            For i = 1 To LRLines(c, 2)
                alignmentStr = alignmentStr & "|"
            Next i
        Next c
        If useBooktabs Then alignmentStr = alignmentStr & "@{}"
        
        ReDim maxStringLengths(numCols)
        
        For r = 1 To numRows
            For c = 1 To numCols
                With cellMatrix(r, c)
                    If .Status = IncompleteCell Then
                        tmpAlignment = ""
                        
                        uniqueCell = (StrComp(.Alignment, LCRs(c)) <> 0)
                        If Not uniqueCell Then uniqueCell = (.LRBorder(1) <> LRLines(c, 1))
                        If Not uniqueCell And c >= numCols Then uniqueCell = (.LRBorder(2) <> LRLines(c, 2))
                        
                        tmpVar = .stringValue
                        
                        If uniqueCell Or .MCols > 1 Then
                            tmpVar = BS & "multicolumn{" & CStr(.MCols) & "}{"
                            
                            If .LRBorder(1) > 0 And (.LRBorder(1) <> LRLines(c, 1) Or c <= 1) Then
                                tmpVar = tmpVar & "|"
                            End If
                            
                            tmpVar = tmpVar & .Alignment
                            
                            If .LRBorder(2) > 0 And (.LRBorder(2) <> LRLines(c, 2) Or c >= numCols) Then
                                tmpVar = tmpVar & "|"
                            End If
                            
                            tmpVar = tmpVar & "}{"
                            
                            If .MCols > 1 Then
                                For j = c + 1 To c + .MCols - 1
                                    cellMatrix(r, j).Status = IgnoreCell
                                Next j
                            End If
                            
                            If .MRows > 1 Then
                                useMultiRow = True
                                For j = c To c + .MCols - 1
                                    For i = r + 1 To r + .MRows - 1
                                        With cellMatrix(i, j)
                                            If j = c Then
                                                .stringValue = tmpVar & "}"
                                                .Status = NormalCell
                                            Else
                                                .Status = IgnoreCell
                                            End If
                                        End With
                                    Next i
                                Next j
                                
                                tmpVar = tmpVar & BS & "multirow{" & CStr(.MRows) & "}{*}{" & .stringValue & "}}"
                            Else
                                tmpVar = tmpVar & .stringValue & "}"
                            End If
                        ElseIf .MRows > 1 Then
                            useMultiRow = True
                            tmpVar = BS & "multirow{" & CStr(.MRows) & "}{*}{" & .stringValue & "}"
                            
                            For i = r + 1 To r + .MRows - 1
                                cellMatrix(i, c).Status = BlankCell
                            Next i
                        End If
                        
                        .Status = NormalCell
                        .stringValue = tmpVar
                        
                        If Not isMac Then
                            .stringValue = Replace(.stringValue, BS, sBS)
                        End If
                        
                        maxStringLengths(c) = Application.Max(maxStringLengths(c), Len(Replace(tmpVar, BS, sBS)))
                    End If
                End With
            Next c
        Next r
        
        'cellMatrix is all ready. Time to build the actual LaTeX string.
        
        tableContentsStr = ""
        
        For r = 1 To numRows
            tmpLineStr = TB
            If r = 1 Then
                ReDim hLineCols(1)
                hLineCols(1) = -1
                For j = 1 To numCols
                    If cellMatrix(r, j).TBBorder(1) > 0 Then
                        If hLineCols(1) > 0 Then
                            ReDim Preserve hLineCols(UBound(hLineCols) + 1)
                        End If
                        hLineCols(UBound(hLineCols)) = j
                    End If
                Next j
                
                If UBound(hLineCols) >= numCols And hLineCols(UBound(hLineCols)) > 0 Then
                    tmpLineStr = tmpLineStr & hLineStrs(1) & CR & TB
                ElseIf hLineCols(1) > 0 Then
                    tmpCLineStr = ""
                    cNums(1) = hLineCols(1)
                    cNums(2) = hLineCols(1)
                    For Each tmpVar In hLineCols
                        If CInt(tmpVar) = cNums(2) + 1 Then
                            cNums(2) = tmpVar
                        ElseIf CInt(tmpVar) > cNums(2) + 1 Then
                            tmpCLineStr = tmpCLineStr & cLineStr & "{" & CStr(cNums(1)) & "-" & CStr(cNums(2)) & "}"
                            cNums(1) = CInt(tmpVar)
                            cNums(2) = CInt(tmpVar)
                        End If
                    Next tmpVar
                    tmpCLineStr = tmpCLineStr & cLineStr & "{" & CStr(cNums(1)) & "-" & CStr(cNums(2)) & "}"
                    tmpLineStr = tmpLineStr & tmpCLineStr & CR & TB
                Else
                    'tmpLineStr = tmpLineStr
                End If
            End If
            
            For c = 1 To numCols
                With cellMatrix(r, c)
                    If .Status = IgnoreCell Then
                        If Not CB_CompressWhiteSpace Then
                            tmpLineStr = tmpLineStr & padString(" ", maxStringLengths(c) + 3)
                        End If
                    Else
                        If c > 1 Then tmpLineStr = tmpLineStr & " & "
                        If OB_CellsShortened Then
                            tmpLineStr = tmpLineStr & .stringValue
                        ElseIf OB_CellsAligned Then
                            tmpLineStr = tmpLineStr & padString(.stringValue, maxStringLengths(c) + (Len(.stringValue) - Len(Replace(.stringValue, BS, sBS))))
                        ElseIf OB_CellsPerLine Then
                            tmpLineStr = tmpLineStr & .stringValue & CR & TB & TB
                        End If
                    End If
                End With
                
                If c >= numCols And r < numRows Then
                    tmpLineStr = tmpLineStr & " " & dBS
                End If
            Next c
            
            
            
            ReDim hLineCols(1)
            hLineCols(1) = -1
            For j = 1 To numCols
                If cellMatrix(r, j).TBBorder(2) > 0 Then
                    If hLineCols(1) > 0 Then
                        ReDim Preserve hLineCols(UBound(hLineCols) + 1)
                    End If
                    hLineCols(UBound(hLineCols)) = j
                End If
            Next j
            
            If UBound(hLineCols) >= numCols And hLineCols(UBound(hLineCols)) > 0 Then
                If r < numRows Then
                    tmpLineStr = tmpLineStr & hLineStrs(2) & CR
                Else
                    tmpLineStr = tmpLineStr & " " & dBS & CR & TB & hLineStrs(3) & CR
                End If
            ElseIf hLineCols(1) > 0 Then
                tmpCLineStr = ""
                cNums(1) = hLineCols(1)
                cNums(2) = hLineCols(1)
                For Each tmpVar In hLineCols
                    If CInt(tmpVar) = cNums(2) + 1 Then
                        cNums(2) = tmpVar
                    ElseIf CInt(tmpVar) > cNums(2) + 1 Then
                        tmpCLineStr = tmpCLineStr & cLineStr & "{" & CStr(cNums(1)) & "-" & CStr(cNums(2)) & "}"
                        cNums(1) = CInt(tmpVar)
                        cNums(2) = CInt(tmpVar)
                    End If
                Next tmpVar
                tmpCLineStr = tmpCLineStr & cLineStr & "{" & CStr(cNums(1)) & "-" & CStr(cNums(2)) & "}"
                If r < numRows Then
                    tmpLineStr = tmpLineStr & tmpCLineStr & CR
                Else
                    tmpLineStr = tmpLineStr & " " & dBS & CR & TB & tmpCLineStr & CR
                End If
            Else
                tmpLineStr = tmpLineStr & CR
            End If
            
            tableContentsStr = tableContentsStr & tmpLineStr
        Next r
        
        If useXColor Or useBooktabs Or useMultiRow Or useGraphicx Or useHyperRef Or useUlem Then
            laTeXStr = laTeXStr & "% Please add the following required packages to your document preamble:" & CR
            If useBooktabs Then laTeXStr = laTeXStr & "% " & BS & "usepackage{booktabs}" & CR
            If useMultiRow Then laTeXStr = laTeXStr & "% " & BS & "usepackage{multirow}" & CR
            If useGraphicx Then laTeXStr = laTeXStr & "% " & BS & "usepackage{graphicx}" & CR
            If useHyperRef Then laTeXStr = laTeXStr & "% " & BS & "usepackage{hyperref}" & CR
            If useUlem Then laTeXStr = laTeXStr & "% " & BS & "usepackage{ulem}" & CR
            If useXColor Then
                If isMac Then
                    laTeXStr = laTeXStr & "% " & BS & "usepackage[table,xcdraw]{xcolor}" & _
                        CR & "% If you use beamer only pass ""xcolor=table"" option, i.e. " & BS & "documentclass[xcolor=table]{beamer}" & CR
                Else
                    laTeXStr = laTeXStr & "% " & BS & "usepackage[table,xcdraw]{xcolor}" & _
                        CR & "% If you use beamer only pass ""xcolor=table"" option, i.e. " & BS & "documentclass[xcolor=table]{beamer}" & CR
                End If
            End If
        End If
        
        laTeXStr = laTeXStr & BS & "begin{table}[h]" & CR
        If useGraphicx Then laTeXStr = laTeXStr & BS & "resizebox{" & BS & "textwidth}{!}{" & CR
        laTeXStr = laTeXStr & BS & "begin{tabular}{" & alignmentStr & "}" & CR & _
            tableContentsStr & BS & "end{tabular}"
        If useGraphicx Then laTeXStr = laTeXStr & "}"
        laTeXStr = laTeXStr & CR & BS & "label{table:label}" & CR & BS & "end{table}"
        
    End If
    
    If InStr(laTeXStr, Chr(180)) > 0 Then
        laTeXStr = Replace(laTeXStr, Chr(180), Chr(92))
        laTeXStr = Replace(laTeXStr, Chr(92) & Chr(92), Chr(92))
    End If
    'laTeXStr = Replace(laTeXStr, Chr(180) & Chr(180), Chr(92) & Chr(92))
    'laTeXStr = Replace(laTeXStr, Chr(92) & Chr(92), "\")
    'laTeXStr = Replace(laTeXStr, "\\", "\")
    laTeXStr = Replace(laTeXStr, BS, sBS)
    
    'If InStr(laTeXStr, dBS) > 0 Then
    '    laTeXStr = Replace(laTeXStr, "\\", "\")
    'End If
    
    With TB_Output
        .Locked = False
        .ScrollBars = fmScrollBarsNone
        .Value = laTeXStr
        .ScrollBars = fmScrollBarsBoth
        .Locked = True
    End With
    
End Sub

Private Sub copyToClipboard()
    Dim laTeXStr As String, tmpVar
    'Dim FilePath As String
    
    laTeXStr = TB_Output.Text

    #If Mac Then
        laTeXStr = Replace(laTeXStr, sBS, BS)
        laTeXStr = Replace(laTeXStr, vbCrLf, Chr(13))
        laTeXStr = Replace(laTeXStr, """", sBS & """")
        laTeXStr = "set the clipboard to """ & laTeXStr & """"
        
        'FilePath = "SSD:Users:Haiiro:Desktop:as.txt"
        'Open FilePath For Output As #1
        'Write #1, laTeXStr
        'Close #1
        tmpVar = MacScript(laTeXStr)
    #Else
        obj.SetText laTeXStr
        obj.PutInClipboard
    #End If
    
    Call MsgBox("Table copied to clipboard as LaTeX code." & vbNewLine & "Paste it into your LaTeX editor.")
End Sub
