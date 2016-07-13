VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub MakeMenu()
    CreateCommandBar
End Sub

Private Sub temp()
    Application.CommandBars("Data").Controls("Convert table to LaTeX").Delete
End Sub


Private Sub Workbook_Open()
    Dim j As Integer
    Dim ComBarNames() As String, ControlNames() As String, MacroNames() As String
    ComBarNames = Split("Data|Data", "|", -1, vbBinaryCompare)
    ControlNames = Split("Convert table to LaTeX|Error Propagation Calculator", "|", -1, vbBinaryCompare)
    MacroNames = Split("ExcelToLaTeXMod.InitExcelToLaTeX|errorPropMod.ErrorProp", "|", -1, vbBinaryCompare)
    
    #If Mac And MAC_OFFICE_VERSION < 15 Then
        For j = LBound(ComBarNames) To UBound(ComBarNames)
            DeleteCommandBar ComBarNames(j), ControlNames(j), MacroNames(j)
        Next j
    #Else
        If ActiveWorkbook Is Nothing Then CreateCommandBar
    #End If
End Sub

Sub DeleteCommandBar(ComBarName As String, ControlName As String, MacroName As String)
    Dim i As Integer
    On Error GoTo 5
    For i = 1 To 100
        Application.CommandBars(ComBarName).Controls(ControlName).Delete
    Next i
5:
    On Error GoTo 0
    CreateMenuItem ControlName, MacroName
End Sub

Sub CreateCommandBar()
    Dim i As Integer
    Dim MacroNames() As String, ControlNames() As String
    MacroNames = Split("ExcelToLaTeXMod.InitExcelToLaTeX|errorPropMod.ErrorProp", "|", -1, vbBinaryCompare)
    ControlNames = Split("Convert table to LaTeX|Error Propagation Calculator", "|", -1, vbBinaryCompare)
    
    For i = LBound(MacroNames) To UBound(MacroNames)
        CreateMenuItem ControlNames(i), MacroNames(i)
    Next i
End Sub

Private Sub CreateMenuItem(ByVal Caption As String, ByVal Action As String)
    Dim ctl As CommandBarControl
    Dim i As Long
    Dim ControlCollection As New Collection
    Dim myMenubar As CommandBar, toolsMenu As CommandBarPopup, newMenuItem As CommandBarControl, newButton As CommandBarControl
    
    Dim DoCode As Boolean
    
    #If MAC_OFFICE_VERSION >= 15 Then
        'MsgBox "Excel 2016 for the Mac"
        DoCode = False
    #End If

    #If Mac Then
        If Val(Application.Version) < 15 Then
            'MsgBox "Excel 2011 or earlier for the Mac"
            DoCode = True
        End If
    #Else
        'MsgBox "Excel for Windows"
        DoCode = True
    #End If
    
    If DoCode Then
    'First Create Menu Item
        Set myMenubar = Application.CommandBars.ActiveMenuBar
        Set toolsMenu = myMenubar.Controls(8)
        Set newMenuItem = myMenubar.FindControl(Tag:=Action, recursive:=True)
        If Not newMenuItem Is Nothing Then newMenuItem.Delete
        Set newMenuItem = toolsMenu.Controls.Add(Type:=msoControlButton, Before:=8)
        newMenuItem.Tag = Action
    
    ' Versions before Office 2007 only (=> ribbons!)
        If CLng(Split(Application.Version, ".")(0)) < 12 Then
            'Now create tool bar
            Dim myToolBar As CommandBar
            On Error Resume Next
            Set myToolBar = Application.CommandBars(Action)
            On Error GoTo 0
            If myToolBar Is Nothing Then
                Set myToolBar = Application.CommandBars.Add(Name:=Action)
            End If
            If myToolBar.Controls.Count > 0 Then myToolBar.Controls(1).Delete
            
            myToolBar.Position = msoBarTop
            myToolBar.Visible = True
            Set newButton = myToolBar.Controls.Add(msoControlButton)
        End If
    End If
      
    If Not newButton Is Nothing Then
        ControlCollection.Add newButton
    End If
    If Not newMenuItem Is Nothing Then
        ControlCollection.Add newMenuItem
    End If
    
    For Each ctl In ControlCollection
        ctl.OnAction = Action
        ctl.FaceId = 107
        ctl.TooltipText = Caption
        ctl.Caption = Caption
    Next
End Sub
