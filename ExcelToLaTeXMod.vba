Attribute VB_Name = "ExcelToLaTeXMod"
Public Sub InitExcelToLaTeX()
    Call AdjustSizeForWin(ExcelToLaTeX)
    With ExcelToLaTeX
        .UserForm_Initialize
        .Show
    End With
End Sub
