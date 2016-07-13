Attribute VB_Name = "RibbonControlMod"
#If Mac Then
    #If MAC_OFFICE_VERSION >= 15 Then
        'Callback for X2LButton onAction
        Sub E2LMacro(control As IRibbonControl)
            Call InitExcelToLaTeX
        End Sub
    
        'Callback for ErrPropButton onAction
        Sub ErrPropMacro(control As IRibbonControl)
            Call ErrorProp
        End Sub
    #End If
#Else
    'Callback for X2LButton onAction
    Sub E2LMacro(control As IRibbonControl)
        Call InitExcelToLaTeX
    End Sub
    
    'Callback for ErrPropButton onAction
    Sub ErrPropMacro(control As IRibbonControl)
        Call ErrorProp
    End Sub
#End If
