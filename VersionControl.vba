Attribute VB_Name = "VersionControl"

Sub ImportCodeModules()
    Dim VBCount As Integer
    
    With ThisWorkbook.VBProject
        VBCount = .VBComponents.Count
        
        For i% = VBCount To 1 Step -1
    
            ModuleName = .VBComponents(i%).CodeModule.Name
    
            If ModuleName <> "VersionControl" Then
                If .VBComponents(i%).Type <> 100 Then
                     .VBComponents.Remove .VBComponents(ModuleName)
                     .VBComponents.Import "C:\Users\Haiiro\Dropbox\Excel_Projects\ExcelTools\" & ModuleName & ".vba"
'                     If .VBComponents(i%).Type = 3 Then
'                         .VBComponents.Import "SDD:Users:Haiiro:Dropbox:SVN:TBP_PRU:" & ModuleName & ".frx"
'                    End If
               End If
            End If
        Next i
        
    End With
    
    Call NameFix

End Sub

Sub NameFix()
    Dim VBCount As Integer
    
    With ThisWorkbook.VBProject
        VBCount = .VBComponents.Count
        
        For j% = VBCount To 1 Step -1
            With .VBComponents(j%)
                If .Name = "MacroModule1" Then
                    .Name = "MacroModule"
                    
                    Exit For
                End If
            End With
        Next j
    End With
End Sub


Sub Whats_In_A_Name()
    ThisWorkbook.VBProject.VBComponents("Module1").Name = "MacroModule"
End Sub

Sub testtype()

    MsgBox ThisWorkbook.VBProject.VBComponents("Sheet4").Type

End Sub


Sub SaveCodeModules()

    'This code Exports all VBA modules
    Dim i%, sName$
    
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.Name
                .VBComponents(i%).Export "C:\Users\Haiiro\Dropbox\Excel_Projects\ExcelTools\" & sName$ & ".vba"
            End If
        Next i
    End With

End Sub
