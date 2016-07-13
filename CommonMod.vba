Attribute VB_Name = "CommonMod"
Option Explicit
'Change size coefficient
Const SizeCoefForMac = 1.2

'   Enlarges userforms that were created on windows so that
'   they are the same size when loaded on mac
Sub AdjustSizeForMac(theForm As Object)
    Dim ControlOnForm As Object

    #If Mac Then
    #Else
        'No mac so not run the code
        Exit Sub
    #End If

    With theForm
        'Change Userform size
        .Width = .Width * SizeCoefForMac
        .Height = .Height * SizeCoefForMac

        'Change controls/font on the userform
        For Each ControlOnForm In .Controls
            With ControlOnForm
                .Top = .Top * SizeCoefForMac
                .Left = .Left * SizeCoefForMac
                .Width = .Width * SizeCoefForMac
                .Height = .Height * SizeCoefForMac
                On Error Resume Next
                .Font.Size = .Font.Size * SizeCoefForMac
                On Error GoTo 0
            End With
        Next
    End With
End Sub

'   Shrinks userforms that were created on Mac so that
'   they are the same size when loaded in Windows
Sub AdjustSizeForWin(theForm As Object)
    Dim ControlOnForm As Object

    #If Mac Then
        'No mac so not run the code
        Exit Sub
    #End If

    With theForm
        'Change Userform size
        .Width = .Width / SizeCoefForMac
        .Height = .Height / SizeCoefForMac

        'Change controls/font on the userform
        For Each ControlOnForm In .Controls
            With ControlOnForm
                .Top = .Top / SizeCoefForMac
                .Left = .Left / SizeCoefForMac
                .Width = .Width / SizeCoefForMac
                .Height = .Height / SizeCoefForMac
                On Error Resume Next
                .Font.Size = .Font.Size / SizeCoefForMac
                On Error GoTo 0
            End With
        Next
    End With
End Sub
