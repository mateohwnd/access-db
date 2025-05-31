Attribute VB_Name = "mdlCustomRightClickMenu"
Option Compare Database
Option Explicit

Public Sub CustomDesignView()

On Error GoTo Err_Handler

    Application.Echo False
    
    SetAccessWindow (SW_SHOW)
    ShowNavigationPane
    ShowRibbon
    
    ' Maximize ribbon
    If CommandBars("ribbon").Height < 100 Then
        CommandBars.ExecuteMso "MinimizeRibbon"
    End If

'    DoCmd.Close acForm, Me.Name
    DoCmd.Close 'acForm, "frmstart"
    DoCmd.OpenForm "frmStart", acDesign

    Application.Echo True
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    strProc = "FormHeader_MouseDown"
    MsgBox "Error " & Err.Number & " in " & strProc & " procedure : " & Err.Description
    Resume Exit_Handler
End Sub
