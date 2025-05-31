Attribute VB_Name = "mdlRibbon"
Option Compare Database
Option Explicit

Public blnShowRibbon As Boolean

Public Function HideRibbon()
    'could run at startup using Autoexec
    'however this also hides the QAT which makes printing reports tricky
     DoCmd.ShowToolbar "Ribbon", acToolbarNo
   '  DoCmd.ShowToolbar "PrintReport", acToolbarYes
End Function

Public Function ShowRibbon()
    'use when opening a report to display print preview ribbon
     DoCmd.ShowToolbar "Ribbon", acToolbarYes
End Function

Public Function ToggleRibbonState()

'hide ribbon if visible & vice versa
    CommandBars.ExecuteMso "MinimizeRibbon"
End Function

Public Function IsRibbonMinimized() As Boolean
    'Result: 0=normal (maximized), -1=autohide (minimized)

    IsRibbonMinimized = (CommandBars("Ribbon").Controls(1).Height < 100)
   ' Debug.Print IsRibbonMinimized
End Function

