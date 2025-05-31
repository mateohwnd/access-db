Attribute VB_Name = "mdlAppWindow"
Option Compare Database
Option Explicit

'************ Code Start **********
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of Dev Ashish

'Additional API code by Daolix

'/* ShowWindow() Commands */

Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3

Global Const SW_SHOW = 5

Public blnShowWindow As Boolean

'###############################################
#If VBA7 Then
    Declare PtrSafe Function ShowWindow Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
        
    Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
            
    Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    Declare PtrSafe Function GetParent Lib "user32" _
        (ByVal hwnd As LongPtr) As LongPtr
        
#Else  '32-bit Office
    Declare Function ShowWindow Lib "user32" _
        (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
        
    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
            
    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
        
    Declare Function GetParent Lib "user32" _
        (ByVal hWnd As Long) As Long
#End If
'###############################################
        
'/* Window field offsets for Set/GetWindowLong() */
Public Const GWL_EXSTYLE       As Long = -20

'/* Extended Window Styles */
Public Const WS_EX_APPWINDOW   As Long = &H40000

Function SetAccessWindow(nCmdShow As Long)

    'Usage Examples
    'Maximize window:
    ' ?SetAccessWindow(SW_SHOWMAXIMIZED)
    'Minimize window:
    ' ?SetAccessWindow(SW_SHOWMINIMIZED)
    'Hide window:
    ' ?SetAccessWindow(SW_HIDE)
    'Normal window:
    ' ?SetAccessWindow(SW_SHOWNORMAL)
    
    Dim loX As Long
   ' Dim loForm As Form
    On Error Resume Next
    
    loX = ShowWindow(hWndAccessApp, nCmdShow)
    SetAccessWindow = (loX <> 0)

End Function

Function RestoreNormalWindow()
    'restore
    SetAccessWindow (SW_SHOWNORMAL)
End Function

Function ShowMaximisedWindow()
    'maximise
    SetAccessWindow (SW_SHOWMAXIMIZED)
End Function

Function ShowLastUsedWindowState()
    'dispalys window maximised or restored ...depending on last state used before hiding
    SetAccessWindow (SW_SHOW)
End Function

Function MinimizeWindow()
    'You can use this in the form load event of your startup form or in an autoexec macro
    SetAccessWindow (SW_SHOWMINIMIZED)
End Function

Function HideWindow()
    'You can use this in the form load event of your startup form or in an autoexec macro
    SetAccessWindow (SW_HIDE)
End Function

Function HideAppWindow(frm As Access.Form)
'new code - app window is NOT restored when taskbar icon clicked
    'omit the ...Or WS_EX_APPWINDOW ...section to hide the taskbar icon
    SetWindowLong frm.hwnd, GWL_EXSTYLE, GetWindowLong(frm.hwnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW
    ShowWindow Application.hWndAccessApp, SW_HIDE
    ShowWindow frm.hwnd, SW_SHOW
 End Function
 
Function HideAppWindowIcon(frm As Access.Form)
    'omit the ...Or WS_EX_APPWINDOW ...section to hide the taskbar icon
    SetWindowLong frm.hwnd, GWL_EXSTYLE, GetWindowLong(frm.hwnd, GWL_EXSTYLE) ' Or WS_EX_APPWINDOW
    ShowWindow Application.hWndAccessApp, SW_HIDE
    ShowWindow frm.hwnd, SW_SHOW
 End Function





