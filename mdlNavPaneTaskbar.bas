Attribute VB_Name = "mdlNavPaneTaskbar"
Option Compare Database
Option Explicit
 
Dim handleW1 As Long
 
'###############################################
#If VBA7 Then 'add PtrSafe
    Private Declare PtrSafe Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
         
    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal handleW1 As Long, _
        ByVal handleW1InsertWhere As Long, ByVal w As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal z As Long, _
        ByVal wFlags As Long) As Long
#ElseIf Win64 Then 'need datatype LongPtr
    Private Declare PtrSafe Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
         
    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal handleW1 As LongPtr, _
        ByVal handleW1InsertWhere As LongPtr, ByVal w As LongPtr, _
        ByVal X As LongPtr, ByVal Y As LongPtr, ByVal z As LongPtr, _
        ByVal wFlags As LongPtr) As LongPtr
#Else '32-bit Office
    Private Declare Function FindWindowA Lib "user32" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
         
    Private Declare Function SetWindowPos Lib "user32" _
        (ByVal handleW1 As Long, _
        ByVal handleW1InsertWhere As Long, ByVal w As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal z As Long, _
        ByVal wFlags As Long) As Long
#End If
'###############################################
 
Const TOGGLE_HIDEWINDOW = &H80
Const TOGGLE_UNHIDEWINDOW = &H40

Public blnShowNavPane As Boolean

Function HideTaskbar()
    handleW1 = FindWindowA("Shell_traywnd", "")
    Call SetWindowPos(handleW1, 0, 0, 0, 0, 0, TOGGLE_HIDEWINDOW)
End Function
 
Function ShowTaskbar()
    Call SetWindowPos(handleW1, 0, 0, 0, 0, 0, TOGGLE_UNHIDEWINDOW)
End Function

Public Function ShowNavigationPane()

On Error GoTo ErrHandler

  '  DoCmd.OpenForm "frmSettings", acDesign
    DoCmd.SelectObject acForm, , True
    
Exit_ErrHandler:
    Exit Function
    
ErrHandler:
    MsgBox "Error " & Err.Number & " in ShowNavigationPane routine : " & Err.Description, vbOKOnly + vbCritical
    Resume Exit_ErrHandler

End Function
Public Function HideNavigationPane()

'CR modified v5263

On Error GoTo ErrHandler

    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.RunCommand acCmdWindowHide
        
Exit_ErrHandler:
    Exit Function
    
ErrHandler:
    MsgBox "Error " & Err.Number & " in HideNavigationPane routine : " & Err.Description, vbOKOnly + vbCritical
    Resume Exit_ErrHandler

End Function

Public Function MinimizeNavigationPane()

On Error GoTo ErrHandler

    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.Minimize
        
Exit_ErrHandler:
    Exit Function
    
ErrHandler:
    MsgBox "Error " & Err.Number & " in MinimizeNavigationPane routine : " & Err.Description, vbOKOnly + vbCritical
    Resume Exit_ErrHandler

End Function

Public Function MaximizeNavigationPane()

On Error GoTo ErrHandler

    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.Maximize
        
Exit_ErrHandler:
    Exit Function
    
ErrHandler:
    MsgBox "Error " & Err.Number & " in MaximizeNavigationPane routine : " & Err.Description, vbOKOnly + vbCritical
    Resume Exit_ErrHandler

End Function





