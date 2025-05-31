Attribute VB_Name = "mdlMoveForm"
Option Compare Database
Option Explicit

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HT_CAPTION = &H2

Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, _
    ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Declare PtrSafe Function ReleaseCapture Lib "user32.dll" () As Long

