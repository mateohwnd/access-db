Attribute VB_Name = "mdlDisableBypassKey"
Option Compare Database
Option Explicit

Sub LockDb()
' disable shift bypass key
On Error GoTo Sub_Error

    Dim db As Database
    Dim prp As Property
        Set db = CurrentDb
        db.Properties("AllowBypassKey") = False
        MsgBox "Lock Confirmed.", vbInformation, "Bypass"
    Exit Sub

Sub_Error:
    Select Case Err
        Case 3270
            Set prp = db.CreateProperty("AllowBypassKey", dbBoolean, True)
            db.Properties.Append prp
            db.Properties("AllowBypassKey") = False
            MsgBox "The bypass property was not found, so it was created.", vbInformation, "Bypass"
        Case Else
            MsgBox "There was an error. (" & Err & ") " & Error
    End Select
    
End Sub

Sub CheckCommand()
    ' bypass key enable
    'Checks open parameter/command
    Dim db As Database
        Set db = CurrentDb
        If Command() = "ABC123" Then
            db.Properties("AllowBypassKey") = True
            MsgBox "Unlock command received. Bypass key enabled.", vbInformation, "Bypass"
        End If
End Sub
