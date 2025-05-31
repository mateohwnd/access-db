Attribute VB_Name = "mdlLogOut"
Option Compare Database

Public Sub Logout()

    Dim f As Access.Form
    Dim i As Long
    
    ' Loop all open forms, from last to first, to avoid problems due to closing forms
    ' (removing them from the Forms collection) in the loop
    For i = Forms.count - 1 To 0 Step -1
        Set f = Forms(i)
        ' Close all forms except the login form
        If f.Name <> "frmStart" Then
            DoCmd.Close acForm, f.Name
        End If
    Next i
    
    Call CloseSession
    Call LogMeOff(lngLoginID)
    
End Sub
