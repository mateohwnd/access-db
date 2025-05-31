Attribute VB_Name = "mdlGlobal"
Public AccountPass As String

Public Function SendEmailUsingGmail()
'For Early Binding, enable Tools > References > Microsoft CDO for Windows 2000 Library

DoCmd.Hourglass True

    Dim NewMail As Object
    Dim mailConfig As Object
    Dim fields As Variant
    Dim msConfigURL As String
    On Error GoTo Err:

    'late binding
    Set NewMail = CreateObject("CDO.Message")
    Set mailConfig = CreateObject("CDO.Configuration")

    ' load all default configurations
    mailConfig.Load -1

    Set fields = mailConfig.fields

    'Set All Email Properties
    With NewMail
        .From = "christianhwnd@gmail.com"
        .To = Forms![frmForgotPass]!txtEmail.value '"kdt987d2@gmail.com"
        .cc = ""
        .BCC = ""
        .Subject = "DB Account"
        .TextBody = "Your accound password is " & AccountPass
        '.AddAttachment "D:\itan\Desktop\DB\todo.txt"
    End With

    msConfigURL = "http://schemas.microsoft.com/cdo/configuration"

    With fields
        .item(msConfigURL & "/smtpusessl") = True             'Enable SSL Authentication
        .item(msConfigURL & "/smtpauthenticate") = 1          'SMTP authentication Enabled
        .item(msConfigURL & "/smtpserver") = "smtp.gmail.com" 'Set the SMTP server details
        .item(msConfigURL & "/smtpserverport") = 465          'Set the SMTP port Details
        .item(msConfigURL & "/sendusing") = 2                 'Send using default setting
        .item(msConfigURL & "/sendusername") = "" 'Your gmail address
        .item(msConfigURL & "/sendpassword") = "" 'Your password or App Password
        .Update                                               'Update the configuration fields
    End With
    
    NewMail.Configuration = mailConfig
    NewMail.Send
    DoEvents
    
DoCmd.Hourglass False
    MsgBox "Your email has been sent", vbInformation, "Email"
    
Exit_Err:
    'Release object memory
    Set NewMail = Nothing
    Set mailConfig = Nothing
    End

Err:
    Select Case Err.Number
    Case -2147220973  'Could be because of Internet Connection
        MsgBox "Check your internet connection." & vbNewLine & Err.Number & ": " & Err.Description
    Case -2147220975  'Incorrect credentials User ID or password
        MsgBox "Check your login credentials and try again." & vbNewLine & Err.Number & ": " & Err.Description
    Case Else   'Report other errors
        MsgBox "Error encountered while sending email." & vbNewLine & Err.Number & ": " & Err.Description
    End Select

    Resume Exit_Err

End Function

Public Function IsValidEmail(sEmailAddress As String) As Boolean
    'Define variables
    Dim sEmailPattern As String
    Dim oRegEx As Object
    Dim bReturn As Boolean
    
    'Use the below regular expressions
    'sEmailPattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$" 'or
    sEmailPattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    'Create Regular Expression Object
    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
    oRegEx.Pattern = sEmailPattern
    bReturn = False
    
    'Check if Email match regex pattern
    If oRegEx.Test(sEmailAddress) Then
        'Debug.Print "Valid Email ('" & sEmailAddress & "')"
        bReturn = True
    Else
        'Debug.Print "Invalid Email('" & sEmailAddress & "')"
        bReturn = False
    End If

    'Return validation result
    IsValidEmail = bReturn
End Function

Public Function IsOpen(ByVal strFormName As String) As Boolean

    IsOpen = False
    ' is form open?
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0 Then
        ' if so make sure its not in design view
        If Forms(strFormName).CurrentView <> 0 Then
            IsOpen = True
        End If
    End If

    Exit Function

End Function

Public Function ListTempVars()
    'TempVar variable types (2 = integer, 8 = string, 7 = date)
    '6 = currency, 11 = boolean, 3 = long integer)
    Dim i                     As Long
    For i = 0 To TempVars.count - 1
        Debug.Print TempVars(i).Name, TempVars(i).value, VarType(TempVars(i))
    Next i
End Function


' Error Logging Subroutine
Public Sub LogErrors(procName As String, errNumber As Long, errDesc As String)
    Dim f As Integer
    Dim logPath As String
    
    logPath = CurrentProject.path & "\error_log.txt"
    f = FreeFile
    
    Open logPath For Append As #f
        Print #f, Now & " | Procedure: " & procName & " | Error " & errNumber & ": " & errDesc
    Close #f
End Sub
