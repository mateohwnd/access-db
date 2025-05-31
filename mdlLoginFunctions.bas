Attribute VB_Name = "mdlLoginFunctions"
Option Compare Database
Option Explicit

Public Function DecryptKey(sStr As String)

'===================================
'NO LONGER USED-replaced with RC4 code
'===================================

'If Len(sStr) < 4 Then Exit Function

'Dim dKey As String
'Dim dWord As String

'Added by Colin Riddington 9 Nov 2018
'Dim I As Integer
'Dim Letter As String


'For I = 1 To Len(Trim(sStr)) Step 4
'    Letter = Trim(Mid(sStr, I, 4))
'    dKey = Chr(Letter Xor 555)
'    dWord = dWord & dKey
'Next I
'DecryptKey = dWord

End Function

Public Function EncryptKey(sStr As String)

'===================================
'NO LONGER USED-replaced with RC4 code
'===================================

'Dim eKey As String

'Added by Colin Riddington 9 Nov 2018
'Dim I As Integer
'Dim Letter As String
'Dim charsInStr As Integer

'charsInStr = Len(Trim(sStr))

'For I = 1 To charsInStr
'    Letter = Mid(sStr, I, 1)
 '   eKey = eKey & CStr(Asc(Letter) Xor 555) & " "
'Next I
'EncryptKey = eKey

End Function

Public Function FindUserName()
    'NOT CURRENTLY USED
    'ONLY use this if you want to use the default network login name
    FindUserName = CreateObject("WScript.Network").UserName
End Function

Public Function GetComputerName()
   GetComputerName = CreateObject("WScript.Network").ComputerName
End Function

Public Function GetUserName()
    'gets logged in user name
    GetUserName = strUserName
End Function

Public Function varUserName() As Boolean

    If strUserName = "" Then
        varUserName = False
    Else
        varUserName = True
    End If

End Function

Public Function GetEncoderCode()
    On Error Resume Next ' Prevent unset variable errors
    
    GetEncoderCode = strEncoderCode
    
    If Err.Number <> 0 Then GetEncoderCode = "" ' Fallback if error
End Function

Public Function GetAccessLevel()
    GetAccessLevel = intAccessLevel
End Function

Public Function GetLoginID()
    'gets loginID for the current session
    GetLoginID = lngLoginID
End Function

Public Function FormattedMsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional Title As String = vbNullString, Optional HelpFile As Variant, Optional Context As Variant) As VbMsgBoxResult

On Error GoTo Err_Handler

'Taken from http://www.trigeminal.com/usenet/usenet015.asp

        FormattedMsgBox = Eval("MsgBox(""" & Prompt & _
         """, " & Buttons & ", """ & Title & """)")

Exit_Handler:
   Exit Function
 
Err_Handler:
   MsgBox "Error " & Err.Number & " in FormattedMsgBox procedure : " & vbCrLf & "   - " & Err.Description
   Resume Exit_Handler

End Function

