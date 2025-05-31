Attribute VB_Name = "mdlLoginEncryption"
Option Compare Database
Option Explicit

Public Const RC4_Key = "isladogs" 'This key is used to store the passwords - ideally it should be encrypted by a different method

'##############################################################
'# RC4 encryption function
'# Author: Andreas J”nsson http://www.freevbcode.com/ShowCode.asp?ID=4398
'# RC4 is a stream cipher designed by Rivest for RSA Security.
'#
'##############################################################
Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
    On Error Resume Next
    
    Dim rb(0 To 255) As Integer, X As Long, Y As Long, z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
    
    If Len(Password) = 0 Then
        Exit Function
    End If
    If Len(Expression) = 0 Then
        Exit Function
    End If
    
    If Len(Password) > 256 Then
        Key() = StrConv(Left$(Password, 256), vbFromUnicode)
    Else
        Key() = StrConv(Password, vbFromUnicode)
    End If
    
    For X = 0 To 255
        rb(X) = X
    Next X
    
    X = 0
    Y = 0
    z = 0
    For X = 0 To 255
        Y = (Y + rb(X) + Key(X Mod Len(Password))) Mod 256
        Temp = rb(X)
        rb(X) = rb(Y)
        rb(Y) = Temp
    Next X
    
    X = 0
    Y = 0
    z = 0
    ByteArray() = StrConv(Expression, vbFromUnicode)
    For X = 0 To Len(Expression)
        Y = (Y + 1) Mod 256
        z = (z + rb(Y)) Mod 256
        Temp = rb(Y)
        rb(Y) = rb(z)
        rb(z) = Temp
        ByteArray(X) = ByteArray(X) Xor (rb((rb(Y) + rb(z)) Mod 256))
    Next X
    
    RC4 = StrConv(ByteArray, vbUnicode)
    
End Function

'##############################################
'Used for testing only during development work
'REMOVE from production database
'##############################################
Public Function TestRC4(Original As String)

Dim Encrypted As String
Dim Decrypted As String

Debug.Print "Org: " & Original

Encrypted = RC4(Original, "RC4_Key")
Debug.Print "Enc: " & Encrypted
'
Decrypted = RC4(Encrypted, "RC4_Key")
Debug.Print "Dec: " & Decrypted

End Function

Sub ZZZz()

'##############################################
'Used for testing only during development work
'REMOVE from production database
'##############################################

    TestRC4 ("123456")

End Sub

Public Function SetDefaultPwd()

    SetDefaultPwd = RC4("Not set", "RC4_Key")
End Function


