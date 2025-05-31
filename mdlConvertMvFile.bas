Attribute VB_Name = "mdlConvertMvFile"
Option Compare Database
Option Explicit

Public Function ConvertMvFile(ByVal strMV As String)

    Dim X As String
    Dim Y() As String
    Dim a As String '
    Dim b As String '
    Dim C As Long
    Dim d As String
    Dim e As Long
    Dim f As String
    Dim i As Long
    Dim z As Long

    'terminate if strMV has no value
    If Len(strMV) = 0 Then 'Or Len(strMV) > 15
        Exit Function
    ElseIf Len(strMV) > 14 Then 'if mvfile length is greater than 14 then simplified.
        a = strMV
        b = Left(a, 4) 'get the first 4 characters to the left
        C = Mid(a, 5) 'get characters starting to 5 position
        e = CStr(val(C)) 'remove trailing zeros
        f = b & "-" & e 'concatenate
        ConvertMvFile = (f)
        Exit Function
    End If
    
    
    'mv file simplification procedure started.
    
    C = InStr(5, strMV, " ")
    X = Trim(Nz(strMV))
    
    d = Mid(strMV, 5, 1)
    
'    Debug.Print d
    
    Select Case d
        Case "-"
            Y = Split(X, "-") ' - line as separator
            'MsgBox "-"
        Case " "
            Y = Split(X, " ", 2) ' space as separator
            'MsgBox "space"
        Case Else
            MsgBox "Error, Wrong MV File format.", vbCritical, "Error"
            Exit Function
    End Select
    
'    If c > 0 Then ' there is a space
'        y = Split(x, " ", 2) ' space as separator
'    ElseIf c = 0 Then
'        y = Split(x, "-") ' - line as separator
'    End If
    
    a = Y(0)
    b = Y(1)
    z = Len(a & b)

'    Debug.Print a
'    Debug.Print b
'    Debug.Print z

    For i = 1 To z
        a = a & "0"
        If Len(a & b) > 14 Then Exit For '15 characters only for mv file format
        Debug.Print i & " " & a
    Next

    strMV = a & b
    ConvertMvFile = (strMV)

    'Debug.Print "mv file number: " & a & b ' & vbCrLf & Len(a & b)

End Function
