Attribute VB_Name = "mdlConnect"
Option Compare Database
Option Explicit

Public Function ReadConfig(path As String) As Object
    Dim fso As Object, ts As Object
    Dim line As String, parts() As String
    Dim config As Object

    Set config = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1) ' ForReading = 1
        Do While Not ts.AtEndOfStream
            line = Trim(ts.ReadLine)
            If line <> "" And InStr(line, "=") > 0 Then
                parts = Split(line, "=")
                config(Trim(parts(0))) = Trim(parts(1))
            End If
        Loop
        ts.Close
    End If

    Set ReadConfig = config
End Function

Public Function ConnectionStatus() As Boolean
    Dim ADOConn As New ADODB.Connection
    Dim ConnStatus As Integer
    Dim config As Object
    Dim connString As String
    Dim configPath As String

    ConnectionStatus = False
    On Error Resume Next
    DoCmd.Hourglass True

    configPath = CurrentProject.path & "\config.txt"
    Set config = ReadConfig(configPath)

    If config.Exists("DRIVER") Then
        connString = "Driver={" & config("DRIVER") & "};" & _
                     "Server=" & config("SERVER") & ";" & _
                     "Database=" & config("DATABASE") & ";" & _
                     "User=" & config("USER") & ";" & _
                     "Password=" & config("PASSWORD") & ";" & _
                     "Option=" & config("OPTION") & ";"
        ADOConn.ConnectionString = connString
        ADOConn.ConnectionTimeout = 15
        ADOConn.Open
        DoCmd.Hourglass False

        ConnStatus = ADOConn.State
        If ConnStatus = 1 Then
            ConnectionStatus = True
        End If
    End If
End Function


