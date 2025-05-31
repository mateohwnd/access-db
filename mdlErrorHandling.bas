Attribute VB_Name = "mdlErrorHandling"
Option Compare Database
Option Explicit

' Error Handler with Timestamped Logs and Optional SQL Context

Public Sub HandleError( _
    ByVal sSource As String, _
    Optional ByVal sCustomMessage As String = "", _
    Optional ByVal sSQLContext As String = "" _
)

    Dim sErrMsg As String
    Dim sErrNum As String
    Dim sErrDesc As String
    Dim sLogPath As String
    Dim sUser As String
    Dim sProject As String

    ' Get error details
    sErrNum = CStr(Err.Number)
    sErrDesc = Err.Description
    sUser = CurrentUser()
    sProject = CurrentProject.Name

    ' Build the error message
    sErrMsg = "An error occurred in: " & sSource & vbCrLf & _
              "Project: " & sProject & vbCrLf & _
              "User: " & sUser & vbCrLf & _
              "Error Number: " & sErrNum & vbCrLf & _
              "Description: " & sErrDesc

    ' Add custom message if provided
    If sCustomMessage <> "" Then
        sErrMsg = sErrMsg & vbCrLf & "Details: " & sCustomMessage
    End If

    ' Add SQL context if provided
    If sSQLContext <> "" Then
        sErrMsg = sErrMsg & vbCrLf & "SQL Context: " & sSQLContext
    End If

    ' --- Error Logging ---
    On Error Resume Next ' Prevent file logging crash

    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create timestamped filename: ErrorLog_yyyymmdd.txt
    sLogPath = CurrentProject.path & "\ErrorLog_" & Format(Now, "yyyymmdd") & ".txt"

    Set ts = fso.OpenTextFile(sLogPath, 8, True) ' 8 = ForAppending
    ts.WriteLine String(40, "-")
    ts.WriteLine "Date/Time: " & Now()
    ts.WriteLine sErrMsg
    ts.WriteLine String(40, "-")
    ts.Close

    Set ts = Nothing
    Set fso = Nothing
    On Error GoTo 0 ' Resume normal error handling

    ' Optional: display message box
    MsgBox sErrMsg, vbCritical, "Application Error"

    ' Clear error object
    Err.Clear

End Sub


