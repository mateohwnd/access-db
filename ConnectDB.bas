Attribute VB_Name = "ConnectDB"
Option Compare Database
Option Explicit

Public Const gStrConnectxString As String = "Driver={MySQL ODBC 8.0 UNICODE Driver};Server=localhost;Database=tkd_db;User=root;Password=123;Option=3;"
Public dbCon As ADODB.Connection

Public Function ConnectDatabase() As ADODB.Connection

    On Error GoTo ConnectionError
    
    ' Check if connections is already open
    If dbCon Is Nothing Then
        ' instantiate new database connection object
        Set dbCon = New ADODB.Connection
    ElseIf dbCon.State = adStateOpen Then
        ' otherwise return existing ADO connection object
        Set ConnectDatabase = dbCon
        Exit Function
    End If
    
    ' define the connection string
    dbCon.ConnectionString = gStrConnectxString
    
    ' open the connection
    dbCon.Open
    Set ConnectDatabase = dbCon
    Exit Function

ConnectionError:
    MsgBox "Failed to connect to database" & vbCrLf & Err.Description & " (" & Err.Number & ")"
    Set ConnectDatabase = Nothing
    
End Function

Public Function CloseDatabase()

    On Error Resume Next
    ' close the connection
    dbCon.Close
    Set dbCon = Nothing
    
    ' reset the error handler
    On Error GoTo 0
    
End Function
