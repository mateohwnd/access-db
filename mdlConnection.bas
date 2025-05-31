Attribute VB_Name = "mdlConnection"
' mdlConnection (Standard Module)
' This module centralizes your ADODB connection logic.
Option Compare Database
Option Explicit

Public g_conn As ADODB.Connection ' Global ADODB Connection object

Public Sub OpenDBConnection()
    On Error GoTo Err_Handler
    
    If g_conn Is Nothing Then
        Set g_conn = New ADODB.Connection
    End If
    
    If g_conn.State = adStateClosed Then
        With g_conn
            ' --- IMPORTANT: Customize your connection string ---
            ' Ensure you have the correct MySQL ODBC Driver installed on client machines.
            ' The driver name (e.g., "MySQL ODBC 8.0 Unicode Driver") must match your installed driver.
            .Provider = "MSDASQL" ' ODBC Driver for MySQL
            .ConnectionString = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
                                "SERVER=localhost;" & _
                                "DATABASE=tkd_db;" & _
                                "UID=root;" & _
                                "PWD=123456;" & _
                                "OPTION=3;" ' Common option for prepared statements, etc.
            .Open
        End With
    End If
    
    Exit Sub
    
Err_Handler:
    Set g_conn = Nothing ' Ensure connection is nulled on error
    Call mdlErrorHandling.HandleError("mdlConnection.OpenDBConnection", "Failed to open database connection.")
    MsgBox "Failed to connect to the database. The application will now close.", vbCritical
    Application.Quit ' Crucial: Exit if database connection cannot be established
End Sub

Public Sub CloseDBConnection()
    On Error Resume Next ' In case connection is already closed or invalid
    If Not g_conn Is Nothing Then
        If g_conn.State = adStateOpen Then
            g_conn.Close
        End If
        Set g_conn = Nothing
    End If
    On Error GoTo 0
End Sub



