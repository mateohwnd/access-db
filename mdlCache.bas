Attribute VB_Name = "mdlCache"
' === mdlCache (Standard Module) ===
Option Compare Database
Option Explicit

' Module-level cached arrays
Public g_MakeModels() As String
Public g_BodyTypes() As String

' When to Refresh Cache
' Any time you update vehicle makes or body types from a CRUD form, just call:
' Call mdlCache.LoadComboBoxCache

Public Sub LoadComboBoxCache()
    On Error GoTo Err_Handler

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim tempList As Collection
    Dim i As Long ' Declare i as Long for loop counter

    Set tempList = New Collection
    Call mdlConnection.OpenDBConnection ' Ensure this opens g_conn properly

    ' --- Load Make Models ---
    sql = "SELECT DISTINCT `MAKE` FROM `VEHICLE MAKE` ORDER BY `MAKE`"
    Set rs = New ADODB.Recordset
    rs.Open sql, g_conn, adOpenForwardOnly, adLockReadOnly

    Do Until rs.EOF
        If Not IsNull(rs!MAKE) Then
            tempList.Add CStr(rs!MAKE) ' Ensure it's added as a string
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    ' Transfer to array - ONLY if there are items
    If tempList.count > 0 Then
        ReDim g_MakeModels(0 To tempList.count - 1)
        For i = 1 To tempList.count
            g_MakeModels(i - 1) = CStr(tempList(i))
        Next i
    Else
        ' If no items, ensure the array is empty or reset
        ReDim g_MakeModels(0 To -1) ' ReDim to an empty array
    End If

    ' Debug Print for g_MakeModels
'    If IsArray(g_MakeModels) And UBound(g_MakeModels) >= LBound(g_MakeModels) Then
'        Debug.Print "g_MakeModels loaded: " & Join(g_MakeModels, "; ")
'    Else
'        Debug.Print "g_MakeModels is empty or not loaded."
'    End If

    ' --- Repeat for Body Types ---
    Set tempList = New Collection ' Re-initialize tempList for the next set of data
    sql = "SELECT DISTINCT `TYPE OF BODY` FROM `TYPE OF BODY` ORDER BY `TYPE OF BODY`"
    Set rs = New ADODB.Recordset
    rs.Open sql, g_conn, adOpenForwardOnly, adLockReadOnly

    Do Until rs.EOF
        If Not IsNull(rs.fields(0)) Then
            tempList.Add CStr(rs.fields(0)) ' Ensure it's added as a string
        End If
        rs.MoveNext
    Loop

    ' Transfer to array - ONLY if there are items
    If tempList.count > 0 Then
        ReDim g_BodyTypes(0 To tempList.count - 1)
        For i = 1 To tempList.count
            g_BodyTypes(i - 1) = CStr(tempList(i))
        Next i
    Else
        ' If no items, ensure the array is empty or reset
        ReDim g_BodyTypes(0 To -1) ' ReDim to an empty array
    End If

    ' Debug Print for g_BodyTypes
'    If IsArray(g_BodyTypes) And UBound(g_BodyTypes) >= LBound(g_BodyTypes) Then
'        Debug.Print "g_BodyTypes loaded: " & Join(g_BodyTypes, "; ")
'    Else
'        Debug.Print "g_BodyTypes is empty or not loaded."
'    End If

    rs.Close
    Set rs = Nothing

    Exit Sub

Err_Handler:
    ' Ensure mdlErrorHandling and sql are properly defined/accessible
    Call mdlErrorHandling.HandleError("LoadComboBoxCache", "Could not load combo box values.", sql)
End Sub

