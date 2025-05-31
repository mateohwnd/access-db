Attribute VB_Name = "Reporting"
Option Compare Database
Option Explicit

Private sSearchText, sSql, sHeader As String

Property Let sql(value As String)
    sSql = value
End Property
Property Get sql() As String
    sql = sSql
End Property
Property Let Header(value As String)
    sHeader = value
End Property
Property Get Header() As String
    Header = sHeader
End Property

Public Function GetHeader() As String
    GetHeader = Header
End Function

Function GetSearchText()
    GetSearchText = sSearchText
End Function

Sub SetKeystroke(cSearchTextBox As Control)
    If Not IsNull(cSearchTextBox.text) Then
        sSearchText = cSearchTextBox.text
    Else
        sSearchText = Nz(cSearchTextBox.text, "") 'backspace as empty
    End If
End Sub

Function NumberOfRecord(sTableName As String, Optional sWhereClause As String = "") As Integer
On Error GoTo HANDLE_ERROR
    Dim rs As Recordset
    Dim sql As String
    
    sql = "select count(*) as NumberOfRecord from " & sTableName
    
    If sWhereClause <> "" Then sql = sql & " " & sWhereClause & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.recordCount > 0 Then NumberOfRecord = rs!NumberOfRecord
    
    Set rs = Nothing
    
HANDLE_ERROR:
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR-NumberOfRecord"
    End If
    
    Exit Function
End Function

