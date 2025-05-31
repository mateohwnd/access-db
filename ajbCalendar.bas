Attribute VB_Name = "ajbCalendar"
Option Compare Database
Option Explicit

'Calendar form variable:
Public gtxtCalTarget As TextBox 'Text box to return the date from the calendar to.

Public Function CalendarFor(txt As TextBox, Optional strTitle As String)
'On Error GoTo Err_Handler
    'Purpose:   Open the calendar form, identifying the text box to return the date to.
    'Arguments: txt = the text box to return the date to.
    '           strTitle = the caption for the calendar form (passed in OpenArgs).
    
    Set gtxtCalTarget = txt
    DoCmd.OpenForm "frmCalendar", windowmode:=acDialog, OpenArgs:=strTitle
    
Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & " - " & Err.Description, vbExclamation, "CalendarFor()"
    Resume Exit_Handler
End Function

Public Function LogError(lngErr As Long, strDescrip As String, strProc As String, _
    Optional bShowUser As Boolean = True, Optional varParam As Variant)
    'Purpose: Minimal substitute for the real error logger function at:
    '               http://allenbrowne.com/ser-23a.html
    
    If bShowUser Then
        MsgBox "Error " & lngErr & ": " & strDescrip, vbExclamation, strProc
    End If
End Function

