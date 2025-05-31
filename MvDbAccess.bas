Attribute VB_Name = "MvDbAccess"
Option Compare Database

Public Function MvExistsById(omv As Mv, sTxtInput As String) As Boolean
On Error GoTo HANDLE_ERROR
    Dim rs As Recordset
    Dim sql As String
    
    sql = "select * from masterlist where bltfn='" & sTxtInput & "' or platn='" & sTxtInput & "' or cocn='" & sTxtInput & "';"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.recordCount > 0 Then
        Set omv = New Mv
        
        omv.Customer = Nz(rs!POLN, "")
        omv.MvFileNumber = Nz(rs!BLTFN, "")
        omv.PlateNumber = Nz(rs!PLATN, "")
        omv.cocNumber = Nz(rs!COCN, "")
        omv.Name = Nz(rs!NAM, "")
        omv.Address = Nz(rs!ADDR, "")
        omv.ChassisNumber = Nz(rs!SERCHAN, "")
        omv.EngineNumber = Nz(rs!MOTN, "")
        omv.BodyType = Nz(rs!TOB, "")
        
        omv.DateIssued = Nz(rs!DATIS, #1/1/1900#)
                
        omv.DateCoverStart = Nz(rs!POIF, #1/1/1900#)
        omv.DateCoverEnd = Nz(rs!POIT, #1/1/1900#)
        omv.amount = Nz(rs!PREM, 0)
        
        MvExistsById = True
    End If
    
    Set rs = Nothing
    
HANDLE_ERROR:
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR-MvExistsById"
    End If
    
    Exit Function
End Function
