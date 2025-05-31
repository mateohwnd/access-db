Attribute VB_Name = "env"
Option Compare Database

Public lngUserID As Long
Public strUserName As String
Public strEncoderCode As String
Public strComputerName As String
Public strPassword As String
Public strEmail As String
Public intAccessLevel As Integer
Public blnChangeOwnPassword As Boolean


Public g_originalValues As Object ' Scripting.Dictionary to cache original DB values
Public g_IsRecordLoaded As Boolean ' Flag to indicate if original values are cached
Public g_currentPlateN As String


' In a Standard Module (e.g., env)
#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pGuid As Any) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As Any) As Long
#End If

' Type definition for the GUID structure
Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' Function to generate a GUID (also in a standard module)
Public Function GenerateGuid() As String
    Dim TGuid As GUID_TYPE
    Dim RetVal As Long
    Dim strGuid As String

    RetVal = CoCreateGuid(TGuid)

    If RetVal = 0 Then ' S_OK
        strGuid = String(32, "0") ' Pre-allocate string for efficiency
        strGuid = Right("00000000" & Hex(TGuid.Data1), 8) & _
                  Right("0000" & Hex(TGuid.Data2), 4) & _
                  Right("0000" & Hex(TGuid.Data3), 4) & _
                  Right("00" & Hex(TGuid.Data4(0)), 2) & _
                  Right("00" & Hex(TGuid.Data4(1)), 2) & _
                  Right("00" & Hex(TGuid.Data4(2)), 2) & _
                  Right("00" & Hex(TGuid.Data4(3)), 2) & _
                  Right("00" & Hex(TGuid.Data4(4)), 2) & _
                  Right("00" & Hex(TGuid.Data4(5)), 2) & _
                  Right("00" & Hex(TGuid.Data4(6)), 2) & _
                  Right("00" & Hex(TGuid.Data4(7)), 2)
        ' Optionally add hyphens if your database expects them (MySQL usually doesn't need them for UUIDs)
        ' GenerateGuid = Left(strGuid, 8) & "-" & Mid(strGuid, 9, 4) & "-" & Mid(strGuid, 13, 4) & "-" & Mid(strGuid, 17, 4) & "-" & Right(strGuid, 12)
        GenerateGuid = strGuid ' Return without hyphens for typical MySQL UUID storage
    Else
        GenerateGuid = "" ' Return empty string on failure
    End If
End Function

