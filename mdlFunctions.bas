Attribute VB_Name = "mdlFunctions"
Option Compare Database
Option Explicit

Public strProc As String

'Assorted functions used in form footers

Function GetProgramName()

    GetProgramName = Nz(DLookup("ItemValue", "tblSettings", "ItemName='ProgramName'"), "")
End Function

Function GetVersion()

    GetVersion = Nz(DLookup("ItemValue", "tblSettings", "ItemName='Version'"), "")
End Function

Function GetVersionDate()

    GetVersionDate = Nz(DLookup("ItemValue", "tblSettings", "ItemName='VersionDate'"), "")
End Function

Function GetCopyright()

    GetCopyright = Nz(DLookup("ItemValue", "tblSettings", "ItemName='Copyright'"), "")
End Function

Function GetWebsite()

    GetWebsite = Nz(DLookup("ItemValue", "tblSettings", "ItemName='ProgramWebsite'"), "")
End Function


