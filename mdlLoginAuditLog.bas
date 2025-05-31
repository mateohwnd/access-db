Attribute VB_Name = "mdlLoginAuditLog"
Option Compare Database
Option Explicit

Public lngLoginID As Long

Function LogMeIn(strUserName As Long)
'Go to the users table and record that the user has logged in and which computer they have logged in from

    CurrentDb.Execute "UPDATE tblUsers SET LoggedIn = True, Computer = GetComputerName()" & _
        " WHERE UserName='" & GetUserName & "' AND tblUsers.Active=True;"
    
End Function

Function LogMeOff(strUserName As Long)
'Go to the users table and record that the user has logged out

    CurrentDb.Execute "UPDATE tblUsers SET LoggedIn = False, Computer = ''" & _
        " WHERE UserName='" & GetUserName & "';"

End Function

Function CreateSession(LoginID As Long)

'This function records the details regarding the login details of the person
'Get the new loginID
'v5 21/11/2018 - added Nz to manage case where no record exists
lngLoginID = Nz(DMax("LoginID", "tblLoginSessions") + 1, 1)

CurrentDb.Execute "INSERT INTO tblLoginSessions ( LoginID, UserName, LoginEvent, ComputerName )" & _
    " VALUES(GetLoginID(), GetUserName(), Now(), GetComputerName());"

End Function

Function CloseSession()

'This closes the open session
    'set logout date/timein tblLoginSessions
    CurrentDb.Execute "UPDATE tblLoginSessions SET LogoutEvent = Now()" & _
        " WHERE LoginID= " & GetLoginID & ";"
    
    'clear user login in tblUsers
    CurrentDb.Execute "UPDATE tblUsers SET LoggedIn = False, Computer = Null" & _
        " WHERE UserName= '" & GetUserName & "';"

End Function


