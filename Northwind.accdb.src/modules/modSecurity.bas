Attribute VB_Name = "modSecurity"
Option Compare Database
Option Explicit

'PURPOSE:
'   Previously, returns true if the current user has been assigned any privilege.
'   Currently, only one privilege exists: approve Purchase Orders
'   If additional privileges are created, additional validation will be required
'   Check specific privilege requested
'
Public Function HasPrivilege(ByVal p As enumPrivileges) As Boolean
10        On Error GoTo Err_Handler

20        HasPrivilege = Not IsNull(DLookup("EmployeePrivilegeID", "EmployeePrivileges", "EmployeeID = " & Get_UserID() & " AND PrivilegeID = " & p))

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modSecurity", "HasPrivilege"
50        Resume Exit_Handler
End Function
