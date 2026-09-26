Attribute VB_Name = "modTableDataMacros"
Option Compare Database
Option Explicit

'PURPOSE:
'   Adds data macros if they don't exist.
'   The reason we even have to do this adding of data macros is because Access templates cannot have data macros
'   that reference custom VBA functions such as our GetAuditFieldsUserName.
Public Sub AddDataMacros()
10        On Error GoTo Err_Handler

          Dim td                  As DAO.TableDef
          Dim strFileName         As String

20        For Each td In g_dbApp().TableDefs
30            If td.Attributes And dbSystemObject Then
                  'Skip system objects like MSys* and USys*
40            Else
50                If HasField(td, "AddedBy") And HasField(td, "AddedOn") And HasField(td, "ModifiedBy") And HasField(td, "ModifiedOn") Then
60                    If HasDataMacro(td) Then
                          'Already has data macro; nothing to do.
70                    Else
                          'Save macro to temp file so we can LoadFromText.
80                        If strFileName = "" Then
                              'File not yet created. Do it now.
90                            strFileName = Environ("TEMP") & "\datamacro.xml"
100                           StringToFile strFileName, DLookup("DataMacro", "Welcome")
110                       End If

120                       Application.LoadFromText acTableDataMacro, td.Name, strFileName
130                   End If
140               End If
150           End If
160       Next td

Exit_Handler:
170       Exit Sub

Err_Handler:
180       clsErrorHandler.HandleError "modTableDataMacros", "AddDataMacros"
190       Resume Exit_Handler
End Sub

'NOTE:
'   If you split the database into a Front End and Back End
'   this module needs to be where the tables are (the Back End)
'   because these function(s) are called by data macros in the tables.
'PURPOSE:
'   Provide name to be used when saving to the Audit Trail fields AddedBy and ModifiedBy which are present in most tables.
'   This function is called from the BeforeChange data macro for most tables.
Public Function GetAuditFieldsUserName() As String
10        On Error GoTo Err_Handler

          Dim strUserName         As String

20        If Get_UserID() = 0 Then
              'User is not logged in.
30            strUserName = GetWindowsUserName()
40        Else
50            strUserName = DLookup("FullNameFNLN", "qryEmployees", "EmployeeID = " & Get_UserID())
60        End If
70        GetAuditFieldsUserName = strUserName

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError "modTableDataMacros", "GetAuditFieldsUserName"
100       Resume Exit_Handler
End Function
