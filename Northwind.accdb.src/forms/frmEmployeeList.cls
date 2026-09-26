Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'This form is in Split Form view, showing both details of the current record, and a list.

Private Sub cboSupervisorID_Enter()
10        On Error GoTo Err_Handler

          'In order to limit the list of Employees who can be assigned as supervisors
          'the combo box is filtered to exclude the current employee's EmployeeID
          'This rowsource must be refreshed for each record when it becomes current
20        Me.cboSupervisorID.Requery

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cboSupervisorID_Enter", True
50        Resume Exit_Handler
End Sub

Private Sub cboSupervisorID_NotInList(NewData As String, Response As Integer)
10        On Error GoTo Err_Handler

          'We need to do two things
          'Suppress the default error message from Acess and replace it with a custom message
          'Allow the user to add a new employee record for the person who will become the new supervisor.

20        Response = acDataErrContinue
30        With Me
40            .cboSupervisorID.Undo
50            If MsgBox(GetString(sNewEmployee, NewData), vbExclamation Or vbOKCancel, "Add " & NewData & "?") = vbOK Then
60                .Dirty = False
70                .Recordset.AddNew
80            End If
90        End With

Exit_Handler:
100       Exit Sub

Err_Handler:
110       clsErrorHandler.HandleError Me.Name, "cboSupervisorID_NotInList", True
120       Resume Exit_Handler
End Sub

Private Sub cmdEmailEmployee_Click()

10        On Error GoTo Err_Handler

          Dim strSender           As String
          Dim strWindowsUserName  As String

20        strWindowsUserName = GetWindowsUserName()
30        strSender = DLookup("FirstName", "Employees", StringFormatSQL("Employees.WindowsUserName = {0}", strWindowsUserName))
40        strSender = vbCrLf & vbCrLf & vbCrLf & strSender & " " & DLookup("LastName", "Employees", StringFormatSQL("Employees.WindowsUserName = {0}", strWindowsUserName))
50        DoCmd.SendObject ObjectType:=acSendNoObject, To:=Nz(Me.EmailAddress, ""), Subject:="NorthWind Traders: ", MessageText:=strSender, editMessage:=True

Exit_Handler:
60        Exit Sub

Err_Handler:
70        If Err <> 2296 Then    'Error 2296 indicates the email password is invalid; the message can't be sent.
80            clsErrorHandler.HandleError Me.Name, "cmdEmailEmployee_Click"
90        End If
100       Resume Exit_Handler
End Sub

Private Sub cmdNewEmployee_Click()
10        On Error GoTo Err_Handler

20        If Not Me.NewRecord Then DoCmd.GoToRecord record:=acNewRec

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdNewEmployee_Click", True
50        Resume Exit_Handler
End Sub

'RETURNS
'   True if the employee can be deleted (no child records exist in four related tables); False otherwise.
'NOTE:
'  This is a private function, used only in this form
Private Function EmployeeCanBeDeleted(ByVal lngEmployeeID As Long) As Boolean
10        On Error GoTo Err_Handler

          Dim blnReturn           As Boolean
          Dim lngEmployeeOrders   As Long
          Dim lngEmployeeSupervisor As Long
          Dim lngEmployeePrivileges As Long
          Dim lngEmployeePurchaseOrders As Long
          Dim sCantDeleteResponse As String
          Dim strOpenArgs         As String
          Dim strMsgCaption       As String

20        blnReturn = True
30        strMsgCaption = "Message from Employees"

40        With Me
50            lngEmployeeOrders = DCount("OrderID", "Orders", StringFormatSQL("EmployeeID ={0}", lngEmployeeID))
60            lngEmployeeSupervisor = DCount("EmployeeID", "Employees", StringFormatSQL("SupervisorID ={0}", lngEmployeeID))
70            lngEmployeePrivileges = DCount("EmployeeID", "EmployeePrivileges", StringFormatSQL("EmployeeID ={0}", lngEmployeeID))
80            lngEmployeePurchaseOrders = DCount("PurchaseOrderID", "PurchaseOrders", StringFormatSQL("SubmittedByID ={0}", lngEmployeeID))

90            If lngEmployeeOrders > 0 Then
100               sCantDeleteResponse = GetString(sEmployeeRestraints, .FullNameFNLN, "has Orders (" & lngEmployeeOrders & ").", "Employees with orders can not be deleted")
110           End If
120           If lngEmployeePurchaseOrders > 0 Then
130               sCantDeleteResponse = sCantDeleteResponse & vbCrLf & vbCrLf & GetString(sEmployeeRestraints, .FullNameFNLN, "has Purchase Orders (" & lngEmployeePurchaseOrders & ").", "Employees with purchase orders can not be deleted")
140           End If
150           If lngEmployeeSupervisor > 0 Then
160               sCantDeleteResponse = sCantDeleteResponse & vbCrLf & vbCrLf & GetString(sEmployeeRestraints, .FullNameFNLN, "is a Supervisor.", "Employees supervising other employees can not be deleted      <br />      until their employees are reassigned")
170           End If
180           If lngEmployeePrivileges > 0 Then
190               sCantDeleteResponse = sCantDeleteResponse & vbCrLf & vbCrLf & GetString(sEmployeeRestraints, .FullNameFNLN, "has elevated Privileges.", "Employees with elevated privileges can not be deleted     <br />      until those privileges are revoked")
200           End If
210           blnReturn = (lngEmployeeOrders + lngEmployeeSupervisor + lngEmployeePrivileges + lngEmployeePurchaseOrders = 0)
220           If blnReturn = True Then
                  ' It's safe to remove any of this Employee's records from MRU if there are no supporting records in other tables.
230               RemoveFromMRU "Orders", Nz(DLookup("PkValue", "MRU", "EmployeeID = " & lngEmployeeID), 0)
240               RemoveFromMRU "PurchaseOrders", Nz(DLookup("PkValue", "MRU", "EmployeeID = " & lngEmployeeID), 0)
                  ' Capture the current employee's name for the deletion confirmation message box
                  ' If multiple employees are deleted in bulk, append all names
250               TempVars.Add Name:="EmployeeToDelete", Value:=TempVars!EmployeeToDelete & Chr(13) & Me.FullNameFNLN.Value
260           Else
270               sCantDeleteResponse = Left(sCantDeleteResponse, Len(sCantDeleteResponse) - 6)    'Remove the trailing <br />
280               strOpenArgs = StringFormat("Header={0}&Message={1}&Actions={2}&FormCaption={3}", "Employee Delete Not Allowed", sCantDeleteResponse, "vbOK", strMsgCaption)
290               DoCmd.OpenForm "frmGenericDialog", acNormal, , , acFormReadOnly, acDialog, strOpenArgs
                  'Defensive programming: currently the dialog would be closed in this case, but that behavior could conceivably be changed.
300               If IsFormOpen("frmGenericDialog") Then
310                   DoCmd.Close acForm, "frmGenericDialog", acSaveNo
320               End If
330           End If
340       End With

350       EmployeeCanBeDeleted = blnReturn

Exit_Handler:
360       Exit Function

Err_Handler:
370       clsErrorHandler.HandleError Me.Name, "EmployeeCanBeDeleted"
380       Resume Exit_Handler
End Function

Private Sub Form_AfterDelConfirm(Status As Integer)
10        On Error GoTo Err_Handler

20        Me.cmdNewEmployee.Visible = Not Me.NewRecord

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterDelConfirm", True
50        Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

          Dim lngNewEmployeeID    As Long

20        With Me
30            lngNewEmployeeID = .EmployeeID
40            DoCmd.OpenForm FormName:="sfrmEmployee_Privileges", _
                             View:=acNormal, _
                             WhereCondition:="EmployeeID = " & lngNewEmployeeID, _
                             DataMode:=acFormAdd, _
                             WindowMode:=acDialog, _
                             OpenArgs:=lngNewEmployeeID
50        End With

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
80        Resume Exit_Handler
End Sub

Private Sub Form_AfterUpdate()
10        On Error GoTo Err_Handler

20        Call RequeryEmployeeList(Me.EmployeeID)
30        ValidateForm_RemoveHighlights Me

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_AfterUpdate", True
60        Resume Exit_Handler
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
10        On Error GoTo Err_Handler

          '  TempVars!EmployeeToDelete returns the correctly identified employee or employees in the confirmation messagebox
20        If MsgBox(GetString(sDeleteEmployee, TempVars!EmployeeToDelete), vbQuestion Or vbYesNo, "Delete Employee?") = vbNo Then
30            Cancel = True
40        End If
50        TempVars.Add Name:="EmployeeToDelete", Value:=""
60        Response = acDataErrContinue

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "Form_BeforeDelConfirm", True
90        Resume Exit_Handler
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Cancel = ValidateForm(Me)

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Current()
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me
30        With Me
40            .cmdNewEmployee.Visible = Not .NewRecord
50            .cboSupervisorID.Requery
60            .lblRecentOrders.Caption = "Recent Orders for " & .FullNameFNLN
70        End With

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "Form_Current", True
100       Resume Exit_Handler
End Sub

Private Sub Form_Delete(Cancel As Integer)
10        On Error GoTo Err_Handler

          'Before allowing an employee to be deleted, confirm that there are no child records in one of the four child record tables
          'Also replace the default Access error message with a custom message if the employee can't be deleted
20        Cancel = Not EmployeeCanBeDeleted(Me.EmployeeID)

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Delete", True
50        Resume Exit_Handler
End Sub

'PURPOSE:
'   Requery the form and go back to the record we were on (otherwise we would be on record 1).
Private Sub RequeryEmployeeList(ByVal lngEmployeeID As Long)
10        On Error GoTo Err_Handler

          Dim rst                 As DAO.Recordset

20        With Me
30            .Requery
40            Set rst = .RecordsetClone
50            rst.FindFirst "EmployeeID = " & lngEmployeeID
60            .Bookmark = rst.Bookmark
70        End With
80        Set rst = Nothing

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError Me.Name, "RequeryEmployeeList"
110       Resume Exit_Handler
End Sub
