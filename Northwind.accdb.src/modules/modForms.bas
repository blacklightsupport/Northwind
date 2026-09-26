Attribute VB_Name = "modForms"
Option Compare Database
Option Explicit

'PURPOSE:
'   Return the CurrentView property of the control's form.
'ALGORITHM:
'   Take into account that if the control is on a Tab control, its parent is a Tab Page, not the form.
Public Function ControlCurrentView(ctl As Access.control) As Integer
10        On Error GoTo Err_Handler

          Dim intReturn           As Integer

20        If TypeOf ctl.Parent Is Page Then
30            intReturn = ctl.Parent.Parent.Parent.CurrentView
40        Else
50            intReturn = ctl.Parent.CurrentView
60        End If
70        ControlCurrentView = intReturn

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError "modForms", "ControlCurrentView"
100       Resume Exit_Handler
End Function

'PURPOSE:
'   Return the DefaultView property of the control's form.
'ALGORITHM:
'   Take into account that if the control is on a Tab control, its parent is a Tab Page, not the form.
Public Function ControlDefaultView(ctl As Access.control) As Integer
10        On Error GoTo Err_Handler

          Dim intReturn           As Integer

20        If TypeOf ctl.Parent Is Page Then
30            intReturn = ctl.Parent.Parent.Parent.DefaultView
40        Else
50            intReturn = ctl.Parent.DefaultView
60        End If
70        ControlDefaultView = intReturn

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError "modForms", "ControlDefaultView"
100       Resume Exit_Handler
End Function

'PURPOSE:
'   Returns if the given form is open or not.
Public Function IsFormOpen(ByVal strFormName As String) As Boolean
10        On Error GoTo Err_Handler

20        IsFormOpen = (SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0)

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modForms", "IsFormOpen"
50        Resume Exit_Handler
End Function

Public Sub RequeryListForms()
10        On Error GoTo Err_Handler
          'Using the Recordset.Requery approach allows allows the form to keep its place
          'Using the Requery approach takes the user to the first record in the list
20        If IsFormOpen("frmCompanyList") Then Forms!frmCompanyList.Recordset.Requery
30        If IsFormOpen("frmEmployeeList") Then Forms!frmEmployeeList.Recordset.Requery
40        If IsFormOpen("frmOrderList") Then Forms!frmOrderList.Recordset.Requery
50        If IsFormOpen("frmProductList") Then Forms!frmProductList.Recordset.Requery
60        If IsFormOpen("frmPurchaseOrderList") Then Forms!frmPurchaseOrderList.Recordset.Requery

Exit_Handler:
70        Exit Sub

Err_Handler:
80        If Err.Number = 3219 Then   '3219 = Invalid operation.
90            Resume Next             'Ignore and continue with the next line.
100       End If
110       clsErrorHandler.HandleError "modForms", "RequeryListForms"
120       Resume Exit_Handler
130       Resume
End Sub

Public Sub RequeryProductList()
10        On Error GoTo Err_Handler

20        If IsFormOpen("frmProductList") Then Forms!frmProductList.Requery

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modForms", "RequeryProductList"
50        Resume Exit_Handler
End Sub
