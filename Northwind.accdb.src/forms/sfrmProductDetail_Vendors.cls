Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Me.txtModifiedBy = Get_EmployeeFNLN(Get_UserID())
30        Me.txtModifiedOn = Now()

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
60        Resume Exit_Handler
End Sub

Private Sub txtVendorID_Click()
10        On Error GoTo Err_Handler

20        If Not Me.NewRecord Then
              'If the Company Detail form is open let them know and ask them to close it
30            If IsFormOpen("frmCompanyDetail") Then
40                MsgBox GetString(sFormAlreadyOpen, "Company Detail"), vbOKOnly, "Already Open"
50                Exit Sub
60            End If
              'Pass Company ID and CompanyTypeID in the OpenArgs Property
70            DoCmd.OpenForm "frmCompanyDetail", OpenArgs:=StringFormat("CompanyID={0}&CompanyTypeID={1}", Me.VendorID, ctVendor)
80        End If

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError Me.Name, "txtVendorID_Click", True
110       Resume Exit_Handler
End Sub
