Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterUpdate", True
50        Resume Exit_Handler
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

Private Sub txtPrimaryPhone_Click()
10        On Error GoTo Err_Handler

          'If the Phone has not been populated yet
          'postion the cursor to the beginning of the field
          'no matter where the user clicked in the field

20        If Len(Me.txtPrimaryPhone) = 0 Then
30            Me.txtPrimaryPhone.SelStart = 0
40        End If

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "txtPrimaryPhone_Click", True
70        Resume Exit_Handler
End Sub

Private Sub txtSecondaryPhone_Click()
10        On Error GoTo Err_Handler

          'If the Phone has not been populated yet
          'postion the cursor to the beginning of the field
          'no matter where the user clicked in the field

20        If Len(Me.txtSecondaryPhone) = 0 Then
30            Me.txtSecondaryPhone.SelStart = 0
40        End If

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "txtSecondaryPhone_Click", True
70        Resume Exit_Handler
End Sub
