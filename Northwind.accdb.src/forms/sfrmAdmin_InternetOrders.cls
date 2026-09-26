Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const MinCount              As Integer = 1
Const MaxCount              As Integer = 5

Private Sub cmdCreate_Click()
10        On Error GoTo Err_Handler

          'Validation
20        If IsNull(Me.txtCount) Then
30            MsgBox GetString(sNumberBetween, MinCount, MaxCount), vbExclamation
40        Else
50            CreateRandomOrders Me.txtCount
60            Me.Parent.UpdateSubLabels

70            Ribbon_RefreshMRU

80            MsgBox "Orders created. Refresh the Order List form to see them.", vbInformation
90        End If

Exit_Handler:
100       Exit Sub

Err_Handler:
110       clsErrorHandler.HandleError Me.Name, "cmdCreate_Click", True
120       Resume Exit_Handler
End Sub

Private Sub txtCount_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

          'Validation
20        If IsNull(Me.txtCount) Or Me.txtCount < MinCount Or Me.txtCount > MaxCount Then
30            MsgBox GetString(sNumberBetween, MinCount, MaxCount), vbExclamation
40            Cancel = True       'Prevents the user from exiting the textbox with this bad value.
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "txtCount_BeforeUpdate", True
80        Resume Exit_Handler
End Sub
