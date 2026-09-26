Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_DblClick(Cancel As Integer)
10        On Error GoTo Err_Handler

20        OpenForm_frmOrderDetails

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Detail_DblClick"
50        Resume Exit_Handler
End Sub

Public Function OpenForm_frmOrderDetails()
10        On Error GoTo Err_Handler

20        OpenOrderDetailsForm Me.OrderID

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "OpenForm_frmOrderDetails", True
50        Resume Exit_Handler
End Function
