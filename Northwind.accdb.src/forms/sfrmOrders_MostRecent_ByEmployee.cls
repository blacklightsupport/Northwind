Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub OrderID_Click()
10        On Error GoTo Err_Handler

20        OpenOrderDetailsForm Me.OrderID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "OrderID_Click", True
50        Resume Exit_Handler
End Sub
