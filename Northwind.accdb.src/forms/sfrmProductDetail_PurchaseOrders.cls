Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub txtPOID_Click()
10        On Error GoTo Err_Handler

20        OpenPurchaseOrderDetailsForm Me.txtPOID

Exit_Handler:

30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "txtPOID_Click"
50        Resume Exit_Handler
60        Resume
End Sub
