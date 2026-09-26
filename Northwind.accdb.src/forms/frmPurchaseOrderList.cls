Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAddPO_Click()
10        On Error GoTo Err_Handler

20        OpenPurchaseOrderDetailsForm

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdAddPO_Click", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

          'Clean up anything left by a previous session

20        Me.FilterOn = False
30        Me.Filter = ""

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_Open", True
60        Resume Exit_Handler
End Sub

Private Sub PurchaseOrderID_Click()
10        On Error GoTo Err_Handler

20        OpenPurchaseOrderDetailsForm Me.PurchaseOrderID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "PurchaseOrderID_Click", True
50        Resume Exit_Handler
End Sub
