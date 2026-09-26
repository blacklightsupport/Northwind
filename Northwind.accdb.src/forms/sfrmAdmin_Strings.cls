Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterDelConfirm(Status As Integer)
10        On Error GoTo Err_Handler

          'Not needed; form does not allow deletions.  Me.Parent.UpdateSubLabels

Exit_Handler:
20        Exit Sub

Err_Handler:
30        clsErrorHandler.HandleError Me.Name, "Form_AfterDelConfirm", True
40        Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

          'Not needed; form does not allow additions.  Me.Parent.UpdateSubLabels

Exit_Handler:
20        Exit Sub

Err_Handler:
30        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
40        Resume Exit_Handler
End Sub
