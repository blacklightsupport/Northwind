Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdResetDates_Click()
10        On Error GoTo Err_Handler

20        SetDatesToCurrent

30        MsgBox "Done. Reopen your forms to see the new dates.", vbInformation

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "cmdResetDates_Click", True
60        Resume Exit_Handler
End Sub
