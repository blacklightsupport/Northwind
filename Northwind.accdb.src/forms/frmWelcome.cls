Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdGetStarted_Click()
10        On Error GoTo Err_Handler

20        DoCmd.Close acForm, Me.Name

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdGetStarted_Click", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Close()
10        On Error GoTo Err_Handler

          'Save the value of the checkbox to the SystemSettings table.
20        SaveSystemSetting ssShowWelcome, IIf(Me.chkShowWelcome, -1, 0)        'Ensure this works in international scenarios.

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Close", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Me.chkShowWelcome = GetSystemSetting(ssShowWelcome)

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Open", True
50        Resume Exit_Handler
End Sub
