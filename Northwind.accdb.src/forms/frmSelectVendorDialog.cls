Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()
10        On Error GoTo Err_Handler

20        DoCmd.Close acForm, Me.Name

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdCancel_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdOK_Click()
10        On Error GoTo Err_Handler

20        If Me.lstVendor.Value > 0 Then
30            Me.Visible = False
40        Else
50            Beep
60            Me.txtMessage.ForeColor = vbRed
70            Me.txtMessage = "Select a Vendor, then OK (or Cancel)"
80            Me.lstVendor.SetFocus
90        End If

Exit_Handler:
100       Exit Sub

Err_Handler:
110       clsErrorHandler.HandleError Me.Name, "cmdOK_Click", True
120       Resume Exit_Handler
End Sub

Private Sub Form_Load()
10        On Error GoTo Err_Handler

          '==================================================================
          '   OpenArgs: <Dialog header caption>|<Message or prompt>|ProductID
          '   ALTERNATE IMPLEMENTATION: use StringToDictionary
          '==================================================================

          Dim arArgs()            As String

20        If Nz(Me.OpenArgs, "") > "" Then
30            arArgs = Split(Me.OpenArgs, "|")
40            Me.Auto_Header0.Caption = arArgs(0)
50            Me.txtMessage = arArgs(1)
60            Me.txtProductID = arArgs(2)
70            Me.lstVendor.Requery
80        End If

90        Me.Visible = True
100       Me.Modal = True

Exit_Handler:
110       Exit Sub

Err_Handler:
120       clsErrorHandler.HandleError Me.Name, "Form_Load", True
130       Resume Exit_Handler
End Sub

Private Sub lstVendor_DblClick(Cancel As Integer)
10        On Error GoTo Err_Handler

20        cmdOK_Click

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "lstVendor_DblClick", True
50        Resume Exit_Handler
End Sub
