Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterDelConfirm(Status As Integer)
10        On Error GoTo Err_Handler

20        If frmLoadedAsSubform Then Me.Parent.UpdateSubLabels

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterDelConfirm", True
50        Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

20        If frmLoadedAsSubform Then Me.Parent.UpdateSubLabels

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

20        With Me
30            If Nz(.OpenArgs, 0) > 0 Then
40                DoCmd.MoveSize Right:=TWIPS_PER_INCH * 3, down:=TWIPS_PER_INCH * 3, Width:=TWIPS_PER_INCH * 4.5, Height:=TWIPS_PER_INCH * 4.5
50                .cboEmployeeID.DefaultValue = .OpenArgs
60            End If
70        End With

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "Form_Open", True
100       Resume Exit_Handler
End Sub

Private Function frmLoadedAsSubform() As Boolean
10        On Error GoTo Err_Handler

20        frmLoadedAsSubform = True
30        Debug.Print Me.Parent.Name

Exit_Handler:
40        Exit Function

Err_Handler:
50        If Err = 2452 Then      '2452 = The expression you entered has an invalid reference to the Parent property.
              '@For example, you may be using the Parent property with a control on a main form or report
              'rather than with a control on a subform or subreport.@@2@1@4852@1
60            frmLoadedAsSubform = False
70        Else
80            clsErrorHandler.HandleError Me.Name, "frmLoadedAsSubform"
90        End If
100       Resume Exit_Handler
End Function
