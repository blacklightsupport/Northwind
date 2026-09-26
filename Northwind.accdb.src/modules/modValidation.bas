Attribute VB_Name = "modValidation"
Option Compare Database
Option Explicit

Private Const BACKGROUND_COLOR As Long = vbWhite        'The "Background 1" color is vbWhite
Private Const HIGHLIGHT_COLOR As Long = vbYellow        'Adjust to your taste

Public Sub HighlightControl(ByRef ctl As Access.control)
10        On Error GoTo Err_Handler

          Dim fc                  As FormatCondition
          Dim intCurrentView      As Integer
          Dim intDefaultView      As Integer

20        intCurrentView = ControlCurrentView(ctl)
30        intDefaultView = ControlDefaultView(ctl)
          'If current=form and default=continuousforms, or current=datasheet
40        If (intCurrentView = 1 And intDefaultView = 1) Or intCurrentView = 2 Then
              'Datasheet does not have BackColor. The way to color a column is through Conditional Formatting.

50            If ctl.FormatConditions.count >= 4 Then     'Just in case we have a runaway process.
60                ctl.FormatConditions(3).Delete
70            End If
80            Set fc = ctl.FormatConditions.Add(AcFormatConditionType.acExpression, , "True")     'In a more elaborate implementation you can limit highlighting to the current row, by making the expression something like "OrderDetailID=123". This requires finding the PK in the underlying recordsource.
90            fc.BackColor = HIGHLIGHT_COLOR
100       Else
              'Transparent controls cannot have a BackColor, so first set their BackStyle to be "Normal".
              'NOTE: We are not setting this property back to what it was.
110           If HasProperty(ctl, "BackStyle") Then
120               ctl.BackStyle = 1                   '1=Normal, 0=Transparent
130               ctl.BackColor = HIGHLIGHT_COLOR     'If a control has a BackStyle, it also has a BackColor. No need for another HasProperty test.
140           End If

150       End If

Exit_Handler:
160       Exit Sub

Err_Handler:
170       clsErrorHandler.HandleError "modValidation", "HighlightControl"
180       Resume Exit_Handler
End Sub

'PURPOSE:
'   Highlight required fields that do not have a value.
'ALGORITHM:
'   Loop over the controls on the form. Test if all required fields have a value.
'   If not, highlight it.
Private Sub HighlightInvalidControls(ByRef frm As Access.Form)
10        On Error GoTo Err_Handler

          Dim ctl                 As Access.control

20        For Each ctl In frm.Controls
30            If IsBoundToRequiredField(ctl) Then
40                If IsNull(ctl.Value) Then
                      'Control bound to required field does not have a value. Highlight it.
50                    HighlightControl ctl
60                End If
70            End If
80        Next ctl

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError "modValidation", "HighlightInvalidControls"
110       Resume Exit_Handler
End Sub

'ALGORITHM:
'   Access the control's Parent which is its form. Get the RecordsetClone for the form.
'   Access the field we are bound to, and check its Required property.
Private Function IsBoundToRequiredField(ByRef ctl As Access.control) As Boolean
10        On Error Resume Next    'Not all controls have a ControlsSource property.
20        IsBoundToRequiredField = ctl.Parent.RecordsetClone.Fields(ctl.ControlSource).Required
End Function

'PURPOSE:
'   Validate the form with respect to required fields.
'ALGORITHM:
'   Loop over the form's controls. Test if all required fields have a value.
'RETURNS:
'   True if the form is valid (all required fields have been entered); False otherwise.
Private Function IsValidForm(ByRef frm As Access.Form) As Boolean
10        On Error GoTo Err_Handler

          Dim blnIsValid          As Boolean
          Dim ctl                 As Access.control

20        blnIsValid = True   'Optimistic

30        For Each ctl In frm.Controls
40            If IsBoundToRequiredField(ctl) Then
50                If IsNull(ctl.Value) Then
                      'This required field is not filled out.
60                    blnIsValid = False
70                    Exit For
80                End If
90            End If
100       Next ctl

110       IsValidForm = blnIsValid

Exit_Handler:
120       Exit Function

Err_Handler:
130       clsErrorHandler.HandleError "modValidation", "IsValidForm"
140       Resume Exit_Handler
End Function

'PURPOSE:
'   Validate the form with respect to required fields.
'   Typically you call this from Form_BeforeUpdate which is just before the data is (attempted to be) saved.
'RETURNS:
'   The "Cancel" value used in BeforeUpdate events (True if invalid).
Public Function ValidateForm(ByRef frm As Access.Form) As Boolean
10        On Error GoTo Err_Handler

          Dim blnCancel           As Boolean

20        blnCancel = Not IsValidForm(frm)
30        If blnCancel Then
40            HighlightInvalidControls frm
50            MsgBox GetString(sRequiredFields), vbExclamation
60        End If
70        ValidateForm = blnCancel

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError "modValidation", "ValidateForm"
100       Resume Exit_Handler
End Function

'PURPOSE:
'   Turn off the highlight color by setting the background to the default white color.
'   Typically you call this procedure from Form_AfterUpdate, which fires after a record has been saved
'   so obviously all required fields have been satisfied.
Public Sub ValidateForm_RemoveHighlights(ByRef frm As Access.Form)
          Dim ctl                 As Access.control
          Dim fc                  As FormatCondition
          Dim intCurrentView      As Integer
          Dim intDefaultView      As Integer

10        Set ctl = frm.Controls(0)   'Any control will do; we will be inspecting parent controls.
20        intCurrentView = ControlCurrentView(ctl)
30        intDefaultView = ControlDefaultView(ctl)
          'If current=form and default=continuousforms, or current=datasheet
40        If (intCurrentView = 1 And intDefaultView = 1) Or intCurrentView = 2 Then

              '    If frm.CurrentView = 2 Then     '2=datasheet
50            On Error Resume Next        'Not all controls have FormatConditions
60            For Each ctl In frm.Controls

70                For Each fc In ctl.FormatConditions
80                    If fc.Expression1 = "True" And fc.BackColor = HIGHLIGHT_COLOR Then
90                        fc.Delete
100                       Exit For
110                   End If
120               Next fc
130           Next ctl
140       Else
150           On Error Resume Next            'Not all controls have a BackColor property
160           For Each ctl In frm.Controls
170               If ctl.BackColor = HIGHLIGHT_COLOR Then ctl.BackColor = BACKGROUND_COLOR
180           Next ctl
190       End If

End Sub
