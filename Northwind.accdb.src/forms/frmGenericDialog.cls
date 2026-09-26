Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancelClose_Click()
10        On Error GoTo Err_Handler

          'If we close the form the calling form will know the User clicked Cancel
20        DoCmd.Close acForm, Me.Name, acSaveNo

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdCancelClose_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdOK_Click()
10        On Error GoTo Err_Handler

          'If we make the form invisible the calling form know the User clicked OK.
20        Me.Visible = False

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdOK_Click", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Load()
10        On Error GoTo Err_Handler

          Dim dict                As Scripting.Dictionary
          Dim key                 As Variant
          Dim strActions          As String
          Dim strFormCaption      As String
          Dim strHeader           As String
          Dim strMsg              As String

          'Set default values for the variables above
20        strActions = "None"
30        strFormCaption = "Message from Northwind"
40        strHeader = "No Message Provided"
50        strMsg = "Please let the IT team know what happened."

          '          ==============================================================================
          '          Passing Multiple values in via the OpenArgs property
          '          ==============================================================================
          '           One approach is to pass the values in a string using a delimiter (for example | ).
          '           The opened form uses the Split() function to break the values apart (see frmSelectVendorDialog).
          '           This approach requires passing the values in the order expected by the opened form.
          '           If there 3 optional values you will need pass those in as well even if you have only the last value ( ||| )
          '
          '           Another approach is to use Key/Value pairs and the Scripting.Dictionary object to parse them.
          '           A benefit of this approach is the values are identified by the KEY portion of the pair not the
          '           Ordinal position - so the order of the values is no longer important.
          '           The NW function StringToDictionary handles the parsing of the Key/Value pairs for you.

60        If Nz(Me.OpenArgs, "") > "" Then

70            Set dict = StringToDictionary(Me.OpenArgs)

80            For Each key In dict.Keys
90                Select Case key
                      Case "FormCaption"
100                       strFormCaption = dict.Item(key)
110                   Case "Header"
120                       strHeader = dict.Item(key)
130                   Case "Message"
140                       strMsg = dict.Item(key)
150                   Case "Actions"
160                       strActions = dict.Item(key)
170               End Select
180           Next key
190       End If

200       With Me
210           .Caption = strFormCaption
220           .Auto_Header0.Caption = strHeader
230           .txtMessage = Trim$(strMsg)

240           If strActions = "None" Then
250               .cmdCancelClose.Visible = True
260               .cmdCancelClose.Caption = "&Close"
270               .cmdOK.Visible = False
280           ElseIf strActions = "vbOKCancel" Then
290               .cmdCancelClose.Visible = True
300               .cmdCancelClose.Caption = "&Cancel"
310               .cmdOK.Visible = True
320               .cmdOK.Caption = "&OK"
330           ElseIf strActions = "vbYesNo" Then
340               .cmdCancelClose.Visible = True
350               .cmdCancelClose.Caption = "&No"
360               .cmdOK.Visible = True
370               .cmdOK.Caption = "&Yes"
380           Else
                  'A value was sent that we don't recognize
390               .cmdCancelClose.Visible = True
400               .cmdCancelClose.Caption = "&Close"
410               .cmdOK.Visible = False
420           End If
430       End With

440       With Me.txtMessage
              'Cause the vertical scroll bar to appear in case
              'the message exceeds the current size of the control
450           .SetFocus
460           .SelStart = 0
470       End With

Exit_Handler:
480       Set dict = Nothing
490       Exit Sub
Err_Handler:
500       clsErrorHandler.HandleError Me.Name, "Form_Load"
510       Resume Exit_Handler
End Sub
