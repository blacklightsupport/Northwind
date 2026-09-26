Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub txtHelpKeywords_Click()
10        On Error GoTo Err_Handler

          Dim strHyperlink        As String

          'Exit if the field is blank to avoid error
20        If Nz(Me.txtHelpKeywords, "") = "" Then GoTo Exit_Handler

30        Select Case Me.OpenMethod
              Case enumNorthwindFeaturesOpenMethod.nfInPageHelp
40                Application.Assistance.SearchHelp Me.txtHelpKeywords

50            Case enumNorthwindFeaturesOpenMethod.nfFollowHyperlink
                  'The field in table is defined as a hyperlink
                  '  and we want it to be a hyperlink so if
                  '   the field is shown and the user clicks on it the webpage will open
                  '  See: txtLearnMore_Click()
                  '
                  'The field type of hyperlink coantains pound signs (#) at the beginning and the end
                  '  we need to remove those to access the webpage from here

60                strHyperlink = Replace(CStr(Me.txtLearnMore), "#", "")

70                Application.FollowHyperlink strHyperlink
80        End Select

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError Me.Name, "txtHelpKeywords_Click", True
110       Resume Exit_Handler
End Sub
