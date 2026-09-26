Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_bolLoginClicked   As Boolean
Private m_UserID            As Long

'This property is used to communicate the selected user to the calling code. See also: Get_UserID().
Public Property Get UserID() As Long
10        UserID = m_UserID
End Property

Private Sub AutoLogin_chk_AfterUpdate()
10        On Error GoTo Err_Handler

20        SaveUserSetting usAutoLogin, Me.AutoLogin_chk

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "AutoLogin_chk_AfterUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub cmdLogin_Click()
10        On Error GoTo Err_Handler

20        If IsNull(Me.cboEmployee.Value) Then
30            MsgBox "You must select a user to login as.  Please select someone now.", vbOKOnly Or vbExclamation, "Select a Login"
40            Me.cboEmployee.SetFocus
50            Me.cboEmployee.Dropdown
60            m_bolLoginClicked = False
70            Exit Sub
80        End If

          'This lets us know the user clicked on the Login button
          'If you look at the forms Unload event below you will see why we need to know this
90        m_bolLoginClicked = True

100       m_UserID = Me.cboEmployee.Value

          'You can log in to the application as any Employee
          'Let the user know who they have logged in as
110       modStartup.SetAppTitle True

120       Me.Visible = False    'Do not close the form. Allow caller to inspect properties.

Exit_Handler:
130       Exit Sub

Err_Handler:
140       clsErrorHandler.HandleError Me.Name, "cmdLogin_Click", True
150       Resume Exit_Handler
End Sub

Private Sub Form_Load()
10        On Error GoTo Err_Handler

          Dim bolAutoLogIn        As Boolean
          Dim lngUserID           As Long

20        bolAutoLogIn = False

30        Me.Disclaimer.Caption = GetString(sDisclaimer)

40        bolAutoLogIn = GetUserSetting(enumUserSettings.usAutoLogin)

50        Me.AutoLogin_chk.Value = bolAutoLogIn

60        lngUserID = Get_UserID()
70        Me.cboEmployee = lngUserID

80        If lngUserID = 0 Then
              ' The function has told us there is no User selected at this time
              ' our business rule says to focus on the default log in
90            m_UserID = DEFAULT_LOGIN_ID
100           Me.AutoLogin_chk.Visible = False
110       ElseIf lngUserID = Get_UserID_ForWindowsUser Then
120           Me.AutoLogin_chk.Visible = True
130       Else
140           Me.AutoLogin_chk.Visible = False
150       End If

160       m_bolLoginClicked = False

170       modStartup.SetAppTitle False

Exit_Handler:
180       Exit Sub

Err_Handler:
190       clsErrorHandler.HandleError Me.Name, "Form_Load", True
200       Resume Exit_Handler
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        On Error GoTo Err_Handler

20        If m_bolLoginClicked = False Then
30            If MsgBox("Do you want to close " & g_strAppName & "?", vbYesNo Or vbExclamation, "Select a Login") = vbNo Then
40                Me.cboEmployee.SetFocus
50                Cancel = True
60            Else
70                Me.Application.Quit acQuitSaveNone
80            End If
90        End If

Exit_Handler:
100       Exit Sub

Err_Handler:
110       clsErrorHandler.HandleError Me.Name, "Form_Unload", True
120       Resume Exit_Handler
End Sub
