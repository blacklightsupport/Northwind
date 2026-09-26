Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_UserID            As Long

'This property is used to communicate the selected user to the calling code. See also: Get_UserID().
Public Property Get UserID() As Long
10        UserID = m_UserID
End Property

Private Sub cmdAddMe_Click()
10        On Error GoTo Err_Handler

20        RunCommand acCmdSaveRecord

30        m_UserID = Me.EmployeeID.Value

          'You can log in to the application as any Employee
          'Let the user know who they have logged in as
40        modStartup.SetAppTitle True

50        Me.Visible = False    'Do not close the form. Allow caller to inspect properties.

Exit_Handler:
60        Exit Sub

Err_Handler:
70        If Err.Number = 3021 Then           ' 3021 = No current record. The form is not valid. Validator has already shown error message.
80            Resume Exit_Handler
90        ElseIf Err.Number = 3022 Then       ' 3022 = The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship. Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again.
100           Err.Description = "This FirstName + LastName combination already exists. Please enter a different name."
110       ElseIf Err.Number = 2046 Then       ' 2046 = The command or action 'SaveRecord' isn't available now.
120           Err.Description = "This record cannot be saved as-is. Does this FirstName + LastName combination already exist?"
130       End If
140       clsErrorHandler.HandleError Me.Name, "cmdAddMe_Click"
150       Resume Exit_Handler
160       Resume
End Sub

Private Sub cmdExit_Click()
10        On Error GoTo Err_Handler

20        If Me.Dirty Then
30            Me.Undo
40        End If

50        DoCmd.Close acForm, Me.Name

60        DoCmd.Quit acQuitSaveAll

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "cmdExit_Click", True
90        Resume Exit_Handler
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Cancel = ValidateForm(Me)

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Load()
10        On Error GoTo Err_Handler

          Dim strUserName         As String

20        strUserName = GetWindowsUserName()

          'this value will be used to find the user in the future
30        Me.WindowsUserName = strUserName

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_Load", True
60        Resume Exit_Handler
End Sub
