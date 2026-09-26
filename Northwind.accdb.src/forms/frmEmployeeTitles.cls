Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

          Dim dict                As Scripting.Dictionary

20        Set dict = StringToDictionary(Me.OpenArgs)
30        If Not Me.NewRecord Then RunCommand acCmdRecordsGoToNew
40        Me.Title = dict("[frmEmployeeList]![cboTitle]")
50        Set dict = Nothing

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "Form_Open", True
80        Resume Exit_Handler
End Sub
