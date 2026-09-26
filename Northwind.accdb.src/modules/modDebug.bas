Attribute VB_Name = "modDebug"
Option Compare Database
Option Explicit

Public Sub DumpTempVars()
10        On Error GoTo Err_Handler

          Dim t                   As TempVar

20        Debug.Print "Dumping " & TempVars.count & " TempVars:"
30        For Each t In TempVars
40            Debug.Print t.Name, t.Value
50        Next t

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError "modDebug", "DumpTempVars"
80        Resume Exit_Handler
End Sub

Public Sub NotImplemented(Optional ByVal strMsg As String)
10        On Error GoTo Err_Handler

20        strMsg = "This feature is not yet implemented." & vbCrLf & strMsg
30        MsgBox strMsg, vbExclamation

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError "modDebug", "NotImplemented"
60        Resume Exit_Handler
End Sub
