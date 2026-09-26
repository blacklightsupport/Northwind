VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'NOTE:
'   This is a "default instance" class, meaning the developer does not need to instantiate it, it is always there, and the instance always has the name of the class.
'   This is somewhat similar to "static" class in other programming languages, but VBA does not have true static classes.
'   To use this class in your own projects you have to use these steps: Export File out of this database, and Import File into your database.
'   A simple copy/paste of the code is NOT enough.

'REFERENCE:
'   https://learn.microsoft.com/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/189fb41b-cc3a-4999-a6d2-ba89f72d2870

Private Const m_VbProjectName As String = "NW2-Dev"       'The name of the VBA project; often used as the Err.Source value.
Private Const LOGFILENAME   As String = "Northwind-Dev.log"
Private m_strLogFilePath    As String

Public Sub ErrorHandlerExample()
10        On Error GoTo Err_Handler

          'Code goes here.

Exit_Handler:
20        Exit Sub

Err_Handler:
30        clsErrorHandler.HandleError "YourModuleName", "YourProcedureName"
40        Resume Exit_Handler
End Sub

Public Sub ErrorHandlerExample_EventProcedure()
10        On Error GoTo Err_Handler

          'Code goes here.

Exit_Handler:
20        Exit Sub

Err_Handler:
30        clsErrorHandler.HandleError "YourModuleName", "YourProcedureName", True
40        Resume Exit_Handler
End Sub

Private Sub Class_Initialize()
          'NOTE: Do not add standard error handler code to this constructor. It messes with the first error.

          'Create full path to error log file in same folder as current database, by concatenating the folder, a backslash, and a filename.
10        m_strLogFilePath = CurrentProject.Path & "\" & LOGFILENAME

End Sub

'Espana UG Identify the source of the error to ensure proper feedback al usario

'ARGUMENTS:
'   strModuleName    - Name of the VBA module where the error occurred.
'   strProcedureName - Name of the VBA procedure where the error occurred.
'   IsEventProcedure - (new for v2.2) True for Event Procedures and other toplevel procedures such as Startup; False for all other procedures.
'   If you forget to set it, the Northwind 2.0 behavior will happen.
'NOTE:
'   This main error handler should not itself have an error handler.
Public Sub HandleError(ByVal strModuleName As String, ByVal strProcedureName As String, _
                       Optional ByVal IsEventProcedure As Boolean = False)

          Dim lngError            As Long
          Dim lngErrorLine        As Long
          Dim strError            As String
          Dim strMsg              As String
          Dim strErrLine          As String
          Dim strErrSource        As String

          'Get the current error information
10        strError = Err.Description
20        lngError = Err.Number
30        lngErrorLine = Erl                  'If no line number, this will be 0.

40        strErrSource = Err.Source
50        If strErrSource = m_VbProjectName Then
60            strErrSource = vbNullString
70        Else
80            strErrSource = " -> " & strErrSource
90        End If

100       If lngErrorLine > 0 Then
110           strErrLine = " (L " & lngErrorLine & ")"
120       End If

130       strErrSource = strModuleName & "." & strProcedureName & strErrLine & strErrSource

140       If lngError = 2501 Then Exit Sub    '2501 = The OpenForm action was canceled. Not really an error.

          'Log the error, even before we show it to the user.
150       LogErrorToFile strError, lngError, lngErrorLine, strModuleName, strProcedureName

160       If Not IsEventProcedure Then
170           Err.Raise lngError, strErrSource, strError
180       End If

          'Compose the error message by concatenating substrings.

190       strMsg = "The following error has occurred: " & vbCrLf
200       strMsg = strMsg & strError & vbCrLf
210       If lngErrorLine <> 0 Then strMsg = strMsg & "On line " & lngErrorLine & vbCrLf
220       strMsg = strMsg & strModuleName & "." & strProcedureName & vbCrLf & vbCrLf
230       strMsg = strMsg & "Error number: " & lngError & vbCrLf & vbCrLf
240       strMsg = strMsg & "Call stack: " & strErrSource & vbCrLf & vbCrLf
250       strMsg = strMsg & "Please notify the Help Desk.  Thank you!"

          'Show the message to the user. A developer running ACCDB can hit Ctrl+Break and Step in the debugger to get to the calling procedure.

260       MsgBox strMsg, vbCritical

End Sub

Private Sub LogErrorToFile(ByVal strError As String, ByVal lngError As Long, ByVal lngErrorLine As Long, ByVal strModuleName As String, ByVal strProcedureName As String)
10        On Error GoTo Err_Handler

          Dim intFile             As Integer

20        intFile = FreeFile
30        Open m_strLogFilePath For Append As intFile

40        Print #intFile, "Current Time   : " & Now
50        Print #intFile, "Module         : " & strModuleName
60        Print #intFile, "Procedure      : " & strProcedureName
70        Print #intFile, "Error String   : " & strError
80        Print #intFile, "Error Number   : " & lngError
90        Print #intFile, "Error Line     : " & IIf(lngErrorLine = 0, "<none>", lngErrorLine)
100       Print #intFile, "User           : " & Environ$("USERNAME")
110       Print #intFile, "Machine        : " & Environ$("COMPUTERNAME")
120       Print #intFile, "---------------"

130       Close #intFile

Exit_Handler:
140       Exit Sub

Err_Handler:
          'Do not call the error handler, or you may get recursive errors. Better to ignore it.  g_ErrorHandler.HandleError "clsErrorHandler", "LogErrorToFile"
150       Debug.Print "Unable to LogErrorToFile. Error number: " & Err.Number & ", Description: " & Err.Description
160       Resume Exit_Handler
End Sub
