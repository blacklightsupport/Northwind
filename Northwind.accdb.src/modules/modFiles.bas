Attribute VB_Name = "modFiles"
Option Compare Database
Option Explicit

Function FileExists(ByVal strFile As String, Optional ByVal blnFindFolders As Boolean) As Boolean
          'Purpose:   Return True if the file exists, even if it is hidden.
          'Arguments: strFile: File name to look for. Current directory searched if no path included.
          '           blnFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
          'Note:      Does not look inside subdirectories for the file.
          'Author:    Allen Browne. http://allenbrowne.com June, 2006.  http://allenbrowne.com/func-11.html
          Dim lngAttributes       As Long

          'Include read-only files, hidden files, system files.
10        lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

20        If blnFindFolders Then
30            lngAttributes = (lngAttributes Or vbDirectory)    'Include folders as well.
40        Else
              'Strip any trailing slash, so Dir does not look inside the folder.
50            Do While Right$(strFile, 1) = "\"
60                strFile = Left$(strFile, Len(strFile) - 1)
70            Loop
80        End If

          'If Dir() returns something, the file exists.
90        On Error Resume Next
100       FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Public Sub StringToFile(ByVal strPath As String, ByVal theString As Variant)
10        On Error GoTo Err_Handler

          Dim intFile             As Integer

20        intFile = FreeFile

30        Open strPath For Output As #intFile
40        Print #intFile, theString
50        Close #intFile

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError "modFiles", "StringToFile"
80        Resume Exit_Handler
End Sub
