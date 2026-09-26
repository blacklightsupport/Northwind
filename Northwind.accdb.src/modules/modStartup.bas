Attribute VB_Name = "modStartup"
Option Compare Database
Option Explicit

'NOTE:
'   The VBA code uses line numbers, so the error handler can report on which line the error occurred.
'   If you want to remove them, search online for "Remove line numbers in VBA" and you can find the code to do so.

Public Const g_strAppName   As String = "Northwind Traders Developer Edition"
Public Const g_dtNorthwindInception As Date = #11/1/2022#

Public Sub CloseAllForms()
10        On Error GoTo Err_Handler

          Dim frm                 As AccessObject

20        For Each frm In CurrentProject.AllForms
30            If frm.IsLoaded Then DoCmd.Close acForm, frm.Name, acSaveNo
40        Next frm

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError "modStartup", "CloseAllForms"
70        Resume Exit_Handler
End Sub

Public Sub CloseAllReports()
10        On Error GoTo Err_Handler

          Dim rpt                 As AccessObject

20        For Each rpt In CurrentProject.AllReports
30            If rpt.IsLoaded Then DoCmd.Close acReport, rpt.Name, acSaveNo
40        Next rpt

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError "modStartup", "CloseAllReports"
70        Resume Exit_Handler
End Sub

'PURPOSE:
'   The opposite of Startup. We are on our way out and want to do some cleanup before we exit for real.
Public Sub Finish()
10        On Error Resume Next    'We are on the way out and don't want to be bothered with error messages.

20        CloseAllOrderDetailsForms
30        CloseAllPurchaseOrderDetailsForms

40        CloseAllReports
50        CloseAllForms

60        RibbonFinish

70        Application.Quit acQuitSaveNone
End Sub

Public Function GetSystemSetting(ByVal SystemSettingID As enumSystemSettings) As Variant
10        On Error GoTo Err_Handler

          Dim v                   As Variant

20        v = DLookup("SettingValue", "SystemSettings", "SettingID = " & SystemSettingID)

          'Conversion to support regional settings.
30        Select Case SystemSettingID
              Case enumSystemSettings.ssTaxRate, enumSystemSettings.ssTaxRate_Vendors
40                v = CSng(v / 1000)    'Divide by 1000 to get Single value. See comment in table SystemSettings.
50            Case enumSystemSettings.ssLastResetDate
60                v = CDate(v)
70            Case enumSystemSettings.ssFirstTimeRun, enumSystemSettings.ssShowWelcome
80                v = CBool(v)
90            Case Else
100               Debug.Assert False      'Unexpected SystemSettingsID passed in.
110       End Select

Exit_Handler:
120       GetSystemSetting = v
130       Exit Function

Err_Handler:
140       clsErrorHandler.HandleError "modStartup", "GetSystemSetting"
150       Resume Exit_Handler
160       Resume
End Function

Public Function GetUserSetting(ByVal UserSettingID As enumUserSettings) As Variant
10        On Error GoTo Err_Handler

20        GetUserSetting = DLookup("SettingValue", "UserSettings", "SettingID = " & UserSettingID)

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modStartup", "GetUserSetting"
50        Resume Exit_Handler
End Function

Public Sub SaveSystemSetting(ByVal SystemSettingID As enumSystemSettings, ByVal varValue As Variant)
10        On Error GoTo Err_Handler

          Dim sql                 As String

          'Conversion to support regional settings.
20        Select Case SystemSettingID
              Case enumSystemSettings.ssLastResetDate
30                varValue = ToAccessDate(varValue)
40        End Select

50        sql = StringFormatSQL("Update SystemSettings set SettingValue = {0} where SettingID = {1};", varValue, SystemSettingID)
60        g_dbApp().Execute sql, dbFailOnError

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError "modStartup", "SaveSystemSetting"
90        Resume Exit_Handler
End Sub

Public Sub SaveUserSetting(ByVal UserSettingID As enumUserSettings, ByVal varValue As Variant)
10        On Error GoTo Err_Handler

          Dim sql                 As String

20        sql = StringFormatSQL("Update UserSettings set SettingValue = {0} where SettingID = {1};", varValue, UserSettingID)
30        g_dbApp().Execute sql, dbFailOnError

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError "modStartup", "SaveUserSetting"
60        Resume Exit_Handler
End Sub

Public Sub SetAppTitle(ByVal bolIncludeUserName As Boolean)
10        On Error GoTo Err_Handler

20        If bolIncludeUserName = False Then
30            CurrentDb.Properties("AppTitle") = g_strAppName & Space(2) & APP_VERSION
40        Else
50            CurrentDb.Properties("AppTitle") = g_strAppName & Space(2) & APP_VERSION & Space(5) & Get_EmployeeFNLN(Get_UserID())
60        End If

70        RefreshTitleBar

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError "modStartup", "SetAppTitle"
100       Resume Exit_Handler
End Sub

'PURPOSE
'   Called from AutoExec macro to start the application.
Public Function Startup()
10        On Error GoTo Err_Handler

          Dim varShowWelcome      As Variant


20        Application.SetOption "Error Trapping", 2   'AddDataMacros calls HasField, which will fail if Error Trapping is set to "Break on all errors.". The setting is in VBA window > Tools > Options > General.

          'Set the Application main window Title
30        modStartup.SetAppTitle False

40        OneTimeProcessing

          'Show the Welcome form
50        varShowWelcome = GetSystemSetting(ssShowWelcome)
60        If varShowWelcome = True Then
70            DoCmd.OpenForm "frmWelcome", acNormal, , , , acDialog
80        End If

90        InitializeUser

100       SetAppTitle True
          'Too soon. ribbonLoaded is a better location anyway.  ActivateTab "tHome"

110       DoCmd.OpenForm "frmOrderList"       'Open the main form, rather than having the user do it.

Exit_Handler:
120       Exit Function

Err_Handler:
130       clsErrorHandler.HandleError "modStartup", "Startup", True
140       Resume Exit_Handler
150       Resume
End Function
