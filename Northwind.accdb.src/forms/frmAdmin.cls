Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'This module was created by converting Macro Code (see design ribbon far right)
'The comments inside the module were created by the conversion process
'We added the error handling
'------------------------------------------------------------
' cmdDeleteTestData_Click
'
'------------------------------------------------------------
Private Sub cmdDeleteTestData_Click()
10        On Error GoTo Err_Handler

          ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
          ' <UserInterfaceMacro For="cmdPrivileges" Event="OnClick" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="SetValue"><Argument Name="Item
          ' _AXL:">[sfrmAdmin].[SourceObject]</Argument><Argument Name="Expression">"sfrmEmployee_Privileges"</Argument></Action></Statements></UserInterfaceMacro>
20        sfrmAdmin.SourceObject = "sfrmAdmin_DeleteTestData"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdDeleteTestData_Click", True
50        Resume Exit_Handler
End Sub

'This module was created by converting Macro Code (see design ribbon far right)
'The comments inside the module were created by the conversion process
'We added the error handling
'------------------------------------------------------------
' cmdInternetOrders_Click
'
'------------------------------------------------------------
Private Sub cmdInternetOrders_Click()
10        On Error GoTo Err_Handler

          ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
          ' <UserInterfaceMacro For="cmdSystemSettings" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="SetValue"><Argument Name="Item">[sfrmAdmin
          ' _AXL:].[SourceObject]</Argument><Argument Name="Expression">"sfrmAdmin_SystemSettings"</Argument></Action></Statements></UserInterfaceMacro>
20        sfrmAdmin.SourceObject = "sfrmAdmin_InternetOrders"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdInternetOrders_Click", True
50        Resume Exit_Handler
End Sub

'This module was created by converting Macro Code (see design ribbon far right)
'The comments inside the module were created by the conversion process
'We added the error handling
'------------------------------------------------------------
' cmdPrivileges_Click
'
'------------------------------------------------------------
Private Sub cmdPrivileges_Click()
10        On Error GoTo Err_Handler

          ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
          ' <UserInterfaceMacro For="cmdStrings" Event="OnClick" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="SetValue"><Argument Name="Item">[
          ' _AXL:sfrmAdmin].[SourceObject]</Argument><Argument Name="Expression">"sfrmAdmin_Strings"</Argument></Action></Statements></UserInterfaceMacro>
20        sfrmAdmin.SourceObject = "sfrmEmployee_Privileges"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdPrivileges_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdProductCategories_Click()
10        On Error GoTo Err_Handler

20        Me.sfrmAdmin.SourceObject = "sfrmProductCategories"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdProductCategories_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdRelationshipsDiagram_Click()
10        On Error GoTo Err_Handler

          Dim strMsg              As String

20        strMsg = "At the time of release of this template, the ACCDT process which creates the Access Template, does not preserve the layout of the Relationships window. " & _
                   "We included this report so you can see it laid out as was intended."
30        MsgBox strMsg, vbInformation
40        DoCmd.OpenReport "rptRelationshipsWindow", acViewReport

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "cmdRelationshipsDiagram_Click", True
70        Resume Exit_Handler
End Sub

'This module was created by converting Macro Code (see design ribbon far right)
'The comments inside the module were created by the conversion process
'We added the error handling
'------------------------------------------------------------
' cmdResetDates_Click
'
'------------------------------------------------------------
Private Sub cmdResetDates_Click()
10        On Error GoTo Err_Handler

          ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
          ' <UserInterfaceMacro For="cmdInternetOrders" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="SetValue"><Argument Name="Item">[sfrmAdmin
          ' _AXL:].[SourceObject]</Argument><Argument Name="Expression">"sfrmAdmin_InternetOrders"</Argument></Action></Statements></UserInterfaceMacro>
20        sfrmAdmin.SourceObject = "sfrmAdmin_ResetDates"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdResetDates_Click", True
50        Resume Exit_Handler
End Sub

'This module was created by converting Macro Code (see design ribbon far right)
'The comments inside the module were created by the conversion process
'We added the error handling
'------------------------------------------------------------
' cmdStrings_Click
'
'------------------------------------------------------------
Private Sub cmdStrings_Click()
10        On Error GoTo Err_Handler

          ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
          ' <UserInterfaceMacro For="cmdResetDates" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><Statements><Action Name="SetValue"><Argument Name="Item">[sfrmAdmin].[S
          ' _AXL:ourceObject]</Argument><Argument Name="Expression">"sfrmAdmin_ResetDates"</Argument></Action></Statements></UserInterfaceMacro>
20        sfrmAdmin.SourceObject = "sfrmAdmin_Strings"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdStrings_Click", True
50        Resume Exit_Handler
End Sub

'This module was created by converting Macro Code (see design ribbon far right)
'The comments inside the module were created by the conversion process
'We added the error handling
'------------------------------------------------------------
' cmdSystemSettings_Click
'
'------------------------------------------------------------
Private Sub cmdSystemSettings_Click()
10        On Error GoTo Err_Handler

20        sfrmAdmin.SourceObject = "sfrmAdmin_SystemSettings"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdSystemSettings_Click", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

20        UpdateSubLabels

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Open", True
50        Resume Exit_Handler
End Sub

'NOTE:
'   Public so the various subforms can update these labels as well.
Public Sub UpdateSubLabels()
10        On Error GoTo Err_Handler

20        Me.lblSystemSettingsSub.Caption = "Count: " & DCount("*", "SystemSettings")
30        Me.lblInternetOrdersSub.Caption = "Last 30 days: " & DCount("*", "Orders", StringFormatSQL("EmployeeID = {0} and OrderDate > {1}", INTERNET_SALES_EMPLOYEEID, DateAdd("d", -30, Date)))
40        Me.lblResetDatesSub.Caption = "Last Reset: " & GetSystemSetting(ssLastResetDate)
50        Me.lblStringsSub.Caption = "Count: " & DCount("*", "Strings")
60        Me.lblPrivilegesSub.Caption = "Count: " & DCount("*", "Privileges")

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "UpdateSubLabels"
90        Resume Exit_Handler
End Sub
