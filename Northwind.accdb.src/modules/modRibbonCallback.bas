Attribute VB_Name = "modRibbonCallback"
Option Compare Database
Option Explicit

'This module contains functions that are called by the Ribbon.

Private m_Ribbon            As Object       'Office.IRibbonUI    'Requires reference to Microsoft Office 16.0 Object Library  (mso.dll).
Private m_rsMRU             As DAO.Recordset
'

Public Sub ActivateTab(ByVal strControlID As String)
10        On Error GoTo Err_Handler

20        If m_Ribbon Is Nothing Then
              'User may have Shift-loaded the app and the startup code did not run.
30        Else
40            m_Ribbon.ActivateTab strControlID
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError "modRibbonCallback", "ActivateTab"
80        Resume Exit_Handler
End Sub

'ARGUMENTS:
'   control - Required parameter for OnAction callback function. Dead Code Review may find this item, but it needs to stay this way.
Public Sub cmdAbout_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmAbout", , , , , acDialog

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdAbout_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdAddOrder_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        OpenOrderDetailsForm

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdAddOrder_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdAddPurchaseOrder_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        OpenPurchaseOrderDetailsForm

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdAddPurchaseOrder_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdAdmin_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmAdmin"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdAdmin_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdCustomers_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmCompanyList"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdCustomers_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdEmployees_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmEmployeeList"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdEmployees_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdExitApplication_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        Finish

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdExitApplication_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdExportToExcel_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        RunCommand acCmdExportExcel

Exit_Handler:
30        Exit Sub

Err_Handler:
40        If Err.Number = 2046 Then       '2046: The command or action 'ExportExcel' isn't available now.
50            MsgBox "This feature requires an exportable object (such as a form) to be selected.", vbInformation
60        Else
70            clsErrorHandler.HandleError "modRibbonCallback", "cmdExportToExcel_OnAction"
80        End If
90        Resume Exit_Handler
End Sub

Public Sub cmdFeatures_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmNorthwindFeatures"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdFeatures_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdLearn_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmLearn"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdLearn_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdNorthwindDocumentation_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        Application.FollowHyperlink "https://support.microsoft.com/topic/32eb79d2-bede-4ea4-b575-0714ca8dc1e2"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdNorthwindDocumentation_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdOrders_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmOrderList"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdOrders_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdProducts_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmProductList"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdProducts_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdPurchaseOrders_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmPurchaseOrderList"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdPurchaseOrders_OnAction"
50        Resume Exit_Handler
End Sub

Public Sub cmdReports_OnAction(ByVal control As Object)
10        On Error GoTo Err_Handler

20        DoCmd.OpenForm "frmReports"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "cmdReports_OnAction"
50        Resume Exit_Handler
End Sub

'PURPOSE:
'   Callback function assisting with populating list of users.
Public Sub ddMRU_GetItemCount(ByVal control As Object, ByRef count As Variant)
10        On Error GoTo Err_Handler

20        If m_rsMRU Is Nothing Then
30            Set m_rsMRU = CurrentDb.OpenRecordset("qryMRU", dbOpenDynaset)      'Dynaset because we may be deleting rows that no longer apply.
40        End If

50        With m_rsMRU
60            .Requery
70            If .RecordCount > 0 Then
80                .MoveLast
90                .MoveFirst
100           End If
110           count = Nz(m_rsMRU.RecordCount, 0)
120       End With

          'Debug.Print Time$, "ddMRU_GetItemCount", control.ID, count

Exit_Handler:
130       Exit Sub

Err_Handler:
140       clsErrorHandler.HandleError "modRibbonCallback", "ddMRU_GetItemCount"
150       Resume Exit_Handler
End Sub

'PURPOSE:
'   Callback function to provide the hidden ID value for this Index position.
Sub ddMRU_GetItemID(ByVal control As Object, ByVal Index As Long, ByRef ID As Variant)
10        On Error GoTo Err_Handler

20        If m_rsMRU.RecordCount > 0 Then
30            m_rsMRU.MoveFirst
40            m_rsMRU.Move Index
50            ID = m_rsMRU.Fields("MRU_ID")
60        End If

          'Debug.Print Time$, "ddMRU_GetItemID", control.ID, Index, ID

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError "modRibbonCallback", "ddMRU_GetItemID"
90        Resume Exit_Handler
End Sub

'PURPOSE:
'   Callback function to provide the label to be shown at this Index position.
Public Sub ddMRU_GetItemLabel(ByVal control As Object, ByVal Index As Long, ByRef Label As Variant)
10        On Error GoTo Err_Handler

          Dim varCompanyID        As Variant

20        If m_rsMRU.RecordCount > 0 Then
30            m_rsMRU.MoveFirst
40            m_rsMRU.Move Index
50            Select Case m_rsMRU!TableName
                  Case "Orders"
60                    varCompanyID = DLookup("CustomerID", "Orders", "OrderID = " & m_rsMRU!PKValue)
70                    If IsNull(varCompanyID) Then
80                        m_rsMRU.Delete     'MRU record points to a no longer existing order.
90                    Else
100                       Label = StringFormat("Order {0}, {1}", m_rsMRU!PKValue, DLookup("CompanyName", "Companies", "CompanyID = " & varCompanyID))
110                   End If

120               Case "PurchaseOrders"
130                   varCompanyID = DLookup("VendorID", "PurchaseOrders", "PurchaseOrderID = " & m_rsMRU!PKValue)
140                   If IsNull(varCompanyID) Then
150                       m_rsMRU.Delete     'MRU record points to a no longer existing purchase order.
160                   Else
170                       Label = StringFormat("PO {0}, {1}", m_rsMRU!PKValue, DLookup("CompanyName", "Companies", "CompanyID = " & varCompanyID))
180                   End If

190               Case Else
200                   Debug.Assert False      'Support for this tablename not yet implemented.
210           End Select
220       End If

          'Debug.Print Time$, "ddMRU_GetItemLabel", control.ID, Index, Label

Exit_Handler:
230       Exit Sub

Err_Handler:
240       clsErrorHandler.HandleError "modRibbonCallback", "ddMRU_GetItemLabel"
250       Resume Exit_Handler
End Sub

'PURPOSE:
'   Supply default value for the MRU list.
Public Sub ddMRU_GetSelectedItemIndex(ByVal control As Object, ByRef Index As Variant)
10        Index = 0       'Select first item in the list.

          'If the active form is an OrderDetail or PurchaseOrderDetail form, select that item in the MRU list.
          'This makes it easier to select from the MRU list, because user is likely to select the non-current item, and the desired Change event happens.
20        On Error Resume Next   'Just in case there is not be an active form.
30        If Screen.ActiveForm.Name = "frmOrderDetails" Then
40            m_rsMRU.FindFirst "PKValue = " & Screen.ActiveForm.OrderID
50            Index = m_rsMRU.AbsolutePosition
60        End If
70        If Screen.ActiveForm.Name = "frmPurchaseOrderDetails" Then
80            m_rsMRU.FindFirst "PKValue = " & Screen.ActiveForm.PurchaseOrderID
90            Index = m_rsMRU.AbsolutePosition
100       End If

          'Debug.Print Time$, "ddMRU_GetSelectedItemIndex", control.ID, Index
End Sub

'PURPOSE:
'   Event procedure that runs when a MRU list item is selected.
Public Sub ddMRU_OnAction(ByVal control As Object, ByVal selectedId As String, ByVal selectedIndex As Integer)
10        On Error GoTo Err_Handler

20        m_rsMRU.FindFirst "MRU_ID = " & selectedId
30        Debug.Assert Not m_rsMRU.NoMatch            'If this asserts, a record that was added to the dropdown is no longer there.

40        Select Case m_rsMRU!TableName
              Case "Orders"
50                OpenOrderDetailsForm m_rsMRU!PKValue

60            Case "PurchaseOrders"
70                OpenPurchaseOrderDetailsForm m_rsMRU!PKValue

80            Case Else
90                Debug.Assert False      'Support for this tablename not yet implemented.
100       End Select

Exit_Handler:
110       Exit Sub

Err_Handler:
120       clsErrorHandler.HandleError "modRibbonCallback", "ddMRU_OnAction"
130       Resume Exit_Handler
End Sub

Public Sub ddMRU_OnChange(ByVal control As Object, strText As String)
10        On Error GoTo Err_Handler

20        Debug.Print Time$, "ddMRU_OnChange", strText

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "ddMRU_OnChange"
50        Resume Exit_Handler
End Sub

Sub gReportOptions_GetVisible(ByVal control As Object, ByRef Visible As Variant)
10        On Error GoTo Err_Handler

20        Visible = (Reports.count > 0)

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "modRibbonCallback", "gReportOptions_GetVisible"
50        Resume Exit_Handler
End Sub

'NOTE:
'   This is not a callback function, but still ribbon-related.
Public Sub RibbonFinish()
10        On Error Resume Next            'Just in case the recordset is not open.
20        m_rsMRU.Close
30        Set m_rsMRU = Nothing
End Sub

'PROCEDURE:
'   ribbonLoaded
'PURPOSE:
'   Cache the ribbonUI object for later use, for example when we want to invalidate it.
'   Called by the onLoad method of the customUI node of the ribbon XML.
'ARGUMENTS:
'   ribbonUI    - Ribbon object.
Public Sub ribbonLoaded(ByVal ribbonUI As Object)       'As Office.IRibbonUI)
10        On Error GoTo Err_Handler

20        Set m_Ribbon = ribbonUI
30        ActivateTab "tHome"

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError "modRibbonCallback", "ribbonLoaded"
60        Resume Exit_Handler
End Sub

Public Sub Ribbon_RefreshMRU()
10        On Error GoTo Err_Handler

20        If m_Ribbon Is Nothing Then
              'User may have Shift-loaded the app and the startup code did not run.
30        Else
40            m_Ribbon.InvalidateControl "ddMRU"
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError "modRibbonCallback", "Ribbon_RefreshMRU"
80        Resume Exit_Handler
End Sub

Public Sub Ribbon_ShowReportsGroup()
10        On Error GoTo Err_Handler

20        If m_Ribbon Is Nothing Then
              'User may have Shift-loaded the app and the startup code did not run.
30        Else
40            m_Ribbon.InvalidateControl "gReportOptions"
              'After InvalidateControl, the group control will call gReportOptions_GetVisible.
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError "modRibbonCallback", "Ribbon_ShowReportsGroup"
80        Resume Exit_Handler
End Sub
