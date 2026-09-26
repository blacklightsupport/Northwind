Attribute VB_Name = "modGlobal"
Option Compare Database
Option Explicit

'Declaring all global variables here.

'Version history
'Version    ReleaseDate     Notes
'2.0        2023-04-24      Initial release.
'2.1        2023-04-27      Support for international dates and quantities.
'2.2        Aug-2023        Round of bug fixing and improvements based on community feedback.
'                           Consolidate one-time code in OneTimeProcessing.
'                           Better support for regional settings. For example dates and currencies should be displayed in your local settings.
'                           Improved StringFormatSQL.
'                           Improved error handler (clsErrorHandler.HandleError).
'2.3        Nov-2023        Fixed a few bugs reported by the community.
'                           Updated GetSystemSetting to support regional settings.
'                           Prevent data entry in OrderDetail subform until parent record is created. This works around a bug in v2.2 which MSFT hasn't fixed yet.
'2.4        Sep-2024        Fix several accessibility issues.
'2.5        Sep-2025        Change BackStyle of 2 textboxes from Transparent to Solid. Improve comments at top of clsErrorHandler. Added HiddenAndSystemObjectsWorkaround.

Public Const APP_VERSION    As String = "2.5"
Public Const INTERNET_SALES_EMPLOYEEID As Long = 10
Public Const DEFAULT_LOGIN_ID As Long = 2
Public Const TWIPS_PER_INCH As Long = 1440

Public Const SINGLE_QUOTE   As String = "'"           'It is common to have to wrap text in a single quote, or search for it in a string.
Public Const TWO_SINGLE_QUOTES As String = "''"     'Sometimes a single quote needs to be escaped into two single quotes. Adding this for readability.

Public g_colOrderDetailsForms As Collection             'To keep track of multiple frmOrderDetails instances (see modOrders).
Public g_colPurchaseOrderDetailsForms As Collection     'To keep track of multiple frmPurchaseOrderDetails instances (see modPurchaseOrders).
Private m_dbApp             As DAO.Database
Private m_UserID            As Long

'Keep this in sync with table CompanyType
'NOTE: the "ct" = "Company Type" prefix ensures there are no Reserved words in the Enum names.
Public Enum enumCompanyType
    ctAll = 0
    ctCustomer = 1
    ctShipper = 2
    ctVendor = 3
    ctNorthwind = 4
End Enum

Public Enum enumOrderStatus             'Keep this in sync with table OrderStatus.
    osClosed = 1                        'NOTE: the "os" = "Order Status" prefix ensures there are no Reserved words in the Enum names.
    osInvoiced = 2
    osNew = 3
    osShipped = 4
    osPaid = 5
End Enum

Public Enum enumPurchaseOrderStatus     'Keep this in sync with PurchaseOrderStatus.
    posApprove = 1
    posClosed = 2
    posNew = 3
    posSubmitted = 4
    posReceived = 5
End Enum

Public Enum enumOrderDetailStatus       'Keep this in sync with table OrderDetailStatus.
    odsAllocated = 1
    odsInvoiced = 2
    odsNew = 3
    odsNoStock = 4
    odsOnOrder = 5
    odsShipped = 6
End Enum

Public Enum enumPrivileges              'Keep this in sync with table Privileges.
    pApprovePO = 1
End Enum

Public Enum enumSystemSettings          'Keep this in sync with table SystemSettings
    ssTaxRate = 1
    ssLastResetDate = 4
    ssShowWelcome = 5
    ssTaxRate_Vendors = 6
    ssFirstTimeRun = 7
End Enum

Public Enum enumUserSettings            'Keep this in sync with table UserSettings
    usAutoLogin = 2
End Enum

'Keep this in sync with table NorthwindFeatures
'NOTE: the "nf" = "Company Type" prefix ensures there are no Reserved words in the Enum names.
Public Enum enumNorthwindFeaturesOpenMethod
    nfInPageHelp = 1
    nfFollowHyperlink = 2
End Enum

'Keep this in sync with table TaxStatus
'NOTE: the "tx" = "Tax Status" prefix ensures there are no Reserved words in the Enum names.
Public Enum enumTaxStatus
    txTaxExempt = 0
    txTaxable = 1
End Enum

'Espana UG - Add to MRU
'PURPOSE:
'   Add the given item to the MRU table and update the MRU dropdown in the Ribbon.
Public Sub AddToMRU(ByVal strTableName As String, ByVal lngPKValue As Long)
10        On Error GoTo Err_Handler

          Const MAX_MRU_COUNT     As Integer = 10
          Dim rs                  As DAO.Recordset
          Dim lngMin              As Long
          Dim sql                 As String

          'ddMRU_OnAction (the callback function used by the Ribbon and implemented in modRibbonCallback) currently supports 2 tables.
20        Debug.Assert strTableName = "Orders" Or strTableName = "PurchaseOrders"

30        sql = StringFormatSQL("insert into MRU(EmployeeID, TableName, PKValue, DateAdded) values ({0}, {1}, {2}, {3});", _
                                Get_UserID(), strTableName, lngPKValue, Now())
40        g_dbApp().Execute sql       'Do not use dbFailOnError because the record may already exist.  ', dbFailOnError

          'Trim back MRU list if it is getting too long. First get the value below which the records must be deleted.
50        sql = StringFormatSQL("select Min(MRU_ID) from (select top {0} MRU_ID from MRU where EmployeeID={1} order by MRU_ID)", MAX_MRU_COUNT, Get_UserID())
60        Set rs = g_dbApp().OpenRecordset(sql, dbOpenSnapshot)
70        If IsNull(rs(0)) Then
              'Empty MRU list. Nothing to trim back.
80        Else
90            lngMin = rs(0)
100           sql = StringFormatSQL("delete * from MRU where MRU_ID < {0} and EmployeeID = {1};", lngMin, Get_UserID())
110           g_dbApp().Execute sql, dbFailOnError
120       End If
130       rs.Close

          'Tell the ribbon element to update itself.
140       Ribbon_RefreshMRU

Exit_Handler:
150       Exit Sub

Err_Handler:
160       clsErrorHandler.HandleError "modGlobal", "AddToMRU"
170       Resume Exit_Handler
End Sub

Public Function GetNorthwindAddress() As String
10        On Error GoTo Err_Handler

          Static strNorthwindAddress As String        'Our address rarely changes, so OK to cache it.
          Dim rs                  As DAO.Recordset
          Dim sql                 As String

20        If strNorthwindAddress = "" Then
30            sql = "select * from Companies where CompanyTypeId = " & enumCompanyType.ctNorthwind
40            Set rs = g_dbApp().OpenRecordset(sql, dbOpenSnapshot)
50            strNorthwindAddress = StringFormat("{0}{1}{2} {3}, {4}", rs!CompanyName & vbCrLf, rs!Address & vbCrLf, rs!City, rs!StateAbbrev, rs!Zip)
60            rs.Close
70            Set rs = Nothing
80        End If
90        GetNorthwindAddress = strNorthwindAddress

Exit_Handler:
100       Exit Function

Err_Handler:
110       clsErrorHandler.HandleError "modGlobal", "GetNorthwindAddress"
120       Resume Exit_Handler
End Function

Public Function GetWindowsUserName() As String
10        On Error GoTo Err_Handler

20        GetWindowsUserName = Left(Environ("UserName"), 255)

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modGlobal", "GetWindowsUserName"
50        Resume Exit_Handler
End Function

Public Function Get_EmployeeFNLN(EmployeeID As Long) As String
10        On Error GoTo Err_Handler

          Dim varLookup           As Variant

20        varLookup = DLookup("FullNameFNLN", "qryEmployees", "EmployeeID = " & EmployeeID)

30        If IsNull(varLookup) Then
              'User not found in the employee table
40            Get_EmployeeFNLN = "Error Employee Not Found"
50        Else
60            Get_EmployeeFNLN = varLookup
70        End If

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError "modGlobal", "Get_EmployeeFNLN"
100       Resume Exit_Handler
End Function

'PURPOSE:
'   Used by queries, which cannot access global variables directly, and other code. Self-healing version.
'   Alternatively we could have used a TempVar.
Public Function Get_UserID() As Long
10        On Error GoTo Err_Handler

20        If m_UserID = 0 Then

30        Else
40            Get_UserID = m_UserID
50        End If

Exit_Handler:
60        Exit Function

Err_Handler:
70        clsErrorHandler.HandleError "modGlobal", "Get_UserID"
80        Resume Exit_Handler
End Function

Public Function Get_UserID_ForWindowsUser() As Long
10        On Error GoTo Err_Handler

          Dim varLookup           As Variant

          'What is the EmployeeID of the the Windows User
20        varLookup = DLookup("EmployeeID", "qryEmployees", StringFormatSQL("WindowsUserName = {0}", GetWindowsUserName()))

30        If IsNull(varLookup) Then
              'Current windows user not found in the employee table
40            Get_UserID_ForWindowsUser = 0
50        Else
60            Get_UserID_ForWindowsUser = varLookup
70        End If

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError "modGlobal", "Get_UserID_ForWindowsUser"
100       Resume Exit_Handler
End Function

'PURPOSE:
'   Self-healing function to return the current database.
'NOTE:
'   Previous versions of NW2 used a global variable, but it would not be set if the app is loaded with the Shift key down, or if the code is reset.
Public Function g_dbApp() As DAO.Database

10        If m_dbApp Is Nothing Then
20            Set m_dbApp = CurrentDb
30        End If

40        Set g_dbApp = m_dbApp

End Function

Public Sub OneTimeProcessing()

10        On Error GoTo Err_Handler

20        If Not GetSystemSetting(ssFirstTimeRun) Then

30            SysCmd acSysCmdSetStatus, "One-Time Processing. Please stand by."
40            DoCmd.Hourglass True

50            HiddenAndSystemObjectsWorkaround

60            SetDatesToCurrent     'So new instance is working with current data.

70            SetCtrlCurrencyFormat

80            AddDataMacros

90            SaveSystemSetting ssFirstTimeRun, -1      '-1 is better than True in international scenarios.

100           SysCmd acSysCmdClearStatus
110       End If

Exit_Handler:
120       DoCmd.Hourglass False
130       Exit Sub

Err_Handler:
140       clsErrorHandler.HandleError "modGlobal", "OneTimeProcessing"
150       Resume Exit_Handler
160       Resume
End Sub

'PURPOSE:
'   Returns standardized string for record create/modification
'   If the Windows UserName has been recorded in the Employees table, use that name rather than
Public Function RecordModifiedString(Optional AddedBy As Variant = "", Optional AddedOn As Variant = "", _
                                     Optional ModifiedBy As Variant = "", Optional ModifiedOn As Variant = "") As String
10        On Error GoTo Err_Handler

          Dim strTemp             As String

20        strTemp = ""

30        If (AddedOn & AddedBy) > "" Then
40            strTemp = "Created"
50            If AddedOn > "" Then strTemp = strTemp & " " & AddedOn
60            If AddedBy > "" Then strTemp = strTemp & " by " & AddedBy
70        End If

80        If (ModifiedBy & ModifiedOn) > "" Then
90            If strTemp > "" Then strTemp = strTemp & " ~ "
100           strTemp = strTemp & " Modified"

110           If ModifiedOn > "" Then strTemp = strTemp & " " & ModifiedOn
120           If ModifiedBy > "" Then strTemp = strTemp & " by " & ModifiedBy
130       End If

140       RecordModifiedString = strTemp

Exit_Handler:
150       Exit Function

Err_Handler:
160       clsErrorHandler.HandleError "modGlobal", "RecordModifiedString"
170       Resume Exit_Handler
End Function

'Espana Remove From MRU
Public Sub RemoveFromMRU(ByVal strTableName As String, ByVal lngPKValue As Long)
10        On Error GoTo Err_Handler

          Dim sql                 As String

20        sql = StringFormatSQL("delete * from MRU where TableName = {0} and PKValue = {1};", strTableName, lngPKValue)
30        g_dbApp().Execute sql, dbFailOnError

          'Tell the ribbon element to update itself.
40        Ribbon_RefreshMRU

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError "modGlobal", "RemoveFromMRU"
70        Resume Exit_Handler
End Sub

'PURPOSE:
'   The template is created in the United States. Other regions will want to display currency values in their format.
'   As of Aug-2023 Access will preserve the US format by changing the Format property to $#.##0,00;($#.##0,00).
'   The code in this procedure will set them back to "Currency".
Private Sub SetCtrlCurrencyFormat()
10        On Error GoTo Err_Handler

          Dim i                   As Integer
          Dim aryForms()          As Variant
          Dim aryControls()       As Variant
          Dim aryReports()        As Variant

20        aryForms = Array("frmOrderDetails", "frmOrderDetails", "frmOrderDetails", "frmOrderList", "frmOrderList", "frmProductDetail", "frmProductDetail", "frmPurchaseOrderDetails", "frmPurchaseOrderDetails", "frmPurchaseOrderDetails", "frmPurchaseOrderList", "frmPurchaseOrderList", _
                           "sfrmCompanyDetail_CustomerOrders", "sfrmCompanyDetail_ShipperOrders", "sfrmOrderLineItems", "sfrmOrderLineItems", "sfrmOrderLineItems", "sfrmOrders_MostRecent_ByEmployee", "sfrmProductDetail_Orders", "sfrmProductDetail_Orders", "sfrmProductDetail_PurchaseOrders", "sfrmProductDetail_PurchaseOrders", _
                           "sfrmPurchaseOrderLineItems", "sfrmPurchaseOrderLineItems", "sfrmPurchaseOrderLineItems")
30        aryControls = Array("ShippingFee", "txtTaxAmount", "txtTotal", "OrderTotal", "txtSumOrderTotal", "txtUnitPrice", "txtStandardUnitCost", "ShippingFee", "txtTaxAmount", "txtTotal", "TotalCost", "txtSumTotalCost", _
                              "txtOrderTotal", "txtOrderTotal", "UnitPrice", "Price", "txtSubTotal", "OrderTotal", "txtUnitPrice", "txtExtendedPrice", "UnitCost", "ExtendedCost", _
                              "UnitCost", "Price", "txtSubTotal")
40        Debug.Assert UBound(aryForms) = UBound(aryControls)

50        For i = 0 To UBound(aryForms)
60            DoCmd.OpenForm aryForms(i), acDesign, WindowMode:=acHidden
70            Forms(aryForms(i)).Controls(aryControls(i)).Format = "Currency"
80            DoCmd.Close acForm, aryForms(i), acSaveYes
90        Next i

100       Erase aryForms
110       Erase aryControls

120       DoCmd.OpenForm "frmReports", , , , , acHidden    'Criteria in reports' recordsources reference the Reports form

          'NOTE: aryReports(0) corresponds to aryControls(0) etc. There are several reports to be fixed. Each may have several controls with Currency.
130       aryReports = Array("rptInvoice", "rptInvoice", "rptInvoice", "rptInvoice", "rptInvoice", _
                             "rptProductCatalog", "rptProductCatalog", "rptProductCatalog", "rptProductCatalog", _
                             "rptSalesByEmployee", "rptSalesByEmployee", "rptSalesByEmployee", _
                             "rptSalesByProduct", "rptSalesByProduct", "rptSalesByProduct", _
                             "rptSalesByProductQuarterly", "rptSalesByProductQuarterly", "rptSalesByProductQuarterly")
140       aryControls = Array("txtUnitPrice", "txtExtendedPrice", "txtSubTotal", "txtShipping", "txtTotal", _
                              "txtTotalSales", "txtMinPrice", "txtMaxPrice", "UnitPrice", _
                              "OrderTotal", "SumOfOrderTotal", "GrandTotalSum", _
                              "OrderTotal", "SumOfOrderTotal", "GrandTotalSum", _
                              "OrderTotal", "SumOfOrderTotal", "GrandTotalSum")
150       Debug.Assert UBound(aryReports) = UBound(aryControls)

160       For i = 0 To UBound(aryReports)
170           DoCmd.OpenReport aryReports(i), acDesign, WindowMode:=acHidden
180           Reports(aryReports(i)).Controls(aryControls(i)).Format = "Currency"
190           DoCmd.Close acReport, aryReports(i), acSaveYes
200       Next i

210       Erase aryReports
220       Erase aryControls

          'Modern chart properties.
230       On Error Resume Next            'Access 2016 and older does not support Modern Chart, so the controls would not exist.
240       aryReports = Array("rptSalesByEmployee", "rptSalesByProduct", "rptSalesByProductQuarterly")
250       aryControls = Array("chrtEmplyeeSalesByMonth", "chrtSalesByProduct", "chrtSalesByProduct")
260       For i = 0 To UBound(aryReports)
270           DoCmd.OpenReport aryReports(i), acDesign, WindowMode:=acHidden
280           Reports(aryReports(i)).Controls(aryControls(i)).PrimaryValuesAxisFormat = "Currency"
290           DoCmd.Close acReport, aryReports(i), acSaveYes
300       Next i

310       DoCmd.Close acForm, "frmReports", acSaveNo
320       Erase aryReports
330       Erase aryControls

Exit_Handler:
340       Exit Sub

Err_Handler:
350       clsErrorHandler.HandleError "modGlobal", "SetCtrlCurrencyFormat"
360       Resume Exit_Handler
370       Resume
End Sub

'PURPOSE:
'   Convert the date to US date format as expected by Access. Needed when building a SQL string with literal.
'   This ISO date format also has the advantage there are no issues with sorting strings, or comparing with > or <.
Public Function ToAccessDate(ByVal dt As Date) As String
10        On Error GoTo Err_Handler

20        ToAccessDate = Format(dt, "yyyy-mm-dd hh:nn:ss")

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modGlobal", "ToAccessDate"
50        Resume Exit_Handler
End Function

Public Sub InitializeUser()
10        On Error GoTo Err_Handler

          Dim bolAutoLogIn        As Boolean
          Dim strWindowsUserName  As String
          Dim varEmployeeID       As Variant

          'Get the User's Login
20        strWindowsUserName = GetWindowsUserName()

          'Is the User in the Employee table?
30        varEmployeeID = DLookup("EmployeeID", "Employees", StringFormatSQL("WindowsUserName = {0}", strWindowsUserName))
40        If IsNull(varEmployeeID) Then
              'The current Windows user is not in the database. Prompt to create a new account.
50            DoCmd.OpenForm "frmCredentials", acNormal, , , , acDialog
              'At this point frmCredentials is hidden, and thus "falls out of the modal loop", so we can inspect its properties before closing it.
60            m_UserID = Form_frmCredentials.UserID
70            DoCmd.Close acForm, "frmCredentials"
80        Else
90            m_UserID = varEmployeeID

              'Has user checked the box on frmLogin requesting that we automatically log them in?
              'If setting not found, assume False
100           bolAutoLogIn = Nz(GetUserSetting(usAutoLogin), False)

110           If bolAutoLogIn = False Then
120               DoCmd.OpenForm "frmLogin", acNormal, , , , acDialog
                  'At this point frmLogin is hidden, and thus "falls out of the modal loop", so we can inspect its properties before closing it.
130               m_UserID = Form_frmLogin.UserID
140               DoCmd.Close acForm, "frmLogin"
150           End If
160       End If

Exit_Handler:
170       Exit Sub

Err_Handler:
180       clsErrorHandler.HandleError "modGlobal", "InitializeUser"
190       Resume Exit_Handler
End Sub

'PURPOSE:
'   Work around imperfection in Access ACCDT generation where USysRibbons table is ignored. This does not happen if hidden and system objects are showing,
'   so we ship NW2 with them showing, and turn them off here, so users experience the normal list of objects, not hidden and system objects.
Private Sub HiddenAndSystemObjectsWorkaround()
10        Application.SetOption "Show Hidden Objects", False
20        Application.SetOption "Show System Objects", False
End Sub
