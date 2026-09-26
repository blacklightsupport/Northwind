Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'PURPOSE:
'   Delete most test data.
Private Sub cmdRemoveTestData_Click()
10        On Error GoTo Err_Handler

          Const MAX_COMPANYID     As Long = 13
          Const MAX_PRODUCTID     As Long = 43
          Const MAX_EMPLOYEEID    As Long = 10
          Const MAX_CONTACTID     As Long = 16
          Dim sql                 As String

20        If MsgBox("This will permanently delete most data. Are you sure? You might want to first make a backup.", vbQuestion Or vbYesNo) = vbYes Then
30            sql = "delete * from MRU;"
40            g_dbApp().Execute sql, dbFailOnError

50            sql = "delete * from Contacts where ContactID > " & MAX_CONTACTID
60            g_dbApp().Execute sql, dbFailOnError

              'Delete all orders and their order details.
70            sql = "delete * from OrderDetails where OrderID in (select OrderID from Orders where EmployeeID > " & MAX_EMPLOYEEID & ");"
80            g_dbApp().Execute sql, dbFailOnError
90            sql = "delete * from OrderDetails where OrderID in (select OrderID from Orders where CustomerID > " & MAX_COMPANYID & ");"
100           g_dbApp().Execute sql, dbFailOnError
110           sql = "delete * from OrderDetails where ProductID > " & MAX_PRODUCTID
120           g_dbApp().Execute sql, dbFailOnError

130           sql = "delete * from Orders where EmployeeID > " & MAX_EMPLOYEEID
140           g_dbApp().Execute sql, dbFailOnError
150           sql = "delete * from Orders where CustomerID > " & MAX_COMPANYID
160           g_dbApp().Execute sql, dbFailOnError

170           sql = "delete * from PurchaseOrderDetails where PurchaseOrderID in (select PurchaseOrderID from PurchaseOrders where SubmittedByID > " & MAX_EMPLOYEEID & ");"
180           g_dbApp().Execute sql, dbFailOnError
190           sql = "delete * from PurchaseOrderDetails where PurchaseOrderID in (select PurchaseOrderID from PurchaseOrders where ApprovedByID > " & MAX_EMPLOYEEID & ");"
200           g_dbApp().Execute sql, dbFailOnError

210           sql = "delete * from PurchaseOrderDetails where PurchaseOrderID in (select PurchaseOrderID from PurchaseOrders where VendorID > " & MAX_COMPANYID & ");"
220           g_dbApp().Execute sql, dbFailOnError
230           sql = "delete * from PurchaseOrderDetails where ProductID > " & MAX_PRODUCTID
240           g_dbApp().Execute sql, dbFailOnError

250           sql = "delete * from PurchaseOrders where SubmittedByID > " & MAX_EMPLOYEEID
260           g_dbApp().Execute sql, dbFailOnError
270           sql = "delete * from PurchaseOrders where ApprovedByID > " & MAX_EMPLOYEEID
280           g_dbApp().Execute sql, dbFailOnError
290           sql = "delete * from PurchaseOrders where VendorID > " & MAX_COMPANYID
300           g_dbApp().Execute sql, dbFailOnError


310           sql = "delete * from Companies where CompanyID > " & MAX_COMPANYID
320           g_dbApp().Execute sql, dbFailOnError

330           sql = "delete * from StockTake where ProductID > " & MAX_PRODUCTID
340           g_dbApp().Execute sql, dbFailOnError
350           sql = "delete * from ProductVendors where ProductID > " & MAX_PRODUCTID
360           g_dbApp().Execute sql, dbFailOnError

370           sql = "delete * from Products where ProductID > " & MAX_PRODUCTID
380           g_dbApp().Execute sql, dbFailOnError


390           sql = "delete * from EmployeePrivileges where EmployeeID > " & MAX_EMPLOYEEID
400           g_dbApp().Execute sql, dbFailOnError
410           sql = "delete * from Employees where EmployeeID > " & MAX_EMPLOYEEID
420           g_dbApp().Execute sql, dbFailOnError

430           sql = "update SystemSettings set SettingValue = -1 where SettingName = 'ShowWelcome'"
440           g_dbApp().Execute sql, dbFailOnError
450           sql = "update UserSettings set SettingValue = 0 where SettingName = 'AutoLogin'"
460           g_dbApp().Execute sql, dbFailOnError

470           MsgBox "Data Deleted. You may want to close and re-open the application.", vbInformation
480       End If

Exit_Handler:
490       Exit Sub

Err_Handler:
500       clsErrorHandler.HandleError Me.Name, "cmdRemoveTestData_Click", True
510       Resume Exit_Handler
520       Resume
End Sub
