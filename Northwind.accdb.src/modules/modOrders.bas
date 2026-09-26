Attribute VB_Name = "modOrders"
Option Compare Database
Option Explicit

Public Sub CloseAllOrderDetailsForms()
10        On Error GoTo Err_Handler

          Dim intCount            As Integer
          Dim i                   As Integer

20        If Not g_colOrderDetailsForms Is Nothing Then
30            intCount = g_colOrderDetailsForms.count
40            For i = 1 To intCount
50                g_colOrderDetailsForms.Remove 1
60            Next i
70        End If

80        Set g_colOrderDetailsForms = Nothing

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError "modOrders", "CloseAllOrderDetailsForms"
110       Resume Exit_Handler
End Sub

'PURPOSE:
'   Called from frmOrderDetails.Form_Close to remove object from collection.
Public Sub CloseOrderDetailsForm(frm As Form)
10        On Error GoTo Err_Handler

          Dim blnRemoved          As Boolean
          Dim intCount            As Integer
          Dim i                   As Integer

20        If g_colOrderDetailsForms Is Nothing Then
              'This may happen when the app is being shut down. Ignore it. We are on the way out anyway.
30        Else
40            intCount = g_colOrderDetailsForms.count
50            For i = 1 To intCount
60                If g_colOrderDetailsForms.Item(i).Hwnd = frm.Hwnd Then
70                    g_colOrderDetailsForms.Remove i
80                    blnRemoved = True
90                    Exit For
100               End If
110           Next i
              'Debug.Assert blnRemoved     'This will assert if form was not opened using OpenOrderDetailsForm.
120       End If

Exit_Handler:
130       Exit Sub

Err_Handler:
140       clsErrorHandler.HandleError "modOrders", "CloseOrderDetailsForm"
150       Resume Exit_Handler
End Sub

'PURPOSE:
'   Simulate internet orders, or just for more testing.
Public Sub CreateRandomOrders(ByVal intOrderCount As Integer)
10        On Error GoTo Err_Handler

          Dim rsOrders            As DAO.Recordset
          Dim rsOrderDetails      As DAO.Recordset
          Dim intAvailable        As Integer
          Dim intOrder            As Integer
          Dim intDetail           As Integer
          Dim intDetails          As Integer
          Dim lngOrderID          As Long

20        Set rsOrders = CurrentDb.OpenRecordset("Orders", dbOpenDynaset)
30        Set rsOrderDetails = CurrentDb.OpenRecordset("OrderDetails", dbOpenDynaset)

40        For intOrder = 1 To intOrderCount
50            With rsOrders
60                .AddNew
70                !EmployeeID = INTERNET_SALES_EMPLOYEEID
80                !CustomerID = GetRandomCustomerID()
90                !OrderDate = Now                            'Now is better than Date: we want to report how long ago the last order was placed. Also for analysis purposes such as when are the busy hours.
100               !TaxRate = GetSystemSetting(ssTaxRate)
110               !TaxStatusID = GetTaxStatusID(!CustomerID)
120               !OrderStatusID = enumOrderStatus.osNew
130               !Notes = "Internet Order"
140               .Update
                  'Get the autonumber value just created.
150               .Move 0, .LastModified
160               lngOrderID = !OrderID
170           End With

              'Create 2-5 random line items for this order.
180           intDetails = GetRandom(2, 5)
190           For intDetail = 2 To intDetails
200               With rsOrderDetails
210                   .AddNew
TryAgain:
220                   !OrderID = lngOrderID
230                   !ProductID = GetRandomProductID()
240                   !Quantity = GetRandom(5, 50)

250                   intAvailable = ProductAvailable(!ProductID)
260                   If intAvailable >= !Quantity Then
                          'We have stock. Allocate it.
270                       !OrderDetailStatusID = enumOrderDetailStatus.odsAllocated
280                   Else
                          'No or not enough stock. The purchasing department should create a new PO for this product (handled in Purchase Order module).
290                       !OrderDetailStatusID = enumOrderDetailStatus.odsNoStock
300                   End If

310                   !UnitPrice = DLookup("UnitPrice", "Products", "ProductID = " & !ProductID)
320                   .Update

330                   .Move 0, .LastModified      'Move to record just added, so we can get the AutoNumber value.

                      'Check if the orderdetailstatus can be advanced from Nostock to OnOrder.
340                   If !OrderDetailStatusID = enumOrderDetailStatus.odsNoStock Then AllocateInventory !ProductID

350               End With

360           Next intDetail

              'Add the new order to the MRU list.
370           AddToMRU "Orders", lngOrderID

380       Next intOrder

390       rsOrderDetails.Close
400       rsOrders.Close

Exit_Handler:
410       Exit Sub

Err_Handler:
420       If Err.Number = 3022 Then   '3022 = The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship. Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again.
              'This is normal: when choosing random numbers, occasionally the same number comes up. Just try again.
430           Resume TryAgain
440       Else
450           clsErrorHandler.HandleError "modOrders", "CreateRamdomOrders"
460           Resume Exit_Handler
470       End If
480       Resume
End Sub

Public Function GetRandomProductID() As Long
10        On Error GoTo Err_Handler

20        GetRandomProductID = GetRandomPkValue("Products", "ProductID")

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modOrders", "GetRandomProductID"
50        Resume Exit_Handler
End Function

'PURPOSE:
'   Support for opening multiple OrderDetails forms.
Public Sub OpenOrderDetailsForm(Optional ByVal varOrderID As Variant)
10        On Error GoTo Err_Handler

          Dim frm                 As Form

20        If g_colOrderDetailsForms Is Nothing Then Set g_colOrderDetailsForms = New Collection

          'Must set TempVars before creating New form, or it is too late for Form_Open.
30        If Not IsMissing(varOrderID) Then TempVars!OpenArgs = "OrderID=" & varOrderID     'Make OpenArgs self-describing by using name=value pairs like a querystring. StringToDictionary function can be used to pick it apart.
40        Set frm = New Form_frmOrderDetails
          'Does not work: OpenArgs is readonly property.  frm.OpenArgs = "OrderID=" & varOrderID
50        frm.Visible = True

60        g_colOrderDetailsForms.Add frm, CStr(frm.Hwnd)

          'POSSIBLE IMPROVEMENT: if using overlapped windows: Move relative to previous item in the collection using frm.Move, Otherwise the windows are stacked exactly on top of each other.

70        Set frm = Nothing

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError "modOrders", "OpenOrderDetailsForm"
100       Resume Exit_Handler
End Sub

'PURPOSE:
'   Set all dates in the database relative to today, so they remain close to today's date.
'   This is unlike previous version of Northwind where the dates were always in 2006.
'NOTE:
'   This is a super simple algorithm. For example it does not work well if all order dates are in 2024, and one is in 2026, and then you click this button. Improve as-needed, and share your code if you can.
Public Sub SetDatesToCurrent()
10        On Error GoTo Err_Handler

          Dim dtMax               As Date
          Dim intDelta            As Integer
          Dim td                  As DAO.TableDef
          Dim fld                 As DAO.Field
          Dim db                  As DAO.Database
          Dim sql                 As String

20        dtMax = Nz(DMax("OrderDate", "Orders"), Date)       'Nz in case there are no orders, and DMax returns Null.
30        intDelta = DateDiff("d", dtMax, Date)
40        If intDelta < 0 Then GoTo Exit_Handler              'Prevent error 3316: Shipped Date and Paid Date must be on or after the Order Date

50        Set db = CurrentDb
60        For Each td In db.TableDefs
70            If td.Attributes And DAO.TableDefAttributeEnum.dbSystemObject Then
                  'System table. Do nothing.
80            Else
90                For Each fld In td.Fields
100                   If fld.Type = DAO.DataTypeEnum.dbDate Then
110                       Debug.Print td.Name, fld.Name
                          'Example: update Orders set OrderDate = dateadd('d', 123, OrderDate)
120                       sql = StringFormat("update {0} set {1} = dateadd('d', {2}, {1})", td.Name, fld.Name, intDelta)
130                       db.Execute sql, dbFailOnError
140                   End If
150               Next fld
160           End If
170       Next td

Exit_Handler:
180       Exit Sub

Err_Handler:
190       clsErrorHandler.HandleError "modOrders", "SetDatesToCurrent"
200       Resume Exit_Handler
210       Resume
End Sub
