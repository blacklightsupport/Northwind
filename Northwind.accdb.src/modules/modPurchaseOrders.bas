Attribute VB_Name = "modPurchaseOrders"
Option Compare Database
Option Explicit

'RETURNS:
'   Long    - New line item ID value, or 0 if the line item could not be added.
Public Function AddPurchaseOrderDetail(ByVal lngPurchaseOrderID As Long, ByVal lngProductID As Long _
                                                                         , ByVal intQuantity As Integer, ByVal curUnitCost As Currency) As Long
10        On Error GoTo Err_Handler

          Dim strSQL              As String
          Dim rs                  As DAO.Recordset

          'Create the Purchase Order Line Item for this product: begin with an 0 or 1 Line Item recordset
20        strSQL = StringFormatSQL("SELECT PurchaseOrderDetailID, PurchaseOrderID, ProductID, Quantity, UnitCost" _
                                   & " FROM PurchaseOrderDetails" _
                                   & " WHERE PurchaseOrderID = {0} and ProductID = {1};", lngPurchaseOrderID, lngProductID)

30        Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
40        If rs.EOF Then
              'Add new line.
50            rs.AddNew
60            rs!PurchaseOrderID = lngPurchaseOrderID
70            rs!ProductID = lngProductID
80            rs!Quantity = intQuantity
90            rs!UnitCost = curUnitCost
100       Else
              'Line item already exists.
              'BUSINESS RULE: Blindly add to the quantity that is already there.
110           rs.Edit
120           rs!Quantity = rs!Quantity + intQuantity
130       End If
140       rs.Update

          'Retrieve the new ID
150       rs.Move 0, rs.LastModified
160       AddPurchaseOrderDetail = rs!PurchaseOrderDetailID
170       rs.Close

Exit_Handler:
180       Set rs = Nothing
190       Exit Function

Err_Handler:
200       clsErrorHandler.HandleError "modPurchaseOrders", "AddPurchaseOrderDetail"
210       Resume Exit_Handler
220       Resume
End Function

Public Sub CloseAllPurchaseOrderDetailsForms()
10        On Error GoTo Err_Handler

          Dim intCount            As Integer
          Dim i                   As Integer

20        If Not g_colPurchaseOrderDetailsForms Is Nothing Then
30            intCount = g_colPurchaseOrderDetailsForms.count
40            For i = 1 To intCount
50                g_colPurchaseOrderDetailsForms.Remove 1
60            Next i
70        End If

80        Set g_colPurchaseOrderDetailsForms = Nothing

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError "modPurchaseOrders", "CloseAllPurchaseOrderDetailsForms"
110       Resume Exit_Handler
End Sub

'PURPOSE:
'   Called from frmPurchaseOrderDetails.Form_Close to remove object from collection.
Public Sub ClosePurchaseOrderDetailsForm(frm As Form)
10        On Error GoTo Err_Handler

          Dim blnRemoved          As Boolean
          Dim intCount            As Integer
          Dim i                   As Integer

20        If g_colPurchaseOrderDetailsForms Is Nothing Then
              'This may happen when the app is being shut down. Ignore it. We are on the way out anyway.
30        Else
40            intCount = g_colPurchaseOrderDetailsForms.count
50            For i = 1 To intCount
60                If g_colPurchaseOrderDetailsForms.Item(i).Hwnd = frm.Hwnd Then
70                    g_colPurchaseOrderDetailsForms.Remove i
80                    blnRemoved = True
90                    Exit For
100               End If
110           Next i
              'Debug.Assert blnRemoved     'This will assert if form was not opened using OpenPurchaseOrderDetailsForm.
120       End If

Exit_Handler:
130       Exit Sub

Err_Handler:
140       clsErrorHandler.HandleError "modPurchaseOrders", "ClosePurchaseOrderDetailsForm"
150       Resume Exit_Handler
End Sub

'PURPOSE:
'   Support for opening multiple PurchaseOrderDetails forms.
Public Sub OpenPurchaseOrderDetailsForm(Optional ByVal varPurchaseOrderID As Variant)
10        On Error GoTo Err_Handler

          Dim frm                 As Form

20        If g_colPurchaseOrderDetailsForms Is Nothing Then Set g_colPurchaseOrderDetailsForms = New Collection

          'Must set TempVars before creating New form, or it is too late.
30        If Not IsMissing(varPurchaseOrderID) Then TempVars!OpenArgs = "PurchaseOrderID=" & varPurchaseOrderID     'Make OpenArgs self-describing by using name=value pairs like a querystring. StringToDictionary function can be used to pick it apart.
40        Set frm = New Form_frmPurchaseOrderDetails
          'Does not work: OpenArgs is readonly property.  frm.OpenArgs = "OrderID=" & varPurchaseOrderID
50        frm.Visible = True

60        g_colPurchaseOrderDetailsForms.Add frm, CStr(frm.Hwnd)

          'POSSIBLE IMPROVEMENT: if using overlapped windows: Move relative to previous item in the collection using frm.Move, Otherwise the windows are stacked exactly on top of each other.


70        Set frm = Nothing

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError "modPurchaseOrders", "OpenPurchaseOrderDetailsForm"
100       Resume Exit_Handler
End Sub

'PURPOSE:
'   Creates new Purchase order and adds a line item.
'RETURNS:
'   New Purchase Order ID
Public Function ReorderProduct(ByVal lngProductID As Long, ByVal lngVendorID As Long, _
                               ByVal intQuantity As Integer, ByVal curUnitCost As Currency) As Long

10        On Error GoTo Err_Handler

          Dim strSQL              As String
          Dim rs                  As DAO.Recordset
          Dim lngNewID            As Long
          Dim lngNewLineID        As Long

20        ReorderProduct = 0

          'Create a new Purchase Order and retrieve the PurchaseOrderID
          'Note: "WHERE PurchaseOrderID = -1" ensures an empty recordset, to which we will add a new record.
30        strSQL = "SELECT PurchaseOrderID, VendorID, StatusID" _
                   & " FROM PurchaseOrders" _
                   & " WHERE PurchaseOrderID = -1;"

40        Debug.Print strSQL

50        Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)

60        With rs
70            .AddNew
80            !VendorID = lngVendorID
90            !StatusID = enumPurchaseOrderStatus.posNew
100           .Update

              'retrieve the new Purchase Order ID
110           .Move 0, .LastModified          'Alternative syntax: .Bookmark = .LastModified
120           lngNewID = !PurchaseOrderID

130           .Close
140       End With

150       ReorderProduct = lngNewID

160       lngNewLineID = AddPurchaseOrderDetail(lngNewID, lngProductID, intQuantity, curUnitCost)

170       If lngNewLineID = 0 Then
180           MsgBox StringFormat("Purchase Order {0} was created, but creating a line item for ProductID={1} failed." _
                                  , lngNewLineID, lngProductID), vbExclamation Or vbOKOnly, "Line Item Fail"
190       End If

Exit_Handler:
200       Set rs = Nothing
210       Exit Function

Err_Handler:
220       clsErrorHandler.HandleError "modPurchaseOrders", "ReorderProduct"
230       Resume Exit_Handler
240       Resume
End Function
