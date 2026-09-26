Attribute VB_Name = "modInventory"
Option Compare Database
Option Explicit

'PURPOSE:
'   Allocate Inventory to orders waiting for them.
Public Sub AllocateInventory(ByVal lngProductID As Long)
10        On Error GoTo Err_Handler

          Dim rsO                 As DAO.Recordset
          Dim intQtyToAllocate    As Integer
          Dim intAvailable        As Integer
          Dim sql                 As String

          'Get list of orderdetails waiting for this product; oldest one first. StatusID=1=Allocated, 4=NoStock, 5=OnOrder. OrderID is the tie breaker if same datetime.
20        sql = StringFormatSQL("select od.* from OrderDetails od " & _
                                "inner join Orders o on o.OrderID = od.OrderID " & _
                                "where od.ProductID = {0} and od.OrderDetailStatusID in (1, 4, 5) " & _
                                "order by o.OrderDate, o.OrderID", lngProductID)
30        Set rsO = g_dbApp().OpenRecordset(sql, dbOpenDynaset)

          'Physical inventory at this moment in time
40        intAvailable = ProductAvailable(lngProductID)

          'Physical Inventory + how much is On Order
50        intQtyToAllocate = intAvailable + ProductOnOrder(lngProductID)

60        While Not rsO.EOF       'Loop over the order line items with this product that are in status of Allocated, No Stock or On Order.
70            If rsO!Quantity <= intAvailable Then    'Only allocate if we can satisfy the entire orderdetail.
80                rsO.Edit
90                rsO!OrderDetailStatusID = enumOrderDetailStatus.odsAllocated
100               intAvailable = intAvailable - rsO!Quantity
110               intQtyToAllocate = intQtyToAllocate - rsO!Quantity
120               rsO.Update

130           ElseIf rsO!Quantity <= intQtyToAllocate Then    'Only OnOrder if we can satisfy the entire orderdetail.
140               rsO.Edit
150               rsO!OrderDetailStatusID = enumOrderDetailStatus.odsOnOrder
160               intQtyToAllocate = intQtyToAllocate - rsO!Quantity
170               rsO.Update

180           Else
                  'If we are here, we don't have enough
190               rsO.Edit
200               rsO!OrderDetailStatusID = enumOrderDetailStatus.odsNoStock
210               rsO.Update
220           End If

230           rsO.MoveNext
240       Wend
250       rsO.Close
260       Set rsO = Nothing

Exit_Handler:
270       Exit Sub

Err_Handler:
280       clsErrorHandler.HandleError "modInventory", "AllocateInventory"
290       Resume Exit_Handler
300       Resume
End Sub

'RETURNS:
'   Integer:    Quantity of this product that is in the given status.
Public Function OrderQuantity_ByStatus(ByVal lngProductID As Long, ByVal ods As enumOrderDetailStatus) As Integer
10        On Error GoTo Err_Handler

          Dim sql                 As String

20        If lngProductID = 0 Then GoTo Exit_Handler

30        sql = StringFormatSQL("ProductID = {0} and OrderDetailStatusID = {1}", lngProductID, ods)

40        OrderQuantity_ByStatus = Nz(DSum("Quantity", "OrderDetails", sql), 0)

Exit_Handler:
50        Exit Function

Err_Handler:
60        clsErrorHandler.HandleError "modInventory", "OrderQuantity_ByStatus"
70        Resume Exit_Handler
End Function

'RETURNS:
'   Integer:    Quantity of this product that is in Allocated status.
Public Function ProductAllocated(ByVal lngProductID As Long) As Integer
10        On Error GoTo Err_Handler

20        ProductAllocated = OrderQuantity_ByStatus(lngProductID, odsAllocated)

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modInventory", "ProductAllocated"
50        Resume Exit_Handler
End Function

'DEFINITION:
'   Last Stock Take Qty + Received since last stocktake - Invoiced since last stocktake.
Public Function ProductAvailable(ByVal lngProductID As Long) As Integer
10        On Error GoTo Err_Handler

          Dim dtLastStockTake     As Date
          Dim intLastStockTake    As Integer
          Dim intAvailable        As Integer
          Dim intBought           As Integer
          Dim intSold             As Integer

20        ProductAvailable = 0

30        If lngProductID = 0 Then GoTo Exit_Handler

          'Last stock take
40        dtLastStockTake = ProductLastStockTakeDate(lngProductID)
50        intLastStockTake = ProductLastStockTakeQuantity(lngProductID)

          'Sold
60        intSold = ProductSold(lngProductID, dtLastStockTake)

          'Bought
70        intBought = ProductBought(lngProductID, dtLastStockTake)

          'Allen Browne's formula - Adapted with permission from http://allenbrowne.com/AppInventory.html
80        intAvailable = intLastStockTake + intBought - intSold

90        ProductAvailable = intAvailable

Exit_Handler:
100       Exit Function

Err_Handler:
110       clsErrorHandler.HandleError "modInventory", "ProductAvailable"
120       Resume Exit_Handler
130       Resume
End Function

'DEFINITION
'   Bought = When product on a PO is Received.
'   ProductBought quantity: Sum of Product received since last stocktake (or since dtAsOf).
'RETURNS:
'   Integer: Quantity of this product that was received after point in time.
'   If you wanted all purchases you would pass g_dtNorthwindInception which is 11/01/2022
Public Function ProductBought(ByVal lngProductID As Long, ByVal dtAsOf As Date) As Integer
10        On Error GoTo Err_Handler

          Dim sql                 As String

20        If lngProductID = 0 Then GoTo Exit_Handler

30        sql = StringFormatSQL("ProductID = {0} and ReceivedDate >= {1}", lngProductID, dtAsOf)

40        ProductBought = Nz(DSum("Quantity", "PurchaseOrderDetails", sql), 0)

Exit_Handler:
50        Exit Function

Err_Handler:
60        clsErrorHandler.HandleError "modInventory", "ProductBought"
70        Resume Exit_Handler
End Function

'PURPOSE:
'   Returns the Date with Time of the Last StockTake for a Product.
'   Northwind Expects there to be a StockTake.  When a Product is added the first StockTake record is entered.
'   If there is no StockTake (it has been deleted), Add a Stock Take record the same as what was created when the Product was added.
'      Set  StockTakeDate using the Date the Product was Added
'           QuantityOnHand = 0
Public Function ProductLastStockTakeDate(ByVal lngProductID As Long) As Date
10        On Error GoTo Err_Handler

          Dim sql                 As String
          Dim rsStockTake         As DAO.Recordset

20        ProductLastStockTakeDate = g_dtNorthwindInception

30        If lngProductID = 0 Then GoTo Exit_Handler

40        sql = StringFormatSQL("select * from StockTake where ProductID = {0} order by StockTakeDate desc;", lngProductID)
50        Set rsStockTake = g_dbApp().OpenRecordset(sql, dbOpenDynaset)
60        If rsStockTake.RecordCount = 0 Then
              'No stock take was ever done for this Product.
70            With rsStockTake
80                .AddNew
90                !StockTakeDate = Nz(DLookup("AddedOn", "Products", "ProductID = " & lngProductID), Now())
100               !ProductID = lngProductID
110               !QuantityOnHand = 0
120               .Update
130               .Move 0, .LastModified    'Move to the record just added, so we can read from it.
140           End With
150       End If

160       ProductLastStockTakeDate = rsStockTake!StockTakeDate

170       rsStockTake.Close
180       Set rsStockTake = Nothing

Exit_Handler:
190       Exit Function

Err_Handler:
200       clsErrorHandler.HandleError "modInventory", "ProductLastStockTakeDate"
210       Resume Exit_Handler
End Function

'RETURNS:
'   Integer: Quantity in the last (most recent) StockTake for the Product
Public Function ProductLastStockTakeQuantity(ByVal lngProductID As Long) As Integer
10        On Error GoTo Err_Handler

          Dim dtStockTakeDate     As Date

20        If lngProductID = 0 Then GoTo Exit_Handler

30        dtStockTakeDate = ProductLastStockTakeDate(lngProductID)

          'ERROR: This may return Null because of roundoff errors.  ProductLastStockTakeQuantity = DLookup("QuantityOnHand", "StockTake", StringFormatSQL("ProductID = {0} and StockTakeDate = CDate({1})", lngProductID, dtStockTakeDate))
          'NOTE: We're using ToAccessDate (which returns a string) on both sides of the expression "StockTakeDate = some_date". This is because of possible roundoff errors if we were to compare
          '      the date values themselves (which are floating point values). In one particular case we had this roundoff error:
          '? DLookup("StockTakeDate", "StockTake", "ProductID = 3") - #11/27/2022 9:02:00 AM#
          '7.27595761418343E-12
          'By comparing the text values (as returned by ToAccessDate), we avoid this roundoff.
40        ProductLastStockTakeQuantity = DLookup("QuantityOnHand", "StockTake", StringFormatSQL("ProductID = {0} and ToAccessDate(StockTakeDate) = {1}", lngProductID, ToAccessDate(dtStockTakeDate)))

Exit_Handler:
50        Exit Function

Err_Handler:
60        clsErrorHandler.HandleError "modInventory", "ProductLastStockTakeQuantity"
70        Resume Exit_Handler
End Function

'RETURNS:
'   Integer: Quantity of this product that is in No Stock status.
Public Function ProductNoStock(ByVal lngProductID As Long) As Integer
10        On Error GoTo Err_Handler

20        ProductNoStock = OrderQuantity_ByStatus(lngProductID, odsNoStock)

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modInventory", "ProductNoStock"
50        Resume Exit_Handler
End Function

'DEFINITION:
'   Sum of Product Quantity for POs in the Approved status.
'NOTE:
'   There can be more than one PO for the same Product, therefore we sum the quantities in qryPOProducts_ByStatus.
'RETURNS:
'   Integer: Quantity of this product that is on an Approved Purchase Order
'
'We need a query in this function because we need two tables which are joined
'   PurchaseOrder - contains the criteria for the POStatus of Approved
'   and PurchaseOrderDetail - to support the criteria for ProductID.
Public Function ProductOnOrder(ByVal lngProductID As Long) As Integer
10        On Error GoTo Err_Handler

          Dim sqlWhere            As String

20        If lngProductID = 0 Then GoTo Exit_Handler

30        sqlWhere = StringFormatSQL("StatusID={0} and ProductID={1}", enumPurchaseOrderStatus.posApprove, lngProductID)

40        ProductOnOrder = Nz(DLookup("Quantity", "qryPOProducts_ByStatus", sqlWhere), 0)     'No need for DSum: qryPOProducts_ByStatus already is a Totals query.

Exit_Handler:
50        Exit Function

Err_Handler:
60        clsErrorHandler.HandleError "modInventory", "ProductOnOrder"
70        Resume Exit_Handler
End Function

'PURPOSE:
'   Calculate the minimum Reorder Quantity
'ALGORITHM:
'To determine minimum Reorder Qty we need to consider
'     how much we have available    intQtyAvailable
'   + how much we have OnOrder      intOnOrder
'then
'     how much NoStock (what we need; some of which may already be On Order
'   + the product Target Level
Public Function ProductReorderQuantity(ByVal lngProductID As Long) As Integer
10        On Error GoTo Err_Handler

          Dim intReorderQty       As Integer
          Dim intNoStock          As Integer
          Dim intMinReorder       As Integer
          Dim intOnOrder          As Integer
          Dim intQtyAvailable     As Integer
          Dim intTargetLevel      As Integer

20        intReorderQty = 0

30        If lngProductID = 0 Then GoTo Exit_Handler

40        intNoStock = ProductNoStock(lngProductID)

50        intOnOrder = ProductOnOrder(lngProductID)    'PO Status Approved.

60        intQtyAvailable = ProductToSell(lngProductID)

70        intMinReorder = Nz(DLookup("MinimumReorderQuantity", "Products", "ProductID = " & lngProductID), 1)
80        intTargetLevel = Nz(DLookup("TargetLevel", "Products", "ProductID = " & lngProductID), 0)

90        If (intQtyAvailable + intOnOrder) >= (intNoStock + intTargetLevel) Then
              ' we dont need to order anything, return the minimum reorder quantity for the Product.
100           intReorderQty = intMinReorder
110       Else
              'we need to reorder.
120           intReorderQty = (intNoStock + intTargetLevel) - (intQtyAvailable + intOnOrder)
130           If intMinReorder > intReorderQty Then    'Return which ever is greater
                  'we need less than the minimum reorder quantity - return the minimum reorder quantity.
140               intReorderQty = intMinReorder
150           End If
160       End If

170       ProductReorderQuantity = intReorderQty

Exit_Handler:
180       Exit Function

Err_Handler:
190       clsErrorHandler.HandleError "modInventory", "ProductReorderQuantity"
200       Resume Exit_Handler
End Function

'RETURNS:
'   Integer: Quantity of this product that was sold after a point in time
'   If you wanted all sales you would pass g_dtNorthwindInception which is 11/01/2022
Public Function ProductSold(ByVal lngProductID As Long, ByVal dtAsOf As Date) As Integer
10        On Error GoTo Err_Handler

          Dim sql                 As String

20        If lngProductID = 0 Then GoTo Exit_Handler

          '         NOTE: Nz(date_field) returns 12/30/1899 (day zero) if it is null, which is less than our dtAsOf so nulls are not returned.
30        sql = StringFormatSQL("ProductID = {0} and OrderID IN (Select OrderID FROM Orders WHERE Nz(InvoiceDate) >= {1})", lngProductID, dtAsOf)

40        ProductSold = Nz(DSum("Quantity", "OrderDetails", sql), 0)

Exit_Handler:
50        Exit Function

Err_Handler:
60        clsErrorHandler.HandleError "modInventory", "ProductSold"
70        Resume Exit_Handler
End Function

'DEFINITION:
'   Last Stock Take Qty + Received since last stocktake - Invoiced since last stocktake - allocated to existing orders.
Public Function ProductToSell(ByVal lngProductID As Long) As Integer
10        On Error GoTo Err_Handler

          Dim dtLastStockTake     As Date
          Dim intAllocated        As Integer
          Dim intAvailable        As Integer
          Dim intBought           As Integer
          Dim intSold             As Integer
          Dim intLastStockTake    As Integer

20        If lngProductID = 0 Then GoTo Exit_Handler

          'Last stock take
30        dtLastStockTake = ProductLastStockTakeDate(lngProductID)
40        intLastStockTake = ProductLastStockTakeQuantity(lngProductID)

          'Sold
50        intSold = ProductSold(lngProductID, dtLastStockTake)

          'Bought
60        intBought = ProductBought(lngProductID, dtLastStockTake)

          'Allocated
70        intAllocated = ProductAllocated(lngProductID)

80        intAvailable = intLastStockTake + intBought - intSold - intAllocated

90        ProductToSell = intAvailable

Exit_Handler:
100       Exit Function

Err_Handler:
110       clsErrorHandler.HandleError "modInventory", "ProductToSell"
120       Resume Exit_Handler
End Function
