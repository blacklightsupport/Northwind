Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'ENHANCEMENTS YOU CAN MAKE:

'1) Table Data Macros:
'   The data macro for the Orders table populates timestamp information, which shown in the form footer, right side.
'   These macros only fire when the data in the Orders table is changed.
'   When you add, change, or delete line items in the Order Detail, the Order table data macros don't fire because
'      you changed only the child records for the Order, NOT the Order record.
'   You could change the data macros on the OrderDetails table to also update the parent Order
'      record.

'2) Invoice Date, Ship Date and Paid Date:
'   We have not implemented rules around these dates, such as:
'      How far in the future can they be?
'      Can you enter them before the Order is Invoiced?
'   You can create any business rules you want and enforce them as needed in your environment.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'ALGORITHM:
'   Use the RecordsetClone to loop over all rows, checking if they all have the required status.
Private Function AllInvoiceLineItemsHaveStatus(ByVal ods As enumOrderDetailStatus) As Boolean
10        On Error GoTo Err_Handler

          Dim blnReturn           As Boolean

20        With Me.sfrmOrderLineItems.Form.RecordsetClone  'This "With" construct has the advantage it automatically cleans up the Recordset object after it ends. And it saves typing.
30            If .RecordCount = 0 Then
40                blnReturn = True
50            Else
60                .MoveFirst
70                blnReturn = True        'Optimistic.

80                Do While Not .EOF
90                    If !OrderDetailStatusID <> ods Then
                          'Different status found. No point in continuing.
100                       blnReturn = False
110                       Exit Do
120                   End If
130                   .MoveNext
140               Loop
150           End If
160       End With

170       AllInvoiceLineItemsHaveStatus = blnReturn

Exit_Handler:
180       Exit Function

Err_Handler:
190       clsErrorHandler.HandleError Me.Name, "AllInvoiceLineItemsHaveStatus"
200       Resume Exit_Handler
End Function

Private Sub cmdAddOrder_Click()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            MsgBox GetString(sAlreadyOnNewRecord), vbInformation
40        Else
50            RunCommand acCmdRecordsGoToNew
60        End If

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "cmdAddOrder_Click", True
90        Resume Exit_Handler
End Sub

Private Sub cmdCloseOrder_Click()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            MsgBox GetString(sOptionNotAvailable_NewRecord), vbExclamation
40        Else
50            If Me.OrderStatusID = enumOrderStatus.osPaid Then
60                RunCommand acCmdSaveRecord

70                If MsgBox(GetString(enumStrings.sOrderClosed), vbYesNo Or vbQuestion) = vbYes Then
80                    Me.OrderStatusID = enumOrderStatus.osClosed
90                    RunCommand acCmdSaveRecord      'Save record, so Requery below will pick up the new value.

100                   SetFormStatus
110               End If
120           Else
130               MsgBox GetString(enumStrings.sOrderMustBePaid), vbExclamation
140           End If
150       End If

Exit_Handler:
160       Exit Sub

Err_Handler:
170       clsErrorHandler.HandleError Me.Name, "cmdCloseOrder_Click", True
180       Resume Exit_Handler
End Sub

Private Sub cmdCreateInvoice_Click()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            MsgBox GetString(sOptionNotAvailable_NewRecord), vbExclamation
40        Else
50            RunCommand acCmdSaveRecord      'The report should pick up saved values.

              'BUSINESS RULE: Only "New" orders can be invoiced.
60            If Me.OrderStatusID = enumOrderStatus.osNew Then

                  'BUSINESS RULE: An order must have at least one line item.
70                If DCount("*", "OrderDetails", "OrderID = " & Me.OrderID) = 0 Then
80                    MsgBox GetString(sOneLineItem), vbExclamation
90                Else
                      'BUSINESS RULE: all order line items must be allocated, so we can invoice all at once. No partial invoices.
100                   If AllInvoiceLineItemsHaveStatus(enumOrderDetailStatus.odsAllocated) Then

110                       If ShippingFeeFilledOut() Then

120                           DoCmd.OpenReport "rptInvoice", acViewPreview, WhereCondition:="OrderID = " & Me.OrderID

130                           If Me.OrderStatusID = enumOrderStatus.osNew Then
140                               Me.OrderStatusID = enumOrderStatus.osInvoiced
150                               Me.InvoiceDate = Now()
160                               Me.cmdPrintInvoice.Visible = True
                                  'Set the status of invoiced line items.
170                               SetLineItemsStatus enumOrderDetailStatus.odsInvoiced
180                           Else
                                  'Status has already advanced from New, and user prints another invoice. No status change needed.
190                           End If

                              'Some controls got new values. Save again. We would not want user to hit Cancel and undo those changes.
                              'This won't work, since the report has focus.  RunCommand acCmdSaveRecord
200                           Me.Dirty = False

210                           SetFormStatus
220                       Else
230                           MsgBox "Shipping Fee is required for the invoice.", vbExclamation
240                       End If
250                   Else
260                       MsgBox GetString(sOrderAllocated), vbExclamation
270                   End If
280               End If
290           Else
300               MsgBox GetString(sNewOrderInvoice), vbExclamation
310           End If
320       End If

Exit_Handler:
330       Exit Sub

Err_Handler:
340       clsErrorHandler.HandleError Me.Name, "cmdCreateInvoice_Click", True
350       Resume Exit_Handler
End Sub

Private Sub cmdDeleteOrder_Click()
10        On Error GoTo Err_Handler

          'BUSINESS RULE: Orders can be Deleted until they are shipped.
20        If IsNull(Me.OrderID) Then      'Alternatively can test for Me.NewRecord.
30            MsgBox "This is a new order. There is nothing to delete.", vbInformation
40        Else
50            If Me.OrderStatusID = enumOrderStatus.osNew Or Me.OrderStatusID = enumOrderStatus.osInvoiced Then
60                RunCommand acCmdDeleteRecord
70            Else
80                MsgBox GetString(enumStrings.sOrderCannotDelete), vbExclamation
90            End If
100       End If

Exit_Handler:
110       Exit Sub

Err_Handler:
120       If Err.Number = 3021 Then   ' 3021 = No current record.
              'This happens when choosing No when prompted to delete the order, and the event is cancelled. Just ignore it and run the next line of code.
130           Resume Next
140       End If
150       clsErrorHandler.HandleError Me.Name, "cmdDeleteOrder_Click"
160       Resume Exit_Handler
170       Resume
End Sub

Private Sub cmdPrintInvoice_Click()
10        On Error GoTo Err_Handler

20        DoCmd.OpenReport "rptInvoice", acViewPreview, WhereCondition:="OrderID = " & Me.OrderID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdPrintInvoice_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdReceivePayment_Click()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            MsgBox GetString(sOptionNotAvailable_NewRecord), vbExclamation
40        Else
              'BUSINESS RULE: order must be shipped before it can be paid.
50            If Me.OrderStatusID = enumOrderStatus.osShipped Then

60                If PaidFieldsFilledOut() Then
70                    Me.OrderStatusID = enumOrderStatus.osPaid
80                    RunCommand acCmdSaveRecord      'This will trigger Form_BeforeUpdate where we will do additional validation.

90                    SetFormStatus
100               End If
110           Else
120               MsgBox GetString(sOrderMustBeShippedBeforePaid), vbExclamation
130           End If
140       End If

Exit_Handler:
150       Exit Sub

Err_Handler:
160       clsErrorHandler.HandleError Me.Name, "cmdReceivePayment_Click", True
170       Resume Exit_Handler
End Sub

Private Sub cmdShipOrder_Click()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            MsgBox GetString(sOptionNotAvailable_NewRecord), vbExclamation
40        Else
              'BUSINESS RULE: order must be invoiced before it can be shipped.
50            If Me.OrderStatusID = enumOrderStatus.osInvoiced Then
60                If ShippingFieldsFilledOut() Then
70                    Me.OrderStatusID = enumOrderStatus.osShipped
80                    RunCommand acCmdSaveRecord      'This will trigger Form_BeforeUpdate where we will do additional validation.

                      'Set line items to shipped.
90                    SetLineItemsStatus enumOrderDetailStatus.odsShipped

100                   SetFormStatus
110               End If
120           Else
130               MsgBox GetString(sOrderMustBeInvoiced), vbExclamation
140           End If
150       End If

Exit_Handler:
160       Exit Sub

Err_Handler:
170       clsErrorHandler.HandleError Me.Name, "cmdShipOrder_Click", True
180       Resume Exit_Handler
End Sub

Private Sub CustomerID_AfterUpdate()
10        On Error GoTo Err_Handler

          'Set default value for tax status.
20        Me.TaxStatusID = DLookup("StandardTaxStatusID", "Companies", "CompanyID = " & Me.CustomerID)

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "CustomerID_AfterUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

20        AddToMRU "Orders", Me.OrderID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
50        Resume Exit_Handler
End Sub

Private Sub Form_AfterUpdate()
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
10        On Error GoTo Err_Handler

20        Response = acDataErrContinue

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_BeforeDelConfirm", True
50        Resume Exit_Handler
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
10        On Error GoTo Err_Handler

          'Set initial values for new record. Fires when the first character is entered into a new record.
20        Me.EmployeeID = Get_UserID()
30        Me.TaxRate = GetSystemSetting(ssTaxRate)
40        Me.OrderStatusID = enumOrderStatus.osNew

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "Form_BeforeInsert", True
70        Resume Exit_Handler
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Cancel = ValidateForm(Me)

30        If Not Cancel Then
40            ValidateForm_RemoveHighlights Me
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
80        Resume Exit_Handler
End Sub

Private Sub Form_Close()
10        On Error GoTo Err_Handler

20        CloseOrderDetailsForm Me

          'Refresh Open List Forms
30        RequeryListForms

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_Close", True
60        Resume Exit_Handler
End Sub

Private Sub Form_Current()
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me        'Needed if user canceled a new record with yellow highlights on the form.

30        SetFormStatus

40        Me.Caption = "Order " & Me.OrderID      'Show the orderID in the caption, so several instances can easily be distinguished.

          'This is too severe. Need to set Locked property on a per-control basis and depends on OrderStatus.
          ' This is not implemented
          ' Me.AllowEdits = (Me.OrderStatusID = enumOrderStatus.osNew)
50        LockControls

          'ALTERNATIVE SYNTAX for .AllowAdditions = (Me.OrderStatusID = enumOrderStatus.osNew):
          'If Me.OrderStatusID = enumOrderStatus.osNew Then .AllowAdditions = True Else .AllowAdditions = False
          'or
          '.AllowAdditions = IIf(Me.OrderStatusID = enumOrderStatus.osNew, True, False)
60        If Me.NewRecord Then
70            With Me.sfrmOrderLineItems.Form
80                .AllowAdditions = True
90                .AllowEdits = True
100               .AllowDeletions = True
110           End With

120           Me.cmdPrintInvoice.Visible = False
130       Else
140           With Me.sfrmOrderLineItems.Form
150               .AllowAdditions = (Me.OrderStatusID = enumOrderStatus.osNew)
160               .AllowEdits = .AllowAdditions
170               .AllowDeletions = .AllowAdditions
180           End With

190           Me.cmdPrintInvoice.Visible = (Me.OrderStatusID <> enumOrderStatus.osNew)
200       End If

Exit_Handler:
210       Exit Sub

Err_Handler:
220       clsErrorHandler.HandleError Me.Name, "Form_Current", True
230       Resume Exit_Handler
End Sub

Private Sub Form_Delete(Cancel As Integer)
10        On Error GoTo Err_Handler

          Dim lngProductID        As Long

20        If MsgBox(GetString(enumStrings.sDeleteRecord, "order"), vbYesNo Or vbQuestion) = vbYes Then

30            lngProductID = 0
              'Delete of the Order causes a cascading delete of order line items
              'We want to delete each line item so we can reallocate inventory for each product
              'before the cascading delete happens
40            With Me.sfrmOrderLineItems.Form.RecordsetClone
50                If .RecordCount > 0 Then
60                    .MoveFirst
70                    While Not .EOF
80                        lngProductID = !ProductID
90                        .Delete
100                       AllocateInventory lngProductID
110                       .MoveNext
120                   Wend
130               End If
140           End With

150           RemoveFromMRU "Orders", Me.OrderID
160       Else
170           Cancel = True
180       End If

Exit_Handler:
190       Exit Sub

Err_Handler:
200       clsErrorHandler.HandleError Me.Name, "Form_Delete", True
210       Resume Exit_Handler
End Sub

Private Sub Form_Load()
10        On Error GoTo Err_Handler

          Dim dict                As Scripting.Dictionary

20        If IsNull(TempVars!OpenArgs) Then
              'New record.
30            If Not Me.NewRecord Then RunCommand acCmdRecordsGoToNew
40        Else
              'TempVars!OpenArgs holds the OrderID to be displayed. Format: "OrderID=123"
50            Set dict = StringToDictionary(TempVars!OpenArgs)

              'Although this code works perfectly, it is not a best practice. As the number of records grows (and Orders will be the fastest growing),
              'it is better to open a single record.

              '        'Use the recordsetclone to select the record with the requested OrderID.
              '        'Note that "with me.recordsetclone" automatically cleans up the temporary objects at "end with".
              '        With Me.RecordsetClone
              '            .FindFirst "OrderID = " & dict("OrderID")
              '            Debug.Assert Not .NoMatch       'If this asserts, the orderID we just clicked on is not in the current recordsource.
              '            Me.Bookmark = .Bookmark         'The bookmark technique selects the current record.
              '        End With

              'Open a single record.
60            Me.RecordSource = "select * from qryOrder where OrderID = " & dict("OrderID")

70            TempVars!OpenArgs = Null
80            Set dict = Nothing

90            If Not Me.NewRecord Then AddToMRU "Orders", Me.OrderID
100       End If

Exit_Handler:
110       Exit Sub

Err_Handler:
120       clsErrorHandler.HandleError Me.Name, "Form_Load", True
130       Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

          'For multi-instance forms, Form_Open is too soon to process OpenArgs and set the current record. Postpone until Form_Load.

Exit_Handler:
20        Exit Sub

Err_Handler:
30        clsErrorHandler.HandleError Me.Name, "Form_Open", True
40        Resume Exit_Handler
End Sub

'PURPOSE:
'   Based on order status, lock controls that can no longer be edited because that would be illogical. E.g. you cannot change the Order date after the order has been created.
'   These are the business rules for Northwind; in another company some fields may remain editable longer.
Private Sub LockControls()
10        On Error GoTo Err_Handler

          Dim varControls         As Variant      'The Array function returns a variant containing an array.
          Dim v                   As Variant

20        varControls = Array()       'Initialize, so it is no longer empty. Without this,error 13 = Type mismatch will occur when looping over varControls.

          'List the controls that should be locked for each status. Some controls like OrderID and TaxAmount can be omitted because they cannot be edited.
30        Select Case Me.OrderStatusID
              Case enumOrderStatus.osNew
40                varControls = Array("OrderID", "CustomerID", "EmployeeID", "OrderStatusName", "OrderDate", "InvoiceDate", "ShippedDate", "ShipperID", "PaymentMethod", "PaidDate", "txtTaxAmount", "txtTotal")

50            Case osInvoiced
60                varControls = Array("OrderID", "CustomerID", "EmployeeID", "OrderStatusName", "OrderDate", "InvoiceDate", "ShippingFee", "PaymentMethod", "PaidDate", "TaxRate", "TaxStatusID", "txtTaxAmount", "txtTotal")

70            Case osShipped
80                varControls = Array("OrderID", "CustomerID", "EmployeeID", "OrderStatusName", "OrderDate", "InvoiceDate", "ShippedDate", "ShipperID", "ShippingFee", "TaxRate", "TaxStatusID", "txtTaxAmount", "txtTotal")

90            Case osPaid
100               varControls = Array("OrderID", "CustomerID", "EmployeeID", "OrderStatusName", "OrderDate", "InvoiceDate", "ShippedDate", "ShipperID", "ShippingFee", "PaymentMethod", "PaidDate", "TaxRate", "TaxStatusID", "txtTaxAmount", "txtTotal")

110           Case osClosed
120               varControls = Array("OrderID", "CustomerID", "EmployeeID", "Notes", "OrderStatusName", "OrderDate", "InvoiceDate", "ShippedDate", "ShipperID", "ShippingFee", "PaymentMethod", "PaidDate", "TaxRate", "TaxStatusID", "txtTaxAmount", "txtTotal")

130           Case Else
140               If IsNull(Me.OrderStatusID) Then            'NOTE: You cannot have a Case Null, because comparisons with Null are Null, not True/False.
150                   varControls = Array("OrderID", "OrderStatusName", "OrderDate", "InvoiceDate", "ShippedDate", "ShipperID", "PaymentMethod", "PaidDate", "txtTaxAmount", "txtTotal")
160               Else
170                   Debug.Assert False      'Unexpected order status.
180               End If
190       End Select

          'Initialize Locked property to False for all controls.
200       On Error Resume Next            'Not all controls have an Enabled Property
210       For Each v In Me.Controls
220           v.Locked = False
230           v.Enabled = True
240           v.BackShade = 100      '100 means no shading.
250       Next v
260       On Error GoTo Err_Handler       'Back to normal error handler.

          'Set Locked property to True for the controls in the array.
270       For Each v In varControls
280           Me.Controls(v).Locked = True
290           Me.Controls(v).Enabled = False
300           Me.Controls(v).BackShade = 95      '95 means Darker 5%
310       Next v

Exit_Handler:
320       Exit Sub

Err_Handler:
330       clsErrorHandler.HandleError Me.Name, "Form_Open", True
340       Resume Exit_Handler
End Sub

Private Function PaidFieldsFilledOut() As Boolean
10        On Error GoTo Err_Handler

          Dim blnIsValid          As Boolean

20        blnIsValid = True       'Optimistic.

30        If IsNull(Me.PaymentMethod) Then
40            blnIsValid = False
50            HighlightControl Me.PaymentMethod
60        End If
70        If IsNull(Me.PaidDate) Then
80            blnIsValid = False
90            HighlightControl Me.PaidDate
100       End If

110       If Not blnIsValid Then
120           MsgBox GetString(sRequiredFields_Paid), vbExclamation
130       End If
140       PaidFieldsFilledOut = blnIsValid

Exit_Handler:
150       Exit Function

Err_Handler:
160       clsErrorHandler.HandleError Me.Name, "PaidFieldsFilledOut"
170       Resume Exit_Handler
End Function

'PURPOSE:
'   Set form and control properties based on current status.
Private Sub SetFormStatus()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            With Me.sfrmOrderLineItems.Form
40                .AllowAdditions = True
50                .AllowEdits = True
60                .AllowDeletions = True
70            End With

80            Me.cmdPrintInvoice.Visible = False
90        Else
100           With Me.sfrmOrderLineItems.Form
110               .AllowAdditions = (Me.OrderStatusID = enumOrderStatus.osNew)
120               .AllowEdits = .AllowAdditions
130               .AllowDeletions = .AllowAdditions
140           End With

150           Me.cmdPrintInvoice.Visible = (Me.OrderStatusID <> enumOrderStatus.osNew)
160       End If

          'Initialize workflow button colors. Alternatively we could set RGB values, but we want to support Themes.
170       Me.cmdCreateInvoice.ForeShade = 95      '95 means Darker 5%
180       Me.cmdShipOrder.ForeShade = 95
190       Me.cmdReceivePayment.ForeShade = 95
200       Me.cmdCloseOrder.ForeShade = 95

210       If Me.NewRecord Then
              'Nothing more to do. OrderStatusID is null.
220       Else
230           LockControls

240           If Me.OrderStatusID = enumOrderStatus.osNew Then GoTo Exit_Handler
250           Me.cmdCreateInvoice.ForeShade = 70      '70 means Darker 30%

260           If Me.OrderStatusID = enumOrderStatus.osInvoiced Then GoTo Exit_Handler
270           Me.cmdShipOrder.ForeShade = 70

280           If Me.OrderStatusID = enumOrderStatus.osShipped Then GoTo Exit_Handler
290           Me.cmdReceivePayment.ForeShade = 70

300           If Me.OrderStatusID = enumOrderStatus.osPaid Then GoTo Exit_Handler
310           Me.cmdCloseOrder.ForeShade = 70
320       End If

Exit_Handler:
330       Exit Sub

Err_Handler:
340       clsErrorHandler.HandleError Me.Name, "SetFormStatus"
350       Resume Exit_Handler
End Sub

Private Sub SetLineItemsStatus(ByVal ods As enumOrderDetailStatus)
10        On Error GoTo Err_Handler

          Dim sql                 As String

20        sql = StringFormatSQL("update OrderDetails set OrderDetailStatusID = {0} where OrderID = {1};", ods, Me.OrderID)
30        g_dbApp().Execute sql, dbFailOnError

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "SetLineItemsStatus"
60        Resume Exit_Handler
End Sub

Private Sub sfrmOrderLineItems_Enter()
          'It doesn't make sense to allow user to (begin to) enter a Child row while the Parent row is not saved, due to Referential integrity. This also prevents a possible bug with subforms and recordsetclones that MSFT has not yet fixed (as of Oct-2023).

          '10        If IsNull(Me.sfrmOrderLineItems.Form.OrderID) Then
          '20            MsgBox GetString(sOrderBeforeOrderLineItems), vbExclamation
          '30            Me.CustomerID.SetFocus
          '40        End If

10        If Me.NewRecord Then
20            MsgBox GetString(sOrderBeforeOrderLineItems), vbExclamation
30            Me.CustomerID.SetFocus
40        End If
End Sub

Private Sub sfrmOrderLineItems_Exit(Cancel As Integer)
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me.sfrmOrderLineItems.Form

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "sfrmOrderLineItems_Exit", True
50        Resume Exit_Handler
End Sub

Private Function ShippingFeeFilledOut() As Boolean
10        On Error GoTo Err_Handler

          Dim blnIsValid          As Boolean

20        blnIsValid = True       'Optimistic.

30        If IsNull(Me.ShippingFee) Then
40            blnIsValid = False
50            HighlightControl Me.ShippingFee
60        End If

70        ShippingFeeFilledOut = blnIsValid

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "ShippingFeeFilledOut"
100       Resume Exit_Handler
End Function

Private Function ShippingFieldsFilledOut() As Boolean
10        On Error GoTo Err_Handler

          Dim blnIsValid          As Boolean

20        blnIsValid = True       'Optimistic.

30        If IsNull(Me.ShippedDate) Then
40            blnIsValid = False
50            HighlightControl Me.ShippedDate
60        End If
70        If IsNull(Me.ShipperID) Then
80            blnIsValid = False
90            HighlightControl Me.ShipperID
100       End If
110       If IsNull(Me.ShippingFee) Then
120           blnIsValid = False
130           HighlightControl Me.ShippingFee
140       End If

150       If Not blnIsValid Then
160           MsgBox GetString(sRequiredFields_Shipping), vbExclamation
170       End If
180       ShippingFieldsFilledOut = blnIsValid

Exit_Handler:
190       Exit Function

Err_Handler:
200       clsErrorHandler.HandleError Me.Name, "ShippingFieldsFilledOut"
210       Resume Exit_Handler
End Function
