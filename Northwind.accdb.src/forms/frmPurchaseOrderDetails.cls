Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'ENHANCEMENTS YOU CAN MAKE:

'1) Table Data Macros:

'   The data macro for the PurchaseOrders table populates timestamp information, which is shown in the form footer, right side.
'   These macros fire only when the data in the PurchaseOrders table is changed.
'   If you add, change, or delete line items in the PurchaseOrders table, data macros don't fire because
'      you changed only the child records for the Purchase Order, NOT the Purchase Order record.
'   You could change the data macros on the PurchaseOrderDetails table to also update
'     the parent PurchaseOrders record.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'ALGORITHM:
'   Allocate (=Set the OrderDetailStatus from OnOrder to Allocated) from the oldest orders to the newest ones.
'   The overage (if any) is added to Stock
'
Private Sub AllocatePurchaseOrder(ByVal lngPurchaseOrderID As Long)
10        On Error GoTo Err_Handler

          Dim rsPO                As DAO.Recordset
          Dim intPOQty            As Integer
          Dim sql                 As String

          'Since update queries quickly became not updatable, we will use recordsets.
20        sql = "select * from PurchaseOrderDetails where PurchaseOrderID = " & lngPurchaseOrderID
30        Set rsPO = g_dbApp().OpenRecordset(sql, dbOpenSnapshot)

40        While Not rsPO.EOF      'Loop over the line items of this PO.

50            intPOQty = rsPO!Quantity
60            AllocateInventory rsPO!ProductID

70            rsPO.MoveNext
80        Wend

90        rsPO.Close
100       Set rsPO = Nothing

Exit_Handler:
110       Exit Sub

Err_Handler:
120       clsErrorHandler.HandleError Me.Name, "AllocatePurchaseOrder"
130       Resume Exit_Handler
End Sub

'PURPOSE: cmdAddNostock_Click
'   Add all products supplied by this vendor, not already on the Purchase Order,
'   with  a status of NO Stock on Order Details; saving the user from having to remember
'   and enter every item they want to order from the vendor.

'FINDING ITEMS WITH NO STOCK FOR THE VENDOR NOT ON THIS PO:
'   The SQL statement below could have been done via a query.
'   We chose not to do that in order to highlight some of the benefits of the StringFormatSQL function
'   and this provides an opportunity to document/explain what we needed to do and how we accomplished that

'IF THIS WERE A QUERY:
'   1) In order to explain what needs to be done (aka document),
'       we would have placed the following statements in the "properties" property of the query
'       and hope that
'       a)you would right click on the query to see if we put anything there
'       or
'       b) you chose the navigation pane option to View By --> Details
'
'   2) There are two parameters required by the Query; PurchaseOrderID and VendorID
'       We would have created tempVars or public functions which can be accessed
'       by the query to pass the parameters and populated them before executing the query
'
'   Using StringFormatSQL means we don't need to do either of these,
'   and StringFormatSQL makes sure the parameters are formatted correctly;
'   i.e. dates have the # before and after a date value, and strings are enclosed in double quotes.

'FINDING THE No Stock ORDER DETAILS FOR THE VENDOR:
'
'   WHERE (((Products.ProductID) Not In (Select ProductID from PurchaseOrderDetails Where PurchaseOrderID = {1}))
'++   Determine which Products are NOT on the Order
'
'   AND ((ProductVendors.VendorID)={2})
'++     Select Only Products this Vendor supplies
'
'   AND ((ProductNoStock([Products].[ProductID]))>0))
'++     Call the function ProductNoStock to ensure there is at least one OrderDetail with a status of NoStock for the Product

'UTILIZING PRE-BUILT CALCULATIONS:
'   This query utilizes two pre-built calculations - meaning calculations done via a public function
'   1) ReorderQuantity
'   2) ProductNoStock
'
'   Those functions were created for use on the Product Detail form.
'   At the time they were created this task was not in the requirements; it came much later.
'   It is a best practice to create functions for all calculations.
'   This ensures that all parts of an application that need the calculation, always do it the same way
'   You can find all places in the application where the calculation is used
'   and if a change is needed to the calculation, you can evaluate if all places it is used need to be changed
'   and if not create a new function for the new calculation.
Private Sub cmdAddNostock_Click()
10        On Error GoTo Err_Handler

          Dim strSQL              As String

20        If Me.NewRecord Then
30            MsgBox "Please first create the main purchase order record. Then this button can add all No Stock line items.", vbInformation
40        Else
50            strSQL = "INSERT INTO PurchaseOrderDetails (PurchaseOrderID, ProductID, Quantity, UnitCost) " & _
                       "SELECT {0} AS POrderID, Products.ProductID, ProductReorderQuantity([Products].[ProductID]) AS ProductReorderQuantity, Products.StandardUnitCost " & _
                       "FROM Products INNER JOIN ProductVendors ON Products.ProductID = ProductVendors.ProductID " & _
                       "WHERE (((Products.ProductID) Not In (Select ProductID from PurchaseOrderDetails Where PurchaseOrderID = {0})) AND (ProductVendors.VendorID = {1}) AND ((ProductNoStock([Products].[ProductID])) > 0));"

60            strSQL = StringFormatSQL(strSQL, Me.PurchaseOrderID, Me.VendorID)

70            g_dbApp().Execute strSQL, dbFailOnError

              'Requery vs Refresh
              'https://support.microsoft.com/office/refresh-or-requery-data-76e8c8b5-b8b4-49eb-a7d5-a472e742366c
              'Refresh does not take into account **ADDED or DELETED** records
              'Requery will take into account all **ADDED, UPDATED and DELETED**

80            Me.sfrmPurchaseOrderLineItems.Requery
90        End If

Exit_Handler:
100       Exit Sub

Err_Handler:
110       clsErrorHandler.HandleError Me.Name, "cmdAddNostock_Click", True
120       Resume Exit_Handler
End Sub

'NOTE:
'   This form does not prevent the form and subform from being edited after that has become illogical (or illegal).
'   Refer to frmOrderDetails.SetFormStatus for a way to implement it.

Private Sub cmdAddPO_Click()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            MsgBox GetString(sAlreadyOnNewRecord), vbInformation
40        Else
50            RunCommand acCmdRecordsGoToNew
60        End If

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "cmdAddPO_Click", True
90        Resume Exit_Handler
End Sub

Private Sub cmdApprovePO_Click()
10        On Error GoTo Err_Handler

          'PO must already have been submitted.
20        If Me.StatusID = enumPurchaseOrderStatus.posSubmitted Then

              'Check if user has privilege to approve POs.
30            If HasPrivilege(enumPrivileges.pApprovePO) Then
40                Me.ApprovedByID = Get_UserID()
50                Me.ApprovedDate = Now
60                Me.StatusID = enumPurchaseOrderStatus.posApprove

70                Me.Dirty = False    'Save the record.

80                AllocatePurchaseOrder Me.PurchaseOrderID

90                SetWorkflowLabelColors

100               MsgBox GetString(sNewStatusSet), vbInformation
110           Else
120               MsgBox GetString(sNoPrivilege, "Approve PO"), vbExclamation
130           End If
140       Else
150           MsgBox GetString(sPOMustBeSubmitted), vbExclamation
160       End If

          'NOTE: implicitly the PO is now sent to Vendor, and we are waiting to receive it.

Exit_Handler:
170       Exit Sub

Err_Handler:
180       clsErrorHandler.HandleError Me.Name, "cmdApprovePO_Click", True
190       Resume Exit_Handler
End Sub

Private Sub cmdClosePO_Click()
10        On Error GoTo Err_Handler

20        If Me.StatusID = enumPurchaseOrderStatus.posReceived Then
30            If PoCloseFieldsFilledOut() Then
40                Me.PaymentDate = Now
50                Me.TaxAmount = Me.txtTaxAmount
60                Me.PaymentAmount = Me.txtTotal
70                Me.StatusID = enumPurchaseOrderStatus.posClosed

80                Me.Dirty = False

90                SetWorkflowLabelColors

100               MsgBox GetString(sNewStatusSet), vbInformation
110           End If
120       Else
130           MsgBox GetString(sPOMustBeReceived), vbExclamation
140       End If

Exit_Handler:
150       Exit Sub

Err_Handler:
160       clsErrorHandler.HandleError Me.Name, "cmdClosePO_Click", True
170       Resume Exit_Handler
End Sub

'NOTE:
'   We allow deletion only for New and Submitted. If you change this business rule, be aware you will need to adjust Inventory along the same lines as when deleting an order or orderlineitem.
Private Sub cmdDeletePO_Click()
10        On Error GoTo Err_Handler

20        If Me.StatusID = enumPurchaseOrderStatus.posNew Or Me.StatusID = enumPurchaseOrderStatus.posSubmitted Then
              'The relationship between PurchaseOrders and PurchaseOrderDetails has cascade delete, so all we need to do is delete the parent record.
30            RunCommand acCmdDeleteRecord
40            If IsFormOpen("frmPurchaseOrderList") Then Forms!frmPurchaseOrderList.Requery
50        Else
60            MsgBox GetString(sPOCannotDelete), vbExclamation
70        End If

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "cmdDeletePO_Click", True
100       Resume Exit_Handler
End Sub

Private Sub cmdReceivePO_Click()
10        On Error GoTo Err_Handler

          Dim sql                 As String

20        If Me.StatusID = enumPurchaseOrderStatus.posApprove Then
30            If MsgBox(GetString(sPostToInventory), vbQuestion Or vbYesNo) = vbYes Then
                  'The entire PO has been delivered. Allocate product to orders that are waiting for it, and put the rest in inventory.
40                Me.ReceivedDate = Now
50                Me.StatusID = enumPurchaseOrderStatus.posReceived

60                Me.Dirty = False    'Save the record.

                  'Set PO line items ReceivedDate to full date/time. Important for inventory calculations.
70                sql = StringFormatSQL("update PurchaseOrderDetails set ReceivedDate = {0} where PurchaseOrderID = {1};", Now(), Me.PurchaseOrderID)
80                g_dbApp().Execute sql, dbFailOnError

90                SetWorkflowLabelColors

100               MsgBox GetString(sNewStatusSet), vbInformation

110               AllocatePurchaseOrder Me.PurchaseOrderID
120           End If
130       Else
140           MsgBox GetString(sPOMustBeApproved), vbExclamation
150       End If

Exit_Handler:
160       Exit Sub

Err_Handler:
170       clsErrorHandler.HandleError Me.Name, "cmdReceivePO_Click", True
180       Resume Exit_Handler
End Sub

Private Sub cmdSubmitPO_Click()
10        On Error GoTo Err_Handler

20        If Me.StatusID = enumPurchaseOrderStatus.posNew Then
30            Me.SubmittedByID = Get_UserID()
40            Me.SubmittedDate = Now
50            Me.StatusID = enumPurchaseOrderStatus.posSubmitted

60            Me.Dirty = False    'Save the record.

70            SetWorkflowLabelColors

80            MsgBox GetString(sNewStatusSet), vbInformation
90        Else
100           MsgBox GetString(sPOMustBeNew), vbExclamation
110       End If

Exit_Handler:
120       Exit Sub

Err_Handler:
130       clsErrorHandler.HandleError Me.Name, "cmdSubmitPO_Click", True
140       Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

20        AddToMRU "PurchaseOrders", Me.PurchaseOrderID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
50        Resume Exit_Handler
End Sub

Private Sub Form_AfterUpdate()
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me

30        If IsFormOpen("frmPurchaseOrderList") Then Forms!frmPurchaseOrderList.Requery

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_AfterUpdate", True
60        Resume Exit_Handler
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
10        On Error GoTo Err_Handler

          'Set initial values for new record. Fires when the first character is entered into a new record.
20        Me.StatusID = enumPurchaseOrderStatus.posNew

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_BeforeInsert", True
50        Resume Exit_Handler
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

20        ClosePurchaseOrderDetailsForm Me

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

30        Me.Caption = "Purchase Order " & Me.PurchaseOrderID      'Show the PurchaseOrderID in the caption, so several instances can easily be distinguised.

40        LockControls

50        SetWorkflowLabelColors

          'Allow edits based on order status
60        If Me.NewRecord Then
70            With Me.sfrmPurchaseOrderLineItems.Form
80                .AllowAdditions = True
90                .AllowEdits = True
100               .AllowDeletions = True
110           End With
120       Else
130           With Me.sfrmPurchaseOrderLineItems.Form
140               .AllowAdditions = (Me.StatusID = enumPurchaseOrderStatus.posNew)
150               .AllowEdits = .AllowAdditions
160               .AllowDeletions = .AllowAdditions
170           End With
180       End If

Exit_Handler:
190       Exit Sub

Err_Handler:
200       clsErrorHandler.HandleError Me.Name, "Form_Current", True
210       Resume Exit_Handler
End Sub

Private Sub Form_Delete(Cancel As Integer)
10        On Error GoTo Err_Handler

20        RemoveFromMRU "PurchaseOrders", Me.PurchaseOrderID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Delete", True
50        Resume Exit_Handler
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

              'Use the recordsetclone technique to select the record with the requested PurchaseOrderID.
              'Note that "with me.recordsetclone" automatically cleans up the temporary objects at "End With".
              'Alternatively you can assign a RecordSource that has only 1 record. See frmOrderDetails.
60            With Me.RecordsetClone
70                .FindFirst "PurchaseOrderID = " & dict("PurchaseOrderID")
80                Debug.Assert Not .NoMatch       'If this asserts, the orderID we just clicked on is not in the current recordsource.
90                Me.Bookmark = .Bookmark
100           End With

110           TempVars!OpenArgs = Null
120           Set dict = Nothing

130           If Not Me.NewRecord Then AddToMRU "PurchaseOrders", Me.PurchaseOrderID
140       End If

Exit_Handler:
150       Exit Sub

Err_Handler:
160       clsErrorHandler.HandleError Me.Name, "Form_Load", True
170       Resume Exit_Handler
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

20        varControls = Array()       'Initialize, so it is no longer empty.

          'List the controls that should be locked for each status. Some controls like OrderID and TaxAmount can be omitted because they cannot be edited.
30        Select Case Me.StatusID
              Case enumPurchaseOrderStatus.posNew
40                varControls = Array("PurchaseOrderID", "VendorID", "StatusName", "SubmittedBy", "SubmittedDate", "ApprovedBy", "ApprovedDate", "ReceivedDate", "PaymentMethod", "ShippingFee", "txtTaxAmount", "txtTotal", "PaymentDate")

50            Case posSubmitted
60                varControls = Array("PurchaseOrderID", "VendorID", "StatusName", "SubmittedBy", "SubmittedDate", "ApprovedBy", "ApprovedDate", "ReceivedDate", "PaymentMethod", "ShippingFee", "txtTaxAmount", "txtTotal", "PaymentDate")

70            Case posApprove
80                varControls = Array("PurchaseOrderID", "VendorID", "StatusName", "SubmittedBy", "SubmittedDate", "ApprovedBy", "ApprovedDate", "PaymentMethod", "ShippingFee", "txtTaxAmount", "txtTotal", "PaymentDate")

90            Case posReceived
100               varControls = Array("PurchaseOrderID", "VendorID", "StatusName", "SubmittedBy", "SubmittedDate", "ApprovedBy", "ApprovedDate", "ReceivedDate", "txtTaxAmount", "txtTotal", "PaymentDate")

110           Case posClosed
120               varControls = Array("PurchaseOrderID", "VendorID", "Notes", "StatusName", "SubmittedBy", "SubmittedDate", "ApprovedBy", "ApprovedDate", "ReceivedDate", "PaymentMethod", "ShippingFee", "txtTaxAmount", "txtTotal", "PaymentDate")
130               Me.AllowEdits = False

140           Case Else
150               If IsNull(Me.StatusID) Then            'NOTE: You cannot have a Case Null, because comparisons with Null are Null, not True/False.
160                   varControls = Array("PurchaseOrderID", "StatusName", "SubmittedBy", "SubmittedDate", "ApprovedBy", "ApprovedDate", "ReceivedDate", "PaymentMethod", "ShippingFee", "txtTaxAmount", "txtTotal", "PaymentDate")
170               Else
180                   Debug.Assert False      'Unexpected order status.
190               End If
200       End Select

          'Initialize Locked property to False for all controls.
210       On Error Resume Next            'Not all controls have an Enabled Property
220       For Each v In Me.Controls
230           v.Locked = False
240           v.Enabled = True
250           v.BackShade = 100      '100 means no shading.
260       Next v
270       On Error GoTo Err_Handler       'Back to normal error handler.

          'Set Locked property to True for the controls in the array.
280       For Each v In varControls
290           Me.Controls(v).Locked = True
300           Me.Controls(v).Enabled = False
310           Me.Controls(v).BackShade = 95      '95 means Darker 5%
320       Next v

Exit_Handler:
330       Exit Sub

Err_Handler:
340       clsErrorHandler.HandleError Me.Name, "PoCloseFieldsFilledOut"
350       Resume Exit_Handler
End Sub

'PURPOSE:
'   Check if all fields for the Close PO status have been filled out.
Private Function PoCloseFieldsFilledOut() As Boolean
10        On Error GoTo Err_Handler

          Dim blnIsValid          As Boolean

20        blnIsValid = True       'Optimistic.

30        If IsNull(Me.ShippingFee) Then
40            blnIsValid = False
50            HighlightControl Me.ShippingFee
60        End If
70        If IsNull(Me.PaymentMethod) Then
80            blnIsValid = False
90            HighlightControl Me.PaymentMethod
100       End If

110       If Not blnIsValid Then
120           MsgBox GetString(sRequiredFields_PO_Close), vbExclamation
130       End If
140       PoCloseFieldsFilledOut = blnIsValid

Exit_Handler:
150       Exit Function

Err_Handler:
160       clsErrorHandler.HandleError Me.Name, "PoCloseFieldsFilledOut"
170       Resume Exit_Handler
End Function

' PURPOSE:
'   Set the foreground color of the 4 workflow labels in the form header.
Private Sub SetWorkflowLabelColors()
10        On Error GoTo Err_Handler

          'Initialize. Alternatively we could set RGB values, but we want to support Themes.
20        Me.cmdSubmitPO.ForeShade = 95      '95 means Darker 5%
30        Me.cmdApprovePO.ForeShade = 95
40        Me.cmdReceivePO.ForeShade = 95
50        Me.cmdClosePO.ForeShade = 95

60        If Me.NewRecord Then
              'Nothing more to do. StatusID is null.
70        Else
80            LockControls

90            If Me.StatusID = enumPurchaseOrderStatus.posNew Then GoTo Exit_Handler
100           Me.cmdSubmitPO.ForeShade = 70      '70 means Darker 30%

110           If Me.StatusID = enumPurchaseOrderStatus.posSubmitted Then GoTo Exit_Handler
120           Me.cmdApprovePO.ForeShade = 70

130           If Me.StatusID = enumPurchaseOrderStatus.posApprove Then GoTo Exit_Handler
140           Me.cmdReceivePO.ForeShade = 70

150           If Me.StatusID = enumPurchaseOrderStatus.posReceived Then GoTo Exit_Handler
160           Me.cmdClosePO.ForeShade = 70
170       End If

Exit_Handler:
180       Exit Sub

Err_Handler:
190       clsErrorHandler.HandleError Me.Name, "SetWorkflowLabelColors"
200       Resume Exit_Handler
End Sub

Private Sub VendorID_AfterUpdate()
10        On Error GoTo Err_Handler

20        Me.sfrmPurchaseOrderLineItems.Requery
30        Me.sfrmPurchaseOrderLineItems.Form.ProductID.Requery

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "VendorID_AfterUpdate", True
60        Resume Exit_Handler
End Sub

Private Sub VendorID_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        If IsNull(Me.VendorID.OldValue) Then
              'Changing the value from Null to a number should not trigger the prompt about changing vendors.
30        Else
40            If MsgBox(GetString(sChangingVendor), vbYesNo Or vbQuestion) = vbYes Then
50                g_dbApp().Execute "delete * from PurchaseOrderDetails where PurchaseOrderID = " & Me.PurchaseOrderID, dbFailOnError
60            Else
70                Cancel = True
80            End If
90        End If

Exit_Handler:
100       Exit Sub

Err_Handler:
110       clsErrorHandler.HandleError Me.Name, "VendorID_BeforeUpdate", True
120       Resume Exit_Handler
End Sub
