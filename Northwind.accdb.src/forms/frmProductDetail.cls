Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'PURPOSE:
'   Sets calculated inventory fields
'   We'll call this each time we change the record
Public Sub CalculateLevels()
10        On Error GoTo Err_Handler

          Dim dtLastStockTake     As Date

20        If Me.NewRecord Then GoTo Exit_Handler

30        Me.txtQtyAvailable = ProductToSell(Me.ProductID)

40        Me.txtQtyNoStock = ProductNoStock(Me.ProductID)

50        dtLastStockTake = ProductLastStockTakeDate(Me.ProductID)

60        Me.txtQtyAllocated = ProductAllocated(Me.ProductID)

70        Me.txtQtyOnOrder = ProductOnOrder(Me.ProductID)

80        Me.txtQtyToReorder = ProductReorderQuantity(Me.ProductID)

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError Me.Name, "CalculateLevels"
110       Resume Exit_Handler
End Sub

Private Sub cboFindProduct_AfterUpdate()
10        On Error GoTo Err_Handler

          Dim lngProductID        As Long

20        lngProductID = Me.cboFindProduct

          'Remove the filter because the recordset is searched as filtered
30        If Me.FilterOn Then Me.FilterOn = False
40        Me.DataEntry = False

50        DoCmd.SearchForRecord acDataForm, Me.Name, acFirst, "[ProductID] = " & lngProductID

          '============UI Option==============
          'Reset Product selection combo
          'Comment if the selected product should remain visible in the combo
60        Me.cboFindProduct = ""
          '===================================

70        ControlStates

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "cboFindProduct_AfterUpdate", True
100       Resume Exit_Handler
End Sub

Private Sub cboProductCategoryID_AfterUpdate()
10        On Error GoTo Err_Handler

20        SetProductCode

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cboProductCategoryID_AfterUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub chkDiscontinued_AfterUpdate()
10        On Error GoTo Err_Handler

20        ControlStates

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "chkDiscontinued_AfterUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub cmdAddProduct_Click()
10        On Error GoTo Err_Handler

20        DoCmd.GoToRecord , "", acNewRec

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdAddProduct_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdDeleteProduct_Click()
10        On Error GoTo Err_Handler

          Dim strMessage          As String

20        strMessage = ""

          'Check if related records will prevent deleting this Product
30        Select Case True
              Case DCount("OrderID", "OrderDetails", "ProductID = " & Me.ProductID) > 0
40                strMessage = "Order"
50            Case DCount("PurchaseOrderDetailID", "PurchaseOrderDetails", "ProductID = " & Me.ProductID) > 0
60                strMessage = "Purchase Order"
70        End Select

80        If strMessage > "" Then

              '================Equivalent Syntax==============================
              '   Two methods of generating a message box
              '   The first is default: Building the messagebox message
              '   The second uses the custom Function GetString in conjunction
              '       with the Enumeration enumStrings.


              'Method 1
              '        MsgBox "Product '" & Me.ProductName & "' cannot be deleted because it has a related " _
                       '        & strMessage & " record.", vbExclamation or vbOKOnly, "Cannot Delete"

              'Method 2
90            MsgBox GetString(enumStrings.sCannotDelete, "Product", Me.ProductName, strMessage) _
                     , vbExclamation Or vbOKOnly, "Cannot Delete"
              '===============================================================

100           GoTo Exit_Handler
110       End If

          'Delete child records in Companies (Vendors) and StockTake (Inventory) first
          '====================Equivalent Syntax==========================
120       g_dbApp().Execute "DELETE * FROM StockTake WHERE ProductID = " & Me.ProductID & ";", dbFailOnError
          '   DoCmd.RunSQL "DELETE * FROM StockTake WHERE ProductID = " & Me.ProductID & ";"
          '   ================================================================

          '================Equivalent Syntax==============================
130       g_dbApp().Execute "DELETE * FROM ProductVendors WHERE ProductID = " & Me.ProductID & ";", dbFailOnError
          '   DoCmd.RunSQL "DELETE * FROM ProductVendors WHERE ProductID = " & Me.ProductID & ";"
          '===============================================================

140       If Not Form.NewRecord Then
150           Me.Refresh
160           DoCmd.RunCommand acCmdDeleteRecord
170           DoCmd.RunCommand acCmdRemoveFilterSort
180           Me.cboFindProduct.Requery
190           RequeryProductList
200       Else
210           Me.Undo
220       End If

Exit_Handler:
230       Exit Sub

Err_Handler:
240       clsErrorHandler.HandleError Me.Name, "cmdDeleteProduct_Click", True
250       Resume Exit_Handler
End Sub

Private Sub cmdReorderProduct_Click()
10        On Error GoTo Err_Handler

          Dim varRelatedID        As Variant
          Dim lngTargetPOID       As Long
          Dim lngVendorID         As Long
          Dim aryVendor()         As String
          Dim strInputBox         As String
          Dim strWhere            As String
          Dim intOrderQuantity    As Integer
          Dim lngTargetLineID     As Long

20        If Me.NewRecord And Me.Dirty = False Then GoTo Exit_Handler

30        If Me.Dirty Then
              '====================Equivalent Syntax==========================
40            Me.Dirty = False
              '        DoCmd.RunCommand acCmdSaveRecord
              '===============================================================
50        End If

          'Check if an Order for this Product is outstanding.
60        If Me.txtQtyOnOrder <> 0 Then

70            strWhere = StringFormatSQL("ProductID = {0} AND StatusID in ({1},{2},{3})", Me.ProductID _
                                                                                          , enumPurchaseOrderStatus.posApprove _
                                                                                            , enumPurchaseOrderStatus.posNew _
                                                                                              , enumPurchaseOrderStatus.posSubmitted)

80            varRelatedID = DLookup("PurchaseOrderID", "qryProductPurchaseOrder", strWhere)

90            If Nz(varRelatedID, 0) > 0 Then

                  'Open frmPurchaseOrders and go to the outstanding Order
100               OpenPurchaseOrderDetailsForm varRelatedID

110               GoTo Exit_Handler
120           End If
130       End If

          'Check if this product has at least one vendor defined before allowing a purchase order.
          'The Purchase Order Line items subform requires a pre-defined vendor for each product
140       varRelatedID = DLookup("Count(ProductVendorID) & ""|"" & First(VendorID)", "ProductVendors", "ProductID = " & Me.ProductID)

150       If Nz(varRelatedID, 0) = "0|" Then

              'Prompt user to enter a Vendor
160           MsgBox "At least one product vendor must be defined for this product before a purchase order can be created.", _
                     vbExclamation Or vbOKOnly, "Vendor Required"

              'Set focus on the Product Vendors tab
170           Me.tabProducts.Pages.Item("pgVendors").SetFocus
180           GoTo Exit_Handler

190       Else
200           aryVendor = Split(varRelatedID, "|")
210       End If

220       If aryVendor(0) = 1 Then     'If we counted a single vendor, we don't have to prompt for a vendor choice
230           lngVendorID = aryVendor(1)
240       Else

              'Prompt user to choose a vendor from multiple possible vendors
              '====================Equivalent Syntax=========================
250           DoCmd.OpenForm "frmSelectVendorDialog", , , , , acDialog, "Vendor Required|Select a Vendor|" & Me.ProductID
              '        DoCmd.OpenForm FormName:="frmSelectVendorDialog", WindowMode:=acDialog, OpenArgs:="Vendor Required|Select a Vendor|" & Me.ProductID
              '===============================================================

260           If IsFormOpen("frmSelectVendorDialog") Then
270               lngVendorID = Form_frmSelectVendorDialog.lstVendor
                  '====================Equivalent Syntax=========================
280               DoCmd.Close acForm, "frmSelectVendorDialog"   'Note the explicit naming of the form to be closed. Without it, the active form is closed which may not always be what you want.
                  '            DoCmd.Close ObjectType:=acForm, ObjectName:="frmSelectVendorDialog"
                  '===============================================================
290           Else
                  'User cancelled.
300               GoTo Exit_Handler
310           End If
320       End If

          'Reorder Quantity: txtQtyReorder > txtMinimumOrderQuantity > prompt user for quantity
330       If Me.txtQtyToReorder = 0 Then
340           intOrderQuantity = Nz(Me.txtMinimumReorderQuantity, 0)
350           If intOrderQuantity = 0 Then
360               strInputBox = InputBox("Specify order quantity:", "Quantity", 0)
370               If strInputBox = "" Then GoTo Exit_Handler      'User clicked Cancel.
380               intOrderQuantity = CInt(strInputBox)
390           End If

400           If intOrderQuantity = 0 Then GoTo Exit_Handler
410       Else
420           intOrderQuantity = Me.txtQtyToReorder
430       End If

          'Check if there's a Purchase Order in "New" status for the Vendor. If so, we'll append to that. DMin will get the oldest PO.
440       varRelatedID = DMin("PurchaseOrderID", "PurchaseOrders" _
                                                 , StringFormatSQL("VendorID = {0} AND StatusID = {1}" _
                                                                   , lngVendorID, enumPurchaseOrderStatus.posNew))

450       If Nz(varRelatedID, 0) > 0 Then
              'Append a new line item to the existing purchase order
460           lngTargetPOID = varRelatedID
470           lngTargetLineID = AddPurchaseOrderDetail(lngTargetPOID, Me.ProductID, intOrderQuantity, Me.StandardUnitCost)
480       Else
              'Create the new Purchase Order
490           lngTargetPOID = ReorderProduct(Me.ProductID, lngVendorID, intOrderQuantity, Me.StandardUnitCost)
500       End If

          'If no new Purchase Order, fail the procedure and exit
510       If lngTargetPOID = 0 Then
520           MsgBox "Reorder Failed.", vbExclamation Or vbOKOnly, "Reorder Fail"
530           GoTo Exit_Handler
540       End If

          'Requery the Quantities and set control formats accordingly
550       ControlStates

          'Set focus on the Purchase Orders tab and requery its subform to reflect the new Purchase Order
560       Me.tabProducts.Pages.Item("pgPurchaseOrders").SetFocus
570       Me.ctlPOSubform.Form.Requery

580       OpenPurchaseOrderDetailsForm lngTargetPOID

590       AddToMRU "PurchaseOrders", lngTargetPOID

          'Update PO list form if loaded
          '=================Equivalent Syntax=============================
600       If IsFormOpen("frmPurchaseOrderList") Then  'This option is a custom function
              '    or
              '    If SysCmd(acSysCmdGetObjectState, acForm, "frmPurchaseOrderList") <> 0 Then
              '    or
              '    If CurrentProject.AllForms("frmPurchaseOrderList").IsLoaded Then
              '===============================================================
610           Forms!frmPurchaseOrderList.Requery
620       End If

          'If the Product List is loaded, requery to reflect the "on order" quantity.
630       RequeryProductList

Exit_Handler:
640       Exit Sub

Err_Handler:
650       Select Case Err.Number
              Case 2186, 91, 13           '2186=error on switching to design view; 91, 13=User cleared input box value or cancelled it
660               Resume Exit_Handler
670           Case Else
680               clsErrorHandler.HandleError Me.Name, "cmdReorderProduct_Click"
690               Resume Exit_Handler
700       End Select
710       Resume
End Sub

'PURPOSE:
'   Sets UI states
Public Sub ControlStates()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            Me.cmdDeleteProduct.Visible = False
40            Me.cmdReorderProduct.Visible = False
50            Me.Caption = "New Product"

              'Ensure focus is not on pgOrders or PgPurchaseOrders, before making them not visible
60            Me.txtProductName.SetFocus
70            Me.pgOrders.Visible = False
80            Me.PgPurchaseOrders.Visible = False

              '===============================================================
              'resolves "ghost" tab display of Orders & Purchase Orders
90            Me.txtProductName.SetFocus
100           Me.tabProducts.Visible = False
110           Me.tabProducts.Visible = True

              '===============================================================

              'Zero out the calculated fields
120           Me.txtQtyAllocated = 0
130           Me.txtQtyAvailable = 0
140           Me.txtQtyToReorder = 0
150           Me.txtQtyOnOrder = 0
160           Me.txtQtyToReorder.ForeColor = Me.txtQtyAvailable.ForeColor
170           Me.lblQtyToReorder.ForeColor = Me.lblQtyAvailable.ForeColor
180           Me.txtQtyToReorder.FontBold = False
190           Me.lblQtyToReorder.FontBold = False
200       Else
              'Populate the inventory level text boxes
210           CalculateLevels

220           Me.cmdDeleteProduct.Visible = True
230           Me.cmdReorderProduct.Visible = True
240           Me.Caption = "Product: " & Me.ProductName
250           Me.pgOrders.Visible = True
260           Me.PgPurchaseOrders.Visible = True

              'Set formatting of inventory level text boxes
270           If Me.txtQtyToReorder > 0 And Me.chkDiscontinued = False Then
280               Me.txtQtyToReorder.ForeColor = vbRed
290               Me.lblQtyToReorder.ForeColor = vbRed
300               Me.txtQtyToReorder.FontBold = True
310               Me.lblQtyToReorder.FontBold = True
320           Else
330               Me.txtQtyToReorder.ForeColor = Me.txtQtyAvailable.ForeColor
340               Me.lblQtyToReorder.ForeColor = Me.lblQtyAvailable.ForeColor
350               Me.txtQtyToReorder.FontBold = False
360               Me.lblQtyToReorder.FontBold = False
370           End If
380       End If

          'Display the Discontinued checbox's label in red if the product is discontinued
390       If Me.chkDiscontinued = True Then
400           Me.cmdReorderProduct.Visible = False
410           Me.lblDiscontinued.BackColor = vbRed
420           Me.lblDiscontinued.ForeColor = vbWhite
430       Else
440           Me.cmdReorderProduct.Visible = Not Me.NewRecord
450           Me.lblDiscontinued.BackColor = vbWhite
460           Me.lblDiscontinued.ForeColor = RGB(89, 89, 89)
470       End If

480       ValidateForm_RemoveHighlights Me        'Needed if user canceled a new record with yellow highlights on the form.

Exit_Handler:
490       Exit Sub

Err_Handler:
500       clsErrorHandler.HandleError Me.Name, "ControlStates"
510       Resume Exit_Handler
End Sub

Private Sub Form_Activate()
10        On Error GoTo Err_Handler

20        If Not Me.NewRecord Then
              'refresh counts that might have been altered on another form
30            ControlStates
40            Me.ctlOrderSubform.Requery
50            Me.ctlPOSubform.Requery
60        End If

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "Form_Activate", True
90        Resume Exit_Handler
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
10        On Error GoTo Err_Handler

20        Me.cboFindProduct.Requery

          'If the Products list is loaded, update it
30        RequeryProductList

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_AfterDelConfirm", True
60        Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

20        Me.cboFindProduct.Requery

          'Set control formatting and calculated quantities
30        ControlStates

          'If the Products list is loaded, update it so the new record will show.
40        RequeryProductList

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
70        Resume Exit_Handler
End Sub

Private Sub Form_AfterUpdate()
10        On Error GoTo Err_Handler

20        RequeryProductList
30        ControlStates

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_AfterUpdate", True
60        Resume Exit_Handler
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Cancel = ValidateForm(Me)

30        If Me.txtStandardUnitCost <= 0 Then
40            MsgBox "Standard Unit Cost must be > 0", vbOKOnly Or vbExclamation, "Cost Required"
50            Cancel = True
60            Me.txtStandardUnitCost.SetFocus
70        End If

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
100       Resume Exit_Handler
End Sub

Private Sub Form_Close()
10        On Error GoTo Err_Handler

          'clear the tab subform source objects to enable lazy loading
20        Me.ctlOrderSubform.SourceObject = ""
30        Me.ctlInventorySubform.SourceObject = ""
40        Me.ctlVendorSubform.SourceObject = ""
50        Me.ctlPOSubform.SourceObject = ""

          'Refresh any open list forms
60        RequeryListForms

Exit_Handler:
70        Exit Sub

Err_Handler:
80        If Err.Number = 2467 Then    '2467 = Error switching to design view
              'do nothing
90        Else
100           clsErrorHandler.HandleError Me.Name, "Form_Close"
110       End If
120       Resume Exit_Handler
End Sub

Private Sub Form_Current()
10        On Error GoTo Err_Handler

20        ControlStates

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Current", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Delete(Cancel As Integer)
10        On Error GoTo Err_Handler

20        RemoveFromMRU "Products", Me.ProductID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Delete", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Load()
10        On Error GoTo Err_Handler

          'Forces the subform to load in the first visible tab
          'This allows the tab order to be changed without causing errors due to late binding.
20        tabProducts_Change

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Load", True
50        Resume Exit_Handler
End Sub

Private Sub imgHelp_Click()
10        On Error GoTo Err_Handler

          '    Application.FollowHyperlink ("<>")


Exit_Handler:
20        Exit Sub

Err_Handler:
30        clsErrorHandler.HandleError Me.Name, "imgHelp_Click", True
40        Resume Exit_Handler
End Sub

Private Sub SetProductCode()
10        On Error GoTo Err_Handler

          Dim varPrefix           As Variant

20        If Nz(Me.ProductID, 0) = 0 Then GoTo Exit_Handler

30        If Nz(Me.cboProductCategoryID, 0) > 0 Then
40            varPrefix = DLookup("ProductCategoryCode", "ProductCategories", "ProductCategoryID = " & Me.cboProductCategoryID)

50            Me.txtProductCode = "NWT" & varPrefix & "-" & Me.ProductID
60        End If

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "SetProductCode"
90        Resume Exit_Handler
End Sub

Private Sub tabProducts_Change()
10        On Error GoTo Err_Handler

          'Late Binding:
          'Setting the source object for the subform only when the tab is selected by the user
          'can improve form performance. This is called Late Binding.

20        Select Case tabProducts.Pages.Item(tabProducts.Value).Name

              Case "pgInventory"
30                If Me.ctlInventorySubform.SourceObject = "" Then Me.ctlInventorySubform.SourceObject = "sfrmProductDetail_StockTake"

40            Case "PgOrders"
50                If Me.ctlOrderSubform.SourceObject = "" Then Me.ctlOrderSubform.SourceObject = "sfrmProductDetail_Orders"

60            Case "PgPurchaseOrders"
70                If Me.ctlPOSubform.SourceObject = "" Then Me.ctlPOSubform.SourceObject = "sfrmProductDetail_PurchaseOrders"

80            Case "pgVendors"
90                If Me.ctlVendorSubform.SourceObject = "" Then Me.ctlVendorSubform.SourceObject = "sfrmProductDetail_Vendors"

100       End Select

Exit_Handler:
110       Exit Sub

Err_Handler:
120       clsErrorHandler.HandleError Me.Name, "tabProducts_Change", True
130       Resume Exit_Handler
End Sub

Private Sub txtTargetLevel_AfterUpdate()
10        On Error GoTo Err_Handler

20        If Not Me.NewRecord Then

              'Save the record
              '====================Equivalent Syntax==========================
30            Me.Dirty = False
              '        DoCmd.RunCommand acCmdSaveRecord   'note that this (docmd) syntax will err in break (debug) mode
              '===============================================================

              'Re-calculate inventory levels and adjust control formatting accordingly
40            ControlStates
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "txtTargetLevel_AfterUpdate", True
80        Resume Exit_Handler
End Sub
