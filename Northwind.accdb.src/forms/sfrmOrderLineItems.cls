Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'PURPOSE:
'   Using a user-defined data type to keep the two values together, unlike two separate variables.
'SEE ALSO:
'   https://learn.microsoft.com/office/vba/language/how-to/user-defined-data-type
Private Type typRowCache
    ProductID               As Long
    Quantity                As Integer
End Type

Private m_RowCache          As typRowCache  'Create a variable of the type that was just defined. Values as they were when entering the row.

'NOTE:
'   This combobox is unbound.
Private Sub cboProductCategories_AfterUpdate()
10        On Error GoTo Err_Handler

          'Note that setting the Rowsource also requeries it so it will have products for all categories or for this category.
20        If Me.cboProductCategories = 0 Then
30            Me.ProductID.RowSource = "qrycboProducts_All"
40        Else
50            Me.ProductID.RowSource = "qrycboProducts"
60        End If

          'If IsNull(Me.ProductID) Or Me.ProductID = 0 Then
70        Me.ProductID = Me.ProductID.Column(0, 0)   'Critical line for cascading comboboxes: it selects the first product, which then causes the value of ProductCategory textbox to be set.
80        ProductID_AfterUpdate                      'Since we just set the value programmatically, we need to call this procedure which would be called if it was done interactively.

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError Me.Name, "cboProductCategories_AfterUpdate", True
110       Resume Exit_Handler
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
10        On Error GoTo Err_Handler

20        If Status = acDeleteOK Then
30            AllocateInventory m_RowCache.ProductID      'Use the cached version, because by now the row is deleted and Me.ProductID is undefined.
40        End If

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "Form_AfterDelConfirm", True
70        Resume Exit_Handler
End Sub

Private Sub Form_AfterUpdate()
10        On Error GoTo Err_Handler

20        ReallocateInventory

30        ValidateForm_RemoveHighlights Me

          'Reset.
40        m_RowCache.ProductID = 0
50        m_RowCache.Quantity = 0

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "Form_AfterUpdate", True
80        Resume Exit_Handler
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
10        On Error GoTo Err_Handler

20        Response = acDataErrContinue        'We already asked user for confirmation in Form_Delete.

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_BeforeDelConfirm", True
50        Resume Exit_Handler
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Me.OrderDetailStatusID = enumOrderDetailStatus.odsNew       'Set initial value for new record. May be updated in Form_BeforeUpdate.

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_BeforeInsert", True
50        Resume Exit_Handler
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Cancel = ValidateForm(Me)
30        If Not Cancel Then SetOrderDetailStatus     'Just in case Quantity was entered before Product.

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
60        Resume Exit_Handler
End Sub

Private Sub Form_Current()
10        On Error GoTo Err_Handler

20        Me.cboProductCategories = Me.ProductCategoryID  'Set the value of the unbound combobox.
30        Me.ProductID.Requery                            'Requery products dropdown so it will have products for this category.

40        ValidateForm_RemoveHighlights Me

          'Cache the row, in case user changes product.
50        If Not Me.NewRecord And Me.OrderDetailStatusID = enumOrderDetailStatus.odsAllocated Then
60            m_RowCache.ProductID = Me.ProductID
70            m_RowCache.Quantity = Me.Quantity
80        End If

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError Me.Name, "Form_Current", True
110       Resume Exit_Handler
End Sub

Private Sub Form_Delete(Cancel As Integer)
10        On Error GoTo Err_Handler

20        If MsgBox(GetString(enumStrings.sDeleteRecord, "order line item"), vbYesNo Or vbQuestion) = vbNo Then
30            Cancel = True
40        End If

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "Form_Delete", True
70        Resume Exit_Handler
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
10        On Error GoTo Err_Handler

          'Alternatively we could have:
          '   * Disabled the subform until the parent form is saved, or at least started (Form_BeforeInsert).
          '   * Checked Me.OrderID in our Form_BeforeInsert, so an early message can be given, instead of waiting until user causes the row to (attempt to) be saved.

20        If DataErr = 3314 Then      '3314 = You must enter a value in the '|' field.
30            If IsNull(Me.OrderID) Then
40                MsgBox GetString(sOrderBeforeOrderLineItems), vbExclamation
50                Response = acDataErrContinue
60            End If
70        End If

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "Form_Error", True
100       Resume Exit_Handler
End Sub

Private Sub ProductCategoryName_GotFocus()
10        On Error GoTo Err_Handler

20        Me.ProductCategoryID.SetFocus

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "ProductCategoryName_GotFocus", True
50        Resume Exit_Handler
End Sub

Private Sub ProductID_AfterUpdate()
10        On Error GoTo Err_Handler

          Const COL_UNITPRICE     As Integer = 2

20        If IsNull(Me.ProductID.Column(COL_UNITPRICE)) Then
              'This can happen when changing product category on a new row without entering any other data.
30            Me.ProductID = 0     'Assigning Null is not possible with this required field. 0 is the next best value.
40        Else
50            Me.UnitPrice = Me.ProductID.Column(COL_UNITPRICE)           'Column at index 2 has the hidden UnitPrice.
60        End If

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "ProductID_AfterUpdate", True
90        Resume Exit_Handler
End Sub

Private Sub ProductID_GotFocus()
10        On Error GoTo Err_Handler

          'Note that setting the Rowsource also requeries it so it will have products for all categories or for this category.
20        If Me.cboProductCategories = 0 Then
30            Me.ProductID.RowSource = "qrycboProducts_All"
40        Else
50            Me.ProductID.RowSource = "qrycboProducts"
60        End If

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "ProductID_GotFocus", True
90        Resume Exit_Handler
End Sub

Private Sub ProductName_DblClick(Cancel As Integer)
10        On Error GoTo Err_Handler

20        If Not Me.NewRecord Then
30            DoCmd.OpenForm FormName:="frmProductDetail", WhereCondition:="ProductID = " & Me.ProductID, WindowMode:=acDialog
40        End If

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "ProductName_DblClick", True
70        Resume Exit_Handler
End Sub

Private Sub ProductName_GotFocus()
10        On Error GoTo Err_Handler

20        Me.ProductID.SetFocus           'So user can type and the dropdown will show matching data.

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "ProductName_GotFocus", True
50        Resume Exit_Handler
End Sub

Private Sub ReallocateInventory()
10        On Error GoTo Err_Handler

20        If m_RowCache.ProductID <> Me.ProductID And m_RowCache.ProductID <> 0 Then
              'The product was changed, and the old quantity of it can now be reallocated.
30            AllocateInventory m_RowCache.ProductID
40        End If
          'Current ProductID
50        AllocateInventory Me.ProductID

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "ReallocateInventory"
80        Resume Exit_Handler
End Sub

Private Sub SetOrderDetailStatus()
10        On Error GoTo Err_Handler

          Dim intAvailable        As Integer

20        If IsNull(Me.ProductID) Or IsNull(Me.Quantity) Then
              'Wait until we have both Quantity and ProductID
30        Else
40            intAvailable = ProductAvailable(Me.ProductID)

50            If intAvailable >= Me.Quantity Then
                  'We have stock. Allocate it.
60                Me.OrderDetailStatusID = enumOrderDetailStatus.odsAllocated
70            Else
                  'None or not enough stock. This status is preliminary, and may be modified in Form_AfterUpdate.
80                Me.OrderDetailStatusID = enumOrderDetailStatus.odsNoStock
90            End If
100       End If

Exit_Handler:
110       Exit Sub

Err_Handler:
120       clsErrorHandler.HandleError Me.Name, "SetOrderDetailStatus"
130       Resume Exit_Handler
End Sub
