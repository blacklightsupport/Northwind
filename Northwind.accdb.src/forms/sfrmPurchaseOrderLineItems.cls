Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Cancel = ValidateForm(Me)       'Validates required fields.

30        If Not Cancel Then
              'Additional validation regarding desired inventory levels.
40            Cancel = ValidateLineItem()
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
80        Resume Exit_Handler
End Sub

Private Sub Form_Current()
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Current", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
10        On Error GoTo Err_Handler

20        If DataErr = 3314 Then      '3314 = You must enter a value in the '|' field.
30            If IsNull(Me.PurchaseOrderID) Then
40                MsgBox "You must first fill out the main Purchase Order form before entering line items. Please cancel this record by hitting Esc, and fill out the main form.", vbExclamation
50                Response = acDataErrContinue
60            End If
70        End If

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "Form_Error", True
100       Resume Exit_Handler
End Sub

Private Sub ProductID_AfterUpdate()
10        On Error GoTo Err_Handler

20        Me.Quantity = ProductReorderQuantity(Me.ProductID)
30        Me.UnitCost = Me.ProductID.Column(2)          'Column at index 2 has the hidden Standard Cost.

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "ProductID_AfterUpdate", True
60        Resume Exit_Handler
End Sub

Private Sub PurchaseOrderDetailID_Click()
10        On Error GoTo Err_Handler

20        If Not Me.NewRecord Then
30            DoCmd.OpenForm "frmProductDetail", WhereCondition:="ProductID = " & Me.ProductID
40        End If

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "PurchaseOrderDetailID_Click", True
70        Resume Exit_Handler
End Sub

'RETURNS:
'   False if the line item is valid = conforms to the business rules; True otherwise.
Private Function ValidateLineItem() As Boolean
10        On Error GoTo Err_Handler

          Dim rs                  As DAO.Recordset
          Dim blnReturn           As Boolean
          Dim intAvailable        As Integer
          Dim sql                 As String

20        blnReturn = False

30        sql = "select * from Products where ProductID = " & Me.ProductID
40        Set rs = g_dbApp().OpenRecordset(sql, dbOpenSnapshot)

          'BUSINESS RULE: Qty should be at least MinimumReorderQuantity
50        If Me.Quantity < Nz(rs!MinimumReorderQuantity, 0) Then
60            blnReturn = True
70            MsgBox GetString(sMinimumReorderQuantity, rs!MinimumReorderQuantity), vbExclamation
80        End If

          'BUSINESS RULE: Qty should be sufficient to get to at least TargetLevel.
90        intAvailable = ProductToSell(Me.ProductID)
100       If intAvailable + Me.Quantity < rs!TargetLevel Then
              'This is a warning, not an error, so we are not setting blnReturn = True.
110           MsgBox GetString(sLessThanTargetLevel, rs!TargetLevel, rs!TargetLevel - intAvailable - Me.Quantity), vbExclamation
120       End If

130       rs.Close
140       Set rs = Nothing

150       ValidateLineItem = blnReturn

Exit_Handler:
160       Exit Function

Err_Handler:
170       clsErrorHandler.HandleError Me.Name, "ValidateLineItem"
180       Resume Exit_Handler
End Function
