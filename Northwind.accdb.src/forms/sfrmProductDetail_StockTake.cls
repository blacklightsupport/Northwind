Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'NOTE:
'   A stock take can have been taken yesterday, and today user wants to enter the data.
'   We do not allow it, because product may have been sold or received between the actual date and now. That would affect product allocation.
'   A good problem to address if this were a real inventory program, but it is not. It is a showcase of Access capabilities.
'   So we take the easy way out and do not allow you to enter a stock take datetime.

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

20        RequeryQuantities

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
50        Resume Exit_Handler
End Sub

Private Sub Form_AfterUpdate()
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me

          'Allocate the new Inventory number
30        AllocateInventory Me.ProductID

40        RequeryQuantities

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "Form_AfterUpdate", True
70        Resume Exit_Handler
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Cancel = ValidateForm(Me)

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
50        Resume Exit_Handler
End Sub

'PURPOSE:
'   Prevent existing records from being edited or deleted.
Private Sub Form_Current()
10        On Error GoTo Err_Handler

20        ValidateForm_RemoveHighlights Me

30        Me.AllowAdditions = True
40        Me.AllowDeletions = False
50        Me.AllowEdits = Me.NewRecord

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "Form_Current", True
80        Resume Exit_Handler
End Sub

Private Sub QuantityOnHand_AfterUpdate()
10        On Error GoTo Err_Handler


20        Me.StockTakeDate = Now()
30        Me.ExpectedQuantity = ProductAvailable(Me.ProductID)

40        Me.Dirty = False    'In QuantityOnHand.BeforeUpdate, the user has confrimed the value for this one and only editable field is correct - force the record to save

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "QuantityOnHand_AfterUpdate", True
70        Resume Exit_Handler
End Sub

Private Sub QuantityOnHand_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

          Dim intOnHand           As Integer
          Dim strMsg              As String

20        intOnHand = ProductAvailable(Me.ProductID)
30        If intOnHand = Me.QuantityOnHand Then
40        Else
50            strMsg = StringFormat("You entered {0} and the expected quantity is {1}. Is {0} correct?", Me.QuantityOnHand, intOnHand)
60            If MsgBox(strMsg, vbYesNo Or vbQuestion, "Quantity On Hand Does Not Match Expected") = vbNo Then
70                Me.Undo
80                Cancel = True
90            End If
100       End If

Exit_Handler:
110       Exit Sub

Err_Handler:
120       clsErrorHandler.HandleError Me.Name, "QuantityOnHand_BeforeUpdate", True
130       Resume Exit_Handler
End Sub

Private Sub RequeryQuantities()
10        On Error GoTo Err_Handler

20        If IsFormOpen("frmProductDetail") Then
              'Example of calling a public sub on a parent form, from a subform
              'Recalculates quantities on the Product Detail form (parent); sets control properties accordingly
30            Form_frmProductDetail.ControlStates

              'Recalculates quantities on the Products List if it's loaded
40            RequeryProductList

50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "RequeryQuantities"
80        Resume Exit_Handler
End Sub
