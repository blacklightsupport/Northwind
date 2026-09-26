Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAdd_Click()
10        On Error GoTo Err_Handler

20        DoCmd.GoToRecord , "", acNewRec

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdAdd_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdDelete_Click()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            Me.Undo
40            GoTo Exit_Handler
50        End If

          'Check if related records will prevent deleting this Product
60        If DCount("ProductID", "Products", "ProductCategoryID = " & Me.ProductCategoryID) > 0 Then

              '================Equivalent Syntax==============================
              '   Two methods of generating a message box
              '   The first is default: Building the messagebox message
              '   The second uses the custom Function GetString in conjunction
              '       with the Enumeration enumStrings.
              'Method 1
              '        MsgBox "Product Category '" & Me.ProductCategory Name & "' cannot be deleted because it has a related " _
                       '        & "Product record.", vbExclamation or vbOKOnly, "Cannot Delete"

              'Method 2
70            MsgBox GetString(enumStrings.sCannotDelete, "Product Category", Me.ProductCategoryName _
                                                                              , "Product"), vbExclamation Or vbOKOnly, "Cannot Delete"
              '===============================================================

80            GoTo Exit_Handler

90        End If

100       DoCmd.RunCommand acCmdDeleteRecord
110       Me.lstCategory.Requery

Exit_Handler:
120       Exit Sub

Err_Handler:
130       clsErrorHandler.HandleError Me.Name, "cmdDelete_Click", True
140       Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

20        Me.lstCategory.Requery

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Current()
10        On Error GoTo Err_Handler

          Dim lngCategoryID       As Long

20        lngCategoryID = Nz(Me.txtProductCategoryID, 0)

          'If user has changed the record using the navigation controls at the bottom of the form,
          'we want to ensure that a different category is not still displaying as selected in the listbox lstCategory
30        If lngCategoryID <> CLng(Nz(Me.lstCategory, lngCategoryID)) Then
40            Me.lstCategory = Null
              '        Me.ctlProductCategory.SetFocus
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "Form_Current", True
80        Resume Exit_Handler
End Sub

Private Sub Form_Delete(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Me.lstCategory.Requery

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_Delete", True
50        Resume Exit_Handler
End Sub

Private Sub lstCategory_AfterUpdate()
10        On Error GoTo Err_Handler

20        If Nz(Me.lstCategory, 0) > 0 Then

              'temporarily disable screen refresh to avoid flicker in title
30            Me.Painting = False

              '=====================Equivalent Syntax=========================
40            DoCmd.SearchForRecord , "", acFirst, "ProductCategoryID = " & Me.lstCategory

              '        With Me.RecordsetClone
              '            .FindFirst "ProductCategoryID = " & Me.lstCategory
              '            If Not .NoMatch Then Me.Bookmark = .Bookmark
              '        End With
              '===============================================================

50        End If

Exit_Handler:
          're-enable screen refresh
60        If Me.Painting = False Then Me.Painting = True
70        Exit Sub

Err_Handler:
          're-enable screen refresh
80        If Me.Painting = False Then Me.Painting = True
90        clsErrorHandler.HandleError Me.Name, "lstCategory_AfterUpdate"
100       Resume Exit_Handler
End Sub

Private Sub txtProductCategoryCode_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then GoTo Exit_Handler

          'check if this product category has been used in a product. If so, don't allow change,
          'since the Products.ProductCode is calculated based on ProductCategories.txtProductCategoryCode
30        If DCount("ProductID", "Products", "ProductCategoryID = " & Me.txtProductCategoryID) > 0 Then

40            MsgBox "You cannot change or delete this Product Code because it is in use by one or more Products" _
                     , vbCritical + vbOKOnly, "Protected Code"
50            Cancel = True
60            Me.txtProductCategoryCode.Undo
70        End If

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "txtProductCategoryCode_BeforeUpdate", True
100       Resume Exit_Handler
End Sub
