Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAddProduct_Click()
10        On Error GoTo Err_Handler

          Dim strForm             As String

20        strForm = "frmProductDetail"

          'Check if frmProductDetail is already open. If so, navigate it to a new record
          '=================Equivalent Syntax=============================
30        If IsFormOpen(strForm) Then  'This option is a custom function
              '    or
              '        If SysCmd(acSysCmdGetObjectState, acForm, strForm) <> 0 Then
              '    or
              '        If CurrentProject.AllForms(strForm).IsLoaded Then
              '===============================================================

40            Forms(strForm).SetFocus
50            DoCmd.GoToRecord acForm, strForm, acNewRec

60        Else
              '=====================Equivalent Syntax=========================
70            DoCmd.OpenForm strForm, , , , acFormAdd
              '        DoCmd.OpenForm FormName:=strForm, DataMode:=acFormAdd
              '===============================================================
80        End If

Exit_Handler:
90        Exit Sub

Err_Handler:
100       clsErrorHandler.HandleError Me.Name, "cmdAddProduct_Click", True
110       Resume Exit_Handler
End Sub

Private Sub cmdExport_Click()
10        On Error GoTo Err_Handler

          Dim strFilePath         As String

          'set path to the same location as Northwind
20        strFilePath = CurrentProject.Path & "\Northwind Products.xlsx"

30        DoCmd.OutputTo acOutputQuery, "qryProductList_Export", acFormatXLSX, strFilePath

40        Application.FollowHyperlink strFilePath

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "cmdExport_Click", True
70        Resume Exit_Handler
End Sub

Private Sub ProductID_Click()
10        On Error GoTo Err_Handler

          '=====================Equivalent Syntax=========================
20        DoCmd.OpenForm "frmProductDetail", , , "ProductID = " & Me.ProductID
          '    DoCmd.OpenForm FormName:="frmProductDetail", WhereCondition:="ProductID = " & Me.ProductID
          '===============================================================

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "ProductID_Click", True
50        Resume Exit_Handler
End Sub
