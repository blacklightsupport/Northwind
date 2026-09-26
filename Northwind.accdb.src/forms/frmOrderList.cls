Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'NOTE:
'   Order Date column is bound to OrderDateOnly (an extra column in qryOrderList), not OrderDate, so it is only a date, not a datetime,
'   facilitating filtering for all orders for that date.

Private Sub cmdAddOrder_Click()
10        On Error GoTo Err_Handler

20        OpenOrderDetailsForm

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdAddOrder_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdShowFilter_Click()
10        On Error GoTo Err_Handler

          Dim strOpenArgs         As String

20        If Me.Form.FilterOn = False Then
30            MsgBox "This form has not been filtered.", vbOKOnly Or vbInformation
40        Else
              'OpenArgs uses Replace function to make the filter more readable for humans.
50            strOpenArgs = StringFormat("Header={0}&Message={1}", "Form Filter", Replace(Me.Form.Filter, "[qryOrderList].", ""))
60            DoCmd.OpenForm "frmGenericDialog", acNormal, , , acFormReadOnly, acDialog, strOpenArgs
70            If IsFormOpen("frmGenericDialog") Then
80                DoCmd.Close acForm, "frmGenericDialog", acSaveNo
90            End If
100       End If

Exit_Handler:
110       Exit Sub

Err_Handler:
120       clsErrorHandler.HandleError Me.Name, "cmdShowFilter_Click", True
130       Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

          'Clean up anything left by a previous session

20        Me.FilterOn = False
30        Me.Filter = ""

          'Apply the filter for the default value of the fraFilter group box.
40        fraFilter_AfterUpdate

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "Form_Open", True
70        Resume Exit_Handler
End Sub

Private Sub fraFilter_AfterUpdate()
10        On Error GoTo Err_Handler

20        Select Case Me.fraFilter
              Case 1                          '1=Open Orders
30                Me.Filter = "OrderStatusID <> 1"    'Open orders are all except in Status=1=Closed.

40            Case 2                          '2=Recent Orders
50                Me.Filter = StringFormatSQL("OrderDate >= {0}", DateAdd("d", -30, Date))

60            Case 3                          '3=My Orders
70                Me.Filter = "EmployeeID = " & Get_UserID()

80            Case 4                          '4=All Orders
90                Me.Filter = ""
100       End Select
110       Me.FilterOn = True                  'This turns the filter on.

120       If Me.fraFilter = 2 And Me.RecordsetClone.RecordCount = 0 Then
130           MsgBox GetString(sNoRecentOrders), vbInformation
140       End If

Exit_Handler:
150       Exit Sub

Err_Handler:
160       clsErrorHandler.HandleError Me.Name, "fraFilter_AfterUpdate", True
170       Resume Exit_Handler
180       Resume
End Sub


Private Sub OrderID_Click()
10        On Error GoTo Err_Handler

20        OpenOrderDetailsForm Me.OrderID

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "OrderID_Click", True
50        Resume Exit_Handler
End Sub
