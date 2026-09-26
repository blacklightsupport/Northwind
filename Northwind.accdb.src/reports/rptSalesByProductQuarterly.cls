Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Activate()
10        On Error GoTo Err_Handler

20        Ribbon_ShowReportsGroup

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Report_Activate", True
50        Resume Exit_Handler
End Sub

Private Sub Report_Close()
10        On Error GoTo Err_Handler

20        Ribbon_ShowReportsGroup

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Report_Close", True
50        Resume Exit_Handler
End Sub

Private Sub Report_NoData(Cancel As Integer)
10        On Error GoTo Err_Handler

20        MsgBox GetString(sReportNoData), vbExclamation
30        Cancel = True

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Report_NoData", True
60        Resume Exit_Handler
End Sub

Private Sub Report_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Me.lblCriteria.Caption = "Order Dates between " & [Forms]![frmReports].[txtStartDate] & " And " & [Forms]![frmReports].[txtEndDate]

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Report_Open", True
50        Resume Exit_Handler
End Sub
