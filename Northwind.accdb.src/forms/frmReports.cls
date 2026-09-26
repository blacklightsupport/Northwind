Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdEmployeeEmailList_Click()
10        On Error GoTo Err_Handler

20        Call ReportFormInitialization(False)
30        TempVars.Add "ReportName", "rptEmployeeEmailList"

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "cmdEmployeeEmailList_Click", True
60        Resume Exit_Handler
End Sub

Private Sub cmdEmployeePhoneList_Click()
10        On Error GoTo Err_Handler

20        Call ReportFormInitialization(False)
30        TempVars.Add "ReportName", "rptEmployeePhoneList"

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "cmdEmployeePhoneList_Click", True
60        Resume Exit_Handler
End Sub

'NOTE:
'   The hardest part of this report is generating the table of contents (ToC). Page numbers will change as more data is added.
'   This requires the report to be rendered fully.
'   Sending the report to a PDF will update the table of contents.
'   The print preview includes an OpenArgs which is used to turn off the deletion of the table contents records
Private Sub cmdFallCatalog_Click()
10        On Error GoTo Err_Handler

20        Call ReportFormInitialization(False)
30        TempVars.Add "ReportName", "rptProductCatalog"
40        TempVars.Add "ReportView", acViewPreview

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "cmdFallCatalog_Click", True
70        Resume Exit_Handler
End Sub

Private Sub cmdPreviewReport_Click()
10        On Error GoTo Err_Handler

          Dim strPath             As String                'this is the directory of the application to send the pdf

20        If TempVars!ReportName = "rptProductCatalog" Then
              'Special processing for this report.
              'get the path of the current application and construct the path to the PDF file.
30            strPath = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\"))
40            strPath = strPath & TempVars!ReportName & ".pdf"    'this is the name of the sacrificial PDF that will be deleted
50            If FileExists(strPath) Then
60                Kill strPath
70            End If

              'Run the report for the first time, to create the ToC.
80            TempVars.Add "GenerateToC", True
90            DoCmd.OutputTo ObjectType:=acOutputReport, ObjectName:=TempVars!ReportName, OutputFormat:=acFormatPDF, OutputFile:=strPath
100           TempVars!GenerateToC = False
110       End If

120       If IsNull(TempVars!ReportView) Then TempVars!ReportView = acViewPreview
130       DoCmd.OpenReport TempVars!ReportName, TempVars!ReportView, "", "", acNormal

Exit_Handler:
140       Exit Sub

Err_Handler:
150       clsErrorHandler.HandleError Me.Name, "cmdPreviewReport_Click", True
160       Resume Exit_Handler
End Sub

Private Sub cmdSalesByemployeeReport_Click()
10        On Error GoTo Err_Handler

20        Call ReportFormInitialization(True)
30        TempVars.Add "ReportName", "rptSalesByEmployee"
40        TempVars.Add "ReportView", acViewReport             'This report must be opened in Report View, so user can filter it.

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "cmdSalesByemployeeReport_Click", True
70        Resume Exit_Handler
End Sub

Private Sub cmdSalesByProductReportMonthly_Click()
10        On Error GoTo Err_Handler

20        Call ReportFormInitialization(True)
30        TempVars.Add "ReportName", "rptSalesByProduct"

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "cmdSalesByProductReportMonthly_Click", True
60        Resume Exit_Handler
End Sub

Private Sub cmdSalesByProductReportQuarterly_Click()
10        On Error GoTo Err_Handler

20        Call ReportFormInitialization(True)
30        TempVars.Add "ReportName", "rptSalesByProductQuarterly"

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "cmdSalesByProductReportQuarterly_Click", True
60        Resume Exit_Handler
End Sub

Private Sub Form_Close()
10        On Error GoTo Err_Handler

20        TempVars.Remove "ReportView"
30        TempVars.Remove "ReportName"

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_Close", True
60        Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

          ' Set default report name to the first report.
20        Me.SetFocus                                         'Prepare for Screen.ActiveControl later on in ReportFormInitialization.
30        Me.cmdSalesByEmployeeReport.SetFocus                'Prepare for Screen.ActiveControl later on in ReportFormInitialization.
40        cmdSalesByemployeeReport_Click                      'Init this form by doing the same thing as if user clicked the first report button.

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "Form_Open", True
70        Resume Exit_Handler
End Sub

'PURPOSE:
'   If blnEnableCriteria is true, show the date range criteria controls and set default date values.
'NOTE:
'   If you add other criteria controls, you must add to this code.
Private Sub ReportFormInitialization(ByVal blnEnableCriteria As Boolean)
10        On Error GoTo Err_Handler

          Dim dtMaxOrderDate      As Date

20        dtMaxOrderDate = DateValue(DMax("OrderDate", "Orders"))       'Set the date range relative to the max order date.
30        With Me
40            .txtStartDate.Enabled = blnEnableCriteria
50            .txtStartDate.Visible = blnEnableCriteria
60            .txtStartDate = IIf(blnEnableCriteria = True, DateAdd("m", -3, DateValue(dtMaxOrderDate)), Null)
70            .txtEndDate.Enabled = blnEnableCriteria
80            .txtEndDate.Visible = blnEnableCriteria
90            .txtEndDate = IIf(blnEnableCriteria = True, dtMaxOrderDate, Null)
100           .lblSelectedReportName.Caption = Screen.ActiveControl.Caption       'Will throw Error 2474 at startup time, when there is no focus yet.
110       End With

Exit_Handler:
120       Exit Sub

Err_Handler:
130       clsErrorHandler.HandleError Me.Name, "ReportFormInitialization"
140       Resume Exit_Handler
End Sub
