Attribute VB_Name = "modReportParameters"
Option Compare Database
Option Explicit

'Functions used to supply the report queries with parameter values. These work (and return default values) even if frmReports is not open.
'Alternatively we could have used TempVars.

Public Function reportParameterStartDate() As Date
10        On Error GoTo Err_Handler

          Dim datStart            As Date

20        If IsFormOpen("frmReports") Then
30            datStart = Forms("frmReports").txtStartDate
40        Else
50            datStart = DateAdd("m", -3, DateValue(DMax("OrderDate", "Orders")))       'Set the date range relative to the max order date.
60        End If

70        reportParameterStartDate = datStart

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError "modReportParameters", "reportParameterStartDate"
100       Resume Exit_Handler
End Function

Public Function reportParameterEndDate() As Date
10        On Error GoTo Err_Handler

          Dim datEnd              As Date

20        If IsFormOpen("frmReports") Then
30            datEnd = Forms("frmReports").txtEndDate
40        Else
50            datEnd = DateValue(DMax("OrderDate", "Orders"))       'Set the date range relative to the max order date.
60        End If

70        reportParameterEndDate = datEnd

Exit_Handler:
80        Exit Function

Err_Handler:
90        clsErrorHandler.HandleError "modReportParameters", "reportParameterEndDate"
100       Resume Exit_Handler
End Function
