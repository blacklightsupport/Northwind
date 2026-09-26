Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub GroupHeader0_Print(Cancel As Integer, PrintCount As Integer)
10        On Error GoTo Err_Handler

          Dim strTitle            As String           'The TOC title
          Dim intPageNumber       As Integer          'Page number grabed from the report
          Dim strSQL              As String           'used to build SQL Statements to execute

          'If TempVars!GenerateToC is not initialized (likely if report is not run from frmReports), then initialize it now.
20        If IsNull(TempVars!GenerateToC) Then TempVars!GenerateToC = True

30        If TempVars!GenerateToC Then
              'Collect ToC information. This is the report output to PDF.

40            strTitle = Me.ProductCategoryName
50            intPageNumber = Me.Page

              'Check to see if the current title already occurs in the Catalog_TableOfContents
60            If DCount("*", "Catalog_TableOfContents", StringFormatSQL("TocTitle = {0}", strTitle)) = 0 Then
70                strSQL = StringFormatSQL("INSERT INTO Catalog_TableOfContents (TocTitle, TocPage) VALUES ({0}, {1});", strTitle, intPageNumber)
80                g_dbApp().Execute strSQL, dbFailOnError
90            End If

              'Activate this section if you have a subfolder with images named according to ProductCategories.ProductCategoryImageFile.
              'For image licensing reasons this template displays a placeholder image.
              '    strImageFile = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\Images\")) & Me.ProductCategoryImageFile
              '    If Len(Dir(strImageFile)) > 0 Then
              '        Me.imgProductCategory.Picture = strImageFile
              '        Me.imgProductCategory.Visible = True
              '    Else
              '        Me.imgProductCategory.Visible = False
              '    End If

100       Else
              'Nothing to do. This is the report in Preview mode.
110       End If

Exit_Handler:
120       Exit Sub

Err_Handler:
130       clsErrorHandler.HandleError Me.Name, "GroupHeader0_Print", True
140       Resume Exit_Handler
End Sub

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

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Report_NoData", True
50        Resume Exit_Handler
End Sub

Private Sub Report_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

          Dim strSQL              As String

          'If TempVars!GenerateToC is not initialized (likely if report is not run from frmReports), then initialize it now.
          'NOTE: this does NOT actually generate the full report with ToC. Run the report from frmReports to get that special processing.
20        If IsNull(TempVars!GenerateToC) Then TempVars!GenerateToC = True

          ' Delete all records from the table of contents table [Catalog_TableOfContents] so we can collect fresh data in GroupHeader0_Print event.
30        If TempVars!GenerateToC Then
40            strSQL = "DELETE * FROM Catalog_TableOfContents;"
50            g_dbApp().Execute strSQL, dbFailOnError
60        Else
              'Don't delete the Table of Contents
70        End If

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "Report_Open", True
100       Resume Exit_Handler
End Sub
