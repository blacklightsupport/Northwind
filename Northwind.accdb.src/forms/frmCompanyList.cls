Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAddNew_Click()
10        On Error GoTo Err_Handler

          'Pass acNewRec in the OpenArgs to the form frmCompanyDetail
          'so it knows you want to add  a new record
20        Open_frmCompanyDetail "Add"

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdAddNew_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdLabelWizard_Click()
10        On Error Resume Next

20        DoCmd.RunCommand acCmdNewObjectLabelsReport

End Sub

Private Sub cmdShowFilter_Click()
10        On Error GoTo Err_Handler

          Dim strOpenArgs         As String

20        If Me.Form.FilterOn = False Then
30            MsgBox "This form has not been filtered.", vbOKOnly Or vbInformation
40        Else
              'OpenArgs uses Replace function to make the filter more readable for humans.
50            strOpenArgs = StringFormat("Header={0}&Message={1}", "Form Filter", Replace(Me.Form.Filter, "[qryCompanyList].", ""))
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

Private Sub cmdShowHideFields_Click()
10        On Error GoTo Err_Handler

20        DoCmd.RunCommand acCmdUnhideColumns

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdShowHideFields_Click", True
50        Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

          Dim ctl                 As Access.control

          'Clean up anything left by a previous session
20        Me.Auto_Header0.Caption = "Companies"
          'In order to have the change to the filter stick
          '  you need to set the FilterOn property to false
          'https://stackoverflow.com/questions/59954653/vba-ms-access-remove-form-filter
30        Me.FilterOn = False
40        Me.Filter = ""

          'The list cannot be edited, but the Filter should still be accessible, so Me.AllowEdits is no good.
50        On Error Resume Next        'Not every type of control has a Locked property.
60        For Each ctl In Me.Controls
70            ctl.Locked = (TypeName(ctl) = "TextBox")
80        Next ctl
90        On Error GoTo Err_Handler

Exit_Handler:
100       Exit Sub

Err_Handler:
110       clsErrorHandler.HandleError Me.Name, "Form_Open", True
120       Resume Exit_Handler
130       Resume
End Sub

Private Sub fraFilter_AfterUpdate()
10        On Error GoTo Err_Handler

20        With Me
30            Select Case .fraFilter

                  Case enumCompanyType.ctCustomer
40                    .Filter = "CompanyTypeID = " & Me.fraFilter
50                    .FilterOn = True
60                    .Auto_Header0.Caption = "Customers"
70                    .NavigationCaption = "Customers"
80                Case enumCompanyType.ctShipper
90                    .Filter = "CompanyTypeID = " & Me.fraFilter
100                   .FilterOn = True
110                   .Auto_Header0.Caption = "Shippers"
120                   .NavigationCaption = "Shippers"
130               Case enumCompanyType.ctVendor
140                   .Filter = "CompanyTypeID = " & Me.fraFilter
150                   .FilterOn = True
160                   .Auto_Header0.Caption = "Vendors"
170                   .NavigationCaption = "Vendors"
180               Case Else
190                   .FilterOn = False
200                   .Filter = ""
210                   .Auto_Header0.Caption = "Companies"
220                   .NavigationCaption = "Companies"
230           End Select
240       End With

Exit_Handler:
250       Exit Sub

Err_Handler:
260       clsErrorHandler.HandleError Me.Name, "fraFilter_AfterUpdate", True
270       Resume Exit_Handler
End Sub

'We want to pass multiple parameters to frmCompanyDetail via
'  the forms OpenArgs property
'To do that we need to pass a string which can be parsed by frmCompanyDetail
'
'To learn more about OpenArgs or Search For Record follow the link below
'https://docs.microsoft.com/office/vba/api/Access.Form.OpenArgs
'https://docs.microsoft.com/office/vba/api/access.docmd.searchforrecord
'
Public Function Open_frmCompanyDetail(ByVal strAction As String)
10        On Error GoTo Err_Handler

          Dim varOpenArgs         As Variant
          Dim strOpenArgAction    As String
          Dim strOpenArgCompanyTypeID As String
          Dim strOpenArgCompanyID As String
          Dim strOpenArgPassedFilter As String

20        If IsFormOpen("frmCompanyDetail") Then
30            MsgBox GetString(sFormAlreadyOpen, Forms!frmCompanyDetail.Caption), vbOKOnly Or vbInformation, "Already Open"
40            GoTo Exit_Handler
50        End If

          'First we will create the string for each parameter individually

60        If Len(strAction) > 0 Then
70            strOpenArgAction = "Action=" & strAction
80        Else
90            strOpenArgCompanyID = "CompanyID=" & Me.txtCompanyID
100       End If

110       If Me.fraFilter > 0 Then
120           strOpenArgCompanyTypeID = "CompanyTypeID=" & Me.fraFilter
130       End If

140       If Me.FilterOn Then
150           strOpenArgPassedFilter = Me.Filter
              'Remove the query name ([qryCompanyList].) which appears in the string
              'so the filter can be applied based only on field names in frmCompanyDetail
160           strOpenArgPassedFilter = "PassedFilter=" & Replace(strOpenArgPassedFilter, "[qryCompanyList].", "")
170       Else
180           strOpenArgPassedFilter = ""
190       End If

          'Secondly we will concatenate the strings above
          'into a single sting to pass in the OpenArgs parameter of the docmd.Open form event

          '***************************************************************
          'Using this place holder in the varOpenArgs string
          '   varOpenArgs = "OpenBy=frmCompanyList"
          '   eliminates the need to continously check
          '   if we are setting or concatenating to varOpenArgs
          '   Now we know we are always concatenating
          '
          '  WHAT WE DONT NEED TO DO FOR EACH POTENTIAL PARAMTER
          '***************************************************************
          '  If Len(varOpenArgs) = 0 then
          '       varOpenArgs = xyz
          '   Else
          '       varOpenArgs = varOpenArgs & "&" & xyz
          '   end if
          '***************************************************************

200       varOpenArgs = "OpenBy=frmCompanyList"    'Place holder

210       If Len(strOpenArgAction) > 0 Then
220           varOpenArgs = varOpenArgs & "&" & strOpenArgAction
230       End If

240       If Len(strOpenArgCompanyID) > 0 Then
250           varOpenArgs = varOpenArgs & "&" & strOpenArgCompanyID
260       End If

270       If Len(strOpenArgCompanyTypeID) > 0 Then
280           varOpenArgs = varOpenArgs & "&" & strOpenArgCompanyTypeID
290       End If

          'Concatenate this last incase it has formatting which interfers
          'with the parameter creation in frmCompanyDetail

300       If Len(strOpenArgPassedFilter) > 0 Then
310           varOpenArgs = varOpenArgs & "&" & strOpenArgPassedFilter
320       End If

330       DoCmd.OpenForm "frmCompanyDetail", acNormal, , , acFormEdit, acWindowNormal, varOpenArgs

Exit_Handler:
340       Exit Function

Err_Handler:
350       clsErrorHandler.HandleError Me.Name, "Open_frmCompanyDetail", True
360       Resume Exit_Handler
End Function

Private Sub txtCompanyID_Click()
10        On Error GoTo Err_Handler

20        Open_frmCompanyDetail ""

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "txtCompanyID_Click", True
50        Resume Exit_Handler
End Sub
