Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_lngCompanyContactCount As Long
Private m_lngCustomerOrderCount As Long
Private m_lngShipperOrderCount As Long
Private m_lngVendorPurchaseOrderCount As Long
Private m_lngVendorProductCount As Long

'What record was the user on when they requested an Add New Company
'0 indicates that the add was requested from the CompanyList
Private m_lngCalledFromID   As Long

Private m_strMsgCaption     As String

Private Sub cboCompanyTypeID_AfterUpdate()
10        On Error GoTo Err_Handler

20        ManageFormOptions

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cboCompanyTypeID_AfterUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub cboCompanyTypeID_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

          Dim strDlgMsg           As String
          Dim strOpenArgs         As String

          'BUSINESS RULE:
          '   If the Company has any existing Orders, Purchase Orders, or Product Vendors
          '   the user cannot change the Company Type

20        If CompanyIsActive = True Or m_lngVendorProductCount > 0 Then
              '        strDlgMsg = "You can not change the Company Type if the Company has<br />" & _
                       '                    "    Customer Orders (" & m_lngCustomerOrderCount & ") <br />" & _
                       '                    "    Shipper Orders (" & m_lngShipperOrderCount & ") <br />" & _
                       '                    "    Vendor Purchase Orders (" & m_lngVendorPurchaseOrderCount & ") <br />" & _
                       '                    "    Vendor Products (" & m_lngVendorProductCount & ") <br />"

30            strDlgMsg = GetString(sCompanyDelete_CompanyTypeChange, "change the Company Type", m_lngCustomerOrderCount, m_lngShipperOrderCount, m_lngVendorPurchaseOrderCount, "Vendor Products (" & m_lngVendorProductCount & ")")

40            strOpenArgs = StringFormat("Header={0}&Message={1}&Actions={2}&FormCaption={3}", "Company Type Change Not Allowed", strDlgMsg, "vbOK", m_strMsgCaption)

50            DoCmd.OpenForm "frmGenericDialog", acNormal, , , acFormReadOnly, acDialog, strOpenArgs

60            If IsFormOpen("frmGenericDialog") Then
70                DoCmd.Close acForm, "frmGenericDialog", acSaveNo
80            End If

90            Me.cboCompanyTypeID.Undo

100           Cancel = True
110       End If

Exit_Handler:
120       Exit Sub

Err_Handler:
130       clsErrorHandler.HandleError Me.Name, "cboCompanyTypeID_BeforeUpdate", True
140       Resume Exit_Handler
End Sub

Private Sub cmdAdd_Click()
10        On Error GoTo Err_Handler

20        If Me.NewRecord Then
30            MsgBox GetString(sAlreadyOnNewRecord), vbInformation
40        Else
50            RunCommand acCmdRecordsGoToNew
60        End If

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "cmdAdd_Click", True
90        Resume Exit_Handler
End Sub

Private Sub cmdCancel_Click()
10        On Error GoTo Err_Handler

20        If m_lngCalledFromID = 0 Then
              'Add Company was requested by frmCompanyList
30            DoCmd.Close acForm, Me.Name
40        Else
50            If Not IsNull(Me.CompanyID) Then
60                If Me.Dirty Then
70                    Me.Undo
80                End If
90            End If
100           DoCmd.SearchForRecord acDataForm, Me.Name, acFirst, "CompanyID = " & m_lngCalledFromID
110       End If

Exit_Handler:
120       Exit Sub

Err_Handler:
130       clsErrorHandler.HandleError Me.Name, "cmdCancel_Click", True
140       Resume Exit_Handler
End Sub

Private Sub cmdClickToMap_Click()
10        On Error GoTo Err_Handler

          Dim strBusinessAddress  As String

          'Concatenate the Address into a single string to pass to the mapping function
20        strBusinessAddress = Me.[Address] & Space(2) & Me.[City] & ", " & Me.[StateAbbrev] & Space(2) & Me.[Zip]
30        Application.FollowHyperlink "http://maps.live.com/default.aspx?where1=" & strBusinessAddress

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "cmdClickToMap_Click", True
60        Resume Exit_Handler
End Sub

Private Sub cmdDeleteCompany_Click()
          'Companies can be deleted two ways
          '   1) User clicks on the Delete Company button (this code executes)
          '   2) User clicks on the record selector (grey vertical bar on the lfst side of the screen)
          '       and then hits the Delete key and the Form_Delete event fires
          '
          'When this code executes the RunCommand acCmdDeleteRecord
          '   the Form_Delete event executes same as option 2
          '
          'All logic for validating the delete request will be found in the Form_Delete event
10        On Error Resume Next

20        RunCommand acCmdDeleteRecord

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdDeleteCompany_Click"
50        Resume Exit_Handler

End Sub

Private Sub cmdEmail_Click()
10        On Error GoTo Err_Handler

20        DoCmd.RunCommand acCmdSend

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "cmdEmail_Click", True
50        Resume Exit_Handler
End Sub

Private Sub cmdSave_Click()
10        On Error GoTo Err_Handler

          Dim Cancel              As Integer       '

          'If you use DoCmd.RunCommand acCmdSaveRecord it will ALWAYS attempt a save
          'DoCmd.RunCommand acCmdSaveRecord

          'If you use Me.Dirty, it will ONLY attempt a save if there is something to save

20        Cancel = ValidateForm(Me)
30        If Not Cancel Then
              'The form is valid. Save it.
40            If Me.Dirty Then Me.Dirty = False

50            ValidateForm_RemoveHighlights Me
60        End If


Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "cmdSave_Click", True
90        Resume Exit_Handler
End Sub

Private Sub cmdShowFilter_Click()
10        On Error GoTo Err_Handler

          Dim strOpenArgs         As String
          Dim strFilter           As String

20        If Me.Form.FilterOn = True Then

30            strFilter = Me.Form.Filter

40            strFilter = Replace(strFilter, "[Lookup_cboCompanyTypeID].", "")
50            strFilter = Replace(strFilter, "[Lookup_cboState].", "")
60            strFilter = Replace(strFilter, "[Lookup_cboStandardTaxStatusID].", "")
70            strFilter = Replace(strFilter, "[qryCompanies].", "")

80            strOpenArgs = StringFormat("Header={0}&Message={1}&Actions={2}&FormCaption={3}", "Form Filter", strFilter, "vbOK", m_strMsgCaption)

90            DoCmd.OpenForm "frmGenericDialog", acNormal, , , acFormReadOnly, acDialog, strOpenArgs
100           If IsFormOpen("frmGenericDialog") Then
110               DoCmd.Close acForm, "frmGenericDialog", acSaveNo
120           End If
130       Else
140           MsgBox "This form has not been filtered.", vbOKOnly Or vbInformation
150       End If

Exit_Handler:
160       Exit Sub

Err_Handler:
170       clsErrorHandler.HandleError Me.Name, "cmdShowFilter_Click", True
180       Resume Exit_Handler
End Sub

Private Sub CompanyCantBeDeleted()
10        On Error GoTo Err_Handler

          Dim strDlgMsg           As String
          Dim strOpenArgs         As String

20        strDlgMsg = GetString(sCompanyDelete_CompanyTypeChange, "delete the Company", m_lngCustomerOrderCount, m_lngShipperOrderCount, m_lngVendorPurchaseOrderCount, "")

30        strOpenArgs = StringFormat("Header={0}&Message={1}&Actions={2}&FormCaption={3}", "Company Delete Not Allowed", strDlgMsg, "vbOK", m_strMsgCaption)

40        DoCmd.OpenForm "frmGenericDialog", acNormal, , , acFormReadOnly, acDialog, strOpenArgs

50        If IsFormOpen("frmGenericDialog") Then
60            DoCmd.Close acForm, "frmGenericDialog", acSaveNo
70        End If

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "CompanyCantBeDeleted"
100       Resume Exit_Handler
End Sub

Private Function CompanyConfirmDelete() As Boolean
10        On Error GoTo Err_Handler

          Dim strDlgMsg           As String
          Dim strOpenArgs         As String

20        If m_lngCompanyContactCount > 0 Or m_lngVendorProductCount > 0 Then
              'Confirm the Users intention to delete if they have Contacts or are a Product Vendor

30            strDlgMsg = "If you delete the Company any <br />" & _
                          "    Contacts (" & m_lngCompanyContactCount & ")or Product Vendors (" & m_lngVendorProductCount & ") will be deleted.<br /><br />" & _
                          "Is this OK?"
40            strOpenArgs = StringFormat("Header={0}&Message={1}&FormCaption={2}&Actions={3}", "Confirm Delete", strDlgMsg, m_strMsgCaption, "vbOKCancel")

50            DoCmd.OpenForm "frmGenericDialog", acNormal, , , acFormReadOnly, acDialog, strOpenArgs

60        Else
              'Confirm the Users intention to delete
70            strDlgMsg = "Are you sure you want to delete this Company?"
80            strOpenArgs = StringFormat("Header={0}&Message={1}&FormCaption={2}&Actions={3}", _
                                         "Confirm Delete", strDlgMsg, m_strMsgCaption, "vbYesNo")

90            DoCmd.OpenForm "frmGenericDialog", acNormal, , , acFormReadOnly, acDialog, strOpenArgs

100       End If

110       If IsFormOpen("frmGenericDialog") Then
              'User Clicked OKAY
120           CompanyConfirmDelete = True
130           DoCmd.Close acForm, "frmGenericDialog", acSaveNo
140       Else
150           CompanyConfirmDelete = False
160       End If

Exit_Handler:
170       Exit Function

Err_Handler:
180       clsErrorHandler.HandleError Me.Name, "CompanyConfirmDelete"
190       Resume Exit_Handler
End Function

Private Function CompanyIsActive() As Boolean
10        On Error GoTo Err_Handler

          Dim lngCount            As Long

          'BUSINESS RULE:
          'Check All Foreign Key Relationships
          'A Company Is Active it they have Orders or Purchase Orders
          'We wont check Contacts or Product Vendors, because we can
          '   delete those for the end user if there are no Orders or Purchase Orders
          '   We count these so we know if we need to delete them

20        m_lngCompanyContactCount = DCount("*", "Contacts", StringFormatSQL("CompanyID={0}", Me.CompanyID))
30        m_lngCustomerOrderCount = DCount("*", "Orders", StringFormatSQL("CustomerID={0}", Me.CompanyID))
40        m_lngShipperOrderCount = DCount("*", "Orders", StringFormatSQL("ShipperID={0}", Me.CompanyID))
50        m_lngVendorPurchaseOrderCount = DCount("*", "PurchaseOrders", StringFormatSQL("VendorID={0}", Me.CompanyID))
60        m_lngVendorProductCount = DCount("*", "ProductVendors", StringFormatSQL("VendorID={0}", Me.CompanyID))

          'Get the Count of Active Orders or Purchase Orders
70        lngCount = m_lngCustomerOrderCount + m_lngShipperOrderCount + m_lngVendorPurchaseOrderCount

80        If lngCount > 0 Then
90            CompanyIsActive = True
100       Else
110           CompanyIsActive = False
120       End If

Exit_Handler:
130       Exit Function

Err_Handler:
140       clsErrorHandler.HandleError Me.Name, "CompanyIsActive"
150       Resume Exit_Handler
End Function

Private Sub Form_AfterDelConfirm(Status As Integer)
10        On Error GoTo Err_Handler

20        Select Case Status
              Case acDeleteOK
30                MsgBox "Deleted!", vbInformation Or vbOKOnly
40                If IsFormOpen("frmCompanyList") = True Then
50                    Forms("frmCompanyList").Requery
60                End If
70        End Select

Exit_Handler:
80        Exit Sub

Err_Handler:
90        clsErrorHandler.HandleError Me.Name, "Form_AfterDelConfirm", True
100       Resume Exit_Handler
End Sub

Private Sub Form_AfterInsert()
10        On Error GoTo Err_Handler

20        If IsFormOpen("frmCompanyList") = True Then
30            Forms("frmCompanyList").Requery
40            Me.SetFocus
50        End If

Exit_Handler:
60        Exit Sub

Err_Handler:
70        clsErrorHandler.HandleError Me.Name, "Form_AfterInsert", True
80        Resume Exit_Handler
End Sub

Private Sub Form_AfterUpdate()
10        On Error GoTo Err_Handler

          'In case there are any left over
20        ValidateForm_RemoveHighlights Me

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_AfterUpdate", True
50        Resume Exit_Handler
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
10        On Error GoTo Err_Handler

          'https://docs.microsoft.com/previous-versions/office/developer/office-2003/aa211249(v=office.11)
          'The link above is for Access 2003 but how this event works has not changed
          'and this document gives a good explanation

          'Access will ask the User to confirm the Delete action
          'We want to Display our own message box - not the default one
          'To do that we set the response to acDataErrContinue

20        Response = acDataErrContinue

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "Form_BeforeDelConfirm", True
50        Resume Exit_Handler
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
10        On Error GoTo Err_Handler

          'This is an example where the business rules are such that saving a record needs to be confirmed.
20        If MsgBox(GetString(sDoYouWantToSaveYourChanges), vbQuestion Or vbYesNo, "Save Your Changes?") = vbNo Then
30            Me.Undo
40        Else
50            Cancel = ValidateForm(Me)
60        End If

Exit_Handler:
70        Exit Sub

Err_Handler:
80        clsErrorHandler.HandleError Me.Name, "Form_BeforeUpdate", True
90        Resume Exit_Handler
End Sub

Private Sub Form_Current()
10        On Error GoTo Err_Handler

20        If Not Me.NewRecord Then m_lngCalledFromID = Me.CompanyID

30        ManageFormOptions

          'In case there are any left over
40        ValidateForm_RemoveHighlights Me

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "Form_Current", True
70        Resume Exit_Handler
End Sub

Private Sub Form_Delete(Cancel As Integer)
10        On Error GoTo Err_Handler

          Dim strSQL              As String

          'BUSINESS RULE:
          '   If the Company has any existing Orders or Purchase Orders
          '           the User can not delete the Company
20        If CompanyIsActive = True Then
30            CompanyCantBeDeleted
40            Cancel = True
50            GoTo Exit_Handler
60        End If

          'Confirm the User Delete Request
70        If CompanyConfirmDelete Then
80            If m_lngCompanyContactCount > 0 Then
90                strSQL = StringFormatSQL("Delete from Contacts where CompanyID = {0}", Me.txtCompanyID.Value)
100               g_dbApp().Execute strSQL, dbFailOnError
110           End If

120           If m_lngVendorProductCount > 0 Then
130               strSQL = StringFormatSQL("Delete from ProductVendors where VendorID = {0}", Me.txtCompanyID.Value)
140               g_dbApp().Execute strSQL, dbFailOnError
150           End If
160       Else
170           Cancel = True
180           Exit Sub
190       End If

Exit_Handler:
200       Exit Sub

Err_Handler:
210       clsErrorHandler.HandleError Me.Name, "Form_Delete", True
220       Resume Exit_Handler
End Sub

'==================================================================
'Using the Calling Form Filter
'==================================================================
'
'The form frmCompanyList is rich with opportunities
'   to filter the Companies in Northwind.
'
'There is an option group at the top of the form to
'   filter the results by the type of Company
'
'The user can right click in any field to access
'   a variety built in filtering options and
'   multiple columns can be filtered
'
'Once the data has been filtered to the subset
'   the user wants to review, double click
'   anywhere in frmCompanyList to open this form
'
'Open this form with only the records which match
'   the subset created in frmCompanyList
'
'To achieve that we are using the filter from
'   frmCompanyList
'
'But it's not that simple.  The data on
'   frmCompanyList is denormalized and the
'   data on this form is not (so we can do data entry)
'
'If the filter from frmCompanyList looks like this
'
'   ([CompanyType]<>"Vendor" Or [CompanyType] IS Null)
'
' and the field CompanyType is not in the table Company
' (CompanyTypeID is) how can we use the filter passed in?
'
'Look at the query for this form and you will see how
'   we got the fields we needed to be able to apply the
'   filter and kept the query updateable
'
'==================================================================
Private Sub Form_Load()
10        On Error GoTo Err_Handler

          Dim dict                As Scripting.Dictionary
          Dim key                 As Variant
          Dim lngOpenArgCompanyID As Long
          Dim strOpenArgAction    As String
          Dim strOpenArgCompanyName As String
          Dim strOpenArgCompanyTypeID As String
          Dim strOpenArgPassedFilter As String
          Dim strFormWhere        As String

          '===================================================================================================
          'NOTE:
          'Often, we need to take different actions depending on the circumstances
          'For example, If a first time Customer give a 10% discount
          'This type of programming is referred to as Branching or Conditional Execution.
          '
          'This form provided the opportunity to demonstrate 4 different
          '  approaches to Branching or Conditional Execution
          '===================================================================================================


          'Parse OpenArgs
          'Split the String created by frmCompanyList into Key Value Pairs
          'Key Value Pairs are a set of two data items - a KEY and a VALUE
          'To see how the string is broken into the Key Value Pairs look at
          '  StringToDictionary function in Northwind

          'Option 1 - Case Statement
          '   Populate the correct variable based on the KEY
          'https://docs.microsoft.com/office/vba/Language/Reference/User-Interface-Help/select-case-statement

          'Initialize the module level variable
20        m_lngCalledFromID = 0    'Indicates List Form
30        m_strMsgCaption = "Message from Company"

40        If IsNull(Me.OpenArgs) Then
              'Form is opened without any specific instructions in OpenArgs.
50        Else
              'Parse Open Args into a Dictionary Object
60            Set dict = StringToDictionary(Me.OpenArgs)
              'Loop thru the Dictionary object
70            For Each key In dict.Keys
80                Select Case key
                      Case "Action"
90                        strOpenArgAction = dict.Item(key)
100                   Case "CompanyName"
110                       strOpenArgCompanyName = dict.Item(key)
120                   Case "CompanyTypeID"
130                       strOpenArgCompanyTypeID = dict.Item(key)
140                   Case "CompanyID"
150                       lngOpenArgCompanyID = dict.Item(key)
160                   Case "PassedFilter"
170                       strOpenArgPassedFilter = dict.Item(key)
180                   Case "[frmOrderDetails]![CustomerID]"           'Opened from the OrderDetails form, entering a new company in the dropdown.
190                       strOpenArgCompanyName = dict.Item(key)
200                       strOpenArgAction = "Add"
210               End Select
220           Next key

230           Set dict = Nothing
240       End If


          'Process In This Order. Don't switch these options around.

          'Option 2 - Nested If Else
          '    Build mFormWhere
          'https://docs.microsoft.com/office/vba/Language/Concepts/Getting-Started/using-ifthenelse-statements

250       If Len(strOpenArgCompanyTypeID) > 0 Then
260           If Len(strOpenArgCompanyTypeID) = 1 Then
270               Me.cboCompanyTypeID.DefaultValue = strOpenArgCompanyTypeID
280               strFormWhere = "CompanyTypeID = " & strOpenArgCompanyTypeID
290           Else
300               strFormWhere = "CompanyTypeID " & strOpenArgCompanyTypeID
310           End If
320       End If


          'Option 3 - ElseIf
          'https://docs.microsoft.com/office/vba/Language/Concepts/Getting-Started/using-ifthenelse-statements

          'See the comment Using the Parent Form Filter at the top of this module

330       If Len(strOpenArgPassedFilter) > 0 Then
340           Me.Filter = strOpenArgPassedFilter
350           Me.FilterOn = True
360       ElseIf Len(strFormWhere) > 0 Then
370           Me.Filter = strFormWhere    'No Passed Filter use strFormWhere
380           Me.FilterOn = True
390       Else
400           Me.FilterOn = False
410       End If


          'Option 4 - Nested If Else and ElseIF
          'https://docs.microsoft.com/office/vba/Language/Concepts/Getting-Started/using-ifthenelse-statements

420       If Len(strOpenArgAction) > 0 Then
430           If strOpenArgAction = "Add" Then
440               DoCmd.GoToRecord acDataForm, Me.Name, acNewRec
450               If Len(strOpenArgCompanyName) > 0 Then
460                   Me.txtCompanyName = strOpenArgCompanyName
470               End If
480           Else
                  'The provided Action is unexpected
                  'Tell the user this has happened
490               MsgBox "An unknown action has been requested. " & vbCrLf & vbCrLf & "Please let IT know this has happened and that action is :  " & strOpenArgAction, vbOKOnly Or vbCritical, "Contact IT"
500           End If
510       ElseIf Len(lngOpenArgCompanyID) > 0 Then
520           DoCmd.SearchForRecord acDataForm, Me.Name, acFirst, "CompanyID = " & lngOpenArgCompanyID
530           m_lngCalledFromID = lngOpenArgCompanyID
540       End If

Exit_Handler:
550       Exit Sub

Err_Handler:
560       clsErrorHandler.HandleError Me.Name, "Form_Load", True
570       Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)
10        On Error GoTo Err_Handler

          'Clean up anything left by a previous session
          'In order to have the clear the filter, you need 2 lines of code as shown below, and they must happen in this order.
          'https://stackoverflow.com/questions/59954653/vba-ms-access-remove-form-filter
20        Me.FilterOn = False
30        Me.Filter = ""

Exit_Handler:
40        Exit Sub

Err_Handler:
50        clsErrorHandler.HandleError Me.Name, "Form_Open", True
60        Resume Exit_Handler
End Sub

Private Sub ManageFormOptions()
10        On Error GoTo Err_Handler

20        With Me
30            .sfrmOrders.SourceObject = ""

40            .Caption = .txtCompanyType.Value

              'Set the subform in the upper right section based on Company Type
50            Select Case .cboCompanyTypeID.Value
                  Case enumCompanyType.ctCustomer
60                    .sfrmOrders.SourceObject = "sfrmCompanyDetail_CustomerOrders"
70                    .lblsfrmOrders.Caption = "Orders"
80                Case enumCompanyType.ctShipper
90                    .sfrmOrders.SourceObject = "sfrmCompanyDetail_ShipperOrders"
100                   .lblsfrmOrders.Caption = "Orders Shipped"
110               Case enumCompanyType.ctVendor
120                   .sfrmOrders.SourceObject = "sfrmCompanyDetail_VendorPurchaseOrders"
130                   .lblsfrmOrders.Caption = "Purchase Orders"
140               Case enumCompanyType.ctNorthwind
150                   .sfrmOrders.SourceObject = ""   'Northwind itself does not place orders.
160                   .lblsfrmOrders.Caption = ""
170               Case Else
180                   Debug.Assert False      'Unexpected company type is not yet implemented.
190           End Select

200       End With

Exit_Handler:
210       Exit Sub

Err_Handler:
220       clsErrorHandler.HandleError Me.Name, "ManageFormOptions"
230       Resume Exit_Handler
End Sub

Private Sub txtBusinessPhone_Click()
10        On Error GoTo Err_Handler

          'If the Phone has not been populated yet
          'position the cursor to the beginning of the field no matter where the user clicked
20        If Len(Me.txtBusinessPhone) = 0 Then
30            Me.txtBusinessPhone.SelStart = 0
40        End If

Exit_Handler:
50        Exit Sub

Err_Handler:
60        clsErrorHandler.HandleError Me.Name, "txtBusinessPhone_Click", True
70        Resume Exit_Handler
End Sub

Private Sub txtWebsite_DblClick(Cancel As Integer)
10        On Error GoTo Err_Handler

20        Application.FollowHyperlink Me.txtWebsite

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError Me.Name, "txtWebsite_DblClick", True
50        Resume Exit_Handler
End Sub
