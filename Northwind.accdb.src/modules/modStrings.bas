Attribute VB_Name = "modStrings"
Option Compare Database
Option Explicit

Public Enum enumStrings     'Corresponds to StringID values in table Strings
    sHelloWorld = 1         'NOTE: the "s" = "Strings" prefix ensures there are no Reserved words in the Enum element names.
    sNumberBetween = 2
    sOrderMustBePaid = 4
    sOrderClosed = 5
    sOrderCannotDelete = 6
    sDeleteRecord = 7
    sRequiredFields = 8
    sRequiredFields_Shipping = 9
    sOrderMustBeInvoiced = 10
    sReportNoData = 11

    sNewEmployee = 12
    sEmployeeRestraints = 13
    sDeleteEmployee = 14
    sInvalidPhone = 15

    sDisclaimer = 16
    sOrderAllocated = 17
    sAlreadyOnNewRecord = 18
    sNoPrivilege = 19
    sPOMustBeNew = 20
    sPOMustBeSubmitted = 21
    sChangingVendor = 22
    sPostToInventory = 23
    sPOMustBeApproved = 24
    sPOMustBeReceived = 25
    sRequiredFields_PO_Close = 26
    sPOCannotDelete = 27
    sMinimumReorderQuantity = 28
    sLessThanTargetLevel = 29
    sNewStatusSet = 30
    sOrderPaid = 32
    sCannotDelete = 33
    sOneLineItem = 34
    sChangeCompanyType = 35
    sRequiredFields_Paid = 36
    sCompanyDelete_CompanyTypeChange = 37
    sCompanyDelete_RelatedRecords = 38
    sOptionNotAvailable_NewRecord = 39
    sNoRecentOrders = 40
    sAbout = 41
    sFormAlreadyOpen = 42

    sCatalog_GastronomicHeader = 43
    sCatalog_Gastronomic = 44
    sCatalog_QualityHeader = 45
    sCatalog_Quality = 46

    sDoYouWantToSaveYourChanges = 48
    sOrderMustBeShippedBeforePaid = 49
    sNewOrderInvoice = 50
    sOrderBeforeOrderLineItems = 51
End Enum

'PURPOSE:
'   Get a string from the Strings table. Supports {0} replacable parameters (see StringFormat).
Public Function GetString(ByVal ID As enumStrings, ParamArray params() As Variant) As String
10        On Error GoTo Err_Handler

          Dim s                   As String

20        s = Nz(DLookup("StringData", "Strings", "StringID = " & ID), "")
30        Debug.Assert s <> ""
40        s = StringFormat(s, params)

50        GetString = s

Exit_Handler:
60        Exit Function

Err_Handler:
70        clsErrorHandler.HandleError "modStrings", "GetString"
80        Resume Exit_Handler
End Function

'NOTES:
'   Patterned after String.Format in .NET.
'   https://docs.microsoft.com/dotnet/api/system.string.format
'   Advantage of this function over traditional string concatenation is that you can focus on the string itself, what you want it to say.
'   Also makes it easier to use strings from a table (see GetString function).
'EXAMPLES:
'   Debug.Print StringFormat("Hello {0}. This is {1}.", "world", "ET")
'   =>  Hello world. This is ET.
'ARGUMENTS:
'   s       - String with zero or more {n} placeholders for parameter values. The first one is {0}.
'   params  - Zero or more parameters to replace the placeholders. Can also handle a paramarray that was passed in from another function taking a paramarray (e.g. GetString).
Public Function StringFormat(ByVal s As String, ParamArray params() As Variant) As String
10        On Error GoTo Err_Handler

          Dim n                   As Integer
          Dim vParams             As Variant

20        If UBound(params) = -1 Then GoTo Exit_Handler       'Zero params passed in.
30        If IsArray(params) And IsArray(params(0)) Then
40            vParams = params(0)
50        Else
60            vParams = params
70        End If

80        For n = 0 To UBound(vParams)
90            s = Replace(s, "{" & n & "}", vParams(n))
100       Next n

Exit_Handler:
110       StringFormat = s
120       Exit Function

Err_Handler:
130       clsErrorHandler.HandleError "modStrings", "StringFormat"
140       Resume Exit_Handler
End Function

'PURPOSE:
'   Format a string using Access SQL rules.
'NOTES:
'   Better than built-in function Application.BuildCriteria which has some unexpected side-effects.
Public Function StringFormatSQL(ByVal s As String, ParamArray params() As Variant) As String
10        On Error GoTo Err_Handler

          Dim i                   As Integer
          Dim vParams             As Variant

20        If IsArray(params) And IsArray(params(0)) Then
30            vParams = params(0)
40        Else
50            vParams = params
60        End If

70        For i = LBound(vParams) To UBound(vParams)
80            If IsNull(vParams(i)) Then
90                vParams(i) = Nz(vParams(i), "NULL")
100           Else
                  'NW 2.0 code to be replaced below.
                  '            Select Case VarType(vParams(i))
                  '                Case vbString:
                  '                    vParams(i) = "'" & vParams(i) & "'"
                  '                Case vbDate:
                  '                    vParams(i) = "#" & vParams(i) & "#"
                  '            End Select

                  'NW 2.2 code will convert international number and date formats to those understood by Access.
                  'We also escape string arguments, in case an argument includes an embedded single-quote (think: ...where LastName = 'O'Brien")
110               Select Case VarType(vParams(i))
                      Case vbCurrency, vbSingle, vbDouble
120                       vParams(i) = LTrim(Str(vParams(i)))                     'Str converts regional numbers to Access standards.
130                   Case vbString:
140                       vParams(i) = SINGLE_QUOTE & Replace(vParams(i), SINGLE_QUOTE, TWO_SINGLE_QUOTES) & SINGLE_QUOTE
150                   Case vbDate:
160                       vParams(i) = "#" & ToAccessDate(vParams(i)) & "#"       'ToAccessDate ensures this works for all regional settings.
170               End Select


180           End If
190       Next i

200       StringFormatSQL = StringFormat(s, vParams)

Exit_Handler:
210       Exit Function

Err_Handler:
220       clsErrorHandler.HandleError "modStrings", "StringFormatSQL"
230       Resume Exit_Handler
240       Resume
End Function

' NOTE:
'   Caller is responsible for cleaning up the dictionary object returned. Example:
'   Dim dict as Scripting.Dictionary
'   Set dict = StringToDictionary("a=1&b=2")
'   'Use the dictionary object
'   Set dict = Nothing      'Cleanup.
' NOTE 2:
'   Note the proper syntax for calling a function that returns an object is by using Set:
'   Set dict = StringToDictionary(...)
' ARGUMENTS:
'   s   - String in querystring format (e.g. "key1=value1&key2=value2&key3=value3"
'         If a key contains an embedded equals sign or ampersand, you may want to call UrlEncode/UrlDecode  (code not included in this template).
' RETURNS:
'   Scripting.Dictionary object with the parsed values.
Public Function StringToDictionary(ByVal v As Variant) As Scripting.Dictionary      'Requires reference to Microsoft Scripting Runtime (scrrun.dll)
10        On Error GoTo Err_Handler

          Const KEYVALUE_DELIMITER As String = "&"
          Const VALUE_DELIMITER   As String = "="
          Dim dict                As New Scripting.Dictionary
          Dim intPos              As Integer
          Dim strTokens()         As String
          Dim varToken            As Variant

20        If IsNull(v) Then
              'Nothing to do.
30        Else
40            dict.CompareMode = vbTextCompare

50            strTokens = Split(v, KEYVALUE_DELIMITER)
60            For Each varToken In strTokens
70                intPos = InStr(varToken, VALUE_DELIMITER)
80                If intPos = 0 Then
                      'it's a token without an = sign
90                Else
100                   dict.Add Left$(varToken, intPos - 1), Mid$(varToken, intPos + 1)
110               End If
120           Next varToken

130           Erase strTokens
140       End If

150       Set StringToDictionary = dict

Exit_Handler:
160       Exit Function

Err_Handler:
170       clsErrorHandler.HandleError "modStrings", "StringToDictionary"
180       Resume Exit_Handler
End Function
