Attribute VB_Name = "modCompanies"
Option Compare Database
Option Explicit

Public Function CalculatePrice() As Currency
    CalculatePrice = 100
End Function


Public Function CalculatePrice200() As Currency
    CalculatePrice = 200
End Function


Public Function GetRandomCustomerID() As Long
10        On Error GoTo Err_Handler

20        GetRandomCustomerID = GetRandomPkValue("qryCustomers", "CompanyID")

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modCompanies", "GetRandomCustomerID"
50        Resume Exit_Handler
End Function

Public Function GetTaxStatusID(ByVal lngCompanyID As Long) As Byte
10        On Error GoTo Err_Handler

20        GetTaxStatusID = DLookup("StandardTaxStatusID", "Companies", "CompanyID = " & lngCompanyID)

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modCompanies", "GetTaxStatusID"
50        Resume Exit_Handler
End Function
