Attribute VB_Name = "modMath"
Option Compare Database
Option Explicit

'REFERENCE:
'   https://docs.microsoft.com/office/vba/language/reference/user-interface-help/rnd-function
Public Function GetRandom(ByVal lngLowerBound As Long, ByVal lngUpperBound As Long) As Long
10        On Error GoTo Err_Handler

          'NOTE: Int function truncates fractional part. Works for Long values as well.
20        GetRandom = Int((lngUpperBound - lngLowerBound + 1) * Rnd + lngLowerBound)

Exit_Handler:
30        Exit Function

Err_Handler:
40        clsErrorHandler.HandleError "modMath", "GetRandom"
50        Resume Exit_Handler
End Function

'PROCEDURE:
'   MaxValue
'PURPOSE:
'   Using variants because user can pass in any data type, including dates.
'   Example: myMax = MaxValue(1, 3.5, 5)
'ARGUMENTS:
'   Any number of values (not an array)
'RETURNS:
'   Variant
Public Function MaxValue(ParamArray varValue() As Variant) As Variant
10        On Error GoTo Err_Handler

          Dim v                   As Variant
          Dim vMax                As Variant

20        If UBound(varValue) = -1 Then
30            vMax = Null        'Zero arguments passed in
40        Else
50            vMax = varValue(0)
60            For Each v In varValue
70                If v > vMax Then vMax = v
80            Next
90        End If
100       MaxValue = vMax

Exit_Handler:
110       Exit Function

Err_Handler:
120       clsErrorHandler.HandleError "modMath", "MaxValue"
130       Resume Exit_Handler
End Function
