Attribute VB_Name = "modDAO"
Option Compare Database
Option Explicit

'ALGORITHM:
'   Open readonly recordset on table. Move a random number of rows down, and select that PK value.
'NOTE:
'   Function assumes PK is a single field and Long. This is a best practice and True for Northwind.
Public Function GetRandomPkValue(ByVal strTable As String, ByVal strPkField As String) As Long
10        On Error GoTo Err_Handler

          Dim rs                  As DAO.Recordset

20        Set rs = CurrentDb.OpenRecordset(strTable, dbOpenSnapshot)
30        With rs
40            .MoveLast     'This assumes there is at least one record
50            .MoveFirst
60            .Move GetRandom(0, .RecordCount - 1)
70            GetRandomPkValue = .Fields(strPkField)
80            .Close
90        End With

Exit_Handler:
100       Exit Function

Err_Handler:
110       clsErrorHandler.HandleError "modDAO", "GetRandomPkValue"
120       Resume Exit_Handler
End Function

'REFERENCE:
'   https://stackoverflow.com/questions/31755802/can-i-use-access-vba-to-determine-if-a-table-has-a-data-macro
'NOTE:
'   While MSysObjects.LvExtra is officially undocumented, if you dump the bytes of this field for a table with table data macro,
'   you will see the XML in plain text. We use that for our algorithm: if the length of the field is at least the length
'   of the data macro code, we assume it has a data macro.
Public Function HasDataMacro(td As DAO.TableDef) As Boolean
10        On Error GoTo Err_Handler

          Dim intDmLength         As Integer
          Dim intLvExtraLength    As Integer

20        intDmLength = Len(DLookup("DataMacro", "Welcome"))
30        intLvExtraLength = Len(Nz(DLookup("LvExtra", "MSysObjects", StringFormatSQL("Name={0}", td.Name)), ""))
40        HasDataMacro = (intLvExtraLength >= intDmLength)

Exit_Handler:
50        Exit Function

Err_Handler:
60        clsErrorHandler.HandleError "modDAO", "HasDataMacro"
70        Resume Exit_Handler
End Function

'ALGORITHM:
'   Attempt to set a Field object to the given field name while "On Error Resume Next" is active,
'   and if no error then the field exists.
'   Alternatively we could have iterated over the Fields collection. That is probably slower.
Public Function HasField(td As DAO.TableDef, ByVal strField As String) As Boolean
          Dim fld                 As DAO.Field

10        On Error Resume Next
20        Set fld = td.Fields(strField)
30        HasField = (Err.Number = 0)
40        Set fld = Nothing
End Function

'PROCEDURE:
'   HasProperty
'PURPOSE:
'   Test if the given object has the named property in its Properties collection.
'ALGORITHM:
'   Set VBA to ignore errors. Try to use the property. If it works, it exists; if not, it does not exist.
'ARGUMENTS:
'   o           - Object that has a Properties collection
'   propName    - Name to look for
'RETURNS:
'   Boolean     - True if found, False otherwise.
'NOTE:
'   The alternative implementation is to loop over the Properties collection. That's several times slower.
Public Function HasProperty(o As Object, ByVal propName As String) As Boolean
          Dim blnResult           As Boolean
          Dim strName             As String

10        On Error Resume Next
20        strName = o.Properties(propName).Name   'All properties have a Name property.
30        blnResult = (Err.Number = 0)
40        HasProperty = blnResult
End Function
