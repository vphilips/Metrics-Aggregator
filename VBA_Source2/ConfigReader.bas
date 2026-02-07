Attribute VB_Name = "ConfigReader"
Option Explicit

' ==============================================================================
' Module: ConfigReader
' Description: Reads configuration data from the 'form_config' worksheet.
'              Provides helper functions to populate UserForm controls.
' ==============================================================================

Private Const SHEET_CONFIG As String = "form_config"
Private Const COL_VAR_NAME As String = "A"
Private Const COL_CLIENT As String = "B"
Private Const COL_ACCOUNTS As String = "C"
Private Const COL_ASOF As String = "D"
Private Const COL_FROM As String = "E"
Private Const COL_TO As String = "F"
Private Const COL_ATTRS As String = "G"

' Cache to avoid reading sheet repeatedly (Optional optimization)
' For simplicity in this version, we will read directly from ranges or arrays.

' ------------------------------------------------------------------------------
' Function: GetDistinctVariables
' Purpose: Returns a Collection of unique Variable Names suitable for the given FormType.
' Note: The SRD implies filtering by FormType might be needed if the sheet mixes them.
'       If 'form_config' lists all variables, we just return all uniques.
' ------------------------------------------------------------------------------
Public Function GetDistinctVariables() As Collection
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim val As String
    Dim unique As New Collection
    Dim cellVal As Variant
    
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    lastRow = ws.Cells(ws.Rows.Count, COL_VAR_NAME).End(xlUp).Row
    
    On Error Resume Next
    For i = 2 To lastRow ' Assuming Header is Row 1
        val = Trim(CStr(ws.Cells(i, COL_VAR_NAME).Value))
        If Len(val) > 0 Then
            unique.Add val, val
        End If
    Next i
    On Error GoTo 0
    
    Set GetDistinctVariables = unique
End Function

' ------------------------------------------------------------------------------
' Function: GetClientsForVariable
' Purpose: Returns a Collection of unique Client names for the selected Variable.
' ------------------------------------------------------------------------------
Public Function GetClientsForVariable(ByVal variableName As String) As Collection
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim varCol As String, clientCol As String
    Dim varVal As String, clientVal As String
    Dim unique As New Collection
    
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    lastRow = ws.Cells(ws.Rows.Count, COL_VAR_NAME).End(xlUp).Row
    
    On Error Resume Next
    For i = 2 To lastRow
        varVal = Trim(CStr(ws.Cells(i, COL_VAR_NAME).Value))
        
        If StrComp(varVal, variableName, vbTextCompare) = 0 Then
            clientVal = Trim(CStr(ws.Cells(i, COL_CLIENT).Value))
            If Len(clientVal) > 0 Then
                unique.Add clientVal, clientVal
            End If
        End If
    Next i
    On Error GoTo 0
    
    Set GetClientsForVariable = unique
End Function

' ------------------------------------------------------------------------------
' Function: GetConfigRowDetails
' Purpose: Retrieves the full row details for a specific Variable + Client combination.
' Returns: A Scripting.Dictionary or a custom Type (using Collection for portability).
'          Keys: "Accounts", "Attributes", "AsOfDate", "FromDate", "ToDate"
' ------------------------------------------------------------------------------
Public Function GetConfigRowDetails(ByVal variableName As String, ByVal clientName As String) As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim varVal As String, clientVal As String
    Dim result As Object ' Scripting.Dictionary
    
    Set result = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    lastRow = ws.Cells(ws.Rows.Count, COL_VAR_NAME).End(xlUp).Row
    
    For i = 2 To lastRow
        varVal = Trim(CStr(ws.Cells(i, COL_VAR_NAME).Value))
        clientVal = Trim(CStr(ws.Cells(i, COL_CLIENT).Value))
        
        If StrComp(varVal, variableName, vbTextCompare) = 0 And _
           StrComp(clientVal, clientName, vbTextCompare) = 0 Then
            
            ' Found the match
            result("Accounts") = Trim(CStr(ws.Cells(i, COL_ACCOUNTS).Value))
            result("Attributes") = Trim(CStr(ws.Cells(i, COL_ATTRS).Value))
            result("AsOfDate") = ws.Cells(i, COL_ASOF).Value
            result("FromDate") = ws.Cells(i, COL_FROM).Value
            result("ToDate") = ws.Cells(i, COL_TO).Value
            
            Set GetConfigRowDetails = result
            Exit Function
        End If
    Next i
    
    ' Return empty if not found
    Set GetConfigRowDetails = result
End Function

' ------------------------------------------------------------------------------
' Function: ParseAccountsString
' Purpose: Splits a semi-colon separated string of accounts into a Collection.
' ------------------------------------------------------------------------------
Public Function ParseAccountsString(ByVal accountsStr As String) As Collection
    Dim parts() As String
    Dim i As Long
    Dim token As String
    Dim coll As New Collection
    
    If Len(accountsStr) = 0 Then
        Set ParseAccountsString = coll
        Exit Function
    End If
    
    parts = Split(accountsStr, ";")
    
    For i = LBound(parts) To UBound(parts)
        token = Trim(parts(i))
        If Len(token) > 0 Then
            coll.Add token
        End If
    Next i
    
    Set ParseAccountsString = coll
End Function
