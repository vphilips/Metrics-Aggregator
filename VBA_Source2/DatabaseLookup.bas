Attribute VB_Name = "DatabaseLookup"
Option Explicit

' ==============================================================================
' Module: DatabaseLookup
' Description: Look up Account Name -> ID using the 'database' worksheet.
'              Caches the lookup table in a Dictionary for performance.
' ==============================================================================

Private Const SHEET_DB As String = "database"
Private Const COL_ACC_NAME As String = "A"
Private Const COL_GLOBAL_ID As String = "C"

Private pAccountCache As Object ' Scripting.Dictionary

' ------------------------------------------------------------------------------
' Sub: InitializeCache
' Purpose: Reads the 'database' sheet into memory if not already done.
' ------------------------------------------------------------------------------
Private Sub InitializeCache()
    If Not pAccountCache Is Nothing Then Exit Sub
    
    Set pAccountCache = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim key As String, val As String
    
    Set ws = ThisWorkbook.Sheets(SHEET_DB)
    lastRow = ws.Cells(ws.Rows.Count, COL_ACC_NAME).End(xlUp).Row
    
    ' Assuming Row 1 is header
    For i = 2 To lastRow
        key = Trim(CStr(ws.Cells(i, COL_ACC_NAME).Value))
        val = Trim(CStr(ws.Cells(i, COL_GLOBAL_ID).Value))
        
        If Len(key) > 0 Then
            ' If duplicates exist, this takes the last one (or use .Exists check to keep first)
            ' Using Case Insensitive key for safety
            If Not pAccountCache.Exists(UCase(key)) Then
                pAccountCache.Add UCase(key), val
            End If
        End If
    Next i
End Sub

' ------------------------------------------------------------------------------
' Function: GetInvestorGlobalId
' Purpose: Returns the Global ID for a given Account Name.
'          Returns Empty string if not found.
' ------------------------------------------------------------------------------
Public Function GetInvestorGlobalId(ByVal accountName As String) As String
    InitializeCache
    
    Dim lookupKey As String
    lookupKey = UCase(Trim(accountName))
    
    If pAccountCache.Exists(lookupKey) Then
        GetInvestorGlobalId = pAccountCache(lookupKey)
    Else
        GetInvestorGlobalId = ""
    End If
End Function

' ------------------------------------------------------------------------------
' Function: GetGlobalIdsForList
' Purpose: Takes a Collection of Account Names and returns a CSV string of IDs.
'          Throws error if an account is missing.
' ------------------------------------------------------------------------------
Public Function GetGlobalIdsForList(ByVal selectedAccounts As Collection) As String
    Dim acc As Variant
    Dim id As String
    Dim result As String
    Dim missingList As String
    
    For Each acc In selectedAccounts
        id = GetInvestorGlobalId(CStr(acc))
        If Len(id) = 0 Then
            missingList = missingList & acc & vbNewLine
        Else
            result = result & id & ","
        End If
    Next acc
    
    If Len(missingList) > 0 Then
        Err.Raise vbObjectError + 1001, "DatabaseLookup", _
                  "The following accounts were not found in the database sheet:" & vbNewLine & missingList
    End If
    
    ' Remove trailing comma
    If Len(result) > 0 Then
        If Right(result, 1) = "," Then result = Left(result, Len(result) - 1)
    End If
    
    GetGlobalIdsForList = result
End Function
