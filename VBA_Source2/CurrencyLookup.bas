Attribute VB_Name = "CurrencyLookup"
Option Explicit

' ==============================================================================
' Module: CurrencyLookup
' Description: Handling currency mapping from 'ccy_map' sheet.
' ==============================================================================

Private Const SHEET_CCY As String = "ccy_map"
' Adjust columns based on actual sheet layout.
' SRD says: Maps currency display/code to Currency Code (A?) and CurrencyId (B?)
' User info implies: Currency Code (e.g., USD) and CurrencyId (e.g. 8).
' We will assume: Col A = Display/Code, Col B = ID.
Private Const COL_CCY_CODE As String = "A"
Private Const COL_CCY_ID As String = "B"

Private pCcyIdCache As Object
Private pCcyCodeCache As Object

Private Sub InitializeCcyCache()
    If Not pCcyIdCache Is Nothing Then Exit Sub
    
    Set pCcyIdCache = CreateObject("Scripting.Dictionary")
    Set pCcyCodeCache = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim code As String, idVal As String
    
    Set ws = ThisWorkbook.Sheets(SHEET_CCY)
    lastRow = ws.Cells(ws.Rows.Count, COL_CCY_CODE).End(xlUp).Row
    
    For i = 2 To lastRow
        code = Trim(CStr(ws.Cells(i, COL_CCY_CODE).Value))
        idVal = Trim(CStr(ws.Cells(i, COL_CCY_ID).Value))
        
        If Len(code) > 0 Then
            If Not pCcyIdCache.Exists(code) Then
                pCcyIdCache.Add code, idVal
            End If
            ' Just in case we need reverse lookup or validity check
            If Not pCcyCodeCache.Exists(code) Then
                pCcyCodeCache.Add code, code
            End If
        End If
    Next i
End Sub

Public Function GetCurrencyId(ByVal ccyCode As String) As String
    InitializeCcyCache
    If pCcyIdCache.Exists(ccyCode) Then
        GetCurrencyId = pCcyIdCache(ccyCode)
    Else
        ' Default or Error?
        GetCurrencyId = "0" ' Return 0 or raise error
    End If
End Function

Public Function GetAllCurrencies() As Collection
    InitializeCcyCache
    Dim coll As New Collection
    Dim key As Variant
    For Each key In pCcyIdCache.Keys
        coll.Add key
    Next key
    Set GetAllCurrencies = coll
End Function

Public Function IsValidCurrency(ByVal ccyCode As String) As Boolean
    InitializeCcyCache
    IsValidCurrency = pCcyIdCache.Exists(ccyCode)
End Function
