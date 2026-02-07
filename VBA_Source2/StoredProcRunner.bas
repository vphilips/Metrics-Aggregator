Attribute VB_Name = "StoredProcRunner"
Option Explicit

' ==============================================================================
' Module: StoredProcRunner
' Description: Executes Stored Procedures and writes output to 'Results' sheet.
'              Implementation matches "Robust/Image" pattern requested.
' ==============================================================================

' Configuration Constants
Private Const USE_SCALABLE_PATTERN As Boolean = True ' Hardcoded choice as requested

' Database Config (Adjust these as needed or read from Config sheet)
' NOTE: These are placeholders based on the image implicating CFG_SERVER constants.
' You should update these to match your actual server/db.
Private Const CFG_SERVER As String = "YOUR_SERVER"
Private Const CFG_DATABASE As String = "YOUR_DB"
Private Const CFG_CONN_TIMEOUT_SEC As Long = 30
Private Const CFG_CMD_TIMEOUT_SEC As Long = 120

Private Const PROVIDER_PRIMARY As String = "MSOLEDBSQL"
Private Const PROVIDER_FALLBACK As String = "SQLOLEDB"

' ------------------------------------------------------------------------------
' Main Execution Router
' ------------------------------------------------------------------------------
Public Sub RunStoredProc(ByVal procName As String, _
                         ByVal params As Object, _
                         ByVal outputSheetName As String)
    
    If USE_SCALABLE_PATTERN Then
        ExecProcToTable_Robust procName, params, outputSheetName
    Else
        ' Legacy/Simple implementation (previously generated)
        ExecProcToTable_Simple procName, params, outputSheetName
    End If

End Sub

' ------------------------------------------------------------------------------
' Robust Implementation (Matching Images)
' ------------------------------------------------------------------------------
Private Sub ExecProcToTable_Robust(ByVal procName As String, _
                                   ByVal params As Object, _
                                   ByVal outputSheetName As String)

    Dim cn As Object ' ADODB.Connection
    Dim cmd As Object ' ADODB.Command
    Dim rs As Object ' ADODB.Recordset
    Dim ws As Worksheet
    
    ' 1. Prepare Target
    Set ws = GetOrAddSheet(outputSheetName)
    ws.Cells.Clear
    
    ' 2. Create Connection
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = CFG_CONN_TIMEOUT_SEC
    cn.CommandTimeout = CFG_CMD_TIMEOUT_SEC
    
    ' 3. Open with Modern/Fallback Logic
    cn.Open BuildConnStringPreferModern()
    
    ' 4. Create Command
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 4 ' adCmdStoredProc
    cmd.CommandText = procName
    cmd.CommandTimeout = CFG_CMD_TIMEOUT_SEC
    
    ' 5. Add Parameters
    If Not params Is Nothing Then
        Dim key As Variant
        For Each key In params.Keys
            ' We simplify type detection here.
            ' Image shows: AddParam cmd, "@name", 200, val, 100
            ' We will infer typical string settings (VarChar, Input, 8000)
            ' Adjust logic if you need strict types (Integer/Date)
            AddParam cmd, CStr(key), 200, CStr(params(key)), 8000
        Next key
    End If
    
    ' 6. Execute & Recordset
    ' Image: rs.CursorLocation = 3 (adUseClient)
    ' Image: rs.Open cmd, , 0, 1 (adOpenForwardOnly, adLockReadOnly)
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open cmd, , 0, 1 ' adOpenForwardOnly, adLockReadOnly
    
    ' 7. Write Output
    If Not rs.EOF Then
        Dim i As Integer
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(1, i + 1).Value = rs.Fields(i).Name
        Next i
        ws.Range("A2").CopyFromRecordset rs
        MsgBox "Query completed successfully.", vbInformation, "Success"
    Else
        MsgBox "Query executed but returned no records.", vbExclamation, "No Data"
    End If
    
    ' Cleanup
    If rs.State = 1 Then rs.Close
    If cn.State = 1 Then cn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set cn = Nothing
    
End Sub

' ------------------------------------------------------------------------------
' Helper: AddParam (Matching Image)
' ------------------------------------------------------------------------------
Private Sub AddParam(ByVal cmd As Object, ByVal name As String, ByVal adoType As Long, _
                     ByVal value As Variant, Optional ByVal size As Long = 0)
    Dim p As Object
    If size > 0 Then
        Set p = cmd.CreateParameter(name, adoType, 1, size) ' 1=adParamInput
    Else
        Set p = cmd.CreateParameter(name, adoType, 1)
    End If
    p.Value = value
    cmd.Parameters.Append p
End Sub

' ------------------------------------------------------------------------------
' Helper: Connection String Builder (Matching Image)
' ------------------------------------------------------------------------------
Private Function BuildConnStringPreferModern() As String
    Dim csPrimary As String, csFallback As String
    Dim cnTest As Object
    
    ' Primary: MSOLEDBSQL
    csPrimary = "Provider=" & PROVIDER_PRIMARY & ";" & _
                "Server=" & CFG_SERVER & ";" & _
                "Database=" & CFG_DATABASE & ";" & _
                "Integrated Security=SSPI;" & _
                "Encrypt=yes;" & _
                "TrustServerCertificate=yes;" & _
                "Application Name=ExcelVBAExtract;"
                
    ' Fallback: SQLOLEDB
    csFallback = "Provider=" & PROVIDER_FALLBACK & ";" & _
                 "Server=" & CFG_SERVER & ";" & _
                 "Database=" & CFG_DATABASE & ";" & _
                 "Integrated Security=SSPI;" & _
                 "Application Name=ExcelVBAExtract;"
                 
    ' Try Primary
    On Error Resume Next
    Set cnTest = CreateObject("ADODB.Connection")
    cnTest.Open csPrimary
    
    If Err.Number = 0 And cnTest.State = 1 Then
        cnTest.Close
        BuildConnStringPreferModern = csPrimary
        Exit Function
    End If
    On Error GoTo 0
    
    ' Use Fallback
    BuildConnStringPreferModern = csFallback
    
End Function

' ------------------------------------------------------------------------------
' Simple Implementation (Legacy backup)
' ------------------------------------------------------------------------------
Private Sub ExecProcToTable_Simple(ByVal procName As String, _
                                   ByVal params As Object, _
                                   ByVal outputSheetName As String)
    ' ... (Original simpler code if needed, omitted for brevity but logic is similar)
    ' Re-implementing basic logic strictly for fallback reference
    Dim cn As Object, cmd As Object, rs As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open BuildConnStringPreferModern() ' Re-use the smart string anyway
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 4
    cmd.CommandText = procName
    
    If Not params Is Nothing Then
        Dim k As Variant
        For Each k In params.Keys
            cmd.CreateParameter(CStr(k), 200, 1, 8000, CStr(params(k)))
        Next k
    End If
    
    Set rs = cmd.Execute
    Dim ws As Worksheet
    Set ws = GetOrAddSheet(outputSheetName)
    ws.Cells.Clear
    If Not rs.EOF Then
        ws.Range("A2").CopyFromRecordset rs
    End If
    rs.Close
    cn.Close
End Sub

' ------------------------------------------------------------------------------
' Helper: Sheet Management
' ------------------------------------------------------------------------------
Private Function GetOrAddSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrAddSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If GetOrAddSheet Is Nothing Then
        Set GetOrAddSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrAddSheet.Name = sheetName
    End If
End Function

' ------------------------------------------------------------------------------
' Wrappers (Unchanged API for FormLogic)
' ------------------------------------------------------------------------------
Public Sub ExecuteMyPerformance(metricName As String, investorIds As String, asOfDateNum As String, rptCurrency As String, outFields As String)
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "@MetricName", metricName
    params.Add "@InvestorNameIds", investorIds
    params.Add "@AsOfDate", asOfDateNum
    params.Add "@ReportingCurrency", rptCurrency
    params.Add "@OutputFields", outFields
    RunStoredProc "edw.usp_InvestorPerformance_Aggregated", params, "Results"
End Sub

Public Sub ExecutePortfolioDiversification(metricName As String, investorIds As String, asOfDateNum As String, rptCcyId As String, outFields As String)
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "@Metric", metricName
    params.Add "@InvestorNameIdsCsv", investorIds
    params.Add "@AsOfDate", asOfDateNum
    params.Add "@ReportingCurrencyId", rptCcyId
    params.Add "@OutputFieldsCsv", outFields
    RunStoredProc "edw.usp_PortfolioDiversification_Aggregated", params, "Results"
End Sub

Public Sub ExecuteCompanyDiversification(metricName As String, investorIds As String, asOfDateStr As String, rptCcyId As String, outFields As String)
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "@MetricName", metricName
    params.Add "@InvestorNameIdsCsv", investorIds
    params.Add "@AsOfDate", asOfDateStr
    params.Add "@ReportingCurrencyId", rptCcyId
    params.Add "@OutputFieldsCsv", outFields
    RunStoredProc "edw.usp_CompanyDiversification_Aggregated", params, "Results"
End Sub

Public Sub ExecuteHistoricalCashflows(metricName As String, investorIds As String, startDate As String, endDate As String, rptCcyId As String, outFields As String)
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "@Metric", metricName
    params.Add "@InvestorIdsCsv", investorIds
    params.Add "@StartDate", startDate
    params.Add "@EndDate", endDate
    params.Add "@ReportingCurrencyId", rptCcyId
    params.Add "@OutputFieldsCsv", outFields
    RunStoredProc "edw.usp_HistoricalCashflowReport_Aggregated", params, "Results"
End Sub
