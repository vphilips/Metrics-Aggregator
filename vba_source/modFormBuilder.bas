Attribute VB_Name = "modFormBuilder"
Option Explicit

' ==================================================================================
' FORM BUILDER SCRIPT (WRAPPER)
' ==================================================================================
' This script delegates to modHistCashBuilder.BuildDynamicForm for form generation.
' ==================================================================================

Public Sub BuildAllUserForms()
    If Not CheckSecurityAccess() Then Exit Sub
    Application.ScreenUpdating = False
    
    ' 1. Historical Cashflows
    ' Updated SP: edw.usp_HistoricalCashflowReport_Aggregated
    modHistCashBuilder.BuildDynamicForm "Historical Cashflows", "Historical Cashflows", "edw.usp_HistoricalCashflowReport_Aggregated"
    
    ' 2. My Performance
    modHistCashBuilder.BuildDynamicForm "My Performance", "My Performance", "edw.pr_usp_investor_transactions3"
    
    ' 3. Portfolio Diversification
    modHistCashBuilder.BuildDynamicForm "Portfolio Diversification", "Portfolio Diversification", "sp_GetPortfolioDiversification"
    
    ' 4. Company Diversification
    modHistCashBuilder.BuildDynamicForm "Company Diversification", "Company Diversification", "edw.pr_usp_investor_transactions_cd"
    
    Application.ScreenUpdating = True
    MsgBox "All UserForms have been built successfully using the new Generic Builder!", vbInformation
End Sub

Private Function CheckSecurityAccess() As Boolean
    Dim vbp As Object
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbp Is Nothing Or vbp.Protection = 1 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' and unlock project.", vbCritical
        CheckSecurityAccess = False
    Else
        CheckSecurityAccess = True
    End If
End Function

' Legacy wrappers (optional, kept for compatibility if buttons invoke these directly)
Public Sub BuildHistCashflowsForm()
    modHistCashBuilder.BuildDynamicForm "Historical Cashflows", "Historical Cashflows", "edw.usp_HistoricalCashflowReport_Aggregated"
End Sub

Public Sub BuildMyPerformanceForm()
    modHistCashBuilder.BuildDynamicForm "My Performance", "My Performance", "edw.pr_usp_investor_transactions3"
End Sub

Public Sub BuildPortfolioDiversificationForm()
    modHistCashBuilder.BuildDynamicForm "Portfolio Diversification", "Portfolio Diversification", "sp_GetPortfolioDiversification"
End Sub

Public Sub BuildCompanyDiversificationForm()
    modHistCashBuilder.BuildDynamicForm "Company Diversification", "Company Diversification", "edw.pr_usp_investor_transactions_cd"
End Sub
