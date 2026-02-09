Attribute VB_Name = "modFormBuilder"
Option Explicit

' ==================================================================================
' FORM BUILDER SCRIPT (WRAPPER)
' ==================================================================================
' This script delegates to modHistCashBuilder.BuildDynamicForm for form generation.
' ==================================================================================

Public Sub BuildAllUserForms()
    If Not modFormHelpers.CheckSecurityAccess() Then Exit Sub
    Application.ScreenUpdating = False
    
    ' 1. Historical Cashflows
    ' Updated SP: edw.usp_HistoricalCashflowReport_Aggregated
    modHistCashBuilder.BuildDynamicForm "Historical Cashflows", "Historical Cashflows", "edw.usp_HistoricalCashflowReport_Aggregated"
    
    ' 2. My Performance
    modHistCashBuilder.BuildDynamicForm "My Performance", "My Performance", "edw.usp_InvestorPerformance_Aggregated"
    
    ' 3. Portfolio Diversification
    modHistCashBuilder.BuildDynamicForm "Portfolio Diversification", "Portfolio Diversification", "edw.usp_PortfolioDiversification_Aggregated"
    
    ' 4. Company Diversification
    modHistCashBuilder.BuildDynamicForm "Company Diversification", "Company Diversification", "edw.usp_CompanyDiversification_Aggregated"
    
    Application.ScreenUpdating = True
    MsgBox "All UserForms have been built successfully using the new Generic Builder!", vbInformation
End Sub



' Legacy wrappers (optional, kept for compatibility if buttons invoke these directly)
Public Sub BuildHistCashflowsForm()
    modHistCashBuilder.BuildDynamicForm "Historical Cashflows", "Historical Cashflows", "edw.usp_HistoricalCashflowReport_Aggregated"
End Sub

Public Sub BuildMyPerformanceForm()
    modHistCashBuilder.BuildDynamicForm "My Performance", "My Performance", "edw.usp_InvestorPerformance_Aggregated"
End Sub

Public Sub BuildPortfolioDiversificationForm()
    modHistCashBuilder.BuildDynamicForm "Portfolio Diversification", "Portfolio Diversification", "edw.usp_PortfolioDiversification_Aggregated"
End Sub

Public Sub BuildCompanyDiversificationForm()
    modHistCashBuilder.BuildDynamicForm "Company Diversification", "Company Diversification", "edw.usp_CompanyDiversification_Aggregated"
End Sub
