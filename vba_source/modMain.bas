Attribute VB_Name = "modMain"
Option Explicit

' Macros to show the forms
' Assign these to Buttons on your Dashboard or invoke via Alt+F8


Public Sub ShowHistoricalCashflows()
    frmHistoricalCashflows.Show
End Sub

Public Sub ShowMyPerformance()
    frmMyPerformance.Show
End Sub

Public Sub ShowPortfolioDiversification()
    frmPortfolioDiversification.Show
End Sub

Public Sub ShowCompanyDiversification()
    frmCompanyDiversification.Show
End Sub
