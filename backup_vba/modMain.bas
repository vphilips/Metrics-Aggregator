Attribute VB_Name = "modMain"
Option Explicit

' Macros to show the forms
' Assign these to Buttons on your Dashboard or invoke via Alt+F8

Public Sub ShowHistoricalCashflows()
    ufHistCashflows.Show
End Sub

Public Sub ShowMyPerformance()
    ufMyPerformance.Show
End Sub

Public Sub ShowPortfolioDiversification()
    ufPortfolioDiversification.Show
End Sub

Public Sub ShowCompanyDiversification()
    ufCompanyDiversification.Show
End Sub
