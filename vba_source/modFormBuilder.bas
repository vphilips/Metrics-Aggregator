Attribute VB_Name = "modFormBuilder"
Option Explicit

' ==================================================================================
' FORM BUILDER SCRIPT
' ==================================================================================
' This script programmatically creates the UserForms, Controls, and Event Code.
'
' PREREQUISITES:
' 1. "Trust access to the VBA project object model" must be ENABLED in Excel Trust Center.
' 2. Reference to "Microsoft Visual Basic for Applications Extensibility 5.3" is recommended
'    but we will use Late Binding to avoid manual reference steps where possible.
' ==================================================================================

Public Sub BuildAllUserForms()
    ' Master macro to build all 4 forms
    
    If Not CheckSecurityAccess() Then Exit Sub
    
    Application.ScreenUpdating = False
    
    BuildHistCashflowsForm
    BuildMyPerformanceForm
    BuildPortfolioDiversificationForm
    BuildCompanyDiversificationForm
    
    Application.ScreenUpdating = True
    
    MsgBox "All UserForms have been built successfully!", vbInformation
End Sub

Private Function CheckSecurityAccess() As Boolean
    Dim vbp As Object
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    On Error GoTo 0
    
    If vbp Is Nothing Then
        MsgBox "Unable to access VBProject. Please enable 'Trust access to the VBA project object model' in File > Options > Trust Center > Trust Center Settings > Macro Settings.", vbCritical
        CheckSecurityAccess = False
        Exit Function
    End If
    
    If vbp.Protection = 1 Then 
        MsgBox "The VBA Project is locked. Please unlock it before running the builder.", vbCritical
        CheckSecurityAccess = False
        Exit Function
    End If
    
    CheckSecurityAccess = True
End Function

Private Sub DeleteFormIfExists(formName As String)
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
    If Not vbComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If
    On Error GoTo 0
End Sub

Private Function CreateUserForm(formName As String, caption As String, width As Double, height As Double) As Object
    DeleteFormIfExists formName
    
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    vbComp.Name = formName
    vbComp.Properties("Caption") = caption
    vbComp.Properties("Width") = width
    vbComp.Properties("Height") = height
    
    Set CreateUserForm = vbComp
End Function

Private Sub AddLabelAndText(formComp As Object, ByRef currentTop As Double, _
                            fieldName As String, labelCap As String, _
                            isMandatory As Boolean, _
                            Optional colIndex As Integer = 0)
    
    ' colIndex: 0 = First Column, 1 = Second Column
    Dim leftMargin As Double
    Dim labelWidth As Double, textWidth As Double
    
    If colIndex = 0 Then
        leftMargin = 10
    Else
        leftMargin = 340 ' Shift second column right
    End If
    
    labelWidth = 90
    textWidth = 120 ' Slightly wider textboxes
    
    ' Add Label
    Dim lbl As Object
    Set lbl = formComp.Designer.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = labelCap & IIf(isMandatory, " (*)", "")
        .Left = leftMargin
        .Top = currentTop + 3
        .Width = labelWidth
        .Height = 18
        ' Name not strictly needed for logic, but cleaner
        .Name = "lbl" & Replace(fieldName, " ", "")
    End With
    
    ' Add TextBox
    Dim txt As Object
    Set txt = formComp.Designer.Controls.Add("Forms.TextBox.1")
    With txt
        .Name = "txt" & Replace(fieldName, " ", "")
        .Left = leftMargin + labelWidth + 5
        .Top = currentTop
        .Width = textWidth
        .Height = 18
        If isMandatory Then .BackColor = RGB(255, 192, 192) ' Light Red tint for mandatory
    End With
    
End Sub

Private Sub InjectCode(formComp As Object, codeText As String)
    formComp.CodeModule.AddFromString codeText
End Sub

' ==================================================================================
' FORM 1: Historical Cashflows
' ==================================================================================
Private Sub BuildHistCashflowsForm()
    Dim uf As Object
    Set uf = CreateUserForm("ufHistCashflows", "Historical Cashflows", 330, 380)
    
    Dim topPos As Double: topPos = 10
    
    ' Fields
    AddLabelAndText uf, topPos, "Account", "Account", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Client", "Client", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Investor", "Investor", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "FromDates", "From Date", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "ToDates", "To Date", True: topPos = topPos + 24
    ' Moved Currency to Optional
    AddLabelAndText uf, topPos, "ViewingCoy", "Viewing Currency", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Metric", "Metric", True: topPos = topPos + 24
    
    AddLabelAndText uf, topPos, "AIVFundGroupID", "AIV Fund Grp ID", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "InvestorTransactionDate", "Inv. Txn Date", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "InvestorTransactionQuarter", "Inv. Txn Quarter", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Currency", "Currency", False: topPos = topPos + 24
    
    ' Buttons
    ' UPDATED: Direct call to edw.pr_usp_investor_transactions2
    
    Dim paramCode As String
    paramCode = ""
    paramCode = paramCode & """@SourceTableVolVal"", Me.txtClient.Value, " & vbCrLf
    paramCode = paramCode & """@MetricName"", Me.txtMetric.Value, " & vbCrLf
    paramCode = paramCode & """@StartDate"", Me.txtFromDates.Value, " & vbCrLf
    paramCode = paramCode & """@EndDate"", Me.txtToDates.Value, " & vbCrLf
    paramCode = paramCode & """@ViewCurrencyCode"", Me.txtViewingCoy.Value, " & vbCrLf
    paramCode = paramCode & """@FilterSourceCurrency"", Me.txtCurrency.Value, " & vbCrLf
    paramCode = paramCode & """@InvestorGroupID"", Me.txtInvestor.Value, " & vbCrLf
    paramCode = paramCode & """@AIVFundGroupID"", Me.txtAIVFundGroupID.Value"

    AddButtons uf, topPos, "edw.pr_usp_investor_transactions2", paramCode, _
        Array("Account", "Client", "Investor", "FromDates", "ToDates", "ViewingCoy", "Metric")

End Sub

' ==================================================================================
' FORM 2: My Performance
' ==================================================================================
Private Sub BuildMyPerformanceForm()
    Dim uf As Object
    Set uf = CreateUserForm("ufMyPerformance", "My Performance", 330, 380)
    
    Dim topPos As Double: topPos = 10
    
    AddLabelAndText uf, topPos, "Account", "Account", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Client", "Client", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Investor", "Investor", True: topPos = topPos + 24
    ' Dates Removed
    AddLabelAndText uf, topPos, "ViewingCoy", "Viewing Currency", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Metric", "Metric", True: topPos = topPos + 24
    
    AddLabelAndText uf, topPos, "Date", "Date", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Currency", "Currency", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "AIVFundGroupID", "AIV Fund Grp ID", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "InvTxnQuarter", "Inv. Txn Quarter", False: topPos = topPos + 24
    
    ' Buttons
    ' UPDATED: Direct call to edw.pr_usp_investor_transactions2
    
    Dim paramCode As String
    paramCode = ""
    paramCode = paramCode & """@SourceTableVolVal"", Me.txtClient.Value, " & vbCrLf
    paramCode = paramCode & """@MetricName"", Me.txtMetric.Value, " & vbCrLf
    ' Dates Removed
    paramCode = paramCode & """@ViewCurrencyCode"", Me.txtViewingCoy.Value, " & vbCrLf
    paramCode = paramCode & """@FilterSourceCurrency"", Me.txtCurrency.Value, " & vbCrLf
    paramCode = paramCode & """@InvestorGroupID"", Me.txtInvestor.Value, " & vbCrLf
    paramCode = paramCode & """@AIVFundGroupID"", Me.txtAIVFundGroupID.Value"
    
    AddButtons uf, topPos, "edw.pr_usp_my_performance", paramCode, _
        Array("Account", "Client", "Investor", "ViewingCoy", "Metric", "Date")
    
End Sub

' ==================================================================================
' FORM 3: Portfolio Diversification
' ==================================================================================
Private Sub BuildPortfolioDiversificationForm()
    Dim uf As Object
    ' Wider for 2 columns - Increased Height for new fields
    Set uf = CreateUserForm("ufPortfolioDiversification", "Portfolio Diversification", 650, 420)
    
    ' Column 1
    Dim topPos1 As Double: topPos1 = 10
    AddLabelAndText uf, topPos1, "Account", "Account", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Client", "Client", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Investor", "Investor", True: topPos1 = topPos1 + 24
    ' Dates Removed
    AddLabelAndText uf, topPos1, "ViewingCoy", "Viewing Currency", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Metric", "Metric", True: topPos1 = topPos1 + 24
    
    AddLabelAndText uf, topPos1, "Date", "Date", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Currency", "Currency", False: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "AIVFundGroupID", "AIV Fund Grp ID", False: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "EntryFund", "Entry Fund", False: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Manager", "Manager", False: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Portfolio", "Portfolio", False: topPos1 = topPos1 + 24
    
    
    ' Column 2
    Dim topPos2 As Double: topPos2 = 10
    AddLabelAndText uf, topPos2, "PortCloseYear", "Close Year", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortCommitYear", "Commit. Year", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortGeo", "Geography", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortGeoBroad", "Geo Broad", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortGeoL3", "Geo L3", False, 1: topPos2 = topPos2 + 24
    ' New Field
    AddLabelAndText uf, topPos2, "PortGeoL5", "Geo L5", False, 1: topPos2 = topPos2 + 24
    
    AddLabelAndText uf, topPos2, "PortInd", "Industry", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortIndL1", "Industry L1", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortStage", "Stage", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortStageBroad", "Stage Broad", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortStatus", "Status", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortTypeBroad", "Type Broad ID", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "PortVintage", "Vintage Year", False, 1: topPos2 = topPos2 + 24
    
    Dim finalTop As Double
    finalTop = IIf(topPos1 > topPos2, topPos1, topPos2)
    
    Dim paramCode As String
    paramCode = ""
    paramCode = paramCode & """@Account"", Me.txtAccount.Value, " & vbCrLf
    paramCode = paramCode & """@Client"", Me.txtClient.Value, " & vbCrLf
    paramCode = paramCode & """@Investor"", Me.txtInvestor.Value, " & vbCrLf
    ' Dates Removed
    paramCode = paramCode & """@ViewCurrencyCode"", Me.txtViewingCoy.Value, " & vbCrLf
    paramCode = paramCode & """@FilterSourceCurrency"", Me.txtCurrency.Value, " & vbCrLf
    paramCode = paramCode & """@Metric"", Me.txtMetric.Value, " & vbCrLf
    paramCode = paramCode & """@Date"", Me.txtDate.Value, " & vbCrLf
    paramCode = paramCode & """@AIVFundGroupID"", Me.txtAIVFundGroupID.Value, " & vbCrLf
    paramCode = paramCode & """@EntryFund"", Me.txtEntryFund.Value, " & vbCrLf
    paramCode = paramCode & """@Manager"", Me.txtManager.Value, " & vbCrLf
    paramCode = paramCode & """@Portfolio"", Me.txtPortfolio.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioCloseYear"", Me.txtPortCloseYear.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioCommitmentYear"", Me.txtPortCommitYear.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioGeography"", Me.txtPortGeo.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioGeographyBroad"", Me.txtPortGeoBroad.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioGeographyL3"", Me.txtPortGeoL3.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioGeographyL5"", Me.txtPortGeoL5.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioIndustry"", Me.txtPortInd.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioIndustryL1"", Me.txtPortIndL1.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioStage"", Me.txtPortStage.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioStageBroad"", Me.txtPortStageBroad.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioStatus"", Me.txtPortStatus.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioTypeBroadID"", Me.txtPortTypeBroad.Value, " & vbCrLf
    paramCode = paramCode & """@PortfolioVintageYear"", Me.txtPortVintage.Value"

    AddButtons uf, finalTop, "sp_GetPortfolioDiversification", paramCode, _
        Array("Account", "Client", "Investor", "ViewingCoy", "Metric", "Date")

    
End Sub

' ==================================================================================
' FORM 4: Company Diversification
' ==================================================================================
Private Sub BuildCompanyDiversificationForm()
    Dim uf As Object
    Set uf = CreateUserForm("ufCompanyDiversification", "Company Diversification", 650, 400)
    
    ' Column 1
    Dim topPos1 As Double: topPos1 = 10
    AddLabelAndText uf, topPos1, "Account", "Account", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Client", "Client", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Investor", "Investor", True: topPos1 = topPos1 + 24
    ' Dates Removed
    AddLabelAndText uf, topPos1, "ViewingCoy", "Viewing Currency", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Metric", "Metric", True: topPos1 = topPos1 + 24
    
    AddLabelAndText uf, topPos1, "Date", "Date", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Currency", "Currency", False: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "AIVFundGroupID", "AIV Fund Grp ID", False: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "CoExpGeoBroad", "Co Exp Geo Broad", False: topPos1 = topPos1 + 24
    
    ' Column 2
    Dim topPos2 As Double: topPos2 = 10
    AddLabelAndText uf, topPos2, "CoExpGeoCount", "Co Exp Geo Ctry", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "CoExpIndID", "Co Exp Ind ID", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "CoExpIndBroad", "Co Exp Ind Broad", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "CoExpIndCat", "Co Exp Ind Cat", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "CoExpInvType", "Co Exp Inv Type", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "CoExpInvYear", "Co Exp Inv Year", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "CoExpPublic", "Co Exp Public", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "CoExpStage", "Co Exp Stage", False, 1: topPos2 = topPos2 + 24
    AddLabelAndText uf, topPos2, "CoExpStageBroad", "Co Exp Stg Broad", False, 1: topPos2 = topPos2 + 24
    
    Dim finalTop As Double
    finalTop = IIf(topPos1 > topPos2, topPos1, topPos2)
    
    Dim paramCode As String
    paramCode = ""
    paramCode = paramCode & """@Account"", Me.txtAccount.Value, " & vbCrLf
    paramCode = paramCode & """@Client"", Me.txtClient.Value, " & vbCrLf
    paramCode = paramCode & """@Investor"", Me.txtInvestor.Value, " & vbCrLf
    ' Dates Removed
    paramCode = paramCode & """@ViewCurrencyCode"", Me.txtViewingCoy.Value, " & vbCrLf
    paramCode = paramCode & """@FilterSourceCurrency"", Me.txtCurrency.Value, " & vbCrLf
    paramCode = paramCode & """@Metric"", Me.txtMetric.Value, " & vbCrLf
    paramCode = paramCode & """@Date"", Me.txtDate.Value, " & vbCrLf
    paramCode = paramCode & """@AIVFundGroupID"", Me.txtAIVFundGroupID.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpGeoBroad"", Me.txtCoExpGeoBroad.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpGeoCountry"", Me.txtCoExpGeoCount.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpIndId"", Me.txtCoExpIndID.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpIndBroad"", Me.txtCoExpIndBroad.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpIndCategory"", Me.txtCoExpIndCat.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpInvType"", Me.txtCoExpInvType.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpInvYear"", Me.txtCoExpInvYear.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpIsPublic"", Me.txtCoExpPublic.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpStage"", Me.txtCoExpStage.Value, " & vbCrLf
    paramCode = paramCode & """@CompExpStageBroad"", Me.txtCoExpStageBroad.Value"

    AddButtons uf, finalTop, "sp_GetCompanyDiversification", paramCode, _
        Array("Account", "Client", "Investor", "ViewingCoy", "Metric", "Date")

End Sub

Private Sub AddButtons(formComp As Object, topPos As Double, spName As String, paramCode As String, mandatoryFields As Variant)
    ' Add Submit and Cancel buttons and associated code
    
    Dim btnSubmit As Object, btnCancel As Object
    
    Set btnSubmit = formComp.Designer.Controls.Add("Forms.CommandButton.1")
    With btnSubmit
        .Caption = "Submit"
        .Name = "btnSubmit"
        .Left = 100
        .Top = topPos + 10
        .Width = 60
        .Height = 24
        .Default = True ' Trigger on Enter
    End With
    
    Set btnCancel = formComp.Designer.Controls.Add("Forms.CommandButton.1")
    With btnCancel
        .Caption = "Cancel"
        .Name = "btnCancel"
        .Left = 170
        .Top = topPos + 10
        .Width = 60
        .Height = 24
        .Cancel = True ' Trigger on Esc
    End With
    
    ' Inject Code
    Dim code As String
    code = "Option Explicit" & vbCrLf & vbCrLf
    
    ' Cancel Logic
    code = code & "Private Sub btnCancel_Click()" & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' Submit Logic
    code = code & "Private Sub btnSubmit_Click()" & vbCrLf
    ' Validation Code Block
    code = code & "    ' 1. Validate Mandatory Fields" & vbCrLf
    
    ' Dynamic Mandatory Validation
    Dim f As Variant
    For Each f In mandatoryFields
        code = code & "    If Not ValidateMandatory(Me.txt" & f & ", """ & f & """) Then Exit Sub" & vbCrLf
        If InStr(1, f, "Date", vbTextCompare) > 0 Then
            code = code & "    If Not ValidateDate(Me.txt" & f & ", """ & f & """) Then Exit Sub" & vbCrLf
        End If
    Next f
    
    code = code & "    On Error GoTo 0" & vbCrLf & vbCrLf
    
    ' Params Building
    code = code & "    ' 2. Prepare Parameters" & vbCrLf
    code = code & "    Dim pNames As Variant, pValues As Variant" & vbCrLf
    code = code & "    BuildParams pNames, pValues, _" & vbCrLf
    code = code & "        " & paramCode & vbCrLf & vbCrLf
    
    ' Execution
    code = code & "    ' 3. Execute SP" & vbCrLf
    code = code & "    Dim rs As Object" & vbCrLf
    code = code & "    Set rs = ExecuteSP(""" & spName & """, pNames, pValues)" & vbCrLf & vbCrLf
    
    ' Output
    code = code & "    ' 4. Output Results" & vbCrLf
    code = code & "    OutputToResults rs" & vbCrLf & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf
    
    InjectCode formComp, code
End Sub
