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
    
    ' 1. Unload from memory if currently running
    Dim i As Integer
    On Error Resume Next
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        If VBA.UserForms(i).Name = formName Then
            Unload VBA.UserForms(i)
        End If
    Next i
    On Error GoTo 0
    
    ' 2. Remove component from project
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
    If Not vbComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If
    On Error GoTo 0
    
    DoEvents ' Allow system to process removal
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
    

Private Sub AddLabelAndCombobox(formComp As Object, ByRef currentTop As Double, _
                                fieldName As String, labelCap As String, _
                                isMandatory As Boolean)
    Dim leftMargin As Double: leftMargin = 10
    Dim labelWidth As Double: labelWidth = 90
    Dim comboWidth As Double: comboWidth = 120
    
    ' Add Label
    Dim lbl As Object
    Set lbl = formComp.Designer.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = labelCap & IIf(isMandatory, " (*)", "")
        .Left = leftMargin
        .Top = currentTop + 3
        .Width = labelWidth
        .Height = 18
        .Name = "lbl" & Replace(fieldName, " ", "")
    End With
    
    ' Add ComboBox
    Dim cbo As Object
    Set cbo = formComp.Designer.Controls.Add("Forms.ComboBox.1")
    With cbo
        .Name = "cbo" & Replace(fieldName, " ", "")
        .Left = leftMargin + labelWidth + 5
        .Top = currentTop
        .Width = comboWidth
        .Height = 18
        .MatchEntry = 1 ' fmMatchEntryComplete - Auto-completion
        If isMandatory Then .BackColor = RGB(255, 192, 192)
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
    AddLabelAndCombobox uf, topPos, "Account", "Account", True: topPos = topPos + 24
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
    
    ' paramCode: Now contains lines like: cmd.Parameters.Append cmd.CreateParameter(...)
    
    Dim adoCode As String
    adoCode = ""
    ' @MetricName varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@MetricName"", 200, 1, 100, NullIfEmpty(Me.txtMetric.Value))" & vbCrLf
    ' @SourceTableVolVal varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@SourceTableVolVal"", 200, 1, 100, NullIfEmpty(Me.txtInvestor.Value))" & vbCrLf
    ' @StartDate date
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@StartDate"", 133, 1, , NullIfEmpty(Me.txtFromDates.Value))" & vbCrLf
    ' @EndDate date
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@EndDate"", 133, 1, , NullIfEmpty(Me.txtToDates.Value))" & vbCrLf
    ' @ViewCurrencyID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@ViewCurrencyID"", 3, 1, , NullIfEmpty(Me.txtViewingCoy.Value))" & vbCrLf
    ' @FilterSourceCurrency varchar(20)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@FilterSourceCurrency"", 200, 1, 20, NullIfEmpty(Me.txtCurrency.Value))" & vbCrLf
    ' @FundTypesExclude varchar(400)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@FundTypesExclude"", 200, 1, 400, NULL)" & vbCrLf
    ' @InvestorGroupID int (Using null for now)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@InvestorGroupID"", 3, 1, , NULL)" & vbCrLf
    ' @AIVFundGroupID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@AIVFundGroupID"", 3, 1, , NullIfEmpty(Me.txtAIVFundGroupID.Value))" & vbCrLf
    ' @MaxRows int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@MaxRows"", 3, 1, , 1000)" & vbCrLf
    ' @OrderBy varchar(20)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@OrderBy"", 200, 1, 20, ""date"")" & vbCrLf

    AddButtons uf, topPos, "edw.pr_usp_investor_transactions2", adoCode, _
        Array("Account", "Client", "Investor", "FromDates", "ToDates", "ViewingCoy", "Metric")

End Sub

' ==================================================================================
' FORM 2: My Performance
' ==================================================================================
Private Sub BuildMyPerformanceForm()
    Dim uf As Object
    Set uf = CreateUserForm("ufMyPerformance", "My Performance", 330, 380)
    
    Dim topPos As Double: topPos = 10
    
    AddLabelAndCombobox uf, topPos, "Account", "Account", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Client", "Client", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Investor", "Investor", True: topPos = topPos + 24
    ' Dates Removed
    AddLabelAndText uf, topPos, "ViewingCoy", "View Currency ID", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Metric", "Metric", True: topPos = topPos + 24
    
    AddLabelAndText uf, topPos, "Date", "Date", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Currency", "Currency Filter", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "AIVFundGroupID", "AIV Fund Grp ID", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "InvTxnQuarter", "Inv. Txn Quarter", False: topPos = topPos + 24
    
    ' Buttons
    ' UPDATED: Explicit ADODB Params
    Dim adoCode As String
    adoCode = ""
    ' @MetricName varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@MetricName"", 200, 1, 100, NullIfEmpty(Me.txtMetric.Value))" & vbCrLf
    ' @SourceTableVolVal varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@SourceTableVolVal"", 200, 1, 100, NullIfEmpty(Me.txtInvestor.Value))" & vbCrLf
    ' @Date date
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@Date"", 133, 1, , NullIfEmpty(Me.txtDate.Value))" & vbCrLf
    ' @ViewCurrencyID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@ViewCurrencyID"", 3, 1, , NullIfEmpty(Me.txtViewingCoy.Value))" & vbCrLf
    ' @FilterSourceCurrency varchar(20)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@FilterSourceCurrency"", 200, 1, 20, NullIfEmpty(Me.txtCurrency.Value))" & vbCrLf
    ' @InvestorGroupID int (Using null for now or mapped?) NULL passed default
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@InvestorGroupID"", 3, 1, , NULL)" & vbCrLf
    ' @AIVFundGroupID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@AIVFundGroupID"", 3, 1, , NullIfEmpty(Me.txtAIVFundGroupID.Value))" & vbCrLf
    
    AddButtons uf, topPos, "edw.pr_usp_investor_transactions3", adoCode, _
        Array("Account", "Client", "Investor", "ViewingCoy", "Metric", "Date")
    
End Sub

' ==================================================================================
' FORM 3: Portfolio Diversification
' ==================================================================================
Private Sub BuildPortfolioDiversificationForm()
    Dim uf As Object
    Set uf = CreateUserForm("ufPortfolioDiversification", "Portfolio Diversification", 650, 420)
    
    ' Column 1
    Dim topPos1 As Double: topPos1 = 10
    AddLabelAndCombobox uf, topPos1, "Account", "Account", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Client", "Client", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Investor", "Investor", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "ViewingCoy", "View Currency ID", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Metric", "Metric", True: topPos1 = topPos1 + 24
    
    AddLabelAndText uf, topPos1, "Date", "Date", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Currency", "Currency Filter", False: topPos1 = topPos1 + 24
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
    
    Dim adoCode As String
    adoCode = ""
    ' @SourceTableVolVal varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@SourceTableVolVal"", 200, 1, 100, NullIfEmpty(Me.txtInvestor.Value))" & vbCrLf
    ' @ViewCurrencyID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@ViewCurrencyID"", 3, 1, , NullIfEmpty(Me.txtViewingCoy.Value))" & vbCrLf
    ' @FilterSourceCurrency varchar(20)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@FilterSourceCurrency"", 200, 1, 20, NullIfEmpty(Me.txtCurrency.Value))" & vbCrLf
    ' @Metric varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@Metric"", 200, 1, 100, NullIfEmpty(Me.txtMetric.Value))" & vbCrLf
    ' @Date date
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@Date"", 133, 1, , NullIfEmpty(Me.txtDate.Value))" & vbCrLf
    ' @AIVFundGroupID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@AIVFundGroupID"", 3, 1, , NullIfEmpty(Me.txtAIVFundGroupID.Value))" & vbCrLf
    ' @EntryFund varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@EntryFund"", 200, 1, 100, NullIfEmpty(Me.txtEntryFund.Value))" & vbCrLf
    ' @Manager varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@Manager"", 200, 1, 100, NullIfEmpty(Me.txtManager.Value))" & vbCrLf
    ' @Portfolio varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@Portfolio"", 200, 1, 100, NullIfEmpty(Me.txtPortfolio.Value))" & vbCrLf
    ' @PortfolioCloseYear int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioCloseYear"", 3, 1, , NullIfEmpty(Me.txtPortCloseYear.Value))" & vbCrLf
    ' @PortfolioCommitmentYear int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioCommitmentYear"", 3, 1, , NullIfEmpty(Me.txtPortCommitYear.Value))" & vbCrLf
    ' @PortfolioGeography varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioGeography"", 200, 1, 100, NullIfEmpty(Me.txtPortGeo.Value))" & vbCrLf
    ' @PortfolioGeographyBroad varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioGeographyBroad"", 200, 1, 100, NullIfEmpty(Me.txtPortGeoBroad.Value))" & vbCrLf
    ' @PortfolioGeographyL3 varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioGeographyL3"", 200, 1, 100, NullIfEmpty(Me.txtPortGeoL3.Value))" & vbCrLf
    ' @PortfolioGeographyL5 varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioGeographyL5"", 200, 1, 100, NullIfEmpty(Me.txtPortGeoL5.Value))" & vbCrLf
    ' @PortfolioIndustry varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioIndustry"", 200, 1, 100, NullIfEmpty(Me.txtPortInd.Value))" & vbCrLf
    ' @PortfolioIndustryL1 varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioIndustryL1"", 200, 1, 100, NullIfEmpty(Me.txtPortIndL1.Value))" & vbCrLf
    ' @PortfolioStage varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioStage"", 200, 1, 100, NullIfEmpty(Me.txtPortStage.Value))" & vbCrLf
    ' @PortfolioStageBroad varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioStageBroad"", 200, 1, 100, NullIfEmpty(Me.txtPortStageBroad.Value))" & vbCrLf
    ' @PortfolioStatus varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioStatus"", 200, 1, 100, NullIfEmpty(Me.txtPortStatus.Value))" & vbCrLf
    ' @PortfolioTypeBroadID varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioTypeBroadID"", 200, 1, 100, NullIfEmpty(Me.txtPortTypeBroad.Value))" & vbCrLf
    ' @PortfolioVintageYear int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@PortfolioVintageYear"", 3, 1, , NullIfEmpty(Me.txtPortVintage.Value))" & vbCrLf

    AddButtons uf, finalTop, "sp_GetPortfolioDiversification", adoCode, _
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
    AddLabelAndCombobox uf, topPos1, "Account", "Account", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Client", "Client", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Investor", "Investor", True: topPos1 = topPos1 + 24
    ' Dates Removed
    AddLabelAndText uf, topPos1, "ViewingCoy", "View Currency ID", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Metric", "Metric", True: topPos1 = topPos1 + 24
    
    AddLabelAndText uf, topPos1, "Date", "Date", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Currency", "Currency Filter", False: topPos1 = topPos1 + 24
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
    
    Dim adoCode As String
    adoCode = ""
    ' @MetricName varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@MetricName"", 200, 1, 100, NullIfEmpty(Me.txtMetric.Value))" & vbCrLf
    ' @SourceTableVolVal varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@SourceTableVolVal"", 200, 1, 100, NullIfEmpty(Me.txtInvestor.Value))" & vbCrLf
    ' @Date date
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@Date"", 133, 1, , NullIfEmpty(Me.txtDate.Value))" & vbCrLf
    ' @ViewCurrencyID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@ViewCurrencyID"", 3, 1, , NullIfEmpty(Me.txtViewingCoy.Value))" & vbCrLf
    ' @FilterSourceCurrency varchar(20)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@FilterSourceCurrency"", 200, 1, 20, NullIfEmpty(Me.txtCurrency.Value))" & vbCrLf
    ' @FundTypesExclude varchar(400)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@FundTypesExclude"", 200, 1, 400, NULL)" & vbCrLf
    ' @InvestorGroupID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@InvestorGroupID"", 3, 1, , NULL)" & vbCrLf
    ' @AIVFundGroupID int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@AIVFundGroupID"", 3, 1, , NullIfEmpty(Me.txtAIVFundGroupID.Value))" & vbCrLf
    ' @MaxRows int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@MaxRows"", 3, 1, , 1000)" & vbCrLf
    ' @OrderBy varchar(20)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@OrderBy"", 200, 1, 20, ""date"")" & vbCrLf

    ' Company Filters
    ' @CompExpGeoBroad varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpGeoBroad"", 200, 1, 100, NullIfEmpty(Me.txtCoExpGeoBroad.Value))" & vbCrLf
    ' @CompExpGeoCountry varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpGeoCountry"", 200, 1, 100, NullIfEmpty(Me.txtCoExpGeoCount.Value))" & vbCrLf
    ' @CompExpIndId varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpIndId"", 200, 1, 100, NullIfEmpty(Me.txtCoExpIndID.Value))" & vbCrLf
    ' @CompExpIndBroad varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpIndBroad"", 200, 1, 100, NullIfEmpty(Me.txtCoExpIndBroad.Value))" & vbCrLf
    ' @CompExpIndCategory varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpIndCategory"", 200, 1, 100, NullIfEmpty(Me.txtCoExpIndCat.Value))" & vbCrLf
    ' @CompExpInvType varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpInvType"", 200, 1, 100, NullIfEmpty(Me.txtCoExpInvType.Value))" & vbCrLf
    ' @CompExpInvYear int
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpInvYear"", 3, 1, , NullIfEmpty(Me.txtCoExpInvYear.Value))" & vbCrLf
    ' @CompExpIsPublic bit - Using int (3) or boolean (11) ? bit is 11 or tinyint 17
    ' ADODB doesn't like adBoolean often, use adInteger (3) or adTinyInt (16) if mapped to int/bit.
    ' SP expects BIT. adBoolean (11) is best provided value is True/False.
    ' If TextBox, it's string. "True".
    ' Let's use adInteger (3) and assume backend casts, or adBoolean (11).
    ' Safest is adInteger if user inputs 0/1. If user inputs "Yes/No", hard.
    ' Assuming user inputs 0 or 1.
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpIsPublic"", 3, 1, , NullIfEmpty(Me.txtCoExpPublic.Value))" & vbCrLf
    ' @CompExpStage varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpStage"", 200, 1, 100, NullIfEmpty(Me.txtCoExpStage.Value))" & vbCrLf
    ' @CompExpStageBroad varchar(100)
    adoCode = adoCode & "    cmd.Parameters.Append cmd.CreateParameter(""@CompExpStageBroad"", 200, 1, 100, NullIfEmpty(Me.txtCoExpStageBroad.Value))" & vbCrLf
    
    AddButtons uf, finalTop, "edw.pr_usp_investor_transactions_cd", adoCode, _
        Array("Account", "Client", "Investor", "ViewingCoy", "MetricName", "Date")

End Sub

Private Sub AddButtons(formComp As Object, topPos As Double, spName As String, adoParamsCode As String, mandatoryFields As Variant)
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
    
    ' Helper: NullIfEmpty
    code = code & "Private Function NullIfEmpty(val As Variant) As Variant" & vbCrLf
    code = code & "    If IsNull(val) Or Trim(val & """") = """" Then" & vbCrLf
    code = code & "        NullIfEmpty = Null" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        NullIfEmpty = val" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf
    
    ' Cancel Logic
    code = code & "Private Sub btnCancel_Click()" & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' ==========================================
    ' Injected Logic for Account Dropdown & Lookup
    ' ==========================================
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim ws As Worksheet" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""database"")" & vbCrLf
    code = code & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    code = code & "    Dim lastRow As Long" & vbCrLf
    code = code & "    lastRow = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row" & vbCrLf
    code = code & "    If lastRow > 1 Then" & vbCrLf
    code = code & "        Me.cboAccount.List = ws.Range(""A2:A"" & lastRow).Value" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub cboAccount_Change()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim ws As Worksheet" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""database"")" & vbCrLf
    code = code & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    code = code & "    Dim idx As Variant" & vbCrLf
    code = code & "    idx = Application.Match(Me.cboAccount.Value, ws.Columns(""A""), 0)" & vbCrLf
    code = code & "    If IsError(idx) Then" & vbCrLf
    code = code & "        ' Me.txtClient.Value = """"" & vbCrLf
    code = code & "        ' Me.txtInvestor.Value = """"" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        Me.txtClient.Value = ws.Cells(idx, 4).Value  ' Column D: ClientShortName" & vbCrLf
    code = code & "        Me.txtInvestor.Value = ws.Cells(idx, 3).Value ' Column C: AccountID (INT)" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' Submit Logic
    code = code & "Private Sub btnSubmit_Click()" & vbCrLf
    ' Validation Code Block
    code = code & "    ' 1. Validate Mandatory Fields" & vbCrLf
    
    ' Dynamic Mandatory Validation
    Dim f As Variant
    For Each f In mandatoryFields
        If f = "Account" Then
            code = code & "    If Not ValidateMandatory(Me.cbo" & f & ", """ & f & """) Then Exit Sub" & vbCrLf
        Else
            code = code & "    If Not ValidateMandatory(Me.txt" & f & ", """ & f & """) Then Exit Sub" & vbCrLf
        End If
        If InStr(1, f, "Date", vbTextCompare) > 0 Then
            code = code & "    If Not ValidateDate(Me.txt" & f & ", """ & f & """) Then Exit Sub" & vbCrLf
        End If
    Next f
    
    code = code & "    On Error GoTo 0" & vbCrLf & vbCrLf
    
    ' Execution
    code = code & "    ' 2. Execute SP via ADODB Explicit Call" & vbCrLf
    code = code & "    Dim conn As Object, cmd As Object, rs As Object" & vbCrLf
    code = code & "    Set conn = modDatabase.GetConnection()" & vbCrLf
    code = code & "    If conn Is Nothing Then Exit Sub" & vbCrLf & vbCrLf

    code = code & "    Set cmd = CreateObject(""ADODB.Command"")" & vbCrLf
    code = code & "    With cmd" & vbCrLf
    code = code & "        .ActiveConnection = conn" & vbCrLf
    code = code & "        .CommandText = """ & spName & """" & vbCrLf
    code = code & "        .CommandType = 4 ' adCmdStoredProc" & vbCrLf
                
    ' Inject ADO Params
    code = code & "        ' --- Parameters ---" & vbCrLf
    code = code & adoParamsCode & vbCrLf
    
    code = code & "    End With" & vbCrLf & vbCrLf
    
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Set rs = cmd.Execute" & vbCrLf
    code = code & "    If Err.Number <> 0 Then" & vbCrLf
    code = code & "        MsgBox ""Error executing SP: "" & Err.Description, vbCritical" & vbCrLf
    code = code & "        conn.Close" & vbCrLf
    code = code & "        Exit Sub" & vbCrLf
    code = code & "    End On Error GoTo 0" & vbCrLf & vbCrLf
    
    ' Output
    code = code & "    ' 3. Output Results" & vbCrLf
    code = code & "    OutputToResults rs" & vbCrLf & vbCrLf
    code = code & "    ' Clean up" & vbCrLf
    code = code & "    Set cmd = Nothing" & vbCrLf
    code = code & "    conn.Close" & vbCrLf
    code = code & "    Set conn = Nothing" & vbCrLf & vbCrLf
    
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf
    
    InjectCode formComp, code
End Sub
