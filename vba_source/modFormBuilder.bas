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
    ' Investor now Mandatory
    AddLabelAndText uf, topPos, "Investor", "Investor", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "FromDates", "From Dates", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "ToDates", "To Dates", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Currency", "Currency", True: topPos = topPos + 24
    ' New Fields
    AddLabelAndText uf, topPos, "ViewingCoy", "Viewing Coy", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Metric", "Metric", True: topPos = topPos + 24
    
    AddLabelAndText uf, topPos, "AIVFundGroupD", "AIV Fund Grp D", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "InvestorTransactionDate", "Inv. Txn Date", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "InvestorTransactionQuarter", "Inv. Txn Quarter", False: topPos = topPos + 24
    
    ' Buttons
    AddButtons uf, topPos, "sp_GetHistoricalCashflows", _
        """@Account"", Me.txtAccount.Value, " & vbCrLf & _
        """@Client"", Me.txtClient.Value, " & vbCrLf & _
        """@Investor"", Me.txtInvestor.Value, " & vbCrLf & _
        """@FromDate"", Me.txtFromDates.Value, " & vbCrLf & _
        """@ToDate"", Me.txtToDates.Value, " & vbCrLf & _
        """@Currency"", Me.txtCurrency.Value, " & vbCrLf & _
        """@ViewingCoy"", Me.txtViewingCoy.Value, " & vbCrLf & _
        """@Metric"", Me.txtMetric.Value, " & vbCrLf & _
        """@AIVFundGroupD"", Me.txtAIVFundGroupD.Value, " & vbCrLf & _
        """@InvestorTransactionDate"", Me.txtInvestorTransactionDate.Value, " & vbCrLf & _
        """@InvestorTransactionQuarter"", Me.txtInvestorTransactionQuarter.Value", _
        Array("Account", "Client", "Investor", "FromDates", "ToDates", "Currency", "ViewingCoy", "Metric")

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
    ' Date now Mandatory
    AddLabelAndText uf, topPos, "Date", "Date", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "FromDates", "From Dates", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "ToDates", "To Dates", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Currency", "Currency", True: topPos = topPos + 24
    ' New Fields
    AddLabelAndText uf, topPos, "ViewingCoy", "Viewing Coy", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Metric", "Metric", True: topPos = topPos + 24
    
    AddLabelAndText uf, topPos, "AIVFundGroupD", "AIV Fund Grp D", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Investor", "Investor", False: topPos = topPos + 24
    AddLabelAndText uf, topPos, "InvTxnQuarter", "Inv. Txn Quarter", False: topPos = topPos + 24
    
    AddButtons uf, topPos, "sp_GetMyPerformance", _
        """@Account"", Me.txtAccount.Value, " & vbCrLf & _
        """@Client"", Me.txtClient.Value, " & vbCrLf & _
        """@Date"", Me.txtDate.Value, " & vbCrLf & _
        """@FromDate"", Me.txtFromDates.Value, " & vbCrLf & _
        """@ToDate"", Me.txtToDates.Value, " & vbCrLf & _
        """@Currency"", Me.txtCurrency.Value, " & vbCrLf & _
        """@ViewingCoy"", Me.txtViewingCoy.Value, " & vbCrLf & _
        """@Metric"", Me.txtMetric.Value, " & vbCrLf & _
        """@AIVFundGroupD"", Me.txtAIVFundGroupD.Value, " & vbCrLf & _
        """@Investor"", Me.txtInvestor.Value, " & vbCrLf & _
        """@InvestorTransactionQuarter"", Me.txtInvTxnQuarter.Value", _
        Array("Account", "Client", "Date", "FromDates", "ToDates", "Currency", "ViewingCoy", "Metric")
    
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
    ' Date now Mandatory
    AddLabelAndText uf, topPos1, "Date", "Date", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "FromDates", "From Dates", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "ToDates", "To Dates", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Currency", "Currency", True: topPos1 = topPos1 + 24
    ' New Fields
    AddLabelAndText uf, topPos1, "ViewingCoy", "Viewing Coy", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Metric", "Metric", True: topPos1 = topPos1 + 24
    
    AddLabelAndText uf, topPos1, "AIVFundGroupD", "AIV Fund Grp D", False: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Investor", "Investor", False: topPos1 = topPos1 + 24
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
    
    AddButtons uf, finalTop, "sp_GetPortfolioDiversification", _
        """@Account"", Me.txtAccount.Value, " & vbCrLf & _
        """@Client"", Me.txtClient.Value, " & vbCrLf & _
        """@Date"", Me.txtDate.Value, " & vbCrLf & _
        """@FromDates"", Me.txtFromDates.Value, " & vbCrLf & _
        """@ToDates"", Me.txtToDates.Value, " & vbCrLf & _
        """@Currency"", Me.txtCurrency.Value, " & vbCrLf & _
        """@ViewingCoy"", Me.txtViewingCoy.Value, " & vbCrLf & _
        """@Metric"", Me.txtMetric.Value, " & vbCrLf & _
        """@AIVFundGroupD"", Me.txtAIVFundGroupD.Value, " & vbCrLf & _
        """@Investor"", Me.txtInvestor.Value, " & vbCrLf & _
        """@EntryFund"", Me.txtEntryFund.Value, " & vbCrLf & _
        """@Manager"", Me.txtManager.Value, " & vbCrLf & _
        """@Portfolio"", Me.txtPortfolio.Value, " & vbCrLf & _
        """@PortfolioCloseYear"", Me.txtPortCloseYear.Value, " & vbCrLf & _
        """@PortfolioCommitmentYear"", Me.txtPortCommitYear.Value, " & vbCrLf & _
        """@PortfolioGeography"", Me.txtPortGeo.Value, " & vbCrLf & _
        """@PortfolioGeographyBroad"", Me.txtPortGeoBroad.Value, " & vbCrLf & _
        """@PortfolioGeographyL3"", Me.txtPortGeoL3.Value, " & vbCrLf & _
        """@PortfolioGeographyL5"", Me.txtPortGeoL5.Value, " & vbCrLf & _
        """@PortfolioIndustry"", Me.txtPortInd.Value, " & vbCrLf & _
        """@PortfolioIndustryL1"", Me.txtPortIndL1.Value, " & vbCrLf & _
        """@PortfolioStage"", Me.txtPortStage.Value, " & vbCrLf & _
        """@PortfolioStageBroad"", Me.txtPortStageBroad.Value, " & vbCrLf & _
        """@PortfolioStatus"", Me.txtPortStatus.Value, " & vbCrLf & _
        """@PortfolioTypeBroadID"", Me.txtPortTypeBroad.Value, " & vbCrLf & _
        """@PortfolioVintageYear"", Me.txtPortVintage.Value", _
        Array("Account", "Client", "Date", "FromDates", "ToDates", "Currency", "ViewingCoy", "Metric")
    
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
    ' Date now Mandatory
    AddLabelAndText uf, topPos1, "Date", "Date", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "FromDates", "From Dates", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "ToDates", "To Dates", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Currency", "Currency", True: topPos1 = topPos1 + 24
    ' New Fields
    AddLabelAndText uf, topPos1, "ViewingCoy", "Viewing Coy", True: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Metric", "Metric", True: topPos1 = topPos1 + 24
    
    AddLabelAndText uf, topPos1, "AIVFundGroupD", "AIV Fund Grp D", False: topPos1 = topPos1 + 24
    AddLabelAndText uf, topPos1, "Investor", "Investor", False: topPos1 = topPos1 + 24
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
    
    AddButtons uf, finalTop, "sp_GetCompanyDiversification", _
        """@Account"", Me.txtAccount.Value, " & vbCrLf & _
        """@Client"", Me.txtClient.Value, " & vbCrLf & _
        """@Date"", Me.txtDate.Value, " & vbCrLf & _
        """@FromDates"", Me.txtFromDates.Value, " & vbCrLf & _
        """@ToDates"", Me.txtToDates.Value, " & vbCrLf & _
        """@Currency"", Me.txtCurrency.Value, " & vbCrLf & _
        """@ViewingCoy"", Me.txtViewingCoy.Value, " & vbCrLf & _
        """@Metric"", Me.txtMetric.Value, " & vbCrLf & _
        """@AIVFundGroupD"", Me.txtAIVFundGroupD.Value, " & vbCrLf & _
        """@Investor"", Me.txtInvestor.Value, " & vbCrLf & _
        """@CompExpGeoBroad"", Me.txtCoExpGeoBroad.Value, " & vbCrLf & _
        """@CompExpGeoCountry"", Me.txtCoExpGeoCount.Value, " & vbCrLf & _
        """@CompExpIndId"", Me.txtCoExpIndID.Value, " & vbCrLf & _
        """@CompExpIndBroad"", Me.txtCoExpIndBroad.Value, " & vbCrLf & _
        """@CompExpIndCategory"", Me.txtCoExpIndCat.Value, " & vbCrLf & _
        """@CompExpInvType"", Me.txtCoExpInvType.Value, " & vbCrLf & _
        """@CompExpInvYear"", Me.txtCoExpInvYear.Value, " & vbCrLf & _
        """@CompExpIsPublic"", Me.txtCoExpPublic.Value, " & vbCrLf & _
        """@CompExpStage"", Me.txtCoExpStage.Value, " & vbCrLf & _
        """@CompExpStageBroad"", Me.txtCoExpStageBroad.Value", _
        Array("Account", "Client", "Date", "FromDates", "ToDates", "Currency", "ViewingCoy", "Metric")

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
