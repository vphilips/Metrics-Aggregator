Attribute VB_Name = "modHistCashBuilder"
Option Explicit

' ==================================================================================
' DYNAMIC FORM BUILDER (CASCADING + GENERIC)
' ==================================================================================
' Programmatically creates forms driven by "form_config" (cascading) 
' and "form_metrics" (variable list).
' ==================================================================================

Public Sub BuildDynamicForm(ByVal FormType As String, ByVal FormCaption As String, ByVal SPName As String)
    If Not modFormHelpers.CheckSecurityAccess() Then Exit Sub
    Application.ScreenUpdating = False
    
    Dim fName As String
    fName = "frm" & Replace(FormType, " ", "")
    DeleteFormIfExists fName
    
    Dim uf As Object
    Set uf = CreateUserForm(fName, FormCaption, 450, 500)
    
    Dim topPos As Double: topPos = 10
    
    ' 1. Logic Fields (Cascading)
    AddLabelAndCombobox uf, topPos, "VariableName", "Variable Name", True: topPos = topPos + 24
    AddLabelAndCombobox uf, topPos, "Client", "Client", True: topPos = topPos + 24
    
    ' 2. Account ListBox
    AddLabelAndListBox uf, topPos, "Accounts", "Account(s)", True, 120
    topPos = topPos + 120 + 24
    
    ' 3. Auto-Filled Fields
    ' Start/End Date vs Asof Date
    If FormType = "Historical Cashflows" Then
        AddLabelAndText uf, topPos, "FromDate", "From Date", True: topPos = topPos + 24
        AddLabelAndText uf, topPos, "ToDate", "To Date", True: topPos = topPos + 24
    Else
        AddLabelAndText uf, topPos, "AsofDate", "Asof Date", True: topPos = topPos + 24
    End If
    
    AddLabelAndText uf, topPos, "Currency", "Currency", True: topPos = topPos + 24
    
    ' 4. Inject Logic
    InjectCascadingLogic uf, SPName, FormType
    
    Application.ScreenUpdating = True
    MsgBox fName & " has been created!", vbInformation
End Sub

' ==================================================================================
' HELPERS
' ==================================================================================


Private Sub DeleteFormIfExists(formName As String)
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
    If Not vbComp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove vbComp
    On Error GoTo 0
End Sub

Private Function CreateUserForm(formName As String, caption As String, width As Double, height As Double) As Object
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3)
    vbComp.Name = formName
    vbComp.Properties("Caption") = caption
    vbComp.Properties("Width") = width
    vbComp.Properties("Height") = height
    Set CreateUserForm = vbComp
End Function

Private Sub AddLabelAndText(formComp As Object, ByRef currentTop As Double, fieldName As String, labelCap As String, isMandatory As Boolean)
    Dim lbl As Object, txt As Object
    Set lbl = formComp.Designer.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = labelCap & IIf(isMandatory, " (*)", ""): .Left = 10: .Top = currentTop + 3: .Width = 100: .Height = 18
    End With
    Set txt = formComp.Designer.Controls.Add("Forms.TextBox.1")
    With txt
        .Name = "txt" & fieldName: .Left = 115: .Top = currentTop: .Width = 300: .Height = 18
    End With
End Sub

Private Sub AddLabelAndCombobox(formComp As Object, ByRef currentTop As Double, fieldName As String, labelCap As String, isMandatory As Boolean)
    Dim lbl As Object, cbo As Object
    Set lbl = formComp.Designer.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = labelCap & IIf(isMandatory, " (*)", ""): .Left = 10: .Top = currentTop + 3: .Width = 100: .Height = 18
    End With
    Set cbo = formComp.Designer.Controls.Add("Forms.ComboBox.1")
    With cbo
        .Name = "cbo" & fieldName: .Left = 115: .Top = currentTop: .Width = 300: .Height = 18
    End With
End Sub

Private Sub AddLabelAndListBox(formComp As Object, ByRef currentTop As Double, fieldName As String, labelCap As String, isMandatory As Boolean, height As Double)
    Dim lbl As Object, chk As Object, lst As Object
    Set lbl = formComp.Designer.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = labelCap & IIf(isMandatory, " (*)", ""): .Left = 10: .Top = currentTop: .Width = 100: .Height = 18
    End With
    Set chk = formComp.Designer.Controls.Add("Forms.CheckBox.1")
    With chk
        .Caption = "Select All": .Name = "chkSelectAll" & fieldName: .Left = 115: .Top = currentTop: .Width = 80: .Height = 18: .Font.Size = 8
    End With
    Set lst = formComp.Designer.Controls.Add("Forms.ListBox.1")
    With lst
        .Name = "lst" & fieldName: .Left = 115: .Top = currentTop + 20: .Width = 300: .Height = height
        .MultiSelect = 1 ' fmMultiSelectMulti
        .ListStyle = 1   ' fmListStyleOption
    End With
End Sub

Private Sub InjectCascadingLogic(formComp As Object, spName As String, FormType As String)
    Dim topPos As Double
    topPos = 120 + 24 + 24 + 24 + 24 + 50 ' approximate bottom
    
    ' Buttons
    Dim btnSubmit As Object, btnCancel As Object
    Set btnSubmit = formComp.Designer.Controls.Add("Forms.CommandButton.1")
    With btnSubmit: .Caption = "Submit": .Name = "btnSubmit": .Left = 150: .Top = topPos: .Width = 80: .Height = 24: .Default = True: End With
    Set btnCancel = formComp.Designer.Controls.Add("Forms.CommandButton.1")
    With btnCancel: .Caption = "Cancel": .Name = "btnCancel": .Left = 240: .Top = topPos: .Width = 80: .Height = 24: .Cancel = True: End With

    Dim code As String
    code = "Option Explicit" & vbCrLf
    code = code & "Private m_AttributesStr As String" & vbCrLf & vbCrLf
    
    ' --- INIT: Load Variable Names from form_metrics ---
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    LoadVariableNames" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub LoadVariableNames()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim ws As Worksheet, rng As Range, cell As Range" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""form_metrics"")" & vbCrLf
    code = code & "    If ws Is Nothing Then MsgBox ""Sheet 'form_metrics' not found!"": Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    ' Logic to find the correct column based on FormType
    code = code & "    Dim headerRng As Range, found As Range, colIndex As Long" & vbCrLf
    code = code & "    Set headerRng = ws.Range(""1:1"")" & vbCrLf
    If FormType = "Company Diversification" Then
        ' Handle known typo in Excel sheet header: "Company Diverisification"
        code = code & "    Set found = headerRng.Find(What:=""Company Diverisification"", LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
        ' Fallback to correct spelling just in case it gets fixed later
        code = code & "    If found Is Nothing Then Set found = headerRng.Find(What:=""Company Diversification"", LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    Else
        code = code & "    Set found = headerRng.Find(What:=""" & FormType & """, LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    End If
    code = code & "    If found Is Nothing Then MsgBox ""Column '" & FormType & "' not found in form_metrics!"": Exit Sub" & vbCrLf
    code = code & "    colIndex = found.Column" & vbCrLf
    code = code & "    " & vbCrLf
    
    code = code & "    Dim lastRow As Long" & vbCrLf
    code = code & "    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row" & vbCrLf
    code = code & "    If lastRow < 2 Then Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    
    code = code & "    For Each cell In ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex))" & vbCrLf
    code = code & "        If SafeStr(cell.Value) <> """" Then" & vbCrLf
    code = code & "            Me.cboVariableName.AddItem SafeStr(cell.Value)" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next cell" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' --- CASCADE 1: Variable -> Client ---
    code = code & "Private Sub cboVariableName_Change()" & vbCrLf
    code = code & "    Me.cboClient.Clear" & vbCrLf
    code = code & "    Me.lstAccounts.Clear" & vbCrLf
    
    ' Ensure fields exist before clearing
    If FormType = "Historical Cashflows" Then
        code = code & "    Me.txtFromDate.Value = """"" & vbCrLf
        code = code & "    Me.txtToDate.Value = """"" & vbCrLf
    Else
        code = code & "    Me.txtAsofDate.Value = """"" & vbCrLf
    End If
    
    code = code & "    Me.txtCurrency.Value = """"" & vbCrLf
    code = code & "    m_AttributesStr = """"" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' Mapping relies on 'form_config' for cascading" & vbCrLf
    code = code & "    Dim ws As Worksheet, lastRow As Long, i As Long" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""form_config"")" & vbCrLf
    code = code & "    lastRow = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    For i = 2 To lastRow" & vbCrLf
    code = code & "        If SafeStr(ws.Cells(i, 1).Value) = Me.cboVariableName.Value Then" & vbCrLf
    code = code & "            Me.cboClient.AddItem SafeStr(ws.Cells(i, 2).Value) ' Col B" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' --- CASCADE 2: Client -> Details ---
    code = code & "Private Sub cboClient_Change()" & vbCrLf
    code = code & "    Me.lstAccounts.Clear" & vbCrLf
    code = code & "    m_AttributesStr = """"" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    Dim ws As Worksheet, lastRow As Long, i As Long" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""form_config"")" & vbCrLf
    code = code & "    lastRow = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' Aggregate from ALL matching rows" & vbCrLf
    code = code & "    For i = 2 To lastRow" & vbCrLf
    code = code & "        If SafeStr(ws.Cells(i, 1).Value) = Me.cboVariableName.Value And SafeStr(ws.Cells(i, 2).Value) = Me.cboClient.Value Then" & vbCrLf
    code = code & "            " & vbCrLf
    code = code & "            ' 1. Populate Accounts (Aggregation)" & vbCrLf
    code = code & "            Dim accRaw As Variant" & vbCrLf
    code = code & "            accRaw = ws.Cells(i, 3).Value" & vbCrLf
    code = code & "            If Not IsError(accRaw) And Not IsEmpty(accRaw) Then" & vbCrLf
    code = code & "                Dim parts() As String" & vbCrLf
    code = code & "                Dim cleanRaw As String" & vbCrLf
    code = code & "                cleanRaw = CStr(accRaw)" & vbCrLf
    code = code & "                parts = Split(cleanRaw, "";"")" & vbCrLf
    code = code & "                Dim p As Variant" & vbCrLf
    code = code & "                For Each p In parts" & vbCrLf
    code = code & "                    Dim cleanVal As String" & vbCrLf
    code = code & "                    cleanVal = p" & vbCrLf
    code = code & "                    ' Handle potential non-breaking spaces (common in Excel data)" & vbCrLf
    code = code & "                    cleanVal = Replace(cleanVal, Chr(160), "" "")" & vbCrLf
    code = code & "                    ' Standard Trim handles leading/trailing regular spaces" & vbCrLf
    code = code & "                    cleanVal = Trim(cleanVal)" & vbCrLf
    code = code & "                    " & vbCrLf
    code = code & "                    If Len(cleanVal) > 0 Then Me.lstAccounts.AddItem cleanVal" & vbCrLf
    code = code & "                Next p" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "            " & vbCrLf
    code = code & "            ' 2. Update Details (Last match updates value)" & vbCrLf
    
    If FormType = "Historical Cashflows" Then
        code = code & "            Me.txtFromDate.Value = SafeStr(ws.Cells(i, 5).Value)" & vbCrLf
        code = code & "            Me.txtToDate.Value = SafeStr(ws.Cells(i, 6).Value)" & vbCrLf
    Else
        code = code & "            Me.txtAsofDate.Value = SafeStr(ws.Cells(i, 4).Value)" & vbCrLf
    End If
    
    code = code & "            Me.txtCurrency.Value = SafeStr(ws.Cells(i, 8).Value) ' Col H" & vbCrLf
    code = code & "            m_AttributesStr = SafeStr(ws.Cells(i, 7).Value) ' Col G" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' --- Select All ---
    code = code & "Private Sub chkSelectAllAccounts_Click()" & vbCrLf
    code = code & "    Dim i As Long" & vbCrLf
    code = code & "    For i = 0 To Me.lstAccounts.ListCount - 1" & vbCrLf
    code = code & "        Me.lstAccounts.Selected(i) = Me.chkSelectAllAccounts.Value" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' --- SUBMIT ---
    code = code & "Private Sub btnSubmit_Click()" & vbCrLf
    code = code & "    ' Validate" & vbCrLf
    code = code & "    If Me.cboVariableName.Value = """" Then MsgBox ""Variable required"": Exit Sub" & vbCrLf
    code = code & "    If Me.cboClient.Value = """" Then MsgBox ""Client required"": Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' 1. Get Metric Name" & vbCrLf
    code = code & "    Dim finalMetric As String" & vbCrLf
    code = code & "    finalMetric = GetMetricFromVariable(Me.cboVariableName.Value)" & vbCrLf
    code = code & "    If finalMetric = """" Then Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' 2. Get Fund IDs (Column D of database sheet)" & vbCrLf
    code = code & "    Dim hasSelection As Boolean: hasSelection = False" & vbCrLf
    code = code & "    Dim k As Long" & vbCrLf
    code = code & "    For k = 0 To Me.lstAccounts.ListCount - 1" & vbCrLf
    code = code & "        If Me.lstAccounts.Selected(k) Then hasSelection = True: Exit For" & vbCrLf
    code = code & "    Next k" & vbCrLf
    code = code & "    If Not hasSelection Then MsgBox ""Select at least one Account"": Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    Dim strFundIds As String" & vbCrLf
    code = code & "    strFundIds = GetFundIds(Me.lstAccounts)" & vbCrLf
    code = code & "    If strFundIds = """" Then Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' 3. Get Client Global ID (Column B of database sheet)" & vbCrLf
    code = code & "    Dim strClientGlobalID As String" & vbCrLf
    code = code & "    strClientGlobalID = GetClientGlobalID(Me.cboClient.Value)" & vbCrLf
    code = code & "    If strClientGlobalID = """" Then MsgBox ""Could not find Client Global ID"": Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    
    code = code & "    ' Execute" & vbCrLf
    code = code & "    Dim conn As Object, cmd As Object, rs As Object" & vbCrLf
    code = code & "    Set conn = modDatabase.GetConnection()" & vbCrLf
    code = code & "    If conn Is Nothing Then Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    Set cmd = CreateObject(""ADODB.Command"")" & vbCrLf
    code = code & "    With cmd" & vbCrLf
    code = code & "        .ActiveConnection = conn" & vbCrLf
    code = code & "        .CommandText = """ & spName & """" & vbCrLf
    code = code & "        .CommandType = 4" & vbCrLf
    code = code & "        " & vbCrLf
    
    If FormType = "Historical Cashflows" Then
        ' SP: edw.usp_HistoricalCashflowReport_Aggregated
        
        code = code & "        .NamedParameters = True" & vbCrLf

        code = code & "        .Parameters.Append .CreateParameter(""@Metric"", 200, 1, 200, NullIfEmpty(finalMetric))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorNameIdsCsv"", 200, 1, 255, NullIfEmpty(strClientGlobalID))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorIdsCsv"", 200, 1, -1, Null)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@FundIdsCsv"", 200, 1, -1, NullIfEmpty(strFundIds))" & vbCrLf
        
        ' Pass Dates as Strings to let SQL handle conversion (avoid int/date clash)
        code = code & "        .Parameters.Append .CreateParameter(""@StartDate"", 200, 1, 20, Format(ParseDateForDB(Me.txtFromDate.Value), ""yyyy-mm-dd""))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@EndDate"", 200, 1, 20, Format(ParseDateForDB(Me.txtToDate.Value), ""yyyy-mm-dd""))" & vbCrLf
        
        ' AsOfDate is a DATE in SQL. Pass as String/Variant (200) with Null to avoid "Int is incompatible with Date" error.
        code = code & "        .Parameters.Append .CreateParameter(""@AsOfDate"", 200, 1, 20, Null)" & vbCrLf
        
        ' Currency ID
        code = code & "        Dim ccyID As Long" & vbCrLf
        code = code & "        ccyID = GetCurrencyID(Me.txtCurrency.Value)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@ReportingCurrencyId"", 3, 1, 4, ccyID)" & vbCrLf

        ' OutputFieldsCsv
        code = code & "        .Parameters.Append .CreateParameter(""@OutputFieldsCsv"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf

    ElseIf FormType = "Company Diversification" Then
        ' SP: edw.usp_CompanyDiversification_Aggregated2
        
        code = code & "        .NamedParameters = True" & vbCrLf
        
        code = code & "        .Parameters.Append .CreateParameter(""@MetricName"", 200, 1, 200, NullIfEmpty(finalMetric))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorNameId"", 3, 1, 4, CLng(strClientGlobalID))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorIdsCsv"", 200, 1, -1, NullIfEmpty(CStr(strClientGlobalID)))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@FundIdsCsv"", 200, 1, -1, NullIfEmpty(strFundIds))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@AsOfDate"", 133, 1, , ParseDateForDB(Me.txtAsofDate.Value))" & vbCrLf
        
        ' Currency ID
        code = code & "        Dim ccyID As Long" & vbCrLf
        code = code & "        ccyID = GetCurrencyID(Me.txtCurrency.Value)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@ReportingCurrencyId"", 3, 1, 4, ccyID)" & vbCrLf

        ' OutputFieldsCsv
        code = code & "        .Parameters.Append .CreateParameter(""@OutputFieldsCsv"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf

        

    ElseIf FormType = "My Performance" Then
        ' SP: edw.usp_InvestorPerformance_Aggregated
        
        ' 1. @Metric (nvarchar(200))
        code = code & "        .Parameters.Append .CreateParameter(""@Metric"", 200, 1, 200, NullIfEmpty(finalMetric))" & vbCrLf
        
        ' 2. @InvestorNameId (int)
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorNameId"", 3, 1, 4, CLng(strClientGlobalID))" & vbCrLf
        
        ' 3. @FundIdCsv (nvarchar(max))
        code = code & "        .Parameters.Append .CreateParameter(""@FundIdCsv"", 200, 1, -1, NullIfEmpty(strFundIds))" & vbCrLf
        
        ' 4. @AsOfDate (int) - strictly yyyymmdd integer
        code = code & "        .Parameters.Append .CreateParameter(""@AsOfDate"", 3, 1, 4, ParseDateToInt(Me.txtAsofDate.Value))" & vbCrLf

        ' 5. @OutputFieldsCsv (nvarchar(max))
        code = code & "        .Parameters.Append .CreateParameter(""@OutputFieldsCsv"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf
        
        ' 6. @ReportingCurrencyId (int)
        code = code & "        Dim ccyID As Long" & vbCrLf
        code = code & "        ccyID = GetCurrencyID(Me.txtCurrency.Value)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@ReportingCurrencyId"", 3, 1, 4, ccyID)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@FriendlyGuard"", 3, 1, 4, 0)" & vbCrLf

    ElseIf FormType = "Portfolio Diversification" Then
        ' SP: edw.usp_PortfolioDiversification_Aggregated
        
        code = code & "        .Parameters.Append .CreateParameter(""@Metric"", 200, 1, 100, NullIfEmpty(finalMetric))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorNameId"", 3, 1, 4, CLng(strClientGlobalID))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@FundIdsCsv"", 200, 1, -1, NullIfEmpty(strFundIds))" & vbCrLf
        
        ' Date as INT
        code = code & "        .Parameters.Append .CreateParameter(""@AsOfDate"", 3, 1, 4, ParseDateToInt(Me.txtAsofDate.Value))" & vbCrLf
        
        ' Currency ID
        code = code & "        Dim ccyID As Long" & vbCrLf
        code = code & "        ccyID = GetCurrencyID(Me.txtCurrency.Value)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@ReportingCurrencyId"", 3, 1, 4, ccyID)" & vbCrLf

        ' OutputFieldsCsv
        code = code & "        .Parameters.Append .CreateParameter(""@OutputFieldsCsv"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf
        
    Else
        ' Fallback to old behavior but using new IDs for safety? No, keep it matching old logic but with new helpers if needed.
        code = code & "        .Parameters.Append .CreateParameter(""@MetricName"", 200, 1, 100, NullIfEmpty(finalMetric))" & vbCrLf
        ' We don't have GetAccountKeys anymore, so using strFundIds as fallback... 
        code = code & "        .Parameters.Append .CreateParameter(""@SourceTableVolVal"", 201, 1, -1, NullIfEmpty(strFundIds))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@Date"", 133, 1, , ParseDateForDB(Me.txtAsofDate.Value))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@AttributeList"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf
        
        code = code & "        Dim ccyID As Long" & vbCrLf
        code = code & "        ccyID = GetCurrencyID(Me.txtCurrency.Value)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@ViewCurrencyID"", 3, 1, 4, ccyID)" & vbCrLf
    End If
    
    code = code & "    End With" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Set rs = cmd.Execute" & vbCrLf
    code = code & "    If Err.Number <> 0 Then" & vbCrLf
    code = code & "        MsgBox ""Error executing SP. Check 'variable_metric_map'. Metric sent: ['"" & finalMetric & ""']"" & vbCrLf & ""DB Error: "" & Err.Description, vbCritical" & vbCrLf
    code = code & "        conn.Close: Exit Sub" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    modUtilities.OutputToResults rs" & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf
    
    code = code & "Private Sub btnCancel_Click(): Unload Me: End Sub" & vbCrLf
    
    formComp.CodeModule.AddFromString code
End Sub
