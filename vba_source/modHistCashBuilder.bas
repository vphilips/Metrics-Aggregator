Attribute VB_Name = "modHistCashBuilder"
Option Explicit

' ==================================================================================
' HISTORICAL CASHFLOW NEW FORM BUILDER (CASCADING)
' ==================================================================================
' Programmatically creates frmHistCashNew driven by "form_config" sheet.
' ==================================================================================

Public Sub BuildHistCashNewForm()
    If Not CheckSecurityAccess() Then Exit Sub
    Application.ScreenUpdating = False
    
    DeleteFormIfExists "frmHistCashNew"
    
    Dim uf As Object
    Set uf = CreateUserForm("frmHistCashNew", "Historical Cashflows (Dynamic)", 450, 500)
    
    Dim topPos As Double: topPos = 10
    
    ' 1. Logic Fields (Cascading)
    AddLabelAndCombobox uf, topPos, "VariableName", "Variable Name", True: topPos = topPos + 24
    AddLabelAndCombobox uf, topPos, "Client", "Client", True: topPos = topPos + 24
    
    ' 2. Account ListBox (Box + Select All)
    AddLabelAndListBox uf, topPos, "Accounts", "Account(s)", True, 120
    topPos = topPos + 120 + 24
    
    ' 3. Auto-Filled Fields
    AddLabelAndText uf, topPos, "FromDate", "From Date", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "ToDate", "To Date", True: topPos = topPos + 24
    AddLabelAndText uf, topPos, "Currency", "Currency", True: topPos = topPos + 24
    
    ' 4. Inject Logic
    InjectCascadingLogic uf, "edw.pr_usp_investor_transactions2"
    
    Application.ScreenUpdating = True
    MsgBox "frmHistCashNew (Cascading) has been created!", vbInformation
End Sub

' ==================================================================================
' HELPERS
' ==================================================================================
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

Private Sub InjectCascadingLogic(formComp As Object, spName As String)
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
    
    ' --- HELPERS ---
    code = code & "Private Function NullIfEmpty(val As Variant) As Variant" & vbCrLf
    code = code & "    If IsNull(val) Or Trim(val & """") = """" Then NullIfEmpty = Null Else NullIfEmpty = val" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function GetSelectedItems(lst As Object) As String" & vbCrLf
    code = code & "    Dim i As Long, s As String: s = """"" & vbCrLf
    code = code & "    For i = 0 To lst.ListCount - 1" & vbCrLf
    code = code & "        If lst.Selected(i) Then s = s & lst.List(i) & "", """ & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "    If Len(s) > 0 Then s = Left(s, Len(s) - 2)" & vbCrLf
    code = code & "    GetSelectedItems = s" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf
    
    ' --- INIT: Load Variable Names ---
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    LoadVariableNames" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub LoadVariableNames()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim ws As Worksheet, rng As Range, cell As Range" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""form_config"")" & vbCrLf
    code = code & "    If ws Is Nothing Then MsgBox ""Sheet 'form_config' not found!"": Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' Dictionary for Unique" & vbCrLf
    code = code & "    Dim dict As Object" & vbCrLf
    code = code & "    Set dict = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    code = code & "    Dim lastRow As Long" & vbCrLf
    code = code & "    lastRow = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row" & vbCrLf
    code = code & "    If lastRow < 2 Then Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    For Each cell In ws.Range(""A2:A"" & lastRow)" & vbCrLf
    code = code & "        If Not IsEmpty(cell.Value) And Not dict.Exists(cell.Value) Then" & vbCrLf
    code = code & "            dict.Add cell.Value, Nothing" & vbCrLf
    code = code & "            Me.cboVariableName.AddItem cell.Value" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next cell" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' --- CASCADE 1: Variable -> Client ---
    code = code & "Private Sub cboVariableName_Change()" & vbCrLf
    code = code & "    Me.cboClient.Clear" & vbCrLf
    code = code & "    Me.lstAccounts.Clear" & vbCrLf
    code = code & "    Me.txtFromDate.Value = """"" & vbCrLf
    code = code & "    Me.txtToDate.Value = """"" & vbCrLf
    code = code & "    Me.txtCurrency.Value = """"" & vbCrLf
    code = code & "    m_AttributesStr = """"" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    Dim ws As Worksheet, lastRow As Long, i As Long" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""form_config"")" & vbCrLf
    code = code & "    lastRow = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    For i = 2 To lastRow" & vbCrLf
    code = code & "        If ws.Cells(i, 1).Value = Me.cboVariableName.Value Then" & vbCrLf
    code = code & "            Me.cboClient.AddItem ws.Cells(i, 2).Value ' Col B" & vbCrLf
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
    code = code & "    ' Find the specific row" & vbCrLf
    code = code & "    Dim r As Long: r = 0" & vbCrLf
    code = code & "    For i = 2 To lastRow" & vbCrLf
    code = code & "        If ws.Cells(i, 1).Value = Me.cboVariableName.Value And ws.Cells(i, 2).Value = Me.cboClient.Value Then" & vbCrLf
    code = code & "            r = i" & vbCrLf
    code = code & "            Exit For" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    If r > 0 Then" & vbCrLf
    code = code & "        ' 1. Populate Accounts (Col C, separated by ;)" & vbCrLf
    code = code & "        Dim accRaw As String" & vbCrLf
    code = code & "        accRaw = ws.Cells(r, 3).Value" & vbCrLf
    code = code & "        Dim parts() As String" & vbCrLf
    code = code & "        parts = Split(accRaw, "";"")" & vbCrLf
    code = code & "        Dim p As Variant" & vbCrLf
    code = code & "        For Each p In parts" & vbCrLf
    code = code & "            Me.lstAccounts.AddItem Trim(p)" & vbCrLf
    code = code & "        Next p" & vbCrLf
    code = code & "        " & vbCrLf
    code = code & "        ' 2. Auto-Fill Details" & vbCrLf
    code = code & "        Me.txtFromDate.Value = ws.Cells(r, 5).Value ' Col E" & vbCrLf
    code = code & "        Me.txtToDate.Value = ws.Cells(r, 6).Value   ' Col F" & vbCrLf
    code = code & "        Me.txtCurrency.Value = ws.Cells(r, 7).Value ' Col G" & vbCrLf
    code = code & "        " & vbCrLf
    code = code & "        ' 3. Capture Attributes (Col H)" & vbCrLf
    code = code & "        m_AttributesStr = ws.Cells(r, 8).Value" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' --- Select All ---
    code = code & "Private Sub chkSelectAllAccounts_Click()" & vbCrLf
    code = code & "    Dim i As Long" & vbCrLf
    code = code & "    For i = 0 To Me.lstAccounts.ListCount - 1" & vbCrLf
    code = code & "        Me.lstAccounts.Selected(i) = Me.chkSelectAllAccounts.Value" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' --- HELPERS FOR LOOKUPS ---
    code = code & "Private Function GetMetricFromVariable(varName As String) As String" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim ws As Worksheet, rng As Range, f As Range" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""variable_metric_map"")" & vbCrLf
    code = code & "    If ws Is Nothing Then GetMetricFromVariable = varName: Exit Function" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' Assumes Variable in Col A, Metric in Col B" & vbCrLf
    code = code & "    Set rng = ws.Range(""A:A"")" & vbCrLf
    code = code & "    Set f = rng.Find(What:=varName, LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    code = code & "    If Not f Is Nothing Then" & vbCrLf
    code = code & "        GetMetricFromVariable = f.Offset(0, 1).Value" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        GetMetricFromVariable = varName ' Fallback" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function GetAccountKeys(lst As Object) As String" & vbCrLf
    code = code & "    ' 1. Load Account Map (Name -> Key) into Dictionary for speed" & vbCrLf
    code = code & "    Dim dict As Object" & vbCrLf
    code = code & "    Set dict = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    code = code & "    dict.CompareMode = 1 ' TextCompare" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    Dim ws As Worksheet, arr As Variant, i As Long, lastRow As Long" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""database"")" & vbCrLf
    code = code & "    If ws Is Nothing Then MsgBox ""Database sheet missing!"": Exit Function" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    lastRow = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row" & vbCrLf
    code = code & "    If lastRow >= 2 Then" & vbCrLf
    code = code & "        arr = ws.Range(""A2:B"" & lastRow).Value ' A=Name, B=Key" & vbCrLf
    code = code & "        For i = 1 To UBound(arr, 1)" & vbCrLf
    code = code & "            If Not IsError(arr(i, 1)) And Not IsEmpty(arr(i, 1)) Then" & vbCrLf
    code = code & "                If Not dict.Exists(arr(i, 1)) Then dict.Add arr(i, 1), arr(i, 2)" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        Next i" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' 2. Map Selected Items" & vbCrLf
    code = code & "    Dim s As String, key As Variant, accName As String" & vbCrLf
    code = code & "    s = """"" & vbCrLf
    code = code & "    For i = 0 To lst.ListCount - 1" & vbCrLf
    code = code & "        If lst.Selected(i) Then" & vbCrLf
    code = code & "            accName = lst.List(i)" & vbCrLf
    code = code & "            If dict.Exists(accName) Then" & vbCrLf
    code = code & "                s = s & dict(accName) & "",""" & vbCrLf
    code = code & "            Else" & vbCrLf
    code = code & "                ' If Key missing, use Name?? User said use Key. " & vbCrLf
    code = code & "                ' We will warn logic or just append Name if no key found (safer fallback)?" & vbCrLf
    code = code & "                ' User requested Account Key. If invalid, SP might fail." & vbCrLf
    code = code & "                ' Let's append the key if found, skip if not? Or append name?" & vbCrLf
    code = code & "                ' We'll append name as fallback." & vbCrLf
    code = code & "                s = s & accName & "",""" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    If Len(s) > 0 Then s = Left(s, Len(s) - 1)" & vbCrLf
    code = code & "    GetAccountKeys = s" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    ' --- HELPERS FOR DATES ---
    code = code & "Private Function ParseDateForDB(val As Variant) As Variant" & vbCrLf
    code = code & "    If IsNull(val) Or Trim(val & """") = """" Then" & vbCrLf
    code = code & "        ParseDateForDB = Null" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        Dim s As String" & vbCrLf
    code = code & "        s = Trim(val)" & vbCrLf
    code = code & "        ' Remove single quotes" & vbCrLf
    code = code & "        s = Replace(s, ""'"", """")" & vbCrLf
    code = code & "        " & vbCrLf
    code = code & "        If IsDate(s) Then" & vbCrLf
    code = code & "            ParseDateForDB = CDate(s)" & vbCrLf
    code = code & "        ElseIf IsNumeric(s) Then" & vbCrLf
    code = code & "            ' Handle Excel serial numbers" & vbCrLf
    code = code & "            On Error Resume Next" & vbCrLf
    code = code & "            ParseDateForDB = CDate(CDbl(s))" & vbCrLf
    code = code & "            If Err.Number <> 0 Then ParseDateForDB = Null" & vbCrLf
    code = code & "            On Error GoTo 0" & vbCrLf
    code = code & "        Else" & vbCrLf
    code = code & "            ParseDateForDB = Null" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    ' --- SUBMIT ---
    code = code & "Private Sub btnSubmit_Click()" & vbCrLf
    code = code & "    ' Validate" & vbCrLf
    code = code & "    If Me.cboVariableName.Value = """" Then MsgBox ""Variable required"": Exit Sub" & vbCrLf
    code = code & "    If Me.cboClient.Value = """" Then MsgBox ""Client required"": Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' 1. Get Metric Name" & vbCrLf
    code = code & "    Dim finalMetric As String" & vbCrLf
    code = code & "    finalMetric = GetMetricFromVariable(Me.cboVariableName.Value)" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' 2. Get Account Keys" & vbCrLf
    code = code & "    Dim strAccountKeys As String" & vbCrLf
    code = code & "    strAccountKeys = GetAccountKeys(Me.lstAccounts)" & vbCrLf
    code = code & "    If strAccountKeys = """" Then MsgBox ""Select at least one Account"": Exit Sub" & vbCrLf
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
    code = code & "        .Parameters.Append .CreateParameter(""@MetricName"", 200, 1, 100, NullIfEmpty(finalMetric))" & vbCrLf
    code = code & "        .Parameters.Append .CreateParameter(""@SourceTableVolVal"", 201, 1, -1, NullIfEmpty(strAccountKeys))" & vbCrLf
    code = code & "        .Parameters.Append .CreateParameter(""@StartDate"", 133, 1, , ParseDateForDB(Me.txtFromDate.Value))" & vbCrLf
    code = code & "        .Parameters.Append .CreateParameter(""@EndDate"", 133, 1, , ParseDateForDB(Me.txtToDate.Value))" & vbCrLf
    code = code & "        .Parameters.Append .CreateParameter(""@AttributeList"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf
    code = code & "        .Parameters.Append .CreateParameter(""@ViewCurrencyID"", 3, 1, 4, 1)" & vbCrLf
    code = code & "    End With" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Set rs = cmd.Execute" & vbCrLf
    code = code & "    If Err.Number <> 0 Then" & vbCrLf
    code = code & "        MsgBox ""Error executing SP: "" & Err.Description, vbCritical" & vbCrLf
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
