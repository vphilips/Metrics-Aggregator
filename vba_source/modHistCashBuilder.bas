Attribute VB_Name = "modHistCashBuilder"
Option Explicit

' ==================================================================================
' DYNAMIC FORM BUILDER (CASCADING + GENERIC)
' ==================================================================================
' Programmatically creates forms driven by "form_config" (cascading) 
' and "form_metrics" (variable list).
' ==================================================================================

Public Sub BuildDynamicForm(ByVal FormType As String, ByVal FormCaption As String, ByVal SPName As String)
    If Not CheckSecurityAccess() Then Exit Sub
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
    
    ' --- HELPERS ---
    code = code & "Private Function NullIfEmpty(val As Variant) As Variant" & vbCrLf
    code = code & "    If IsNull(val) Or Trim(val & """") = """" Then NullIfEmpty = Null Else NullIfEmpty = val" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function SafeStr(val As Variant) As String" & vbCrLf
    code = code & "    If IsError(val) Then SafeStr = """" Else SafeStr = Trim(val & """")" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function GetSelectedItems(lst As Object) As String" & vbCrLf
    code = code & "    Dim i As Long, s As String: s = """"" & vbCrLf
    code = code & "    For i = 0 To lst.ListCount - 1" & vbCrLf
    code = code & "        If lst.Selected(i) Then s = s & lst.List(i) & "", """ & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "    If Len(s) > 0 Then s = Left(s, Len(s) - 2)" & vbCrLf
    code = code & "    GetSelectedItems = s" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf
    
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
    code = code & "    ' Mapping relies on 'form_config' for cascading regardless of where Valid Variables came from?" & vbCrLf
    code = code & "    ' Assumption: All valid Variables in form_metrics MUST exist in form_config col A to drive cascading." & vbCrLf
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
    code = code & "    ' Find the specific row" & vbCrLf
    code = code & "    Dim r As Long: r = 0" & vbCrLf
    code = code & "    For i = 2 To lastRow" & vbCrLf
    code = code & "        If SafeStr(ws.Cells(i, 1).Value) = Me.cboVariableName.Value And SafeStr(ws.Cells(i, 2).Value) = Me.cboClient.Value Then" & vbCrLf
    code = code & "            r = i" & vbCrLf
    code = code & "            Exit For" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    If r > 0 Then" & vbCrLf
    code = code & "        ' 1. Populate Accounts (Col C, separated by ;)" & vbCrLf
    code = code & "        Dim accRaw As Variant" & vbCrLf
    code = code & "        accRaw = ws.Cells(r, 3).Value" & vbCrLf
    code = code & "        If Not IsError(accRaw) And Not IsEmpty(accRaw) Then" & vbCrLf
    code = code & "            Dim parts() As String" & vbCrLf
    code = code & "            Dim cleanRaw As String" & vbCrLf
    code = code & "            cleanRaw = Replace(CStr(accRaw), "","", "";"") ' Handle commas too" & vbCrLf
    code = code & "            parts = Split(cleanRaw, "";"")" & vbCrLf
    code = code & "            Dim p As Variant" & vbCrLf
    code = code & "            For Each p In parts" & vbCrLf
    code = code & "                Me.lstAccounts.AddItem Trim(p)" & vbCrLf
    code = code & "            Next p" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "        " & vbCrLf
    code = code & "        ' 2. Auto-Fill Details" & vbCrLf
    
    ' Only fill logic that matches UI controls
    If FormType = "Historical Cashflows" Then
        code = code & "        Me.txtFromDate.Value = SafeStr(ws.Cells(r, 5).Value) ' Col E" & vbCrLf
        code = code & "        Me.txtToDate.Value = SafeStr(ws.Cells(r, 6).Value)   ' Col F" & vbCrLf
    Else
        ' As of Date is now in Col D (4)
        code = code & "        Me.txtAsofDate.Value = SafeStr(ws.Cells(r, 4).Value)   ' Col D maps to Asof" & vbCrLf
    End If
    
    ' Currency is in Col H (8)
    code = code & "        Me.txtCurrency.Value = SafeStr(ws.Cells(r, 8).Value) ' Col H" & vbCrLf
    code = code & "        " & vbCrLf
    ' Attributes (Output Fields) is in Col G (7)
    code = code & "        m_AttributesStr = SafeStr(ws.Cells(r, 7).Value) ' Col G" & vbCrLf
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
    code = code & "    Set f = rng.Find(What:=Trim(varName), LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    code = code & "    If Not f Is Nothing Then" & vbCrLf
    code = code & "        GetMetricFromVariable = SafeStr(f.Offset(0, 1).Value)" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        MsgBox ""Variable '"" & varName & ""' not found in 'variable_metric_map' sheet."", vbCritical" & vbCrLf
    code = code & "        GetMetricFromVariable = """"" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    code = code & "Private Function GetAccountKeys(lst As Object) As String" & vbCrLf
    code = code & "    ' 1. Load Account Map (Name -> GlobalID) into Dictionary for speed" & vbCrLf
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
    code = code & "        ' Read Columns A through C" & vbCrLf
    code = code & "        arr = ws.Range(""A2:C"" & lastRow).Value ' A=Name, B=Key, C=GlobalID" & vbCrLf
    code = code & "        For i = 1 To UBound(arr, 1)" & vbCrLf
    code = code & "            If SafeStr(arr(i, 1)) <> """" Then" & vbCrLf
    code = code & "                ' Map Name (Col 1) to GlobalID (Col 3)" & vbCrLf
    code = code & "                Dim nKey As String: nKey = SafeStr(arr(i, 1))" & vbCrLf
    code = code & "                If Not dict.Exists(nKey) Then dict.Add nKey, arr(i, 3)" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        Next i" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' 2. Map Selected items" & vbCrLf
    code = code & "    Dim s As String, key As Variant, accName As String" & vbCrLf
    code = code & "    s = """"" & vbCrLf
    code = code & "    Dim missingList As String: missingList = """"" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    For i = 0 To lst.ListCount - 1" & vbCrLf
    code = code & "        If lst.Selected(i) Then" & vbCrLf
    code = code & "            accName = SafeStr(lst.List(i))" & vbCrLf
    code = code & "            If dict.Exists(accName) Then" & vbCrLf
    code = code & "                s = s & dict(accName) & "",""" & vbCrLf
    code = code & "            Else" & vbCrLf
    code = code & "                missingList = missingList & vbCrLf & ""- "" & accName" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    If Len(missingList) > 0 Then" & vbCrLf
    code = code & "        MsgBox ""The following accounts could not be mapped to an Investor ID (check 'database' sheet):"" & vbCrLf & missingList, vbCritical" & vbCrLf
    code = code & "        GetAccountKeys = """"" & vbCrLf
    code = code & "        Exit Function" & vbCrLf
    code = code & "    End If" & vbCrLf
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
    code = code & "        s = Replace(s, ""'"", """")" & vbCrLf
    code = code & "        " & vbCrLf
    code = code & "        If IsDate(s) Then" & vbCrLf
    code = code & "            ParseDateForDB = CDate(s)" & vbCrLf
    code = code & "        ElseIf IsNumeric(s) Then" & vbCrLf
    code = code & "            On Error Resume Next" & vbCrLf
    code = code & "            ParseDateForDB = CDate(CDbl(s))" & vbCrLf
    code = code & "            If Err.Number <> 0 Then ParseDateForDB = Null" & vbCrLf
    code = code & "            On Error GoTo 0" & vbCrLf
    code = code & "        Else" & vbCrLf
    code = code & "            ParseDateForDB = Null" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    ' --- HELPER: DATE INT (YYYYMMDD) ---
    code = code & "Private Function ParseDateToInt(val As Variant) As Variant" & vbCrLf
    code = code & "    If IsNull(val) Or Trim(val & """") = """" Then" & vbCrLf
    code = code & "        ParseDateToInt = Null" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        Dim s As String" & vbCrLf
    code = code & "        s = Trim(val)" & vbCrLf
    code = code & "        If IsDate(s) Then" & vbCrLf
    code = code & "            ParseDateToInt = CLng(Format(CDate(s), ""yyyymmdd""))" & vbCrLf
    code = code & "        ElseIf IsNumeric(s) And Len(s) = 8 Then" & vbCrLf
    code = code & "            ParseDateToInt = CLng(s)" & vbCrLf
    code = code & "        Else" & vbCrLf
    code = code & "            ParseDateToInt = Null" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf

    ' --- HELPER: CURRENCY LOOKUP ---
    code = code & "Private Function GetCurrencyID(ccyName As String) As Long" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim ws As Worksheet, rng As Range, f As Range" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Sheets(""ccy_map"")" & vbCrLf
    code = code & "    If ws Is Nothing Then GetCurrencyID = 0: Exit Function" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' Look for Currency Code in Col A" & vbCrLf
    code = code & "    Set rng = ws.Range(""A:A"")" & vbCrLf
    code = code & "    Set f = rng.Find(What:=Trim(ccyName), LookIn:=xlValues, LookAt:=xlWhole)" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    If Not f Is Nothing Then" & vbCrLf
    code = code & "        GetCurrencyID = CLng(f.Offset(0, 1).Value) ' Col B has ID" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        GetCurrencyID = 0 ' Not found" & vbCrLf
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
    code = code & "    If finalMetric = """" Then Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    ' 2. Get Account Keys (GLOBAL IDs for ALL forms)" & vbCrLf
    code = code & "    Dim hasSelection As Boolean: hasSelection = False" & vbCrLf
    code = code & "    Dim k As Long" & vbCrLf
    code = code & "    For k = 0 To Me.lstAccounts.ListCount - 1" & vbCrLf
    code = code & "        If Me.lstAccounts.Selected(k) Then hasSelection = True: Exit For" & vbCrLf
    code = code & "    Next k" & vbCrLf
    code = code & "    If Not hasSelection Then MsgBox ""Select at least one Account"": Exit Sub" & vbCrLf
    code = code & "    " & vbCrLf
    code = code & "    Dim strAccountKeys As String" & vbCrLf
    code = code & "    strAccountKeys = GetAccountKeys(Me.lstAccounts)" & vbCrLf
    code = code & "    If strAccountKeys = """" Then Exit Sub" & vbCrLf
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
        ' Params: @Metric (200), @InvestorIdsCsv (MAX), @StartDate (DATE), @EndDate (DATE), @ReportingCurrencyId (INT), @OutputFieldsCsv (MAX)
        
        code = code & "        .Parameters.Append .CreateParameter(""@Metric"", 200, 1, 200, NullIfEmpty(finalMetric))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorIdsCsv"", 200, 1, -1, NullIfEmpty(strAccountKeys))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@StartDate"", 133, 1, , ParseDateForDB(Me.txtFromDate.Value))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@EndDate"", 133, 1, , ParseDateForDB(Me.txtToDate.Value))" & vbCrLf
        
        ' OutputFieldsCsv
        code = code & "        .Parameters.Append .CreateParameter(""@OutputFieldsCsv"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf
        
        ' Currency ID
        code = code & "        Dim ccyID As Long" & vbCrLf
        code = code & "        ccyID = GetCurrencyID(Me.txtCurrency.Value)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@ReportingCurrencyId"", 3, 1, 4, ccyID)" & vbCrLf

    ElseIf FormType = "Company Diversification" Then
        ' SP: edw.usp_CompanyDiversification_Aggregated
        ' Params: @MetricName (200), @InvestorNameIdsCsv (MAX), @AsOfDate (DATE/String?), @ReportingCurrencyId (INT), @OutputFieldsCsv (MAX)
        ' Note: Image shows @InvestorNameIdsCsv for Company Div, sticking to that.
        
        code = code & "        .Parameters.Append .CreateParameter(""@MetricName"", 200, 1, 200, NullIfEmpty(finalMetric))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorIdsCsv"", 200, 1, -1, NullIfEmpty(strAccountKeys))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@AsOfDate"", 133, 1, , ParseDateForDB(Me.txtAsofDate.Value))" & vbCrLf
        
        ' OutputFieldsCsv
        code = code & "        .Parameters.Append .CreateParameter(""@OutputFieldsCsv"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf

        ' Currency ID
        code = code & "        Dim ccyID As Long" & vbCrLf
        code = code & "        ccyID = GetCurrencyID(Me.txtCurrency.Value)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@ReportingCurrencyId"", 3, 1, 4, ccyID)" & vbCrLf

    ElseIf FormType = "My Performance" Then
        ' SP: edw.usp_InvestorPerformance_Aggregated
        ' Params: @MetricName (200), @InvestorIds (MAX), @AsOfDate (INT YYYYMMDD), @ReportingCurrency (STRING), @OutputFields (MAX)
        
        code = code & "        .Parameters.Append .CreateParameter(""@MetricName"", 200, 1, 200, NullIfEmpty(finalMetric))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorIds"", 200, 1, -1, NullIfEmpty(strAccountKeys))" & vbCrLf
        
        ' Date as INT
        code = code & "        .Parameters.Append .CreateParameter(""@AsOfDate"", 3, 1, 4, ParseDateToInt(Me.txtAsofDate.Value))" & vbCrLf
        
        ' OutputFields
        code = code & "        .Parameters.Append .CreateParameter(""@OutputFields"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf

        ' Currency as String
        code = code & "        .Parameters.Append .CreateParameter(""@ReportingCurrency"", 200, 1, 50, Me.txtCurrency.Value)" & vbCrLf

    ElseIf FormType = "Portfolio Diversification" Then
        ' SP: edw.usp_PortfolioDiversification_Aggregated
        ' Params: @Metric (100), @InvestorNameIdsCsv (NULL), @InvestorIdsCsv (MAX), @AsOfDate (INT), @ReportingCurrencyId (INT), @OutputFieldsCsv (MAX)
        
        code = code & "        .Parameters.Append .CreateParameter(""@Metric"", 200, 1, 100, NullIfEmpty(finalMetric))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorNameIdsCsv"", 200, 1, -1, Null)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@InvestorIdsCsv"", 200, 1, -1, NullIfEmpty(strAccountKeys))" & vbCrLf
        
        ' Date as INT
        code = code & "        .Parameters.Append .CreateParameter(""@AsOfDate"", 3, 1, 4, ParseDateToInt(Me.txtAsofDate.Value))" & vbCrLf
        
        ' OutputFieldsCsv
        code = code & "        .Parameters.Append .CreateParameter(""@OutputFieldsCsv"", 201, 1, -1, NullIfEmpty(m_AttributesStr))" & vbCrLf

        ' Currency ID
        code = code & "        Dim ccyID As Long" & vbCrLf
        code = code & "        ccyID = GetCurrencyID(Me.txtCurrency.Value)" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@ReportingCurrencyId"", 3, 1, 4, ccyID)" & vbCrLf

    Else
        ' Legacy / Fallback
        code = code & "        .Parameters.Append .CreateParameter(""@MetricName"", 200, 1, 100, NullIfEmpty(finalMetric))" & vbCrLf
        code = code & "        .Parameters.Append .CreateParameter(""@SourceTableVolVal"", 201, 1, -1, NullIfEmpty(strAccountKeys))" & vbCrLf
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
