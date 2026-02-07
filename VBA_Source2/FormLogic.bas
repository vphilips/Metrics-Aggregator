Attribute VB_Name = "FormLogic"
Option Explicit

' ==============================================================================
' Module: FormLogic
' Description: Contains the event handling logic for UserForms.
'              The Form code-behind will delegate to these Subs.
' ==============================================================================

' Control Names (Must match FormBuilder generation)
Public Const CTL_VAR_NAME As String = "cmbVariable"
Public Const CTL_CLIENT As String = "cmbClient"
Public Const CTL_ACCOUNTS As String = "lstAccounts"
Public Const CTL_CHK_ALL As String = "chkSelectAll"
Public Const CTL_ASOF As String = "txtAsOfDate"
Public Const CTL_FROM As String = "txtFromDate"
Public Const CTL_TO As String = "txtToDate"
Public Const CTL_CURRENCY As String = "cmbCurrency"

' ------------------------------------------------------------------------------
' Common: Initialize
' ------------------------------------------------------------------------------
Public Sub Common_Initialize(frm As Object)
    ' Load Variable Name Dropdown
    Dim vars As Collection
    Dim v As Variant
    
    Set vars = ConfigReader.GetDistinctVariables()
    frm.Controls(CTL_VAR_NAME).Clear
    For Each v In vars
        frm.Controls(CTL_VAR_NAME).AddItem v
    Next v
    
    ' Load Currency Dropdown
    ' Note: SRD "Currency dropdown loads from ccy_map"
    ' We can allow both Code or ID, usually Display Code is best for human
    Dim ccys As Collection
    Set ccys = CurrencyLookup.GetAllCurrencies()
    
    frm.Controls(CTL_CURRENCY).Clear
    For Each v In ccys
        frm.Controls(CTL_CURRENCY).AddItem v
    Next v
    
End Sub

' ------------------------------------------------------------------------------
' Common: Variable Change -> Filter Client
' ------------------------------------------------------------------------------
Public Sub Common_VarChange(frm As Object)
    Dim selVar As String
    selVar = frm.Controls(CTL_VAR_NAME).Value
    
    frm.Controls(CTL_CLIENT).Clear
    frm.Controls(CTL_ACCOUNTS).Clear
    
    If Len(selVar) = 0 Then Exit Sub
    
    Dim clients As Collection
    Set clients = ConfigReader.GetClientsForVariable(selVar)
    
    Dim c As Variant
    For Each c In clients
        frm.Controls(CTL_CLIENT).AddItem c
    Next c
    
    ' Auto-select if only one
    If frm.Controls(CTL_CLIENT).ListCount = 1 Then
        frm.Controls(CTL_CLIENT).ListIndex = 0
        Common_ClientChange frm ' Trigger cascade
    End If
End Sub

' ------------------------------------------------------------------------------
' Common: Client Change -> Populate Accounts
' ------------------------------------------------------------------------------
Public Sub Common_ClientChange(frm As Object)
    Dim selVar As String
    Dim selClient As String
    
    selVar = frm.Controls(CTL_VAR_NAME).Value
    selClient = frm.Controls(CTL_CLIENT).Value
    
    frm.Controls(CTL_ACCOUNTS).Clear
    
    If Len(selVar) = 0 Or Len(selClient) = 0 Then Exit Sub
    
    Dim config As Object
    Set config = ConfigReader.GetConfigRowDetails(selVar, selClient)
    
    If config.Count = 0 Then Exit Sub
    
    ' Parse Accounts
    Dim accColl As Collection
    Set accColl = ConfigReader.ParseAccountsString(CStr(config("Accounts")))
    
    Dim acc As Variant
    For Each acc In accColl
        frm.Controls(CTL_ACCOUNTS).AddItem acc
    Next acc
    
    ' Pre-fill Dates if they exist and controls exist
    On Error Resume Next
    If Not IsEmpty(config("AsOfDate")) Then frm.Controls(CTL_ASOF).Value = Format(config("AsOfDate"), "yyyy-mm-dd")
    If Not IsEmpty(config("FromDate")) Then frm.Controls(CTL_FROM).Value = Format(config("FromDate"), "yyyy-mm-dd")
    If Not IsEmpty(config("ToDate")) Then frm.Controls(CTL_TO).Value = Format(config("ToDate"), "yyyy-mm-dd")
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' Common: Select All Checkbox
' ------------------------------------------------------------------------------
Public Sub Common_SelectAllChange(frm As Object)
    Dim chk As Boolean
    Dim i As Long
    chk = frm.Controls(CTL_CHK_ALL).Value
    
    With frm.Controls(CTL_ACCOUNTS)
        For i = 0 To .ListCount - 1
            .Selected(i) = chk
        Next i
    With End With
End Sub

' ------------------------------------------------------------------------------
' Helpers for Validation & Extraction
' ------------------------------------------------------------------------------
Private Function GetSelectedAccounts(frm As Object) As Collection
    Dim coll As New Collection
    Dim i As Long
    With frm.Controls(CTL_ACCOUNTS)
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                coll.Add .List(i)
            End If
        Next i
    End With
    Set GetSelectedAccounts = coll
End Function

Private Function ValidateCommon(frm As Object, ByRef outVar As String, ByRef outClient As String, ByRef outCcy As String, ByRef outIds As String, ByRef outConfig As Object) As Boolean
    ' 1. Check Variable
    outVar = frm.Controls(CTL_VAR_NAME).Value
    If Len(outVar) = 0 Then
        MsgBox "Please select a Variable Name.", vbExclamation
        Exit Function
    End If
    
    ' 2. Check Client
    outClient = frm.Controls(CTL_CLIENT).Value
    If Len(outClient) = 0 Then
        MsgBox "Please select a Client.", vbExclamation
        Exit Function
    End If
    
    ' 3. Check Accounts
    Dim selAcc As Collection
    Set selAcc = GetSelectedAccounts(frm)
    If selAcc.Count = 0 Then
        MsgBox "Please select at least one Account.", vbExclamation
        Exit Function
    End If
    
    ' 4. Look up IDs
    On Error Resume Next
    outIds = DatabaseLookup.GetGlobalIdsForList(selAcc)
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        Exit Function
    End If
    On Error GoTo 0
    
    ' 5. Check Currency
    outCcy = frm.Controls(CTL_CURRENCY).Value
    If Len(outCcy) = 0 Then
        MsgBox "Please select a Currency.", vbExclamation
        Exit Function
    End If
    
    ' 6. Get Config Row for Attributes
    Set outConfig = ConfigReader.GetConfigRowDetails(outVar, outClient)
    
    ValidateCommon = True
End Function

' ------------------------------------------------------------------------------
' SUBMIT: My Performance
' ------------------------------------------------------------------------------
Public Sub SubmitMyPerformance(frm As Object)
    Dim varName As String, client As String, ccy As String, ids As String
    Dim config As Object
    
    If Not ValidateCommon(frm, varName, client, ccy, ids, config) Then Exit Sub
    
    ' Date Validation (AsOf)
    Dim asOfStr As String
    asOfStr = frm.Controls(CTL_ASOF).Value
    If Not IsDate(asOfStr) Then
        MsgBox "Please enter a valid As Of Date.", vbExclamation
        Exit Sub
    End If
    
    ' Format: yyyymmdd numeric
    Dim dateNum As String
    dateNum = Format(CDate(asOfStr), "yyyymmdd")
   
    ' Currency: Code (e.g. USD) per SRD 4.1
    ' already in ccy variable
    
    ' Attributes
    Dim attrs As String
    attrs = config("Attributes")
    
    ' EXECUTE
    StoredProcRunner.ExecuteMyPerformance varName, ids, dateNum, ccy, attrs
End Sub

' ------------------------------------------------------------------------------
' SUBMIT: Portfolio Diversification
' ------------------------------------------------------------------------------
Public Sub SubmitPortfolio(frm As Object)
    Dim varName As String, client As String, ccy As String, ids As String
    Dim config As Object
    
    If Not ValidateCommon(frm, varName, client, ccy, ids, config) Then Exit Sub
    
    Dim asOfStr As String
    asOfStr = frm.Controls(CTL_ASOF).Value
    If Not IsDate(asOfStr) Then
        MsgBox "Please enter a valid As Of Date.", vbExclamation
        Exit Sub
    End If
    
    Dim dateNum As String
    dateNum = Format(CDate(asOfStr), "yyyymmdd")
   
    ' Currency: ID (e.g. 8) per SRD 4.2
    Dim ccyId As String
    ccyId = CurrencyLookup.GetCurrencyId(ccy)
    
    Dim attrs As String
    attrs = config("Attributes")
    
    StoredProcRunner.ExecutePortfolioDiversification varName, ids, dateNum, ccyId, attrs
End Sub

' ------------------------------------------------------------------------------
' SUBMIT: Company Diversification
' ------------------------------------------------------------------------------
Public Sub SubmitCompany(frm As Object)
    Dim varName As String, client As String, ccy As String, ids As String
    Dim config As Object
    
    If Not ValidateCommon(frm, varName, client, ccy, ids, config) Then Exit Sub
    
    Dim asOfStr As String
    asOfStr = frm.Controls(CTL_ASOF).Value
    If Not IsDate(asOfStr) Then
        MsgBox "Please enter a valid As Of Date.", vbExclamation
        Exit Sub
    End If
    
    ' Format: yyyy-mm-dd per SRD 4.3
    Dim dateIso As String
    dateIso = Format(CDate(asOfStr), "yyyy-mm-dd")
   
    Dim ccyId As String
    ccyId = CurrencyLookup.GetCurrencyId(ccy)
    
    Dim attrs As String
    attrs = config("Attributes")
    
    StoredProcRunner.ExecuteCompanyDiversification varName, ids, dateIso, ccyId, attrs
End Sub

' ------------------------------------------------------------------------------
' SUBMIT: Historical Cashflows
' ------------------------------------------------------------------------------
Public Sub SubmitHistorical(frm As Object)
    Dim varName As String, client As String, ccy As String, ids As String
    Dim config As Object
    
    If Not ValidateCommon(frm, varName, client, ccy, ids, config) Then Exit Sub
    
    ' Dates: From/To
    Dim fromStr As String, toStr As String
    fromStr = frm.Controls(CTL_FROM).Value
    toStr = frm.Controls(CTL_TO).Value
    
    If Not IsDate(fromStr) Or Not IsDate(toStr) Then
        MsgBox "Please enter valid From and To dates.", vbExclamation
        Exit Sub
    End If
    
    ' Check From <= To
    If CDate(fromStr) > CDate(toStr) Then
        MsgBox "From Date cannot be later than To Date.", vbExclamation
        Exit Sub
    End If
    
    ' Format: yyyy-mm-dd
    Dim startIso As String, endIso As String
    startIso = Format(CDate(fromStr), "yyyy-mm-dd")
    endIso = Format(CDate(toStr), "yyyy-mm-dd")
   
    Dim ccyId As String
    ccyId = CurrencyLookup.GetCurrencyId(ccy)
    
    Dim attrs As String
    attrs = config("Attributes")
    
    StoredProcRunner.ExecuteHistoricalCashflows varName, ids, startIso, endIso, ccyId, attrs
End Sub

' ------------------------------------------------------------------------------
' Helper: Cancel
' ------------------------------------------------------------------------------
Public Sub Common_Cancel(frm As Object)
    Unload frm
End Sub
