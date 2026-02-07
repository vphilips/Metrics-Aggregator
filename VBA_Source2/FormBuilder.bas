Attribute VB_Name = "FormBuilder"
Option Explicit

' ==============================================================================
' Module: FormBuilder
' Description: Generates the UserForms programmatically.
'              Requires: Microsoft Visual Basic for Applications Extensibility 5.3
' ==============================================================================

' Constants for Control Layout
Private Const MARGIN_LEFT As Single = 12
Private Const MARGIN_TOP_START As Single = 12
Private Const CTRL_HEIGHT As Single = 18
Private Const CTRL_WIDTH_STD As Single = 200
Private Const LABEL_WIDTH As Single = 80
Private Const GAP_Y As Single = 6

Private CurrentTop As Single

' ------------------------------------------------------------------------------
' Main Entry Point
' ------------------------------------------------------------------------------
Public Sub BuildAllForms()
    ' Check for Trusted Access (Can't easily check via code without error, but we try)
    On Error Resume Next
    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Please enable 'Trust Access to the VBA Project Object Model' in Macro Settings.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Ensure References (Optional - Extensibility)
    AddExtensibilityReference
    
    ' Build My Performance
    BuildForm_MyPerformance
    
    ' Build Portfolio
    BuildForm_Portfolio
    
    ' Build Company
    BuildForm_Company
    
    ' Build Historical
    BuildForm_Historical
    
    MsgBox "All UserForms generated successfully!", vbInformation
End Sub

Private Sub AddExtensibilityReference()
    On Error Resume Next
    ' GUID for VBIDE
    ThisWorkbook.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' Generic Form Builder
' ------------------------------------------------------------------------------
Private Sub BuildForm_Generic(formName As String, formCaption As String, formType As String)
    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject
    
    ' 1. Delete Existing
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbp.VBComponents(formName)
    If Not vbComp Is Nothing Then
        vbp.VBComponents.Remove vbComp
    End If
    On Error GoTo 0
    
    ' 2. Create New
    Set vbComp = vbp.VBComponents.Add(3) ' vbext_ct_MSForm
    vbComp.Name = formName
    vbComp.Properties("Caption") = formCaption
    vbComp.Properties("Width") = 350
    vbComp.Properties("Height") = 400 ' Dynamic adjust later?
    
    Dim frm As Object
    Set frm = vbComp.Designer
    
    CurrentTop = MARGIN_TOP_START
    
    ' 3. Add Common Controls
    ' Variable Name
    AddLabel frm, "lblVariable", "Variable Name (*)"
    AddComboBox frm, "cmbVariable"
    NextRow
    
    ' Client
    AddLabel frm, "lblClient", "Client (*)"
    AddComboBox frm, "cmbClient"
    NextRow
    
    ' Accounts (+ Select All)
    AddLabel frm, "lblAccounts", "Account(s) (*)"
    AddCheckbox frm, "chkSelectAll", "Select All"
    NextRow
    
    ' ListBox (Taller)
    AddListBox frm, "lstAccounts", 100
    CurrentTop = CurrentTop + 100 + GAP_Y
    
    ' ASOF Date OR From/To
    If formType = "HISTORICAL" Then
        AddLabel frm, "lblFrom", "From Date (*)"
        AddTextBox frm, "txtFromDate"
        NextRow
        
        AddLabel frm, "lblTo", "To Date (*)"
        AddTextBox frm, "txtToDate"
        NextRow
    Else
        AddLabel frm, "lblAsOf", "As Of Date (*)"
        AddTextBox frm, "txtAsOfDate"
        NextRow
    End If
    
    ' Currency
    AddLabel frm, "lblCurrency", "Currency (*)"
    AddComboBox frm, "cmbCurrency"
    NextRow
    
    ' Buttons
    CurrentTop = CurrentTop + 10
    AddCommandButton frm, "cmdSubmit", "Submit", MARGIN_LEFT, CurrentTop
    AddCommandButton frm, "cmdCancel", "Cancel", MARGIN_LEFT + 100, CurrentTop
    
    ' 4. Inject Code
    InjectCode vbComp, formType
    
End Sub

' ------------------------------------------------------------------------------
' Specific Wrappers
' ------------------------------------------------------------------------------
Private Sub BuildForm_MyPerformance()
    BuildForm_Generic "frmMyPerformance", "My Performance", "MY_PERFORMANCE"
End Sub

Private Sub BuildForm_Portfolio()
    BuildForm_Generic "frmPortfolioDiversification", "Portfolio Diversification", "PORTFOLIO"
End Sub

Private Sub BuildForm_Company()
    BuildForm_Generic "frmCompanyDiversification", "Company Diversification", "COMPANY"
End Sub

Private Sub BuildForm_Historical()
    BuildForm_Generic "frmHistoricalCashflows", "Historical Cashflows", "HISTORICAL"
End Sub

' ------------------------------------------------------------------------------
' Helpers: Add Controls
' ------------------------------------------------------------------------------
Private Sub AddLabel(frm As Object, name As String, caption As String)
    Dim lbl As Object
    Set lbl = frm.Controls.Add("Forms.Label.1", name)
    lbl.Caption = caption
    lbl.Left = MARGIN_LEFT
    lbl.Top = CurrentTop
    lbl.Width = LABEL_WIDTH
    lbl.Height = CTRL_HEIGHT
End Sub

Private Sub AddComboBox(frm As Object, name As String)
    Dim cmb As Object
    Set cmb = frm.Controls.Add("Forms.ComboBox.1", name)
    cmb.Left = MARGIN_LEFT + LABEL_WIDTH + 10
    cmb.Top = CurrentTop
    cmb.Width = CTRL_WIDTH_STD
    cmb.Height = CTRL_HEIGHT
End Sub

Private Sub AddTextBox(frm As Object, name As String)
    Dim txt As Object
    Set txt = frm.Controls.Add("Forms.TextBox.1", name)
    txt.Left = MARGIN_LEFT + LABEL_WIDTH + 10
    txt.Top = CurrentTop
    txt.Width = CTRL_WIDTH_STD
    txt.Height = CTRL_HEIGHT
End Sub

Private Sub AddCheckbox(frm As Object, name As String, caption As String)
    Dim chk As Object
    Set chk = frm.Controls.Add("Forms.CheckBox.1", name)
    chk.Caption = caption
    chk.Left = MARGIN_LEFT + LABEL_WIDTH + 10
    chk.Top = CurrentTop
    chk.Width = 100
    chk.Height = CTRL_HEIGHT
    ' Note: Checkbox shares row with Label usually, or its own
End Sub

Private Sub AddListBox(frm As Object, name As String, height As Single)
    Dim lst As Object
    Set lst = frm.Controls.Add("Forms.ListBox.1", name)
    lst.Left = MARGIN_LEFT
    lst.Top = CurrentTop
    lst.Width = LABEL_WIDTH + 10 + CTRL_WIDTH_STD
    lst.Height = height
    lst.MultiSelect = 1 ' fmMultiSelectMulti
End Sub

Private Sub AddCommandButton(frm As Object, name As String, caption As String, left As Single, top As Single)
    Dim cmd As Object
    Set cmd = frm.Controls.Add("Forms.CommandButton.1", name)
    cmd.Caption = caption
    cmd.Left = left
    cmd.Top = top
    cmd.Width = 80
    cmd.Height = 24
End Sub

Private Sub NextRow()
    CurrentTop = CurrentTop + CTRL_HEIGHT + GAP_Y
End Sub

' ------------------------------------------------------------------------------
' Inject Code Helper
' ------------------------------------------------------------------------------
Private Sub InjectCode(vbComp As Object, formType As String)
    Dim code As String
    
    ' Common Init
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    FormLogic.Common_Initialize Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' Variable Change
    code = code & "Private Sub cmbVariable_Change()" & vbCrLf
    code = code & "    FormLogic.Common_VarChange Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' Client Change
    code = code & "Private Sub cmbClient_Change()" & vbCrLf
    code = code & "    FormLogic.Common_ClientChange Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' Select All
    code = code & "Private Sub chkSelectAll_Change()" & vbCrLf
    code = code & "    FormLogic.Common_SelectAllChange Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' Cancel
    code = code & "Private Sub cmdCancel_Click()" & vbCrLf
    code = code & "    FormLogic.Common_Cancel Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' Submit (Based on Type)
    code = code & "Private Sub cmdSubmit_Click()" & vbCrLf
    Select Case formType
        Case "MY_PERFORMANCE"
            code = code & "    FormLogic.SubmitMyPerformance Me" & vbCrLf
        Case "PORTFOLIO"
            code = code & "    FormLogic.SubmitPortfolio Me" & vbCrLf
        Case "COMPANY"
            code = code & "    FormLogic.SubmitCompany Me" & vbCrLf
        Case "HISTORICAL"
            code = code & "    FormLogic.SubmitHistorical Me" & vbCrLf
    End Select
    code = code & "End Sub" & vbCrLf
    
    vbComp.CodeModule.AddFromString code
End Sub
