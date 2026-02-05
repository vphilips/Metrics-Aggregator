VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMyPerformance 
   Caption         =   "My Performance"
   ClientHeight    =   320
   ClientWidth     =   330
   OleObjectBlob   =   "ufMyPerformance.frx":0000
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   24
      Left            =   170
      TabIndex        =   19
      Top             =   250
      Width           =   60
   End
   Begin MSForms.CommandButton btnSubmit 
      Caption         =   "Submit"
      Default         =   -1  'True
      Height          =   24
      Left            =   100
      TabIndex        =   18
      Top             =   250
      Width           =   60
   End
   Begin MSForms.TextBox txtInvTxnQuarter 
      Height          =   18
      Left            =   102
      TabIndex        =   17
      Top             =   198
      Width           =   200
   End
   Begin MSForms.Label lblInvTxnQuarter 
      Caption         =   "Inv. Txn Quarter"
      Height          =   18
      Left            =   6
      TabIndex        =   16
      Top             =   198
      Width           =   90
   End
   Begin MSForms.TextBox txtInvestor 
      Height          =   18
      Left            =   102
      TabIndex        =   15
      Top             =   174
      Width           =   200
   End
   Begin MSForms.Label lblInvestor 
      Caption         =   "Investor"
      Height          =   18
      Left            =   6
      TabIndex        =   14
      Top             =   174
      Width           =   90
   End
   Begin MSForms.TextBox txtAIVFundGroupD 
      Height          =   18
      Left            =   102
      TabIndex        =   13
      Top             =   150
      Width           =   200
   End
   Begin MSForms.Label lblAIVFundGroupD 
      Caption         =   "AIV Fund Grp D"
      Height          =   18
      Left            =   6
      TabIndex        =   12
      Top             =   150
      Width           =   90
   End
   Begin MSForms.TextBox txtDate 
      Height          =   18
      Left            =   102
      TabIndex        =   11
      Top             =   126
      Width           =   200
   End
   Begin MSForms.Label lblDate 
      Caption         =   "Date"
      Height          =   18
      Left            =   6
      TabIndex        =   10
      Top             =   126
      Width           =   90
   End
   Begin MSForms.TextBox txtCurrency 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   9
      Top             =   102
      Width           =   200
   End
   Begin MSForms.Label lblCurrency 
      Caption         =   "Currency (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   8
      Top             =   102
      Width           =   90
   End
   Begin MSForms.TextBox txtToDates 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   7
      Top             =   78
      Width           =   200
   End
   Begin MSForms.Label lblToDates 
      Caption         =   "To Dates (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   6
      Top             =   78
      Width           =   90
   End
   Begin MSForms.TextBox txtFromDates 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   5
      Top             =   54
      Width           =   200
   End
   Begin MSForms.Label lblFromDates 
      Caption         =   "From Dates (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   4
      Top             =   54
      Width           =   90
   End
   Begin MSForms.TextBox txtClient 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   3
      Top             =   30
      Width           =   200
   End
   Begin MSForms.Label lblClient 
      Caption         =   "Client (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   2
      Top             =   30
      Width           =   90
   End
   Begin MSForms.TextBox txtAccount 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   1
      Top             =   6
      Width           =   200
   End
   Begin MSForms.Label lblAccount 
      Caption         =   "Account (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   0
      Top             =   6
      Width           =   90
   End
End
Attribute VB_Name = "ufMyPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSubmit_Click()
    ' 1. Validate Mandatory Fields
    If Not ValidateMandatory(Me.txtAccount, "Account") Then Exit Sub
    If Not ValidateMandatory(Me.txtClient, "Client") Then Exit Sub
    If Not ValidateMandatory(Me.txtFromDates, "From Dates") Then Exit Sub
    If Not ValidateDate(Me.txtFromDates, "From Dates") Then Exit Sub
    If Not ValidateMandatory(Me.txtToDates, "To Dates") Then Exit Sub
    If Not ValidateDate(Me.txtToDates, "To Dates") Then Exit Sub
    If Not ValidateMandatory(Me.txtCurrency, "Currency") Then Exit Sub
    If Not ValidateDate(Me.txtDate, "Date") Then Exit Sub
    
    ' 2. Prepare Parameters
    Dim pNames As Variant, pValues As Variant
    BuildParams pNames, pValues, _
        "@Account", Me.txtAccount.Value, _
        "@Client", Me.txtClient.Value, _
        "@FromDate", Me.txtFromDates.Value, _
        "@ToDate", Me.txtToDates.Value, _
        "@Currency", Me.txtCurrency.Value, _
        "@Date", Me.txtDate.Value, _
        "@AIVFundGroupD", Me.txtAIVFundGroupD.Value, _
        "@Investor", Me.txtInvestor.Value, _
        "@InvestorTransactionQuarter", Me.txtInvTxnQuarter.Value
    
    ' 3. Execute SP
    Dim rs As Object
    Set rs = ExecuteSP("sp_GetMyPerformance", pNames, pValues)
    
    ' 4. Output Results
    OutputToResults rs
    
    Unload Me
End Sub
