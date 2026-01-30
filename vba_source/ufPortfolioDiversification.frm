VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPortfolioDiversification 
   Caption         =   "Portfolio Diversification"
   ClientHeight    =   380
   ClientWidth     =   650
   OleObjectBlob   =   "ufPortfolioDiversification.frx":0000
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   24
      Left            =   340
      TabIndex        =   47
      Top             =   320
      Width           =   60
   End
   Begin MSForms.CommandButton btnSubmit 
      Caption         =   "Submit"
      Default         =   -1  'True
      Height          =   24
      Left            =   270
      TabIndex        =   46
      Top             =   320
      Width           =   60
   End
   
   ' COLUMN 1
   Begin MSForms.Label lblAccount 
      Caption         =   "Account (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   0
      Top             =   6
      Width           =   90
   End
   Begin MSForms.TextBox txtAccount 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   1
      Top             =   6
      Width           =   120
   End
   
   Begin MSForms.Label lblClient 
      Caption         =   "Client (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   2
      Top             =   30
      Width           =   90
   End
   Begin MSForms.TextBox txtClient 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   3
      Top             =   30
      Width           =   120
   End
   
   Begin MSForms.Label lblFromDates 
      Caption         =   "From Dates (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   4
      Top             =   54
      Width           =   90
   End
   Begin MSForms.TextBox txtFromDates 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   5
      Top             =   54
      Width           =   120
   End
   
   Begin MSForms.Label lblToDates 
      Caption         =   "To Dates (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   6
      Top             =   78
      Width           =   90
   End
   Begin MSForms.TextBox txtToDates 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   7
      Top             =   78
      Width           =   120
   End
   
   Begin MSForms.Label lblCurrency 
      Caption         =   "Currency (*)"
      Height          =   18
      Left            =   6
      TabIndex        =   8
      Top             =   102
      Width           =   90
   End
   Begin MSForms.TextBox txtCurrency 
      BackColor       =   &H00FFC0C0&
      Height          =   18
      Left            =   102
      TabIndex        =   9
      Top             =   102
      Width           =   120
   End
   
   Begin MSForms.Label lblDate 
      Caption         =   "Date"
      Height          =   18
      Left            =   6
      TabIndex        =   10
      Top             =   126
      Width           =   90
   End
   Begin MSForms.TextBox txtDate 
      Height          =   18
      Left            =   102
      TabIndex        =   11
      Top             =   126
      Width           =   120
   End
   
   Begin MSForms.Label lblAIVFundGroupD 
      Caption         =   "AIV Fund Grp D"
      Height          =   18
      Left            =   6
      TabIndex        =   12
      Top             =   150
      Width           =   90
   End
   Begin MSForms.TextBox txtAIVFundGroupD 
      Height          =   18
      Left            =   102
      TabIndex        =   13
      Top             =   150
      Width           =   120
   End
   
   Begin MSForms.Label lblInvestor 
      Caption         =   "Investor"
      Height          =   18
      Left            =   6
      TabIndex        =   14
      Top             =   174
      Width           =   90
   End
   Begin MSForms.TextBox txtInvestor 
      Height          =   18
      Left            =   102
      TabIndex        =   15
      Top             =   174
      Width           =   120
   End
   
   Begin MSForms.Label lblEntryFund 
      Caption         =   "Entry Fund"
      Height          =   18
      Left            =   6
      TabIndex        =   16
      Top             =   198
      Width           =   90
   End
   Begin MSForms.TextBox txtEntryFund 
      Height          =   18
      Left            =   102
      TabIndex        =   17
      Top             =   198
      Width           =   120
   End
   
   Begin MSForms.Label lblManager 
      Caption         =   "Manager"
      Height          =   18
      Left            =   6
      TabIndex        =   18
      Top             =   222
      Width           =   90
   End
   Begin MSForms.TextBox txtManager 
      Height          =   18
      Left            =   102
      TabIndex        =   19
      Top             =   222
      Width           =   120
   End
   
   Begin MSForms.Label lblPortfolio 
      Caption         =   "Portfolio"
      Height          =   18
      Left            =   6
      TabIndex        =   20
      Top             =   246
      Width           =   90
   End
   Begin MSForms.TextBox txtPortfolio 
      Height          =   18
      Left            =   102
      TabIndex        =   21
      Top             =   246
      Width           =   120
   End
   
   Begin MSForms.Label lblPortCloseYear 
      Caption         =   "Close Year"
      Height          =   18
      Left            =   6
      TabIndex        =   22
      Top             =   270
      Width           =   90
   End
   Begin MSForms.TextBox txtPortCloseYear 
      Height          =   18
      Left            =   102
      TabIndex        =   23
      Top             =   270
      Width           =   120
   End
   
   ' COLUMN 2
   Begin MSForms.Label lblPortCommitYear 
      Caption         =   "Commit. Year"
      Height          =   18
      Left            =   340
      TabIndex        =   24
      Top             =   6
      Width           =   90
   End
   Begin MSForms.TextBox txtPortCommitYear 
      Height          =   18
      Left            =   436
      TabIndex        =   25
      Top             =   6
      Width           =   120
   End
   
   Begin MSForms.Label lblPortGeo 
      Caption         =   "Geography"
      Height          =   18
      Left            =   340
      TabIndex        =   26
      Top             =   30
      Width           =   90
   End
   Begin MSForms.TextBox txtPortGeo 
      Height          =   18
      Left            =   436
      TabIndex        =   27
      Top             =   30
      Width           =   120
   End
   
   Begin MSForms.Label lblPortGeoBroad 
      Caption         =   "Geo Broad"
      Height          =   18
      Left            =   340
      TabIndex        =   28
      Top             =   54
      Width           =   90
   End
   Begin MSForms.TextBox txtPortGeoBroad 
      Height          =   18
      Left            =   436
      TabIndex        =   29
      Top             =   54
      Width           =   120
   End
   
   Begin MSForms.Label lblPortGeoL3 
      Caption         =   "Geo L3"
      Height          =   18
      Left            =   340
      TabIndex        =   30
      Top             =   78
      Width           =   90
   End
   Begin MSForms.TextBox txtPortGeoL3 
      Height          =   18
      Left            =   436
      TabIndex        =   31
      Top             =   78
      Width           =   120
   End
   
   Begin MSForms.Label lblPortInd 
      Caption         =   "Industry"
      Height          =   18
      Left            =   340
      TabIndex        =   32
      Top             =   102
      Width           =   90
   End
   Begin MSForms.TextBox txtPortInd 
      Height          =   18
      Left            =   436
      TabIndex        =   33
      Top             =   102
      Width           =   120
   End
   
   Begin MSForms.Label lblPortIndL1 
      Caption         =   "Industry L1"
      Height          =   18
      Left            =   340
      TabIndex        =   34
      Top             =   126
      Width           =   90
   End
   Begin MSForms.TextBox txtPortIndL1 
      Height          =   18
      Left            =   436
      TabIndex        =   35
      Top             =   126
      Width           =   120
   End
   
   Begin MSForms.Label lblPortStage 
      Caption         =   "Stage"
      Height          =   18
      Left            =   340
      TabIndex        =   36
      Top             =   150
      Width           =   90
   End
   Begin MSForms.TextBox txtPortStage 
      Height          =   18
      Left            =   436
      TabIndex        =   37
      Top             =   150
      Width           =   120
   End
   
   Begin MSForms.Label lblPortStageBroad 
      Caption         =   "Stage Broad"
      Height          =   18
      Left            =   340
      TabIndex        =   38
      Top             =   174
      Width           =   90
   End
   Begin MSForms.TextBox txtPortStageBroad 
      Height          =   18
      Left            =   436
      TabIndex        =   39
      Top             =   174
      Width           =   120
   End
   
   Begin MSForms.Label lblPortStatus 
      Caption         =   "Status"
      Height          =   18
      Left            =   340
      TabIndex        =   40
      Top             =   198
      Width           =   90
   End
   Begin MSForms.TextBox txtPortStatus 
      Height          =   18
      Left            =   436
      TabIndex        =   41
      Top             =   198
      Width           =   120
   End
   
   Begin MSForms.Label lblPortTypeBroad 
      Caption         =   "Type Broad ID"
      Height          =   18
      Left            =   340
      TabIndex        =   42
      Top             =   222
      Width           =   90
   End
   Begin MSForms.TextBox txtPortTypeBroad 
      Height          =   18
      Left            =   436
      TabIndex        =   43
      Top             =   222
      Width           =   120
   End
   
   Begin MSForms.Label lblPortVintage 
      Caption         =   "Vintage Year"
      Height          =   18
      Left            =   340
      TabIndex        =   44
      Top             =   246
      Width           =   90
   End
   Begin MSForms.TextBox txtPortVintage 
      Height          =   18
      Left            =   436
      TabIndex        =   45
      Top             =   246
      Width           =   120
   End
   
End
Attribute VB_Name = "ufPortfolioDiversification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSubmit_Click()
    ' 1. Validate
    If Not ValidateMandatory(Me.txtAccount, "Account") Then Exit Sub
    If Not ValidateMandatory(Me.txtClient, "Client") Then Exit Sub
    If Not ValidateMandatory(Me.txtFromDates, "From Dates") Then Exit Sub
    If Not ValidateDate(Me.txtFromDates, "From Dates") Then Exit Sub
    If Not ValidateMandatory(Me.txtToDates, "To Dates") Then Exit Sub
    If Not ValidateDate(Me.txtToDates, "To Dates") Then Exit Sub
    If Not ValidateMandatory(Me.txtCurrency, "Currency") Then Exit Sub
    ' Date optional check
    If Not ValidateDate(Me.txtDate, "Date") Then Exit Sub
    
    ' 2. Params
    Dim pNames As Variant, pValues As Variant
    BuildParams pNames, pValues, _
        "@Account", Me.txtAccount.Value, _
        "@Client", Me.txtClient.Value, _
        "@FromDates", Me.txtFromDates.Value, _
        "@ToDates", Me.txtToDates.Value, _
        "@Currency", Me.txtCurrency.Value, _
        "@Date", Me.txtDate.Value, _
        "@AIVFundGroupD", Me.txtAIVFundGroupD.Value, _
        "@Investor", Me.txtInvestor.Value, _
        "@EntryFund", Me.txtEntryFund.Value, _
        "@Manager", Me.txtManager.Value, _
        "@Portfolio", Me.txtPortfolio.Value, _
        "@PortfolioCloseYear", Me.txtPortCloseYear.Value, _
        "@PortfolioCommitmentYear", Me.txtPortCommitYear.Value, _
        "@PortfolioGeography", Me.txtPortGeo.Value, _
        "@PortfolioGeographyBroad", Me.txtPortGeoBroad.Value, _
        "@PortfolioGeographyL3", Me.txtPortGeoL3.Value, _
        "@PortfolioIndustry", Me.txtPortInd.Value, _
        "@PortfolioIndustryL1", Me.txtPortIndL1.Value, _
        "@PortfolioStage", Me.txtPortStage.Value, _
        "@PortfolioStageBroad", Me.txtPortStageBroad.Value, _
        "@PortfolioStatus", Me.txtPortStatus.Value, _
        "@PortfolioTypeBroadID", Me.txtPortTypeBroad.Value, _
        "@PortfolioVintageYear", Me.txtPortVintage.Value

    ' 3. Execute
    Dim rs As Object
    Set rs = ExecuteSP("sp_GetPortfolioDiversification", pNames, pValues)
    
    ' 4. Output
    OutputToResults rs
    
    Unload Me
End Sub
