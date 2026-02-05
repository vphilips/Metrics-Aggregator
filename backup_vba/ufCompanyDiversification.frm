VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufCompanyDiversification 
   Caption         =   "Company Diversification"
   ClientHeight    =   300
   ClientWidth     =   650
   OleObjectBlob   =   "ufCompanyDiversification.frx":0000
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   24
      Left            =   340
      TabIndex        =   40
      Top             =   240
      Width           =   60
   End
   Begin MSForms.CommandButton btnSubmit 
      Caption         =   "Submit"
      Default         =   -1  'True
      Height          =   24
      Left            =   270
      TabIndex        =   39
      Top             =   240
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
   
   Begin MSForms.Label lblCoExpGeoBroad 
      Caption         =   "Co Exp Geo Broad"
      Height          =   18
      Left            =   6
      TabIndex        =   16
      Top             =   198
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpGeoBroad 
      Height          =   18
      Left            =   102
      TabIndex        =   17
      Top             =   198
      Width           =   120
   End
   
   ' COLUMN 2
   Begin MSForms.Label lblCoExpGeoCount 
      Caption         =   "Co Exp Geo Ctry"
      Height          =   18
      Left            =   340
      TabIndex        =   18
      Top             =   6
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpGeoCount 
      Height          =   18
      Left            =   436
      TabIndex        =   19
      Top             =   6
      Width           =   120
   End
   
   Begin MSForms.Label lblCoExpIndID 
      Caption         =   "Co Exp Ind ID"
      Height          =   18
      Left            =   340
      TabIndex        =   20
      Top             =   30
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpIndID 
      Height          =   18
      Left            =   436
      TabIndex        =   21
      Top             =   30
      Width           =   120
   End
   
   Begin MSForms.Label lblCoExpIndBroad 
      Caption         =   "Co Exp Ind Broad"
      Height          =   18
      Left            =   340
      TabIndex        =   22
      Top             =   54
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpIndBroad 
      Height          =   18
      Left            =   436
      TabIndex        =   23
      Top             =   54
      Width           =   120
   End
   
   Begin MSForms.Label lblCoExpIndCat 
      Caption         =   "Co Exp Ind Cat"
      Height          =   18
      Left            =   340
      TabIndex        =   24
      Top             =   78
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpIndCat 
      Height          =   18
      Left            =   436
      TabIndex        =   25
      Top             =   78
      Width           =   120
   End
   
   Begin MSForms.Label lblCoExpInvType 
      Caption         =   "Co Exp Inv Type"
      Height          =   18
      Left            =   340
      TabIndex        =   26
      Top             =   102
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpInvType 
      Height          =   18
      Left            =   436
      TabIndex        =   27
      Top             =   102
      Width           =   120
   End
   
   Begin MSForms.Label lblCoExpInvYear 
      Caption         =   "Co Exp Inv Year"
      Height          =   18
      Left            =   340
      TabIndex        =   28
      Top             =   126
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpInvYear 
      Height          =   18
      Left            =   436
      TabIndex        =   29
      Top             =   126
      Width           =   120
   End
   
   Begin MSForms.Label lblCoExpPublic 
      Caption         =   "Co Exp Public"
      Height          =   18
      Left            =   340
      TabIndex        =   30
      Top             =   150
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpPublic 
      Height          =   18
      Left            =   436
      TabIndex        =   31
      Top             =   150
      Width           =   120
   End
   
   Begin MSForms.Label lblCoExpStage 
      Caption         =   "Co Exp Stage"
      Height          =   18
      Left            =   340
      TabIndex        =   32
      Top             =   174
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpStage 
      Height          =   18
      Left            =   436
      TabIndex        =   33
      Top             =   174
      Width           =   120
   End
   
   Begin MSForms.Label lblCoExpStageBroad 
      Caption         =   "Co Exp Stg Broad"
      Height          =   18
      Left            =   340
      TabIndex        =   34
      Top             =   198
      Width           =   90
   End
   Begin MSForms.TextBox txtCoExpStageBroad 
      Height          =   18
      Left            =   436
      TabIndex        =   35
      Top             =   198
      Width           =   120
   End
End
Attribute VB_Name = "ufCompanyDiversification"
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
        "@CompExpGeoBroad", Me.txtCoExpGeoBroad.Value, _
        "@CompExpGeoCountry", Me.txtCoExpGeoCount.Value, _
        "@CompExpIndId", Me.txtCoExpIndID.Value, _
        "@CompExpIndBroad", Me.txtCoExpIndBroad.Value, _
        "@CompExpIndCategory", Me.txtCoExpIndCat.Value, _
        "@CompExpInvType", Me.txtCoExpInvType.Value, _
        "@CompExpInvYear", Me.txtCoExpInvYear.Value, _
        "@CompExpIsPublic", Me.txtCoExpPublic.Value, _
        "@CompExpStage", Me.txtCoExpStage.Value, _
        "@CompExpStageBroad", Me.txtCoExpStageBroad.Value
    
    ' 3. Execute
    Dim rs As Object
    Set rs = ExecuteSP("sp_GetCompanyDiversification", pNames, pValues)
    
    ' 4. Output
    OutputToResults rs
    
    Unload Me
End Sub
