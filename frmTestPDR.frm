VERSION 5.00
Begin VB.Form frmTestPDR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraConnectionSettings 
      Caption         =   "Connection settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtToken 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtDomain 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   9
         Text            =   "Winter!2#"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   "karentest"
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "ClkALSQLDev\CLINTRAKSQLDEV01,1433"
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Default         =   -1  'True
         Height          =   360
         Left            =   2880
         TabIndex        =   5
         Top             =   2280
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Token"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domain"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1965
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1605
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1245
         Width           =   720
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame fraPDR 
      Caption         =   "PDR"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   4455
      Begin VB.TextBox txtJobLogId 
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Text            =   "35092"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtCodingNum 
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Text            =   "0"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtFileLinksId 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "649390"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblJobLog 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Log Id"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblCodingNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coding #"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Links Id"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmTestPDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarUser As ClintrakCommon.ApplicationUser

Private Sub Form_Load()
    Dim WshNetwork As Object
    
    Set WshNetwork = CreateObject("WScript.Network")
    Me.txtDomain.Text = WshNetwork.UserDomain
    Me.txtUsername.Text = WshNetwork.Username
    Set WshNetwork = Nothing
End Sub


Private Sub cmdConnect_Click()
    If mvarUser Is Nothing Then
        If Authenticate(Me.txtUsername.Text, Me.txtPassword.Text, Me.txtDomain.Text) Then
            Me.txtToken.Text = mvarUser.Token
            Me.fraPDR.Enabled = True
        Else
            Set mvarUser = Nothing
        End If
    End If
End Sub

Private Function Authenticate(ByVal Username As String, ByVal Password As String, ByVal Domain As String) As Boolean
    Set mvarUser = New ApplicationUser
    With mvarUser
        If .Authenticate(Username, Password, Domain, AuthenticationPurpose.Login, Me.txtServer.Text, Me.txtDatabase.Text, "LabelProof") Then 'Application.ProductName
            Authenticate = True
        Else
            Authenticate = False
        End If
    End With
End Function

Private Sub Command1_Click()
    Dim pdr As ProductionRuns.ProdRunMain
    
    Set pdr = New ProductionRuns.ProdRunMain
    With pdr
        If .Initialize(Me.txtUsername.Text, Me.txtToken.Text, Me.txtServer.Text, Me.txtDatabase.Text, "\\ClkAlData\Clintrak_Data\ICONS\") Then
            Call .Load_Prod_Run(CLng(txtFileLinksId.Text), CLng(Me.txtCodingNum.Text), CLng(Me.txtJobLogId.Text))
            
        End If
      
    End With
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not mvarUser Is Nothing Then
        mvarUser.DisposeToken
    End If

End Sub
