VERSION 5.00
Begin VB.Form frmTestPDR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   9750
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
   ScaleHeight     =   2175
   ScaleWidth      =   9750
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
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9495
      Begin VB.TextBox txtToken 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "ross.lombardi 926062e1-e43f-464e-b65a-dc36bf5ad90f USBOH-SQLDEV01\USBOHSQLDEV01 QIClintrak"
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Token"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame fraPDR 
      Caption         =   "PDR"
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
      Top             =   720
      Width           =   9495
      Begin VB.TextBox txtJobLogId 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "39277"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtCodingNum 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "274784"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtFileLinksId 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "664688"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpen 
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
         TabIndex        =   10
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblCodingNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coding #"
         Height          =   195
         Left            =   120
         TabIndex        =   8
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
    Me.txtToken.Text = Command$
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

Private Sub cmdOpen_Click()
    Dim pdr As ProductionRuns.ProdRunMain
    
    Set pdr = New ProductionRuns.ProdRunMain
    With pdr
        Dim loginString() As String
        loginString = Split(Me.txtToken.Text, " ")
        If .Initialize(loginString(0), loginString(1), loginString(2), loginString(3), "\\ClkAlData\Clintrak_Data\ICONS\") Then
            Call .Load_Prod_Run(CLng(txtFileLinksId.Text), CLng(Me.txtCodingNum.Text), CLng(Me.txtJobLogId.Text))
        End If
      
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not mvarUser Is Nothing Then
        mvarUser.DisposeToken
    End If

End Sub

