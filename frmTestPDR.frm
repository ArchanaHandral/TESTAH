VERSION 5.00
Begin VB.Form frmTestPDR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Runs Stub"
   ClientHeight    =   2295
   ClientLeft      =   13200
   ClientTop       =   4305
   ClientWidth     =   9735
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
   ScaleHeight     =   2295
   ScaleWidth      =   9735
   Begin VB.Frame fraRepReport 
      Caption         =   "Replacement Report"
      Height          =   1335
      Left            =   4800
      TabIndex        =   11
      Top             =   840
      Width           =   4815
      Begin VB.TextBox txtRunBarcode 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Text            =   "PDR278149"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpenRepReport 
         Caption         =   "Open"
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Run Barcode"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   915
      End
   End
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
         Text            =   "Daniel.Wenke 5fcc1a4a-81bb-4268-b6f4-7b2bb4b6dfe5 USBOH-SQLDEV01\USBOHSQLDEV01 DWClintrak"
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
      Top             =   840
      Width           =   4455
      Begin VB.TextBox txtJobLogId 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "58828"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtCodingNum 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "376483"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtFileLinksId 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "723635"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpenPDRWindow 
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

Private Sub cmdOpenRepReport_Click()
    Dim reprpt As ProductionRuns.ProdRunMain
    Dim loginString() As String
    
    loginString = Split(Me.txtToken.Text, " ")
    Set reprpt = New ProductionRuns.ProdRunMain
    If reprpt.Initialize(loginString(0), loginString(1), loginString(2), loginString(3), "\\ClkAlData\Clintrak_Data\ICONS\") Then
        reprpt.CreateMultiReport
        If Left$(Me.txtRunBarcode.Text, 3) = "PRG" Then
            reprpt.Print_PRGPlanning_Form Me.txtRunBarcode.Text, True
        End If
        reprpt.PrintAll_ProdPlanning_Forms Me.txtRunBarcode.Text, True
        On Error Resume Next
        Shell "del C:\yo.pdf"
        On Error GoTo 0
        reprpt.SaveToFile "C:\yo.pdf"
        Set reprpt = Nothing
    Else
        Set reprpt = Nothing
    End If
End Sub

Private Sub Form_Load()
    Me.txtToken.Text = Command$
End Sub



Private Sub cmdOpenPDRWindow_Click()
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

