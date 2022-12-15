VERSION 5.00
Begin VB.Form frmReason 
   Caption         =   "Client Inventory Reason"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   3510
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   3480
      TabIndex        =   2
      Top             =   3000
      Width           =   990
   End
   Begin VB.TextBox txtReason 
      Height          =   1935
      Left            =   240
      MaxLength       =   999
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "frmReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarBackupReason As String
Private Sub cmdOK_Click()
    If CheckForChanges() Then
        ProductionRun.UseClientInventoryReason = Trim(Me.txtReason.text)
    End If

    Unload Me
End Sub
Private Sub Form_Load()
    Me.txtReason.text = ProductionRun.UseClientInventoryReason
    ' store a copy of the original values for check for changes check
    BackupOnLoad
End Sub
Private Sub BackupOnLoad()
    mvarBackupReason = Trim(Me.txtReason.text)
End Sub
Private Function CheckForChanges() As Boolean
    If Trim(Me.txtReason.text) <> mvarBackupReason Then
        CheckForChanges = True
    Else
        CheckForChanges = False
    End If
End Function
Public Function SetReasonTextStatus(ByVal enabled As Boolean)
    Me.txtReason.enabled = enabled
End Function
