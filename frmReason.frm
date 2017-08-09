VERSION 5.00
Begin VB.Form frmReason 
   Caption         =   "Client Inventory Reason"
   ClientHeight    =   3975
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
   ScaleHeight     =   3975
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   3480
      TabIndex        =   6
      Top             =   3360
      Width           =   990
   End
   Begin VB.TextBox txtReason 
      Height          =   1935
      Left            =   240
      MaxLength       =   999
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   7575
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   2200
   End
   Begin VB.TextBox txtModifiedBy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   885
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   195
      Left            =   5220
      TabIndex        =   2
      Top             =   240
      Width           =   525
   End
   Begin VB.Label lblModifiedBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Modified By:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1230
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
    If CheckForChanges Then
        ProductionRun.UseClientInventoryModifiedBy = gApplicationUser.employeeId
        ProductionRun.UseClientInventoryDate = gLocationHandler.ConvertToEST(Now, gApplicationUser.ClintrakLocationId)
        ProductionRun.UseClientInventoryReason = Trim(Me.txtReason.text)
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    
    If ProductionRun.UseClientInventoryModifiedBy = 0 Then
        Me.txtModifiedBy.text = GetEmployeeName(gApplicationUser.employeeId)
        Me.txtDate.text = gLocationHandler.ConvertToLocal(Now, gApplicationUser.ClintrakLocationId) & " " & gUserLocation.Time_Zone_Display
        Me.txtReason.text = ""
    Else
        Me.txtModifiedBy.text = GetEmployeeName(ProductionRun.UseClientInventoryModifiedBy)
        Me.txtDate.text = gLocationHandler.ConvertToLocal(ProductionRun.UseClientInventoryDate, gApplicationUser.ClintrakLocationId) & " " & gUserLocation.Time_Zone_Display
        Me.txtReason.text = ProductionRun.UseClientInventoryReason
    End If
    
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
