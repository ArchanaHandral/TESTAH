VERSION 5.00
Begin VB.Form frmCancelNote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancellation Reason"
   ClientHeight    =   4185
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbCancelReasons 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.TextBox txtCancelNotes 
      Height          =   2295
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblReasonRequired 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason (required):"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label lblNotesOptional 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes (optional):"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Characters remaining:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Tag             =   "Characters remaining:"
      Top             =   3600
      Width           =   1530
   End
End
Attribute VB_Name = "frmCancelNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public reason As String
Public notes As String

'Private Sub CancelButton_Click()
'    reason = ""
'    Unload Me
'End Sub


Private Sub Form_Load()
    basGlobals.GetPDRCancellationReasons CmbCancelReasons
    SetCharactersRemaining

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    reason = Me.CmbCancelReasons.text
    
    If reason = "" Then
        frmProdPlan.mvarPDRCancelDirtyflag = False
        MsgBox _
            "You did not select a reason, so your Cancellation attempt will be undone.", _
            vbExclamation
    Else
        UpdatePDRCancelUI
    End If
    
End Sub
Private Sub UpdatePDRCancelUI()
    notes = Me.txtCancelNotes.text
    If Trim$(notes) = "" Then
        notes = "N/A"
    End If
   frmProdPlan.mvarPDRCancelDirtyflag = True
   ProductionRun.StatusLookupId = basGlobals.GetLookupId("PDRStatus", "Cancelled")
   ProductionRun.CancellationReasonLookupId = CmbCancelReasons.itemData(CmbCancelReasons.ListIndex)
   ProductionRun.CancellationNotes = notes
End Sub
    
Private Sub Label1_Click()

End Sub

Private Sub cmdOK_Click()
    If Len(Trim$(CmbCancelReasons.text)) = 0 Then
        MsgBox _
            "You must select a reason or close this screen to undo your Cacellation attempt.", _
            vbExclamation
        Exit Sub
    End If
    'reason = Me.ComReason.text
    'UpdatePDRCancelUI
    Unload Me
End Sub

Private Sub txtCancelNotes_Change()
   SetCharactersRemaining
    
End Sub

Private Sub SetCharactersRemaining()
    Me.lblLabel1.Caption = Me.lblLabel1.Tag & CStr(Me.txtCancelNotes.MaxLength - Len(Me.txtCancelNotes.text))
End Sub
