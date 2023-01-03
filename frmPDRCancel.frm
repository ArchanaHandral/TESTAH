VERSION 5.00
Begin VB.Form frmPDRCancel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancellation Reason"
   ClientHeight    =   5295
   ClientLeft      =   2505
   ClientTop       =   4605
   ClientWidth     =   6510
   Icon            =   "frmPDRCancel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4802.721
   ScaleMode       =   0  'User
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   5135
      TabIndex        =   6
      Top             =   4852
      Width           =   1211
   End
   Begin VB.Frame fraNonBillableDetailsFrame 
      Height          =   4680
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox cmbPDRCancelReasons 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox txtPDRCancelNotes 
         Height          =   3000
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1500
         Width           =   6015
      End
      Begin VB.Label lblNonBillableAuthorizedByLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason (required):"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes (optional):"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1155
      End
   End
   Begin VB.Label lblcharlimit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character remaining: 200"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4874
      Width           =   1800
   End
End
Attribute VB_Name = "frmPDRCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    If Me.cmbPDRCancelReasons.ListIndex > -1 Then
     frmProdPlan.mvarPDRCancelDirtyflag = True
        
        ProductionRun.StatusLookupId = basGlobals.GetLookupId("PDRStatus", "Cancelled")
        ProductionRun.CancellationReasonLookupId = cmbPDRCancelReasons.itemData(cmbPDRCancelReasons.ListIndex)
        ProductionRun.CancellationNotes = Trim(Me.txtPDRCancelNotes.text)
        Unload Me
            
    Else
       
       MsgBox "You must select a reason to cancel the PDR or you can close this screen to undo your Cancellation attempt.", vbOKOnly + vbInformation, "Cancel PDR"
       frmProdPlan.mvarPDRCancelDirtyflag = False
    
    End If

End Sub

Private Sub Form_Load()
    basGlobals.GetPDRCancellationReasons cmbPDRCancelReasons
    UpdatePDRCancelUI

    ' store a copy of the original values for check for changes check
    'BackupOnLoad
End Sub

Private Sub UpdatePDRCancelUI()
    
    If ProductionRun.CancellationReasonLookupId > -1 Then
     Me.cmbPDRCancelReasons.ListIndex = GetPDRCancellationReasonListIndex(ProductionRun.CancellationReasonLookupId)
    Else
        Me.cmbPDRCancelReasons.ListIndex = GetPDRCancellationReasonListIndex(0)
    End If

    ' ProductionRun.CancellationReasonLookupId = cmbPDRCancelReasons.itemData(cmbPDRCancelReasons.ListIndex)
    Me.txtPDRCancelNotes.text = ProductionRun.CancellationNotes
    '  MsgBox "ProductionRun.CancellationNotes:" & ProductionRun.CancellationNotes & "IsValidSelection:" & IsValidSelection, vbOKOnly + vbInformation, "Cancel PDR"

End Sub

Private Function GetPDRCancellationReasonListIndex(PDRCancellationReasonLookupId As Integer)

    Dim i As Long

    For i = 0 To Me.cmbPDRCancelReasons.ListCount - 1

        If cmbPDRCancelReasons.itemData(i) = PDRCancellationReasonLookupId Then
            GetPDRCancellationReasonListIndex = i
            Exit Function

        End If

    Next i
    
    GetPDRCancellationReasonListIndex = -1

End Function

Public Static Sub ShowPDRCancel()

    Dim frm As frmPDRCancel

    Set frm = New frmPDRCancel
    frm.Show vbModal

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If IsValidSelection() = False Then
        frmProdPlan.mvarPDRCancelDirtyflag = False
       
    Else
        frmProdPlan.mvarPDRCancelDirtyflag = True
        ProductionRun.StatusLookupId = basGlobals.GetLookupId("PDRStatus", "Cancelled")
        ProductionRun.CancellationReasonLookupId = cmbPDRCancelReasons.itemData(cmbPDRCancelReasons.ListIndex)
        ProductionRun.CancellationNotes = Trim(Me.txtPDRCancelNotes.text)
    End If

End Sub
    
Private Function IsValidSelection() As Boolean

    If Me.cmbPDRCancelReasons.ListIndex > -1 Then
        IsValidSelection = True
    Else
        MsgBox "You did not select a reason, so your Cancellation attempt will be undone.", vbOKOnly + vbInformation, "Cancel PDR"
        IsValidSelection = False
    End If
    
End Function
