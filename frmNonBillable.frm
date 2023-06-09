VERSION 5.00
Begin VB.Form frmNonBillable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Billing Details"
   ClientHeight    =   3270
   ClientLeft      =   2505
   ClientTop       =   4605
   ClientWidth     =   6495
   Icon            =   "frmNonBillable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraNonBillableDetailsFrame 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox cmbNonBillableReason 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox txtNonBillablePRNumber 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1560
         Width           =   4680
      End
      Begin VB.TextBox txtNonBillableNotes 
         Height          =   1125
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1920
         Width           =   4680
      End
      Begin VB.Label lblNonBillableAuthorizedByLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorized By:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblNonBillableAuthorizedDateLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorized Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label lblNonBillableAuthorizedBy 
         BackStyle       =   0  'Transparent
         Caption         =   "____________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   4680
      End
      Begin VB.Label lblNonBillableAuthorizedDate 
         BackStyle       =   0  'Transparent
         Caption         =   "____________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         Width           =   4680
      End
      Begin VB.Label lblNonBillablePRNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PR Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmNonBillable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UpdateNonBillableUIReadOnlyStatus(isJobLevel As Boolean)
    Dim isLocalWithReason As Boolean
    
    isLocalWithReason = ((isJobLevel = False) And (GetNonBillableReasonSelected() > 0))
    
    Me.cmbNonBillableReason.Enabled = (isJobLevel = False) ' it is assumed job level is set to false if the PDR is billable
    Me.txtNonBillablePRNumber.Enabled = isLocalWithReason
    Me.txtNonBillableNotes.Enabled = isLocalWithReason
End Sub

Private Sub UpdateNonBillableUI(nb As CNonBillable)
    If nb.IsBillable Then
        Me.cmbNonBillableReason.ListIndex = GetNonBillableReasonListIndex(0)
    Else
        ' also come here if we have no reason, since that will fail IsBillable
        Me.cmbNonBillableReason.ListIndex = GetNonBillableReasonListIndex(nb.reasonId)
    End If

    Me.txtNonBillablePRNumber.text = nb.PRNumber
    Me.txtNonBillableNotes.text = nb.notes

    If nb.AuthorizedBy > 0 Then
        Me.lblNonBillableAuthorizedBy.Caption = GetEmployeeName(nb.AuthorizedBy)
        Me.lblNonBillableAuthorizedDate.Caption = basGlobals.ConvertDateWithTimeZone(nb.AuthorizedDate, gApplicationUser.ClintrakLocationId)
    Else
        Me.lblNonBillableAuthorizedBy.Caption = ""
        Me.lblNonBillableAuthorizedDate.Caption = ""
    End If
    
    ' may get called when changing ListIndex but this ensures it was called
    UpdateNonBillableUIReadOnlyStatus nb.isJobLevel
End Sub

Private Function GetNonBillableReasonListIndex(reasonId As Long)
    Dim i As Long
    For i = 0 To Me.cmbNonBillableReason.ListCount - 1
        If cmbNonBillableReason.itemData(i) = reasonId Then
            GetNonBillableReasonListIndex = i
            Exit Function
        End If
    Next i
    
    GetNonBillableReasonListIndex = -1
End Function

Private Function GetNonBillableReasonSelected() As Long
    If cmbNonBillableReason.ListIndex >= 0 Then
        GetNonBillableReasonSelected = cmbNonBillableReason.itemData(cmbNonBillableReason.ListIndex)
    Else
        GetNonBillableReasonSelected = -1
    End If
End Function

Public Static Sub ShowNonBillable()
    Dim frm As frmNonBillable
    Set frm = New frmNonBillable
    frm.Show vbModal
End Sub

Private Sub cmbNonBillableReason_Click()
    UpdateNonBillableUIReadOnlyStatus ProductionRun.NonBillable.isJobLevel
End Sub

Private Sub Form_Load()
    CenterForm Me
    basGlobals.GetNonBillableReasons cmbNonBillableReason
    UpdateNonBillableUI ProductionRun.NonBillable
    frmProdPlan.mvarBillingDirty = False
End Sub

Private Function IsValidSelection() As Boolean
    Dim newDept As String
    
    newDept = basGlobals.GetNonBillableDepartment(GetNonBillableReasonSelected)
    
    ' must use WasFCSDepartment because the original reason id is (intentionally) not a public property
    If ProductionRun.NonBillable.WasFCSDepartment And newDept <> "FCS" And GetNonBillableReasonSelected > 0 And basGlobals.gReprintFile_Type = "REPLACEMENT" Then
        ' This exists because the DB trigger that fires off an email when a reason is selected from billable
        ' cannot determine a reason change since it is on the PDR table. We want to send out an email when
        ' going from FCS to any other department but can't because the trigger can't detect that. So, as an
        ' interim solution, let's block that scenario: you can't go from FCS to anyone else. You must go
        ' through billable first.
        MsgBox "Please make this replacement PDR billable and save before switching the non-billable reason from FCS.", vbOKOnly + vbInformation, "Non-Billable Reason"
        IsValidSelection = False
    Else
        IsValidSelection = True
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    If IsValidSelection() = False Then
        Cancel = 1
    Else
        ProductionRun.NonBillable.reasonId = GetNonBillableReasonSelected
        'UpdateNonBillableUI ProductionRun.NonBillable
        
        If ProductionRun.NonBillable.HasChange() = True Then
            frmProdPlan.txtDirtyFlag.text = "Y"
            frmProdPlan.mvarBillingDirty = True
        Else
            frmProdPlan.mvarBillingDirty = False
        End If
        
        If ProductionRun.NonBillable.reasonId > 0 Then
            ProductionRun.NonBillable.PRNumber = txtNonBillablePRNumber.text
            ProductionRun.NonBillable.notes = txtNonBillableNotes.text
        End If
    End If
End Sub

