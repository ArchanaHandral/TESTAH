VERSION 5.00
Begin VB.Form frmNotes 
   Caption         =   "Notes"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1770
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtNotes 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This form is called from frmSmpConfig.  Its purpose is to allow sample configuration notes to be viewed and modified.</summary>
'</comment>

Option Explicit

Private Sub cmdOK_Click()
    If Trim$(txtNotes) <> Trim$(frmSmpConfig.notes) Then
        frmSmpConfig.dirtyFlag = "Y"
        frmSmpConfig.notes = Trim$(txtNotes)
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    txtNotes = Trim$(frmSmpConfig.notes)
End Sub

