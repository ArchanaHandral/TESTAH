VERSION 5.00
Begin VB.Form frmSpecInst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Special Instructions"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInstructions 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpecInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This form is called from frmProdPlan.  Its purpose is to allow a production run's special instructions to be viewed and modified.</summary>
'</comment>

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()

    ProductionRun.Special_Inst = Me.txtInstructions.text
    Unload Me

End Sub
