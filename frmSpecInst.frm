VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
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
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbInstructions 
      Height          =   3240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5715
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmSpecInst.frx":0000
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
'Private madoData As nADOData.CADOData

Private Sub CancelButton_Click()

    Unload Me

End Sub

Private Sub OKButton_Click()

    If Me.rtbInstructions.TextRTF > " " And ProductionRun.Production_Run_Id > 0 Then
        Call writetext
    End If
    
    Unload Me

End Sub

Private Sub writetext()
        
        Dim strSQL As String
        Dim objData As nADOData.CADOData
        
        Screen.MousePointer = vbHourglass
        
        Set objData = New CADOData
        With objData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
            .ResetParameters
            .CommandType = adCmdText
        End With

        ' update the RTF Order Special Instruction
'        madoData.ResetParameters
'        madoData.CommandType = adCmdText
        
        
        strSQL = "declare @textpointer varbinary(16) " & _
                "select @textpointer = textptr(Production_Runs.Special_Inst) from Production_Runs where Production_Runs.Production_Run_Id = " & ProductionRun.Production_Run_Id & _
                "writetext Production_Runs.Special_Inst @textpointer " & " '" & CheckQuotes(Me.rtbInstructions.TextRTF) & "'"
                    
        objData.SQL = strSQL
        objData.Execute True
               
        ProductionRun.Special_Inst = Me.rtbInstructions
            
        Screen.MousePointer = vbDefault

End Sub
