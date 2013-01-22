VERSION 5.00
Begin VB.Form frmDuplicate 
   Caption         =   "Duplicate"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   5775
      Begin VB.ListBox listPDRSameCoding 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   720
         Width           =   2160
      End
      Begin VB.ListBox listDupSetSameCoding 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   3480
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   720
         Width           =   2160
      End
      Begin VB.CommandButton cmdRemoveSameCoding 
         Caption         =   "<--Remove"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdAddSameCoding 
         Caption         =   "Add-->"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "* When duplicating using this method, the manual config. changes to the repeat values are copied over."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "PDR List (Same Coding Only)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "PDR's to Duplicate"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Keep Manual Configs*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   5775
      Begin VB.ListBox listPDR 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   135
         MultiSelect     =   2  'Extended
         TabIndex        =   16
         Top             =   480
         Width           =   2160
      End
      Begin VB.ListBox listDupSet 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   3480
         MultiSelect     =   2  'Extended
         TabIndex        =   15
         Top             =   480
         Width           =   2160
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add-->"
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<--Remove"
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PDR List (All Codings)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "PDR's to Duplicate"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox txtPDRNum 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5160
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Duplicate Using:"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmDuplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This form is called from frmProdPlan and is used to select PDRs to duplicate the sample configurations for the current PDR into.</summary>
'</comment>

Option Explicit

Dim PDR_List As CCOLdupFiles    'holder collection of PDR's that can be duplicated
' kbg 2008-009 added
Private PDR_SameCoding_List As CCOLdupFiles

Private Sub cmdAdd_Click()

Dim i As Long

    If listPDR.ListCount > 0 And listPDR.ListIndex > -1 Then
        For i = listPDR.ListCount To 1 Step -1
            If listPDR.Selected(i - 1) Then
            
                ' kbg 2008-009 added to remove from the dup with same coding box
                Dim n  As Long
                For n = dupSameCodingData.count To 1 Step -1
                    If dupSameCodingData.Item(n).productionRun_Barcode = PDR_List.Item(i).productionRun_Barcode Then
                        'add the item to the holder collection of possible PDR's
                        Call PDR_SameCoding_List.Add(dupSameCodingData.Item(n).productionId, dupSameCodingData.Item(n).fileName, dupSameCodingData.Item(n).productionRun_Barcode)
                        'adds the item to the listbox of possible PDR's to duplicate
                        Call listPDRSameCoding.AddItem(PDR_SameCoding_List.Item(PDR_SameCoding_List.count).productionRun_Barcode)
                        'removes the item from the collection of PDR's to duplicate
                        Call dupSameCodingData.Remove(n)
                        'removes the item from the listbox of PDR's to duplicate
                        Call listDupSetSameCoding.RemoveItem(n - 1)
                    End If
                Next
            
                'adds the item to the collection of PDR's to be duplicated
                Call dupData.Add(PDR_List.Item(i).productionId, PDR_List.Item(i).fileName, PDR_List.Item(i).productionRun_Barcode)
                'add the item to the listbox of PDR's to duplicate
                Call listDupSet.AddItem(dupData.Item(dupData.count).productionRun_Barcode)
                'removes the item from the holder collection
                Call PDR_List.Remove(i)
                'removes the item from the listbox of possible PDR's to duplicate
                Call listPDR.RemoveItem(i - 1)
            End If
        Next
    End If
End Sub

Private Sub cmdOK_Click()
    ' kbg 2008-009 added message box here
    If dupData.count = 0 And dupSameCodingData.count = 0 Then
        MsgBox "Production Runs have not been choosen!", _
                vbInformation + vbOKOnly, "Duplicate Samples"
        Exit Sub
    End If
    
    Unload Me
End Sub


Private Sub cmdRemove_Click()

Dim i As Long

    If listDupSet.ListCount > 0 And listDupSet.ListIndex > -1 Then
        For i = listDupSet.ListCount To 1 Step -1
            If listDupSet.Selected(i - 1) Then
                'add the item to the holder collection of possible PDR's
                Call PDR_List.Add(dupData.Item(i).productionId, dupData.Item(i).fileName, dupData.Item(i).productionRun_Barcode)
                'adds the item to the listbox of possible PDR's to duplicate
                Call listPDR.AddItem(PDR_List.Item(PDR_List.count).productionRun_Barcode)
                'removes the item from the collection of PDR's to duplicate
                Call dupData.Remove(i)
                'removes the item from the listbox of PDR's to duplicate
                Call listDupSet.RemoveItem(i - 1)
            End If
        Next
    End If
    
End Sub

Private Sub Form_Load()

Dim i As Long

    Set PDR_List = New CCOLdupFiles
    'loads the holder collection with all PDR's that are associated with the rand except for the current
    ' kbg 2008-009 added booAllCodings param
    Call LookUpSetToDuplicate(gRandomizationId, ProductionRun.Production_Run_Id, PDR_List, True)
    'loads the PDR listbox with the PDR's
    For i = 1 To PDR_List.count
        Call listPDR.AddItem(PDR_List.Item(i).productionRun_Barcode)
    Next
    
    ' kbg 2008-009 added another list box
    Set PDR_SameCoding_List = New CCOLdupFiles
    Call LookUpSetToDuplicate(gRandomizationId, ProductionRun.Production_Run_Id, PDR_SameCoding_List, False)
    For i = 1 To PDR_SameCoding_List.count
        Call listPDRSameCoding.AddItem(PDR_SameCoding_List.Item(i).productionRun_Barcode)
    Next
    
    txtPDRNum = frmProdPlan.txtBarcodeId
    
End Sub

' kbg 2008-009 added
Private Sub cmdRemoveSameCoding_Click()
    On Error GoTo Handle_Error

    Dim i As Long

    If listDupSetSameCoding.ListCount > 0 And listDupSetSameCoding.ListIndex > -1 Then
        For i = listDupSetSameCoding.ListCount To 1 Step -1
            If listDupSetSameCoding.Selected(i - 1) Then
                'add the item to the holder collection of possible PDR's
                Call PDR_SameCoding_List.Add(dupSameCodingData.Item(i).productionId, dupSameCodingData.Item(i).fileName, dupSameCodingData.Item(i).productionRun_Barcode)
                'adds the item to the listbox of possible PDR's to duplicate
                Call listPDRSameCoding.AddItem(PDR_SameCoding_List.Item(PDR_SameCoding_List.count).productionRun_Barcode)
                'removes the item from the collection of PDR's to duplicate
                Call dupSameCodingData.Remove(i)
                'removes the item from the listbox of PDR's to duplicate
                Call listDupSetSameCoding.RemoveItem(i - 1)
            End If
        Next
    End If

Cleanup_Exit:
     Exit Sub
Handle_Error:
     MsgBox Err.description & vbCrLf & _
         "in frmDuplicate.cmdRemoveSameCoding_Click ", _
         vbCritical + vbOKOnly, "Application Error"
     Resume Cleanup_Exit
End Sub

' kbg 2008-009 added
Private Sub cmdAddSameCoding_Click()
    On Error GoTo Handle_Error

    Dim i As Long

    If listPDRSameCoding.ListCount > 0 And listPDRSameCoding.ListIndex > -1 Then
        For i = listPDRSameCoding.ListCount To 1 Step -1
            If listPDRSameCoding.Selected(i - 1) Then
            
                Dim n  As Long
                For n = dupData.count To 1 Step -1
                    If dupData.Item(n).productionRun_Barcode = PDR_SameCoding_List.Item(i).productionRun_Barcode Then
                        'add the item to the holder collection of possible PDR's
                        Call PDR_List.Add(dupData.Item(n).productionId, dupData.Item(n).fileName, dupData.Item(n).productionRun_Barcode)
                        'adds the item to the listbox of possible PDR's to duplicate
                        Call listPDR.AddItem(PDR_List.Item(PDR_List.count).productionRun_Barcode)
                        'removes the item from the collection of PDR's to duplicate
                        Call dupData.Remove(n)
                        'removes the item from the listbox of PDR's to duplicate
                        Call listDupSet.RemoveItem(n - 1)
                    End If
                Next
            
            
                'adds the item to the collection of PDR's to be duplicated
                Call dupSameCodingData.Add(PDR_SameCoding_List.Item(i).productionId, PDR_SameCoding_List.Item(i).fileName, PDR_SameCoding_List.Item(i).productionRun_Barcode)
                'add the item to the listbox of PDR's to duplicate
                Call listDupSetSameCoding.AddItem(dupSameCodingData.Item(dupSameCodingData.count).productionRun_Barcode)
                'removes the item from the holder collection
                Call PDR_SameCoding_List.Remove(i)
                'removes the item from the listbox of possible PDR's to duplicate
                Call listPDRSameCoding.RemoveItem(i - 1)
            End If
        Next
    End If

Cleanup_Exit:
     Exit Sub
Handle_Error:
     MsgBox Err.description & vbCrLf & _
         "in frmDuplicate.cmdAddSameCoding_Click ", _
         vbCritical + vbOKOnly, "Application Error"
     Resume Cleanup_Exit
End Sub

' kbg 2008-009 added
Private Sub cmdCancel_Click()
    Dim i As Long
    For i = dupData.count To 1 Step -1
        dupData.Remove (i)
    Next
    
    For i = dupSameCodingData.count To 1 Step -1
        dupSameCodingData.Remove (i)
    Next
    Unload Me
    
End Sub
