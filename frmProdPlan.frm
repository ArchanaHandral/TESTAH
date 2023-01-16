VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmProdPlan 
   BorderStyle     =   0  'None
   Caption         =   "Computerization Order"
   ClientHeight    =   9675
   ClientLeft      =   6780
   ClientTop       =   4170
   ClientWidth     =   10650
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
   Icon            =   "frmProdPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPdrSpecialInstructions 
      Height          =   285
      Left            =   1320
      TabIndex        =   68
      Top             =   9000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtLinksSpecialInstructions 
      Height          =   285
      Left            =   1680
      TabIndex        =   67
      Top             =   9000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   120
      TabIndex        =   58
      Top             =   1080
      Width           =   5295
      Begin VB.TextBox txtClientName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "XXXXXXXXXXXXXXXXX"
         Top             =   120
         Width           =   4080
      End
      Begin VB.TextBox txtJobNumber 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "0000-000-00 RL#00"
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtProtocol 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   600
         Width           =   3945
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job No.:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Protocol:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   690
      End
   End
   Begin VB.Frame FrameClientInventory 
      Height          =   1335
      Left            =   120
      TabIndex        =   53
      Top             =   3720
      Width           =   5295
      Begin VB.CommandButton cmdReason 
         Caption         =   "Reason"
         Height          =   360
         Left            =   4080
         TabIndex        =   8
         Top             =   720
         Width           =   990
      End
      Begin VB.CheckBox chkUseClientInventory 
         Caption         =   "Use Client Inventory"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "000000"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtStockNo 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "XXXXXXXXXXXX"
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clinical Supply Labels:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Label Stock:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.TextBox txtNonBillableNotice 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   6600
      TabIndex        =   52
      Text            =   "NON-BILLABLE"
      Top             =   9350
      Width           =   1215
   End
   Begin VB.Frame fraDescription 
      Height          =   4215
      Left            =   5520
      TabIndex        =   39
      Top             =   3000
      Width           =   5055
      Begin VB.CommandButton cmdSpecInst 
         Caption         =   "Special Instructions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   3700
         Width           =   1725
      End
      Begin VB.CommandButton cmdAddtlData 
         Caption         =   "Client Data"
         Height          =   315
         Left            =   3840
         TabIndex        =   5
         Top             =   3220
         Width           =   975
      End
      Begin VB.TextBox txtReferanceNo 
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "NOTE: Reference No. will be appended to the Description."
         Top             =   3240
         Width           =   3480
      End
      Begin VB.TextBox txtProdDesc 
         Height          =   2460
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label lblReferenceNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No.:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   3000
         Width           =   1110
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtPDRStatus 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   3240
      TabIndex        =   47
      Text            =   "COMBINED/ PDR HAS BEEN PROCESSED"
      Top             =   9350
      Width           =   3255
   End
   Begin VB.TextBox txtReplacement 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1920
      TabIndex        =   46
      Text            =   "REPLACEMENT"
      Top             =   9350
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1085
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   10605
      Begin VB.TextBox txtBarcodeId 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "PDR______"
         Top             =   80
         Width           =   10230
      End
      Begin VB.TextBox txtProducedBy 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "XXXXX XXXXXXXXXX"
         Top             =   600
         Width           =   10215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Created By:"
         Height          =   195
         Left            =   210
         TabIndex        =   45
         Top             =   405
         Width           =   10245
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1935
      Left            =   5520
      TabIndex        =   37
      Top             =   1080
      Width           =   5055
      Begin VB.CheckBox chkPrintAtPackager 
         Caption         =   "Print at packager"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdViewShipping 
         Caption         =   "View Shipping"
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBComboShip 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4695
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196614
         DataMode        =   2
         Cols            =   3
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnHeaders   =   0   'False
         FieldDelimiter  =   "!"
         FieldSeparator  =   ","
         BackColorEven   =   14737632
         BackColorOdd    =   16777152
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   8281
         _ExtentY        =   529
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   585
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   7680
      Width           =   5295
      Begin VB.CheckBox chkReOrientation 
         Caption         =   "PDR is to be Reoriented"
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
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtScratchStockNo 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "XXXXXXXX"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blinding Laminate:"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   810
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblApplyBlindLam 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apply to Labels and Samples"
         ForeColor       =   &H00008000&
         Height          =   390
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   2295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   5295
      Begin VB.TextBox txtClientTGId 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "XXXXXXXXXXXXXXXXXXX"
         Top             =   1200
         Width           =   4020
      End
      Begin VB.TextBox txtGroupName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   4020
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "XXXXXXXXXXXXXXXXXXX"
         Top             =   240
         Width           =   4020
      End
      Begin VB.Label lblClientTGId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client TG ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment Group:"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   885
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   9360
      TabIndex        =   13
      Top             =   8910
      Width           =   1215
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
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   5295
      Begin VB.TextBox txtOnsertBaseRollStock 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   "RLP12345 (min width: 1.0625)"
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtOnsertDie 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "XXXX.XXXX.XX.XXX.XXXX.XX"
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   4065
      End
      Begin VB.TextBox txtLabelId 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "XXXXXXXXXXXX"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblOnsertBaseRollStock 
         BackStyle       =   0  'Transparent
         Caption         =   "Base Roll Stock:"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   120
         TabIndex        =   66
         Top             =   1920
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblOnsertDie 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Onsert Die:"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Description:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Label ID:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.TextBox txtScratchIRQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Frame fraSamples 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5520
      TabIndex        =   16
      Top             =   7200
      Width           =   5055
      Begin VB.TextBox txtSampleGroups 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2835
         TabIndex        =   10
         Top             =   600
         Width           =   840
      End
      Begin VB.CommandButton cmdSamples 
         Caption         =   "Configure Types"
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtSamples 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2835
         TabIndex        =   9
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label6 
         Caption         =   "Samples:"
         Height          =   270
         Left            =   1680
         TabIndex        =   17
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label lblSampleGroups 
         Caption         =   "Sample Types:"
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   615
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdDeleteProdRun 
      Caption         =   "&Delete"
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
      Left            =   120
      TabIndex        =   15
      Top             =   8910
      Width           =   1095
   End
   Begin VB.TextBox txtDirtyFlag 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   8040
      TabIndex        =   12
      Top             =   8910
      Width           =   1215
   End
   Begin VB.TextBox txtStockIRQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   180
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   285
      Left            =   0
      TabIndex        =   26
      Top             =   9390
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3087
            MinWidth        =   3087
            Text            =   "<User Location>"
            TextSave        =   "<User Location>"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2293
            MinWidth        =   2293
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5909
            MinWidth        =   5909
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4780
            MinWidth        =   2505
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "stsMessage"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuReplacementList 
      Caption         =   "&Replacements"
      Begin VB.Menu mnuReplacements 
         Caption         =   "&Create Production Run"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepProdRuns 
         Caption         =   ""
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuCancelPDR 
         Caption         =   "Cancel PDR"
      End
      Begin VB.Menu mnuViewAuditTrail 
         Caption         =   "View Audit Trail"
      End
      Begin VB.Menu duplicateConfig 
         Caption         =   "&Duplicate Sample Configuration"
      End
      Begin VB.Menu mnuViewComputerizationOrder 
         Caption         =   "Vie&w Computerization Order Form"
      End
   End
   Begin VB.Menu mnuNonBillable 
      Caption         =   "&Billing"
      Begin VB.Menu mnuNonBillableView 
         Caption         =   "View..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmProdPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This is the main form.  Its purpose is to allow a Production Run to be viewed and modified.</summary>
'</comment>

Option Explicit
Option Base 1                              'set array(s) lower bound index at 1

Public sampleTypes As Integer              'total existing sample types for Prod Run
Public quantityTotal As Long               'total existing quantity for Prod Run
Public SampleColumns As Integer            'number of columns in sample
Public codingColumns As Integer            'number of columns in coding file

Public oCollIRQInfo As CCollIRQInfo
Public oIRQInfo As CIRQInfo
Private HoldStockProofId As Long
Private HoldScratchStockProofId As Long

Private mSampleFileName As String           ' local copy to be used to populate global one
Private mReprintFile_id As Long             ' Used to help determine client sample
Private mInputFileName As String
Private booSamplesQTYChanged As Boolean     ' used to determine if the samples quantities have been changed
Private mReprintProdRunId As Long

Public mColClientFields As CColClientReqdFields    ' used to hold Client Required fields and values collection

Private mvarOrigStockProofId As Long
Private mvarOrigScratchStockProofId As Long
Private mvarOrigBlindLamApply As Integer
Private mvarOrigStockId As String
Private mvarOrigScratchStockId As String
Private mvarOrigOnsertDieToolId As Long
Private mvarOrigOnsertDiePartNumber As String
Private mvarOriginalOnsertBaseRollStockComponentId As Long
Private mvarOriginalOnsertBaseRollStockPartNumber As String
Private mvarOriginalOnsertBaseRollStockMinimumWidth As Double

Public booSamplesDirtyFlag As Boolean
Private mvarDebugText As String
Private mvarDebugIcoPath As String
Public mvarBillingDirty As Boolean
Public mvarLinksReportDirty As Boolean
Public mvarPDRCancelDirtyflag As Boolean



Private Function CheckOverrides() As Boolean
    If ProductionRun.NonBillable.PerformOverride = False Then
        CheckOverrides = False
    Else
        CheckOverrides = True
    End If
End Function

Private Sub CheckSave()
    If HasDataChanged Then
        Select Case MsgBox("Do you want to save changes?", vbYesNo + vbQuestion)
            Case vbNo
                Me.txtDirtyFlag = ""
            Case vbYes
                cmdSave_Click
        End Select
    Else
        Me.txtDirtyFlag = ""
    End If
End Sub
Private Sub chkUseClientInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim response As Integer
    If Button = vbLeftButton Then
            
            ProductionRun.UseClientInventory = IIf(Me.chkUseClientInventory.value = 0, True, False)
            Me.chkUseClientInventory.value = Abs(ProductionRun.UseClientInventory)
            
            If IsUseClientInventorySetToDefault Then
                response = MsgBox("Setting Use Client Inventory back to the default value " & _
                    "will erase the existing reason." & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, "Confirm Erase Reason")
                    Select Case response
                        Case vbYes
                            ProductionRun.ResetUseClientInventoryReasonDetails
                            Me.cmdReason.enabled = False
                        Case vbNo
                            ProductionRun.UseClientInventory = Not ProductionRun.UseClientInventory
                            Me.chkUseClientInventory.value = Abs(ProductionRun.UseClientInventory)
                    End Select
                    
            ElseIf ProductionRun.HasUseClientInventoryAccess Then
                Call cmdReason_Click
                Me.cmdReason.enabled = True
                
            Else
                MsgBox "You do not have permission to change 'Use Client Inventory'." & vbCrLf & vbCrLf & "Please contact your supervisor or manager.", vbExclamation + vbOKOnly, "No Permission"
                ProductionRun.UseClientInventory = Not ProductionRun.UseClientInventory
                Me.chkUseClientInventory.value = Abs(ProductionRun.UseClientInventory)
            End If
      
    End If
End Sub

Private Function IsUseClientInventorySetToDefault() As Boolean
    IsUseClientInventorySetToDefault = (ProductionRun.UseClientInventory = ProductionRun.NonBillable.IsBillable)
End Function


Private Sub cmdReason_Click()
    Dim frmReason As New frmReason
    
    frmReason.SetReasonTextStatus Me.chkUseClientInventory.enabled
    frmReason.Show vbModal

    Set frmReason = Nothing
                
End Sub

Private Sub Form_Load()
    
    On Error GoTo Error_this_Sub
    
    CenterForm Me
    Dim strLabelId As String
    Dim lngQtyReq As Long
    Dim lngProofId As Long
    Dim lngRandId As Long
    Dim strGroupNum As String
    Dim i As Long
    Dim strMessageCC As String
    Dim strMessage As String
    Dim booIRQsExist As Boolean
    Dim booChange As Boolean
    Dim strReason As String
    Dim booPdrHasRun As Boolean
    Dim booPdrIsCombined As Boolean
    Dim xmls As New ClintrakCommon.XmlSettings
    Dim strFileNameSamples As String
    Dim strTempFileName As String
        
    booSamplesDirtyFlag = False
    Call Me.txtReplacement.Move(Me.StatusBar1.Panels(2).Left + 50, Me.StatusBar1.Top + 50)
    Call Me.txtPDRStatus.Move(Me.StatusBar1.Panels(3).Left + 50, Me.StatusBar1.Top + 50)
    Call Me.txtNonBillableNotice.Move(Me.StatusBar1.Panels(4).Left + 50, Me.StatusBar1.Top + 50)
    
    Set CCBlindLamApplyCol = New Collection
    Call CCBlindLamApplyCol.Add(LABELS_ONLY_TEXT, "0")
    Call CCBlindLamApplyCol.Add(LABEL_AND_SAMPLES_TEXT, "1")
    Call CCBlindLamApplyCol.Add(NOT_BLINDED_TEXT, "2")
    Call CCBlindLamApplyCol.Add(NA_TEXT, "3")
    
    Set BlindLamApplyCol = New Collection
    Call BlindLamApplyCol.Add(NOT_BLINDED_TEXT, "0")
    Call BlindLamApplyCol.Add(LABEL_AND_SAMPLES_TEXT, "1")
    Call BlindLamApplyCol.Add(LABELS_ONLY_TEXT, "2")
    Call BlindLamApplyCol.Add(NA_TEXT, "3")
           
    'Load debug info from apps.xml
    gDebugMode = xmls.GetSetting("DebugMode")
    xmls.AppName = "ProductionRuns"
    gRnrBaseFilePath = xmls.GetSetting("RNRFilePath")
    gRnrDebugFilePath = gRnrBaseFilePath & "DEBUG\"
    gWebUtilitiesUrl = xmls.GetSetting("WebUtilitiesUrl")
    gWebAuditPRUrl = gWebUtilitiesUrl & "Audit/AuditLog/Index?AuditType=Production%20Run"
     
    ' came from ProdRunMain ------------------
    Call getFileLinksInfo
    Call GetJobInformation
    Call GetClientName
    Call GetRandIDNumber
    Call GetRandDelimiter
    Call GetLabelCurrentValues

    'set this flag each time this screen is entered
    booReplacement = False
    Me.txtReplacement.Visible = False
    Me.mnuReplacements.enabled = False
    Me.mnuViewAuditTrail.enabled = False
    booSamplesQTYChanged = False
    booPdrIsCombined = False
    Call Enable_Fields
    
    Set ProductionRun = New CProdrun
    ProductionRun.NonBillable.NoticeForm = frmProdPlan
    Set Planning = New CPlanningMethods
        
    Screen.MousePointer = vbHourglass
    
    Me.txtClientName = gClientName
    Me.txtJobNumber = gJobNumber
    Me.txtProtocol = gProtocol
        
    ProductionRun.Production_Run_Id = gProductionRun_Id
    lngProofId = gProofId
    lngRandId = gRandomizationId
    lngQtyReq = gQuantity
    strLabelId = gLabelId
    Me.txtFileName = GetFileNameFromFilePath(gCodingFileName)
    Me.txtGroupName = gGroupName
    Me.txtClientTGId = gClientGroupName
    strGroupNum = gGroupNumber
    'loads the ship to combobox with all associated shipping addresses
    Call LoadShipToCombo2Column(Me.SSDBComboShip)
    
    HoldStockProofId = 0
    HoldScratchStockProofId = 0
    
    If ProductionRun.Production_Run_Id <> 0 Then

        If gDebugMode = True Then

            ProductionRun.LookupRecord
            'if coding path doesn't have debug, lets copy coding files

            If InStr(1, ProductionRun.File_Name, gRnrDebugFilePath, vbTextCompare) <> 1 Then

                strTempFileName = Replace$(ProductionRun.File_Name, gRnrBaseFilePath, gRnrDebugFilePath, 1, -1, vbTextCompare)
                'coding doesn't have debug in the path must copy files and update coding db path
                Call CopyFiles(ProductionRun.File_Name, strTempFileName)
                Call UpdateCodingFilePath(gCodingFileId, ProductionRun.Production_Run_Id, strTempFileName, 0)
                
                'create sample path from coding path and check if it's debug
                strFileNameSamples = Replace$(ProductionRun.File_Name, "\cod\", "\smp\", 1, -1, vbTextCompare)  'it's already doing this in the code
                strFileNameSamples = GetFilePath(strFileNameSamples) 'drop the file in the path
                strTempFileName = Replace$(strFileNameSamples, gRnrBaseFilePath, gRnrDebugFilePath, 1, -1, vbTextCompare)
                ProductionRun.LookupRecord
                
                Call CopySampleFiles(strFileNameSamples, strTempFileName, ProductionRun.Barcode_Id)
            End If

        Else
            ProductionRun.LookupRecord
        End If
    
    
        'if the production run is not found then initialize the screen and populate some
        'of the values from the labels/files screen
        'else populate then screen with the data from the produciton run
        If ProductionRun.Prod_Run_Found Then
            booNewProdRun = False
            Me.chkReOrientation.value = ProductionRun.Reorient_Ind
            Me.txtQty = ProductionRun.Qty_Requested
            Me.txtSamples = ProductionRun.Samples_Requested
            Me.txtSampleGroups = ProductionRun.Sample_Number
            Me.txtStockNo.text = ProductionRun.stock
            Me.txtOnsertDie = ProductionRun.OnsertDiePartNumber
            HoldStockProofId = ProductionRun.Stock_Proof_Id
            Me.txtProducedBy = ProductionRun.Produced_By
            Me.txtScratchStockNo.text = ProductionRun.Scratch_Stock
            Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
            HoldScratchStockProofId = ProductionRun.Scratch_Proof_Id

            Me.lblApplyBlindLam.Caption = BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff))
            If Me.lblApplyBlindLam.Caption = NA_TEXT Or Me.lblApplyBlindLam.Caption = NOT_BLINDED_TEXT Then
                Me.lblApplyBlindLam.Caption = ""
            End If
                    
            Call SetSSDBComboText(Me.SSDBComboShip, ProductionRun.Ship_To_Id)
            Me.txtProdDesc = ProductionRun.Prod_Description
            Me.txtReferanceNo = ProductionRun.Reference_No
            Me.txtBarcodeId = ProductionRun.Barcode_Id
            
            Set gBarcodeFile = New CodingFile
            gBarcodeFile.Init gApplicationUser
            gBarcodeFile.LoadByPDR ProductionRun.Barcode_Id
            
            If Not CheckShippingExist(gJob_Id) Then  'checks to see whether the Shipping Address has been selected
                MsgBox _
                    "The Shipping Address has NOT been selected!" & vbCrLf & _
                    "Please Select a Shipping Address before making changes", vbExclamation
                ActivateEditOption (False)
            Else
                ActivateEditOption (True)
            End If
            Call Load_Menu   'load any associated production runs
            If lngQtyReq <> ProductionRun.Qty_Requested Then
                If MsgBox( _
                    "The quantity on the selected coding file does not match the existing Production Run quantity." & vbCrLf _
                    & "The Production Run quantity will be replaced with the Quantity from the selected file." & vbCrLf _
                    & vbTab & "EXISTING PRODUCTION RUN QTY: " & ProductionRun.Qty_Requested & vbCrLf _
                    & vbTab & "FILE RECORD QTY: " & lngQtyReq & vbCrLf & vbCrLf _
                    & "Choose YES to accept this replacement NO to leave the quantity as is." & vbCrLf _
                    & "(NOTE: You must save the Production run for this change to become permanent!)", _
                    vbQuestion + vbYesNo) = vbYes Then
               
                    ProductionRun.Qty_Requested = lngQtyReq
                    Me.txtDirtyFlag = "Y"
                    mvarLinksReportDirty = True
                    Me.txtProducedBy = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
                End If
            End If
            
            Call UpdateExistingIRQList

        Else
            MsgBox _
                "Production run was not found and one should exist for this Link." & vbCrLf & _
                "Please contact IT with this error.", vbCritical
            Exit Sub
        End If
    Else
        booNewProdRun = True
        Me.txtDirtyFlag = "Y"
        Me.mvarLinksReportDirty = True
        Me.txtProducedBy = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
        Me.cmdSpecInst.enabled = False
        ProductionRun.Job_Log_Id = gJob_Id
        ProductionRun.Randomization_Id = lngRandId
        ProductionRun.Client_Id = gClientId
        ProductionRun.Qty_Requested = lngQtyReq
        ProductionRun.Proof_Id = lngProofId
        ProductionRun.Campaign_No = strGroupNum
        ProductionRun.File_Name = gCodingFileName
        ProductionRun.OnsertBaseRollStockComponentId = gOnsertBaseRollStockComponentId
        ProductionRun.OnsertBaseRollStockMinimumWidth = gOnsertBaseRollStockMinimumWidth
        ProductionRun.OnsertBaseRollStockPartNumber = gOnsertBaseRollStockPartNumber
       
        
        Me.txtStockNo.text = gStockLabelId
        ProductionRun.stock = gStockLabelId
        ProductionRun.Stock_Proof_Id = gStockProofId
        Me.txtScratchStockNo.text = gBlindLabelId
        ProductionRun.Scratch_Stock = gBlindLabelId
        ProductionRun.Scratch_Proof_Id = gBlindProofId
        Me.txtOnsertDie.text = gOnsertDiePartNumber
        Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock()
        
        ProductionRun.OnsertDiePartNumber = gOnsertDiePartNumber
        ProductionRun.OnsertDieToolId = gOnsertDieToolId
        
        Set gBarcodeFile = New CodingFile
        gBarcodeFile.Init gApplicationUser
        gBarcodeFile.LoadByCDF "", gCodingFileName
        
        txtSamples.text = "0"
        txtSampleGroups.text = "0"
        ProductionRun.NonBillable.SetNew ProductionRun.Job_Log_Id
        
        Me.lblApplyBlindLam.Caption = CCBlindLamApplyCol(CStr(gBlindLamApply))
        If Me.lblApplyBlindLam.Caption = NA_TEXT Or Me.lblApplyBlindLam.Caption = NOT_BLINDED_TEXT Then
            Me.lblApplyBlindLam.Caption = ""
        End If
        
        For i = 1 To BlindLamApplyCol.count
            If BlindLamApplyCol(i) = CCBlindLamApplyCol(CStr(gBlindLamApply)) Then
                ProductionRun.Apply_ScratchOff = i - 1
                Exit For
            End If
        Next i
    End If

    'Load Client Required fields and values
    Call LoadClientLabelFields(gClientId, ProductionRun.Production_Run_Id)
    
    If ProductionRun.Production_Run_Id > 0 Then
        Call LoadBarcodeInfo(ProductionRun.Production_Run_Id)
    End If
    
    'Initialize Holder
    Set mColClientFields = ClientReqdFields.Clone
    
    If gClientReqFieldInd = True Or ClientReqdFields.count > 0 Then
        Me.cmdAddtlData.enabled = True
    Else
        Me.cmdAddtlData.enabled = False
    End If
    
    'need to call the label proof table to get the label description of the label id
    ProductionRun.GetLabelDesc
    Me.txtDesc = ProductionRun.LabelDescription
    Me.txtLabelId = strLabelId
    Me.txtQty = ProductionRun.Qty_Requested
    
    If booNewProdRun Then
        Me.txtProdDesc = ProductionRun.LabelDescription & " - " & Me.txtGroupName
    End If
    
    ProductionRun.Client_Name = Me.txtClientName
    
    Call CheckColumnNumbers
    
    ' This option is only enabled for Replacements
    Me.chkPrintAtPackager.enabled = False
       
    'md added for clintrak samples - determine if the PDR has been RUN.  If so, block all
    'fields on the form.
    If Not booNewProdRun Then
        If Determine_If_PDR_HasRun Then
            booPdrHasRun = True
            Call Lock_Out_PDRForm(False, True)
            
            ' check to see if the PDR is on a PKS
            ' b/c the Job Shipping Flag isn't the most reliable indicator
            If Determine_Shipping_Flag_On Or DeterminePDROnPKS Then
                Me.txtProdDesc.enabled = False
                SSDBComboShip.enabled = False
                cmdSave.enabled = False
                ' check here to see if any samples are on PKS
                ' and if so, disable the button.  only do the check if the
                ' number of sample types is greater than 1
                If Me.txtSampleGroups.text > 1 Then
                    If DeterminePDROnPKS("S") Then
                        cmdSamples.enabled = False
                    End If
                Else
                    Me.cmdSamples.enabled = False
                End If
                chkReOrientation.enabled = False
                Me.txtReferanceNo.enabled = False
                If gClientReqFieldInd = True And ClientReqdFields.count > 0 Then
                    If ClientReqdFields.Item(1).Production_Run_Client_Fields_Id = 0 Then
                        Me.cmdAddtlData.enabled = False
                    End If
                End If
            End If
            If ProductionRun.Barcode_Id <> "" Then
                If Planning.CheckIfCombined(ProductionRun.Barcode_Id) Then
                    booPdrIsCombined = True
                    txtPDRStatus.text = "COMBINED PDR HAS BEEN PROCESSED"
                End If
            End If
        Else
            booPdrHasRun = False
            Call Lock_Out_PDRForm(True, False)
            Me.txtProdDesc.enabled = True
            SSDBComboShip.enabled = True
            cmdSave.enabled = True
            cmdSamples.enabled = True
            chkReOrientation.enabled = True
            Me.txtReferanceNo.enabled = True
            
            If ProductionRun.Barcode_Id <> "" Then
                If Planning.CheckIfCombined(ProductionRun.Barcode_Id) Then
                    booPdrIsCombined = True
                    Call Lock_Out_PDRForm(False, True)
                    txtProdDesc.enabled = True
                    txtPDRStatus.text = "COMBINED PDR"
                End If
            End If
        End If
    Else
        booPdrHasRun = False
        Call Lock_Out_PDRForm(True, False)
        SSDBComboShip.enabled = True
        cmdSamples.enabled = True
        cmdSave.enabled = True
        chkReOrientation.enabled = True
        Me.txtReferanceNo.enabled = True
        
        If ProductionRun.Barcode_Id <> "" Then
            If Planning.CheckIfCombined(ProductionRun.Barcode_Id) Then
                booPdrIsCombined = True
                Call Lock_Out_PDRForm(False, True)
                cmdSamples.enabled = False
                txtProdDesc.enabled = True
                txtPDRStatus.text = "COMBINED PDR"
            End If
        End If
    End If
    

    mvarOrigStockProofId = 0
    mvarOrigStockId = ""
    mvarOrigScratchStockProofId = 0
    mvarOrigScratchStockId = ""
    mvarOrigBlindLamApply = 0
    mvarOrigOnsertDieToolId = 0
    mvarOrigOnsertDiePartNumber = ""
    mvarOriginalOnsertBaseRollStockMinimumWidth = 0
    mvarOriginalOnsertBaseRollStockComponentId = 0
    mvarOriginalOnsertBaseRollStockPartNumber = ""
    booIRQsExist = False
    
    ' when loading an existing PDR is not on a packing slip, check if the stock/blinding
    ' lam/blinding lam application matches what is saved in Label Specs for the label
    ' and if not and it is not grouped, has no IRQs, and is not approved, ask user if
    ' they would like to change it so it does match.  If it doesn't match and it has run,
    ' has an IRQ, is grouped, or approved, just display a message letting them know it
    ' doesn't match but can't be fixed because of one of those reasons listed.
    If booNewProdRun = False And DeterminePDROnPKS = False Then
        If Not oCollIRQInfo Is Nothing Then
            If oCollIRQInfo.count > 0 Then
                booIRQsExist = True
            End If
        End If
        ' capture the last stage the PDR has reached.  This will be used to:
        ' a) Determine if the stock/blinding lam info be update
        ' b) If the stock/blinding lam can't be updated, it will give the user the reason
        '       it can't be updated to it they can to have it updated, they know where in
        '       the process the PDR is and they can start from there to get it to a position
        '       where it can be updated.
        strReason = ""
        
'        If InStr(txtPDRStatus.Text, "PROCESSED") > 0 And Me.txtPDRStatus.Visible = True Then
        If booPdrHasRun Then
            strReason = "has been run."
        ElseIf booIRQsExist Then
            strReason = "has an IRQ."
'        ElseIf InStr(txtPDRStatus.Text, "COMBINED") > 0 And Me.txtPDRStatus.Visible = True Then
        ElseIf booPdrIsCombined Then
            strReason = "is part of a PRG."
        ElseIf ProductionRun.ApprovalDate <> "1/1/1900" Then
            strReason = "has been Approved."
        End If
        
        If (ProductionRun.Stock_Proof_Id <> gStockProofId) Then
            If strReason = "" Then
                If MsgBox("The PDR's stock is currently " & ProductionRun.stock & "." & vbCrLf & _
                        "Label Specs's stock is currently " & gStockLabelId & "." & vbCrLf & vbCrLf & _
                        "Do you want to update the PDR Stock to match Label Specs?", _
                        vbYesNo + vbQuestion, "Update Stock") = vbYes Then
                    Me.txtStockNo.text = gStockLabelId
                    ProductionRun.stock = gStockLabelId
                    ProductionRun.Stock_Proof_Id = gStockProofId
                    
                    MsgBox "The PDR's stock has been updated to match Label Specs." & vbCrLf & _
                                "You must SAVE the PDR in order to save this change.", vbInformation, "Stock Updated"
            
                    booChange = True
                    mvarLinksReportDirty = True
                End If
            Else
                MsgBox "The PDR's stock is currently " & ProductionRun.stock & "." & vbCrLf & _
                        "Label Specs's stock is currently " & gStockLabelId & "." & vbCrLf & vbCrLf & _
                        "The PDR's stock CANNOT be updated at this time because it " & strReason, _
                        vbOKOnly + vbInformation, "Stock Difference"
            End If
                    
        End If
        ' the blinding laminate and blinding laminate application will go hand in hand
        If (ProductionRun.Scratch_Proof_Id <> gBlindProofId) Or (gBlindProofId > 0 And (BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff)) <> CCBlindLamApplyCol(CStr(gBlindLamApply)))) Then
            
            strMessageCC = gBlindLabelId
            If gBlindLabelId <> "N/A" Then
                strMessageCC = strMessageCC & " (" & CCBlindLamApplyCol(CStr(gBlindLamApply)) & ")"
            End If
            
            strMessage = ProductionRun.Scratch_Stock
            If ProductionRun.Scratch_Stock <> "N/A" Then
                strMessage = strMessage & " (" & BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff)) & ")"
            End If
            
            If strReason = "" Then
                If MsgBox("The PDR's Blinding Laminate is currently " & strMessage & "." & vbCrLf & _
                        "Label Specs's Blinding Laminate is currently " & strMessageCC & "." & vbCrLf & vbCrLf & _
                        "Do you want to update the PDR Blinding Laminate to match Label Specs?", _
                        vbYesNo + vbQuestion, "Update Blinding Laminate") = vbYes Then
                     
                    Me.lblApplyBlindLam.Caption = CCBlindLamApplyCol(CStr(gBlindLamApply))
                    If Me.lblApplyBlindLam.Caption = NA_TEXT Or Me.lblApplyBlindLam.Caption = NOT_BLINDED_TEXT Then
                        Me.lblApplyBlindLam.Caption = ""
                    End If
        
                    For i = 1 To BlindLamApplyCol.count
                        If BlindLamApplyCol(i) = CCBlindLamApplyCol(CStr(gBlindLamApply)) Then
                            ProductionRun.Apply_ScratchOff = i - 1
                            Exit For
                        End If
                    Next i
                    
                    Me.txtScratchStockNo.text = gBlindLabelId
                    ProductionRun.Scratch_Stock = gBlindLabelId
                    ProductionRun.Scratch_Proof_Id = gBlindProofId
                    
                    MsgBox "The PDR's Blinding Laminate has been updated to match Label Specs." & vbCrLf & _
                                "You must SAVE the PDR in order to save this change.", vbInformation, "Blinding Laminate Updated"
                    
                    booChange = True
                    mvarLinksReportDirty = True
                End If
            
            Else
            
                MsgBox "The PDR's Blinding Laminate is currently " & strMessage & "." & vbCrLf & _
                        "Label Specs's Blinding Laminate is currently " & strMessageCC & "." & vbCrLf & vbCrLf & _
                        "The PDR's Blinding Laminate CANNOT be updated at this time because it " & strReason, _
                        vbOKOnly + vbInformation, "Blinding Laminate Difference"
            
            End If
            
        End If
        
        ' Check for changes to the Onsert die after the PDR has been created
         If (ProductionRun.OnsertDieToolId <> gOnsertDieToolId) Then
            If strReason = "" Then
                If MsgBox( _
                    "The PDR's Onsert die is currently " & ProductionRun.OnsertDiePartNumber & "." & vbCrLf & _
                    "Label Specs's Onsert die is currently " & gOnsertDiePartNumber & "." & vbCrLf & vbCrLf & _
                    "Do you want to update the PDR's die to match Label Specs?", vbYesNo + vbQuestion, "Update Die") = vbYes Then
                    Me.txtOnsertDie.text = gOnsertDiePartNumber
                    ProductionRun.OnsertDiePartNumber = gOnsertDiePartNumber
                    ProductionRun.OnsertDieToolId = gOnsertDieToolId
                    MsgBox _
                        "The PDR's Onsert die has been updated to match Label Specs." & vbCrLf & _
                        "You must SAVE the PDR in order to save this change.", vbInformation, "Die Updated"
                    booChange = True
                End If
            Else
                MsgBox "The PDR's Onsert die is currently " & ProductionRun.OnsertDiePartNumber & "." & vbCrLf & _
                        "Label Specs's Onsert die is currently " & gOnsertDiePartNumber & "." & vbCrLf & vbCrLf & _
                        "The PDR's die CANNOT be updated at this time because it " & strReason, _
                        vbOKOnly + vbInformation, "Die Difference"
            End If
        End If
        
        If (ProductionRun.OnsertBaseRollStockComponentId <> gOnsertBaseRollStockComponentId) Then
            If basGlobals.ShowChangePromptForLabel("Onsert base roll stock", FormatEmptyStringNA(ProductionRun.OnsertBaseRollStockPartNumber), FormatEmptyStringNA(gOnsertBaseRollStockPartNumber), strReason) Then
                ProductionRun.OnsertBaseRollStockPartNumber = gOnsertBaseRollStockPartNumber
                ProductionRun.OnsertBaseRollStockComponentId = gOnsertBaseRollStockComponentId
                Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
            End If
        End If
        
        If (ProductionRun.OnsertBaseRollStockMinimumWidth <> gOnsertBaseRollStockMinimumWidth) Then
            If basGlobals.ShowChangePromptForLabel("Onsert base roll stock minimum width", FormatZeroNumberNA(ProductionRun.OnsertBaseRollStockMinimumWidth), FormatZeroNumberNA(gOnsertBaseRollStockMinimumWidth), strReason) Then
                ProductionRun.OnsertBaseRollStockMinimumWidth = gOnsertBaseRollStockMinimumWidth
                Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
            End If
        End If
        
        If booChange = True Or (mvarLinksReportDirty And booReplacement = False) Then
            Me.txtDirtyFlag.text = "Y"
            Me.txtProducedBy = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
        End If
    End If
        
    ' DW 2012-001 added
    If txtPDRStatus.text = "COMBINED PDR" Then Me.chkReOrientation.enabled = False
    
    If ProductionRun.StatusLookupId = basGlobals.GetStatusID(CANCEL_STATUS) Then
        Me.mnuCancelPDR.Caption = UNCANCEL_TEXT
        Me.txtPDRStatus.text = STATUS_FOOTER_TEXT
        Me.txtPDRStatus.Visible = True
    End If
     
    ' DW 2008-017 added
    Me.StatusBar1.Panels(1).text = gClintrakLocations(gApplicationUser.ClintrakLocationId).Display & " User "
    Me.StatusBar1.Panels(1).Picture = LoadPicture(gDomainIconPath & gApplicationUser.ClintrakLocationId & ".ico")
    
    mvarDebugText = xmls.GetSetting("DebugText")
    mvarDebugIcoPath = xmls.GetSetting("DebugIcoPath")
    
    If Len(mvarDebugText) > 0 Then
        StatusBar1.Panels("stsMessage").Visible = True
        StatusBar1.Panels("stsMessage").text = mvarDebugText
        StatusBar1.Panels("stsMessage").ToolTipText = mvarDebugText
        
        If Dir(mvarDebugIcoPath) <> "" Then
            StatusBar1.Panels("stsMessage").Picture = LoadPicture(mvarDebugIcoPath)
        End If
        
    Else
        StatusBar1.Panels("stsMessage").Visible = False
    End If
    
    ManageClientInventoryCheckBox
' If not print at packager check for samples . Allow correction if file length is 0
    If Not CBool(Me.chkPrintAtPackager.value) Then
        If Me.txtSamples > 0 Then
            If FileLen(gSampleFileName) = 0 Then
               If vbYes = MsgBox("The Internal Sample Files were not found. Do you want the system to re-create the Internal Samples based on today’s data?", vbYesNo) Then
                    If Fix_Samples_FromPDR_Screen Then
'change PDR sample counts for new samples just configured
                        ProductionRun.Samples_Requested = CLng(Me.txtSamples)
                        ProductionRun.Sample_Number = CInt(Me.txtSampleGroups)
                        MsgBox "Internal Samples Fixed", vbOKOnly
                    Else
                        MsgBox _
                            "Unable to Fix Internal Samples" & vbCrLf & _
                            "Please contact IT.", vbExclamation
                    End If
                End If
            End If
        End If
    End If
Exit_this_Sub:
    Screen.MousePointer = vbDefault
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR in Loading Prod Run"
    Resume Exit_this_Sub
    
End Sub

Private Sub cmdAddtlData_Click()
    Load frmClientReqdFields
    frmClientReqdFields.Show vbModal
End Sub

Private Sub cmdSamples_Click()
    Dim lngExistingSampleQTY As Long
    Dim lngNewSampleQTY As Long
    Dim lngExistingSampleTypes As Long
    Dim lngNewSampleTypes As Long
    
    ' Checks to see whether the production run exists or not
    If Not CheckProdRunExist Then
        Me.txtSamples = 0
        Me.txtSampleGroups = 0
        MsgBox _
            "Please save Production Run First Before Configuring Samples.", vbExclamation
        Exit Sub
    End If
        
    If Not CheckExistingSample Then
        If ProductionRun.PrintAtPackager Then
            MsgBox _
            "There are no samples to configure for this run.", vbInformation
            Exit Sub
        End If
    End If
   
    lngExistingSampleQTY = CLng(Me.txtSamples.text)
    lngExistingSampleTypes = CLng(Me.txtSampleGroups.text)
    
    frmSmpConfig.Show vbModal
    
    lngNewSampleQTY = CLng(Me.txtSamples.text)
    lngNewSampleTypes = CLng(Me.txtSampleGroups.text)
    
    ' Check to see if the sample quantities changed
    If (lngExistingSampleQTY <> lngNewSampleQTY) Or (lngExistingSampleTypes <> lngNewSampleTypes) Or booSamplesDirtyFlag = True Then
        Me.txtProducedBy.text = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
        ProductionRun.Produced_By = Me.txtProducedBy
        ProductionRun.UpdateSampleQuantities
        booSamplesDirtyFlag = False
    End If
        
End Sub

Private Sub cmdSpecInst_Click()

    Load frmSpecInst
    frmSpecInst.txtInstructions = ProductionRun.Special_Inst
    ' Do not allow users to save changes if the PDR has run
    frmSpecInst.OKButton.enabled = Not Me.txtPDRStatus.Visible
    frmSpecInst.Show vbModal

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   0 - The user has chosen the Close command from the Control-menu box on the form, or hits the big X on the other side.
'   1 - The Unload method has been invoked from code.
'   2 - The current Windows-environment session is ending.
'   3 - The Microsoft Windows Task Manager is closing the application.
'   4 - An MDI child form is closing because the MDI form is closing.
    If ProductionRun.IsProductionRunDeleted Then Exit Sub
    CheckSave
    
    ' DW 2010-002 added
    Select Case UnloadMode
        Case 0, 1
            If Me.txtDirtyFlag = "Y" Then Cancel = True
    End Select
    
End Sub

Private Sub mnuCancelPDR_Click()
     
    If IsThereAnIRQ Then
        MsgBox "The PDR cannot be cancelled because it already has IRQs." & vbCrLf & "Have the PDR removed from its IRQs and try again.", vbOKOnly + vbInformation, "Cancel PDR"
        Exit Sub

    End If

    If ProductionRun.IsOnRushRequest Then
        MsgBox "The PDR cannot be cancelled because it is on a rush request." & vbCrLf & "Remove the PDR from its rush request and try again.", vbOKOnly + vbInformation, "Cancel PDR"
        Exit Sub

    End If

    If ProductionRun.IsScheduled Then
        MsgBox "The PDR cannot be cancelled because it is scheduled." & vbCrLf & "Have the PDR unscheduled and try again.", vbOKOnly + vbInformation, "Cancel PDR"
        Exit Sub

    End If
    
    If ProductionRun.IsCombined Then
        MsgBox "The PDR cannot be cancelled because it is part of a PRG." & vbCrLf & "Delete the PRG and try again.", vbOKOnly + vbInformation, "Cancel PDR"
        Exit Sub

    End If
  
    If Me.mnuCancelPDR.Caption = CANCEL_TEXT Then
        CancelPDR
    Else
        UNCancelPDR

    End If
    
End Sub

Private Sub CancelPDR()
    frmCancelNote.ShowCancelNote

    If frmProdPlan.mvarPDRCancelDirtyflag = True Then
        Me.mnuCancelPDR.Caption = UNCANCEL_TEXT
        Me.txtPDRStatus.text = STATUS_FOOTER_TEXT
        Me.txtPDRStatus.Visible = True
        txtDirtyFlag.text = "Y"
        frmProdPlan.mvarLinksReportDirty = True
    End If
    
End Sub

Private Sub UNCancelPDR()
    Me.mnuCancelPDR.Caption = CANCEL_TEXT
    txtPDRStatus.text = ""
    ProductionRun.CancellationReasonLookupId = 0
    ProductionRun.CancellationNotes = "N/A"

    If (ProductionRun.ApprovalDate <> "1/1/1900") Or booReplacement Then
        ProductionRun.StatusLookupId = basGlobals.GetStatusID(APPROVED_STATUS)
    Else
        ProductionRun.StatusLookupId = basGlobals.GetStatusID(AWAITINGAPPROVAL_STATUS)

    End If
    txtDirtyFlag.text = "Y"
End Sub

Private Sub mnuNonBillableView_Click()
    Dim originalIsBillable As Boolean
    Dim priorBillableStatusSet As Boolean
    originalIsBillable = ProductionRun.NonBillable.IsBillable
    priorBillableStatusSet = ProductionRun.NonBillable.IsBillableStatusSet
    
    frmNonBillable.ShowNonBillable
    
    If Me.mvarBillingDirty And (Not cmdSave.enabled) Then
        cmdSave.enabled = True
    End If
    
    If Not IsThereAnIRQ Then
        If priorBillableStatusSet <> ProductionRun.NonBillable.IsBillableStatusSet Or (originalIsBillable <> ProductionRun.NonBillable.IsBillable) Then
            If ProductionRun.UseClientInventoryReason <> "" Then
                MsgBox "The prior reason entered for the Use Client Inventory change will be deleted.", vbInformation, "Reason Cleared by Billing Change"
                ProductionRun.ResetUseClientInventoryReasonDetails
            End If
            ProductionRun.UseClientInventory = ProductionRun.NonBillable.IsBillable
            Me.chkUseClientInventory.value = Abs(ProductionRun.UseClientInventory)
            Me.cmdReason.enabled = False
        End If
    End If
End Sub
Private Sub mnuViewAuditTrail_Click()
    Dim sWebAuditPRUrl As String
    sWebAuditPRUrl = gWebAuditPRUrl & Chr(38) & "id=" & ProductionRun.Production_Run_Id
    Shell "cmd /c start " & Chr(34) & Chr(34) & " " & Chr(34) & sWebAuditPRUrl & Chr(34), vbHide
End Sub

Private Sub mnuViewComputerizationOrder_Click()

 Dim rptProdPlan As New ARProdPlan
 Dim oPDF As New ARExportPDF

    If ProductionRun.Production_Run_Id > 0 Then
        
        'md added for clintrak samples
        Call ProductionRun.Determine_Clintrak_Samples
        
        If Planning.CheckIfCombined(txtBarcodeId) = True Then
            gShowQuarantineTagFlag = False
        Else
          gShowQuarantineTagFlag = True
        End If
                
        rptProdPlan.Printer.Orientation = ddOLandscape
        
        ' CC v2.7.3 disable Print button and other buttons on ActiveReport Viewer
        rptProdPlan.Toolbar.Tools(1).Visible = False
        rptProdPlan.Toolbar.Tools(2).Visible = False
        rptProdPlan.Toolbar.Tools(4).Visible = False
        rptProdPlan.Toolbar.Tools(0).Visible = False
        
        rptProdPlan.Show vbModal
    
    Else
        MsgBox "Create the Production Run first by clicking the Save Button.", vbInformation + vbOKOnly, "Production Run Error"
    End If
End Sub

Private Sub cmdDeleteProdRun_Click()
    Dim digitalOverageOrder As String
    
    On Error GoTo Error_this_Sub
    
    ' Validate Delete
    ' -------------------------------------------------
    If ProductionRun.Production_Run_Id = 0 Then
        GoTo Exit_this_Sub
    End If
    
    If ProductionRun.ApprovalDate <> "1/1/1900" Then
        MsgBox _
            "This PDR cannot be deleted because it is already approved." & vbCrLf & _
            "If possible, unapprove the PDR/Links and then try again.", vbExclamation
        Exit Sub
    End If
    
    ' to not allow delete if an IRQ exists
    If Me.txtStockIRQ <> "" And Me.txtScratchIRQ <> "" Then
        MsgBox _
            "This PDR cannot be deleted." & vbCrLf & _
            "IRQs have already been created for Stock and Blinding Laminates.", _
            vbExclamation
        Exit Sub
    End If
    
    If Me.txtStockIRQ <> "" Then
        MsgBox _
            "This PDR cannot be deleted." & vbCrLf & _
            "An IRQ has already been created.", vbExclamation
        Exit Sub
    End If
    
    If Me.txtScratchIRQ <> "" Then
        MsgBox _
            "This PDR cannot be deleted." & vbCrLf & _
            "An IRQ has already been created for Blinding Laminates.", vbExclamation
        Exit Sub
    End If
    
    ' check for existing replacements b/c when the main PDR is deleted,
    ' it orphans the replacements and throughs things out of whack.
    If Not booReplacement And Me.mnuRepProdRuns(0).Visible = True Then
        MsgBox _
            "This PDR cannot be deleted." & vbCrLf & _
            "A Replacement PDR has already been created.", vbExclamation
        Exit Sub
    End If

    ' Check to see if the PDR is associated to a Digital Overage Order and prompt user to unassociate it
    digitalOverageOrder = checkForAssociatedDigitalOverageOrder(ProductionRun.Barcode_Id)
    If Not digitalOverageOrder = "N/A" Then
        MsgBox _
            "This PDR cannot be deleted." & vbCrLf & _
            "An associated Digital Overage Order exists: " & digitalOverageOrder, vbExclamation
        Exit Sub
    End If
    
    If ProductionRun.IsOnRushRequest Then
        MsgBox "This PDR cannot be deleted." & vbCrLf & _
        "It is associated to a rush request.", vbExclamation
        Exit Sub
    End If
    
    If gLinkLock.IsLockStillValid = False Then
        MsgBox "You cannot delete this PDR because your lock on the Links has been deleted." & vbCrLf & vbCrLf & "Please reload before continuing.", vbInformation + vbOKOnly, "Invalid Lock"
        Exit Sub
    End If

    If MsgBox("Are you sure you want to Delete this Production Run?", _
        vbQuestion + vbYesNo) = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    ' Process Delete
    ' -------------------------------------------------
    Screen.MousePointer = vbHourglass
    
    ProductionRun.Prod_Run_Updated = False
    ProductionRun.DeleteRecord
    If ProductionRun.Prod_Run_Updated Then
        MsgBox "Production Run Record was Successfully Deleted.", vbInformation
        Call UpdateCompletelyShippedFlag(ProductionRun.Job_Log_Id)
        txtDirtyFlag = ""
        mvarLinksReportDirty = False
    Else
        MsgBox "Production Run Record was NOT Deleted." & Chr$(13) & "Check data and retry.", vbCritical
    End If
    
Exit_this_Sub:
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR in Deleting Prod Run"
    Resume Exit_this_Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    ' clean up all global variables here
    Set CCBlindLamApplyCol = Nothing
    Set BlindLamApplyCol = Nothing
    Set oCollIRQInfo = Nothing
    Set oIRQInfo = Nothing
    Set mColClientFields = Nothing
    Set BarcodeInfo = Nothing
    Set dupSameCodingData = Nothing
    Set smpSameCodingData = Nothing
    Set BarcodeInfo = Nothing
    Set ProductionRun = Nothing
    Set mData = Nothing
    Set dupData = Nothing
    Set smpData = Nothing
    Set Planning = Nothing
    Set PlanningList = Nothing
    Set ProductionGroup = Nothing
    Set gApplicationUser = Nothing
    Set ClientReqdFields = Nothing
    Set gadoConnection = Nothing
    Set gClintrakLocations = Nothing
    booReplacement = False
    gJob_Id = 0
    gJobNumber = ""
    gProtocol = ""
    gFileLinksId = 0
    gRandomizationId = 0
    booNewProdRun = False
    gClientName = ""
    gClientId = 0
    gClientRefReqInd = False
    gClientReqFieldInd = False
    gReprintFileName = ""
    gProductionRun_Id = 0
    gLabelId = ""
    gProofId = 0
    gQuantity = 0
    gCodingFileName = ""
    gRandDelimiter = ""
    gOriginalPDRBarcode = ""
    gCodingNumber = 0
    gGroupNumber = 0
    gGroupName = ""
    gClientGroupName = ""
    gRandIDNumber = ""
    vdata = ""
    columnNumber = 0
    gSampleFileName = ""
    gSampleTypeId = 0
    writeData = ""
    gSampleFlag = False
    gLinksSpecInstr = ""
    gReprintFile_Type = ""
    gShowQuarantineTagFlag = False
    gReOrientFlag = False
    gRandBarcode = ""
    gOrigPDRNumber = ""
    gStockProofId = 0
    gStockLabelId = ""
    gBlindProofId = 0
    gBlindLabelId = ""
    gBlindLamApply = 0
    gOnsertDieToolId = 0
    gOnsertDiePartNumber = ""
    gCodingRepeatCnt = 0

    ' clean up all module varialbes here
    Set oCollIRQInfo = Nothing
    Set oIRQInfo = Nothing
    Set mColClientFields = Nothing
    sampleTypes = 0
    quantityTotal = 0
    SampleColumns = 0
    codingColumns = 0
    HoldStockProofId = 0
    HoldScratchStockProofId = 0
    mReprintFile_id = 0
    mReprintProdRunId = 0
    mvarOrigStockProofId = 0
    mvarOrigStockId = ""
    mvarOrigScratchStockProofId = 0
    mvarOrigScratchStockId = ""
    mvarOrigBlindLamApply = 0
    mvarOrigOnsertDieToolId = 0
    mvarOrigOnsertDiePartNumber = ""
    mvarOriginalOnsertBaseRollStockMinimumWidth = 0
    mvarOriginalOnsertBaseRollStockComponentId = 0
    mvarOriginalOnsertBaseRollStockPartNumber = ""
    mSampleFileName = ""
    mInputFileName = ""
    booSamplesQTYChanged = False
End Sub

Private Sub mnuAbout_Click()
    Dim abt As New AboutForm
    
    abt.Initialize gApplicationUser, Me.Icon, App.title, App.ProductName, App.Major, App.Minor, App.Revision
    abt.Show

End Sub

Private Sub cmdViewShipping_Click()
    If SSDBComboShip.text = "" Then
        MsgBox "Please select a Ship To.", vbExclamation
    Else
        frmShipTo.Show vbModal
    End If
End Sub

Private Sub duplicateConfig_Click()
    Dim i As Long
    Dim lngLink As Long
    Dim strTempFileName As String
    
    On Error GoTo Handle_Error

    If Not CheckExistingSample() Then
        MsgBox "Data does not Exist!" & vbCrLf & "Please create a Sample Configuration.", _
                vbInformation + vbOKOnly, "Warning - Error Duplicating"
        Exit Sub
    ElseIf Me.txtDirtyFlag = "Y" Then
        MsgBox "Data has not been saved!" & vbCrLf _
                        & "Please save the data first.", _
                vbInformation + vbOKOnly, "Warning - Error Duplicating"
        Exit Sub
    End If
    
    If gLinkLock.IsLockStillValid = False Then
        MsgBox "You cannot duplicate samples because your lock on the Links has been deleted." & vbCrLf & vbCrLf & "Please reload before continuing.", vbInformation + vbOKOnly, "Invalid Lock"
        Exit Sub
    End If

    Set dupData = New CCOLdupFiles
    Set dupSameCodingData = New CCOLdupFiles
    
    Screen.MousePointer = vbHourglass
    
    'loads the pdr selection form for duplication
    Load frmDuplicate
    
    Screen.MousePointer = vbDefault
    frmDuplicate.Show vbModal
    
    If dupData.count > 0 Or dupSameCodingData.count > 0 Then
        If MsgBox( _
            "Are you sure you want to Duplicate for Production Runs?" & vbCrLf & _
            "Yes: This will replace preconfigured samples.", _
            vbInformation + vbYesNo) = vbYes Then
        
            
            Set smpData = New CCOLsmpFiles
            Set smpSameCodingData = New CCOLsmpFiles
            
            Call LookUpDuplicateSet(ProductionRun.Production_Run_Id, smpData)
            Call LookUpDuplicateSet(ProductionRun.Production_Run_Id, smpSameCodingData)
            gSampleFlag = False
            
            '***********************************************************************
            '***********************************************************************
            '***********************************************************************
            
            Call DuplicateFilesCopyMods(smpSameCodingData, dupSameCodingData)
            ' at this point, smpData contains the newly created samples from the
            ' duping and none of the orig. dup info pdr samples
            If gSampleFlag = True Then
                GoTo Cleanup_Exit
            End If
            Call DuplicateUpdateSmpTable(smpSameCodingData)
            
            For i = 1 To dupSameCodingData.count
               
               Call DuplicateUpdateProdTable(dupSameCodingData.Item(i).productionId)
               
               If gDebugMode = True Then
                
                    If InStr(1, dupSameCodingData.Item(i).fileName, gRnrDebugFilePath, vbTextCompare) <> 1 Then
                        lngLink = GetDuplicateFileLinksInfo(dupSameCodingData.Item(i).productionId)
                        strTempFileName = Replace$(dupSameCodingData.Item(i).fileName, gRnrBaseFilePath, gRnrDebugFilePath, 1, -1, vbTextCompare)
                        'coding doesn't have debug in the path must copy files and update coding db path
                        Call CopyFiles(dupSameCodingData.Item(i).fileName, strTempFileName)
                        Call UpdateCodingFilePath(lngLink, dupSameCodingData.Item(i).productionId, strTempFileName, 0)
                    End If
                
               End If
               
            Next i
            
            '***********************************************************************
            '***********************************************************************
            '***********************************************************************
            
            Call DuplicateFiles(smpData, dupData)
            ' at this point, smpData contains the newly created samples from the
            ' duping and none of the orig. dup info pdr samples
            If gSampleFlag = True Then
                GoTo Cleanup_Exit
            End If
            Call DuplicateUpdateSmpTable(smpData)

            For i = 1 To dupData.count
                Call DuplicateUpdateProdTable(dupData.Item(i).productionId)
                    If gDebugMode = True Then
                        If InStr(1, dupData.Item(i).fileName, gRnrDebugFilePath, vbTextCompare) <> 1 Then
                            strTempFileName = Replace$(dupData.Item(i).fileName, gRnrBaseFilePath, gRnrDebugFilePath, 1, -1, vbTextCompare)
                            'coding doesn't have debug in the path must copy files and update coding db path
                            Call CopyFiles(dupData.Item(i).fileName, strTempFileName)
                            lngLink = GetDuplicateFileLinksInfo(dupData.Item(i).productionId)
                            Call UpdateCodingFilePath(lngLink, dupData.Item(i).productionId, strTempFileName, 0)
                        End If
                    End If
            Next i
            If booReplacement = False Then
                basGlobals.UpdateLinkLockVersionAndUpdatedBy
            Else
                basGlobals.UpdateLinkLockVersion
            End If
                
            MsgBox _
                "Samples have been duplicated for Production Runs.", vbInformation
        End If
                
    End If


Cleanup_Exit:
    Screen.MousePointer = vbDefault
    Exit Sub
Handle_Error:
     MsgBox Err.description & vbCrLf & _
         "in frmProdPlan.mnuDuplicate_Click ", _
         vbCritical + vbOKOnly, "Application Error"
     Resume Cleanup_Exit
End Sub

Private Sub Load_Menu()
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
    
    Dim i As Long

    Me.mnuReplacements.enabled = True
    Me.mnuViewAuditTrail.enabled = True
    

    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        .RowsetSize = 1

        ' Call the SP to create the resultset
        .AddParameter "orig prod run id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ProductionRunByOrigId"
        i = 0
        If Not .Recordset.EOF Then
            Me.mnuLine1.Visible = True
            Do Until .Recordset.EOF
                If i < 10 Then
                    Me.mnuRepProdRuns(i).Visible = True
                    Me.mnuRepProdRuns(i).Caption = .Recordset!Production_Run_Barcode
                    i = i + 1
                    .Recordset.MoveNext
                Else
' DW 2012-001 Issue # 136 move msgbox after recordset.close
'                    MsgBox _
'                        "More than 10 Replacement Production Runs exist. " & vbCrLf & _
'                        "Contact IT if you need to access a Production Run that " & _
'                        "is not visible in the menu.", vbExclamation
                    Exit Do
                End If
            Loop
        End If
        .Recordset.Close
        ' DW 2012-001 Issue #136 No-No
        If i >= 10 Then
            MsgBox _
                "More than 10 Replacement Production Runs exist. " & vbCrLf & _
                "Contact IT if you need to access a Production Run that " & _
                "is not visible in the menu.", vbExclamation
        End If
    End With


Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR - Loading Menu"
    Resume Exit_this_Sub

End Sub

' this creates a new Replacement Run
Private Sub mnuReplacements_Click()
    Dim n As Long
    Dim strMessageCC As String
    Dim strMessage As String
    
    On Error GoTo Error_this_Sub
      
    CheckSave
    
    Dim Message, title, Default
    Dim txtInputSeqNum As String
    Dim booFound As Boolean
    Dim strReprintFileName As String
    Dim lngReprintQty As Long
    Dim intLastSlashPos As Integer
    Dim objData As nADOData.CADOData
    
    Message = "Scan the Barcode from the Reprint Report!"
    title = "Reprint File"
    Default = "0"
    txtInputSeqNum = InputBox(Message, title, Default, 4320, 4320)
    If Len(txtInputSeqNum) = 0 Then
        Exit Sub
    End If
    txtInputSeqNum = UCase$(txtInputSeqNum)
    
    'try to find the file
    booFound = False
    
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
           
       
        .AddParameter "Reprint File Barcode", txtInputSeqNum, adVarChar, adParamInput
        .OpenRecordSetFromSP "get_ReprintFileByBarcode"
        If .Recordset.EOF Then
            booFound = False
        Else
            booFound = True
            strReprintFileName = .Recordset!Reprint_File_Name
            lngReprintQty = .Recordset!File_Record_Cnt
            'md added for clintrak samples
            mReprintFile_id = .Recordset!Reprint_File_Id
            gReprintFile_Type = .Recordset!Reprint_File_Type
            mInputFileName = .Recordset!Input_File_Name
            mReprintProdRunId = .Recordset!Production_Run_Id
        End If
        .Recordset.Close
    
    End With
    
    If booFound = False Then
        MsgBox "The scanned Reprint File, " & txtInputSeqNum & ", was not found.", _
                vbExclamation + vbOKOnly, "Error - File Not Found!"
        Exit Sub
    End If
    
    'md new check - should only be scanning a file intended for a Replacement
    If gReprintFile_Type <> "REPLACEMENT" Then
        MsgBox "The scanned file, " & txtInputSeqNum & ", was not intended to be a Replacement. " & _
             "Please try again!!", vbExclamation + vbOKOnly, "Error - Incorrect File Scanned!"
        Exit Sub
    End If
         
    If mReprintProdRunId <> ProductionRun.Production_Run_Id Then
        MsgBox "The scanned file, " & txtInputSeqNum & ", does not belong to " & txtBarcodeId & ". " & _
             "Please try again!!", vbExclamation + vbOKOnly, "Error - Incorrect File Scanned!"
        Exit Sub
    End If
    
     'verify that the data file scanned exists on the server before continuing
    If Not FileExists(strReprintFileName) Then
        MsgBox "The Reprint File Does Not Exist on the server!" & vbCrLf _
            & "Verify that the file was not archived and try again.", vbExclamation + vbOKOnly, "Input File Does Not Exist!"
        Exit Sub
    End If
        
    Set gBarcodeFile = New CodingFile
    gBarcodeFile.Init gApplicationUser
    gBarcodeFile.LoadByRPF "", mReprintFile_id
    
    'reset screen fields and collection fields
    booReplacement = True
    Me.txtReplacement.Visible = True
    Me.chkReOrientation.enabled = True
    
    'md added for Samples project
     If gReprintFile_Type = "REPLACEMENT" Then
        Me.txtReplacement.text = "REPLACEMENT"
     Else
        Me.txtReplacement.text = "RESUPPLY"
        Me.txtProdDesc = ProductionRun.LabelDescription
     End If
     
  
    'md added for clintrak samples - must do this in the correct order to preserve the
    'integrity of the replacement labels.
    Me.SSDBComboShip.enabled = True
    Me.cmdSamples.enabled = True
    Me.cmdSave.enabled = True
    Me.mnuReplacements.enabled = False
    Me.mnuViewAuditTrail.enabled = False
    
    ' tj IRQ stuff 2 - //HoldStockProofId\\ and //HoldScratchStockProofId\\ would have been set from initial load and should be the same
    Me.txtStockIRQ = ""
    Me.txtScratchIRQ = ""
        
    gReprintFileName = strReprintFileName
    intLastSlashPos = InStrRev(strReprintFileName, "\")
    Me.txtFileName = Mid$(strReprintFileName, intLastSlashPos + 1, Len(strReprintFileName) - 1)
    Me.txtQty = lngReprintQty
    Me.txtSamples = 0
    Me.txtSampleGroups = 0
    Me.txtBarcodeId = "PDR______"
    
    ProductionRun.File_Name = strReprintFileName
    ProductionRun.Orig_Prod_Run_Id = ProductionRun.Production_Run_Id
    ProductionRun.Production_Run_Id = 0
    ProductionRun.Qty_Requested = lngReprintQty
    ProductionRun.Samples_Requested = 0
    ProductionRun.Reprint_File_Id = mReprintFile_id
    ProductionRun.NonBillable.SetNew ProductionRun.Job_Log_Id
    If ProductionRun.NonBillable.isJobLevel = False Then
        ' If we haven't found a job-level non-billable, force the combo box to
        ' empty. This forces the user to pick a value on save.
        ProductionRun.NonBillable.reasonId = -1
    End If
    
    Me.txtProducedBy = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
    Me.chkReOrientation.value = 0
    
    Call Lock_Out_PDRForm(True, False)
    
    'client reqd fields (if they exist) are carried over from the original PDR.
    'This will clear out the Production_Run_Client_Fields_Id
    'so they will be saved for the new Replacement PDR
    If ClientReqdFields.count > 0 Then
        If ClientReqdFields.Item(1).Field_Name_Value <> "" Then
            For n = 1 To ClientReqdFields.count
                ClientReqdFields.Item(n).Production_Run_Client_Fields_Id = 0
            Next n
        Else
            For n = ClientReqdFields.count To 1 Step -1
                ClientReqdFields.Remove (n)
            Next n
            Me.cmdAddtlData.enabled = False
        End If
    End If
    
    'the reference number (if one existed) is carried over
    'from the original PDR so it is not editable ever
    Me.txtReferanceNo.enabled = False
    
    ' the special instructions should not be enabled until the replacement run is initially
    ' saved b/c that screen writes directly to the DB and if the PDR isn't in the DB, the
    ' changes are just not saved.
    Me.cmdSpecInst.enabled = False
    
    
    If (ProductionRun.Stock_Proof_Id <> gStockProofId) Then
        If MsgBox("The original PDR's stock is " & ProductionRun.stock & "." & vbCrLf & _
                "Label Specs's stock is currently " & gStockLabelId & "." & vbCrLf & vbCrLf & _
                "Do you want to update this Replacement PDR Stock to match Label Specs?", _
                vbYesNo + vbQuestion, "Update Stock") = vbYes Then
            Me.txtStockNo.text = gStockLabelId
            ProductionRun.stock = gStockLabelId
            ProductionRun.Stock_Proof_Id = gStockProofId

        End If
                
    End If
    
    If (ProductionRun.Scratch_Proof_Id <> gBlindProofId) Or (BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff)) <> CCBlindLamApplyCol(CStr(gBlindLamApply))) Then
        
        strMessageCC = gBlindLabelId
        If gBlindLabelId <> "N/A" Then
            strMessageCC = strMessageCC & " (" & CCBlindLamApplyCol(CStr(gBlindLamApply)) & ")"
        End If
        
        strMessage = ProductionRun.Scratch_Stock
        If ProductionRun.Scratch_Stock <> "N/A" Then
            strMessage = strMessage & " (" & BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff)) & ")"
        End If
        
        If MsgBox("The original PDR's Blinding Laminate is " & strMessage & "." & vbCrLf & _
                "Label Specs's Blinding Laminate is currently " & strMessageCC & "." & vbCrLf & vbCrLf & _
                "Do you want to update this Replacement PDR Blinding Laminate to match Label Specs?", _
                vbYesNo + vbQuestion, "Update Blinding Laminate") = vbYes Then
             
            Me.lblApplyBlindLam.Caption = CCBlindLamApplyCol(CStr(gBlindLamApply))
            If Me.lblApplyBlindLam.Caption = NA_TEXT Or Me.lblApplyBlindLam.Caption = NOT_BLINDED_TEXT Then
                Me.lblApplyBlindLam.Caption = ""
            End If

            For n = 1 To BlindLamApplyCol.count
                If BlindLamApplyCol(n) = CCBlindLamApplyCol(CStr(gBlindLamApply)) Then
                    ProductionRun.Apply_ScratchOff = n - 1
                    Exit For
                End If
            Next n
            
            Me.txtScratchStockNo.text = gBlindLabelId
            ProductionRun.Scratch_Stock = gBlindLabelId
            ProductionRun.Scratch_Proof_Id = gBlindProofId

        End If
    End If
    
    ' Check if Onsert die has changed in Label Specs after the original PDR was created.
    If (ProductionRun.OnsertDieToolId <> gOnsertDieToolId) Then
        If MsgBox( _
            "The original PDR's Onsert die is " & ProductionRun.OnsertDiePartNumber & "." & vbCrLf & _
            "Label Specs's Onsert die is currently " & gOnsertDiePartNumber & "." & vbCrLf & vbCrLf & _
            "Do you want to update the Replacement PDR die to match Label Specs?", _
            vbYesNo + vbQuestion, "Update Die") = vbYes Then
            Me.txtOnsertDie.text = gOnsertDiePartNumber
            ProductionRun.OnsertDiePartNumber = gOnsertDiePartNumber
            ProductionRun.OnsertDieToolId = gOnsertDieToolId
        End If
    End If
    
    If (ProductionRun.OnsertBaseRollStockComponentId <> gOnsertBaseRollStockComponentId) Then
        If basGlobals.ShowChangePromptForNewReplacement("Onsert base roll stock", FormatEmptyStringNA(ProductionRun.OnsertBaseRollStockPartNumber), FormatEmptyStringNA(gOnsertBaseRollStockPartNumber), "") Then
            ProductionRun.OnsertBaseRollStockPartNumber = gOnsertBaseRollStockPartNumber
            ProductionRun.OnsertBaseRollStockComponentId = gOnsertBaseRollStockComponentId
            Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
        End If
    End If
    
    If (ProductionRun.OnsertBaseRollStockMinimumWidth <> gOnsertBaseRollStockMinimumWidth) Then
        If basGlobals.ShowChangePromptForNewReplacement("Onsert base roll stock minimum width", FormatZeroNumberNA(ProductionRun.OnsertBaseRollStockMinimumWidth), FormatZeroNumberNA(gOnsertBaseRollStockMinimumWidth), "") Then
            ProductionRun.OnsertBaseRollStockMinimumWidth = gOnsertBaseRollStockMinimumWidth
            Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
        End If
    End If
    
    Me.chkPrintAtPackager.enabled = True
    ManageClientInventoryCheckBox
    
Proc_EXIT:
    Exit Sub
  
Error_this_Sub:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Error Processing Reprint Name"
    Resume Proc_EXIT
  
End Sub

' this loads an existing Replacement Run
Private Sub mnuRepProdRuns_Click(index As Integer)
    Dim objData As nADOData.CADOData
    Dim strMessageOrig As String
    Dim strMessage As String
    Dim booIRQsExist As Boolean
    Dim i As Long
    Dim strMessageCC As String
    Dim booChange As Boolean
    Dim strReason As String
    Dim strTempFileName As String
    Dim strFileNameSamples As String
    
    On Error GoTo Error_this_Sub
    
    Call CheckSave

    Dim intLastSlashPos As Integer
    Dim RPFound As Boolean
    Dim strCodingFileHeader As String
    
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        .RowsetSize = 1
        .AddParameter "prod run barcode", mnuRepProdRuns(index).Caption, adVarChar, adParamInput
        .OpenRecordSetFromSP "get_ProductionRunByBarcode"
        If .Recordset.EOF Then
            MsgBox "The selected Production Run, " & mnuRepProdRuns(index).Caption & ", was not found.", vbExclamation
            Exit Sub
         End If
                         
        'load the replacement production run data into the collection
        ProductionRun.Production_Run_Id = .Recordset!Production_Run_Id
        gProductionRun_Id = ProductionRun.Production_Run_Id
        ProductionRun.Reprint_File_Id = .Recordset!Reprint_File_Id

        If mvarOrigStockProofId = 0 Then
            mvarOrigStockProofId = ProductionRun.Stock_Proof_Id
            mvarOrigStockId = ProductionRun.stock
            mvarOrigScratchStockProofId = ProductionRun.Scratch_Proof_Id
            mvarOrigScratchStockId = ProductionRun.Scratch_Stock
            mvarOrigBlindLamApply = ProductionRun.Apply_ScratchOff
            mvarOrigOnsertDieToolId = ProductionRun.OnsertDieToolId
            mvarOriginalOnsertBaseRollStockMinimumWidth = ProductionRun.OnsertBaseRollStockMinimumWidth
            mvarOriginalOnsertBaseRollStockComponentId = ProductionRun.OnsertBaseRollStockComponentId
            mvarOriginalOnsertBaseRollStockPartNumber = ProductionRun.OnsertBaseRollStockPartNumber
            mvarOrigOnsertDiePartNumber = ProductionRun.OnsertDiePartNumber
        End If
        
        strCodingFileHeader = ProductionRun.Coding_File_Header
         
        'reset screen fields and collection fields
        booReplacement = True
        Me.txtReplacement.Visible = True
        
        ' reset the text of the PDR status field back to the default value from when the screen was originally loaded.
        txtPDRStatus.text = "PDR HAS BEEN PROCESSED"
    
        ProductionRun.LookupRecord
        
        'handle replacements created in prod when in debug mode that never loaded in debug before
         If gDebugMode = True Then
            'if coding path doesn't have debug, lets copy coding files
            
            If InStr(1, ProductionRun.File_Name, gRnrDebugFilePath, vbTextCompare) <> 1 Then

                strTempFileName = Replace$(ProductionRun.File_Name, gRnrBaseFilePath, gRnrDebugFilePath, 1, -1, vbTextCompare)
                'coding doesn't have debug in the path must copy files and update coding db path
                Call CopyFiles(ProductionRun.File_Name, strTempFileName)
                Call UpdateCodingFilePath(gCodingFileId, ProductionRun.Production_Run_Id, strTempFileName, 1)
                
                'create sample path from coding path and check if it's debug
                strFileNameSamples = Replace$(ProductionRun.File_Name, "\reprints\", "\smp\", 1, -1, vbTextCompare)  'it's already doing this in the code
                strFileNameSamples = GetFilePath(strFileNameSamples) 'drop the file in the path
                strTempFileName = Replace$(strFileNameSamples, gRnrBaseFilePath, gRnrDebugFilePath, 1, -1, vbTextCompare)
                ProductionRun.LookupRecord
                
                Call CopySampleFiles(strFileNameSamples, strTempFileName, ProductionRun.Barcode_Id)
            End If
        End If
        
        ProductionRun.Coding_File_Header = strCodingFileHeader
        
        'Load Client Required fields and values
        Call LoadClientLabelFields(gClientId, ProductionRun.Production_Run_Id)
        
        'Initialize Holder
        Set mColClientFields = ClientReqdFields.Clone
        
        If gClientReqFieldInd = True Or ClientReqdFields.count > 0 Then
            If ClientReqdFields.Item(1).Production_Run_Client_Fields_Id = 0 Then
                Me.cmdAddtlData.enabled = False
            Else
                Me.cmdAddtlData.enabled = True
            End If
            
        Else
            Me.cmdAddtlData.enabled = False
        End If
        
        'the reference number (if one existed) was carried over from the original
        'PDR when this Replacement PDR was created so it is not editable ever
        Me.txtReferanceNo.enabled = False
                
        Me.chkPrintAtPackager.enabled = True
        
        'md added for clintrak samples - determine if the PDR has been RUN.  If so, block all
        'fields on the form.
        If Determine_If_PDR_HasRun Then
            Call Lock_Out_PDRForm(False, True)
            Me.cmdSamples.enabled = True
            ' check to see if the PDR is on a PKS
            ' b/c the Job Shipping Flag isn't the most reliable indicator
            If Determine_Shipping_Flag_On Or DeterminePDROnPKS Then
                SSDBComboShip.enabled = False
                Me.chkPrintAtPackager.enabled = False
                ' check for samples on PKS
                If Me.txtSampleGroups.text > 1 Then
                    If DeterminePDROnPKS("S") Then
                        cmdSamples.enabled = False
                    End If
                Else
                    Me.cmdSamples.enabled = False
                End If
                Me.chkReOrientation.enabled = False
                cmdSave.enabled = False
            End If
            If ProductionRun.Barcode_Id <> "" Then
                If Planning.CheckIfCombined(ProductionRun.Barcode_Id) Then
                    txtPDRStatus.text = "COMBINED PDR HAS BEEN PROCESSED"
                    chkPrintAtPackager.enabled = False
                End If
            End If
        Else
            Call Lock_Out_PDRForm(True, False)
            SSDBComboShip.enabled = True
            cmdSamples.enabled = True
            cmdSave.enabled = True
            Me.chkReOrientation.enabled = True
            If ProductionRun.Barcode_Id <> "" Then
                If Planning.CheckIfCombined(ProductionRun.Barcode_Id) Then
                    Call Lock_Out_PDRForm(False, True)
                    txtProdDesc.enabled = True
                    txtPDRStatus.text = "COMBINED PDR"
                    chkPrintAtPackager.enabled = False
                End If
            End If
        End If
        
        Me.mnuReplacements.enabled = False
        Me.mnuViewAuditTrail.enabled = True
       
        gReprintFileName = ProductionRun.File_Name
        intLastSlashPos = InStrRev(gReprintFileName, "\")
        Me.txtFileName = Mid$(gReprintFileName, intLastSlashPos + 1, Len(gReprintFileName) - 1)
        Me.txtQty = ProductionRun.Qty_Requested
        Me.txtSamples = ProductionRun.Samples_Requested
        Me.txtSampleGroups = ProductionRun.Sample_Number
        Me.txtProducedBy = ProductionRun.Produced_By
        Call SetSSDBComboText(Me.SSDBComboShip, ProductionRun.Ship_To_Id)
        ' hold the orig pdr number
        If gOrigPDRNumber = "" Then
            gOrigPDRNumber = Me.txtBarcodeId.text
        End If
        Me.txtBarcodeId = ProductionRun.Barcode_Id
        Me.txtProdDesc = ProductionRun.Prod_Description
        Me.txtReferanceNo = ProductionRun.Reference_No
        ' tj - IRQ stuff 2
        HoldStockProofId = ProductionRun.Stock_Proof_Id
        HoldScratchStockProofId = ProductionRun.Scratch_Proof_Id
        Me.txtStockNo.text = ProductionRun.stock
        Me.txtOnsertDie.text = ProductionRun.OnsertDiePartNumber
        Me.txtScratchStockNo.text = ProductionRun.Scratch_Stock
        Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
        Me.lblApplyBlindLam.Caption = BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff))
        If Me.lblApplyBlindLam.Caption = NA_TEXT Or Me.lblApplyBlindLam.Caption = NOT_BLINDED_TEXT Then
            Me.lblApplyBlindLam.Caption = ""
        End If
        
        Call UpdateExistingIRQList
        
        Me.chkPrintAtPackager.value = ProductionRun.PrintAtPackager
    End With
    
 'md added for Samples project
    Call GetReprintFile_For_PDR(RPFound)
    If RPFound Then
        If gReprintFile_Type = "REPLACEMENT" Then
            Me.txtReplacement.text = "REPLACEMENT"
        Else
            Me.txtReplacement.text = "RESUPPLY"
        End If
    Else
        'use as default for any older PDR's created before this impelementation
        Me.txtReplacement.text = "REPLACEMENT"
    End If
    
    Me.chkReOrientation.value = ProductionRun.Reorient_Ind
    
    booIRQsExist = False
    If Not oCollIRQInfo Is Nothing Then
        If oCollIRQInfo.count > 0 Then
            booIRQsExist = True
        End If
    End If
    
    ' Clear dirty flag here because if the flag was tripped as
    ' part of the data load process it doesn't count.
    Me.txtDirtyFlag.text = ""
    mvarLinksReportDirty = False
    
    If DeterminePDROnPKS = False Then
        strReason = ""
        If InStr(txtPDRStatus.text, "PROCESSED") > 0 And Me.txtPDRStatus.Visible = True Then
            strReason = "has been run."
        ElseIf booIRQsExist Then
            strReason = "has an IRQ."
        ElseIf InStr(txtPDRStatus.text, "COMBINED") > 0 And Me.txtPDRStatus.Visible = True Then
            strReason = "is part of a PRG."
        End If
    
        If (ProductionRun.Stock_Proof_Id <> mvarOrigStockProofId) Then
            If strReason = "" Then
                If MsgBox("This Replacement PDR's stock is currently " & ProductionRun.stock & "." & vbCrLf & _
                        "The Original PDR's stock is " & mvarOrigStockId & "." & vbCrLf & vbCrLf & _
                        "Do you want to update this Replacement PDR Stock to match the Original PDR?", _
                        vbYesNo + vbQuestion, "Update Stock") = vbYes Then
                    Me.txtStockNo.text = mvarOrigStockId
                    ProductionRun.stock = mvarOrigStockId
                    ProductionRun.Stock_Proof_Id = mvarOrigStockProofId
                    
                    MsgBox "This Replacement PDR's stock has been updated to match the Original PDR." & vbCrLf & _
                                "You must SAVE this Replacement PDR in order to save this change.", vbInformation, "Stock Updated"
                    
                    booChange = True
                End If
            Else
                MsgBox "This Replacement PDR's stock is currently " & ProductionRun.stock & "." & vbCrLf & _
                        "The Original PDR's stock is " & mvarOrigStockId & "." & vbCrLf & vbCrLf & _
                        "The PDR's stock CANNOT be updated at this time because it " & strReason, _
                        vbOKOnly + vbInformation, "Stock Difference"
            End If
                
        ElseIf (ProductionRun.Stock_Proof_Id <> gStockProofId) Then
            If strReason = "" Then
                If MsgBox("This Replacement PDR's stock is currently " & ProductionRun.stock & "." & vbCrLf & _
                        "Label Specs's stock is currently " & gStockLabelId & "." & vbCrLf & vbCrLf & _
                        "Do you want to update this Replacement PDR Stock to match Label Specs?", _
                        vbYesNo + vbQuestion, "Update Stock") = vbYes Then
                    Me.txtStockNo.text = gStockLabelId
                    ProductionRun.stock = gStockLabelId
                    ProductionRun.Stock_Proof_Id = gStockProofId
                    
                    MsgBox "This Replacement PDR's stock has been updated to match Label Specs." & vbCrLf & _
                                "You must SAVE this Replacement PDR in order to save this change.", vbInformation, "Stock Updated"
                    
                    booChange = True
                End If
            Else
                MsgBox "This Replacement PDR's stock is currently " & ProductionRun.stock & "." & vbCrLf & _
                        "Label Specs's stock is currently " & gStockLabelId & "." & vbCrLf & vbCrLf & _
                        "The PDR's stock CANNOT be updated at this time because it " & strReason, _
                        vbOKOnly + vbInformation, "Stock Difference"
            End If
    
                    
        End If
        If (ProductionRun.Scratch_Proof_Id <> mvarOrigScratchStockProofId) Or (mvarOrigScratchStockProofId > 0 And (BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff)) <> BlindLamApplyCol(CStr(mvarOrigBlindLamApply)))) Then
            
            strMessageOrig = mvarOrigScratchStockId
            If gBlindLabelId <> "N/A" Then
                strMessageOrig = strMessageOrig & " (" & BlindLamApplyCol(CStr(mvarOrigBlindLamApply)) & ")"
            End If
            
            strMessage = ProductionRun.Scratch_Stock
            If ProductionRun.Scratch_Stock <> "N/A" Then
                strMessage = strMessage & " (" & BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff)) & ")"
            End If
            
            
            If strReason = "" Then
                If MsgBox("This Replacement PDR's Blinding Laminate is currently " & strMessage & "." & vbCrLf & _
                        "The Original PDR's Blinding Laminate is currently " & strMessageOrig & "." & vbCrLf & vbCrLf & _
                        "Do you want to update this Replacement PDR Blinding Laminate to match the Original PDR?", _
                        vbYesNo + vbQuestion, "Update Blinding Laminate") = vbYes Then
                     
                    Me.lblApplyBlindLam.Caption = BlindLamApplyCol(CStr(mvarOrigBlindLamApply))
                    If Me.lblApplyBlindLam.Caption = NA_TEXT Or Me.lblApplyBlindLam.Caption = NOT_BLINDED_TEXT Then
                        Me.lblApplyBlindLam.Caption = ""
                    End If
        
                    For i = 1 To BlindLamApplyCol.count
                        If BlindLamApplyCol(i) = BlindLamApplyCol(CStr(mvarOrigBlindLamApply)) Then
                            ProductionRun.Apply_ScratchOff = i - 1
                            Exit For
                        End If
                    Next i
                    
                    Me.txtScratchStockNo.text = mvarOrigScratchStockId
                    ProductionRun.Scratch_Stock = mvarOrigScratchStockId
                    ProductionRun.Scratch_Proof_Id = mvarOrigScratchStockProofId
                    
                    MsgBox "This Replacement PDR's Blinding Laminate has been updated to match the Original PDR." & vbCrLf & _
                                "You must SAVE this Replacement PDR in order to save this change.", vbInformation, "Blinding Laminate Updated"
                    
                    booChange = True
                End If
            Else
                MsgBox "This Replacement PDR's Blinding Laminate is currently " & strMessage & "." & vbCrLf & _
                        "The Original PDR's Blinding Laminate is currently " & strMessageOrig & "." & vbCrLf & vbCrLf & _
                        "The PDR's Blinding Laminate CANNOT be updated at this time because it " & strReason, _
                        vbOKOnly + vbInformation, "Blinding Laminate Difference"
            End If
            
                
        ElseIf (ProductionRun.Scratch_Proof_Id <> gBlindProofId) Or (gBlindProofId > 0 And (BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff)) <> CCBlindLamApplyCol(CStr(gBlindLamApply)))) Then
            
            strMessageCC = gBlindLabelId
            If gBlindLabelId <> "N/A" Then
                strMessageCC = strMessageCC & " (" & CCBlindLamApplyCol(CStr(gBlindLamApply)) & ")"
            End If
            
            strMessage = ProductionRun.Scratch_Stock
            If ProductionRun.Scratch_Stock <> "N/A" Then
                strMessage = strMessage & " (" & BlindLamApplyCol(CStr(ProductionRun.Apply_ScratchOff)) & ")"
            End If
            
            If strReason = "" Then
                If MsgBox("This Replacement PDR's Blinding Laminate is currently " & strMessage & "." & vbCrLf & _
                        "Label Specs's Blinding Laminate is currently " & strMessageCC & "." & vbCrLf & vbCrLf & _
                        "Do you want to update this Replacement PDR Blinding Laminate to match Label Specs?", _
                        vbYesNo + vbQuestion, "Update Blinding Laminate") = vbYes Then
                     
                    Me.lblApplyBlindLam.Caption = CCBlindLamApplyCol(CStr(gBlindLamApply))
                    If Me.lblApplyBlindLam.Caption = NA_TEXT Or Me.lblApplyBlindLam.Caption = NOT_BLINDED_TEXT Then
                        Me.lblApplyBlindLam.Caption = ""
                    End If
        
                    For i = 1 To BlindLamApplyCol.count
                        If BlindLamApplyCol(i) = CCBlindLamApplyCol(CStr(gBlindLamApply)) Then
                            ProductionRun.Apply_ScratchOff = i - 1
                            Exit For
                        End If
                    Next i
                    
                    Me.txtScratchStockNo.text = gBlindLabelId
                    ProductionRun.Scratch_Stock = gBlindLabelId
                    ProductionRun.Scratch_Proof_Id = gBlindProofId
                    
                    MsgBox "This Replacement PDR's Blinding Laminate has been updated to match Label Specs." & vbCrLf & _
                                "You must SAVE this Replacement PDR in order to save this change.", vbInformation, "Blinding Laminate Updated"
                    
                    booChange = True
                End If
            Else
                MsgBox "This Replacement PDR's Blinding Laminate is currently " & strMessage & "." & vbCrLf & _
                        "Label Specs's Blinding Laminate is currently " & strMessageCC & "." & vbCrLf & vbCrLf & _
                        "The PDR's Blinding Laminate CANNOT be updated at this time because it " & strReason, _
                        vbOKOnly + vbInformation, "Blinding Laminate Difference"
            End If
                
        End If

        ' Check to see if the Replacement PDR's Onsert die is different from the original PDR
        If (ProductionRun.OnsertDieToolId <> mvarOrigOnsertDieToolId) Then
            If strReason = "" Then
                If MsgBox( _
                    "This Replacement PDR's Onsert die is currently " & ProductionRun.OnsertDiePartNumber & "." & vbCrLf & _
                    "The Original PDR's Onsert die is " & mvarOrigOnsertDiePartNumber & "." & vbCrLf & vbCrLf & _
                    "Do you want to update the Replacement PDR die to match the Original PDR?", _
                    vbYesNo + vbQuestion, "Update Die") = vbYes Then
                    Me.txtOnsertDie.text = mvarOrigOnsertDiePartNumber
                    ProductionRun.OnsertDiePartNumber = mvarOrigOnsertDiePartNumber
                    ProductionRun.OnsertDieToolId = mvarOrigOnsertDieToolId
                    MsgBox _
                        "The Replacement PDR's Onsert die has been updated to match the Original PDR." & vbCrLf & _
                        "You must SAVE this Replacement PDR in order to save this change.", vbInformation, "Die Updated"
                    booChange = True
                End If
            Else
                MsgBox _
                    "This Replacement PDR's Onsert die is currently " & ProductionRun.OnsertDiePartNumber & "." & vbCrLf & _
                    "The Original PDR's Onsert die is " & mvarOrigOnsertDiePartNumber & "." & vbCrLf & vbCrLf & _
                    "The PDR's die CANNOT be updated at this time because it " & strReason, _
                    vbOKOnly + vbInformation, "Die Difference"
            End If
                
        ' Check to see if the Replacement PDR's Onsert die is different from the Label
        ElseIf (ProductionRun.OnsertDieToolId <> gOnsertDieToolId) Then
            If strReason = "" Then
                If MsgBox( _
                    "This Replacement PDR's Onsert die is currently " & ProductionRun.OnsertDiePartNumber & "." & vbCrLf & _
                    "Label Specs's Onsert die is currently " & gOnsertDiePartNumber & "." & vbCrLf & vbCrLf & _
                    "Do you want to update the Replacement PDR die to match Label Specs?", _
                    vbYesNo + vbQuestion, "Update Die") = vbYes Then
                    Me.txtOnsertDie.text = gOnsertDiePartNumber
                    ProductionRun.OnsertDiePartNumber = gOnsertDiePartNumber
                    ProductionRun.OnsertDieToolId = gOnsertDieToolId
                    MsgBox _
                        "The Replacement PDR's die has been updated to match Label Specs." & vbCrLf & _
                        "You must SAVE this Replacement PDR in order to save this change.", vbInformation, "Die Updated"
                    booChange = True
                End If
            Else
                MsgBox "This Replacement PDR's Onsert die is currently " & ProductionRun.OnsertDiePartNumber & "." & vbCrLf & _
                        "Label Specs's Onsert die is currently " & gOnsertDiePartNumber & "." & vbCrLf & vbCrLf & _
                        "The PDR's die CANNOT be updated at this time because it " & strReason, _
                        vbOKOnly + vbInformation, "Die Difference"
            End If
        End If
        
        If (ProductionRun.OnsertBaseRollStockPartNumber <> mvarOriginalOnsertBaseRollStockPartNumber) Then
            If basGlobals.ShowChangePromptForExistingReplacement("Onsert base roll stock", FormatEmptyStringNA(ProductionRun.OnsertBaseRollStockPartNumber), FormatEmptyStringNA(mvarOriginalOnsertBaseRollStockPartNumber), strReason) Then
                ProductionRun.OnsertBaseRollStockPartNumber = mvarOriginalOnsertBaseRollStockPartNumber
                ProductionRun.OnsertBaseRollStockComponentId = mvarOriginalOnsertBaseRollStockComponentId
                Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
            End If
            
        ElseIf (ProductionRun.OnsertBaseRollStockPartNumber <> basGlobals.gOnsertBaseRollStockPartNumber) Then
            If basGlobals.ShowChangePromptForLabel("Onsert base roll stock", FormatEmptyStringNA(ProductionRun.OnsertBaseRollStockPartNumber), FormatEmptyStringNA(basGlobals.gOnsertBaseRollStockPartNumber), strReason) Then
                ProductionRun.OnsertBaseRollStockPartNumber = basGlobals.gOnsertBaseRollStockPartNumber
                ProductionRun.OnsertBaseRollStockComponentId = basGlobals.gOnsertBaseRollStockComponentId
                Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
            End If
      
        End If
        
        If (ProductionRun.OnsertBaseRollStockMinimumWidth <> mvarOriginalOnsertBaseRollStockMinimumWidth) Then
            If basGlobals.ShowChangePromptForExistingReplacement("Onsert base roll minimum width", FormatZeroNumberNA(ProductionRun.OnsertBaseRollStockMinimumWidth), FormatZeroNumberNA(mvarOriginalOnsertBaseRollStockMinimumWidth), strReason) Then
                ProductionRun.OnsertBaseRollStockMinimumWidth = mvarOriginalOnsertBaseRollStockMinimumWidth
                Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
            End If
        
        ElseIf (ProductionRun.OnsertBaseRollStockMinimumWidth <> basGlobals.gOnsertBaseRollStockMinimumWidth) Then
            If basGlobals.ShowChangePromptForLabel("Onsert base roll minimum width", FormatZeroNumberNA(ProductionRun.OnsertBaseRollStockMinimumWidth), FormatZeroNumberNA(basGlobals.gOnsertBaseRollStockMinimumWidth), strReason) Then
                ProductionRun.OnsertBaseRollStockMinimumWidth = basGlobals.gOnsertBaseRollStockMinimumWidth
                Me.txtOnsertBaseRollStock = ProductionRun.OnsertStock
            End If
       
        End If
       
        
        If booChange = True Then
            txtDirtyFlag = "Y"
            Me.txtProducedBy = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
        End If
           
        ManageClientInventoryCheckBox
' If not print at packager check for samples . Allow correction if file length is 0
        If Not CBool(Me.chkPrintAtPackager.value) Then
            If Me.txtSamples > 0 Then
' need to set gSampleFilename to reprint PDR. Currently set to original PDR
                Call GetSampleFileName
                If FileLen(gSampleFileName) = 0 Then
                    If vbYes = MsgBox("The Internal Sample Files were not found. Do you want the system to re-create the Internal Samples based on today’s data?", vbYesNo) Then
                        If Fix_Samples_FromPDR_Screen Then
'change PDR sample counts for new samples just configured
                            ProductionRun.Samples_Requested = CLng(Me.txtSamples)
                            ProductionRun.Sample_Number = CInt(Me.txtSampleGroups)
                            MsgBox "Internal Samples Fixed", vbOKOnly
                        Else
                            MsgBox _
                                "Unable to Fix Internal Samples" & vbCrLf & _
                                "Please contact IT.", vbExclamation
                        End If
                    End If
                End If
            End If
        End If
    
    End If
    
Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR - Loading Menu"
    Resume Exit_this_Sub
    
End Sub

Private Sub Enable_Fields()
    Me.txtProdDesc.enabled = True
    Me.txtReferanceNo.enabled = True
End Sub

Private Sub SSDBComboShip_Click()
    Me.cmdViewShipping.enabled = (Me.SSDBComboShip.text <> "")
End Sub

Private Sub SSDBComboShip_InitColumnProps()
    SSDBComboShip.Columns(0).Width = SSDBComboShip.Width
    SSDBComboShip.Columns(1).Visible = False
    SSDBComboShip.Columns(2).Width = SSDBComboShip.Width
End Sub

Private Sub cmdSave_Click()
    Dim booInitialSave As Boolean
    Dim strCodingFileHeader As String
    Dim strTempFileName As String
        
    On Error GoTo Error_this_Sub

    ' DW 2010-002 added
    ' Check for Replacement PDRs (Add/Remove Clintrak Samples based on Print at Packager
    ' ----------------------------------------------------------------------------------
    If booReplacement Then
        Screen.MousePointer = vbHourglass
        If ProductionRun.Production_Run_Id > 0 Then
            If Me.chkPrintAtPackager.value <> ProductionRun.PrintAtPackager Then
                Call DeleteSample(ProductionRun.Production_Run_Id, 0)
                'change PDR sample counts for new samples just configured
                Me.txtSamples.text = 0
                Me.txtSampleGroups.text = 0
                ProductionRun.Samples_Requested = CLng(Me.txtSamples)
                ProductionRun.Sample_Number = CInt(Me.txtSampleGroups)
                ProductionRun.UpdateSampleQuantities
                mReprintFile_id = ProductionRun.Reprint_File_Id
            End If
        End If
        Screen.MousePointer = vbDefault
    End If
    
    If Not ValidScreen Then
        GoTo Exit_this_Sub
    End If
    
    If Not CheckOverrides Then
        GoTo Exit_this_Sub
    End If
    
    Screen.MousePointer = vbHourglass

    'added the following to prevent production runs from being saved if
    'their IRQs have been issued & the user is attempting to change the quantity or stock or IRQs have
    'been created and the user is attempting to change the stock.
    If Not oCollIRQInfo Is Nothing Then
        For Each oIRQInfo In oCollIRQInfo
            'If oIRQInfo.IRQ_Number = Me.txtStockIRQ Then
            If InStr(1, Me.txtStockIRQ, oIRQInfo.IRQ_Number) > 0 Then
                If HoldStockProofId <> ProductionRun.Stock_Proof_Id Then
                    MsgBox _
                        "Cannot Save the Production Run." & vbCrLf & _
                        "The Stock cannot be changed after an IRQ has been created.", vbExclamation
                    GoTo Exit_this_Sub
                End If
            ElseIf oIRQInfo.IRQ_Number = Me.txtScratchIRQ Then
                If HoldScratchStockProofId <> ProductionRun.Scratch_Proof_Id Then
                    MsgBox _
                        "Cannot Save the Production Run." & vbCrLf & _
                        "The Blinding Laminate Stock cannot be changed after an IRQ has been created.", vbExclamation
                    GoTo Exit_this_Sub
                End If
            End If
        Next
    End If
    
    If gLinkLock.IsLockStillValid = False Then
        MsgBox "You cannot save because your lock on the Links has been deleted." & vbCrLf & vbCrLf & "Please reload before continuing.", vbInformation + vbOKOnly, "Invalid Lock"
        GoTo Exit_this_Sub
    End If

    Call AppendRefNumber
    Call AppendClientReqdFields
    
    HasDataChanged
    
    If (Not (Determine_If_PDR_HasRun And (Not booNewProdRun))) Or booReplacement Then
            Call DupLinksSpecInstructions
    End If
    
    If booReplacement Then
        Call AppendOrigRunToSpecInstructions
    End If

    
    Set smpData = New CCOLsmpFiles
    booInitialSave = (ProductionRun.Production_Run_Id = 0)
    
    With ProductionRun
        .Prod_Run_Updated = False
        If Me.txtDirtyFlag.text = "Y" Then
            Me.txtProducedBy.text = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
        End If
        .Produced_By = Me.txtProducedBy
        .Samples_Requested = CLng(Me.txtSamples)
        .Sample_Number = CInt(Me.txtSampleGroups)
        .Prod_Description = Me.txtProdDesc
        .Reference_No = Me.txtReferanceNo
        If Not booReplacement Then
            .Reprint_File_Id = 0
        End If
        If CBool(Me.chkPrintAtPackager.value) Then
            .Ship_To_Id = 0
            .Ship_Description = ""
        Else
            .Ship_To_Id = SSDBComboShip.Columns.Item(1).text
            .Ship_Description = SSDBComboShip.text
        End If
        .PrintAtPackager = Me.chkPrintAtPackager.value  ' DW 2010-002 added
        If Not .Clintrak_Location_Id > 0 Then
            .Clintrak_Location_Id = gApplicationUser.ClintrakLocationId
        End If
        .Reorient_Ind = Me.chkReOrientation.value
        ' non billable reason is updated as the combo box is changed (so that the UI can be made read-only as necessary)
    End With
    
    'if we're in debug mode for a new production run and coding isn't updated we need to update coding for a new run
    'then update the coding link
    If (gDebugMode = True) And (booNewProdRun = True) And (InStr(1, ProductionRun.File_Name, gRnrDebugFilePath, vbTextCompare) <> 1) Then
        strTempFileName = Replace$(ProductionRun.File_Name, gRnrBaseFilePath, gRnrDebugFilePath, 1, -1, vbTextCompare)
        Call CopyFiles(ProductionRun.File_Name, strTempFileName)
        ProductionRun.File_Name = strTempFileName
        ProductionRun.SaveProdRun
        Call UpdateCodingFilePath(gCodingFileId, ProductionRun.Production_Run_Id, strTempFileName, 0)
    Else
       ProductionRun.SaveProdRun
    End If
    
    If ProductionRun.Prod_Run_Updated Then
        gProductionRun_Id = ProductionRun.Production_Run_Id     'assigns the global to the new production run id
        If Not CheckExistingSample Then
            If Save_Samples_FromPDR_Screen Then
                'change PDR sample counts for new samples just configured
                ProductionRun.Samples_Requested = CLng(Me.txtSamples)
                ProductionRun.Sample_Number = CInt(Me.txtSampleGroups)
                ProductionRun.UpdateSampleQuantities
            Else
                MsgBox _
                    "Production Run was Saved but the Samples were not Configured" & vbCrLf & _
                    "Please contact IT.", vbExclamation
            End If
        End If
        
        If booInitialSave Then
            Call LoadBarcodeInfo(ProductionRun.Production_Run_Id)
            Me.chkUseClientInventory.enabled = True
            If booReplacement Then
                ProductionRun.StatusLookupId = basGlobals.GetStatusID(APPROVED_STATUS)
            Else
                ProductionRun.StatusLookupId = basGlobals.GetStatusID(AWAITINGAPPROVAL_STATUS)
            End If
        End If
        
        'Save client required fields values
        If ClientReqdFields.count > 0 Then
            Call SaveClientReqdFields
        End If
        

        Call UpdateExistingIRQList

        MsgBox "Production Run Record successfully processed.", vbInformation
        
        Call UpdateCompletelyShippedFlag(ProductionRun.Job_Log_Id)
        
        txtDirtyFlag = ""
        mvarLinksReportDirty = False
        Me.cmdSpecInst.enabled = True
        Me.mvarBillingDirty = False
        
    Else
        MsgBox "Record was not Saved." & Chr$(13) & "Check data and retry.", vbCritical
        GoTo Exit_this_Sub
    End If
    
    ' hold main's barcode info
    If booReplacement = True Then
        strCodingFileHeader = ProductionRun.Coding_File_Header
    End If
    
    'refresh the class
    If (Not (Determine_If_PDR_HasRun And (Not booNewProdRun))) Or booReplacement Then
        'if we are here because this pdr is processed, then billing was changed and ,
        'we should not do this!
        ProductionRun.LookupRecord
        Me.txtBarcodeId = ProductionRun.Barcode_Id
        Me.cmdDeleteProdRun.enabled = True
        Call Load_Menu
    
    End If
    
    ' hold main's barcode info
    If booReplacement = True Then
        ProductionRun.Coding_File_Header = strCodingFileHeader
    End If
        
Exit_this_Sub:
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Saving Function"
    Resume Exit_this_Sub

End Sub

Private Function ValidScreen() As Boolean
    On Error GoTo Handle_Error
    
    Dim i As Long
    Dim intClientSamplesNotShipping As Long
    
    ValidScreen = False
    
    ' PDR Checks
    ' ---------------------------------------------------
    If Trim$(Me.txtDesc) = "" Then
        Me.txtDesc.SetFocus
        MsgBox "Description cannot be blank.", vbExclamation
        Exit Function
    End If
    If Trim$(Me.txtQty) = "" Or Not IsNumeric(Me.txtQty) Then
        Me.txtQty.SetFocus
        MsgBox "Quantity Requested must be entered.", vbExclamation
        Exit Function
    End If
    If ProductionRun.Apply_ScratchOff = NA Then
        MsgBox _
            "The Label is missing scratchoff information." & vbCrLf & _
            "This information must be entered in Label Specs before creating a Production Run.", vbExclamation
        Exit Function
    End If
    If Trim$(Me.SSDBComboShip.text) = "" And Not CBool(Me.chkPrintAtPackager.value) Then
        Me.SSDBComboShip.SetFocus
        MsgBox "Must enter the Ship To.", vbExclamation
        Exit Function
    End If
    If gClientRefReqInd = True And Trim$(Me.txtReferanceNo.text) = "" And Me.txtReplacement.Visible = False Then
        Me.txtReferanceNo.SetFocus
        MsgBox "Reference Number cannot be blank.", vbExclamation
        Exit Function
    End If
    If ProductionRun.NonBillable.reasonId < 0 Then
        MsgBox "The PDR must either be marked as billable or a non-billable reason must be selected.", vbExclamation
        Exit Function
    End If
    ' checking reference number for commas
    If InStr(Me.txtReferanceNo.text, ",") Then
        Me.txtReferanceNo.SetFocus
        MsgBox "Reference Number cannot contain any commas.", vbExclamation
        Exit Function
    End If
    'Check that all Client Required Fields have values
    For i = 1 To ClientReqdFields.count
        If Trim$(ClientReqdFields.Item(i).Field_Name_Value) = "" Then
            Me.cmdAddtlData.SetFocus
            MsgBox "Required Client Fields cannot be blank.", vbExclamation
            Exit Function
        End If
        If InStr(ClientReqdFields.Item(i).Field_Name_Value, ",") <> 0 Then
            Me.cmdAddtlData.SetFocus
            MsgBox "Required Client Fields cannot contain any commas.", vbExclamation
            Exit Function
        End If
    Next
    
    ' Sample Checks
    ' ---------------------------------------------------
    ' Call CheckExistingSample() to populate quantityTotal, sampleTypes
    ' and intClientSamplesNotShipping.
    Call CheckExistingSample(intClientSamplesNotShipping)
    If Me.quantityTotal <> CLng(Me.txtSamples) Then
        MsgBox _
            "Quantity Samples do not match configured." & vbCrLf & _
            "Samples:" & quantityTotal, vbExclamation
        Exit Function
    End If
    If Me.sampleTypes <> CInt(Me.txtSampleGroups) Then
        MsgBox _
            "Number of Sample Types do not match configured." & vbCrLf & _
            "Sample Types:" & sampleTypes, vbExclamation
        Exit Function
    End If
    If Not CBool(Me.chkPrintAtPackager.value) And intClientSamplesNotShipping > 0 Then
        MsgBox _
            "There are Client samples with no shipping address selected." & vbCrLf & _
            "Please make sure that all Client samples are properly configured.", vbExclamation
        Me.cmdSamples.SetFocus
        Exit Function
    End If
   
    ' Check Use Client Inventory for required reason
    If Not IsThereAnIRQ Then
        If Me.chkUseClientInventory.value <> Abs(ProductionRun.NonBillable.IsBillable) And _
        ProductionRun.UseClientInventoryReason = "" Then
            MsgBox _
                "The status of Use Client Inventory has changed and requires a reason to be entered." & vbCrLf & _
                "Please enter a reason and try again.", vbExclamation
                If Me.cmdReason.enabled = True Then Me.cmdReason.SetFocus
                Exit Function
        End If
    End If
    
    ValidScreen = True

Cleanup_Exit:
    Exit Function
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Function


'<comment>
' <summary>
'       Checks to see whether there are existing sample data for the current PDR
'       and sets the form variables quantityTotal and sampleTypes.</summary>
' <param name="intClientSamplesNotShipping">Optional ByRef parameter to tally up the
'       number of Sample Types that are not shipping (Job_Shipping_Id=0).</param>
' <return>Returns True if sample data exists.</return>
'</comment>
Private Function CheckExistingSample(Optional ByRef intClientSamplesNotShipping As Long = -1) As Boolean
    On Error GoTo Handle_Error

    Dim objData As nADOData.CADOData
    sampleTypes = 0
    quantityTotal = 0

    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .AddParameter "Type Number", 0, adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                sampleTypes = sampleTypes + 1
                quantityTotal = quantityTotal + .Recordset!quantity
                ' Tally up how many client samples are not shipping
                If intClientSamplesNotShipping > -1 Then
                    If .Recordset!Sample_Type <> "CLINTRAK" And .Recordset!Job_Shipping_Id = 0 Then
                        intClientSamplesNotShipping = intClientSamplesNotShipping + 1
                    End If
                End If
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With

    If sampleTypes >= 1 Then
        CheckExistingSample = True
    Else
        CheckExistingSample = False
    End If
    
Cleanup_Exit:
    Exit Function
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Function


'<comment>
' <summary>
'       Checks to see whether there was any data changed to the form, if there are
'       changes it prompts the user to save.</summary>
'</comment>
Private Function HasDataChanged() As Boolean
    On Error GoTo Handle_Error

    Dim i As Long
    
    HasDataChanged = False
    
    If ProductionRun.Production_Run_Id = 0 Then GoTo Flag_Dirty_Links
    If Trim$(Me.txtProdDesc) <> Trim$(ProductionRun.Prod_Description) Then GoTo Flag_Dirty_Links
    If Trim$(Me.txtReferanceNo) <> Trim$(ProductionRun.Reference_No) Then GoTo Flag_Dirty
    If Me.chkReOrientation.value <> ProductionRun.Reorient_Ind Then GoTo Flag_Dirty_Links
    If ProductionRun.Special_Inst <> ProductionRun.SpecialInstructions Then GoTo Flag_Dirty_Links ' Special_Inst = current, SpecialInstructions = original
    If ProductionRun.NonBillable.HasChange() Then GoTo Flag_Dirty
    If txtDirtyFlag.text = "Y" Then GoTo Flag_Dirty_Links
    If ProductionRun.HasChanges Then GoTo Flag_Dirty
    
    ' Determine whether client required fields have changeds
    If Not ClientReqdFields Is Nothing Then
        If ClientReqdFields.count > 0 Then
            For i = 1 To ClientReqdFields.count
                If mColClientFields.Item(i).Field_Name_Value <> ClientReqdFields.Item(i).Field_Name_Value Then
                    GoTo Flag_Dirty
                End If
            Next i
        End If
    End If
    
    ' check that descriptions match still
    If Not CBool(Me.chkPrintAtPackager.value) Then
        If (ProductionRun.Ship_To_Id <> SSDBComboShip.Columns.Item(1).text) Or _
                ProductionRun.Ship_Description <> SSDBComboShip.Columns(0).text Then
            GoTo Flag_Dirty_Links
        End If
    End If
    
    If txtDirtyFlag.text = "Y" Then
        ' if we found no changes above, still go through the process if the flag was manually set
        GoTo Flag_Dirty
    End If
    
    HasDataChanged = False
    
Cleanup_Exit:
    Exit Function
    
Flag_Dirty_Links:
    If booReplacement = False Then
        Me.mvarLinksReportDirty = True
    End If

Flag_Dirty:
    Me.txtDirtyFlag = "Y"
    Me.txtProducedBy.text = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
    HasDataChanged = True
    GoTo Cleanup_Exit
    
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Function

Private Sub CheckColumnNumbers()
'
'comments: This sub checks to see whether the number of columns in the
'          coding file matches the sample files. if not then it deletes all
'          the sample file and Sample_Types DB entries
'parameters:  none

Dim tmpString As String
Dim tmpLength As Integer
Dim tmpTotal As String
Dim SmplString As String
Dim CodeString As String
Dim ChgCodingFile As Boolean
Dim HoldSmplId As Long
Dim objData As nADOData.CADOData
Dim lngNumberSampleFileLines As Long
Dim strData As String
Dim i As Long


    'check to see if samples exist
    If Trim$(txtSampleGroups.text) = "" Or txtSampleGroups.text = "0" Then
        Exit Sub
    End If

    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
            .CursorType = adOpenForwardOnly
            .CommandType = adCmdStoredProc
            .LockType = adLockReadOnly

            .ResetParameters
            .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
            .AddParameter "Type Number", 0, adInteger, adParamInput
            .OpenRecordSetFromSP "get_SampleTypeInfo"
            
            'check to see if the column numbers are the same
            If Not .Recordset.EOF Then
                gSampleFileName = .Recordset!Sample_File_Name
                lngNumberSampleFileLines = GetNumberLinesInFile(gSampleFileName)
                                
                ChgCodingFile = False
                
                ' loop through all records in the sample file to compare to
                ' the corresponding record in the coding file
                For i = 1 To lngNumberSampleFileLines
                    ' get the line of data from the sample file
                    strData = GetLineOfData(gSampleFileName, i)
                    ' get the number of columns in the line of data
                    SampleColumns = CountDelimitedWords(strData, gRandDelimiter)
                    tmpString = GetDelimitedWord(strData, 1, gRandDelimiter)
                    tmpLength = Len(tmpString)
                    tmpTotal = Len(strData)
                    SmplString = Mid$(strData, tmpLength + 2, tmpTotal)
                                    
                    ' get the line of data from the coding file
                    strData = GetLineOfData(gCodingFileName, i)
                    ' get the number of columns in the line of data
                    codingColumns = CountDelimitedWords(strData, gRandDelimiter)
                    tmpString = GetDelimitedWord(strData, 1, gRandDelimiter)
                    tmpLength = Len(tmpString)
                    tmpTotal = Len(strData)
                    CodeString = Mid$(strData, tmpLength + 2, tmpTotal)
                    
                    ' this is required because it sets vdata to the first line and
                    ' this data is then used when creating a replacement PDR
                    If i = 1 Then
                        Read_File (gSampleFileName)
                        Read_File (gCodingFileName)
                    End If
                    
                              
                    ' compare the sample data to the coding data.
                    ' if different, display a message and exit the loop bc there's no need to continue checking.
                    If CodeString <> SmplString Then
                        ChgCodingFile = True
                        If booReplacement = False Then
                            basGlobals.UpdateLinkLockVersionAndUpdatedBy  ' deleted samples, so immediately update the links version
                        Else
                            basGlobals.UpdateLinkLockVersion
                        End If
                        If codingColumns <> SampleColumns Then
                            MsgBox _
                                "The number of coded fields from the original coding file has changed." & vbCrLf & _
                                "All the sample files will be DELETED! You MUST reconfigure ALL samples again!!!", vbExclamation
                        Else
                            MsgBox _
                                "The Coding File is different than the Sample Files" & vbCrLf & _
                                "The sample files will be deleted!  You MUST reconfigure ALL samples again!!!", vbExclamation
                        End If
                        Exit For
                    End If
                Next i
                 
            End If
            
            ' if there was a change in the coding file, then we MUST delete ALL
            ' sample files and make them reconfigure the samples.
            If ChgCodingFile Then
                Do While Not .Recordset.EOF
                    HoldSmplId = .Recordset!sample_type_id
                    gSampleFileName = .Recordset!Sample_File_Name
                    
                    If DeleteSample_By_SmplID(HoldSmplId) Then
                        Kill (gSampleFileName)
                    Else
                        MsgBox _
                            "There was an error deleting sample file " & gSampleFileName & vbCrLf & _
                            "Please contact IT.", vbExclamation
                            
                        .Recordset.Close
                        Exit Sub
                    End If
                    .Recordset.MoveNext
                Loop
            End If
            
            .Recordset.Close
    End With
        
End Sub
Private Sub GetSampleFileName()
'
'comments: Get sample file name for check
'parameters:  none

Dim objData As nADOData.CADOData
'check to see if samples exist
    If Trim$(txtSampleGroups.text) = "" Or txtSampleGroups.text = "0" Then
        Exit Sub
    End If

    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
            .CursorType = adOpenForwardOnly
            .CommandType = adCmdStoredProc
            .LockType = adLockReadOnly

            .ResetParameters
            .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
            .AddParameter "Type Number", 0, adInteger, adParamInput
            .OpenRecordSetFromSP "get_SampleTypeInfo"
                        
            If Not .Recordset.EOF Then
                gSampleFileName = .Recordset!Sample_File_Name
            End If
            .Recordset.Close
    End With
       
End Sub

Public Sub ActivateEditOption(on_off As Boolean)
    cmdSamples.enabled = on_off
    cmdSpecInst.enabled = on_off
    cmdDeleteProdRun.enabled = on_off
    cmdSave.enabled = on_off
    txtProdDesc.enabled = on_off
    txtReferanceNo.enabled = on_off
End Sub

Private Function CheckProdRunExist() As Boolean
'
'comments: this function checks to see whether the production run exists or not
'parameters: none
'returns: boolean - true if production run exists

    If ProductionRun.Production_Run_Id = 0 Or IsNull(ProductionRun.Production_Run_Id) Then
        CheckProdRunExist = False
    Else
        CheckProdRunExist = True
    End If

End Function

Private Function UpdateCompletelyShippedFlag(lngJobLogId As Long) As Boolean
    On Error GoTo Handle_Error

    Dim lngError As Long
    Dim objData As nADOData.CADOData
    
    UpdateCompletelyShippedFlag = False

    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
        .ResetParameters
        .AddParameter "Job_Log_Id", lngJobLogId, adInteger, adParamInput
        .AddParameter "@in_completely_shipped_flag", " ", adVarChar, adParamInput
        .AddParameter "error", "      ", adInteger, adParamOutput
        .ExecuteSP "update_Job_CompletelyShipped_Flag", True
        .RetrieveParameters
        lngError = .GetParameterValue("error")
    End With

    UpdateCompletelyShippedFlag = (lngError = 0)
    
    If Not UpdateCompletelyShippedFlag Then
        MsgBox _
            "The Job Completely Shipping flag may not have been set correctly." & vbCrLf & _
            "Please contact IT.", vbExclamation
    End If
    
Cleanup_Exit:
    Exit Function
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Function

Private Function Save_Samples_FromPDR_Screen() As Boolean
'md added for clintrak sampels called from from prodplan to automatically configure
'2 clintrak samples on save of PDR with any replacement samples as well

On Error GoTo PROC_ERR

    Save_Samples_FromPDR_Screen = False
    mSampleFileName = ""
    
    'check to ensure that we can use the coding file to retrieve the live data for the Clintrak samples
    'if there are only samples for this PDR then we must force the creation of dummy Clintrak samples
    If Me.txtQty = 0 Then
        Call Get_SampleFile_Layout
    Else
        ' load 2 clintrak samples to collection for use in creation of file with live data
        ' If Print at Packager is true do not add the "Clintrak" Samples
        If Not CBool(Me.chkPrintAtPackager.value) Then
            Call ReadProcess_File_CTK(ProductionRun.File_Name, 2, "CLINTRAK")
        End If
    End If
    
    ' If Print at Packager is true do not add the "Clintrak" Samples
    If Not CBool(Me.chkPrintAtPackager.value) Then
        'create clintrak sample file
        Call CollectionToFileCTK(mSampleFileName)
        gSampleFileName = mSampleFileName
        
        'insert the samples to a collection to load the sample_types table
        'this will be Clintrak and any other samples created for a replacement PDR if one is being created
        Call Format_SampleType_Table_For_Clintrak
    End If
    
    If booReplacement Then
        Call Get_ReprintSamples_ID_For_Replacements
    End If

    'md save the collection of samples to the table
    If Save_PDRSample_Info() Then
         Save_Samples_FromPDR_Screen = True
    End If
    
Proc_EXIT:
    Exit Function
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Error Saving Samples from PDR screen", vbExclamation
    Resume Proc_EXIT

End Function
Private Function Fix_Samples_FromPDR_Screen() As Boolean
'sample fix code
'fix for 0 length sample file
'md added for clintrak sampels called from from prodplan to automatically configure
'2 clintrak samples on save of PDR with any replacement samples as well

On Error GoTo PROC_ERR
    
      
    Fix_Samples_FromPDR_Screen = False
    mSampleFileName = ""
    
    'check to ensure that we can use the coding file to retrieve the live data for the Clintrak samples
    'if there are only samples for this PDR then we must force the creation of dummy Clintrak samples
    If Me.txtQty = 0 Then
        Call Get_SampleFile_Layout
    Else
        ' load 2 clintrak samples to collection for use in creation of file with live data
        ' If Print at Packager is true do not add the "Clintrak" Samples
        If Not CBool(Me.chkPrintAtPackager.value) Then
            Call ReadProcess_File_CTK(ProductionRun.File_Name, 2, "CLINTRAK")
        End If
    End If
    
    ' If Print at Packager is true do not add the "Clintrak" Samples
    If Not CBool(Me.chkPrintAtPackager.value) Then
        'create clintrak sample file
        Call CollectionToFileCTK(mSampleFileName)
        gSampleFileName = mSampleFileName
        If FileLen(gSampleFileName) <> 0 Then
            Fix_Samples_FromPDR_Screen = True
        End If
    End If
Proc_EXIT:
    Exit Function
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Error Saving Samples from PDR screen", vbExclamation
    Resume Proc_EXIT
End Function
Private Sub CollectionToFileCTK(strDestination As String)

'md new for clintrak samples
'
'comments: takes the data from the collection and places into a file
'parameters: data - collection to parse through
'          : strDestination - destination of file
'returns: True if the collection is saved to file
'
On Error GoTo Error_this_Sub
  
   Dim strFileName As String
   
    'checks to see if the "smp" directory exists
    If Not FileExists(strDestination) Then
         strFileName = GetFilePath(ProductionRun.File_Name)
         strFileName = GetFilePath(strFileName) & "\smp\"
         If Not DirExists(strFileName) Then
            'creates the smp directory if it doesn't
             MkDir (strFileName)
         End If
             strFileName = strFileName & Trim$(ProductionRun.Barcode_Id) & "_1.smp"
             mSampleFileName = strFileName
    Else
             strFileName = strDestination
             mSampleFileName = strDestination
    End If
            
    'calls the function to write the collection to a file
    Call WriteFile(mData, "CTK", Trim$(strFileName))
                 
Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Saving sample Clintrak samples to file"
    Resume Exit_this_Sub
    
End Sub

Private Function Save_PDRSample_Info() As Boolean
'
'md new for clintrak samples
'comments: this sub calls the stored procedure to save or update the sample info
'
On Error GoTo Error_this_Sub

    Dim nreturn As Long
    Dim i As Long
    Dim objData As nADOData.CADOData
    
    Save_PDRSample_Info = False
    
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
    End With
    
    gadoConnection.Connection.BeginTrans
    
    For i = 1 To smpData.count
    
'      With madoData
    With objData
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
                            
        .AddParameter "Sample Type Id", 0, adInteger, adParamInput
        .AddParameter "Production Run Id", smpData.Item(i).productionId, adInteger, adParamInput
        .AddParameter "Type Number", smpData.Item(i).typeNumber, adInteger, adParamInput
        .AddParameter "Sample Type", smpData.Item(i).sampleType, adVarChar, adParamInput
        .AddParameter "Ship To Id", smpData.Item(i).shipTo, adInteger, adParamInput
        .AddParameter "Quantity", smpData.Item(i).quantity, adInteger, adParamInput
        .AddParameter "Sample File Name", smpData.Item(i).smpfileName, adVarChar, adParamInput
        .AddParameter "Sample Description", CheckNulls(smpData.Item(i).smpDescription), adChar, adParamInput

        .AddParameter "return", "   ", adInteger, adParamOutput ' the "   " is for a length value
        .AddParameter "identity", "   ", adInteger, adParamOutput ' the "   " is for a length value
        
        .ExecuteSP "save_SampleFiles", True
        
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
            .Connection.RollbackTrans
            Save_PDRSample_Info = False
            Exit Function
        End If
        
      End With
      
    Next i
    
    gadoConnection.Connection.CommitTrans
    Save_PDRSample_Info = True
    
Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Saving Clintrak Sample Information"
    gadoConnection.Connection.RollbackTrans
    Resume Exit_this_Sub

End Function

Public Sub ReadProcess_File_CTK(FileSource As String, procquantity As Long, cmbotype As String)
'md new code for clintrak samples on automatic population
'comments:  reads the file and processes the correct records
'parameters: FileSource - path of file to read
'            procquantity - what position in the file to start the read
'            cmbotype - the selected sample type

Dim lngSourceFile As Long
Dim i As Long

On Error GoTo PROC_ERR

' Open the source file
lngSourceFile = FreeFile

vdata = ""

Open FileSource For Input Access Read As lngSourceFile

'first initialize the existing collection grid then populate with count
Set mData = New CCOLPDRFILES

' can only read the file up to the total number of file entries, so replace
If procquantity > Me.txtQty Then
    Line Input #lngSourceFile, vdata
    For i = 1 To procquantity
        Call mData.Add(vdata, i, cmbotype)
    Next i
Else
    For i = 1 To procquantity
        Line Input #lngSourceFile, vdata
        Call mData.Add(vdata, i, cmbotype)
    Next i

End If

' Close file
Close lngSourceFile

Proc_EXIT:
    Exit Sub
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "ReadProcess File"
    Resume Proc_EXIT

End Sub

Private Sub Get_ReprintSamples_ID_For_Replacements()
'get sample reprints

    Dim index As Long
    Dim objData As nADOData.CADOData
    
    On Error GoTo Error_this_Sub

    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        
        ' Call the SP to create the resultset
        .ResetParameters
        
        .AddParameter "Reprint File Id", mReprintFile_id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_Reprint_Sample_Files_By_ReprintFile_Id"

        ' loading record set
        Do While Not .Recordset.EOF
            Me.txtSamples = Me.txtSamples + .Recordset!File_Record_Count
            Me.txtSampleGroups = Me.txtSampleGroups + 1
            index = Me.txtSampleGroups
            
            'load the collection
              Call smpData.Add(ProductionRun.Production_Run_Id, Me.txtSampleGroups, .Recordset!Sample_Type, ProductionRun.Ship_To_Id, _
                .Recordset!File_Record_Count, .Recordset!Sample_Type, Determine_ReplacementSample_Path(index), .Recordset!sample_type_id)
                
            'create the file from the reprint sample file
            Call FileCopy(.Recordset!Reprint_Sample_File_Name, smpData.Item(index).smpfileName)
            
            .Recordset.MoveNext
        Loop
        .Recordset.Close
    End With
    
Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR - Get Sample Reprints"
    Resume Exit_this_Sub
    
End Sub

Private Sub Format_SampleType_Table_For_Clintrak()
'md new for clintrak samples
'format/add to the collection for the clintrak samples for the sample_type table

Dim tnotes As String

 Me.txtSamples = 2
 Me.txtSampleGroups = 1

 Call smpData.Add(ProductionRun.Production_Run_Id, 1, "CLINTRAK", 0, 2, "CLINTRAK", gSampleFileName, 0) ' DW 2010-002 added final parameter
    
End Sub

Private Function Determine_ReplacementSample_Path(typecount As Long) As String

'md new for clintrak samples
'
'comments: gets the correct path for samples from reprints

On Error GoTo Error_this_Sub
  
   Dim strFileName As String

   Determine_ReplacementSample_Path = ""
      
   strFileName = GetFilePath(ProductionRun.File_Name)
   strFileName = GetFilePath(strFileName) & "\smp\"
   If Not DirExists(strFileName) Then
       'creates the smp directory if it doesn't
     MkDir (strFileName)
   End If
   strFileName = strFileName & Trim$(ProductionRun.Barcode_Id) & "_" & typecount & ".smp"
   Determine_ReplacementSample_Path = strFileName
              
Exit_this_Sub:
    Exit Function

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Determine Replacement File Path"
    Resume Exit_this_Sub
    
End Function


Private Sub Lock_Out_PDRForm(booval1 As Boolean, booval2 As Boolean)
    
    cmdDeleteProdRun.enabled = (ProductionRun.Production_Run_Id <> 0) And booval1
    
    txtPDRStatus.Visible = booval2
    
Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR - Lock Out PDR Form"
    Resume Exit_this_Sub
    
End Sub

Private Sub GetReprintFile_For_PDR(Found As Boolean)

  'md added for samples project - call to get the reprint sample type for display use
    Dim objData As nADOData.CADOData
    On Error GoTo PROC_ERR
  
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        .RowsetSize = 1
   
        .AddParameter "Reprint File Id", ProductionRun.Reprint_File_Id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ReprintFileByID"
        If .Recordset.EOF Then
            Found = False
            MsgBox "The selected Production Run's Reprint File was not found.", vbExclamation
            Exit Sub
         Else
            Found = True
            gReprintFile_Type = .Recordset!Reprint_File_Type
         End If
    End With
         
Proc_EXIT:
    Exit Sub
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Error Getting Reprint File for PDR"
    Resume Proc_EXIT
         
End Sub

Public Function Determine_Shipping_Flag_On()

'md new for clintrak samples
'comments: this calls Job Assignment Log for Shipped flag
'
    Dim objData As nADOData.CADOData
     
    On Error GoTo Error_this_Sub
   
    
    Determine_Shipping_Flag_On = False
        
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
                    
        .AddParameter "Job Log ID", ProductionRun.Job_Log_Id, adInteger, adParamInput
                      
        .OpenRecordSetFromSP "get_JobAssignmentByJobLog"
        
        If Not .Recordset.EOF Then
            If .Recordset!Completely_Shipped_Flag = "Y" Then
               Determine_Shipping_Flag_On = True
            End If
        End If
        .Recordset.Close
    End With
                      
Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Checking Shipping Flag"
    Resume Exit_this_Sub

End Function

Private Sub Get_SampleFile_Layout()

Dim subdata As String
Dim lngSourceFile As Long
Dim i As Long

On Error GoTo Error_this_Sub

    'first initialize the existing collection grid then populate with count
    Set mData = New CCOLPDRFILES
    
    subdata = ""
    ' Open the source file
    lngSourceFile = FreeFile

    Open mInputFileName For Input Access Read As lngSourceFile
   
    For i = 1 To 2
        Line Input #lngSourceFile, subdata
        Call mData.Add(subdata, i, "CTK")
    Next i
    
    Close lngSourceFile
    
Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR - Get Sample File Layout"
    Resume Exit_this_Sub
    
End Sub

Public Sub GetIRQInfo()
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
            
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_PDR_IRQ_Info"
            
        If Not oCollIRQInfo Is Nothing Then
            Set oCollIRQInfo = Nothing
        End If
        
        If Not .Recordset.EOF Then
            Set oCollIRQInfo = New CCollIRQInfo
            Do Until .Recordset.EOF
                Set oIRQInfo = oCollIRQInfo.Add
                             oIRQInfo.IRQ_Proof_Id = .Recordset!Proof_Id
                             oIRQInfo.IRQ_Id = .Recordset!Inventory_Request_Id
                             oIRQInfo.IRQ_Number = .Recordset!IRQ_Number
                             oIRQInfo.PDR_Count = 0
                             oIRQInfo.IRQ_Label_Identification = .Recordset!label_identification
                             oIRQInfo.IRQ_Details_Id = .Recordset!Inventory_Request_Details_Id
                             oIRQInfo.IRQ_Details_Qty_Requested = .Recordset!Qty_Requested
                             oIRQInfo.IRQ_Status = .Recordset!Status
                             oIRQInfo.IRQ_Main_Proof_Id = .Recordset!Main_Proof_Id
                             oIRQInfo.IRQToolComponentId = .Recordset!Tool_Component_ID
                .Recordset.MoveNext
            Loop
            
            .Recordset.Close
        End If
    End With


Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error - "
    Resume Exit_this_Sub

End Sub


Private Function GetNumberPDRsForIRQ(lngIRQ_Id As Long) As Long
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
            
    GetNumberPDRsForIRQ = 0
            
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "IRQ", lngIRQ_Id, adInteger, adParamInput
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_IRQ_PDR_Count"
            
        If Not .Recordset.EOF Then
            '
            GetNumberPDRsForIRQ = .Recordset!PDR_Count
            .Recordset.Close
        End If
    End With


Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error - "
    GetNumberPDRsForIRQ = 0
    Resume Exit_this_Sub

End Function


' DW 2012-001 added
Private Function checkForAssociatedDigitalOverageOrder(prodrunbarcode As String) As String
'
'comments:  This function gets the Order Number from the Orders table where the PDR Barcode matches the Onsert_Computerization_Run field
'parameter: prodrunbarcode - production run barcode
'returns:   string representing the CLK Order Number or N/A if it doesn't
'
On Error GoTo Handle_Error
    
    Dim objData As nADOData.CADOData
    
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        .RowsetSize = 1

        ' Call the SP to create the resultset
        .ResetParameters
        .AddParameter "PDRBarcode", prodrunbarcode, adVarChar, adParamInput
        .OpenRecordSetFromSP "get_Digital_Open_Overage_Orders_By_RunBarCode"

        ' loading record set
        ' Only cares about the first record returned.
        ' Order_Number:
        '               N/A             - No associations what so ever
        '               N/A w/Orders    - IDO Orders are available, but we're still N/A
        '               Order           - CLK of associated Order
        If Not .Recordset.EOF Then
            checkForAssociatedDigitalOverageOrder = .Recordset!Order_Number
        End If
        
        .Recordset.Close
    End With
    
   
exit_function:
    Exit Function

Handle_Error:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "checkForAssociatedDigitalOverageOrder", vbExclamation
    Resume exit_function

End Function

Private Function Get_SampleQTY(prodrunid As Long, SampletypeNum As Integer) As Long
'
'comments:  This function gets the sameple quantities for the sample type number passed in
'parameter: prodrunid - production run id, sampletypenum - sample type number (use 0 to return totals for all sample types)
'returns:   total for the specified sample type
'
On Error GoTo Handle_Error
    
    Dim QtyTotal As Long
    Dim objData As nADOData.CADOData
    
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        .RowsetSize = 1

        ' Call the SP to create the resultset
        .ResetParameters
        .AddParameter "Production Run Id", prodrunid, adInteger, adParamInput
        .AddParameter "Type Number", SampletypeNum, adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"

        ' loading record set
        Do While Not .Recordset.EOF
            QtyTotal = QtyTotal + .Recordset!quantity
            .Recordset.MoveNext
        Loop
        .Recordset.Close
    End With
    
    Get_SampleQTY = QtyTotal

exit_function:
    Exit Function

Handle_Error:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Get_SampleQTY", vbExclamation
    Resume exit_function
    
End Function

'<comment>
' <summary>
'       This sub checks to see if there are links special instructions and if so,
'       if they aren't already in the PDR's special instructions and if not,
'       asks the user if he/she would like to append them and if so,
'       appends them and saves to the database</summary>
'</comment>
Private Sub DupLinksSpecInstructions()
    On Error GoTo Handle_Error

    If Trim$(Me.txtLinksSpecialInstructions.text) <> "" Then
        Me.txtPdrSpecialInstructions.text = ProductionRun.Special_Inst
        If InStr(Me.txtPdrSpecialInstructions.text, Trim$(Me.txtLinksSpecialInstructions.text)) = 0 Then
        
            If MsgBox( _
                "Would you like to append the Links Special Instructions " & vbCrLf & _
                "to the PDR's Special Instructions?", vbQuestion + vbYesNo) = vbYes Then
                
                If Trim$(Me.txtPdrSpecialInstructions.text) <> "" Then
                    Me.txtPdrSpecialInstructions.text = _
                        Me.txtPdrSpecialInstructions.text & vbCrLf & vbCrLf & _
                        Me.txtLinksSpecialInstructions.text & " - " & _
                        Now & " " & gClintrakLocations(gApplicationUser.ClintrakLocationId).Time_Zone_Display
                Else
                    Me.txtPdrSpecialInstructions.text = _
                        Me.txtLinksSpecialInstructions.text & " - " & _
                        Now & " " & gClintrakLocations(gApplicationUser.ClintrakLocationId).Time_Zone_Display
                End If
                
                
                If Len(Me.txtPdrSpecialInstructions.text) > _
                    Len(Me.txtLinksSpecialInstructions.text & " - " & _
                        Now & " " & gClintrakLocations(gApplicationUser.ClintrakLocationId).Time_Zone_Display) Then
                    Load frmSpecInst
                    frmSpecInst.txtInstructions = Me.txtPdrSpecialInstructions.text
                    ' Do not allow users to cancel
                    frmSpecInst.OKButton.enabled = True
                    frmSpecInst.CancelButton.enabled = False
                    frmSpecInst.Show vbModal
                End If
            End If
        End If
        
        ProductionRun.Special_Inst = Me.txtPdrSpecialInstructions.text
        
    End If
    
Cleanup_Exit:
    Exit Sub
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Sub


Private Sub AppendRefNumber()

    On Error GoTo Error_this_Sub

    If Trim$(ProductionRun.Reference_No) <> "" Then 'there had previously been a reference no. for the pdr
            If InStr(Me.txtProdDesc.text, Trim$(ProductionRun.Reference_No)) > 0 Then 'the old reference no. had been in the desc.
                If Trim$(Me.txtReferanceNo.text) = "" Then 'the reference no. has been removed so remove it from the desc.
                    Me.txtProdDesc.text = Replace(Me.txtProdDesc.text, Trim$(ProductionRun.Reference_No), "")
                Else 'the reference no. has just been changed, so swap the old one in the desc. for the new one
                    Me.txtProdDesc.text = Replace(Me.txtProdDesc.text, Trim$(ProductionRun.Reference_No), Trim$(Me.txtReferanceNo.text))
                End If
            Else 'the old reference no. was not in the desc. so add it
                If InStr(Me.txtProdDesc.text, Me.txtReferanceNo.text) = 0 Then
                    Me.txtProdDesc.text = Me.txtProdDesc.text & " " & Trim$(Me.txtReferanceNo.text)
                End If
            End If
    Else 'there hadn't previously been a reference no. for the pdr
        If Trim$(Me.txtReferanceNo.text) <> "" Then 'there is now a reference no for the pdr, so add to desc
            If InStr(Me.txtProdDesc.text, Me.txtReferanceNo.text) = 0 Then
                Me.txtProdDesc.text = Me.txtProdDesc.text & " " & Trim$(Me.txtReferanceNo.text)
            End If
        End If
    End If
    
Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR - Append Reference Number to Description"
    Resume Exit_this_Sub
    
End Sub

Private Sub AppendClientReqdFields()
    Dim i As Long
    
    On Error GoTo Error_this_Sub
        
    For i = 1 To ClientReqdFields.count
        If mColClientFields.Item(i).Field_Name_Value <> "" Then ' there had previously been a value for the client field
            If InStr(Me.txtProdDesc.text, mColClientFields.Item(i).Client_Required_Field_Name & " " & mColClientFields.Item(i).Field_Name_Value) > 0 Then 'the old value had been in the desc.
                If ClientReqdFields.Item(i).Field_Name_Value <> "" Then
                    Me.txtProdDesc.text = Replace(Me.txtProdDesc.text, mColClientFields.Item(i).Field_Name_Value, ClientReqdFields.Item(i).Field_Name_Value)
                End If
            Else 'the old value was not in the desc. so add it
                If InStr(Me.txtProdDesc.text, ClientReqdFields.Item(i).Client_Required_Field_Name & " " & ClientReqdFields.Item(i).Field_Name_Value) = 0 Then
                    Me.txtProdDesc.text = Me.txtProdDesc.text & " " & _
                                            ClientReqdFields.Item(i).Client_Required_Field_Name & " " & _
                                    ClientReqdFields.Item(i).Field_Name_Value
                End If
            End If

        Else ' there hadn't previously been a value for this client field
            If ClientReqdFields.Item(i).Field_Name_Value <> "" Then 'there is now a value for the client field
                Me.txtProdDesc.text = Me.txtProdDesc.text & " " & _
                                    ClientReqdFields.Item(i).Client_Required_Field_Name & " " & _
                                    ClientReqdFields.Item(i).Field_Name_Value
            End If
        End If
    Next i
    
Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR - Append Client Required Fields to Description"
    Resume Exit_this_Sub
    
End Sub

'************************************************************
' Saves collection to the Production_Run_Client_Fields table
'************************************************************
Private Sub SaveClientReqdFields()
    Dim i As Long
    Dim objData As nADOData.CADOData

    On Error GoTo Handle_Error
    
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        .RowsetSize = 1

        For i = 1 To ClientReqdFields.count
       
            .ResetParameters
            .AddParameter "Production Run Client Fields Id", ClientReqdFields.Item(i).Production_Run_Client_Fields_Id, adInteger, adParamInput
            .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
            .AddParameter "Client Required Field Name", ClientReqdFields.Item(i).Client_Required_Field_Name, adVarChar, adParamInput
            .AddParameter "Field Name Value", CheckNulls(ClientReqdFields.Item(i).Field_Name_Value), adVarChar, adParamInput
            .AddParameter "ModifiedByEmployeeId", gApplicationUser.employeeId, adInteger, adParamInput

            .AddParameter "error", "    ", adInteger, adParamOutput
            .AddParameter "identity", "    ", adInteger, adParamOutput
                
            .ExecuteSP "save_Production_Run_Client_Fields", True
    
            .RetrieveParameters
    
            If .GetParameterValue("error") <> 0 Then
                GoTo Handle_Error
            Else
                ClientReqdFields.Item(i).Production_Run_Client_Fields_Id = .GetParameterValue("identity")
            End If
            
        Next i
    
    End With
    
    ' Refresh State Holder
    Set mColClientFields = ClientReqdFields.Clone
              
Exit_Sub:
    Exit Sub
    
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Exit_Sub
    
End Sub

Private Sub AppendOrigRunToSpecInstructions()
'
'comments:  this sub appends the replacement run's original PRG or PDR to
'           the special instrutions
'parameters: none
'returns:    nothing

    Dim strOrigRun As String
    
    strOrigRun = "Original Run: " & GetOriginalRunBarcode

        Me.txtPdrSpecialInstructions.text = ProductionRun.Special_Inst
        
        If InStr(Me.txtPdrSpecialInstructions.text, strOrigRun) = 0 Then
            Me.txtPdrSpecialInstructions.text = Me.txtPdrSpecialInstructions.text & vbCrLf & strOrigRun
        End If
        
    ProductionRun.Special_Inst = Me.txtPdrSpecialInstructions.text
    
End Sub


Private Function GetOriginalRunBarcode() As String
    Dim objData As nADOData.CADOData
    Dim strPDR As String
    Dim strPRG As String
    On Error GoTo Error_this_Sub
            
    GetOriginalRunBarcode = ""
            
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Production Run Id", ProductionRun.Orig_Prod_Run_Id, adInteger, adParamInput
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_OriginalRunGroup_by_ProdRunId"
            
        If Not .Recordset.EOF Then
            '
            strPDR = .Recordset!Orig_PDR_Barcode
            strPRG = .Recordset!Orig_PRG_Barcode
            .Recordset.Close
        End If
    End With
    
    If strPRG = "" Then
        GetOriginalRunBarcode = strPDR
    Else
        GetOriginalRunBarcode = strPRG
    End If


Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error - "
    GetOriginalRunBarcode = ""
    Resume Exit_this_Sub

End Function

Private Sub chkPrintAtPackager_Click()
    On Error GoTo Handle_Error
    
    If CBool(Me.chkPrintAtPackager.value) Then
        If Me.txtStockIRQ.text <> "" Then
            MsgBox _
                "Inventory for this PDR has been requested on " & Me.txtStockIRQ.text & "." & vbCrLf & _
                "Please remove the PDR from the IRQ, reload this screen and try again.", vbExclamation
            Me.chkPrintAtPackager.value = 0
        Else
            Me.SSDBComboShip.text = ""
            Me.SSDBComboShip.enabled = False
            Me.cmdViewShipping.enabled = False
        End If
    Else
        ' Enable combo if it was screen editable to begin with
        Me.SSDBComboShip.enabled = Me.cmdSave.enabled
        Me.cmdViewShipping.enabled = (Me.SSDBComboShip.text <> "")
    End If
    
    Me.txtDirtyFlag.text = "Y"
    
Cleanup_Exit:
     Exit Sub
Handle_Error:
     MsgBox Err.description & vbCrLf & _
         "in frmProdPlan.chkComputerizeAtPackager_Click ", _
         vbCritical + vbOKOnly, "Application Error"
     Resume Cleanup_Exit
End Sub

Private Sub txtDi1rtyFlag_Change()
    If LCase$(Me.txtDirtyFlag.text) = "y" Then
        Me.txtProducedBy.text = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
    End If
End Sub

' maintain the current use or status of the Client Inventory checkbox
Private Sub ManageClientInventoryCheckBox()
    Select Case ProductionRuns.booNewProdRun
        Case True
            Me.chkUseClientInventory.value = Abs(ProductionRun.NonBillable.IsBillable)
            Me.chkUseClientInventory.enabled = False
            Me.cmdReason.enabled = False
            ProductionRun.UseClientInventory = IIf(Me.chkUseClientInventory.value = 1, True, False)
        Case False  'existing PDR
            If ProductionRun.Production_Run_Id = 0 Then
                If ProductionRun.HasExistingUseClientInventoryReason Then ProductionRun.ResetUseClientInventoryReasonDetails
                If booReplacement Then
                    Me.chkUseClientInventory.enabled = False
                    Me.chkUseClientInventory.value = Abs(ProductionRun.NonBillable.IsBillable)
                End If
                Me.cmdReason.enabled = False
            Else
                Me.chkUseClientInventory.value = Abs(ProductionRun.UseClientInventory)
                'Reason button should be available if there exists a reason entry
                If Not ProductionRun.HasExistingUseClientInventoryReason Then Me.cmdReason.enabled = False
                Me.chkUseClientInventory.enabled = Not IsThereAnIRQ
            End If
            ProductionRun.UseClientInventory = IIf(Me.chkUseClientInventory.value = 1, True, False)
    End Select
   
End Sub

Private Function IsThereAnIRQ() As Boolean
    If frmProdPlan.oCollIRQInfo Is Nothing Then Call GetIRQInfo
    If Not frmProdPlan.oCollIRQInfo Is Nothing Then
        For Each oIRQInfo In frmProdPlan.oCollIRQInfo
            If (IIf(InStr(1, frmProdPlan.txtStockIRQ, CStr(oIRQInfo.IRQ_Number)) > 0, True, False) Or oIRQInfo.IRQ_Number = frmProdPlan.txtScratchIRQ) Then
                    IsThereAnIRQ = True
            End If
        Next
    End If
    
End Function

Public Sub FlagChange()
    Me.txtDirtyFlag = "Y"
    Me.txtProducedBy = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
End Sub

Private Sub UpdateExistingIRQList()
    txtStockIRQ.text = ""
    txtScratchIRQ.text = ""
    Call GetIRQInfo
    If Not oCollIRQInfo Is Nothing Then
        For Each oIRQInfo In oCollIRQInfo
            oIRQInfo.PDR_Count = GetNumberPDRsForIRQ(oIRQInfo.IRQ_Id)
            
            If oIRQInfo.IRQType = Label And (oIRQInfo.IRQ_Proof_Id = ProductionRun.Stock_Proof_Id Or oIRQInfo.IRQ_Main_Proof_Id = ProductionRun.Stock_Proof_Id) Then
                Me.txtStockIRQ = IIf(Me.txtStockIRQ = "", oIRQInfo.IRQ_Number, Me.txtStockIRQ & "," & oIRQInfo.IRQ_Number)
                
            ElseIf (oIRQInfo.IRQType = Label) And (oIRQInfo.IRQ_Proof_Id = ProductionRun.Scratch_Proof_Id) Then
                Me.txtScratchIRQ = oIRQInfo.IRQ_Number
                
            ElseIf (oIRQInfo.IRQType = RollStock) And (oIRQInfo.IRQToolComponentId = ProductionRun.OnsertBaseRollStockComponentId) Then
                Me.txtStockIRQ = IIf(Me.txtStockIRQ = "", oIRQInfo.IRQ_Number, Me.txtStockIRQ & "," & oIRQInfo.IRQ_Number)
                
            End If
        Next
    End If
End Sub

