VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSmpConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Configuration"
   ClientHeight    =   6675
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8415
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
   ScaleHeight     =   6675
   ScaleWidth      =   8415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New Type"
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
      Left            =   1320
      TabIndex        =   41
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   8415
      Begin VB.TextBox txtSmpTypeNum 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "1"
         Top             =   120
         Width           =   495
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Configuration Type "
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
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   7785
      End
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
      Height          =   2625
      Left            =   120
      TabIndex        =   24
      Top             =   680
      Width           =   8175
      Begin VB.TextBox txtAdd3 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1360
         Width           =   3135
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBComboSmpType 
         Height          =   315
         Left            =   1155
         TabIndex        =   5
         Top             =   240
         Width           =   2415
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
         DividerType     =   0
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.CommandButton Update 
         Caption         =   "Append/Delete"
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
         Left            =   2010
         TabIndex        =   8
         Top             =   1170
         Width           =   1560
      End
      Begin VB.TextBox txtShip 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   160
         Width           =   3135
      End
      Begin VB.TextBox txtAttn 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   460
         Width           =   3135
      End
      Begin VB.TextBox txtAdd1 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   760
         Width           =   3135
      End
      Begin VB.TextBox txtAdd2 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1060
         Width           =   3135
      End
      Begin VB.TextBox txtCity 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1660
         Width           =   3135
      End
      Begin VB.TextBox txtState 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1960
         Width           =   3135
      End
      Begin VB.TextBox txtZip 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2260
         Width           =   3135
      End
      Begin VB.TextBox txtQtynumber 
         Height          =   285
         Left            =   1155
         TabIndex        =   7
         Top             =   1170
         Width           =   735
      End
      Begin VB.CommandButton cmdLoadCoding 
         Caption         =   "Load 'As Per Sequence'"
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
         Left            =   1155
         TabIndex        =   9
         Top             =   1530
         Width           =   2415
      End
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   1155
         TabIndex        =   10
         Top             =   2040
         Width           =   2415
      End
      Begin SSDataWidgets_B.SSDBCombo SSDBComboShip 
         Height          =   315
         Left            =   1155
         TabIndex        =   6
         Top             =   720
         Width           =   2415
         DataFieldList   =   "Column 0"
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
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lblAddr3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addr Line 3:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   40
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sample Type:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblShipToCbo 
         Caption         =   "Ship To:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1185
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Description:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblShipTo 
         Caption         =   "Ship To:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   31
         Top             =   255
         Width           =   735
      End
      Begin VB.Label lblAttn 
         Caption         =   "Attention:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   555
         Width           =   735
      End
      Begin VB.Label lblAddr1 
         Caption         =   "Addr Line 1:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   29
         Top             =   855
         Width           =   975
      End
      Begin VB.Label lblAddr2 
         Caption         =   "Addr Line 2:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   1155
         Width           =   975
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   1695
         Width           =   495
      End
      Begin VB.Label lblState 
         Caption         =   "State:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   1995
         Width           =   615
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   2295
         Width           =   375
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
      Height          =   2655
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   8175
      Begin VB.CommandButton cmdNotes 
         Caption         =   "Notes"
         Height          =   255
         Left            =   7080
         TabIndex        =   13
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdConfigure 
         Caption         =   "Configure"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtcolumn 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2265
         Width           =   1215
      End
      Begin GridEX20.GridEX jgrdData 
         Height          =   1935
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3413
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         RowHeight       =   19
         MultiSelect     =   -1  'True
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "MS Sans Serif"
         FontName        =   "MS Sans Serif"
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "smpConfig.frx":0000
         FormatStyle(2)  =   "smpConfig.frx":0138
         FormatStyle(3)  =   "smpConfig.frx":01E8
         FormatStyle(4)  =   "smpConfig.frx":029C
         FormatStyle(5)  =   "smpConfig.frx":0374
         FormatStyle(6)  =   "smpConfig.frx":042C
         FormatStyle(7)  =   "smpConfig.frx":050C
         ImageCount      =   0
         PrinterProperties=   "smpConfig.frx":052C
      End
      Begin VB.Label lblSelectedColumn 
         Caption         =   "Selected Column:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdDeleteButton 
      Caption         =   "Delete"
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
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBackButton 
      Caption         =   "<<Back"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdNextButton 
      Caption         =   "Next>>"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdCloseButton 
      Caption         =   "Close"
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
      Left            =   7200
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveButton 
      Caption         =   "Save"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuLoadSample 
         Caption         =   "&Load Existing Sample Configuration"
      End
   End
End
Attribute VB_Name = "frmSmpConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This form is called from frmProdPlan and is used to view and modify sample configurations.</summary>
'</comment>

Option Explicit

Public dirtyFlag As String             'flag to signal dirty form
Public notes As String
Public oIRQInfo As CIRQInfo

Dim jobShippingId As Long
Dim comboSmp As String         'temp to test for change in sampletype combo
Dim description As String      'temp to test for change in the Description
Private mvarIRQsExist As Boolean
Private mvarCombinedOrRun As Boolean
Private intCodingColCount As Integer


Private Sub Form_Load()
    CenterForm Me
    Dim colTemp As JSColumn
    Dim colCount As Integer
    Dim i As Long
    Dim j As Long
    Dim arrHeader() As String
    Dim strCodingFileHeaders As String

    Call LoadSmpTypeCombo("SPTY", SSDBComboSmpType)
    Call LoadShipToCombo2Column(SSDBComboShip)
    
    jgrdData.Columns.Clear
    
    mvarIRQsExist = False
    mvarCombinedOrRun = False
    intCodingColCount = 0
    strCodingFileHeaders = ProductionRun.Coding_File_Header
    
    If BarcodeInfo.count > 0 Then
        For i = 1 To BarcodeInfo.count
            strCodingFileHeaders = strCodingFileHeaders & gRandDelimiter & BarcodeInfo.Item(i).BarcodeFValue & " " & BarcodeInfo.Item(i).BarcodeDesc
        Next i
    End If
    
    ' Extract headers into array
    arrHeader = Split(strCodingFileHeaders, gRandDelimiter)

    'initializes the grid with 30 columns
    For i = 1 To 30
        Set colTemp = jgrdData.Columns.Add()
        If i = 1 Then
            colTemp.Caption = "Type"
        Else
            If (i - 1) <= UBound(arrHeader) Then
                colTemp.Caption = arrHeader(i - 1)
            End If
        End If
    Next
    
    'sets first column to be not editable
    jgrdData.Columns(1).Selectable = False
    
    'reads the production run file from the Production Plan Form
    Call getFileLinksInfo
    
    If SampleDataExists(Me.txtSmpTypeNum) Then
        Call SetSSDBComboText(SSDBComboSmpType, "", SSDBComboSmpType.text)
        Call Read_SampleFile(gSampleFileName, SSDBComboSmpType.Columns.Item(2).text)
        'counts the number of columns
        colCount = CountDelimitedWords(vdata, gRandDelimiter)
        ' account for barcode column
        intCodingColCount = colCount
        If BarcodeInfo.count > 0 Then
            colCount = colCount + BarcodeInfo.count
        End If
        If Not booReplacement Then
            Read_File (ProductionRun.File_Name)
        End If
        dirtyFlag = ""              'sets the dirty flag to "" for existing data
        Call LoadShippingInfo(jobShippingId, txtShip, txtAttn, txtAdd1, txtAdd2, _
                        txtCity, txtState, txtZip, txtAdd3)
        cmdDeleteButton.enabled = True
    Else
        'Clintrak sample data does not exist
        Read_File (ProductionRun.File_Name)
        'populates the collection with data from file
        Set mData = New CCOLPDRFILES
        gSampleTypeId = 0
        Call SetSSDBComboText(Me.SSDBComboSmpType, "CLINTRAK", "CLINTRAK")
        Call SSDBComboSmpType_Click
        Me.txtQtynumber.text = 2
        Call Update_Click
        jobShippingId = 0
        cmdDeleteButton.enabled = False
        dirtyFlag = "Y"
        'counts the number of columns
        colCount = CountDelimitedWords(vdata, gRandDelimiter)
         ' account for barcode column
        intCodingColCount = colCount
        If BarcodeInfo.count > 0 Then
            colCount = colCount + BarcodeInfo.count
        End If
        frmProdPlan.sampleTypes = 1     'no data exists initialize to 1
    End If
    
    'hides columns that do not have data in it.
    For j = colCount + 1 To 30
        jgrdData.Columns.Item(j).Visible = False
    Next j
    
    'greys out the back button at initial form launch
    cmdBackButton.enabled = False
    
    If getExistingSampleTypes(0) < 1 Then
        mnuLoadSample.enabled = False
    Else
        mnuLoadSample.enabled = True
    End If
    
    Call CheckAvailableNext
    Call SetScreenEdit

    If Not frmProdPlan.oCollIRQInfo Is Nothing Then
        For Each oIRQInfo In frmProdPlan.oCollIRQInfo
            If (IIf(InStr(1, frmProdPlan.txtStockIRQ, CStr(oIRQInfo.IRQ_Number)) > 0, True, False) Or oIRQInfo.IRQ_Number = frmProdPlan.txtScratchIRQ) Then
                    mvarIRQsExist = True
            End If
        Next
    End If
    
    ' only enable if not a replacement and not clintrak samples
    If booReplacement Or Me.SSDBComboSmpType.text = "CLINTRAK" Or mvarIRQsExist = True Then
        Me.mnuLoadSample.enabled = False
        Me.cmdDeleteButton.enabled = False
    End If
    
    If booReplacement Or mvarIRQsExist = True Then
        Update.enabled = False
        txtQtynumber.enabled = False
        Me.cmdLoadCoding.enabled = False
    Else
        Update.enabled = True
        txtQtynumber.enabled = True
        If Me.SSDBComboSmpType.text <> "CLINTRAK" Then Me.cmdLoadCoding.enabled = True
    End If

    If Determine_If_PDR_HasRun = True Or Planning.CheckIfCombined(ProductionRun.Barcode_Id) = True Then
        mvarCombinedOrRun = True
        Call LockOutForm
    End If
    
    jgrdData.ItemCount = mData.count
    jgrdData.Update

End Sub

Private Sub cmdBackButton_Click()
    'checks to see whether the form was modified
    Call CheckChanges
    
    'save was canceled
    If dirtyFlag = "C" Then
        dirtyFlag = "Y"
        Exit Sub
    'was not saved
    ElseIf dirtyFlag = "Y" Then
        'adjust the total samples existing for unsaved data
        If Not SampleDataExists(Me.txtSmpTypeNum) Then
            frmProdPlan.sampleTypes = frmProdPlan.sampleTypes - 1
        End If
    End If
    
    'decrements the Type Number and retrieves the sample data
    txtSmpTypeNum = txtSmpTypeNum - 1
    
    If SampleDataExists(Me.txtSmpTypeNum) Then
        Call SetSSDBComboText(SSDBComboSmpType, "", SSDBComboSmpType.text)
        Call Read_SampleFile(gSampleFileName, _
             SSDBComboSmpType.Columns.Item(2).text)
        Read_File (gCodingFileName)
        dirtyFlag = ""
        Call SetScreenEdit
        If SSDBComboSmpType.text <> "CLINTRAK" Then
            Call LoadShippingInfo(jobShippingId, txtShip, txtAttn, txtAdd1, txtAdd2, _
                        txtCity, txtState, txtZip, txtAdd3)
            If jobShippingId = 0 Then ClearShipFields
        End If
    
        cmdDeleteButton.enabled = True
    Else
        Read_File (gCodingFileName)
        Set mData = New CCOLPDRFILES
        dirtyFlag = "Y"
        gSampleTypeId = 0
        SSDBComboSmpType.text = ""
        SSDBComboShip.text = ""
        comboSmp = ""
        Call ClearShipFields
        txtDescription.text = ""
        description = ""
        jobShippingId = 0
        Call SetScreenEdit
        Call mData.Add(vdata, 1, _
          SSDBComboSmpType.Columns.Item(2).text)
        cmdDeleteButton.enabled = False
    End If
    
    ' Set to dirty if Print @ Packager has changed
    If gSampleTypeId > 0 And Me.SSDBComboShip.enabled And Trim(Me.SSDBComboShip.text) = "" Then
        dirtyFlag = "Y"
    End If
    
    'greys out back button when first sample type is reached
    If txtSmpTypeNum = 1 Then
        cmdBackButton.enabled = False
    End If
    
    ' only enable if not a replacement and not clintrak samples
    If booReplacement Or Me.SSDBComboSmpType.text = "CLINTRAK" Or mvarIRQsExist = True Then
        Me.mnuLoadSample.enabled = False
        cmdDeleteButton.enabled = False
    Else
        Me.mnuLoadSample.enabled = True
    End If
    
    Call CheckAvailableNext
    
    If mvarCombinedOrRun = True Then
        Call LockOutForm
    End If
    
    'resets the configure column data
    txtcolumn.text = ""
    columnNumber = 0
    
    jgrdData.ItemCount = mData.count
    jgrdData.Update
    jgrdData.Refresh
    ' Need to deselect for processing on the Configure screen
    jgrdData.Row = 1
    jgrdData.RowSelected(1) = False
    
    ' DW 2010-002 added to fix the lack of scroll-return
    jgrdData.EnsureVisible 1, 1
    
End Sub

Private Sub cmdCloseButton_Click()
    'check see whether the data was modified
    Call CheckChanges
    
    'save was cancelled
    If dirtyFlag = "C" Then
        dirtyFlag = "Y"
        Exit Sub
    'was not saved
    ElseIf dirtyFlag = "Y" Then
        Me.txtQtynumber = mData.count
        columnNumber = 0    'resets the selected column when the form is closed
        Unload Me
        Exit Sub
    End If

    columnNumber = 0    'resets the selected column when the form is closed
    Unload Me
    
End Sub

Private Sub cmdConfigure_Click()
    If columnNumber = 0 Then
        MsgBox "Please Select a Column first.", vbExclamation
        Exit Sub
    ElseIf columnNumber = 1 Then
        MsgBox "The First Column cannot be edited.", vbExclamation
        Exit Sub
    End If
    frmGridConfig.Show vbModal
End Sub

Private Sub cmdDeleteButton_Click()
    On Error GoTo Handle_Error
        
    If MsgBox("Are you sure you want to delete this sample?", _
        vbQuestion + vbYesNo) = vbYes Then
        'remove the data from the existing quantity and sample types
        frmProdPlan.sampleTypes = frmProdPlan.sampleTypes - 1
        Call DeleteSample(CLng(ProductionRun.Production_Run_Id), Me.txtSmpTypeNum)
    Else
        Exit Sub
    End If
        
    'decrements the type number
    txtSmpTypeNum = txtSmpTypeNum - 1
    
    'checks to see if the type number goes to zero
    If CInt(txtSmpTypeNum.text) < 1 Then
        txtSmpTypeNum.text = 1
        cmdBackButton.enabled = False
    End If
    
    
    Call CheckAvailableNext
    
    If SampleDataExists(Me.txtSmpTypeNum) Then
        Call SetSSDBComboText(SSDBComboSmpType, "", SSDBComboSmpType.text)
        Call Read_SampleFile(gSampleFileName, _
             SSDBComboSmpType.Columns.Item(2).text)
        Read_File (gCodingFileName)
        dirtyFlag = ""
        Call LoadShippingInfo(jobShippingId, txtShip, txtAttn, txtAdd1, txtAdd2, _
                        txtCity, txtState, txtZip, txtAdd3)
        cmdDeleteButton.enabled = True
    Else
        Read_File (gCodingFileName)
        'populates the collection with data from file
        Set mData = New CCOLPDRFILES
        Call mData.Add(vdata, 1, SSDBComboSmpType.text)
        gSampleTypeId = 0
        Call ClearShipFields
        SSDBComboSmpType.text = ""
        SSDBComboShip.text = ""
        Me.cmdDeleteButton.enabled = False
    End If
    
    columnNumber = 0
    txtcolumn.text = ""
    
    ' only enable if not a replacement and not clintrak samples
    If booReplacement Or Me.SSDBComboSmpType.text = "CLINTRAK" Or mvarIRQsExist = True Then
        Me.mnuLoadSample.enabled = False
        Me.cmdDeleteButton.enabled = False
    Else
        Me.mnuLoadSample.enabled = True
    End If
    
    If txtSmpTypeNum = 1 Then
        cmdBackButton.enabled = False
    End If
    
    If getExistingSampleTypes(0) < 1 Then
        mnuLoadSample.enabled = False
        cmdDeleteButton.enabled = False
    End If
    
    Call CheckAvailableNext
    
    frmProdPlan.booSamplesDirtyFlag = True
    frmProdPlan.txtProducedBy = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
    
    'md added calls to correct the sample file(s) when the delete is done
    'we are not aligning the sample file name with the sample type number
    Call AlignSmplFileName_After_Delete
    
    MsgBox "Sample Configuration has been deleted.", vbInformation
    
    jgrdData.ItemCount = mData.count
    jgrdData.Update
    jgrdData.Refresh
    
Cleanup_Exit:
     Exit Sub
Handle_Error:
     MsgBox Err.description & vbCrLf & _
         "in frmSmpConfig.cmdDeleteButton_Click ", _
         vbCritical + vbOKOnly, "Application Error"
     Resume Cleanup_Exit
End Sub

Private Sub cmdLoadCoding_Click()
    'Updates the grid with the quantity selected
    Dim temp As Long
    Dim i As Long
    Dim tempstr As String
    
     'checks to see whether the sample type was filled in first
    If SSDBComboSmpType.text = "" Then
        MsgBox "The Sample Type must be entered!", vbExclamation
        txtQtynumber = 1
        Exit Sub
    End If
    
    'checks to see whether the quantity number is valid
    If Not IsNumeric(txtQtynumber) Then
        MsgBox "The Quantity Value must be numeric!", vbExclamation
        Exit Sub
    End If
    
    ' Checks to see if there are enough product labels to copy.
    If CLng(Me.txtQtynumber.text) > ProductionRun.Qty_Requested Then
        MsgBox "Cannot load ""As Per Sequence"" since there are more sample labels than product.", vbExclamation
        Exit Sub
    End If
       
    temp = txtQtynumber
    
    If SSDBComboSmpType.text <> "CLINTRAK" Then
        'removes all the old data on the grid
        For i = 1 To mData.count
            mData.Remove (mData.count)
        Next
       
        tempstr = SSDBComboSmpType.Columns.Item(2).text
        Call ReadProcess_File(ProductionRun.File_Name, temp, tempstr)
            
        dirtyFlag = "Y"
    End If
    
    Me.txtQtynumber = mData.count
    
    jgrdData.ItemCount = mData.count
    jgrdData.Update
    jgrdData.Refresh
End Sub

Private Sub cmdNextButton_Click()
    'checks to see whether the data was changed
    Call CheckChanges
    
    'data is not saved
    If dirtyFlag = "Y" Then
        Call UnSavedNext
        Exit Sub
    'save was cancelled
    ElseIf dirtyFlag = "C" Then
        dirtyFlag = "Y"
        Exit Sub
    End If

    Me.txtSmpTypeNum = Me.txtSmpTypeNum + 1
    
    If SampleDataExists(Me.txtSmpTypeNum) Then
        Call SetSSDBComboText(SSDBComboSmpType, "", SSDBComboSmpType.text)
        Call Read_SampleFile(gSampleFileName, _
             SSDBComboSmpType.Columns.Item(2).text)
        Read_File (gCodingFileName)
        dirtyFlag = ""
        
        Call SetScreenEdit
        
        If SSDBComboSmpType.text <> "CLINTRAK" Then
            Call LoadShippingInfo(jobShippingId, txtShip, txtAttn, txtAdd1, txtAdd2, _
                        txtCity, txtState, txtZip, txtAdd3)
            If jobShippingId = 0 Then ClearShipFields   ' DW 2010-002 added
        End If
        cmdDeleteButton.enabled = True
    Else
        Read_File (gCodingFileName)
        Set mData = New CCOLPDRFILES
        dirtyFlag = "Y"
        gSampleTypeId = 0
        SSDBComboSmpType.text = ""
        SSDBComboShip.text = ""
        comboSmp = ""
        Call ClearShipFields
        txtDescription.text = ""
        description = ""
        jobShippingId = 0
        cmdDeleteButton.enabled = False
        Call SetScreenEdit
        Call mData.Add(vdata, 1, SSDBComboSmpType.text)
        frmProdPlan.sampleTypes = frmProdPlan.sampleTypes + 1
    End If
    
    'resets configure column data
    txtcolumn.text = ""
    columnNumber = 0
    txtQtynumber = mData.count
    
    ' Set to dirty if Print @ Packager has changed
    If gSampleTypeId > 0 And Me.SSDBComboShip.enabled And Trim(Me.SSDBComboShip.text) = "" Then
        dirtyFlag = "Y"
    End If
    
    'if sample number is 1 then back button is greyed out
    If txtSmpTypeNum = 1 Then
        cmdBackButton.enabled = False
    Else
        cmdBackButton.enabled = True
    End If
    
    ' only enable if not a replacement and not clintrak samples
    If booReplacement Or Me.SSDBComboSmpType.text = "CLINTRAK" Or mvarIRQsExist = True Then
        Me.mnuLoadSample.enabled = False
        ' the delete button is enabled/disable above as well, so only disable if certain criteria met
        cmdDeleteButton.enabled = False
    Else
        Me.mnuLoadSample.enabled = True
    End If

    
    Call CheckAvailableNext
    
    If mvarCombinedOrRun = True Then
        Call LockOutForm
    End If
    
    jgrdData.ItemCount = mData.count
    jgrdData.Update
    jgrdData.Refresh
    ' Need to deselect for processing on the Configure screen
    jgrdData.Row = 1
    jgrdData.RowSelected(1) = False
    
    ' DW 2010-002 added to fix the lack of scroll-return
    jgrdData.EnsureVisible 1, 1
    
End Sub

Private Sub cmdNotes_Click()
    Load frmNotes
    frmNotes.Show vbModal
End Sub

Private Sub cmdSaveButton_Click()
    Dim colCount As Long
    Dim j As Long
    
    jgrdData.ItemCount = mData.count
    jgrdData.Update
    jgrdData.Refresh
    
    'checks if the form is valid
    If Valid_Sample_Form() Then
        If SSDBComboSmpType.text = "CLINTRAK" Then
            notes = "This is an INTERNAL SAMPLE - DO NOT SHIP!!!"
            jobShippingId = 0
        Else
            If notes = "This is a CLINTRAK SAMPLE - DO NOT SHIP!!!" Or _
                    notes = "This is an INTERNAL SAMPLE - DO NOT SHIP!!!" Then
                notes = " "
            End If
        End If
                
        Call CollectionToFile(gSampleFileName)
        Call Save_Sample_Info
        SampleDataExists (Me.txtSmpTypeNum)
    Else
        dirtyFlag = "C"
        Exit Sub
    End If
    
    ' only enable if not a replacement and not clintrak samples
    If booReplacement Or Me.SSDBComboSmpType.text = "CLINTRAK" Or mvarIRQsExist = True Then
        Me.mnuLoadSample.enabled = False
        Me.cmdDeleteButton.enabled = False
    Else
        Me.mnuLoadSample.enabled = True
        Me.cmdDeleteButton.enabled = True
    End If
    
    frmProdPlan.booSamplesDirtyFlag = True
    frmProdPlan.txtProducedBy = gApplicationUser.LastName & ", " & gApplicationUser.FirstName
    
    MsgBox "Sample Configuration has been saved", vbInformation
    
    CheckAvailableNext
       
    colCount = CountDelimitedWords(vdata, gRandDelimiter)
    ' account for barcode column
    If BarcodeInfo.count > 0 Then
        colCount = colCount + BarcodeInfo.count
    End If
    
    For j = colCount + 1 To 30
        jgrdData.Columns.Item(j).Visible = False
    Next j
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmProdPlan.txtSampleGroups = getExistingSampleTypes(0)
    frmProdPlan.txtSamples = getExistingSampleQTY(0)
    
    ProductionRun.Sample_Number = frmProdPlan.txtSampleGroups
    ProductionRun.Samples_Requested = frmProdPlan.txtSamples
    
'    ProductionRun.UpdateSampleQuantities
End Sub

Private Sub jgrdData_ColumnHeaderClick(ByVal column As GridEX20.JSColumn)

    'current column that we have selected
    columnNumber = column.index
    txtcolumn.text = columnNumber
    
End Sub

Private Sub jgrdData_GroupByBoxHeaderClick(ByVal Group As GridEX20.JSGroup)

    'When clicking in a group by box header we change SortOrder for that group
    
    Group.SortOrder = -Group.SortOrder
    
End Sub

Private Sub jgrdData_LostFocus()
    jgrdData.Update
    jgrdData.Refresh
End Sub

Private Sub jgrdData_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)

    'Delete the item in the collection
    mData.Remove RowIndex
    
End Sub


Private Sub jgrdData_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim recTemp As CPRDFiles
    Dim Coding_File_Data() As String
    Dim selected_columns() As String
    Dim strData As String
    Dim strBarcode As String
    Dim i As Long

    'don't worry if the grid is sorted or if the column positions are changed by the user
    Set recTemp = mData.Item(RowIndex)
    With recTemp
        Values(1) = IIf(.Field1 = "", BLANK_RPT, .Field1)
        Values(2) = IIf(.Field2 = "", BLANK_RPT, .Field2)
        Values(3) = IIf(.Field3 = "", BLANK_RPT, .Field3)
        Values(4) = IIf(.Field4 = "", BLANK_RPT, .Field4)
        Values(5) = IIf(.Field5 = "", BLANK_RPT, .Field5)
        Values(6) = IIf(.Field6 = "", BLANK_RPT, .Field6)
        Values(7) = IIf(.Field7 = "", BLANK_RPT, .Field7)
        Values(8) = IIf(.Field8 = "", BLANK_RPT, .Field8)
        Values(9) = IIf(.Field9 = "", BLANK_RPT, .Field9)
        Values(10) = IIf(.Field10 = "", BLANK_RPT, .Field10)
        Values(11) = IIf(.Field11 = "", BLANK_RPT, .Field11)
        Values(12) = IIf(.Field12 = "", BLANK_RPT, .Field12)
        Values(13) = IIf(.Field13 = "", BLANK_RPT, .Field13)
        Values(14) = IIf(.Field14 = "", BLANK_RPT, .Field14)
        Values(15) = IIf(.Field15 = "", BLANK_RPT, .Field15)
        Values(16) = IIf(.Field16 = "", BLANK_RPT, .Field16)
        Values(17) = IIf(.Field17 = "", BLANK_RPT, .Field17)
        Values(18) = IIf(.Field18 = "", BLANK_RPT, .Field18)
        Values(19) = IIf(.Field19 = "", BLANK_RPT, .Field19)
        Values(20) = IIf(.Field20 = "", BLANK_RPT, .Field20)
        Values(21) = IIf(.Field21 = "", BLANK_RPT, .Field21)
        Values(22) = IIf(.Field22 = "", BLANK_RPT, .Field22)
        Values(23) = IIf(.Field23 = "", BLANK_RPT, .Field23)
        Values(24) = IIf(.Field24 = "", BLANK_RPT, .Field24)
        Values(25) = IIf(.Field25 = "", BLANK_RPT, .Field25)
        Values(26) = IIf(.Field26 = "", BLANK_RPT, .Field26)
        Values(27) = IIf(.Field27 = "", BLANK_RPT, .Field27)
        Values(28) = IIf(.Field28 = "", BLANK_RPT, .Field28)
        Values(29) = IIf(.Field29 = "", BLANK_RPT, .Field29)
        Values(30) = IIf(.Field30 = "", BLANK_RPT, .Field30)
    End With
    

    If BarcodeInfo.count > 0 Then

        For i = 1 To intCodingColCount
            If i < intCodingColCount Then
                strData = strData & Values(i) & gRandDelimiter
            Else
                strData = strData & Values(i)
            End If
        Next i
        
        'separate into an array the data from the grid
        Coding_File_Data = Split(strData, gRandDelimiter)
        
        For i = 1 To BarcodeInfo.count
            strBarcode = ""
            
            ReDim selected_columns(0)
        
            'place into an array the selected column numbers
            Call Extract_MergeBarcode_Columns(BarcodeInfo.Item(i).BarcodeFields, selected_columns, True)
            
            'build the merge barcode
            strBarcode = Display_Barcode(selected_columns, Coding_File_Data)
            
            Values(intCodingColCount + i) = strBarcode
        Next i
    End If

End Sub


Private Sub jgrdData_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim recTemp As CPRDFiles
    
    Set recTemp = mData.Item(RowIndex)
    With recTemp
        .Field1 = IIf(Trim(UCase(Values(1))) = BLANK_RPT, "", Values(1))
        .Field2 = IIf(Trim(UCase(Values(2))) = BLANK_RPT, "", Values(2))
        .Field3 = IIf(Trim(UCase(Values(3))) = BLANK_RPT, "", Values(3))
        .Field4 = IIf(Trim(UCase(Values(4))) = BLANK_RPT, "", Values(4))
        .Field5 = IIf(Trim(UCase(Values(5))) = BLANK_RPT, "", Values(5))
        .Field6 = IIf(Trim(UCase(Values(6))) = BLANK_RPT, "", Values(6))
        .Field7 = IIf(Trim(UCase(Values(7))) = BLANK_RPT, "", Values(7))
        .Field8 = IIf(Trim(UCase(Values(8))) = BLANK_RPT, "", Values(8))
        .Field9 = IIf(Trim(UCase(Values(9))) = BLANK_RPT, "", Values(9))
        .Field10 = IIf(Trim(UCase(Values(10))) = BLANK_RPT, "", Values(10))
        .Field11 = IIf(Trim(UCase(Values(11))) = BLANK_RPT, "", Values(11))
        .Field12 = IIf(Trim(UCase(Values(12))) = BLANK_RPT, "", Values(12))
        .Field13 = IIf(Trim(UCase(Values(13))) = BLANK_RPT, "", Values(13))
        .Field14 = IIf(Trim(UCase(Values(14))) = BLANK_RPT, "", Values(14))
        .Field15 = IIf(Trim(UCase(Values(15))) = BLANK_RPT, "", Values(15))
        .Field16 = IIf(Trim(UCase(Values(16))) = BLANK_RPT, "", Values(16))
        .Field17 = IIf(Trim(UCase(Values(17))) = BLANK_RPT, "", Values(17))
        .Field18 = IIf(Trim(UCase(Values(18))) = BLANK_RPT, "", Values(18))
        .Field19 = IIf(Trim(UCase(Values(19))) = BLANK_RPT, "", Values(19))
        .Field20 = IIf(Trim(UCase(Values(20))) = BLANK_RPT, "", Values(20))
        .Field21 = IIf(Trim(UCase(Values(21))) = BLANK_RPT, "", Values(21))
        .Field22 = IIf(Trim(UCase(Values(22))) = BLANK_RPT, "", Values(22))
        .Field23 = IIf(Trim(UCase(Values(23))) = BLANK_RPT, "", Values(23))
        .Field24 = IIf(Trim(UCase(Values(24))) = BLANK_RPT, "", Values(24))
        .Field25 = IIf(Trim(UCase(Values(25))) = BLANK_RPT, "", Values(25))
        .Field26 = IIf(Trim(UCase(Values(26))) = BLANK_RPT, "", Values(26))
        .Field27 = IIf(Trim(UCase(Values(27))) = BLANK_RPT, "", Values(27))
        .Field28 = IIf(Trim(UCase(Values(28))) = BLANK_RPT, "", Values(28))
        .Field29 = IIf(Trim(UCase(Values(29))) = BLANK_RPT, "", Values(29))
        .Field30 = IIf(Trim(UCase(Values(30))) = BLANK_RPT, "", Values(30))
    End With
    
    dirtyFlag = "Y"
    
End Sub

Private Sub mnuLoadSample_Click()
    On Error GoTo Handle_Error
        
    Dim strInput As String
    Dim quantity As Long
    Dim file As String
    Dim stype As String

    If MsgBox( _
        "Are you sure you want to load existing data?" & vbCrLf & _
        "Any changes made to this form will be lost!", vbQuestion + vbYesNo) = vbYes Then
        
        Do
            strInput = InputBox( _
                "Enter Sample Type Number or Cancel to Quit!", _
                "Load Existing Sample", _
                1)
            If Len(strInput) = 0 Then Exit Sub
        Loop Until CheckSampleNum(CInt(strInput), file, quantity, _
            "Sample Type Number Does Not Exists. Please Try Again.", stype)
      
        ' update the descrition too
        'if the description value was the default one change to the newly selected type
        If Trim(txtDescription.text) = Trim(SSDBComboSmpType.text) Then
            txtDescription.text = stype
        End If

        SSDBComboSmpType.text = stype
        Call SetSSDBComboText(SSDBComboSmpType, "", SSDBComboSmpType.text)
        comboSmp = Me.SSDBComboSmpType.text
          
        Call Read_SampleFile(file, _
            SSDBComboSmpType.Columns.Item(2).text)
        Call Read_File(ProductionRun.File_Name)
        
        dirtyFlag = "Y"
        txtQtynumber.text = mData.count
        jgrdData.ItemCount = mData.count
        jgrdData.Update
        jgrdData.Refresh
    
    End If

Cleanup_Exit:
     Exit Sub
Handle_Error:
     MsgBox Err.description & vbCrLf & _
         "in frmSmpConfig.mnuLoadSample_Click ", _
         vbCritical + vbOKOnly, "Application Error"
     Resume Cleanup_Exit
End Sub

Private Sub SSDBComboShip_Click()
    
    If jobShippingId <> SSDBComboShip.Columns.Item(1).text Then
        dirtyFlag = "Y"
        jobShippingId = SSDBComboShip.Columns.Item(1).text
    End If
    
    Call LoadShippingInfo(jobShippingId, txtShip, txtAttn, txtAdd1, txtAdd2, _
                        txtCity, txtState, txtZip, txtAdd3)
                        
End Sub

Private Sub SSDBComboShip_InitColumnProps()

    SSDBComboShip.Columns(0).Width = SSDBComboShip.Width
    SSDBComboShip.Columns(1).Visible = False
    SSDBComboShip.Columns(2).Width = SSDBComboShip.Width
    
End Sub

Private Sub SSDBComboSmpType_InitColumnProps()

    SSDBComboSmpType.Columns(0).Width = SSDBComboSmpType.Width
    SSDBComboSmpType.Columns(1).Visible = False
    SSDBComboSmpType.Columns(2).Visible = False
    
End Sub

Private Sub SSDBComboSmpType_Click()
    
    On Error GoTo Handle_Error
    
    'checks to see if data was changed
    If comboSmp <> SSDBComboSmpType.Columns.Item(0).text Then
        ' preventing the manual selection of CLINTRAK sample type
        If SSDBComboSmpType.Columns.Item(0).text = "CLINTRAK" And Me.txtSmpTypeNum.text <> "1" Then
            Call SetSSDBComboText(Me.SSDBComboSmpType, "", comboSmp)
            MsgBox "CLINTRAK samples cannot be configured manually.", vbCritical, "Invalid Sample Type"
            Exit Sub
        ElseIf Me.SSDBComboSmpType.Columns.Item(0).text <> "CLINTRAK" And Me.txtSmpTypeNum.text = "1" Then
            Call SetSSDBComboText(Me.SSDBComboSmpType, "", "CLINTRAK")
            MsgBox "The first sample set must be CLINTRAK samples.", vbCritical, "Invalid Sample Type"
        End If
           
        dirtyFlag = "Y"
        'if the description value was the default one change to the newly selected type
        If Trim(txtDescription) = Trim(comboSmp) Then
            txtDescription = SSDBComboSmpType.text
        End If
        
        Call UpdateFirstColumn
        
        comboSmp = Me.SSDBComboSmpType.Columns(0).text 'comboSmpType.List(comboSmpType.ListIndex)
    End If
    
    Call SetScreenEdit
  
    jgrdData.ItemCount = mData.count
    jgrdData.Update
    jgrdData.Refresh

Cleanup_Exit:
     Exit Sub
Handle_Error:
     MsgBox Err.description & vbCrLf & _
         "in frmSmpConfig.SSDBComboSmpType_Click ", _
         vbCritical + vbOKOnly, "Application Error"
     Resume Cleanup_Exit
End Sub


Private Sub txtDescription_LostFocus()
    'checks to see whether the description text was changed
    If Trim(description) <> Trim(txtDescription.text) Then
        dirtyFlag = "Y"
        description = Trim(txtDescription.text)
    End If
End Sub

Private Sub Update_Click()
    'Updates the grid with the quantity selected
    Dim temp As Long
    Dim i As Long
    Dim tempstr As String
        
     'checks to see whether the sample type was filled in first
    If SSDBComboSmpType.text = "" Then
        MsgBox "The Sample Type Must be Entered!!", vbExclamation
        txtQtynumber = 1
        Exit Sub
    End If
    
     'checks to see whether the quantity number is valid
    If Not IsNumeric(txtQtynumber) Then
        MsgBox "The Quantity Value must be Numeric!", vbExclamation
        Exit Sub
    End If
    
    temp = txtQtynumber
    
    'check to see if the clintrak samples requested are > than total quantity
    'if clintrak samples - must read the coding file and populate with real data
    'to the screen, not just default the values.
    If SSDBComboSmpType.text = "CLINTRAK" Then
         'check the sample quantity
          'load the grid with live data for Clintrak samples
        tempstr = SSDBComboSmpType.Columns.Item(2).text
        Call ReadProcess_File(ProductionRun.File_Name, temp, tempstr)
        dirtyFlag = "Y"
    End If
    'adds new rows to the table if not clintrak sample request
    'md added check for <> CLINTRAK
    If SSDBComboSmpType.text <> "CLINTRAK" Then
        ' DW 2010-002 added
        For i = 1 To mData.count
            mData.Remove (mData.count)
        Next
        ' Re-Read vdata record 1
        Read_File (ProductionRun.File_Name)
        
        If temp > mData.count Then
            For i = mData.count To (temp - 1)
                Call mData.Add(vdata, (i + 1), SSDBComboSmpType.Columns.Item(2).text)
            Next i
            dirtyFlag = "Y"
        End If
    End If
    
    'removes rows from the table
    If temp < mData.count Then
        For i = (temp + 1) To mData.count
            mData.Remove (mData.count)
        Next
    dirtyFlag = "Y"
    End If
    
    Me.txtQtynumber = mData.count

    jgrdData.ItemCount = mData.count
    jgrdData.Update
    jgrdData.Refresh
End Sub

Public Sub CollectionToFile(strDestination As String)
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
                strFileName = strFileName & Trim(ProductionRun.Barcode_Id) & "_" & CInt(Me.txtSmpTypeNum.text) & ".smp"
                gSampleFileName = strFileName
            Else
                strFileName = strDestination
                gSampleFileName = strDestination
            End If
            
            'calls the function to write the collection to a file
            'md added code to write to file with 3 char abbrev.
            Call WriteFile(mData, _
                SSDBComboSmpType.Columns.Item(2).text, _
               Trim(strFileName))
               
            dirtyFlag = ""
            
Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Saving sample configuration to file"
    Resume Exit_this_Sub
    
End Sub

Private Sub Save_Sample_Info()
'
'comments: this sub calls the stored procedure to save or update the sample info
'parameters: none
'returns: nothing
'
On Error GoTo Error_this_Sub

    Dim nreturn As Long
    Dim objData As nADOData.CADOData

    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly

        .AddParameter "Sample Type Id", gSampleTypeId, adInteger, adParamInput
        .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .AddParameter "Type Number", CInt(txtSmpTypeNum), adInteger, adParamInput
        .AddParameter "Sample Type", SSDBComboSmpType, adVarChar, adParamInput
        .AddParameter "Ship To Id", jobShippingId, adInteger, adParamInput
        .AddParameter "Quantity", mData.count, adInteger, adParamInput
        .AddParameter "Sample File Name", gSampleFileName, adVarChar, adParamInput
        .AddParameter "Sample Description", CheckNulls(txtDescription), adChar, adParamInput
        .AddParameter "Notes", CheckNulls(notes), adVarChar, adParamInput

        .AddParameter "return", "   ", adInteger, adParamOutput ' the "   " is for a length value
        .AddParameter "identity", "   ", adInteger, adParamOutput ' the "   " is for a length value

        .ExecuteSP "save_SampleFiles", True

        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        If IsNull(nreturn) Or Trim(nreturn) = "" Or nreturn <> 0 Then
            .Connection.RollbackTrans
            Exit Sub
        End If
        dirtyFlag = ""
    End With

Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Saving sample information"
    Resume Exit_this_Sub

End Sub

Private Function SampleDataExists(smpNum As Integer) As Boolean
'
'comments: Loads the form with existing data from the table Sample_Types that has
'           matching production run id and sample type number
'parameters: smpNum - sample type number to look up
'returns: boolean
'
On Error GoTo Error_this_Function
    Dim objData As nADOData.CADOData
    
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .AddParameter "Type Number", CInt(smpNum), adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"
        
        If Not .Recordset.EOF Then
            gSampleTypeId = .Recordset!sample_type_id
            gProductionRun_Id = .Recordset!Production_Run_Id
            txtSmpTypeNum = .Recordset!Type_Number
            SSDBComboSmpType.text = .Recordset!Sample_Type
            SSDBComboShip = .Recordset!description
            txtQtynumber = .Recordset!quantity
            gSampleFileName = .Recordset!Sample_File_Name
            txtDescription.text = Trim(.Recordset!Sample_Description)
            jobShippingId = .Recordset!Job_Shipping_Id
            comboSmp = SSDBComboSmpType.text
            description = txtDescription.text
            notes = .Recordset!notes
        Else
            txtQtynumber = "1"
            gSampleFileName = ""
            txtDescription = SSDBComboSmpType.text
            gSampleTypeId = 0
            notes = ""
        End If
        
        'checks to see if the a sample file exists for this particular production run
        'id and type number
        If FileExists(gSampleFileName) Then
            SampleDataExists = True
        Else
            SampleDataExists = False
        End If
        .Recordset.Close
    End With

Exit_this_Function:
    Exit Function

Error_this_Function:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Loading existing sample data information"
    Resume Exit_this_Function

End Function

'<comment>
' <summary>
'       Check that all necessary data is valid for sample config form.</summary>
' <return>True if all data fields necessary are populated.</return>
'</comment>
Public Function Valid_Sample_Form() As Boolean
    On Error GoTo Handle_Error

    Dim i As Long
    Dim columnSqueezeLength As Long
    Dim columnSqueezeCapacity As Long
    Dim columnSqueeze As String
    Dim tmpField As String
    Dim r As Long, c As Long
        
    Valid_Sample_Form = False
    columnSqueeze = ""
    columnSqueezeLength = 0
    columnSqueezeCapacity = 0
    
    For r = 1 To jgrdData.RowCount
        For c = 1 To jgrdData.Columns.count
            If InStr(1, mData.Item(r).Fields(c), gRandDelimiter) Then
                MsgBox "Row " & r & " column " & c & " contains the rand delimiter: """ & gRandDelimiter & """", vbExclamation
                Exit Function
            End If
        Next c
    Next r
    
    If mData.count = 0 Then
        Me.txtQtynumber.SetFocus
        MsgBox "Please Select Valid Quantity And Hit Update!", vbExclamation
        Exit Function
    End If
    
    If SSDBComboSmpType.text = "" Then
        Me.SSDBComboSmpType.SetFocus
        MsgBox "Please Select a Sample Type from Drop Down Box", vbExclamation
        Exit Function
    End If
    
    'md added code to bypass this check if TYPE = Clintrak
    If SSDBComboSmpType.text <> "CLINTRAK" And Not CBool(frmProdPlan.chkPrintAtPackager) Then
        If Trim(SSDBComboShip.text) = "" Then
            Me.SSDBComboShip.SetFocus
            MsgBox "Please Select a Ship To Location from Drop Down Box!", vbExclamation
            Exit Function
        End If
    End If
    
    ' DW 2010-002 added - when print at packager is checked Clintrak samples are not required
    If Not CBool(frmProdPlan.chkPrintAtPackager.value) Then
        ' check that the first set of samples is CLINTRAK and there are 2
        If Me.txtSmpTypeNum.text = "1" Then
            If Me.SSDBComboSmpType.text <> "CLINTRAK" Then
                MsgBox "The first sample set must be CLINTRAK samples.", vbExclamation
                Exit Function
            End If
            If Me.txtQtynumber.text <> "2" Then
                MsgBox "There must be 2 CLINTRAK samples.", vbExclamation
                Exit Function
            End If
            If mData.count <> 2 Then
                MsgBox "There must be 2 CLINTRAK samples.", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    columnSqueezeCapacity = 32
    columnSqueeze = Space$(columnSqueezeCapacity)
    
    For i = 1 To mData.count
            tmpField = mData.Item(i).Field1
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
            
            tmpField = mData.Item(i).Field2
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field3
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field4
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field5
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
            
            tmpField = mData.Item(i).Field6
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field7
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field8
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field9
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field10
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field11
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field12
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field13
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field14
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field15
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field16
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field17
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field18
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field19
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field20
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field21
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field22
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field23
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field24
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field25
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field26
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field27
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field28
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field29
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
                
            tmpField = mData.Item(i).Field30
            If Len(tmpField) + columnSqueezeLength >= columnSqueezeCapacity Then columnSqueezeCapacity = columnSqueezeCapacity * 2: columnSqueeze = columnSqueeze & Space$(columnSqueezeCapacity)
                Mid(columnSqueeze, columnSqueezeLength + 1, Len(tmpField)) = tmpField: columnSqueezeLength = columnSqueezeLength + Len(tmpField)
        
        columnSqueezeLength = 0
        
    Next i
    
    Valid_Sample_Form = True

Cleanup_Exit:
    Exit Function
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Function

Private Function getExistingSampleQTY(typenum As Integer) As Long
'
'comments: this function gets the original quantity of a sample type
'parameters: typeNum - sample type number to be checked
'returns: quantity of sample type as integer
'
On Error GoTo Error_this_Function

Dim total As Long    'temp holder for total of quantities
total = 0
Dim objData As nADOData.CADOData

    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .AddParameter "Type Number", typenum, adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                total = total + .Recordset!quantity
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
    
    getExistingSampleQTY = total

Exit_this_Function:
    Exit Function
    
Error_this_Function:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error getting existing sample quantity totals"
    Resume Exit_this_Function

End Function


'<comment>
' <summary>
'       Check to see whether the dirty flag has been raised.  If so, prompt
'       the user with options to save.</summary>
'</comment>
Private Sub CheckChanges()
    On Error GoTo Handle_Error

    '"Y" = its not saved yet
    '"C" = save aborted
    '"" = data saved

    'checks to see if the data has been changed
    If dirtyFlag = "Y" Then
        
        Select Case MsgBox( _
            "Changes have not been saved, Save Changes?" & vbCrLf & _
            "No: Changes will be lost!", vbQuestion + vbYesNoCancel)
            Case vbYes
                Call cmdSaveButton_Click
            Case vbNo
                dirtyFlag = "Y"
            Case vbCancel
                dirtyFlag = "C"
        End Select
    End If
    
Cleanup_Exit:
    Exit Sub
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Sub

Private Sub UnSavedNext()
'
'comments: this function is used by the command next button to evaulate the data for the
'           next sample type number. the function populates the next sample type form
'           with the appropraite data
'parameters: none
'returns: nothing
'
'checks to see if the next sample type exists
On Error GoTo Error_this_Sub

    If SampleDataExists(Me.txtSmpTypeNum + 1) Then
        ' this code is executed if the next sample exists in the DB and the file exists.
        ' it loads the screen with the saved info for the next sample set.
        Call Read_SampleFile(gSampleFileName, SSDBComboSmpType.text)
        Read_File (gCodingFileName)
        dirtyFlag = ""
        Call LoadShippingInfo(jobShippingId, txtShip, txtAttn, txtAdd1, txtAdd2, _
                            txtCity, txtState, txtZip, txtAdd3)
        cmdDeleteButton.enabled = True
        cmdBackButton.enabled = True
    'checks to see if the current sample type exists
    ElseIf SampleDataExists(Me.txtSmpTypeNum) Then
        ' this code is executed if the current sample exsists in the DB and the file exists
        ' it will prep the next sample type as a new one since the IF above failed b/c the
        ' next one doesn't exist yet.
        txtSmpTypeNum = txtSmpTypeNum + 1
        Read_File (gCodingFileName)
        Set mData = New CCOLPDRFILES
        Call mData.Add(vdata, 1, _
            SSDBComboSmpType.Columns.Item(2).text)
        frmProdPlan.sampleTypes = frmProdPlan.sampleTypes + 1
        dirtyFlag = "Y"
        SSDBComboSmpType.text = ""
        SSDBComboShip.text = ""
        comboSmp = ""
        txtDescription.text = ""
        Call ClearShipFields
        description = ""
        jobShippingId = 0
        txtQtynumber = mData.count
        cmdDeleteButton.enabled = False
        cmdBackButton.enabled = True
    Else
        ' this code is exeuted if the next sample type and the current sample set both
        ' don't exists.  It loads leaves the current sample set loaded, but refreshes
        ' all fields back to scratch.
    'current sample type does not exist
        Read_File (gCodingFileName)
        'populates the collection with data from file
        Set mData = New CCOLPDRFILES
         Call mData.Add(vdata, 1, _
             SSDBComboSmpType.Columns.Item(2).text)
        dirtyFlag = "Y"
        SSDBComboSmpType.text = ""
        SSDBComboShip.text = ""
        comboSmp = ""
        txtDescription.text = ""
        Call ClearShipFields
        description = ""
        jobShippingId = 0
        cmdDeleteButton.enabled = False
    End If

    txtcolumn.text = ""
    columnNumber = 0
    
    ' only enable if not a replacement and not clintrak samples
    If booReplacement Or Me.SSDBComboSmpType.text = "CLINTRAK" Or mvarIRQsExist = True Then
        Me.mnuLoadSample.enabled = False
        Me.cmdDeleteButton.enabled = False
    Else
        Me.mnuLoadSample.enabled = True
    End If
    
    If txtSmpTypeNum = 1 Then
        cmdBackButton.enabled = False
    Else
        cmdBackButton.enabled = True
    End If
    
    Call CheckAvailableNext

    jgrdData.ItemCount = mData.count
    jgrdData.Update
    jgrdData.Refresh

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Loading sample data for next sample number"
    Resume Exit_this_Sub

End Sub

Private Function CheckSampleNum(num As Integer, file As String, _
quantity As Long, Message, smpltype As String) As Boolean
'
'comments: this function checks whether there is sample data for a particular sample
'           number.
'parameters:    num - sample number to check,
'               file - variable that holds file if found
'               quantity - variable that holds the quantity if found
'returns:   True if data is found
'
Dim objData As nADOData.CADOData

On Error GoTo Error_this_Function

'md made changes for clintrak samples project

    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .AddParameter "Type Number", CInt(num), adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"

        If Not .Recordset.EOF Then
            file = .Recordset!Sample_File_Name
            quantity = .Recordset!quantity
            smpltype = .Recordset!Sample_Type
            'md added for clintrak samples
            If .Recordset!Job_Shipping_Id = 0 Then
                CheckSampleNum = False
                file = ""
                quantity = 0
                Message = "Cannot Load a CLINTRAK sample. Please Try Again!"
                GoTo Exit_this_Function
            End If
        End If
        .Recordset.Close
    End With

        If FileExists(file) Then
            CheckSampleNum = True
        Else
            CheckSampleNum = False
            file = ""
            quantity = 0
        End If

Exit_this_Function:
    Exit Function

Error_this_Function:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Checking sample number information"
    Resume Exit_this_Function

End Function

Private Sub UpdateFirstColumn()
Dim i As Long

    For i = 1 To mData.count
         mData.Item(i).Field1 = _
        SSDBComboSmpType.Columns.Item(2).text & "-" & i
    Next

End Sub



Private Function getExistingSampleTypes(typenum As Integer) As Long
'
'comments: this function gets the original quantity of a sample tpye
'parameters: typeNum - sample type number to be checked
'returns: quantity of sample type as integer
'
On Error GoTo Error_this_Function

Dim total As Long    'temp holder for total of quantities
total = 0
Dim objData As nADOData.CADOData

    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .AddParameter "Type Number", typenum, adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                total = total + 1
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
    
    getExistingSampleTypes = total

Exit_this_Function:
    Exit Function
    
Error_this_Function:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error getting existing sample quantity totals"
    Resume Exit_this_Function

End Function

Private Sub CheckAvailableNext()
'checks to see if the last sample has been reached and greys out next button

'    If Me.txtSmpTypeNum = frmProdPlan.txtSampleGroups Then
'        cmdNextButton.Enabled = False
'        cmdAdd.Enabled = True
'    Else
'        cmdNextButton.Enabled = True
'        cmdAdd.Enabled = False
'    End If
    

    If Me.txtSmpTypeNum = frmProdPlan.sampleTypes Then
        cmdNextButton.enabled = False
        If mvarIRQsExist = True Or booReplacement = True Then
            cmdAdd.enabled = False
        Else
            cmdAdd.enabled = True
        End If
        
    Else
        cmdNextButton.enabled = True
        cmdAdd.enabled = False
    End If
''     If CLng(Me.txtSmpTypeNum.Text) >= CLng(frmProdPlan.txtSampleGroups.Text) Then
''        cmdNextButton.Caption = "Add New Type"
''    Else
''        cmdNextButton.Caption = "Next>>"
''    End If

End Sub

Private Sub SetScreenEdit()
    On Error GoTo Handle_Error

    Dim i As Long
    Dim booEditable As Boolean

    booEditable = (SSDBComboSmpType.text <> "CLINTRAK")
    If Not booEditable Then ClearShipFields
    
    Me.lblSelectedColumn.enabled = booEditable
    Me.txtcolumn.enabled = booEditable
    Me.cmdConfigure.enabled = booEditable
    Me.cmdLoadCoding.enabled = booEditable
    
    Me.SSDBComboShip.enabled = (booEditable And Not CBool(frmProdPlan.chkPrintAtPackager.value))
    Me.lblShipToCbo.enabled = SSDBComboShip.enabled
    Me.lblShipTo.enabled = SSDBComboShip.enabled
    Me.lblAttn.enabled = SSDBComboShip.enabled
    Me.lblAddr1.enabled = SSDBComboShip.enabled
    Me.lblAddr2.enabled = SSDBComboShip.enabled
    Me.lblAddr3.enabled = SSDBComboShip.enabled
    Me.lblCity.enabled = SSDBComboShip.enabled
    Me.lblState.enabled = SSDBComboShip.enabled
    Me.lblZip.enabled = SSDBComboShip.enabled
    Me.txtShip.enabled = SSDBComboShip.enabled
    Me.txtAttn.enabled = SSDBComboShip.enabled
    Me.txtAdd1.enabled = SSDBComboShip.enabled
    Me.txtAdd2.enabled = SSDBComboShip.enabled
    Me.txtAdd3.enabled = SSDBComboShip.enabled
    Me.txtCity.enabled = SSDBComboShip.enabled
    Me.txtState.enabled = SSDBComboShip.enabled
    Me.txtZip.enabled = SSDBComboShip.enabled
    
    SSDBComboSmpType.enabled = booEditable
    ' Do not allow the user to edit Type column
    Me.jgrdData.Columns(1).Selectable = False
    For i = 2 To 20
        Me.jgrdData.Columns(i).Selectable = SSDBComboSmpType.enabled
        ' so barcode columns are not selectable
        If i > intCodingColCount Then
            Me.jgrdData.Columns(i).Selectable = False
        End If
    Next
     
Cleanup_Exit:
    Exit Sub
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Sub

Private Sub ClearShipFields()
    On Error GoTo Handle_Error
    Me.SSDBComboShip.text = ""
    Me.txtShip.text = ""
    Me.txtAttn.text = ""
    Me.txtAdd1.text = ""
    Me.txtAdd2.text = ""
    Me.txtAdd3.text = ""
    Me.txtCity.text = ""
    Me.txtState.text = ""
    Me.txtZip.text = ""
Cleanup_Exit:
    Exit Sub
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Sub


Private Sub Process_Rename_Files(TmpSmpFile As String, TmpTypeNum As Integer)

'md new for clintrak samples project
'This sub will rename the sample file to match the file type number after a delete. The
'sample type table will be update as well.

Dim StringCnt As String
Dim stringfl As String
Dim strFileName As String

    
    StringCnt = GetDelimitedFirstLine(TmpSmpFile, 1, ".", False)
    'StringCnt = GetDelimitedFirstLine(StringCnt, 2, "_", False) ' DW 2010-002 commented out for modification
    ' Explanation: Path changed to UNC, introduced second underscore, desired value is now the third element
    StringCnt = GetDelimitedFirstLine(StringCnt, 3, "_", False)
    
    'if the type and the file match, then this file does not have to be renamed
    If CInt(StringCnt) = TmpTypeNum Then
        Exit Sub
    End If

    'create a temporary sample file
    strFileName = ""
    strFileName = GetFilePath(TmpSmpFile)
    strFileName = strFileName & "\temporary.smp"
    Call FileCopy(TmpSmpFile, strFileName)
    'deletes the old sample file from the directory
    Kill (TmpSmpFile)
    
    'create a copy of the temporary file to the correct name
    stringfl = ""
    stringfl = GetFilePath(TmpSmpFile)
    stringfl = stringfl & "\" & ProductionRun.Barcode_Id & "_" & TmpTypeNum & ".smp"
    'create the new file
    Call FileCopy(strFileName, stringfl)
    
    'remove the temporary sample
    Kill (strFileName)
    
    'call to upate the sample table
    Call UpdateSmplFile_On_DB(TmpTypeNum, stringfl)

End Sub

Private Sub AlignSmplFileName_After_Delete()
'md go through all samples to correct files if necessary
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
        .AddParameter "Type Number", 0, adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                Call Process_Rename_Files(.Recordset!Sample_File_Name, .Recordset!Type_Number)
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
    
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error getting existing sample files"
    Resume Exit_this_Sub
    
End Sub

Private Sub UpdateSmplFile_On_DB(typenum As Integer, NewSmplFile As String)
  'md new for Clintrak Samples project - Update the sample type with the correct file name
  
 On Error GoTo Error_this_Sub

    Dim nreturn As Long
    Dim mikeoData As nADOData.CADOData
    
    If mikeoData Is Nothing Then
        Set mikeoData = New CADOData
        With mikeoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If
    
    With mikeoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly

        .AddParameter "Production Run Id", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .AddParameter "Type Number", CInt(typenum), adInteger, adParamInput
        .AddParameter "Smpl File", NewSmplFile, adVarChar, adParamInput
        
        .AddParameter "return", "   ", adInteger, adParamOutput
        
        .ExecuteSP "Update_SampleFiles", True
        
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        If IsNull(nreturn) Or Trim(nreturn) = "" Or nreturn <> 0 Then
            .Connection.RollbackTrans
            Exit Sub
        End If
    End With
    
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Saving sample File information"
    Resume Exit_this_Sub
  
End Sub

Private Sub Extract_MergeBarcode_Columns(mergeCoding As String, ByRef colArray() As String, With_Delimiters As Boolean)
'
'comments:  This function extracts the merged barcode data into an array separating column numbers and delimiters
'parameter: mergeCoding - delimited coding string, colArray - array to put extracted data into
'           with_delimiter - determines whether or not to included the delimiter in the array
'returns:   true if data exists
'
On Error GoTo Handle_Error

    Dim i As Long
    Dim tempHolder As String    'holds the temporary substrings
    ReDim colArray(0)
    
    For i = 1 To Len(Trim(mergeCoding))
        
        If IsNumeric(Mid$(Trim(mergeCoding), i, 1)) Then
            tempHolder = tempHolder & Mid$(Trim(mergeCoding), i, 1)
        Else
            If With_Delimiters Then
                ReDim Preserve colArray(UBound(colArray) + 2)
                colArray(UBound(colArray)) = Mid$(Trim(mergeCoding), i, 1)
                colArray(UBound(colArray) - 1) = tempHolder
            Else
                ReDim Preserve colArray(UBound(colArray) + 1)
                colArray(UBound(colArray)) = tempHolder
            End If
            tempHolder = ""
        End If
    Next
    
    'add the last column to the array
    If IsNumeric(tempHolder) Then
        ReDim Preserve colArray(UBound(colArray) + 1)
        colArray(UBound(colArray)) = tempHolder
    End If
    
exit_sub:
    Exit Sub

Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description, , "Extract_MergeBarcode_Columns"
    Resume exit_sub

End Sub

Private Function Display_Barcode(Selected_Column_Array() As String, Coding_File_Data() As String) As String
'Function to return Barcode from the Input String Array
'comments:  This function returns a string of what the barcode data will look like
'returns:   string representing the merge barcode data

On Error GoTo Handle_Error

    Dim i As Long
    Dim outputstring As String
    
    For i = 1 To UBound(Selected_Column_Array)
        If outputstring = "" Then
            If IsNumeric(Selected_Column_Array(i)) Then
                If CInt(Selected_Column_Array(i)) - 1 <= UBound(Coding_File_Data) Then
                    outputstring = Coding_File_Data(CInt(Selected_Column_Array(i)) - 1)
                End If
            Else
                outputstring = CheckBarCodeOutPut(Selected_Column_Array(i))
            End If
        Else
            If IsNumeric(Selected_Column_Array(i)) Then
                If (CInt(Selected_Column_Array(i)) - 1) <= UBound(Coding_File_Data) Then
                    outputstring = outputstring & Coding_File_Data(CInt(Selected_Column_Array(i)) - 1)
                End If
            Else
                If Not Selected_Column_Array(i) = CNODELIMITER Then outputstring = outputstring & CheckBarCodeOutPut(Selected_Column_Array(i))
          End If
        End If
    Next
    
    Display_Barcode = outputstring

exit_function:
    Exit Function

Handle_Error:
   'Log errors and pass to calling procedure
    Err.Raise Err.Number, Err.Source & "->DisplayBarcode()", Err.description
    Resume exit_function

End Function

Private Function CheckBarCodeOutPut(strColumn As String) As String
    ' Returns the desired output pending the input column designation
    On Error GoTo Handle_Error
    Select Case strColumn
        Case "R"
            CheckBarCodeOutPut = Format$(Mid$(gRandBarcode, 4, Len(gRandBarcode) - 3), "000000")
        Case "P"
            If booReplacement Then
                CheckBarCodeOutPut = Format$(Mid$(gOrigPDRNumber, 4, Len(gOrigPDRNumber) - 3), "0000000")
            Else
                CheckBarCodeOutPut = Format$(Mid$(ProductionRun.Barcode_Id, 4, Len(ProductionRun.Barcode_Id) - 3), "0000000")
            End If
        Case Is <> CNODELIMITER
            CheckBarCodeOutPut = strColumn
    End Select

exit_function:
    Exit Function

Handle_Error:
   'Log errors and pass to calling procedure
    Err.Raise Err.Number, Err.Source & "->CheckBarCodeOutPut()", Err.description
    Resume exit_function
End Function

Private Sub LockOutForm()
    On Error GoTo Handle_Error

    If mvarCombinedOrRun = True Then
        Me.jgrdData.AllowEdit = False
        Me.SSDBComboSmpType.enabled = False
        Me.txtQtynumber.enabled = False
        Me.Update.enabled = False
        Me.cmdLoadCoding.enabled = False
        Me.cmdConfigure.enabled = False
        Me.mnuLoadSample.enabled = False
        Me.cmdDeleteButton.enabled = False
        Me.cmdAdd.enabled = False
    End If
        
Cleanup_Exit:
    Exit Sub
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit

End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Handle_Error

    Call cmdNextButton_Click
    cmdAdd.enabled = False
                

Cleanup_Exit:
     Exit Sub
Handle_Error:
     MsgBox Err.description & vbCrLf & _
         "in frmSmpConfig.cmdAdd_Click ", _
         vbCritical + vbOKOnly, "Application Error"
     Resume Cleanup_Exit

End Sub



