VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmClientReqdFields 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Required Data"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin GridEX20.GridEX GridExClientFields 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4048
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmClientReqdFields.frx":0000
      Column(2)       =   "frmClientReqdFields.frx":0154
      Column(3)       =   "frmClientReqdFields.frx":028C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmClientReqdFields.frx":03AC
      FormatStyle(2)  =   "frmClientReqdFields.frx":04E4
      FormatStyle(3)  =   "frmClientReqdFields.frx":0594
      FormatStyle(4)  =   "frmClientReqdFields.frx":0648
      FormatStyle(5)  =   "frmClientReqdFields.frx":0720
      FormatStyle(6)  =   "frmClientReqdFields.frx":07D8
      ImageCount      =   0
      PrinterProperties=   "frmClientReqdFields.frx":08B8
   End
End
Attribute VB_Name = "frmClientReqdFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This form is called from frmProdPlan and is used to capture values for client specific fields.</summary>
'</comment>

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim booLockScreen As Boolean
    
    booLockScreen = False
    
    Me.GridExClientFields.ItemCount = 0
    
    For i = 1 To ClientReqdFields.count
        ClientReqdFields.Item(i).Temp_Field_Name_Value = ClientReqdFields.Item(i).Field_Name_Value
    Next i
    
    Me.GridExClientFields.ItemCount = ClientReqdFields.count
    
    Me.GridExClientFields.Refetch
    
    With frmProdPlan
        If Not booNewProdRun Then
        ' if it is not a new PDR, check if the PDR has been put on a PKS,
        ' or if the Job's completely shipped flag is on and if so, lock the screen
            If .Determine_Shipping_Flag_On = True Or _
                    DeterminePDROnPKS = True Then
                booLockScreen = True
            End If
        End If
        
        ' if the save button is not enabled then these shouldn't be editable b/c they can't be saved
        ' and if this is a replacement PDR, these should never be editable
        If .cmdSave.Enabled = False Or _
                .txtReplacement.Visible = True Then
            booLockScreen = True
        End If
    End With
    
    If booLockScreen = True Then
        Me.GridExClientFields.Enabled = False
        Me.GridExClientFields.ForeColor = &H80000011      'set disabled forecolor gray in grid
    Else
        Me.GridExClientFields.Enabled = True
        Me.GridExClientFields.ForeColor = &H80000008      'set enabled forecolor Black in grid
    End If
       
End Sub

Private Sub GridExClientFields_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim oFields As CClientReqdField

    On Error GoTo Handle_Error
    
    If Not ClientReqdFields Is Nothing And ClientReqdFields.count > 0 Then
        Set oFields = ClientReqdFields.Item(RowIndex)
        With oFields
            Values(1) = .Client_Required_Field_Name
            Values(2) = .Temp_Field_Name_Value
            Values(3) = .Production_Run_Client_Fields_Id
        End With
    End If
    
exit_function:
    Exit Sub
    
Handle_Error:
    Err.Raise Err.Number, "GridExClientFields_UnboundReadData()->" & Err.Source, Err.description
    GoTo exit_function
    
End Sub

Private Sub GridExClientFields_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim oFields As CClientReqdField

    On Error GoTo Handle_Error
    
    If Not ClientReqdFields Is Nothing And ClientReqdFields.count > 0 Then
        Set oFields = ClientReqdFields.Item(RowIndex)
        With oFields
            .Client_Required_Field_Name = Trim$(Values(1))
            .Temp_Field_Name_Value = Trim$(Values(2))
            .Production_Run_Client_Fields_Id = Trim$(Values(3))
        End With
    End If

exit_function:
    Exit Sub
    
Handle_Error:
    Err.Raise Err.Number, "GridExClientFields_UnboundUpdate()->" & Err.Source, Err.description
    GoTo exit_function
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    Me.GridExClientFields.Update
    
    For i = 1 To ClientReqdFields.count
        ClientReqdFields.Item(i).Field_Name_Value = ClientReqdFields.Item(i).Temp_Field_Name_Value
    Next i

    Unload Me
    
End Sub

