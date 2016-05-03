VERSION 5.00
Begin VB.Form frmOverride 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supervisor Override"
   ClientHeight    =   2580
   ClientLeft      =   7875
   ClientTop       =   3555
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1524.35
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   750
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1620
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2085
      TabIndex        =   5
      Top             =   1620
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1125
      Width           =   2325
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   $"frmOverride.frx":0000
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "*Password is case sensitive."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   638
      TabIndex        =   6
      Top             =   2280
      Width           =   2475
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "*&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   1140
      Width           =   1080
   End
End
Attribute VB_Name = "frmOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarOverrideUser As ClintrakCommon.ApplicationUser
Private mvarRequiredOverride As Long

Public Property Let RequiredOverride(value As Long)
    mvarRequiredOverride = value
End Property

Public Property Get OverrideUserEmployeeId() As Long
    If mvarOverrideUser Is Nothing Then
        OverrideUserEmployeeId = 0
    Else
        OverrideUserEmployeeId = mvarOverrideUser.employeeId
    End If
End Property

Private Sub cmdCancel_Click()
    Set mvarOverrideUser = Nothing
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo PROC_ERR
    
    Screen.MousePointer = vbHourglass
    
    If mvarOverrideUser.Authenticate(txtUserName.text, txtPassword.text, gApplicationUser.Domain, override, gApplicationUser.SQLServer, gApplicationUser.SQLDatabase, App.Title) Then
        ' Check Override permission via ACL
        If mvarOverrideUser.HasAccess(mvarRequiredOverride) Then
            'Override Access as defined by ACL
            Me.Hide
        Else
            'Application Level access failed.
            MsgBox "You don't have permission for the specified override." & vbCrLf & _
            "Please contact your Supervisor.", vbInformation, "Override Permissions"

            Set mvarOverrideUser = Nothing
            Me.Hide
        End If
    Else
        txtUserName.SetFocus
    End If

Proc_EXIT:
    Screen.MousePointer = vbDefault
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "ERROR - Override"
    Resume Proc_EXIT
End Sub

Private Sub Form_Load()
    Set mvarOverrideUser = New ClintrakCommon.ApplicationUser
    
    CenterForm Me
End Sub

Public Static Function GetOverrideUser(override As Long, prompt As String) As Long
    Dim it As frmOverride
    
    Set it = New frmOverride
    it.RequiredOverride = override
    it.lblPrompt.Caption = prompt
    it.Show vbModal
    GetOverrideUser = it.OverrideUserEmployeeId
    
    Unload it
End Function

