VERSION 5.00
Begin VB.Form frmGridConfig 
   Caption         =   "Configure"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   3510
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSuffix 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtPrefix 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtChangeCol 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.CheckBox chkSequence 
      Caption         =   "Sequence Data"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1335
      Width           =   1695
   End
   Begin VB.CheckBox chkSuffix 
      Caption         =   "Add Suffix:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   975
      Width           =   1815
   End
   Begin VB.CheckBox chkPrefix 
      Caption         =   "Add Prefix:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   615
      Width           =   1695
   End
   Begin VB.CheckBox chkChangeCol 
      Caption         =   "Change Column Data:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   255
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmGridConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This form is called from frmSmpConfig and is used to modify the Sample Configuration Data for a column of data.</summary>
'</comment>

Option Explicit
Dim booConfigSelect As Boolean

Private Sub chkChangeCol_Click()
    If chkChangeCol = 1 Then
        chkSequence.Enabled = False
    Else
        chkSequence.Enabled = True
    End If
End Sub

Private Sub chkSequence_Click()
    If chkSequence = 1 Then
       chkChangeCol.Enabled = False
    Else
       chkChangeCol.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If chkSequence = 1 Then
        If Not SequenceData(columnNumber) Then
            chkSequence = 0
            Exit Sub
        End If
    End If
    
    If chkChangeCol = 1 Then
         If Not ChangeAll(txtChangeCol.Text) Then
            Exit Sub
         End If
    End If
    
    If chkPrefix = 1 Then
        ChangePrefix (txtPrefix.Text)
    End If
    
    If chkSuffix = 1 Then
        ChangeSuffix (txtSuffix.Text)
    End If
    
    frmSmpConfig.jgrdData.ItemCount = mData.count
    frmSmpConfig.jgrdData.Update
    frmSmpConfig.jgrdData.Refresh
    frmSmpConfig.dirtyFlag = "Y"
    Unload Me
End Sub

'<comment>
' <summary>
'       this function checks whether the column of data is numeric, if it is
'       then it sequences the data in it.</summary>
' <param name="col">column to sequence</param>
' <return>true if the column data is numeric</return>
'</comment>
Private Function SequenceData(col As Integer) As Boolean
    On Error GoTo Handle_Error
    
    Dim i As Long
    Dim booColNumeric As Boolean
    
    SequenceData = False
    
    For i = 1 To mData.count
        Select Case col
            Case 1
                booColNumeric = IsNumeric(mData.Item(i).Field1)
            Case 2
                booColNumeric = IsNumeric(mData.Item(i).Field2)
            Case 3
                booColNumeric = IsNumeric(mData.Item(i).Field3)
            Case 4
                booColNumeric = IsNumeric(mData.Item(i).Field4)
            Case 5
                booColNumeric = IsNumeric(mData.Item(i).Field5)
            Case 6
                booColNumeric = IsNumeric(mData.Item(i).Field6)
            Case 7
                booColNumeric = IsNumeric(mData.Item(i).Field7)
            Case 8
                booColNumeric = IsNumeric(mData.Item(i).Field8)
            Case 9
                booColNumeric = IsNumeric(mData.Item(i).Field9)
            Case 10
                booColNumeric = IsNumeric(mData.Item(i).Field10)
            Case 11
                booColNumeric = IsNumeric(mData.Item(i).Field11)
            Case 12
                booColNumeric = IsNumeric(mData.Item(i).Field12)
            Case 13
                booColNumeric = IsNumeric(mData.Item(i).Field13)
            Case 14
                booColNumeric = IsNumeric(mData.Item(i).Field14)
            Case 15
                booColNumeric = IsNumeric(mData.Item(i).Field15)
            Case 16
                booColNumeric = IsNumeric(mData.Item(i).Field16)
            Case 17
                booColNumeric = IsNumeric(mData.Item(i).Field17)
            Case 18
                booColNumeric = IsNumeric(mData.Item(i).Field18)
            Case 19
                booColNumeric = IsNumeric(mData.Item(i).Field19)
            Case 20
                booColNumeric = IsNumeric(mData.Item(i).Field20)
            ' DW increasing # of columns from 20 to 30 based on client supplied data
            Case 21
                booColNumeric = IsNumeric(mData.Item(i).Field21)
            Case 22
                booColNumeric = IsNumeric(mData.Item(i).Field22)
            Case 23
                booColNumeric = IsNumeric(mData.Item(i).Field23)
            Case 24
                booColNumeric = IsNumeric(mData.Item(i).Field24)
            Case 25
                booColNumeric = IsNumeric(mData.Item(i).Field25)
            Case 26
                booColNumeric = IsNumeric(mData.Item(i).Field26)
            Case 27
                booColNumeric = IsNumeric(mData.Item(i).Field27)
            Case 28
                booColNumeric = IsNumeric(mData.Item(i).Field28)
            Case 29
                booColNumeric = IsNumeric(mData.Item(i).Field29)
            Case 30
                booColNumeric = IsNumeric(mData.Item(i).Field30)
            
        End Select
        
        If Not booColNumeric Then
            MsgBox "The column data is not numeric!", vbExclamation
            Exit Function
        End If
    Next


    For i = 1 To mData.count
        Select Case col
            Case 1
                mData.Item(i).Field1 = CLng(mData.Item(1).Field1) + i - 1
            Case 2
                mData.Item(i).Field2 = CLng(mData.Item(1).Field2) + i - 1
            Case 3
                mData.Item(i).Field3 = CLng(mData.Item(1).Field3) + i - 1
            Case 4
                mData.Item(i).Field4 = CLng(mData.Item(1).Field4) + i - 1
            Case 5
                mData.Item(i).Field5 = CLng(mData.Item(1).Field5) + i - 1
            Case 6
                mData.Item(i).Field6 = CLng(mData.Item(1).Field6) + i - 1
            Case 7
                mData.Item(i).Field7 = CLng(mData.Item(1).Field7) + i - 1
            Case 8
                mData.Item(i).Field8 = CLng(mData.Item(1).Field8) + i - 1
            Case 9
                mData.Item(i).Field9 = CLng(mData.Item(1).Field9) + i - 1
            Case 10
                mData.Item(i).Field10 = CLng(mData.Item(1).Field10) + i - 1
            Case 11
                mData.Item(i).Field11 = CLng(mData.Item(1).Field11) + i - 1
            Case 12
                mData.Item(i).Field12 = CLng(mData.Item(1).Field12) + i - 1
            Case 13
                mData.Item(i).Field13 = CLng(mData.Item(1).Field13) + i - 1
            Case 14
                mData.Item(i).Field14 = CLng(mData.Item(1).Field14) + i - 1
            Case 15
                mData.Item(i).Field15 = CLng(mData.Item(1).Field15) + i - 1
            Case 16
                mData.Item(i).Field16 = CLng(mData.Item(1).Field16) + i - 1
            Case 17
                mData.Item(i).Field17 = CLng(mData.Item(1).Field17) + i - 1
            Case 18
                mData.Item(i).Field18 = CLng(mData.Item(1).Field18) + i - 1
            Case 19
                mData.Item(i).Field19 = CLng(mData.Item(1).Field19) + i - 1
            Case 20
                mData.Item(i).Field20 = CLng(mData.Item(1).Field20) + i - 1
            ' DW increasing # of columns from 20 to 30 based on client supplied data
            Case 21
                mData.Item(i).Field21 = CLng(mData.Item(1).Field21) + i - 1
            Case 22
                mData.Item(i).Field22 = CLng(mData.Item(1).Field22) + i - 1
            Case 23
                mData.Item(i).Field23 = CLng(mData.Item(1).Field23) + i - 1
            Case 24
                mData.Item(i).Field24 = CLng(mData.Item(1).Field24) + i - 1
            Case 25
                mData.Item(i).Field25 = CLng(mData.Item(1).Field25) + i - 1
            Case 26
                mData.Item(i).Field26 = CLng(mData.Item(1).Field26) + i - 1
            Case 27
                mData.Item(i).Field27 = CLng(mData.Item(1).Field27) + i - 1
            Case 28
                mData.Item(i).Field28 = CLng(mData.Item(1).Field28) + i - 1
            Case 29
                mData.Item(i).Field29 = CLng(mData.Item(1).Field29) + i - 1
            Case 30
                mData.Item(i).Field30 = CLng(mData.Item(1).Field30) + i - 1
        End Select
    Next

    SequenceData = True

Cleanup_Exit:
    Exit Function
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Function

Private Function ChangeAll(txt As String) As Boolean
'
'comments:  this function changes all the data inside a column to the specified text
'parameters:    txt - text to change to
'returns:   true text is entered and changed to
'
Dim i As Long
ChangeAll = False

    If Not booConfigSelect Then
        For i = 1 To mData.count
            Select Case columnNumber
                Case 1
                    mData.Item(i).Field1 = txt
                Case 2
                    mData.Item(i).Field2 = txt
                Case 3
                    mData.Item(i).Field3 = txt
                Case 4
                    mData.Item(i).Field4 = txt
                Case 5
                    mData.Item(i).Field5 = txt
                Case 6
                    mData.Item(i).Field6 = txt
                Case 7
                    mData.Item(i).Field7 = txt
                Case 8
                    mData.Item(i).Field8 = txt
                Case 9
                    mData.Item(i).Field9 = txt
                Case 10
                    mData.Item(i).Field10 = txt
                Case 11
                    mData.Item(i).Field11 = txt
                Case 12
                    mData.Item(i).Field12 = txt
                Case 13
                    mData.Item(i).Field13 = txt
                Case 14
                    mData.Item(i).Field14 = txt
                Case 15
                    mData.Item(i).Field15 = txt
                Case 16
                    mData.Item(i).Field16 = txt
                Case 17
                    mData.Item(i).Field17 = txt
                Case 18
                    mData.Item(i).Field18 = txt
                Case 19
                    mData.Item(i).Field19 = txt
                Case 20
                    mData.Item(i).Field20 = txt
                ' DW increasing # of columns from 20 to 30 based on client supplied data
                Case 21
                    mData.Item(i).Field21 = txt
                Case 22
                    mData.Item(i).Field22 = txt
                Case 23
                    mData.Item(i).Field23 = txt
                Case 24
                    mData.Item(i).Field24 = txt
                Case 25
                    mData.Item(i).Field25 = txt
                Case 26
                    mData.Item(i).Field26 = txt
                Case 27
                    mData.Item(i).Field27 = txt
                Case 28
                    mData.Item(i).Field28 = txt
                Case 29
                    mData.Item(i).Field29 = txt
                Case 30
                    mData.Item(i).Field30 = txt
            End Select
        Next
    Else
        ' gave the user the option of only configuring changes to selected rows
        For i = 1 To frmSmpConfig.jgrdData.SelectedItems.count
            Select Case columnNumber
                Case 1
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field1 = txt
                Case 2
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field2 = txt
                Case 3
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field3 = txt
                Case 4
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field4 = txt
                Case 5
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field5 = txt
                Case 6
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field6 = txt
                Case 7
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field7 = txt
                Case 8
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field8 = txt
                Case 9
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field9 = txt
                Case 10
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field10 = txt
                Case 11
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field11 = txt
                Case 12
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field12 = txt
                Case 13
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field13 = txt
                Case 14
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field14 = txt
                Case 15
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field15 = txt
                Case 16
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field16 = txt
                Case 17
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field17 = txt
                Case 18
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field18 = txt
                Case 19
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field19 = txt
                Case 20
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field20 = txt
                ' DW increasing # of columns from 20 to 30 based on client supplied data
                Case 21
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field21 = txt
                Case 22
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field22 = txt
                Case 23
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field23 = txt
                Case 24
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field24 = txt
                Case 25
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field25 = txt
                Case 26
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field26 = txt
                Case 27
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field27 = txt
                Case 28
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field28 = txt
                Case 29
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field29 = txt
                Case 30
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field30 = txt
            End Select
        Next
    End If
        

ChangeAll = True

End Function

Private Sub ChangePrefix(txt As String)
Dim i As Long

    If Not booConfigSelect Then
        For i = 1 To mData.count
            Select Case columnNumber
                Case 1
                    mData.Item(i).Field1 = txt & mData.Item(i).Field1
                Case 2
                    mData.Item(i).Field2 = txt & mData.Item(i).Field2
                Case 3
                    mData.Item(i).Field3 = txt & mData.Item(i).Field3
                Case 4
                    mData.Item(i).Field4 = txt & mData.Item(i).Field4
                Case 5
                    mData.Item(i).Field5 = txt & mData.Item(i).Field5
                Case 6
                    mData.Item(i).Field6 = txt & mData.Item(i).Field6
                Case 7
                    mData.Item(i).Field7 = txt & mData.Item(i).Field7
                Case 8
                    mData.Item(i).Field8 = txt & mData.Item(i).Field8
                Case 9
                    mData.Item(i).Field9 = txt & mData.Item(i).Field9
                Case 10
                    mData.Item(i).Field10 = txt & mData.Item(i).Field10
                Case 11
                    mData.Item(i).Field11 = txt & mData.Item(i).Field11
                Case 12
                    mData.Item(i).Field12 = txt & mData.Item(i).Field12
                Case 13
                    mData.Item(i).Field13 = txt & mData.Item(i).Field13
                Case 14
                    mData.Item(i).Field14 = txt & mData.Item(i).Field14
                Case 15
                    mData.Item(i).Field15 = txt & mData.Item(i).Field15
                Case 16
                    mData.Item(i).Field16 = txt & mData.Item(i).Field16
                Case 17
                    mData.Item(i).Field17 = txt & mData.Item(i).Field17
                Case 18
                    mData.Item(i).Field18 = txt & mData.Item(i).Field18
                Case 19
                    mData.Item(i).Field19 = txt & mData.Item(i).Field19
                Case 20
                    mData.Item(i).Field20 = txt & mData.Item(i).Field20
                ' DW increasing # of columns from 20 to 30 based on client supplied data
                Case 21
                    mData.Item(i).Field21 = txt & mData.Item(i).Field21
                Case 22
                    mData.Item(i).Field22 = txt & mData.Item(i).Field22
                Case 23
                    mData.Item(i).Field23 = txt & mData.Item(i).Field23
                Case 24
                    mData.Item(i).Field24 = txt & mData.Item(i).Field24
                Case 25
                    mData.Item(i).Field25 = txt & mData.Item(i).Field25
                Case 26
                    mData.Item(i).Field26 = txt & mData.Item(i).Field26
                Case 27
                    mData.Item(i).Field27 = txt & mData.Item(i).Field27
                Case 28
                    mData.Item(i).Field28 = txt & mData.Item(i).Field28
                Case 29
                    mData.Item(i).Field29 = txt & mData.Item(i).Field29
                Case 30
                    mData.Item(i).Field30 = txt & mData.Item(i).Field30
            End Select
        Next
    Else
        ' gave the user the option of only configuring changes to selected rows
        For i = 1 To frmSmpConfig.jgrdData.SelectedItems.count
            Select Case columnNumber
                Case 1
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field1 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field1
                Case 2
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field2 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field2
                Case 3
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field3 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field3
                Case 4
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field4 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field4
                Case 5
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field5 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field5
                Case 6
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field6 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field6
                Case 7
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field7 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field7
                Case 8
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field8 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field8
                Case 9
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field9 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field9
                Case 10
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field10 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field10
                Case 11
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field11 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field11
                Case 12
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field12 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field12
                Case 13
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field13 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field13
                Case 14
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field14 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field14
                Case 15
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field15 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field15
                Case 16
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field16 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field16
                Case 17
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field17 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field17
                Case 18
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field18 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field18
                Case 19
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field19 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field19
                Case 20
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field20 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field20
                ' DW increasing # of columns from 20 to 30 based on client supplied data
                Case 21
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field21 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field21
                Case 22
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field22 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field22
                Case 23
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field23 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field23
                Case 24
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field24 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field24
                Case 25
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field25 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field25
                Case 26
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field26 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field26
                Case 27
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field27 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field27
                Case 28
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field28 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field28
                Case 29
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field29 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field29
                Case 30
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field30 = txt & mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field30
            End Select
        Next
    End If

End Sub

Private Sub ChangeSuffix(txt As String)
Dim i As Long

    If Not booConfigSelect Then
        For i = 1 To mData.count
            Select Case columnNumber
                Case 1
                    mData.Item(i).Field1 = mData.Item(i).Field1 & txt
                Case 2
                    mData.Item(i).Field2 = mData.Item(i).Field2 & txt
                Case 3
                    mData.Item(i).Field3 = mData.Item(i).Field3 & txt
                Case 4
                    mData.Item(i).Field4 = mData.Item(i).Field4 & txt
                Case 5
                    mData.Item(i).Field5 = mData.Item(i).Field5 & txt
                Case 6
                    mData.Item(i).Field6 = mData.Item(i).Field6 & txt
                Case 7
                    mData.Item(i).Field7 = mData.Item(i).Field7 & txt
                Case 8
                    mData.Item(i).Field8 = mData.Item(i).Field8 & txt
                Case 9
                    mData.Item(i).Field9 = mData.Item(i).Field9 & txt
                Case 10
                    mData.Item(i).Field10 = mData.Item(i).Field10 & txt
                Case 11
                    mData.Item(i).Field11 = mData.Item(i).Field11 & txt
                Case 12
                    mData.Item(i).Field12 = mData.Item(i).Field12 & txt
                Case 13
                    mData.Item(i).Field13 = mData.Item(i).Field13 & txt
                Case 14
                    mData.Item(i).Field14 = mData.Item(i).Field14 & txt
                Case 15
                    mData.Item(i).Field15 = mData.Item(i).Field15 & txt
                Case 16
                    mData.Item(i).Field16 = mData.Item(i).Field16 & txt
                Case 17
                    mData.Item(i).Field17 = mData.Item(i).Field17 & txt
                Case 18
                    mData.Item(i).Field18 = mData.Item(i).Field18 & txt
                Case 19
                    mData.Item(i).Field19 = mData.Item(i).Field19 & txt
                Case 20
                    mData.Item(i).Field20 = mData.Item(i).Field20 & txt
                ' DW increasing # of columns from 20 to 30 based on client supplied data
                Case 21
                    mData.Item(i).Field21 = mData.Item(i).Field21 & txt
                Case 22
                    mData.Item(i).Field22 = mData.Item(i).Field22 & txt
                Case 23
                    mData.Item(i).Field23 = mData.Item(i).Field23 & txt
                Case 24
                    mData.Item(i).Field24 = mData.Item(i).Field24 & txt
                Case 25
                    mData.Item(i).Field25 = mData.Item(i).Field25 & txt
                Case 26
                    mData.Item(i).Field26 = mData.Item(i).Field26 & txt
                Case 27
                    mData.Item(i).Field27 = mData.Item(i).Field27 & txt
                Case 28
                    mData.Item(i).Field28 = mData.Item(i).Field28 & txt
                Case 29
                    mData.Item(i).Field29 = mData.Item(i).Field29 & txt
                Case 30
                    mData.Item(i).Field30 = mData.Item(i).Field30 & txt
            End Select
        Next
    Else
        ' gave the user the option of only configuring changes to selected rows
         For i = 1 To frmSmpConfig.jgrdData.SelectedItems.count
            Select Case columnNumber
                Case 1
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field1 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field1 & txt
                Case 2
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field2 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field2 & txt
                Case 3
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field3 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field3 & txt
                Case 4
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field4 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field4 & txt
                Case 5
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field5 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field5 & txt
                Case 6
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field6 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field6 & txt
                Case 7
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field7 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field7 & txt
                Case 8
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field8 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field8 & txt
                Case 9
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field9 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field9 & txt
                Case 10
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field10 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field10 & txt
                Case 11
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field11 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field11 & txt
                Case 12
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field12 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field12 & txt
                Case 13
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field13 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field13 & txt
                Case 14
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field14 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field14 & txt
                Case 15
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field15 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field15 & txt
                Case 16
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field16 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field16 & txt
                Case 17
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field17 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field17 & txt
                Case 18
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field18 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field18 & txt
                Case 19
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field19 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field19 & txt
                Case 20
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field20 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field20 & txt
                ' DW increasing # of columns from 20 to 30 based on client supplied data
                Case 21
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field21 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field21 & txt
                Case 22
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field22 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field22 & txt
                Case 23
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field23 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field23 & txt
                Case 24
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field24 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field24 & txt
                Case 25
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field25 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field25 & txt
                Case 26
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field26 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field26 & txt
                Case 27
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field27 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field27 & txt
                Case 28
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field28 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field28 & txt
                Case 29
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field29 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field29 & txt
                Case 30
                    mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field30 = mData.Item(frmSmpConfig.jgrdData.SelectedItems(i).RowIndex).Field30 & txt
            End Select
        Next
    End If

End Sub
Private Sub Form_Load()
    ' ability to configure only selected samples
    If frmSmpConfig.jgrdData.SelectedItems.count > 0 Then
    
        If MsgBox("Do you wish to configure the selected samples only?", _
            vbQuestion + vbYesNo) = vbYes Then
            
            booConfigSelect = True
            Caption = "Configure Selected Samples"
            chkSequence.Enabled = False
        Else
            booConfigSelect = False
            Caption = "Configure All Samples"
            chkSequence.Enabled = True
        End If
        
    Else
        Caption = "Configure All Samples"
        booConfigSelect = False
    End If
End Sub
