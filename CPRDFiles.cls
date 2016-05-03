VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPRDFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object holds data that will be written to a line in a file.</summary>
'</comment>

Option Explicit

Public Field1 As String
Public Field2 As String
Public Field3 As String
Public Field4 As String
Public Field5 As String
Public Field6 As String
Public Field7 As String
Public Field8 As String
Public Field9 As String
Public Field10 As String
Public Field11 As String
Public Field12 As String
Public Field13 As String
Public Field14 As String
Public Field15 As String
Public Field16 As String
Public Field17 As String
Public Field18 As String
Public Field19 As String
Public Field20 As String
' DW increasing # of columns from 20 to 30 based on client supplied data
Public Field21 As String
Public Field22 As String
Public Field23 As String
Public Field24 As String
Public Field25 As String
Public Field26 As String
Public Field27 As String
Public Field28 As String
Public Field29 As String
Public Field30 As String

Public Property Get Fields(index As Long) As String
Select Case index
    Case 1
        Fields = Field1
    Case 2
        Fields = Field2
    Case 3
        Fields = Field3
    Case 4
        Fields = Field4
    Case 5
        Fields = Field5
    Case 6
        Fields = Field6
    Case 7
        Fields = Field7
    Case 8
        Fields = Field8
    Case 9
        Fields = Field9
    Case 10
        Fields = Field10
        
    Case 11
        Fields = Field11
    Case 12
        Fields = Field12
    Case 13
        Fields = Field13
    Case 14
        Fields = Field14
    Case 15
        Fields = Field15
    Case 16
        Fields = Field16
    Case 17
        Fields = Field17
    Case 18
        Fields = Field18
    Case 19
        Fields = Field19
    Case 20
        Fields = Field20
        
    Case 21
        Fields = Field21
    Case 22
        Fields = Field22
    Case 23
        Fields = Field23
    Case 24
        Fields = Field24
    Case 25
        Fields = Field25
    Case 26
        Fields = Field26
    Case 27
        Fields = Field27
    Case 28
        Fields = Field28
    Case 29
        Fields = Field29
    Case 30
        Fields = Field30

    Case Else
        MsgBox "No such field: " & index
        Fields = ""
End Select
End Property
