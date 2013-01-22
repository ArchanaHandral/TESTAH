VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CsmpFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object holds information about sample type configurations.</summary>
'</comment>

Option Explicit

Public productionId As Long
Public typeNumber As Integer
Public sampleType As String
Public shipTo As Long           ' DW ECR12-008 update integer to long
Public quantity As Long
Public smpDescription As String
Public smpfileName As String
Public notes As String
Public sample_type_id As Long
