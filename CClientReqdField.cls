VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClientReqdField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object holds information about client required fields for a production run.</summary>
'</comment>

Option Explicit

'Client Required Fields
'Public Client_Required_Fields_Id As Long
Public Client_Id As Long
Public Client_Required_Field_Name As String

'Production Run Client Fields (Values)
Public Production_Run_Client_Fields_Id As Long
Public Production_Run_Id As Long
Public Field_Name_Value As String
Public Temp_Field_Name_Value As String
