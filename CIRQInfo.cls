VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIRQInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object holds information about inventory requests.</summary>
'</comment>

Option Explicit

Public IRQ_Proof_Id As Long
Public IRQ_Id As Long
Public IRQ_Number As String
Public PDR_Count As Integer
Public IRQ_Label_Identification As String
Public IRQ_Details_Id As Long
Public IRQ_Details_Qty_Requested As Long
Public IRQ_Status As String
Public IRQ_Main_Proof_Id As Long

