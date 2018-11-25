VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollIRQInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object is a Collection of CIRQInfo.</summary>
'</comment>

Option Explicit

Private mcol As Collection

Public Function Add() As CIRQInfo

    'create new object
    Dim objNewMember As CIRQInfo
    Set objNewMember = New CIRQInfo

    mcol.Add objNewMember

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As CIRQInfo
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    count = mcol.count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mcol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mcol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mcol = Nothing
End Sub

Public Function Clone() As CCollIRQInfo
    'returns a separate copy of the
    'current collection
    'Syntax: Set y = x.Clone
    'Function added 4/10/03 by CR; used in Delete Production Record
    
    Dim i As Long
    Dim CCollClone As CCollIRQInfo
    Dim oIRQ As CIRQInfo
    
    Set CCollClone = New CCollIRQInfo
    
    For i = 1 To Me.count
    
        Set oIRQ = CCollClone.Add
            With CCollClone.Item(i)
                .IRQ_Proof_Id = Me.Item(i).IRQ_Proof_Id
                .IRQ_Id = Me.Item(i).IRQ_Id
                .IRQ_Number = Me.Item(i).IRQ_Number
                .PDR_Count = Me.Item(i).PDR_Count
                .IRQ_Label_Identification = Me.Item(i).IRQ_Label_Identification
                .IRQ_Details_Id = Me.Item(i).IRQ_Details_Id
                .IRQ_Details_Qty_Requested = Me.Item(i).IRQ_Details_Qty_Requested
                .IRQ_Status = Me.Item(i).IRQ_Status
                .IRQ_Main_Proof_Id = Me.Item(i).IRQ_Main_Proof_Id
                .IRQToolComponentId = Me.Item(i).IRQToolComponentId
            End With
    Next i
    
    Set Clone = CCollClone

End Function
