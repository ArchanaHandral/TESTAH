VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColBarcodeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object is a Collection of CBarcodeInfo.</summary>
'</comment>

Option Explicit

'local variable to hold collection
Private mcol As Collection              'Collection of CBarcodeInfo objects

Public Function Add(Optional sKey As String) As CBarcodeInfo
                    
                    
    'create a new CBarcodeInfo object
    Dim objNewMember As CBarcodeInfo
    Set objNewMember = New CBarcodeInfo
    
    
    'Optional default for collection key
    If Len(sKey) = 0 Then
        mcol.Add objNewMember
    Else
        mcol.Add objNewMember, sKey
    End If
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As CBarcodeInfo
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


Public Function Clone() As CColBarcodeInfo
    'used to return a separate copy of the collection
    'Syntax: Set y = x.Clone()
    
    Dim i As Long
    Dim CCollClone As CColBarcodeInfo
    Dim oValues As CBarcodeInfo
    
    Set CCollClone = New CColBarcodeInfo
    
    For i = 1 To Me.count
    
        Set oValues = CCollClone.Add
        With CCollClone.Item(i)
            .BarcodeFields = Me.Item(i).BarcodeFields
            .BarcodeFValue = Me.Item(i).BarcodeFValue
            .BarcodeDesc = Me.Item(i).BarcodeDesc
        End With
                                                                                 
    Next i
    
    Set Clone = CCollClone
    
End Function
