VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCOLsmpFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object is a Collection of CsmpFiles.</summary>
'</comment>

Option Explicit

Private mcol As Collection
' DW ECR12-008 update ship as Integer to ship as Long
Public Function Add(prodId As Long, typenum As Integer, _
                    smptype As String, ship As Long, qty As Long, _
                    description As String, file As String, sample_type_id As Long) As CsmpFiles

Dim objNewMember As CsmpFiles
Set objNewMember = New CsmpFiles

objNewMember.productionId = prodId
objNewMember.typeNumber = typenum
objNewMember.sampleType = smptype
objNewMember.shipTo = ship
objNewMember.quantity = qty
objNewMember.smpDescription = description
objNewMember.smpfileName = file
objNewMember.sample_type_id = sample_type_id

mcol.Add objNewMember

'return the object created
Set Add = objNewMember
Set objNewMember = Nothing

End Function
Public Property Get Item(vntIndexKey As Variant) As CsmpFiles
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



