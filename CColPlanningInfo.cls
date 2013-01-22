VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColPlanningInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object holds information about a PDR/PRG and its link.</summary>
'</comment>

Option Explicit

Private mcol As Collection

Public Function Add(Prgbarcode As String, formIdentification As String, filelinkid As Long, jobLogId As Long, _
                    codingNum As Integer) As CPlanningInfo

Dim objNewMember As CPlanningInfo
Set objNewMember = New CPlanningInfo

objNewMember.Form_Idectification = formIdentification
objNewMember.File_Links_ID = filelinkid
objNewMember.Job_Log_Id = jobLogId
objNewMember.Coding_Number = codingNum


mcol.Add objNewMember

'return the object created
Set Add = objNewMember
Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As CPlanningInfo
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
