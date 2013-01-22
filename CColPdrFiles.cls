VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCOLPDRFILES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object is a Collection of CPRDFiles.</summary>
'</comment>

Option Explicit

Private mcol As Collection

Public Function Add(Field1 As String, count As Long, smptype As String, Optional sKey As String) As CPRDFiles
    
    'create a new object
    Dim objNewMember As CPRDFiles
    Dim temp As String
    
    Set objNewMember = New CPRDFiles

    'set the properties passed into the method
    'md added code to correct Samples when reading from a file. Replacement files were
    'not being displayed correctly in the Grid on the sample config. screen
    
    If (booReplacement) And (smptype <> "CTK") Then
        ' kbg 2008-009 added InsertBLANK param
       temp = GetDelimitedFirstLine(vdata, 1, gRandDelimiter, False)
       ' kbg 2008-009 added InsertBLANK param
       objNewMember.Field1 = smptype & "-" & GetDelimitedFirstLine(temp, 3, "-", False)
    Else
        objNewMember.Field1 = smptype & "-" & count
    End If
   
   ' kbg 2008-009 added InsertBLANK param
    objNewMember.Field2 = GetDelimitedFirstLine(vdata, 2, gRandDelimiter, True)
    objNewMember.Field3 = GetDelimitedFirstLine(vdata, 3, gRandDelimiter, True)
    objNewMember.Field4 = GetDelimitedFirstLine(vdata, 4, gRandDelimiter, True)
    objNewMember.Field5 = GetDelimitedFirstLine(vdata, 5, gRandDelimiter, True)
    objNewMember.Field6 = GetDelimitedFirstLine(vdata, 6, gRandDelimiter, True)
    objNewMember.Field7 = GetDelimitedFirstLine(vdata, 7, gRandDelimiter, True)
    objNewMember.Field8 = GetDelimitedFirstLine(vdata, 8, gRandDelimiter, True)
    objNewMember.Field9 = GetDelimitedFirstLine(vdata, 9, gRandDelimiter, True)
    objNewMember.Field10 = GetDelimitedFirstLine(vdata, 10, gRandDelimiter, True)
    objNewMember.Field11 = GetDelimitedFirstLine(vdata, 11, gRandDelimiter, True)
    objNewMember.Field12 = GetDelimitedFirstLine(vdata, 12, gRandDelimiter, True)
    objNewMember.Field13 = GetDelimitedFirstLine(vdata, 13, gRandDelimiter, True)
    objNewMember.Field14 = GetDelimitedFirstLine(vdata, 14, gRandDelimiter, True)
    objNewMember.Field15 = GetDelimitedFirstLine(vdata, 15, gRandDelimiter, True)
    objNewMember.Field16 = GetDelimitedFirstLine(vdata, 16, gRandDelimiter, True)
    objNewMember.Field17 = GetDelimitedFirstLine(vdata, 17, gRandDelimiter, True)
    objNewMember.Field18 = GetDelimitedFirstLine(vdata, 18, gRandDelimiter, True)
    objNewMember.Field19 = GetDelimitedFirstLine(vdata, 19, gRandDelimiter, True)
    objNewMember.Field20 = GetDelimitedFirstLine(vdata, 20, gRandDelimiter, True)
    ' DW increasing # of columns from 20 to 30 based on client supplied data
    objNewMember.Field21 = GetDelimitedFirstLine(vdata, 21, gRandDelimiter, True)
    objNewMember.Field22 = GetDelimitedFirstLine(vdata, 22, gRandDelimiter, True)
    objNewMember.Field23 = GetDelimitedFirstLine(vdata, 23, gRandDelimiter, True)
    objNewMember.Field24 = GetDelimitedFirstLine(vdata, 24, gRandDelimiter, True)
    objNewMember.Field25 = GetDelimitedFirstLine(vdata, 25, gRandDelimiter, True)
    objNewMember.Field26 = GetDelimitedFirstLine(vdata, 26, gRandDelimiter, True)
    objNewMember.Field27 = GetDelimitedFirstLine(vdata, 27, gRandDelimiter, True)
    objNewMember.Field28 = GetDelimitedFirstLine(vdata, 28, gRandDelimiter, True)
    objNewMember.Field29 = GetDelimitedFirstLine(vdata, 29, gRandDelimiter, True)
    objNewMember.Field30 = GetDelimitedFirstLine(vdata, 30, gRandDelimiter, True)
    
    If Len(sKey) = 0 Then
        mcol.Add objNewMember
    Else
        mcol.Add objNewMember, sKey
    End If
   
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As CPRDFiles
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

