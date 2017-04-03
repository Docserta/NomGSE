VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ResCheck3Ds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_ResCheck3Ds
Private mcol As Collection

Public Function Add(Id As String, _
        Optional Libel As String = "", _
        Optional Result As String = "", _
        Optional Lig As String) As c_ResCheck3D
    
    Dim objNewMember As c_ResCheck3D
    Set objNewMember = New c_ResCheck3D
    
    objNewMember.Id = Id
    objNewMember.Libel = Libel
    objNewMember.Result = Result
    objNewMember.Lig = Lig

    mcol.Add objNewMember, Id

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_ResCheck3D
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub






