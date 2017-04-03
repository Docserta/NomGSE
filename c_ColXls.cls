VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ColXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_ColXls
'Collection des colonnes Excel des fichiers de rapports

Private mcol As Collection

Public Function Add(Id As String, _
        Optional Col As String = "", _
        Optional Nom As String = "", _
        Optional lCol As Double = 0) As c_ColXl
    
    Dim objNewMember As c_ColXl
    Set objNewMember = New c_ColXl
    
    objNewMember.Id = Id
    objNewMember.Col = Col
    objNewMember.Nom = Nom
    objNewMember.lCol = lCol

    mcol.Add objNewMember, Id

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_ColXl
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








