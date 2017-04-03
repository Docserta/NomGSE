VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ResCheck3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Check3D

Private mvarId As String
Private mvarLibel As String
Private mvarResult As String
Private mvarLig As String

Public Property Let Id(ByVal vData As String)
    mvarId = vData
End Property

Public Property Get Id() As String
    Id = mvarId
End Property

Public Property Let Libel(ByVal vData As String)
    mvarLibel = vData
End Property

Public Property Get Libel() As String
    Libel = mvarLibel
End Property

Public Property Let Result(ByVal vData As String)
    mvarResult = vData
End Property

Public Property Get Result() As String
    Result = mvarResult
End Property

Public Property Let Lig(ByVal vData As String)
    mvarLig = vData
End Property

Public Property Get Lig() As String
    Lig = mvarLig
End Property
