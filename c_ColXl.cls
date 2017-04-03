VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ColXl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_ColXl
'classe des colonnes excel des rapport de controle
Private mvarId As String
Private mvarNom As String
Private mvarCol As String
Private mvarlCol As Double

Public Property Let Id(ByVal vData As String)
    mvarId = vData
End Property

Public Property Get Id() As String
    Id = mvarId
End Property

Public Property Let Nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get Nom() As String
    Nom = mvarNom
End Property

Public Property Let Col(ByVal vData As String)
    mvarCol = vData
End Property

Public Property Get Col() As String
    Col = mvarCol
End Property

Public Property Let lCol(ByVal vData As Double)
    mvarlCol = vData
End Property

Public Property Get lCol() As Double
    lCol = mvarlCol
End Property





