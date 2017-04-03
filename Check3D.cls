VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Check3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : Check_3D
'*
'* Fonctions :  vérifie un certain nomre d'item sur les 3D
'*
'*
'* Version : 1
'* Création :  CFR
'* Modification : 10/11/15
'*
'**********************************************************************
Option Explicit

Private Doclie As Document
Private ProdLie As Product
Private Rad11Digt As String
Public NomInst As String
Private IsPart As Boolean, IsProduct As Boolean
Private PartBody As Body

Property Let Charge3D(DocEC)
    'Réinitialisation de l'objet
    Set Doclie = DocEC
    'Charge le 3D, teste s'il s'agit d'un Part ou d'un Product
    If InStr(Doclie.Name, ".CATPart") <> 0 Then
        Set ProdLie = Doclie.GetItem(1)
        Set ProdLie = ProdLie.ReferenceProduct
        IsPart = True
        IsProduct = False
    ElseIf InStr(Doclie.Name, ".CATProduct") <> 0 Then
        Set ProdLie = Doclie.Product.ReferenceProduct
        IsPart = False
        IsProduct = True
    Else
        IsPart = False
        IsProduct = False
        Err.Raise vbObjectError + 520, "Check3D", "Ce document n'est ni un Part ni un Product"
    End If

End Property

    Property Get PN()
    'Renvoi le PartNumber du 3d de référence
        PN = Doclie.Product.PartNumber
    End Property
    
    Property Get EstPart() As Boolean
    'test si le DocEC est un CATPart ou un CATProduct
        EstPart = IsPart
    End Property

    Property Get EstProduct() As Boolean
    'test si le DocEC est un CATPart ou un CATProduct
        EstProduct = IsProduct
    End Property
    
    Property Get PartEC() As Part
        'Renvoi le part du doclie
        Set PartEC = Doclie.Part
    End Property
    
    Property Get ProdEC() As Product
        'Renvoi le product du doclié
        Set ProdEC = Doclie.Product
    End Property
    
'#############################################################
'# Les PartBody
'#############################################################
    Property Get Coll_bodies() As Bodies
        'renvoi la collection de partbodies
        If Me.EstPart Then
            Set Coll_bodies = Me.PartEC.Bodies
        End If
    End Property
    
    Property Get Coll_Hbodies() As HybridBodies
        'Renvoi la collection des Hybridbodies
        If Me.EstPart Then
            Set Coll_Hbodies = Me.PartEC.HybridBodies
        End If
    End Property
    
    Property Get NbBodies() As Integer
        'Renvoi le nombre de part bodies
        If Me.EstPart Then
            NbBodies = Me.Coll_bodies.Count
        End If
    End Property
    
    Property Get Mbody() As Body
        'Renvoi le Corps de pièce principal
        If Me.EstPart Then
            Set Mbody = Me.PartEC.MainBody
        End If
    End Property
    
    Property Get InWorkObj()
        'Renvoi l'objet de travail
        If Me.EstPart Then
           Set InWorkObj = Me.PartEC.InWorkObject
        End If
    End Property
    
'#############################################################
'# Les Contraintes et relations
'#############################################################
    
    Property Get col_Consts()
        'Renvoi la collection des contraintes
        If Me.EstProduct Then
            Set col_Consts = ProdEC.Connections("CATIAConstraints")
        End If
    End Property
   
   Property Get Col_Relations()
    'Renvoi la collection de relations
    If Me.EstPart Then
        Set Col_Relations = PartEC.Relations
    End If
   End Property
   
'#############################################################
'# Les Paramètres 3D
'#############################################################
    Property Get paramsGSe() As Parameters
    'Renvoi la collection de paramètres GSE
        Set paramsGSe = ProdLie.UserRefProperties
    End Property
    
    Property Get ParamsPart() As Parameters
    'Renvoi la collection des paramètre du part ou du product
            Set ParamsPart = ProdLie.Parameters
    End Property
    
    Public Function TestParamExist(TPE_NomParam As String) As Boolean
    'test si le paramètre passé en argument existe dans le part.
    'si oui renvoi true sinon False,
    On Error Resume Next
        Dim TPE_ParamCherche As StrParam
        Set TPE_ParamCherche = paramsGSe.Item(TPE_NomParam)
        If (Err.Number <> 0) Then
            Err.Clear
            TestParamExist = False
        Else
            TestParamExist = True
        End If
    On Error GoTo 0
    End Function
        
    Public Function LectureParam(LP_NomParam As String) As String
    'Renvoi le contenu du paramètre passé en argument dans le part de la grille.
    'Si le paramètre n'existe pas renvois une chaine vide
        If TestParamExist(LP_NomParam) Then
            If paramsGSe.Item(LP_NomParam).Value <> "" Then
                LectureParam = paramsGSe.Item(LP_NomParam).Value
            Else
                LectureParam = "Vide"
            End If
        Else
            LectureParam = "Absent"
        End If
    End Function
    
    Property Get ListParamGSE() As Variant
    'Renvoi la liste des paramètre spécifiques GSE
    ListParamGSE = Array("NomPulsGSE_DesignOutillage", "NomPulsGSE_NoOutillage", "NomPulsGSE_SiteAB", "NomPulsGSE_CHK", "NomPulsGSE_DatePlan" _
                          , "NomPulsGSE_Sheet", "NomPulsGSE_ItemNb", "NomPulsGSE_Dimension", "NomPulsGSE_Material", "NomPulsGSE_Protect" _
                          , "NomPulsGSE_SupplierRef", "NomPulsGSE_Weight", "NomPulsGSE_Miscellanous", "NomPulsGSE_MecanoSoude", "NomPulsGSE_CE" _
                          , "NomPulsGSE_PresUserGuide", "NomPulsGSE_PresCaisse", "NomPulsGSE_NoCaisse")
    End Property
    
    Property Get Val_Description()
    'Renvoi le contenu de l'attribut "Description"
        Val_Description = ProdEC.DescriptionRef
    End Property
    
    Property Get Val_NomPulsGSE_DesignOutillage()
    'Renvoi le contenu du paramètre "NomPulsGSE_DesignOutillage"
        Val_NomPulsGSE_DesignOutillage = LectureParam("NomPulsGSE_DesignOutillage")
    End Property
    
    Property Get Val_NomPulsGSE_NoOutillage()
    'Renvoi le contenu du paramètre "NomPulsGSE_NoOutillage"
        Val_NomPulsGSE_NoOutillage = LectureParam("NomPulsGSE_NoOutillage")
    End Property
    
    Property Get Val_NomPulsGSE_SiteAB()
    'Renvoi le contenu du paramètre "NomPulsGSE_SiteAB"
        Val_NomPulsGSE_SiteAB = LectureParam("NomPulsGSE_SiteAB")
    End Property
    
    Property Get Val_NomPulsGSE_CHK()
    'Renvoi le contenu du paramètre "NomPulsGSE_CHK"
        Val_NomPulsGSE_CHK = LectureParam("NomPulsGSE_CHK")
    End Property
    
    Property Get Val_NomPulsGSE_DatePlan() As String
    'Renvoi le contenu du paramètre "NomPulsGSE_DatePlan"
        Val_NomPulsGSE_DatePlan = LectureParam("NomPulsGSE_DatePlan")
    End Property
    
    Property Get Val_NomPulsGSE_CE()
    'Renvoi le contenu du paramètre "NomPulsGSE_CE"
        Val_NomPulsGSE_CE = LectureParam("NomPulsGSE_CE")
    End Property
    
    Property Get Val_NomPulsGSE_PresUserGuide()
    'Renvoi le contenu du paramètre "NomPulsGSE_PresUserGuide"
        Val_NomPulsGSE_PresUserGuide = LectureParam("NomPulsGSE_PresUserGuide")
    End Property
    
    Property Get Val_NomPulsGSE_PresCaisse()
    'Renvoi le contenu du paramètre "NomPulsGSE_PresCaisse"
        Val_NomPulsGSE_PresCaisse = LectureParam("NomPulsGSE_PresCaisse")
    End Property
    
    Property Get Val_NomPulsGSE_NoCaisse()
    'Renvoi le contenu du paramètre "NomPulsGSE_NoCaisse"
        Val_NomPulsGSE_NoCaisse = LectureParam("NomPulsGSE_NoCaisse")
    End Property
    
    Property Get Val_NomPulsGSE_Sheet()
    'Renvoi le contenu du paramètre "NomPulsGSE_Sheet"
        Val_NomPulsGSE_Sheet = LectureParam("NomPulsGSE_Sheet")
    End Property
    
    Property Get Val_NomPulsGSE_ItemNb()
    'Renvoi le contenu du paramètre "NomPulsGSE_ItemNb"
        Val_NomPulsGSE_ItemNb = LectureParam("NomPulsGSE_ItemNb")
    End Property
    
    Property Get Val_NomPulsGSE_Dimension()
    'Renvoi le contenu du paramètre "NomPulsGSE_Dimension"
        Val_NomPulsGSE_Dimension = LectureParam("NomPulsGSE_Dimension")
    End Property
    
    Property Get Val_NomPulsGSE_Material()
    'Renvoi le contenu du paramètre "NomPulsGSE_Material"
        Val_NomPulsGSE_Material = LectureParam("NomPulsGSE_Material")
    End Property
    
    Property Get Val_NomPulsGSE_Protect()
    'Renvoi le contenu du paramètre "NomPulsGSE_Protect"
        Val_NomPulsGSE_Protect = LectureParam("NomPulsGSE_Protect")
    End Property
    
    Property Get Val_NomPulsGSE_Miscellanous()
    'Renvoi le contenu du paramètre "NomPulsGSE_Miscellanous"
        Val_NomPulsGSE_Miscellanous = LectureParam("NomPulsGSE_Miscellanous")
    End Property
    
    Property Get Val_NomPulsGSE_SupplierRef()
    'Renvoi le contenu du paramètre "NomPulsGSE_SupplierRef"
        Val_NomPulsGSE_SupplierRef = LectureParam("NomPulsGSE_SupplierRef")
    End Property
    
    Property Get Val_NomPulsGSE_Weight()
    'Renvoi le contenu du paramètre "NomPulsGSE_Weight"
        Val_NomPulsGSE_Weight = LectureParam("NomPulsGSE_Weight")
    End Property
    
    Property Get Val_NomPulsGSE_MecanoSoude()
    'Renvoi le contenu du paramètre "NomPulsGSE_MecanoSoude"
        Val_NomPulsGSE_MecanoSoude = LectureParam("NomPulsGSE_MecanoSoude")
    End Property
    
    Property Get Val_NomPulsGSE_TypeNum()
    'Renvoi le contenu du paramètre "NomPulsGSE_TypeNum"
        Val_NomPulsGSE_TypeNum = LectureParam("NomPulsGSE_TypeNum")
    End Property

Property Get partmat() As String
'Renvoi le contenu du paramètre "Matiere"
'Si plusieurs corps de pièces, vérifie chaque corps
Dim ParamMat As Parameter
Dim Mat As String, lstMat As String
Dim nBody As String
    On Error Resume Next
    For Each ParamMat In ParamsPart
        Mat = ""
        If InStr(1, ParamMat.Name, "\Material", vbTextCompare) > 0 Then
            nBody = ExtractNameBody(ParamMat.Name)
            If ParamMat.ValueAsString = "" Then
                Mat = nBody & "=Absent "
            Else
                Mat = nBody & "=" & ParamMat.ValueAsString & " "
            End If
        End If
        lstMat = lstMat & Mat
    Next
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        partmat = "Absent"
    Else
        partmat = lstMat
    End If
    End Property

'################## FIN Gestion de paramètres
    
    
'#############################################################
'# Les Checks 3D
'#############################################################

Public Function TestAllProperties() As Boolean
'test si toutes les propriétés GSE sont présentes
Dim i As Integer
    TestAllProperties = True
    'Dim exist As Boolean
    'exist = True
    For i = 1 To 13
        If Not (TestParamExist(CStr(ListParamGSE(i)))) Then
            TestAllProperties = False
            'exist = False
        End If
    Next i
    'TestAllProperties = False
End Function

Property Get CK_Properties() As Boolean
'Check si les propriétés GSE sont présentes
    CK_Properties = TestAllProperties

End Property

'################## La Numérotation des fichiers.

Public Function NomFic() As String
'renvois le nom du fichier sans l'extention
    NomFic = Left(Doclie.Name, InStr(1, Doclie.Name, ".", vbTextCompare) - 1)
End Function

Property Get CK_LgNomFic() As Boolean
' tests si le fichier à bien 14 carratères
    If Len(NomFic) = 14 Then
        CK_LgNomFic = True
    Else
        CK_LgNomFic = False
    End If
End Property

Property Let initRad11Digt(RadNomFic As String)
'Récupère le radical pour comparaison avec les noms de fichier,d'instance et de PartNumber
    Rad11Digt = Left(RadNomFic, 11)
End Property

Property Get CK_Radical() As Boolean
' test si les 11 premiers carratère du  PartNumber sont égal au radical
    If Left(Me.PN, 11) = Rad11Digt Then CK_Radical = True Else CK_Radical = False
End Property

Property Get CK_NumEgal() As Boolean
' Test si le partNumber et le n° du fichier sont egal
    If Me.PN = Me.NomFic Then
        CK_NumEgal = True
    End If
End Property

Property Get CK_Impaire() As Boolean
'test si les 3 dernier chiffres du PartNumber forment un nombre impaire
'et compris entre 100 et 499
    If Right(Me.PN, 3) > 100 And Right(Me.PN, 3) < 499 Then
        Dim DerDigit As String
        DerDigit = Right(Me.PN, 1)
        Select Case DerDigit
            Case 1, 3, 5, 7, 9
                CK_Impaire = True
            Case 0, 2, 4, 6, 8
                CK_Impaire = False
        End Select
    Else
        CK_Impaire = False
    End If
End Property

'################## Part Body

    Property Get CK_NbBodies() As Boolean
    'Test qu'il n'y a qu'un seul body pour les pièce entre 200 et 699
        Dim i As Integer
        CK_NbBodies = True
        If Me.NbBodies = 1 Then
            CK_NbBodies = True
        Else
            For i = 2 To Me.NbBodies
                If Me.Coll_bodies.Item(i).InBooleanOperation Then
                Else 'problème d'inversion de booleen avec fonction Not()
                    CK_NbBodies = False
                End If
            Next i
        End If
    End Property

    Property Get CK_InWorkObj() As Boolean
    'Test si le corps de pièce principal est actif
    'If Me.InWorkObj <> Me.Mbody Then
    If Me.InWorkObj.Name = Me.Mbody.Name Then
        CK_InWorkObj = True
    Else
        CK_InWorkObj = False
    End If
    End Property

    Property Get CK_EstHybride() As Boolean
    'test si le part est en conception hybride
    'Tente de créer un point dans le corp de pièce (l'efface ensuite)
    'et détecte l'erreur
    On Error Resume Next
        Dim HS_test As HybridShapeFactory
        Set HS_test = Me.PartEC.HybridShapeFactory
        Dim Pt_test As HybridShapePointCoord
        Set Pt_test = HS_test.AddNewPointCoord(0, 0, 0)
        Dim Sel_test As Selection
        Set Sel_test = Doclie.Selection
        If IsMissing(Mbody.HybridShapes.Count) Then
            CK_EstHybride = False
        Else
            If Mbody.HybridShapes.Count > 0 Then
                CK_EstHybride = True
            Else
                CK_EstHybride = False
            End If
        End If
        Me.Mbody.InsertHybridShape Pt_test
        If Err.Number <> 0 Then
            Err.Clear
            CK_EstHybride = False
        Else
            Sel_test.Add Pt_test
            Sel_test.Delete
            CK_EstHybride = True
        End If
    End Property

    Property Get Broken_relations() As Boolean
    'test si les relation contiennent le terme "delated_" dans leur valeur
    On Error Resume Next
    Dim Relat As Relation
        Broken_relations = False
    Dim Val_Reltn As Integer
    For Each Relat In Me.Col_Relations
        Val_Reltn = InStr(1, Relat.Value, "deleted_", vbTextCompare)
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If Val_Reltn > 0 Then
                Broken_relations = True
            End If
        End If
    Next
    End Property




