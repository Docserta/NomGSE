VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Check2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : Check_2D
'*
'* Fonctions :  vérifie un certain nomre d'item sur les 2D
'*
'*
'* Version : 1
'* Création :  CFR
'* Modification : 02/10/15
'*
'**********************************************************************

Private CK_CollDocs As Documents
Private CK_drawDoc As DrawingDocument
Private CK_Calques As DrawingSheets

Private CK_CalqueVue As DrawingSheet
Private CK_DrawingVues As DrawingViews
Private CK_BackGroundVue As DrawingView, CK_DrawingVue As DrawingView
Private CK_CollTxt As DrawingTexts
Private CK_CollDim As DrawingDimensions
Private LiensVue As DrawingViewGenerativeLinks 'Collection des liens des vues
Private PremLien As String 'Premier lien de la vue
Private Doclie As Document
'Private DocProdLie As Document
Private ProdLie As Product
Private CK_Tables As DrawingTables
Private CK_TableauNom2D As DrawingTable, CK_TableauTitres As DrawingTable

Private mvarVueLock As Boolean

Private paramsGSe As Parameters

'
'===========================
' Initialisation de l'objet
'===========================
'
'Private Sub Class_initialize()
'    On Error Resume Next
'End Sub

Property Let CK_OuvreDraw(CheminPlusNomFichier As String)
    'Réinitialisation de l'objet
    
    Set CK_CollDocs = CATIA.Documents
    
    'Ouvre le fichier Catdrawing
    Set CK_drawDoc = CK_CollDocs.Open(CheminPlusNomFichier)
    Set CK_Calques = CK_drawDoc.Sheets
    Set CK_CalqueVue = CK_Calques.activeSheet
    
    'Les calques
    Set CK_DrawingVues = CK_CalqueVue.Views
    Set CK_BackGroundVue = CK_DrawingVues.Item("Background View")
    Set CK_DrawingVue = CK_DrawingVues.Item("Main View")
    
    On Error GoTo Erreur1
    PremLien = CK_Lien3D(CK_DrawingVues)
    If PremLien <> "" Then
        Set Doclie = CK_CollDocs.Open(PremLien)
        On Error GoTo 0
        'Charge le 3D lié, teste s'il s'agit d'un Part ou d'un Product
        If InStr(PremLien, ".CATPart") <> 0 Then
            Set ProdLie = Doclie.GetItem(1)
            Set ProdLie = ProdLie.ReferenceProduct
        ElseIf InStr(PremLien, ".CATProduct") <> 0 Then
            Set ProdLie = Doclie.Product.ReferenceProduct
        End If
        'Paramètres
        Set paramsGSe = ProdLie.UserRefProperties
    End If
    GoTo Fin:
    
Erreur1:
    Err.Raise vbObjectError + 515, "Check2D", "Impossible d'ouvrir le 3D de référence"
    GoTo Fin
Fin:
End Property

Private Function CK_Lien3D(Vues As DrawingViews) As String
'Renvoi le lien du 3D lié au catdrawing
'Recherche dans la première vue puis dans les suivantes
'si la première vue pointe sur un corps de pièce le nom du 3D de référence ne remonte pas
'certain plan avec des vues iso ne remonte pas le fullname
Dim Liens As DrawingViewGenerativeLinks
Dim Lien As AnyObject
Dim vue As DrawingView
Dim VueLock As Boolean
Dim NomLien As String
Dim NomFullLien As String
Dim ParentDoc As Document
Dim i As Long

    For i = 3 To Vues.Count 'On saute les 2 première vues (MainView et Background View)
        On Error Resume Next
        Set vue = Vues.Item(i)
        VueLock = vue.LockStatus 'mémorise l'etat d'origine de la vue
        vue.LockStatus = False ' Déblocage de la vue pour accéder aux liens
        Set Liens = vue.GenerativeLinks
        Set Lien = Liens.FirstLink.Parent
        NomFullLien = Lien.FullName
        NomLien = Lien.Name
        If Err.Number <> 0 Then 'Erreur sur le fullname
            Err.Clear
        Else
            Exit For
        End If
        vue.LockStatus = VueLock 'Rétabli l'etat d'origine de la vue
    Next i
    If NomFullLien = "" And NomLien <> "" Then
    'on a récupéré un Name mais pas de FullName on va chercher le document correspondant au name
        On Error Resume Next
        For Each ParentDoc In CK_CollDocs
            If InStr(1, ParentDoc.Name, NomLien, vbTextCompare) > 0 Then
                NomFullLien = ParentDoc.FullName
                Exit For
            End If
        Next
    End If
    CK_Lien3D = NomFullLien
    If CK_Lien3D = "" Then
        Err.Raise vbObjectError + 514, "Check2D", "Liens de la vue non reconnus ou brisés"
    End If
    
    Set ParentDoc = Nothing
    'Set CK_CollDocs = Nothing
End Function

Property Get CK_Lien3DName() As String
'renvoi le nom du fichier 3D lié à la première vue du plan
Dim Pt As String
Dim Slh As String
    If PremLien = "" Then
        CK_Lien3DName = PremLien
    Else
        Pt = InStr(1, PremLien, ".", vbTextCompare)
        If Pt > 0 Then
            CK_Lien3DName = Left(PremLien, Pt - 1)
        End If
        Slh = InStr(1, CK_Lien3DName, "\", vbTextCompare)
        While Slh > 0
            Slh = InStr(1, CK_Lien3DName, "\", vbTextCompare)
            CK_Lien3DName = Right(CK_Lien3DName, Len(CK_Lien3DName) - Slh)
        Wend
    End If
    
End Property

Property Get Val_Matiere_PartLie()
'Renvoi le paramètre Material du part lié
    On Error Resume Next
    Dim DocLie_Part As Part
    Set DocLie_Part = Doclie.Part
    Val_Matiere_PartLie = DocLie_Part.Parameters.GetItem("Material").Value
    If Err.Number <> 0 Then
        Err.Clear
        Val_Matiere_PartLie = ""
    End If
    On Error GoTo 0
End Property

Property Get Val_Masse() As Double
'Renvoi la masse du part lié
    On Error Resume Next
    Dim Doclié_Prod As Product
    Set DocLie_Prod = Doclie.Product
    Val_Masse = Round(DocLie_Prod.Analyze.Mass, 3)

End Property

Public Function CK_Masse5P(CKM_masse As String) As Boolean
'Renvoi true si la valeur passée en argument ne dépasse pas de  +/- 5% la valeur de la masse
On Error Resume Next
temp_masse = CDbl(CKM_masse)
If Err.Number <> 0 Then
    Err.Clear
    CK_Masse5P = False
Else
    If temp_masse / Val_Masse > 1.05 Or temp_masse / Val_Masse < 0.95 Then
        CK_Masse5P = True
    Else
        CK_Masse5P = False
    End If
End If
End Function

'################## Nommage des fichiers

    Property Get Val_NumDrawing()
    'Renvoi le numero du drawing
    Val_NumDrawing = CK_drawDoc.Name
    End Property
    
    Property Get CK_Numerotation() As String
    'Check si la numérotation est correcte
    Dim NumTemp As String
    CK_Numerotation = "KO"
        's'il y a un slash c'est la planche générale
        If CK_Prem_Planche Then
            NumTemp = CK_Cart_NumeroPlan & "-" & Left(CK_Cart_Sheet, InStr(CK_Cart_Sheet, "/") - 1) & ".CATDrawing"
        Else
            NumTemp = CK_Cart_NumeroPlan & "-" & CK_Cart_Sheet & ".CATDrawing"
        End If
        If NumTemp = Val_NumDrawing Then
            CK_Numerotation = "OK"
        End If
    End Property

    Property Get Val_ItemNbNomFich()
    'Renvoi les 3 derniers carractère du nom du fichier 3D. ce qui doit correspondre a l'ItemNumber
    Dim TempItem As String
    If InStr(PremLien, ".CATPart") > 0 Then
        TempItem = Left(PremLien, InStr(PremLien, ".CATPart") - 1)
        TempItem = Right(TempItem, 3)
    End If
    Val_ItemNbNomFich = TempItem
    End Property
    
    Property Get Val_Num3D()
    'Renvoi le PartNumber du 3d de référence
        Val_Num3D = Doclie.Product.PartNumber
    End Property
'################## FIN Nommage des fichiers

'################## Gestion des vues

    Property Get CK_WorkingVueActive() As Boolean
    'Check si la vue active est dans l'environnement Active Vue
    CK_WorkingVueActive = False
        If CK_DrawingVues.ActiveView.Name <> "Background View" Then
            CK_WorkingVueActive = True
        End If
    End Property

    Property Get CK_LockVue() As Boolean
    'Check si toutes les vues sont bloquées
    CK_LockVue = True
        Dim i As Long
        'If PremVueLock Then
            For i = 4 To CK_DrawingVues.Count
                If CK_DrawingVues.Item(i).LockStatus Then
                Else
                    CK_LockVue = False
                    Exit For
                End If
            Next
        'Else
        '    CK_LockVue = False
        'End If
    End Property
    
    Property Get CK_CadreVue() As Boolean
    'Check si toutes les cadres de vues sont masqués
    CK_CadreVue = True
        Dim i As Long
        For i = 3 To CK_DrawingVues.Count
        'debug.print CK_DrawingVues.Item(i).Name
        'debug.print CK_DrawingVues.Item(i).FrameVisualization
            If CK_DrawingVues.Item(i).FrameVisualization Then
                CK_CadreVue = False
                Exit For
            End If
        Next
    End Property

    Property Get CK_VueIso() As Boolean
    'Check si toutesfalse"KO"
        Dim i As Long
        For i = 1 To CK_DrawingVues.Count
            If CK_DrawingVues.Item(i).ViewType = 8 Then
                CK_VueIso = True
                Exit For
            End If
        Next
    End Property

'################## FIN Gestion des vues

'#############################################################
'# Les Cotes et bullage
'#############################################################
    Property Get Val_List_Bulle() As String()
    'renvoi la liste des bulles détectée dans le plan
    Dim VueActive As DrawingView
    Dim Listebulles() As String
    ReDim Listebulles(0)
    Dim NbRep As Long
        NbRep = 1
    Dim TexteActif As DrawingText
    On Error Resume Next
    For Each VueActive In CK_DrawingVues
        If Err.Number <> 0 Then
            Err.Clear
            Listebulles(0) = "KO"
        Else
            On Error GoTo 0
            ' on passe les 2 premières vues
            If Not (VueActive.Name = "Background View") And Not (VueActive.Name = "Main View") Then
            'pour chaque texte de la vue
                Set CK_CollTxt = VueActive.Texts
                If CK_CollTxt.Count > 0 Then
                    For Each TexteActif In CK_CollTxt
                        If InStr(1, CStr(TexteActif.Name), "Numéro de pièce.") > 0 Or InStr(1, CStr(TexteActif.Name), "Balloon.") > 0 Then
                            ReDim Preserve Listebulles(NbRep)
                            Listebulles(NbRep) = TexteActif.Text
                            NbRep = NbRep + 1
                        End If
                    Next
                End If
            End If
        End If
    Next
    Val_List_Bulle = Listebulles()
    End Property

    Property Get CK_Bullage6() As String()
    'Check si tous les bullages ont une taille de 6
    ' et si la ligne de rappel est une fleche
    Dim i As Integer, j As Integer
    Dim VueActive As DrawingView
    Dim ListeRep() As String
    ReDim ListeRep(2, 0)
    ListeRep(0, 0) = "OK"
    Dim NbRep As Long
        NbRep = 1
    Dim TexteActif As DrawingText
    On Error Resume Next
    Debug.Print CK_DrawingVues.Name
    'For i = 1 To CK_DrawingVues.Count
    For Each VueActive In CK_DrawingVues
    'Set VueActive = CK_DrawingVues.Item(i)
    Debug.Print "nom vue : " & VueActive.Name
        If Err.Number <> 0 Then
            Err.Clear
            ListeRep(0, 0) = "KO"
        Else
            On Error GoTo Err_CK_Bullage6
            ' on passe les 2 premières vues
            If Not (VueActive.Name = "Background View") And Not (VueActive.Name = "Main View") Then
            'pour chaque texte de la vue
                Set CK_CollTxt = VueActive.Texts
                If CK_CollTxt.Count > 0 Then
                    'For j = 1 To CK_CollTxt.Count
                    For Each TexteActif In CK_CollTxt
                    'Set TexteActif = CK_CollTxt.Item(j)
                        If InStr(1, CStr(TexteActif.Name), "Numéro de pièce.") > 0 Or InStr(1, CStr(TexteActif.Name), "Balloon.") > 0 Then
                            Debug.Print "Nom texte : " & TexteActif.Name & "Nbrep : " & NbRep
                            If TexteActif.GetFontSize(0, 0) <> 6 Then
                                ListeRep(0, 0) = "KO"
                                ReDim Preserve ListeRep(2, NbRep)
                                ListeRep(0, NbRep) = "N° Bulle : " & TexteActif.Text
                                ListeRep(1, NbRep) = "Nom vue : " & VueActive.Name
                                ListeRep(2, NbRep) = "Taille texte = " & TexteActif.GetFontSize(0, 0)
                                NbRep = NbRep + 1
                            End If
                            On Error Resume Next
                            Temp_Leader = TexteActif.Leaders.Item(1).HeadSymbol
                            Debug.Print "Nom texte : " & TexteActif.Name & "Nbrep : " & NbRep & "Type leader : " & Temp_Leader
                            If Err.Number <> 0 Then
                                Err.Clear
                                Debug.Print "Nom texte : " & TexteActif.Name & "Nbrep : " & NbRep & " Pas de leader : : " & Temp_Leader & " Texte : " & TexteActif.Text
                            Else
                                On Error GoTo Err_CK_Bullage6
                                If Temp_Leader <> 17 Then
                                'If TexteActif.Leaders.Item(1).HeadSymbol <> 17 Then
                                    ListeRep(0, 0) = "KO"
                                    Debug.Print "Nom texte : " & TexteActif.Name & "Nbrep : " & NbRep & " Type Leader : " & Temp_Leader & " Texte : " & TexteActif.Text
                                    ReDim Preserve ListeRep(2, NbRep)
                                    ListeRep(0, NbRep) = "N° Bulle : " & TexteActif.Text
                                    ListeRep(1, NbRep) = "Nom vue : " & VueActive.Name
                                    ListeRep(2, NbRep) = "Type de flèche incorrect : " & Temp_Leader
                                    NbRep = NbRep + 1
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
GoTo fin_CK_Bullage6
    
Err_CK_Bullage6:
    ListeRep(0, 0) = "KO"
    ListeRep(1, 0) = Err.Number
    ListeRep(2, 0) = Err.Description
    
fin_CK_Bullage6:
    CK_Bullage6 = ListeRep()
    End Property

Property Get CK_CoteCasse() As String()
'renvois la liste des Cotes "cassée"
    Dim i As Integer, j As Integer
    Dim VueActive As DrawingView
    Dim ListeCotes() As String
    ReDim ListeCotes(2, 0)
    Dim CoteActive As DrawingDimension
    ListeCotes(0, 0) = "OK"
        Dim NbCote As Long
        NbCote = 1
    For Each VueActive In CK_DrawingVues
        If Err.Number <> 0 Then
            Err.Clear
            ListeCotes(0, 0) = "KO"
        Else
            ' on passe les 2 premières vues
            If Not (VueActive.Name = "Background View") And Not (VueActive.Name = "Main View") Then
                'pour chaque Cotes de la vue
                Set CK_CollDim = VueActive.Dimensions
                If CK_CollDim.Count > 0 Then
                    'For j = 1 To CK_CollTxt.Count
                    For Each CoteActive In CK_CollDim
                        If CoteActive.DimStatus = catUnUpdatableDim Or CoteActive.DimStatus = catFakeDim Then
                            ListeCotes(0, 0) = "KO"
                            ReDim Preserve ListeCotes(2, NbCote)
                            ListeCotes(0, NbCote) = "Cote : " & CoteActive.Name
                            ListeCotes(1, NbCote) = "Vue :  " & VueActive.Name
                            If CoteActive.DimStatus = catUnUpdatableDim Then
                                ListeCotes(2, NbCote) = "Cote Brisée"
                            ElseIf CoteActive.DimStatus = catFakeDim Then
                                ListeCotes(2, NbCote) = "Fausse cote"
                            End If
                            NbCote = NbCote + 1
                        End If
                    Next
                End If
            End If
  
        End If
    Next
 CK_CoteCasse = ListeCotes()
End Property


'################## FIN Gestion des Cotes et bullage

'#############################################################
'# Les Info dans le cartouche
'#############################################################

    Public Function TestTxtExist(TTE_NomTxt As String) As Boolean
    'test si le texte passé en argument existe dans le Cartouche.
    'si oui renvoi true sinon False,
    On Error Resume Next
        Dim TTE_drawTexte As DrawingText
        Set TTE_drawTexte = CK_BackGroundVue.Texts.GetItem(TTE_NomTxt)
        If (Err.Number <> 0) Then
            Err.Clear
            TestTxtExist = False
        Else
            TestTxtExist = True
        End If
    On Error GoTo 0
    End Function
    
    Public Function LectureTxt(LT_NomTxt As String) As String
    'Renvoi le contenu du Texte passé en argument dans le cartouche.
    'Si le paramètre n'existe pas renvois une chaine vide
        Dim temptxt As DrawingText
        If TestTxtExist(LT_NomTxt) Then
            LectureTxt = CK_BackGroundVue.Texts.GetItem(LT_NomTxt).Text
        Else
            LectureTxt = ""
        End If
    End Function
    
    Property Get CK_TextesCartouche() As DrawingTexts
    'Renvoi la collection des textes du calque de fond
        Set CK_TextesCartouche = CK_BackGroundVue.Texts
    End Property
    
    Property Get CK_Cart_LimitNotStaded()
    'Renvoi le contenu du texte Texte.LimitNotStaded
    If TestTxtExist("Texte.LimitNotStaded") Then
        CK_Cart_LimitNotStaded = LectureTxt("Texte.LimitNotStaded")
    Else
        CK_Cart_LimitNotStaded = ""
    End If
    End Property

    Property Get Val_Cart_SurfFinish()
    'Renvoi le contenu du texte Texte.SurfFinish
    If TestTxtExist("Texte.SurfFinish") Then
        Val_Cart_SurfFinish = LectureTxt("Texte.SurfFinish")
    Else
        Val_Cart_SurfFinish = ""
    End If
    End Property

    Property Get CK_Cart_Division()
    'Renvoi le contenu du texte Texte.Division
    If TestTxtExist("Texte.Division") Then
        CK_Cart_Division = LectureTxt("Texte.Division")
    Else
        CK_Cart_Division = ""
    End If
    End Property
    
    Property Get Val_Cart_Title()
    'Renvoi le contenu du text texte.title
    If TestTxtExist("Texte.Title") Then
        Val_Cart_Title = LectureTxt("Texte.Title")
    Else
        Val_Cart_Title = ""
    End If
    End Property

    Property Get CK_Cart_NumeroPlan()
    'Renvoi le contenu du text texte.NumeroPlan
    If TestTxtExist("Texte.NumeroPlan") Then
        CK_Cart_NumeroPlan = LectureTxt("Texte.NumeroPlan")
    Else
        CK_Cart_NumeroPlan = ""
    End If
    End Property

    Property Get CK_Cart_Sheet()
    'Renvoi le contenu du texte Text.sheet
    If TestTxtExist("Texte.sheet") Then
        CK_Cart_Sheet = LectureTxt("Texte.sheet")
    Else
        CK_Cart_Sheet = ""
    End If
    
    End Property
    
    Property Get Val_Cart_Echelle()
       'Renvoi le contenu du texte "Texte.Scale"
    If TestTxtExist("Texte.Scale") Then
        Val_Cart_Echelle = LectureTxt("Texte.Scale")
    Else
        Val_Cart_Echelle = "pas de valeur dans la case echelle du plan"
    End If
        
    End Property
    
    
    Private Function Conv_Cart_Echelle() As Double
    'Converti la valeur d'echelle de la vue en valeur de vartouche EX "1/1" dans la vue = "1" dans le cartouche
 
    Dim temp_echelle As String
    If TestTxtExist("Texte.Scale") Then
        temp_echelle = LectureTxt("Texte.Scale")
        Select Case temp_echelle
            Case "1/1"
                Conv_Cart_Echelle = 1
            Case "1/2"
                Conv_Cart_Echelle = 0.5
            Case "1/3"
                conv_cart_echalle = 0.3
            Case "2/5"
                Conv_Cart_Echelle = 0.4
            Case "1/5"
                Conv_Cart_Echelle = 0.2
            Case "1/10"
                Conv_Cart_Echelle = 0.1
            Case "2/1"
                Conv_Cart_Echelle = 2
            Case "3/1"
                conv_cart_echalle = 3
            Case "5/2"
                Conv_Cart_Echelle = 5
            Case "10/1"
                Conv_Cart_Echelle = 10
        End Select
    Else
        Conv_Cart_Echelle = temp_echelle
    End If
    End Function
    
    Property Get CK_Echelle()
    'Vérifie qu'au moins une vue est a l'echelle spécifié dans le plan
    Dim Temp_DrawingVue As DrawingView
    Dim i As Long
    
    CK_Echelle = False
    For i = 3 To CK_DrawingVues.Count
    'For Each Temp_DrawingVue In CK_DrawingVues
        Set Temp_DrawingVue = CK_DrawingVues.Item(i)
        If Temp_DrawingVue.Scale = Conv_Cart_Echelle Then
            CK_Echelle = True
        End If
    Next
    End Property
    
    Property Get CK_Prem_Planche()
    'Vérifie si c'est la première planche (N° pl / Nb Planche)
    CK_Prem_Planche = False
    If TestTxtExist("Texte.sheet") Then
        If InStr(LectureTxt("Texte.sheet"), "/") > 0 Then
            CK_Prem_Planche = True
        End If
    End If
    End Property
    
    Property Get Val_Cart_TotPlanches()
    'Renvoi le nombre de planches total
    If CK_Prem_Planche Then
            Val_Cart_TotPlanches = Right(LectureTxt("Texte.sheet"), Len(LectureTxt("Texte.sheet")) - InStr(LectureTxt("Texte.sheet"), "/"))
    Else
            Val_Cart_TotPlanches = ""
    End If
    End Property
    
    Property Get CK_Cart_DRN()
    'Renvois le contnu du texte "Texte.DRN"
    If TestTxtExist("Texte.DRN") Then
        CK_Cart_DRN = LectureTxt("Texte.DRN")
    Else
        CK_Cart_DRN = ""
    End If
    End Property
    
    Property Get CK_Cart_DOOrig()
    'Renvois le contnu du texte "Texte.DOOrig"
    If TestTxtExist("Text.DOOrig") Then
        CK_Cart_DOOrig = LectureTxt("Text.DOOrig")
    Else
        CK_Cart_DOOrig = ""
    End If
    End Property
    
    Property Get CK_Cart_DWGSys()
    'Renvois le contnu du texte "Texte.DwgSyst
    If TestTxtExist("Texte.DwgSyst") Then
        CK_Cart_DWGSys = LectureTxt("Texte.DwgSyst")
    Else
        CK_Cart_DWGSys = ""
    End If
    End Property
    
    Property Get CK_Cart_Process()
    'Renvois le contnu du texte "Texte.Process"
    If TestTxtExist("Texte.Process") Then
        CK_Cart_Process = LectureTxt("Texte.Process")
    Else
        CK_Cart_Process = ""
    End If
    End Property
    
    Property Get CK_Cart_CHK()
    'Renvois le contnu du texte "Texte.CHK"
    If TestTxtExist("Texte.CHK") Then
        CK_Cart_CHK = LectureTxt("Texte.CHK")
    Else
        CK_Cart_CHK = ""
    End If
    End Property

    Property Get Val_Cart_Issue()
    'Renvois le contnu du texte "Texte.Issue"
    If TestTxtExist("Texte.Issue") Then
        Val_Cart_Issue = LectureTxt("Texte.Issue")
    Else
        Val_Cart_Issue = ""
    End If
    End Property

    'Indice de modif
     Property Get CK_Mod_Issue()
    'Renvois le contnu du texte "Modifications_Issue"
    If TestTxtExist("Modifications_Issue") Then
        CK_Mod_Issue = LectureTxt("Modifications_Issue")
    Else
        CK_Mod_Issue = ""
    End If
    End Property

     Property Get CK_Mod_Date() As String
    'Renvois le contnu du texte "Texte.Datemod"
    If TestTxtExist("Texte.DateMod") Then
        CK_Mod_Date = LectureTxt("Texte.DateMod")
    Else
        CK_Mod_Date = ""
    End If
    End Property

     Property Get CK_Mod_Design()
    'Renvois le contnu du texte "Modification_Designer"
    If TestTxtExist("Modification_Designer") Then
        CK_Mod_Design = LectureTxt("Modification_Designer")
    Else
        CK_Mod_Design = ""
    End If
    End Property

    Property Get CK_Cart_Format()
    'Renvoi le contenu du texte "Texte.Size"
    If TestTxtExist("Texte.Size") Then
        CK_Cart_Format = LectureTxt("Texte.Size")
    Else
        CK_Cart_Format = ""
    End If
    End Property

    Property Get CK_Format_Plan()
    'Renvoi le format du plan
    CK_Format_Plan = CK_CalqueVue.PaperName
    End Property


'################## FIN Info dans le cartouche

'#############################################################
'# Les Paramètres 3D
'#############################################################
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
    
'    Public Sub EcritParam(EP_NomParam As String, EP_ValeurParam As String)
'    'test si le paramètre passé en argument existe dans le part.
'    'si oui si oui lui affecte la valeur passée en argument,
'    'sinon le crée et lui affecte la valeur passée en argument
'        Dim EP_ParamTemp As StrParam
'        If TestParamExist(EP_NomParam) Then
'            ParamsGSE.Item(EP_NomParam).Value = EP_ValeurParam
'        Else
'            Set EP_ParamTemp = ParamsGSE.CreateString(EP_NomParam, EP_ValeurParam)
'        End If
'    End Sub
    
    Public Function LectureParam(LP_NomParam As String) As String
    'Renvoi le contenu du paramètre passé en argument dans le part de la grille.
    'Si le paramètre n'existe pas renvois une chaine vide
        If TestParamExist(LP_NomParam) Then
            LectureParam = paramsGSe.Item(LP_NomParam).Value
        Else
            LectureParam = ""
        End If
    End Function
    
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
    
    Property Get CK_NomPulsGSE_CHK()
    'Renvoi le contenu du paramètre "NomPulsGSE_CHK"
        CK_NomPulsGSE_CHK = LectureParam("NomPulsGSE_CHK")
    End Property
    
    Property Get CK_NomPulsGSE_DatePlan() As String
    'Renvoi le contenu du paramètre "NomPulsGSE_DatePlan"
        CK_NomPulsGSE_DatePlan = LectureParam("NomPulsGSE_DatePlan")
    End Property
    
    Property Get CK_NomPulsGSE_CE()
    'Renvoi le contenu du paramètre "NomPulsGSE_CE"
        CK_NomPulsGSE_CE = LectureParam("NomPulsGSE_CE")
    End Property
    
    Property Get Val_NomPulsGSE_PresUserGuide()
    'Renvoi le contenu du paramètre "NomPulsGSE_PresUserGuide"
        Val_NomPulsGSE_PresUserGuide = LectureParam("NomPulsGSE_PresUserGuide")
    End Property
    
    Property Get CK_NomPulsGSE_PresCaisse()
    'Renvoi le contenu du paramètre "NomPulsGSE_PresCaisse"
        CK_NomPulsGSE_PresCaisse = LectureParam("NomPulsGSE_PresCaisse")
    End Property
    
    Property Get CK_NomPulsGSE_NoCaisse()
    'Renvoi le contenu du paramètre "NomPulsGSE_NoCaisse"
        CK_NomPulsGSE_NoCaisse = LectureParam("NomPulsGSE_NoCaisse")
    End Property
    
    Property Get CK_NomPulsGSE_Sheet()
    'Renvoi le contenu du paramètre "NomPulsGSE_Sheet"
        CK_NomPulsGSE_Sheet = LectureParam("NomPulsGSE_Sheet")
    End Property
    
    Property Get CK_NomPulsGSE_ItemNb()
    'Renvoi le contenu du paramètre "NomPulsGSE_ItemNb"
        CK_NomPulsGSE_ItemNb = LectureParam("NomPulsGSE_ItemNb")
    End Property
    
    Property Get CK_NomPulsGSE_Dimension()
    'Renvoi le contenu du paramètre "NomPulsGSE_Dimension"
        CK_NomPulsGSE_Dimension = LectureParam("NomPulsGSE_Dimension")
    End Property
    
    Property Get Val_NomPulsGSE_Material()
    'Renvoi le contenu du paramètre "NomPulsGSE_Material"
        Val_NomPulsGSE_Material = LectureParam("NomPulsGSE_Material")
    End Property
    
    Property Get Val_NomPulsGSE_Protect()
    'Renvoi le contenu du paramètre "NomPulsGSE_Protect"
        Val_NomPulsGSE_Protect = LectureParam("NomPulsGSE_Protect")
    End Property
    
    Property Get CK_NomPulsGSE_Miscellanous()
    'Renvoi le contenu du paramètre "NomPulsGSE_Miscellanous"
        CK_NomPulsGSE_Miscellanous = LectureParam("NomPulsGSE_Miscellanous")
    End Property
    
    Property Get CK_NomPulsGSE_SupplierRef()
    'Renvoi le contenu du paramètre "NomPulsGSE_SupplierRef"
        CK_NomPulsGSE_SupplierRef = LectureParam("NomPulsGSE_SupplierRef")
    End Property
    
    Property Get CK_NomPulsGSE_Weight()
    'Renvoi le contenu du paramètre "NomPulsGSE_Weight"
        CK_NomPulsGSE_Weight = LectureParam("NomPulsGSE_Weight")
    End Property
    
    Property Get CK_NomPulsGSE_MecanoSoude()
    'Renvoi le contenu du paramètre "NomPulsGSE_MecanoSoude"
        CK_NomPulsGSE_MecanoSoude = LectureParam("NomPulsGSE_MecanoSoude")
    End Property
    
    Property Get CK_NomPulsGSE_TypeNum()
    'Renvoi le contenu du paramètre "NomPulsGSE_TypeNum"
        CK_NomPulsGSE_TypeNum = LectureParam("NomPulsGSE_TypeNum")
    End Property

'################## FIN Gestion de paramètres

'#############################
'Le tableau de la nomenclature
'#############################

Property Get CK_Exist_TabNom()
'Test si un tableau de nomenclature est présent dans le plan
 CK_Exist_TabNom = False
    For i = 1 To CK_BackGroundVue.Tables.Count
        If CK_BackGroundVue.Tables.Item(i).Name = "TableauNom" Then
            CK_Exist_TabNom = True
        End If
    Next i
End Property

Property Get CK_Table_Nom()
'Renvoi le tableau de nomenclature dans le drawing
    For i = 1 To CK_BackGroundVue.Tables.Count
        If CK_BackGroundVue.Tables.Item(i).Name = "TableauNom" Then
            Set CK_Table_Nom = CK_BackGroundVue.Tables.Item(i)
        End If
    Next i
End Property

Property Get CK_Table_Titre()
'Renvoi le tableau de nomenclature dans le drawing
    For i = 1 To CK_BackGroundVue.Tables.Count
        If CK_BackGroundVue.Tables.Item(i).Name = "TableauTitre" Then
            Set CK_Table_Titre = CK_BackGroundVue.Tables.Item(i)
        End If
    Next i
End Property

Property Get Table_Nom_Ens()
'Renvoi 'true' si c'est une nomenclature d'un ensemble
'la nomenclature d'un ensemble contient 10 colones
'celle d'un détail en contient 5
    If CK_Table_Nom.NumberOfColumns > 5 Then
        Table_Nom_Ens = True
    Else
        Table_Nom_Ens = False
    End If
End Property

Property Get Table_Nom_Det()
'Renvoi 'true' si c'est une nomenclature d'un détail
'la nomenclature d'un détail contient 5 colones
    If CK_Table_Nom.NumberOfColumns = 5 Then
        Table_Nom_Det = True
    Else
        Table_Nom_Det = False
    End If
End Property

Property Get Table_Nom_Ensligne1()
'Renvoi la première ligne de nomenclature pour les ensembles
    Table_Nom_Ensligne1 = CK_Table_Nom.NumberOfRows - 1
End Property

Property Get Prem_Col_TableNomEns()
'renvoi le N° de la première colonne a partir de laquelle on va trouver les attributs Shett, Item NB etc ..
'sur certain ensemble contenant des symétrique ou des variantes, une ou plusieurs colonnes sont ajoutée en tête
    Prem_Col_TableNomEns = 0
    If CK_Table_Nom.NumberOfColumns > 10 Then
        Prem_Col_TableNomEns = CK_Table_Nom.NumberOfColumns - 10
    End If
End Property

Property Get CK_Table_Sheet()
'renvoi le contenu de la cellule "Sheet"
    If Table_Nom_Ens Then
        CK_Table_Sheet = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 2 + Prem_Col_TableNomEns)
    Else
        CK_Table_Sheet = ""
    End If
End Property

Property Get CK_Table_ItemNb()
'renvoi le contenu de la cellule "ItemNb"
    If Table_Nom_Ens Then
        CK_Table_ItemNb = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 3 + Prem_Col_TableNomEns)
    Else
        CK_Table_ItemNb = ""
    End If
End Property

Property Get cK_Table_PArtNb()
'renvoi le contenu de la cellule "PartNumber"
    If Table_Nom_Ens Then
        cK_Table_PArtNb = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 4 + Prem_Col_TableNomEns)
    Else
        cK_Table_PArtNb = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 1)
    End If
End Property

Property Get CK_Table_Description()
'renvoi le contenu de la cellule "Description"
    If Table_Nom_Ens Then
        CK_Table_Description = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 5 + Prem_Col_TableNomEns)
    Else
        CK_Table_Description = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 2)
    End If
End Property

Property Get CK_Table_Dimension()
'renvoi le contenu de la cellule "Dimension"
    If Table_Nom_Ens Then
        CK_Table_Dimension = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 6 + Prem_Col_TableNomEns)
    Else
        CK_Table_Dimension = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 3)
    End If
End Property

Property Get CK_Table_Material()
'renvoi le contenu de la cellule "Material"
    If Table_Nom_Ens Then
        CK_Table_Material = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 7 + Prem_Col_TableNomEns)
    Else
        CK_Table_Material = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 4)
    End If
End Property

Property Get Val_Table_Protect()
'renvoi le contenu de la cellule "Protect"
    If Table_Nom_Ens Then
        Val_Table_Protect = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 8 + Prem_Col_TableNomEns)
    Else
        Val_Table_Protect = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 5)
    End If
End Property

Property Get CK_Table_Miscel()
'renvoi le contenu de la cellule "Miscelanous"
    If Table_Nom_Ens Then
        CK_Table_Miscel = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 9 + Prem_Col_TableNomEns)
    Else
        CK_Table_Miscel = ""
    End If
End Property

Property Get Val_Table_Weight()
'renvoi le contenu de la cellule "Weight"
    If Table_Nom_Ens Then
        Val_Table_Weight = CK_Table_Nom.GetCellString(Table_Nom_Ensligne1, 10 + Prem_Col_TableNomEns)
    Else
        Val_Table_Weight = ""
    End If
End Property

Property Get ValNum_table_weight() As Double
'renvoi la valeur numérique de la cellule "Weight"
Dim i As Integer
Dim temp_ValStr_table_weight As String
    For i = 1 To Len(Val_Table_Weight)
        If IsNumeric(Mid(Val_Table_Weight, i, 1)) Then
            temp_ValStr_table_weight = temp_ValStr_table_weight & Mid(Val_Table_Weight, i, 1)
        End If
    Next i
    On Error Resume Next
    ValNum_table_weight = CDbl(temp_ValStr_table_weight)
    If Err.Number <> 0 Then
        Err.Clear
            ValNum_table_weight = 0
        On Error GoTo 0
    End If
End Property

Property Get CK_Weight() As Boolean
'Check si la valeur de la masse du drawing est différente de =/- 5% ou > 25Kg de la masse réelle.
CK_Weight = False
If ValNum_table_weight > 25 Then
    CK_Weight = True
ElseIf (ValNum_table_weight / Val_Masse < 0.95) Or (ValNum_table_weight / Val_Masse > 1.05) Then
    CK_Weight = True
End If

End Property
Property Get Val_Gimp()
'Renvoi le Numéro de la Gimp
    If Table_Nom_Ens And CK_Table_Nom.GetCellString(1, 5 + Prem_Col_TableNomEns) = "USER GUIDE" Then
            Val_Gimp = CK_Table_Nom.GetCellString(1, 4 + Prem_Col_TableNomEns)
    Else
        Val_Gimp = ""
    End If
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    CATIA.DisplayFileAlerts = False
    Doclie.Close
    CK_drawDoc.Close
    CATIA.DisplayFileAlerts = True
    On Error GoTo 0

End Sub
