Attribute VB_Name = "b_Nomenclature"
Option Explicit


Sub catmain()
' *****************************************************************
' * Création d'attributs pour génération de nomenclatures automatique
' * Construit une liste des parts contenu dans le product pointé par l'utilisateur
' * ouvre un boite de dialogue permettant de modifier ou de renseigner les paramètres pour chaques parts
' * puis met a jour les attibuts de chaque part avec les valeurs de la boite de dialogue
' * Création CFR le 24/10/2012
' * modification le : 18/09/14
' *    Ajout module de classe xMacroLocation
' * modification le : 28/10/14
' *    Prise en compte de 2 systemes de numérotation des achats 500 à 999 ou 700 à 900
' *****************************************************************
'On Error Resume Next
On Error GoTo err_Main

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "b_Nomenclature", VMacro

'Chargement des variables
Dim MacroLocation As New xMacroLocation
If Not (MacroLocation.FicIniExist("VarNomenclatureGSE.ini")) Then ' on est pas dans l'environnement GSE_R21
    MsgBox "Vous n'etes pas dans l'environnement GRE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
    Exit Sub
Else
    MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
    CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
    CheminDestNomenclature = MacroLocation.ValVar("CheminDestNomenclature")
End If
     
'tous les documents ouverts
Set Coll_Documents = CATIA.Documents
NbPartsNames = Coll_Documents.Count

'Ensemble Sélectionné
Dim DocumentSelectionne As Document
Set ActiveDoc = CATIA.ActiveDocument
Dim ActiveProd, ProductSelectionne As Product
Set ActiveProd = ActiveDoc.Product

Dim FichTxt As String
Dim i As Long, j As Long
Dim CrLf As String
    CrLf = Chr(13) & Chr(10)
 Dim TestErreur As String, Msg_Err As String
    
'  paramètres de l'ensemble sélectionné
Dim MesParametres As Parameters
'DesignOutillage   'affiché dans le formulaire  mais pas mis a jour dans les parts et products
'NoOutillage       'affiché dans le formulaire  mais pas mis a jour dans les parts et products
'NomPulsGSE_SiteAB
'NomPulsGSE_CHK
'NomPulsGSE_DatePlan
'NomPulsGSE_CE
'NomPulsGSE_PresUserGuide
'NomPulsGSE_PresCaisse
'NomPulsGSE_NoCaisse
'NomPulsGSE_Sheet
'NomPulsGSE_ItemNb
'NomPulsGSE_Dimension
'NomPulsGSE_Material
'NomPulsGSE_Protect
'NomPulsGSE_Miscellanous
'NomPulsGSE_SupplierRef
'NomPulsGSE_Weight
'NomPulsGSE_MecanoSoude
'NomPulsGSE_TypeNum

'Test si le product général est vide
Dim NbrParts As Integer
NbrParts = Coll_Documents.Count
If NbrParts = 0 Then
    MsgBox "Ce product est vide", vbOKOnly, "Erreur"
    Exit Sub
End If

'Récupération des Nom et Numéro d'outillage
Set MesParametres = ActiveProd.UserRefProperties
Dim TmpNoOut As String, TmpDesOut As String
TmpNoOut = ActiveProd.PartNumber
TmpDesOut = RecupParam(MesParametres, "NomPulsGSE_DesignOutillage")

'Type de Numérotation
'NomPulsGSE_TypeNum
TypeNum = RecupParam(MesParametres, "NomPulsGSE_TypeNum")

'Initialisation du tableau des paramètres
ReDim TableauPartsParam(NbParam, 0)

'Initialisation de la liste des parts
ReDim ListPartsNames(0)
ListPartsNames(0) = ""

'Demande à l'utilisateur de sélectionner un product
    Dim varfilter(0) As Variant
    Dim objSel As Selection
    Dim objSelLB As Object
    Dim strReturn As String
    Dim strMsg As String
    
    varfilter(0) = "Product"
    Set objSel = CATIA.ActiveDocument.Selection
    Set objSelLB = objSel
    MsgBox "Sélectionner un product dans l'arbre"
    strMsg = "Selection product"
    objSel.Clear
    strReturn = objSelLB.SelectElement2(varfilter, strMsg, False)
    Dim NomObjetSel, NumObjetSel As String
    'Nom de l'objet sélectionné dans l'arbre
    NomObjetSel = objSel.Item2(1).Value.Name
    'Partnumber de l'objet sélectionné dans l'arbre
    NumObjetSel = objSel.Item2(1).Value.PartNumber
    NumObjetSel = NumObjetSel & ".CATProduct"
    
    Set DocumentSelectionne = Coll_Documents.Item(CStr(NumObjetSel))
    Set ProductSelectionne = DocumentSelectionne.Product

'Chargement en mode conception
    ProductSelectionne.ApplyWorkMode DESIGN_MODE
    
'Chargement de la barre de progression
    Load Frm_Progression
    Frm_Progression.Show vbModeless
    Frm_Progression.Caption = " Récupération des paramètres. Veuillez patienter..."
    ProgressBar (1)
    CompteurLimiteBarre = Coll_Documents.Count

'Appel de la procedure de contitution de la liste des parts et recupération des attributs
    ListPartTableau ProductSelectionne
    ReDim Preserve TableauPartsParam(NbParam, UBound(TableauPartsParam, 2) - 1)
    Unload Frm_Progression
   
'Chargement de la boite de dialogue "attribut"
'Elle ne sera activée qu'après un DblClick sur une ligne de nomenclature
    Load Frm_Attributs
    
'Création objet fichier texte
    Dim fs, f, f_Export
    Set fs = CreateObject("scripting.filesystemobject")

'Remplissage de la liste Designation
    FichTxt = CheminSourcesMacro & List_Designation
    Set f = fs.opentextfile(FichTxt, ForReading, 1)
    Do While Not f.AtEndOfStream
        Frm_Attributs.Cbx_Designation.AddItem (f.ReadLine)
    Loop
'Remplissage de la liste Material
    FichTxt = CheminSourcesMacro & List_Material
    Set f = fs.opentextfile(FichTxt, ForReading, 1)
    Do While Not f.AtEndOfStream
        Frm_Attributs.Cbx_Material.AddItem (f.ReadLine)
    Loop
'Remplissage de la liste Protect
    FichTxt = CheminSourcesMacro & List_Protect
    Set f = fs.opentextfile(FichTxt, ForReading, 1)
    Do While Not f.AtEndOfStream
        Frm_Attributs.Cbx_Protect.AddItem (f.ReadLine)
    Loop
'Remplissage de la liste Miscellanous
    FichTxt = CheminSourcesMacro & List_Miscellanous
    Set f = fs.opentextfile(FichTxt, ForReading, 1)
    Do While Not f.AtEndOfStream
        Frm_Attributs.Cbx_Miscellanous.AddItem (f.ReadLine)
    Loop

'Remplissage de la liste Catalogue dans le formulaire Catalogue
    Dim Table_Catalogue() As String
    i = 0
    j = 0
    Dim ligneEC As String
    Dim PosSep1, PosSep2, PosSep3, PosSep4, PosSep5 As Integer 'Position des séparateurs de champs dans la liste
    FichTxt = CheminSourcesMacro & List_Catalogue
    Set f = fs.opentextfile(FichTxt, ForReading, 1)
    Do While Not f.AtEndOfStream
        ligneEC = f.ReadLine
        ReDim Preserve Table_Catalogue(3, i)
        PosSep1 = InStr(1, ligneEC, ";")
        PosSep2 = InStr(PosSep1 + 1, ligneEC, ";")
        PosSep3 = InStr(PosSep2 + 1, ligneEC, ";")
        Table_Catalogue(0, i) = Left(ligneEC, PosSep1 - 1)
        Table_Catalogue(1, i) = Mid(ligneEC, PosSep1 + 1, PosSep2 - PosSep1 - 1)
        Table_Catalogue(2, i) = Mid(ligneEC, PosSep2 + 1, PosSep3 - PosSep2 - 1)
       Table_Catalogue(3, i) = Right(ligneEC, Len(ligneEC) - PosSep3)
       i = i + 1
    Loop
    Load Frm_Catalogue
    Frm_Catalogue.LBx_Catalogue.ColumnCount = 4
    Frm_Catalogue.LBx_Catalogue.List = TriList2D(TranspositionTabl(Table_Catalogue), 1, True)

'Chargement de la boite de dialogue Nomenclature
    Load Frm_Nomenclature
    Frm_Nomenclature.LBx_Nomenclature.ColumnCount = NbParam
    Frm_Nomenclature.Tbx_Designation = TmpDesOut
    Frm_Nomenclature.Tbx_Reference = TmpNoOut
    Frm_Nomenclature.LBx_Nomenclature.List = TriList2D(TranspositionTabl(TableauPartsParam), 1, True)
    Frm_Nomenclature.Show
    
    If Not (Frm_Nomenclature.ChB_OkAnnule) Then Exit Sub

' Export des modifications vers fichier texte
' afin de pouvoir les récuperer si plantage "Automation error"
    Dim FichExport, TableLigne As String
    FichExport = CheminDestNomenclature & "Export_Attributs.txt"
    Set f_Export = fs.createtextfile(FichExport, True)
    For i = 0 To UBound(Frm_Nomenclature.LBx_Nomenclature.List, 1)
        For j = 0 To UBound(Frm_Nomenclature.LBx_Nomenclature.List, 2)
            TableLigne = TableLigne & Frm_Nomenclature.LBx_Nomenclature.List(i, j) & "|"
        Next j
        
        TableLigne = Replace(TableLigne, CrLf, " ")
        f_Export.Writeline (TableLigne)
    TableLigne = ""
    Next i
    f_Export.Close

'test l'erreur "Automation error"
    On Error Resume Next
        TestErreur = DocumentSelectionne.Name
    If (Err.Number <> 0) Then
        Err.Clear
        Msg_Err = "Une erreur s'est produite. les modifications apportée à la nomenclature sont sauvegardées dans : " & FichExport & Chr(13)
        Msg_Err = Msg_Err & "relancez la macro et cliquez sur le bouton d'import pour les récupérer."
        MsgBox Msg_Err, vbCritical, "Automation error"
        Exit Sub
    End If

'Mise à jours des attributs du 3D avec les modification apportées par l'utilisateur
    MajAttributs
    
'Déchargement des boites de dialogue
    Unload Frm_Progression
    Unload Frm_Attributs
    Unload Frm_Nomenclature
    
GoTo Quit_err_Main

err_Main:
MsgBox Err.Number & Err.Description

Quit_err_Main:
End Sub

Public Sub ListPartTableau(LPT_Products As Products)
' *****************************************************************
' * Construit la liste des parts et récupère leurs attributs
' * balaye l'ensembles des Item du Product outillage et test s'il s'agit d'un part ou d'un product
' * Pour les parts, récupére les attibuts
' * pour les products, appelle la procedure en recursif pour rechercher les "sous parts"
' * Création CFR le 05/11/2012
' * Dernière modification le
' *****************************************************************
On Error GoTo err_ListPartTableau
Dim LPT_Coll_documents As Documents
Set LPT_Coll_documents = CATIA.Documents
Dim LPT_Product As Product
'Nom de l'Item en cours de traitement dans l'arbre
Dim LPT_NomItemEC As String
Dim LPT_CompPart, LPT_CompProduct As Boolean
Dim LPT_Document As Document
Dim LPT_ListeAttribPartEC(10) As String
Dim i As Long
Dim PartExistinList As Boolean

'Paramètres
Dim LPT_MesParametres As Parameters
'DesignOutillage
'NoOutillage
'Site_Airbus
'CHK
'DatePlan
'Sheet         'NomPulsGSE_Sheet
'ItemNb        'NomPulsGSE_ItemNb
'Dimension     'NomPulsGSE_Dimension
'Material      'NomPulsGSE_Material
'Protect       'NomPulsGSE_Protect
'Miscellanous  'NomPulsGSE_Miscellanous
'SupplierRef   'NomPulsGSE_SupplierRef
'Weight        'NomPulsGSE_Weight
'MecanoSoude   'NomPulsGSE_MecanoSoude
'Type de Numérotation   'NomPulsGSE_TypeNum

'Nombre d'Item dans le product
Dim NbrPartsOut As Integer
    NbrPartsOut = LPT_Products.Count
For i = 1 To NbrPartsOut
    PartExistinList = False
    Set LPT_Product = LPT_Products.Item(i)
    'Vérification que le part n'a pas déja été traité
    'Cas des parts communs a plusieurs produits
    If ListPartUnique(LPT_Product.PartNumber, ListPartsNames) Then
        ProgressBar (100 / CompteurLimiteBarre * UBound(ListPartsNames(), 1))
        'Ajout a la liste des noms de parts
        ReDim Preserve ListPartsNames(UBound(ListPartsNames, 1) + 1)
        ListPartsNames(UBound(ListPartsNames(), 1)) = LPT_Product.PartNumber
        'test s'il s'agit d'un Part ou d'un Product
        LPT_CompPart = True
        LPT_CompProduct = True
        On Error Resume Next
        Err.Clear
        LPT_NomItemEC = CStr(LPT_Product.PartNumber) & ".CATProduct"
        Set LPT_Document = LPT_Coll_documents.Item(LPT_NomItemEC)
        If (Err.Number <> 0) Then
            'Le composant n'est pas un CATProduct
            Err.Clear
            LPT_CompProduct = False
        End If
        LPT_NomItemEC = CStr(LPT_Product.PartNumber) & ".CATPart"
        Set LPT_Document = LPT_Coll_documents.Item(LPT_NomItemEC)
        If (Err.Number <> 0) Then
            'Le composant n'est pas un CATPart
            Err.Clear
            LPT_CompPart = False
        End If
        'On Error GoTo err_ListPartTableau
        On Error Resume Next
        If LPT_CompProduct Then
            'Récupération des paramètres des Products dans le tableau
            Set LPT_MesParametres = LPT_Product.ReferenceProduct.UserRefProperties
            LPT_ListeAttribPartEC(0) = RecupParam(LPT_MesParametres, "NomPulsGSE_Sheet")                'Sheet
            'cas des outillages et variantes
            If TypeElement(LPT_Product.PartNumber, TypeNum) >= 1 And TypeElement(LPT_Product.PartNumber, TypeNum) <= 2 Then
                LPT_ListeAttribPartEC(1) = ""
            Else
                'LPT_ListeAttribPartEC(1) = TestParamExist(LPT_MesParametres, "NomPulsGSE_ItemNb")
                LPT_ListeAttribPartEC(1) = RecupParam(LPT_MesParametres, "NomPulsGSE_ItemNb")           'Item Nbr
            End If
            LPT_ListeAttribPartEC(10) = LPT_Product.PartNumber                                          'PartNumber
            LPT_ListeAttribPartEC(2) = LPT_Product.PartNumber                                           'SupplierRef
            LPT_ListeAttribPartEC(3) = LPT_Product.DescriptionRef
            'LPT_ListeAttribPartEC(4) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Dimension")
            LPT_ListeAttribPartEC(4) = RecupParam(LPT_MesParametres, "NomPulsGSE_Dimension")            'Dimension
            'LPT_ListeAttribPartEC(5) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Material")
            LPT_ListeAttribPartEC(5) = RecupParam(LPT_MesParametres, "NomPulsGSE_Material")             'Material
            'LPT_ListeAttribPartEC(6) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Protect")
            LPT_ListeAttribPartEC(6) = RecupParam(LPT_MesParametres, "NomPulsGSE_Protect")              'Protect
            'LPT_ListeAttribPartEC(7) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Miscellanous")
            LPT_ListeAttribPartEC(7) = RecupParam(LPT_MesParametres, "NomPulsGSE_Miscellanous")         'Miscellanous
            'LPT_ListeAttribPartEC(8) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Weight")
            LPT_ListeAttribPartEC(8) = RecupParam(LPT_MesParametres, "NomPulsGSE_Weight")               'Weight
            'LPT_ListeAttribPartEC(9) = TestParamExist(LPT_MesParametres, "NomPulsGSE_MecanoSoude")
            LPT_ListeAttribPartEC(9) = RecupParam(LPT_MesParametres, "NomPulsGSE_MecanoSoude")         'Mecano-Soudé
            
            'Ajout à la liste de nomenclature
            AddCompNom LPT_ListeAttribPartEC, NbParam
            
            'C'est un product, on relance la procedure en reccursif
            ListPartTableau LPT_Product

        ElseIf LPT_CompPart Then
            'Récupération des paramètres du part dans le tableau
            Set LPT_MesParametres = LPT_Product.ReferenceProduct.UserRefProperties
            'LPT_ListeAttribPartEC(0) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Sheet")
            LPT_ListeAttribPartEC(0) = RecupParam(LPT_MesParametres, "NomPulsGSE_Sheet")                'Sheet
            'LPT_ListeAttribPartEC(1) = TestParamExist(LPT_MesParametres, "NomPulsGSE_ItemNb")
            LPT_ListeAttribPartEC(1) = RecupParam(LPT_MesParametres, "NomPulsGSE_ItemNb")               'Item Nbr
            'Detection des elements du commerce ou fabriqués
            'Si pièce fabriquée, on récupère le parNumber
            If TypeElement(LPT_Product.PartNumber, TypeNum) = 9 Then
                'c'est une pièce du commerce
                LPT_ListeAttribPartEC(10) = LPT_Product.PartNumber                                       'PartNumber
                'LPT_ListeAttribPartEC(2) = TestParamExist(LPT_MesParametres, "NomPulsGSE_SupplierRef")
                LPT_ListeAttribPartEC(2) = RecupParam(LPT_MesParametres, "NomPulsGSE_SupplierRef")      'SupplierRef
            Else
                'C'est une pièce fabriquée
                LPT_ListeAttribPartEC(10) = LPT_Product.PartNumber                                       'PartNumber
                LPT_ListeAttribPartEC(2) = LPT_Product.PartNumber                                       'SupplierRef
            End If
            LPT_ListeAttribPartEC(3) = LPT_Product.DescriptionRef                                       'DescriptionRef
            'LPT_ListeAttribPartEC(4) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Dimension")
            LPT_ListeAttribPartEC(4) = RecupParam(LPT_MesParametres, "NomPulsGSE_Dimension")            'Dimension
            'LPT_ListeAttribPartEC(5) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Material")
            LPT_ListeAttribPartEC(5) = RecupParam(LPT_MesParametres, "NomPulsGSE_Material")             'Material
            'LPT_ListeAttribPartEC(6) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Protect")
            LPT_ListeAttribPartEC(6) = RecupParam(LPT_MesParametres, "NomPulsGSE_Protect")              'Protect
            'LPT_ListeAttribPartEC(7) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Miscellanous")
            LPT_ListeAttribPartEC(7) = RecupParam(LPT_MesParametres, "NomPulsGSE_Miscellanous")         'Miscellanous
            'LPT_ListeAttribPartEC(8) = TestParamExist(LPT_MesParametres, "NomPulsGSE_Weight")
            LPT_ListeAttribPartEC(8) = RecupParam(LPT_MesParametres, "NomPulsGSE_Weight")               'Weight
            'LPT_ListeAttribPartEC(9) = TestParamExist(LPT_MesParametres, "NomPulsGSE_MecanoSoude")
            LPT_ListeAttribPartEC(9) = RecupParam(LPT_MesParametres, "NomPulsGSE_MecanoSoude")         'Mecano-Soudé
            'Ajout à la liste de nomenclature
            AddCompNom LPT_ListeAttribPartEC, NbParam

        End If
    End If
Next i

GoTo quit_err_ListPartTableau

err_ListPartTableau:
MsgBox Err.Number & Err.Description

quit_err_ListPartTableau:
End Sub


 Public Sub MajAttributs()
 ' *****************************************************************
' * Mise à jours des attributs du 3D avec les modification apportées par l'utilisateur
' * Création CFR le 05/11/2012
' * Dernière modification le 03/10/14
' *     externalisation dans une sous procédure
' *****************************************************************
On Error GoTo err_MajAttrib
Dim MA_NomComposantCherche As String, MA_NomComposantEC As String
'Document à mettre à jour
Dim DocumentMaj As Document
Set Coll_Documents = CATIA.Documents
Dim MA_Parametres As Parameters
Dim CompProduct, CompPart As Boolean
Dim TypEltemp As Integer
Dim i As Long
Dim MA_CompPart As Boolean, MA_CompProduct As Boolean
Dim tmp As String

'Chargement de la barre de progression
    Load Frm_Progression
    Frm_Progression.Show vbModeless
    Frm_Progression.Caption = " Enregistrement des attributs dans les parts. Veuillez patienter..."
    ProgressBar (1)
    
For i = 0 To UBound(Frm_Nomenclature.LBx_Nomenclature.List)
    'Recherche du Part ou du Product dans la collection des documents
    'test s'il s'agit d'un Part ou d'un Product
    ProgressBar (100 / UBound(Frm_Nomenclature.LBx_Nomenclature.List) * i)
    MA_CompPart = False
    MA_CompProduct = False
    MA_NomComposantCherche = Frm_Nomenclature.LBx_Nomenclature.List(i, 10)                     'PartNumber
    
    'For j = 1 To Coll_documents.Count
    '    MA_NomComposantEC = Coll_documents.Item(j).Name
    '    pospt = InStr(1, MA_NomComposantEC, ".", vbTextCompare)
    '    If Left(MA_NomComposantEC, pospt - 1) = MA_NomComposantCherche Then
    '        If Right(MA_NomComposantEC, Len(MA_NomComposantEC) - pospt) = "CATProduct" Then
    '            MA_CompPart = False
    '            MA_CompProduct = True
    '        ElseIf Right(MA_NomComposantEC, Len(MA_NomComposantEC) - pospt) = "CATPart" Then
    '            MA_CompPart = True
    '            MA_CompProduct = False
    '        End If
    '        Set DocumentMaj = Coll_documents.Item(j)
    '    End If
    'Next j

    If TypeElement(CStr(MA_NomComposantCherche), TypeNum) < 7 Then
        MA_NomComposantEC = MA_NomComposantCherche & ".CATProduct"
        MA_CompPart = False
        MA_CompProduct = True
    Else
        MA_NomComposantEC = MA_NomComposantCherche & ".CATPart"
        MA_CompPart = True
        MA_CompProduct = False
    End If
    Set DocumentMaj = Coll_Documents.Item(CStr(MA_NomComposantEC))

    If MA_CompProduct Or MA_CompPart Then 'Paramètres communs aux parts et products
        Set MA_Parametres = DocumentMaj.Product.UserRefProperties
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Sheet", Frm_Nomenclature.LBx_Nomenclature.List(i, 0))             'Sheet
        'ItemNb calcul automatique
        'PartNumber calcul automatique
        DocumentMaj.Product.DescriptionRef = Frm_Nomenclature.LBx_Nomenclature.List(i, 3)                                  'DescriptionRef
    End If
    TypEltemp = TypeElement(Frm_Nomenclature.LBx_Nomenclature.List(i, 10), TypeNum)
    If MA_CompProduct Then
        'On reprend le PartNumber car Pas de SupplierRef sur les products
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_SupplierRef", DocumentMaj.Product.PartNumber)   'On reprend le PartNumber car Pas de SupplierRef sur les products
        'Pas de Dimension sur Products
        'Pas de Matérial sur les products
        
        'Cas grand S-ens et petits S-ens
        If TypEltemp >= 3 And TypEltemp <= 6 Then 'PartNumber
            tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Protect", Frm_Nomenclature.LBx_Nomenclature.List(i, 6))       'Protect
        End If
        'Pas de Miscelanous sur les products 'sauf pour les SYM des petits S-ENS
        If TypEltemp = 6 Then                                             'PartNumber
            tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Miscellanous", Frm_Nomenclature.LBx_Nomenclature.List(i, 7))  'Miscellanous
        End If
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Weight", Frm_Nomenclature.LBx_Nomenclature.List(i, 8))            'Weight
        'Cas de MécanoSoudé
        If TypEltemp >= 3 And TypEltemp <= 8 Then 'PartNumber
            tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_MecanoSoude", Frm_Nomenclature.LBx_Nomenclature.List(i, 9))  'Mecano-Soudé
        End If
        
    ElseIf MA_CompPart Then
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_SupplierRef", AjoutQuote(Frm_Nomenclature.LBx_Nomenclature.List(i, 2)))       'SupplierRef
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Dimension", Frm_Nomenclature.LBx_Nomenclature.List(i, 4))        'Dimension
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Material", Frm_Nomenclature.LBx_Nomenclature.List(i, 5))         'Material
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Protect", Frm_Nomenclature.LBx_Nomenclature.List(i, 6))           'Protect
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Miscellanous", Frm_Nomenclature.LBx_Nomenclature.List(i, 7))      'Miscellanous
        tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_Weight", Frm_Nomenclature.LBx_Nomenclature.List(i, 8))            'Weight
        'Cas de mécano Soudé
        If TypEltemp >= 7 And TypEltemp <= 8 Then 'PartNumber
            tmp = CreateParamExist(MA_Parametres, "NomPulsGSE_MecanoSoude", Frm_Nomenclature.LBx_Nomenclature.List(i, 9))  'Mecano-Soudé
        End If
    End If
Next i

GoTo quit_err_MajAttrib

err_MajAttrib:
MsgBox Err.Number & Err.Description

quit_err_MajAttrib:
 End Sub
 
