Attribute VB_Name = "a_Infos_Outillage"
Option Explicit

'Liste unique du nom des parts
'Collections des documents ouverts
Sub catmain()
On Error GoTo Err_CatMain
' *****************************************************************
' * Création d'attributs pour génération de nomenclatures automatique
' * Construit une liste des parts contenu dans le product pointé par l'utilisateur
' * ouvre un boite de dialogue permettant de documenter les infos générales
' * puis ajoute ou met a jour les attibuts de chaque part avec les valeurs de la boite de dialogue
' * Création CFR le 17/07/2013
' * modification le : 18/09/14
' *    Ajout module de classe xMacroLocation
' * modification le : 28/10/14
' *    Prise en compte de 2 systemes de numérotation des achats 500 à 999 ou 700 à 900
' * modification le 09/02/2015
' *     Numerotation des caisses = 040 à 199 au lie de 100 à 199
' *****************************************************************
On Error Resume Next

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "a_Infos_Outillage", VMacro

'Chargement des variables
Dim MacroLocation As New xMacroLocation
If Not (MacroLocation.FicIniExist("VarNomenclatureGSE.ini")) Then ' on est pas dans l'environnement GSE_R21
    MsgBox "Vous n'etes pas dans l'environnement GRE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
    Exit Sub
Else
    MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
    CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
End If

'tous les documents ouverts
Set Coll_Documents = CATIA.Documents
NbPartsNames = Coll_Documents.Count

'Ensemble général
Dim EnsGeneralDocument As ProductDocument
Set EnsGeneralDocument = CATIA.ActiveDocument
Dim EnsGeneralProduct As Product
Set EnsGeneralProduct = EnsGeneralDocument.Product
Dim EnsGeneralProducts As Products
Set EnsGeneralProducts = EnsGeneralProduct.Products

Dim Tmp_UserGuide As String, tmp_Site_AB As String, tmp_Client As String, tmp_CHK As String, Tmp_DatePlan As String
Dim Tmp_CE As String, Tmp_OutilBase As String, Tmp_Caisse   As String
Dim Tmp_ListeVariante() As String
Dim tmp As String, tmpVal As String
Dim i As Long

' paramètres de l'ensemble général(EG)
Dim ParametresEG As Parameters
'NomPulsGSE_DesignOutillage
'NomPulsGSE_NoOutillage
'NomPulsGSE_SiteAB
'NomPulsGSE_CHK
'NomPulsGSE_Client
'NomPulsGSE_DatePlan
'NomPulsGSE_CE
'NomPulsGSE_PresUserGuide
'NomPulsGSE_PresCaisse
'NomPulsGSE_NoCaisse
'NomPulsGSE_Sheet
'NomPulsGSE_ItemNb
'Dimension         'Pas sur product de tête
'Material          'Pas sur product de tête
'Protect           'Pas sur product de tête
'Miscellanous      'Pas sur product de tête
'SupplierRef       'Pas sur product de tête
'Weight            'Pas sur product de tête
'MecanoSoude       'Pas sur product de tête
'NomPulsGSE_TypeNum

'Test si le product général est vide
Dim NbrParts As Integer
NbrParts = EnsGeneralProducts.Count
If NbrParts = 0 Then
    MsgBox "Ce product est vide", vbOKOnly, "Erreur"
    Exit Sub
End If

'Chargement en mode conception
 '   EnsGeneralProduct.ApplyWorkMode DESIGN_MODE

'Initialisation de la liste des parts
ReDim ListPartsNames(0)
ListPartsNames(0) = ""

'Ouverture de la boite de dialogue de renseignement des info générales
    Load Frm_NomOutillage
    'documente les paramètres dans les champs
    Set ParametresEG = EnsGeneralProduct.UserRefProperties
    Frm_NomOutillage.Tbx_Designation = TestParamExist(ParametresEG, "NomPulsGSE_DesignOutillage")
    Tmp_UserGuide = TestParamExist(ParametresEG, "NomPulsGSE_PresUserGuide")
    If Tmp_UserGuide = "OUI" Then Frm_NomOutillage.ChB_UserGuide = True Else Frm_NomOutillage.ChB_UserGuide = False
    
    'Site Airbus
    tmp_Site_AB = TestParamExist(ParametresEG, "NomPulsGSE_SiteAB")
    If tmp_Site_AB <> "" Then
        Frm_NomOutillage.Cbx_SiteAirbus = tmp_Site_AB
    End If
    ' CHK
    tmp_CHK = TestParamExist(ParametresEG, "NomPulsGSE_CHK")
    If tmp_CHK <> "" Then
        Frm_NomOutillage.Tbx_CHK = tmp_CHK
    End If
    'Client
    tmp_Client = TestParamExist(ParametresEG, "NomPulsGSE_Client")
    If tmp_Client <> "" Then
        Frm_NomOutillage.Cbx_Client = tmp_Client
    End If
    'Date Plan
    Tmp_DatePlan = TestParamExist(ParametresEG, "NomPulsGSE_DatePlan")
    If Tmp_DatePlan = "" Then
        Frm_NomOutillage.Tbx_DatePlan = CreateParamExist(ParametresEG, "NomPulsGSE_DatePlan", Txt2Digit(Day(Date)) & "/" & Txt2Digit(Month(Date)) & "/" & Year(Date))
    Else
        Frm_NomOutillage.Tbx_DatePlan = Tmp_DatePlan
    End If
    Tmp_CE = TestParamExist(ParametresEG, "NomPulsGSE_CE")
    If Tmp_CE = "OUI" Then Frm_NomOutillage.ChB_CE = True Else Frm_NomOutillage.ChB_CE = False
'Recherche du Numero de l'outillage
    Frm_NomOutillage.Tbx_NoOutillage = EnsGeneralProduct.PartNumber 'Numero Outillage est le PartNumber
    
'Detection de l'outil de base
    Tmp_OutilBase = DetectOutilBase(EnsGeneralProduct)
    If Tmp_OutilBase = "NON" Then
        MsgBox "Pas d'outillage de base (000) détecté dans ce Product. Relancez la macro dans un ensemble général.", vbCritical, "Erreur de modèle"
        Exit Sub
    Else
        Frm_NomOutillage.Tbx_NoOutilBase = Tmp_OutilBase
    End If
    
'Détection des variantes
    If DetectVariante(EnsGeneralProduct) Then
        Frm_NomOutillage.ChB_VarianteOutillage = True
        Tmp_ListeVariante = ListeVariante(EnsGeneralProduct)
        Frm_NomOutillage.LD_NomVariantes = Tmp_ListeVariante(0)
        For i = 0 To UBound(Tmp_ListeVariante, 1)
            Frm_NomOutillage.LD_NomVariantes.AddItem (Tmp_ListeVariante(i))
        Next
    End If

 'Détection Caisse
    Tmp_Caisse = DetectCaisse(EnsGeneralProduct)
    If Tmp_Caisse = "NON" Then
        Frm_NomOutillage.ChB_Caisse = False
    Else
        Frm_NomOutillage.ChB_Caisse = True
        Frm_NomOutillage.Tbx_NoCaisse = Tmp_Caisse
    End If
   'Détecte si le Type de Numérotation est déja renseigné.
   ' Si non, c'est un ancien plan avec l'ancienen numérotation (Type1)
    If TestParamExist(ParametresEG, "NomPulsGSE_TypeNum") = 1 Then
        Frm_NomOutillage.RB_TypNum1 = True
    Else
        Frm_NomOutillage.RB_TypNum2 = True
    End If
    
    Frm_NomOutillage.Show
    If Not (Frm_NomOutillage.ChB_OkAnnule) Then Exit Sub

'Création ou mise a jour des Paramètres sur le product général
    CreateParamExist2 ParametresEG, "NomPulsGSE_DesignOutillage", Frm_NomOutillage.Tbx_Designation
    CreateParamExist2 ParametresEG, "NomPulsGSE_NoOutillage", Frm_NomOutillage.Tbx_NoOutillage
    tmpVal = Frm_NomOutillage.Cbx_SiteAirbus
    CreateParamExist2 ParametresEG, "NomPulsGSE_SiteAB", CStr(tmpVal)
    tmpVal = Frm_NomOutillage.Tbx_CHK
    CreateParamExist2 ParametresEG, "NomPulsGSE_CHK", CStr(tmpVal)
    tmpVal = Frm_NomOutillage.Cbx_Client
    CreateParamExist2 ParametresEG, "NomPulsGSE_Client", CStr(tmpVal)
    CreateParamExist2 ParametresEG, "NomPulsGSE_DatePlan", Frm_NomOutillage.Tbx_DatePlan
    'Traitement du logo CE
    If Frm_NomOutillage.ChB_CE Then tmpVal = "OUI" Else tmpVal = "NON"
    CreateParamExist2 ParametresEG, "NomPulsGSE_CE", CStr(tmpVal)
    'Traitement du userGuide
    If Frm_NomOutillage.ChB_UserGuide Then tmpVal = "OUI" Else tmpVal = "NON"
    CreateParamExist2 ParametresEG, "NomPulsGSE_PresUserGuide", CStr(tmpVal)
    'Traitement des caisses
    If Frm_NomOutillage.ChB_Caisse Then tmpVal = "OUI" Else tmpVal = "NON"
    CreateParamExist2 ParametresEG, "NomPulsGSE_PresCaisse", CStr(tmpVal)
    CreateParamExist2 ParametresEG, "NomPulsGSE_NoCaisse", Frm_NomOutillage.Tbx_NoCaisse
    CreateParamExist2 ParametresEG, "NomPulsGSE_Sheet", ""
    CreateParamExist2 ParametresEG, "NomPulsGSE_ItemNb", ""  'Vide dans le cas d'un Product de tête
    'Type de Numérotation (Documente la variable publique pour le reste de la procédure)
    If Frm_NomOutillage.RB_TypNum1 Then
        TypeNum = "1"
    ElseIf Frm_NomOutillage.RB_TypNum2 Then
        TypeNum = "2"
    End If
        CreateParamExist2 ParametresEG, "NomPulsGSE_TypeNum", TypeNum
        
    'Traitement des variantes
    'If Frm_NomOutillage.ChB_VarianteOutillage Then
    '    tmpVal = "OUI"
    '    For i = 1 To Frm_NomOutillage.LD_NomVariantes.LineCount
    '        tmpListe = tmpListe & Frm_NomOutillage.LD_NomVariantes(i) & ";"
    '    Next
    '        tmpListe = Left(tmpListe, Len(tmpListe) - 1)
    'Else
    '    tmpVal = "NON"
     '   tmpListe = ""
    'End If
    'CreateParamExist2(ParametresEG, "NompulsGSE_PresVariante", CStr(tmpVal))
    'CreateParamExist2(ParametresEG, "NompulsGSE_NoVariante", CStr(tmpListe))
    
    Frm_NomOutillage.Hide
    
'Chargement de la barre de progression
    Load Frm_Progression
    Frm_Progression.Show vbModeless
    Frm_Progression.Caption = " Création des paramètres. Veuillez patienter..."
    ProgressBar (1)
    
'Création ou mise à jour des Paramètres sur tous les products et les parts du products général
    PropageAttributs EnsGeneralProducts

    Unload Frm_Progression
    Unload Frm_NomOutillage
    MsgBox "Fin de l'initialisation des attibuts", vbInformation, "Initialisation des attributs"

Quit_CatMain:
Exit Sub
Err_CatMain:
MsgBox Err.Number & " - " & Err.Description
End Sub

Public Sub PropageAttributs(PA_Products As Products)
' *****************************************************************
' * Construit la liste des parts et récupère leurs attributs
' * balaye l'ensembles des Item de la collection des products  et teste s'il s'agit d'un part ou d'un product
' * Pour les parts, Crée la série d'attibuts
' * pour les products crée la série des attibuts puis appelle la procedure en recursif pour rechercher les "sous parts"
' * Création CFR le 23/07/2013
' * Dernière modification le
' *****************************************************************
On Error Resume Next
Dim PA_Product As Product
'Nom de l'Item en cours de traitement dans l'arbre
Dim PA_NomItemEC As String
Dim PA_CompPart, PA_CompProduct As Boolean
Dim PA_Document As Document
Dim PA_NBrPRoducts As Integer

'Paramètres

Dim PA_MesParametres As Parameters
'NomPulsGSE_DesignOutillage
'NomPulsGSE_NoOutillage
'NomPulsGSE_SiteAB
'NomPulsGSE_CHK
'NomPulsGSE_DatePlan
'CE                 'Uniquement sur product de tête
'PresUserGuide      'Uniquement sur product de tête
'PresCaisse         'Uniquement sur product de tête
'NoCaisse           'Uniquement sur product de tête
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

Dim tmp
Dim Tmp_Val As String
Dim TypElNum As Integer
Dim PartNum As String
Dim i As Long
Dim PartExistinList As Boolean
Dim tmp_Misc As String

'Nombre d'Item dans le products
PA_NBrPRoducts = PA_Products.Count
    For i = 1 To PA_NBrPRoducts
        Set PA_Product = PA_Products.Item(i)
        PartExistinList = False
        TypElNum = CInt(Left(TypeElement(PA_Product.PartNumber, TypeNum), 1))
        'Récupération du Numéro de Part (sans le -FLXxx)
        If TypeElementFlx(PA_Product.PartNumber) Then
            'PartNum = Left(PA_Product.PartNumber, Len(PA_Product.PartNumber) - 5)
            PartNum = Left(PA_Product.PartNumber, Len(PA_Product.PartNumber) - 6)
        Else
            PartNum = PA_Product.PartNumber
        End If
        'Vérification que le part n'a pas déja été traité
        'Cas des parts communs a plusieurs produits
        If ListPartUnique(PA_Product.PartNumber, ListPartsNames) Then

            ProgressBar (100 / NbPartsNames * UBound(ListPartsNames(), 1))
            'Ajout a la liste des noms de parts
            ReDim Preserve ListPartsNames(UBound(ListPartsNames, 1) + 1)
            ListPartsNames(UBound(ListPartsNames(), 1)) = PA_Product.PartNumber
            'test s'il s'agit d'un Part ou d'un Product
            PA_CompPart = True
            PA_CompProduct = True
            Err.Clear
            PA_NomItemEC = CStr(PA_Product.PartNumber) & ".CATProduct"
            Set PA_Document = Coll_Documents.Item(PA_NomItemEC)
            If (Err.Number <> 0) Then
                'Le composant n'est pas un CATProduct
                Err.Clear
                PA_CompProduct = False
            End If
            PA_NomItemEC = CStr(PA_Product.PartNumber) & ".CATPart"
            Set PA_Document = Coll_Documents.Item(PA_NomItemEC)
            If (Err.Number <> 0) Then
                'Le composant n'est pas un CATPart
                Err.Clear
                PA_CompPart = False
            End If
            If PA_CompProduct Then
            'Chargement en mode conception
            PA_Product.ApplyWorkMode DESIGN_MODE
                'Création des attributs sans valeur
                'sauf pour N° outillage, Nom outillage, Site Airbus, CHK et date plan, qui sont créés avec les valeurs enregistrées dans le formulaire
                Set PA_MesParametres = PA_Product.ReferenceProduct.UserRefProperties
                'Cas général (tous les products)
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_DesignOutillage", CStr(Frm_NomOutillage.Tbx_Designation))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_NoOutillage", CStr(Frm_NomOutillage.Tbx_NoOutillage))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_SiteAB", CStr(Frm_NomOutillage.Cbx_SiteAirbus))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_CHK", CStr(Frm_NomOutillage.Tbx_CHK))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_Client", CStr(Frm_NomOutillage.Cbx_Client))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_DatePlan", CStr(Frm_NomOutillage.Tbx_DatePlan))
                tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Sheet")
                tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Weight")
                'Type de Numérotation
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_TypeNum", TypeNum)
                'Cas des outillages 000 et des variantes 001 -> 039
                If TypElNum = 1 Or TypElNum = 2 Then
                    tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_ItemNb", "")
                    'Pas de Protect
                End If
                'Cas des petits et Grands S-Ens
                If TypElNum >= 3 And TypElNum <= 6 Then
                    tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_ItemNb", CStr(Right(PartNum, 3)))
                    tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Protect")
                End If
                'Cas des grand S-Ens SYM et petit S-Ens SYM
                If TypElNum = 4 Or TypElNum = 6 Then
                    tmp_Misc = "SYM TO " & CStr(CInt(Right(PartNum, 3)) - 1)
                    tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_Miscellanous", CStr(tmp_Misc))
                Else
                    tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Miscellanous")
                End If
                'Cas des mécano-soudé
                If TypElNum >= 3 And TypElNum <= 6 Then
                    tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_MecanoSoude")
                End If
                              
                ' Si C'est une caisse on document la désignation
                If Frm_NomOutillage.ChB_Caisse Then
                    If PA_Product.PartNumber = Frm_NomOutillage.Tbx_NoCaisse Then
                        PA_Product.DescriptionRef = "STORAGE BOX"
                    End If
                End If
                
                ' C'est un product, on relance la procedure en reccursif
                PropageAttributs PA_Document.Product.Products
            
            ElseIf PA_CompPart Then
                'Création des attributs sans valeur
                'sauf pour N° outillage, Nom outillage qui sont créés avec les valeurs enregistrées dans le formulaire
                Set PA_MesParametres = PA_Product.ReferenceProduct.UserRefProperties
                CreateParamExist2 PA_MesParametres, "NomPulsGSE_DesignOutillage", CStr(Frm_NomOutillage.Tbx_Designation)
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_NoOutillage", CStr(Frm_NomOutillage.Tbx_NoOutillage))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_SiteAB", CStr(Frm_NomOutillage.Cbx_SiteAirbus))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_CHK", CStr(Frm_NomOutillage.Tbx_CHK))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_Client", CStr(Frm_NomOutillage.Cbx_Client))
                'tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_DatePlan", Txt2Digit(Day(Date)) & "/" & Txt2Digit(Month(Date)) & "/" & Year(Date))
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_DatePlan", CStr(Frm_NomOutillage.Tbx_DatePlan))
                tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Sheet")
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_ItemNb", CStr(Right(PartNum, 3)))
                tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Dimension")
                tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Material")
                tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Protect")
                tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_SupplierRef")
                tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Weight")
                'Type de Numérotation
                tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_TypeNum", TypeNum)
                'Cas des Parts SYM
                If TypElNum = 8 Then
                    tmp_Misc = "SYM TO " & CStr(CInt(Right(PartNum, 3)) - 1)
                    tmp = CreateParamExist(PA_MesParametres, "NomPulsGSE_Miscellanous", CStr(tmp_Misc))
                Else
                    tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_Miscellanous")
                End If
                'Cas des mécano-soudé
                If TypElNum >= 7 And TypElNum <= 8 Then
                    tmp = TestParamExist(PA_MesParametres, "NomPulsGSE_MecanoSoude")
                End If
                     
            End If
        End If
    Next
End Sub
