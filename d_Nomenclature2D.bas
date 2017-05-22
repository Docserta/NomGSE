Attribute VB_Name = "d_Nomenclature2D"
Option Explicit

Sub catmain()
' *****************************************************************
' * Apartir du 2D, demande a l'utilisateur de quel part ou product il souhaite extraire la nomenclature
' * puis constitue la nomenclature et l'integre dans le 2D
' * Prise en compte des particularités comme l'ajout de la ligne User Guide et Caisse dans la nom du product outillage.
' * Création CFR le 05/11/2012
' * Version 2.4
' * Dernière modification le : 29/02/14
' *     Ajout détection Langue de catia pour champs de Nomenclature
' * modification le : 18/09/14
' *    Ajout module de classe xMacroLocation
' * modification le : 28/10/14
' *    Prise en compte de 2 systemes de numérotation des achats 500 à 999 ou 700 à 900
' * modification le : 24/11/14
' *    Ajout dans le calque de détail d'un texte portant le numéro du part/product lié au plan
' *    pour initialisation de la macro d_nomenclature2D
' * modification le : 22/12/14
' *    Changé coordonnées d'insertion du nota Welding
' * modification le : 14/01/15
' *    Modification du saut de ligne pour les GSE allemnad (plus de saut entre 200/500  ou 500/900
' * modification le : 15/12/15
' *    Prise en charge des nomenclatures miltiplan ( symétriques et variantes)
' * Modification le  : 21/04/16
' *     Prise en compte du site Airbus pour formatage des N° de planche (2 ou 3 digit)
' * Modification le : 21/09/16
' *     Suppression du dito "WELDING_NOTE" s'il est déja présent
' * Modification le : 18/05/17
' *     Ajout nombre de planche total sur planche 1
' *****************************************************************

On Error GoTo Err_Nomenclature2D
'On Error Resume Next

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "d_Nomenclature2D", VMacro

'Chargement des variables
Dim MacroLocation As New xMacroLocation
If Not (MacroLocation.FicIniExist("VarNomenclatureGSE.ini")) Then ' on est pas dans l'environnement GSE_R21
    MsgBox "Vous n'etes pas dans l'environnement GRE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
    Exit Sub
Else
    MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
    CheminDestNomenclature = MacroLocation.ValVar("CheminDestNomenclature")
    CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
    Nom_FicCageCodes = MacroLocation.ValVar("Nom_FicCageCodes")
End If
    
'Test si le document actif est un Drawing
    Set ActiveDoc = CATIA.ActiveDocument
    On Error Resume Next
    Dim ActiveDrawingDoc As DrawingDocument
    Set ActiveDrawingDoc = ActiveDoc
    If (Err <> 0) Then
        MsgBox "Le document actif n'est pas un drawing. Activez un drawing avant de lancer cette macro.", vbCritical, "Erreur"
        Exit Sub
    End If
    On Error GoTo 0
    
'Test si le drawing contient les tableaux de nomenclature
'Si non la macro ne peux pas descendre la nomenclature
    Dim Col_CalquesDrawActif As DrawingSheets
    Set Col_CalquesDrawActif = ActiveDrawingDoc.Sheets
    
    Dim ActiveDrawingCalque As DrawingSheet, DrawingCalqueDet As DrawingSheet
    Set ActiveDrawingCalque = ActiveDrawingDoc.Sheets.activeSheet
    
    Dim ActiveDrawingVues As DrawingViews
    Set ActiveDrawingVues = ActiveDrawingCalque.Views
    Dim ActiveDrawingVue As DrawingView, DrawingVueTxt As DrawingView
    Set ActiveDrawingVue = ActiveDrawingVues.Item("Background View")
    
    Dim ActiveDrawingTables As DrawingTables
    Set ActiveDrawingTables = ActiveDrawingVue.Tables
    Dim TabNom, TabTitre As Boolean
        TabNom = False
    Dim i As Long, j As Long
    
    If ActiveDrawingTables.Count > 0 Then
        Dim ActiveDrawingTable As DrawingTable
        For i = 1 To ActiveDrawingTables.Count
            If ActiveDrawingTables.Item(i).Name = "TableauNom" Then TabNom = True ' la table existe
        Next i
    End If
    If Not (TabNom) Then
        MsgBox "Il manque un tableau nécéssaire à la création de la nomenclature. Recréez le CatDrawing avec la macro 'b_Creation_plan'", vbCritical, "Erreur"
        Exit Sub
    End If

'Recherche dans le calque de détail si un texte (TxtNumDetail) contenant des numéros de part ou de product est présent
    If Col_CalquesDrawActif.Item(2).Name = "Calque.2 (Détail)" Then 'pour les plans Français
        Set DrawingCalqueDet = Col_CalquesDrawActif.Item("Calque.2 (Détail)")
    ElseIf Col_CalquesDrawActif.Item(2).Name = "Sheet.2 (Detail)" Then 'pour les plans anglais
        Set DrawingCalqueDet = Col_CalquesDrawActif.Item("Sheet.2 (Detail)")
    End If
    Dim MesDetails2D As DrawingViews
        Set MesDetails2D = DrawingCalqueDet.Views
    Dim MonDetail2D As DrawingView
    
    Dim DetailExplicite As String
    Dim LstDetailsExplicite() As String
    
    Dim TxtNumDetail As DrawingText
    For j = 1 To MesDetails2D.Count
        If MesDetails2D.Item(j).Name = "NumDetailView" Then
            Set MonDetail2D = MesDetails2D.Item(j)
        End If
    Next j
    On Error Resume Next
    Set TxtNumDetail = MonDetail2D.Texts.Item(1)
    If (Err.Number <> 0) Then
        Err.Clear
        DetailExplicite = ""
    Else
        DetailExplicite = TxtNumDetail.Text
    End If
    On Error GoTo 0
    
'Construction de la liste des part a détailler sur la planche
    Dim ListePart() As String
    ListePart = ListPart(DetailExplicite)

'Affichage du formulaire permettant de choisir le fichier dont on veux extraire la nomenclature
    Load FRM_ListFichiers
    'If Right(DetailExplicite, Len(DetailExplicite) - InStr(DetailExplicite, ".")) = "CATPart" Then
    If InStr(DetailExplicite, "CATPart") <> 0 Then
        FRM_ListFichiers.RBt_TypePlan2.Value = True
    Else
        FRM_ListFichiers.RBt_TypePlan1.Value = True
    End If
    FRM_ListFichiers.Cbx_FicaTraiter = ListePart(0)
    FRM_ListFichiers.Show
    Dim CatiaLangue As String
    If FRM_ListFichiers.ChB_CatAnglais Then
        CatiaLangue = "Anglais"
    Else
        CatiaLangue = "Français"
    End If
    
'Test si un Part autre que celui prédéfini a été choisi dans la liste
    If FRM_ListFichiers.Cbx_FicaTraiter <> ListePart(0) Then
        MsgBox "Vous avez choisi de tracer la nomenclature de : " & FRM_ListFichiers.Cbx_FicaTraiter & Chr(10) & "alors que le plan était prévu pour : " & DetailExplicite, vbInformation
        ReDim ListePart(0)
        ListePart(0) = FRM_ListFichiers.Cbx_FicaTraiter
    End If
    
'Tracé de la nomenclature
    If FRM_ListFichiers.ChB_OkAnnule Then
        If FRM_ListFichiers.RBt_TypePlan1 Then 'Plan d'ensemble
            TraceNom2DProduct CatiaLangue, ListePart
        ElseIf FRM_ListFichiers.RBt_TypePlan2 Then 'Plan de détails
            TraceNom2DPart ListePart
        End If
    End If
    Unload FRM_ListFichiers

On Error Resume Next
    ActiveDrawingCalque.Update
Err.Clear
GoTo Quit_Nomenclature2D

Quit_Nomenclature2D:
Exit Sub

Err_Nomenclature2D:
MsgBox Err.Number & " - " & Err.Description
End Sub

Public Sub TraceNom2DProduct(TNP_CatiaLangue As String, ListProd() As String)
' *****************************************************************
' * Documente la nomenclature dans le 2D du Product passé en argument sur le 2D
' *
' *
' * Création CFR le 05/11/2012
' * Dernière modification le 17/12/15
' *     Ajout d'une colonne Qté pour le Sym ou les variantes
' *     Ajout des Cage Codes
' *****************************************************************
On Error GoTo ErrTraceNom2DProduct
'On Error Resume Next
Dim ProductGenDoc As Document
Dim TNP_Product As Product
Dim ProductGenProd As Product
Dim ListParamProd() As String
Dim ListCageCode() As String
Dim i As Long, j As Integer
Dim TmpTypeEl As Integer
Dim ParametresProduct As Parameters
Dim ProductGenParams As Parameters
Dim Param_PG_PresCaisse As StrParam
Dim Param_PG_NoCaisse As StrParam
Dim Param_PresUserGuide As StrParam
Dim NomProductGen As String, Tmp_SiteAB As String, Tmp_MecanoSoude As String, TypeNum As String
Dim NotaSymExiste As Boolean
Dim NotaSoudExiste As Boolean
Dim objexcel
Dim objWorkBook
Dim TablNomProduct() As String 'Tableau de la nomenclature complète avant tracé dans le 2D
Dim TableCompile() As String 'Compilation des lignes des différents ensembles (sym ou  variantes)
Dim LigneNomTempo As String
Dim Boucle As Integer
Dim LigActive As Integer, NoEndNom  As Integer
Dim NbColPlus As Integer 'Nombre d'ensembles pour calcul du décallage de colonne
Dim NoSheetEG As String 'N0 de planche de l'ensemble gégnéral
Dim NoTotSheet As String 'Nombre de planches totale
Dim Coll_Tables As DrawingTables 'Tableau de nomenclature dans le drawing
    
    'Initialisation des variables
    Set Coll_Documents = CATIA.Documents
    Set ProductDoc = Coll_Documents.Item(ListProd(0))
    Set TNP_Product = ProductDoc.Product
    ReDim ListParamProd(UBound(ListProd), 3)
    Boucle = 0
    NoEndNom = 0
    LigActive = 1
    
'Récupération des paramètres du Product sélectionné par l'utilisateur
    Set ParametresProduct = TNP_Product.UserRefProperties
      
'Recupération des parametre du product général dans le product en cours
    NomProductGen = RecupParam(ParametresProduct, "NomPulsGSE_NoOutillage")
    NomProductGen = NomProductGen & ".CATProduct"
    Tmp_SiteAB = RecupParam(ParametresProduct, "NomPulsGSE_SiteAB")  'Site_AB
    Tmp_MecanoSoude = RecupParam(ParametresProduct, "NomPulsGSE_MecanoSoude")  'MecanoSoude
    TypeNum = RecupParam(ParametresProduct, "NomPulsGSE_TypeNum")
    NoSheetEG = RecupParam(ParametresProduct, "NomPulsGSE_Sheet")  'Sheet
    
'Récupération du product général
    Set ProductGenDoc = Coll_Documents.Item(NomProductGen)
    Set ProductGenProd = ProductGenDoc.Product

'Ajout des paramètres "Sheet", "Description" et "Weights" de chaque ensembles a la liste des ensembles a nomenclaturer
    For i = 0 To UBound(ListProd, 1)
        Set ProductDoc = Coll_Documents.Item(ListProd(i))
        Set TNP_Product = ProductDoc.Product
        Set ParametresProduct = TNP_Product.UserRefProperties
        ListParamProd(i, 0) = ProductDoc.Name
        ListParamProd(i, 1) = RecupParam(ParametresProduct, "NomPulsGSE_Sheet") 'sheet
        ListParamProd(i, 2) = SautLigne(ProductDoc.Product.DescriptionRef)  'Description
        ListParamProd(i, 3) = RecupParam(ParametresProduct, "NomPulsGSE_Weight") 'sheet
    Next i

'Barre de progression
    Load Frm_Progression
    Frm_Progression.Show vbModeless
    ProgressBar (1)

'Formatage de la nomenclature en fonction de la langue du Catia
    'Variable temp de traduction des descriptions
    Dim TNP_LangueQt, TNP_LangueRef, TNP_LangueDesc As String
    
    If TNP_CatiaLangue = "Anglais" Then
        TNP_LangueQt = "Quantity"
        TNP_LangueRef = "Part Number"
        TNP_LangueDesc = "Product Description"
    ElseIf TNP_CatiaLangue = "Français" Then
        TNP_LangueQt = "Quantité"
        TNP_LangueRef = "Référence"
        TNP_LangueDesc = "Description du produit"
    Else
        MsgBox "Erreur dans la détection de la langue paramétré dans Catia."
        GoTo QuitTraceNom2DProduct
    End If

'verifie si un fichier de nomenclature est déja présent et l'efface
    If Not (EffaceFicNom(CheminDestNomenclature, ProductGenProd.Name & ".xls")) Then
        End
    End If

'Extraction de la nomenclature du product général et sauvegarde dans un fichier excel
    Dim assemblyConvertor1Variant
    Dim assemblyConvertor1 As AssemblyConvertor
    Set assemblyConvertor1 = ProductGenProd.GetItem("BillOfMaterial")

    Dim arrayOfVariantOfBSTR1(10)
    arrayOfVariantOfBSTR1(0) = CStr(TNP_LangueQt)
    arrayOfVariantOfBSTR1(1) = "NomPulsGSE_Sheet"
    arrayOfVariantOfBSTR1(2) = "NomPulsGSE_ItemNb"
    arrayOfVariantOfBSTR1(3) = CStr(TNP_LangueRef)
    arrayOfVariantOfBSTR1(4) = "NomPulsGSE_SupplierRef"
    arrayOfVariantOfBSTR1(5) = CStr(TNP_LangueDesc)
    arrayOfVariantOfBSTR1(6) = "NomPulsGSE_Dimension"
    arrayOfVariantOfBSTR1(7) = "NomPulsGSE_Material"
    arrayOfVariantOfBSTR1(8) = "NomPulsGSE_Protect"
    arrayOfVariantOfBSTR1(9) = "NomPulsGSE_Miscellanous"
    arrayOfVariantOfBSTR1(10) = "NomPulsGSE_Weight"

    Set assemblyConvertor1Variant = assemblyConvertor1
    assemblyConvertor1Variant.SetCurrentFormat arrayOfVariantOfBSTR1

    Dim arrayOfVariantOfBSTR2(10)
    arrayOfVariantOfBSTR2(0) = CStr(TNP_LangueQt)
    arrayOfVariantOfBSTR2(1) = "NomPulsGSE_Sheet"
    arrayOfVariantOfBSTR2(2) = "NomPulsGSE_ItemNb"
    arrayOfVariantOfBSTR2(3) = CStr(TNP_LangueRef)
    arrayOfVariantOfBSTR2(4) = "NomPulsGSE_SupplierRef"
    arrayOfVariantOfBSTR2(5) = CStr(TNP_LangueDesc)
    arrayOfVariantOfBSTR2(6) = "NomPulsGSE_Dimension"
    arrayOfVariantOfBSTR2(7) = "NomPulsGSE_Material"
    arrayOfVariantOfBSTR2(8) = "NomPulsGSE_Protect"
    arrayOfVariantOfBSTR2(9) = "NomPulsGSE_Miscellanous"
    arrayOfVariantOfBSTR2(10) = "NomPulsGSE_Weight"

    Set assemblyConvertor1Variant = assemblyConvertor1
    assemblyConvertor1Variant.SetSecondaryFormat arrayOfVariantOfBSTR2
    assemblyConvertor1.[Print] "XLS", CStr(CheminDestNomenclature & ProductGenProd.Name & ".xls"), ProductGenProd

'On Error GoTo ErrTraceNom2DProduct
  
    'Creation d'un objet eXcel et ouverture de la nomenclature précédement générée
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objexcel.Workbooks.Open(CStr(CheminDestNomenclature & ProductGenProd.Name & ".xls"))
    objexcel.Visible = True
    objWorkBook.activeSheet.Visible = True

    'Compilation des lignes des différents ensembles (sym ou  variantes)
    TableCompile = CompileNom(objWorkBook, ListProd, True, Tmp_MecanoSoude, Tmp_SiteAB)

    'Nombre d'ensembles pour calcul du décallage de colonne
    NbColPlus = UBound(TableCompile, 1) - 9

    'recherhe le N° de planche maxi
    NoTotSheet = MaxSheet(TableCompile, Tmp_SiteAB)

    'Création de la ligne N°1 (Ligne de l'assemblage) et des lignes N°1 des Sym ou des variantes
    For i = 0 To NbColPlus
        'Récupération de la désignation et du poids du product à détailler dans les paramètre du Product de tète
        ReDim Preserve TablNomProduct(NbColPlus + 9, Boucle)
        TablNomProduct(NbColPlus, Boucle) = "" 'Qte du général
        'Cas des GSE Allemands
        If Tmp_SiteAB = "Allemand" Then
            TablNomProduct(NbColPlus + 1, Boucle) = "" ' pas de N° de planche
        Else
            TablNomProduct(NbColPlus + 1, Boucle) = NoSheetEG 'Sheet
        End If
        TablNomProduct(NbColPlus + 2, Boucle) = "" 'Item Nbr
        TablNomProduct(NbColPlus + 3, Boucle) = Left(ListParamProd(i, 0), InStr(ListParamProd(i, 0), ".") - 1) 'Part Nbr
        TablNomProduct(NbColPlus + 4, Boucle) = SautLigne(ListParamProd(i, 2)) 'DecriptionRef
        TablNomProduct(NbColPlus + 5, Boucle) = "" 'Dimension
        TablNomProduct(NbColPlus + 6, Boucle) = "" 'Material
        'Cas des petits sous-Ens
        TmpTypeEl = TypeElement(TablNomProduct(NbColPlus + 3, Boucle), TypeNum)
        If TmpTypeEl >= 5 And TmpTypeEl <= 6 Then
            TablNomProduct(NbColPlus + 7, Boucle) = SautLigne(RecupParam(ParametresProduct, "NomPulsGSE_Protect")) 'Protect
        Else
            TablNomProduct(NbColPlus + 7, Boucle) = "" 'Protect
        End If
        'Cas des symétriques
        If TmpTypeEl = 4 Or TmpTypeEl = 6 Or TmpTypeEl = 8 Then
            TablNomProduct(NbColPlus + 8, Boucle) = SymMiscellanous(TNP_Product.Name) 'Miscellanous
        Else
            TablNomProduct(NbColPlus + 8, Boucle) = "" 'Miscellaneous
        End If
        
        TablNomProduct(NbColPlus + 9, Boucle) = ListParamProd(i, 3) 'Weights
        Boucle = Boucle + 1
    Next i
    
    'Ajout d'une ligne vide
    ReDim Preserve TablNomProduct(NbColPlus + 9, Boucle)
    For i = 0 To 9 + NbColPlus
        TablNomProduct(i, Boucle) = ""
    Next i
    Boucle = Boucle + 1
   
    'ajout des lignes compilées a la table
    For i = 0 To UBound(TableCompile, 2)
        ReDim Preserve TablNomProduct(NbColPlus + 9, Boucle)
        For j = 0 To NbColPlus + 9
            TablNomProduct(j, Boucle) = TableCompile(j, i)
        Next j
        Boucle = Boucle + 1
    Next i
            
    'Si c'est le Product de l'outillage, traite les lignes User Guide et caisse
    If TypeElement(TNP_Product.Name, TypeNum) = 1 Then
    
        'Récupération des attributs concernant la caisse et le Userguide dans le product général
        Set ProductGenParams = ProductGenProd.UserRefProperties
        
        'Détection et ajout de userGuide
        Dim NBLigneTablNomProd As Long
        Dim Tmp_UserGuide As String
            Tmp_UserGuide = RecupParam(ProductGenParams, "NomPulsGSE_PresUserGuide")
        If Not (Tmp_UserGuide = "NON") Then
            NBLigneTablNomProd = UBound(TablNomProduct(), 2) + 1
            ReDim Preserve TablNomProduct(NbColPlus + 9, NBLigneTablNomProd)
            TablNomProduct(NbColPlus + 0, NBLigneTablNomProd) = "1" 'Qte
            TablNomProduct(NbColPlus + 1, NBLigneTablNomProd) = "01" 'Sheet
            If TypeNum = "1" Then
                TablNomProduct(NbColPlus + 2, NBLigneTablNomProd) = "900" 'Item Nbr
            ElseIf TypeNum = "2" Then
                TablNomProduct(NbColPlus + 2, NBLigneTablNomProd) = "990" 'Item Nbr
            End If
            TablNomProduct(NbColPlus + 3, NBLigneTablNomProd) = Left(TNP_Product.Name, 11) & "-GIM" 'Part Nbr
            TablNomProduct(NbColPlus + 4, NBLigneTablNomProd) = "USER GUIDE" 'Description
            TablNomProduct(NbColPlus + 5, NBLigneTablNomProd) = "" 'Dimension
            TablNomProduct(NbColPlus + 6, NBLigneTablNomProd) = "" 'Material
            TablNomProduct(NbColPlus + 7, NBLigneTablNomProd) = "" 'Protect
            TablNomProduct(NbColPlus + 8, NBLigneTablNomProd) = "" 'Miscellaneous
            TablNomProduct(NbColPlus + 9, NBLigneTablNomProd) = "" 'Weights
        End If
        
        'Détection et ajout Caisse
        Dim Temp_PresCaisse As String, Temp_NoCaisse As String
        Temp_PresCaisse = RecupParam(ProductGenParams, "NomPulsGSE_PresCaisse")
        Temp_NoCaisse = RecupParam(ProductGenParams, "NomPulsGSE_NoCaisse")
        
        Dim NoEndInsertLigneVideNom As Integer
        If Not (Temp_PresCaisse = "NON") Then
            'reprise du fichier excel de la nomenclature pour aller chercher la caise dans le product général
            'Pointage sur la premiere ligne des composants du product a analyser
            LigActive = 1
            NoEndInsertLigneVideNom = 0
            LigneNomTempo = ""
            Do While Not LigneNomTempo = Temp_NoCaisse And NoEndNom <= 2
                LigActive = LigActive + 1
                'On recherche dans la colone 4 (PartNumber)
                LigneNomTempo = Right(objWorkBook.activeSheet.cells(LigActive, 4).Value, Len(Temp_NoCaisse))
                'Si 2 lignes vides consecutive => EOF du fichier excel
                If objWorkBook.activeSheet.cells(LigActive, 1).Value = "" Then
                    NoEndNom = NoEndNom + 1
                Else
                    NoEndNom = 0
                End If
            Loop
            If NoEndNom <= 2 Then ' l'EOF n'a pas été atteint, la ligne de la caisse à été trouvée
                NBLigneTablNomProd = UBound(TablNomProduct(), 2) + 1
                ReDim Preserve TablNomProduct(NbColPlus + 9, NBLigneTablNomProd)
                TablNomProduct(NbColPlus + 0, NBLigneTablNomProd) = objWorkBook.activeSheet.cells(LigActive, 1).Value 'Qte
                TablNomProduct(NbColPlus + 1, NBLigneTablNomProd) = FormatNoSheet(CStr(objWorkBook.activeSheet.cells(LigActive, 2).Value), Tmp_SiteAB) 'Sheet
                TablNomProduct(NbColPlus + 2, NBLigneTablNomProd) = Txt3Digit(objWorkBook.activeSheet.cells(LigActive, 3).Value) 'Item Nbr
                TablNomProduct(NbColPlus + 3, NBLigneTablNomProd) = SautLigne(objWorkBook.activeSheet.cells(LigActive, 4).Value) 'Part Nbr
                TablNomProduct(NbColPlus + 4, NBLigneTablNomProd) = SautLigne(objWorkBook.activeSheet.cells(LigActive, 6).Value) 'Description
                TablNomProduct(NbColPlus + 5, NBLigneTablNomProd) = SautLigne(objWorkBook.activeSheet.cells(LigActive, 7).Value) 'Dimension
                TablNomProduct(NbColPlus + 6, NBLigneTablNomProd) = SautLigne(VigPt(objWorkBook.activeSheet.cells(LigActive, 8).Value)) 'Material
                TablNomProduct(NbColPlus + 7, NBLigneTablNomProd) = SautLigne(objWorkBook.activeSheet.cells(LigActive, 9).Value) 'Protect
                TablNomProduct(NbColPlus + 8, NBLigneTablNomProd) = SautLigne(objWorkBook.activeSheet.cells(LigActive, 10).Value) 'Miscellaneous
                TablNomProduct(NbColPlus + 9, NBLigneTablNomProd) = objWorkBook.activeSheet.cells(LigActive, 11).Value 'Weights
            End If
        End If
        
        'Tri du TablNomProduct pour mettre la caisse au bon endroit
        'Création d'un tableau temporaire sans les 2 premières lignes
        Dim Temp_TablNomProduct() As String
        NBLigneTablNomProd = UBound(TablNomProduct(), 2) - 2
        ReDim Temp_TablNomProduct(NbColPlus + 9, NBLigneTablNomProd)
        For i = 0 To NBLigneTablNomProd
            For j = 0 To NbColPlus + 9
                Temp_TablNomProduct(j, i) = TablNomProduct(j, i + 2)
            Next
        Next
        Temp_TablNomProduct() = TranspositionTabl(TriList2D(TranspositionTabl(Temp_TablNomProduct), 2, True))
        
        'Remplacement des lignes du tableau par les lignes triées
        For i = 0 To NBLigneTablNomProd
            For j = 0 To NbColPlus + 9
                TablNomProduct(j, i + 2) = Temp_TablNomProduct(j, i)
            Next
        Next
    End If

'Calcule des Cage Code
    ListCageCode() = CageCode(objWorkBook.activeSheet, 10, TypeNum)
    
'Fermeture du fichier excel
    objWorkBook.Close

'Insertion des lignes blanches entre les centaines pour les GSE allemand
    If Tmp_SiteAB = "Allemand" Then
        TablNomProduct() = InsertLigneVide(TablNomProduct(), TypeNum, NbColPlus)
    End If

'Copie de la nomenclature dans le 2D
    'Documents
    Dim DrawActif As DrawingDocument
    Set DrawActif = CATIA.ActiveDocument
    Dim Coll_Calques As DrawingSheets
    Set Coll_Calques = DrawActif.Sheets
    Dim CalqueActif As DrawingSheet
    Set CalqueActif = DrawActif.Sheets.activeSheet
    Dim Coll_Vues As DrawingViews
    Set Coll_Vues = CalqueActif.Views
    Dim Vue_Back As DrawingView
    Set Vue_Back = Coll_Vues.Item("Background View")
    Dim TNP_VueSym As DrawingView
    'Set TNP_VueC = Coll_Vues.Item("")

'Recherche du format du calque et paramétrage de la position du tableau
    Dim CalqueActifPaperSize As String
    CalqueActifPaperSize = CalqueActif.PaperName
    Dim Dim_Calque_X, Dim_Calque_Y As Integer
    If CalqueActifPaperSize = "A0 ISO" Then
        Dim_Calque_X = 1189
        Dim_Calque_Y = 841
    ElseIf CalqueActifPaperSize = "A1 ISO" Then
        Dim_Calque_X = 841
        Dim_Calque_Y = 594
    ElseIf CalqueActifPaperSize = "A2 ISO" Then
        Dim_Calque_X = 594
        Dim_Calque_Y = 420
    ElseIf CalqueActifPaperSize = "A3 ISO" Then
        Dim_Calque_X = 420
        Dim_Calque_Y = 297
    End If

'Mise a jour du numero de planche
    Dim Item
    For Each Item In Vue_Back.Texts
        If Item.Name = "Texte.sheet" Then
            If NoSheetEG = "01" Then
                Item.Text = NoSheetEG & "/" & NoTotSheet
            Else
                Item.Text = NoSheetEG
            Else
        End If
    Next
    
'Recupération du tableau de nomenclature dans le drawing
    Set Coll_Tables = Vue_Back.Tables
    Dim TablNom2D, TablTitres As DrawingTable
    For i = 1 To 2
        If Coll_Tables.Item(i).Name = "TableauNom" Then Set TablNom2D = Coll_Tables.Item(i)
        If Coll_Tables.Item(i).Name = "TableauTitre" Then Set TablTitres = Coll_Tables.Item(i)
    Next i

'Remplissage et formatage du tableau de nomenclature
    TablNom2D.AnchorPoint = CatTableBottomRight
    TablNom2D.X = Dim_Calque_X - 90
    TablNom2D.Y = 170
    ProgressBar (10)
    
    'Supprime les lignes existantes sauf la ligne des entètes
    For i = TablNom2D.NumberOfRows To 2 Step -1
        TablNom2D.RemoveRow i - 1
    Next

    'Création du nombre de ligne nécéssaire dans le Tableau 2D
    '+1 ligne pour les titres
    NBLigneTablNomProd = UBound(TablNomProduct(), 2)
    For i = 1 To NBLigneTablNomProd + 1
        If TablNom2D.NumberOfRows <= i Then TablNom2D.AddRow (1)
    Next i
    
    'Suppression du nombre de colonne en trop (cas d'une régénération d'une nom prééxistante)
    While TablNom2D.NumberOfColumns > 10
        TablNom2D.RemoveColumn 1
    Wend
    
    'Ajout du nombre de colonnes correspondant au sym ou aux variantes
    If NbColPlus > 0 Then
        For i = 1 To NbColPlus
            TablNom2D.AddColumn (i)
            TablNom2D.SetCellString TablNom2D.NumberOfRows, 1, "QTY"
            TablNom2D.GetCellObject(TablNom2D.NumberOfRows, 1).SetFontSize 0, 0, 2.5
            TablNom2D.SetCellAlignment TablNom2D.NumberOfRows, 1, CatTableMiddleCenter
        Next i
    End If

    Dim Tmp_NoSheet As String, tmp_Value As String
    For i = 1 To NBLigneTablNomProd + 1
    ProgressBar (10 + (90 / (NBLigneTablNomProd) * SupDivZero(i)))
        For j = 1 To NbColPlus + 10
                If j = 2 + NbColPlus Then 'N° de planche active pour les éléments du commerce
                    Tmp_NoSheet = NoSheetSupplier(TablNomProduct(j, NBLigneTablNomProd + 1 - i), TablNomProduct(j - 1, NBLigneTablNomProd + 1 - i), NoSheetEG, TypeNum)
                    TablNom2D.SetCellString i, j, Tmp_NoSheet
                Else
                    tmp_Value = TablNomProduct(j - 1, NBLigneTablNomProd + 1 - i)
                    On Error Resume Next
                    TablNom2D.SetCellString i, j, tmp_Value
                End If
                If j = 5 + NbColPlus Then
                    TablNom2D.SetCellAlignment i, j, CatTableMiddleLeft
                Else
                    TablNom2D.SetCellAlignment i, j, CatTableMiddleCenter
                End If
                
                If j <= (1 + NbColPlus) Or j = 3 + NbColPlus Then
                    TablNom2D.GetCellObject(i, j).SetFontSize 0, 0, 4.2
                ElseIf j = 2 + NbColPlus And Len(Tmp_NoSheet) <= 3 Then
                    TablNom2D.GetCellObject(i, j).SetFontSize 0, 0, 4.2
                Else
                    TablNom2D.GetCellObject(i, j).SetFontSize 0, 0, 3.5
                End If
        Next j
    Next i

'Remplissage du tableau des Titres
    TablTitres.AnchorPoint = CatTableTopRight
    TablTitres.X = Dim_Calque_X - 462
    TablTitres.Y = 170
    
    'Suppression du nombre de lignes en trop (cas d'une régénération d'une nom prééxistante)
    While TablTitres.NumberOfRows > 2
        TablTitres.RemoveRow 1
    Wend
    
    'Ajout des lignes en plus pour le nom des ensembles
    TablTitres.SetCellString 1, 1, TablNomProduct(NbColPlus + 3, 0)
    TablTitres.SetCellAlignment 1, 1, CatTableMiddleCenter
    If NbColPlus > 0 Then
        For i = 1 To NbColPlus
            TablTitres.AddRow (i)
            TablTitres.SetCellString 1, 1, TablNomProduct(NbColPlus + 3, 1)
            TablTitres.SetCellAlignment 1, 1, CatTableMiddleCenter
        Next i
    End If
    
    'Modification de la position du tableau
    TablTitres.X = TablTitres.X - (NbColPlus * 12)
    
'Si c'est un symétrique, on ajoute un texte

    If UBound(ListProd) = 1 Then
        Dim TNP_Y As Double
            TNP_Y = TablNom2D.Y
        Dim TNP_X As Double
            TNP_X = TablNom2D.X
        Dim Sym As Boolean
            Sym = False
        If CInt(Mid(ListProd(0), 12, 3)) - CInt(Mid(ListProd(1), 12, 3)) = 1 Then
            ListProd(0) = "-" & Left(ListProd(0), InStr(ListProd(0), ".") - 1) & " NOT REPRESENTED"
            ListProd(1) = "-" & Left(ListProd(1), InStr(ListProd(1), ".") - 1) & " DRAWN"
            Sym = True
        ElseIf CInt(Mid(ListProd(0), 12, 3)) - CInt(Mid(ListProd(1), 12, 3)) = -1 Then
            ListProd(0) = "-" & Left(ListProd(0), InStr(ListProd(0), ".") - 1) & " DRAWN"
            ListProd(1) = "-" & Left(ListProd(1), InStr(ListProd(1), ".") - 1) & " NOT REPRESENTED"
            Sym = True
        End If
        
        'teste l'existance d'un nota
        For i = 1 To Vue_Back.Texts.Count
            If Vue_Back.Texts.Item(i).Name = "TxtNotaSym" Then
                NotaSymExiste = True
            End If
        Next i
        If Sym And Not (NotaSymExiste) Then
            Dim TxtNotaSym As DrawingText
            'Set TNP_VueSym = Coll_Vues.Add("NotaSym")
            TNP_X = TNP_X - 375
            TNP_Y = TNP_Y + 70 + TablNom2D.NumberOfRows * 10
            'Set TxtNotaSym = TNP_VueSym.Texts.Add(ListProd(0) & Chr(10) & ListProd(1), TNP_X, TNP_Y)
            Set TxtNotaSym = Vue_Back.Texts.Add(ListProd(0) & Chr(10) & ListProd(1), TNP_X, TNP_Y)
            TxtNotaSym.TextProperties.FONTSIZE = 5
            TxtNotaSym.Name = "TxtNotaSym"
        End If
    End If
      
'Ajout des notas
    Dim CalqueDet As DrawingSheet
    'collection des Dito du calque de détails
    If Coll_Calques.Item(2).Name = "Calque.2 (Détail)" Then 'pour les plans Français
        Set CalqueDet = Coll_Calques.Item("Calque.2 (Détail)")
    ElseIf Coll_Calques.Item(2).Name = "Sheet.2 (Detail)" Then 'pour les plans anglais
        Set CalqueDet = Coll_Calques.Item("Sheet.2 (Detail)")
    End If
    Set Coll_Vues = CalqueDet.Views
    Dim DetailsInstancies As DrawingComponents
    Set DetailsInstancies = Vue_Back.Components
    
'Ajout du Nota Mecano-Soudé

    'detection du nota s'il existe
    For i = 1 To DetailsInstancies.Count
        If Left(DetailsInstancies.Item(i).Name, Len("WELDING_NOTE")) = "WELDING_NOTE" Then
            NotaSoudExiste = True
        End If
    Next i
    
    If Tmp_MecanoSoude = "OUI" And Not (NotaSoudExiste) Then
        Dim PosNotaWelding_X, PosNotaWelding_Y As Integer
            PosNotaWelding_X = Dim_Calque_X - 425
            PosNotaWelding_Y = 80
        Dim DetNotaWeldingSource As DrawingView
        Dim DetNotaWeldingCible As DrawingComponent
        Set DetNotaWeldingSource = Coll_Vues.Item("WELDING_NOTE")
        Set DetNotaWeldingCible = DetailsInstancies.Add(DetNotaWeldingSource, PosNotaWelding_X, PosNotaWelding_Y)
    End If

'Ajout des cages Codes
    If TypeElement(TNP_Product.Name, TypeNum) = 1 Then
        Dim DrawTxtEC As DrawingText
        Dim TxtEC As String
        'Ajout du Titre
        TxtEC = "SUPPLIER"
        Dim PosNotaSupplier_X As Integer, PosNotaSupplier_y As Integer
        PosNotaSupplier_X = 26 + 150
        PosNotaSupplier_y = Dim_Calque_Y - 30
          
        'Ajout des lignes de fournisseurs
        For i = 0 To UBound(ListCageCode, 2)
            If ListCageCode(0, i) <> "" And ListCageCode(0, i) <> " " Then
                TxtEC = TxtEC & Chr(10)
                TxtEC = TxtEC & Chr(10)
                TxtEC = TxtEC & ListCageCode(0, i)
                TxtEC = TxtEC & Chr(10)
                TxtEC = TxtEC & "CAGE CODE : " & ListCageCode(1, i)
            End If
        Next
        Set DrawTxtEC = Vue_Back.Texts.Add(TxtEC, PosNotaSupplier_X, PosNotaSupplier_y)
        'Mise en forme
        DrawTxtEC.SetFontSize 1, 9, 8
        DrawTxtEC.SetFontSize 10, 0, 5
        DrawTxtEC.TextProperties.Justification = catCenter
        DrawTxtEC.SetParameterOnSubString catUnderline, 1, 9, 1
        
    End If
 Unload Frm_Progression

QuitTraceNom2DProduct:
Exit Sub
ErrTraceNom2DProduct:
    MsgBox Err.Number & " - " & Err.Description
    'Resume Next
End Sub
 
Public Sub TraceNom2DPart(TN2P_ListPart() As String)
' *****************************************************************
' * Documente la nomenclature du Part sur le 2D
' *
' *
' * Création CFR le 05/11/2012
' * Dernière modification le
' *****************************************************************

'Documents
    Set Coll_Documents = CATIA.Documents
    Dim TN2P_PartDoc As PartDocument
    Dim TN2P_Product As Product
    
'Drawing
    Dim TN2P_Drawing As DrawingDocument
    Set TN2P_Drawing = CATIA.ActiveDocument
    Dim TN2P_Calque As DrawingSheet
    Set TN2P_Calque = TN2P_Drawing.Sheets.activeSheet
    Dim TN2P_Vues As DrawingViews
    Set TN2P_Vues = TN2P_Calque.Views
    Dim TN2P_Vue As DrawingView
    Set TN2P_Vue = TN2P_Vues.Item("Background View")
    
'Texts
    Dim TxtNotaSym As DrawingText
        
'Recupération du tableau de nomenclature
    Dim TN2P_Tables As DrawingTables
    Set TN2P_Tables = TN2P_Vue.Tables
    Dim TN2P_TableauNom2D As DrawingTable
    Set TN2P_TableauNom2D = TN2P_Tables.Item(1)
    Dim TN2P_Y As Double
        TN2P_Y = TN2P_TableauNom2D.Y
    Dim TN2P_X As Double
        TN2P_X = TN2P_TableauNom2D.X
        
    Dim i As Integer, j As Integer

'Suppression des lignes si nomenclature déja présente
    For i = TN2P_TableauNom2D.NumberOfRows - 1 To 2 Step -1
        TN2P_TableauNom2D.RemoveRow i
        TN2P_TableauNom2D.Y = TN2P_Y - 10.03
        TN2P_Y = TN2P_TableauNom2D.Y
    Next i

'Suppression du nota s'il existe
    For i = 1 To TN2P_Vues.Count
        If TN2P_Vues.Item(i).Name = "NotaSym" Then
            TN2P_Vues.Remove (i)
        End If
    Next i
     
'Pour chaque Part de la liste
For i = 0 To UBound(TN2P_ListPart, 1)
    'For j = 1 To Coll_Documents.Count
        'If Coll_Documents.Item(j).Name = TN2P_ListPart(i) Then
        'Set TN2P_PartDoc = Coll_Documents.Item(j)
        On Error Resume Next
        Set TN2P_PartDoc = Coll_Documents.Item(TN2P_ListPart(i))
        If Err.Number <> 0 Then
            Err.Clear
            MsgBox "Le Part : " & TN2P_ListPart(i) & " n'est pas chargé, veuillez ouvri le Part du détail ou l'assemblage général.", vbCritical, "Fichier non chargé"
            End
        End If
        On Error GoTo 0
        'End If
    'Next j
    Set TN2P_Product = TN2P_PartDoc.Product
    Dim TN2P_Parametres As Parameters
    Set TN2P_Parametres = TN2P_Product.Parameters
     
'Recupération des paramètres
    Dim Param_Dimension As StrParam
    Dim Param_Material As StrParam
    Dim Param_Protect As StrParam
    Dim Param_Sheet As StrParam
    
    Set Param_Dimension = TN2P_Parametres.Item("NomPulsGSE_Dimension")
    Set Param_Material = TN2P_Parametres.Item("NomPulsGSE_Material")
    Set Param_Protect = TN2P_Parametres.Item("NomPulsGSE_Protect")
    Set Param_Sheet = TN2P_Parametres.Item("NomPulsGSE_Sheet")
    
'Création d'une ligne si plusieur part
    If i > 0 Then
        TN2P_TableauNom2D.AddRow (1)
        TN2P_TableauNom2D.Y = TN2P_Y + 10.03
    End If
Dim toto
'Remplissage de la nom + Mise en forme
    TN2P_TableauNom2D.SetCellString 1, 1, SautLigne(TN2P_Product.PartNumber)
    TN2P_TableauNom2D.SetCellAlignment 1, 1, CatTableMiddleCenter
    TN2P_TableauNom2D.SetCellString 1, 2, SautLigne(TN2P_Product.DescriptionRef)
    TN2P_TableauNom2D.SetCellAlignment 1, 2, CatTableMiddleLeft
    TN2P_TableauNom2D.SetCellString 1, 3, SautLigne(Param_Dimension.Value)
    TN2P_TableauNom2D.SetCellAlignment 1, 3, CatTableMiddleCenter
    TN2P_TableauNom2D.SetCellString 1, 4, SautLigne(Param_Material.Value)
    TN2P_TableauNom2D.SetCellAlignment 1, 4, CatTableMiddleCenter
    TN2P_TableauNom2D.SetCellString 1, 5, SautLigne(Param_Protect.Value)
    TN2P_TableauNom2D.SetCellAlignment 1, 5, CatTableMiddleCenter
Next

'Si c'est un symétrique, on ajoute un texte
    If UBound(TN2P_ListPart) = 1 Then
        Dim Sym As Boolean
            Sym = False
        If CInt(Mid(TN2P_ListPart(0), 12, 3)) - CInt(Mid(TN2P_ListPart(1), 12, 3)) = 1 Then
            TN2P_ListPart(0) = "-" & Left(TN2P_ListPart(0), InStr(TN2P_ListPart(0), ".") - 1) & " SYMMETRICAL NOT REPRESENTED"
            TN2P_ListPart(1) = "-" & Left(TN2P_ListPart(1), InStr(TN2P_ListPart(1), ".") - 1) & " DRAWN"
            Sym = True
        ElseIf CInt(Mid(TN2P_ListPart(0), 12, 3)) - CInt(Mid(TN2P_ListPart(1), 12, 3)) = -1 Then
            TN2P_ListPart(0) = "-" & Left(TN2P_ListPart(0), InStr(TN2P_ListPart(0), ".") - 1) & " DRAWN"
            TN2P_ListPart(1) = "-" & Left(TN2P_ListPart(1), InStr(TN2P_ListPart(1), ".") - 1) & " SYMMETRICAL NOT REPRESENTED"
            Sym = True
        End If
        If Sym Then
            TN2P_X = TN2P_X
            TN2P_Y = TN2P_Y + 30
            Set TxtNotaSym = TN2P_Vue.Texts.Add(TN2P_ListPart(0) & Chr(10) & TN2P_ListPart(1), TN2P_X, TN2P_Y)
            TxtNotaSym.TextProperties.FONTSIZE = 5
            TxtNotaSym.TextProperties.Bold = True
            TxtNotaSym.Name = "TxtNotaSym"
            TxtNotaSym.TextProperties.Update
        End If
    End If
    
'Maj du numero de planche
    Dim Item
        For Each Item In TN2P_Vue.Texts
            If Item.Name = "Texte.sheet" Then
                Item.Text = Param_Sheet.Value
            End If
        Next

QuitTraceNom2DPart:
Exit Sub
ErrTraceNom2DPart:
MsgBox Err.Number
End Sub


Public Function CompileNom(CN_Workbook, ListeEns, SymVar, CN_MecanoSoude, CN_SiteAB As String) As String()
'Compile dans un même tableau les nomenclatures des plusieurs ensembles
Dim CN_LigActive As Long
Dim CN_LigneTemp As String
Dim NoEndNom  As Integer
Dim NomEns As String
Dim CN_TableauNom() As String
Dim CN_TableauCompil() As String
Dim i As Long, j As Integer, NewLine As Long
Dim NbEns As Integer
    NbEns = UBound(ListeEns)
Dim NoEns As Integer
    NoEns = 1
'pour chaque ensemble
    For i = 0 To NbEns
        'suppression de l'extension ".CATProduct"
        NomEns = Left(ListeEns(i), InStr(1, ListeEns(i), ".", vbTextCompare) - 1)
        NoEndNom = 0
        CN_LigActive = 1
        'Pointage sur la premiere ligne des composants du product a analyser
        CN_LigneTemp = ""
        Do While Not CN_LigneTemp = NomEns And NoEndNom <= 2
            CN_LigActive = CN_LigActive + 1
            'On recherche dans la colone 1 (Nomenclature de xxxx)
            CN_LigneTemp = Right(CN_Workbook.activeSheet.cells(CN_LigActive, 1).Value, Len(NomEns))
            'Si 2 lignes vides consecutive => EOF du fichier excel
            If CN_Workbook.activeSheet.cells(CN_LigActive, 1).Value = "" Then
                NoEndNom = NoEndNom + 1
            Else
                NoEndNom = 0
            End If
        Loop
        CN_LigActive = CN_LigActive + 2 'saut des entètes
        
        'Ajout des lignes de nomenclature
        Do While Not CN_Workbook.activeSheet.cells(CN_LigActive, 1).Value = ""
        
            ReDim Preserve CN_TableauNom(10, NewLine)
            CN_TableauNom(0, NewLine) = CN_Workbook.activeSheet.cells(CN_LigActive, 1).Value 'Qte
            CN_TableauNom(1, NewLine) = FormatNoSheet(CStr(CN_Workbook.activeSheet.cells(CN_LigActive, 2).Value), CN_SiteAB) 'Sheet
            CN_TableauNom(2, NewLine) = Txt3Digit(CN_Workbook.activeSheet.cells(CN_LigActive, 3).Value) 'Item Nbr
            CN_TableauNom(3, NewLine) = SautLigne(CN_Workbook.activeSheet.cells(CN_LigActive, 5).Value) 'Part Nbr
            CN_TableauNom(4, NewLine) = SautLigne(CN_Workbook.activeSheet.cells(CN_LigActive, 6).Value) 'Description
            'cas des mécano soudés
            If CN_MecanoSoude = "OUI" Or TypeElement(CN_Workbook.activeSheet.cells(CN_LigActive, 4).Value, TypeNum) = 9 Then
                CN_TableauNom(5, NewLine) = SautLigne(CN_Workbook.activeSheet.cells(CN_LigActive, 7).Value) 'Dimension
                CN_TableauNom(6, NewLine) = SautLigne(VigPt(CN_Workbook.activeSheet.cells(CN_LigActive, 8).Value)) 'Material
                CN_TableauNom(7, NewLine) = SautLigne(CN_Workbook.activeSheet.cells(CN_LigActive, 9).Value) 'Protect
            Else
                CN_TableauNom(5, NewLine) = "" 'Dimension
                CN_TableauNom(6, NewLine) = "" 'Material"
                CN_TableauNom(7, NewLine) = "" 'Protect
            End If
                CN_TableauNom(8, NewLine) = SautLigne(CN_Workbook.activeSheet.cells(CN_LigActive, 10).Value) 'Miscellaneous
                CN_TableauNom(9, NewLine) = CN_Workbook.activeSheet.cells(CN_LigActive, 11).Value 'Weights
                CN_TableauNom(10, NewLine) = NoEns
            NewLine = NewLine + 1
            CN_LigActive = CN_LigActive + 1
        Loop
        NoEns = NoEns + 1
    Next
    
'tri du tableau sur la colone ItemNumber pour regrouper les Item
CN_TableauNom = TranspositionTabl(CN_TableauNom())
CN_TableauNom = TriList2D(CN_TableauNom(), 2, True)
CN_TableauNom = TranspositionTabl(CN_TableauNom())
Dim tmp As String

'Decalage de colonne
    NewLine = -1
    'Nombre de colonnes
    Dim CN_NBCol As Integer, NBColCompil As Integer
    CN_NBCol = 10 - 1
    NBColCompil = 10 + NbEns - 1
    Dim ItenNBEC As String
        ItenNBEC = ""
   
   'si c'est un sym
   If UBound(ListeEns, 1) > 0 Then
        For i = 0 To UBound(CN_TableauNom, 2)
            NoEns = CN_TableauNom(10, i)
            If CN_TableauNom(2, i) = ItenNBEC Then 'on as déja renseigné cette ligne
                 If NoEns = 1 Then
                    'Ajout de la Quantité dans la colone du sym
                    CN_TableauCompil(1, NewLine) = CN_TableauNom(0, i)
                Else
                    CN_TableauCompil(0, NewLine) = CN_TableauNom(0, i)
                End If
                
            Else 'c'est un nouvel item, on ajoute une ligne
                'on vérifie si la ligne appartiens a l'ensemble de base(colonne 10 = 1 ou au sym (colonne 10 = 2)
                NewLine = NewLine + 1
                ReDim Preserve CN_TableauCompil(NBColCompil, NewLine)
                If NoEns = 1 Then
                    'Ajout de la valeur du champs Qte dans la colonne de base (2)
                    CN_TableauCompil(1, NewLine) = CN_TableauNom(0, i)
                Else
                    'Ajout de la valeur du champs Qte dans la colonne du sym (1)
                    CN_TableauCompil(0, NewLine) = CN_TableauNom(0, i)
                End If
                'Récupèration du reste de la ligne
                For j = 1 To CN_NBCol
                    CN_TableauCompil(j + NbEns, NewLine) = CN_TableauNom(j, i)
                Next
                ItenNBEC = CN_TableauNom(2, i)
                
            End If
        Next
    Else
    
    'si c'est un seul ensemble
        NewLine = 0
        For i = 0 To UBound(CN_TableauNom, 2)
            ReDim Preserve CN_TableauCompil(NBColCompil, NewLine)
            For j = 0 To CN_NBCol
                CN_TableauCompil(j + NbEns, NewLine) = CN_TableauNom(j, i)
            Next
            NewLine = NewLine + 1
        Next i
    End If

CompileNom = CN_TableauCompil()
End Function

Private Function MaxSheet(mTable, SiteAB As String) As String
'Renvoi le N° de la planche maxi
Dim SheetMax As Integer
Dim i As Long
    
    MaxSheet = 0
    For i = 0 To UBound(mTable, 2)
        If IsNumeric(mTable(1, i)) Then
            If CInt(mTable(1, i)) > MaxSheet Then
                SheetMax = CInt(mTable(1, i))
            End If
        End If
    Next
    MaxSheet = FormatNoSheet(CStr(SheetMax), SiteAB)
End Function
