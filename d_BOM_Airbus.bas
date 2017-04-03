Attribute VB_Name = "d_BOM_Airbus"
Option Explicit

Sub catmain()
' *****************************************************************
' * Extraction de la BOM AIRBUS Type A350
' *
' *
' * Création CFR le 21/09/2015
' * Version 2.9
' * Dernière modification le :
' *
' *****************************************************************
'On Error Resume Next

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "d_BOM_Airbus", VMacro

'Chargement des variables
Dim MacroLocation As New xMacroLocation
If Not (MacroLocation.FicIniExist("VarNomenclatureGSE.ini")) Then ' on est pas dans l'environnement GSE_R21
    MsgBox "Vous n'etes pas dans l'environnement GRE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
    Exit Sub
Else
    MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
    CheminDestNomenclature = MacroLocation.ValVar("CheminDestNomenclature")
    CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
    Nom_TemplateAirbus = MacroLocation.ValVar("TemplateBomA350")
End If

Dim Lig_EC As Long, i As Long, j As Long
    Lig_EC = 3
    
Dim Ens_GeneralDoc As Document
Set Ens_GeneralDoc = CATIA.ActiveDocument
Dim Ens_GeneralProduct As Product
Dim PartProd As Product
Set Ens_GeneralProduct = Ens_GeneralDoc.Product
    
'  paramètres de l'ensemble sélectionné
Dim paramsGSe As Parameters
Set paramsGSe = Ens_GeneralProduct.UserRefProperties
'Type de Numérotation
'NomPulsGSE_TypeNum
TypeNum = RecupParam(paramsGSe, "NomPulsGSE_TypeNum")
    
'  'Détection Caisse
'  Dim Tmp_Caisse As String
'    Tmp_Caisse = DetectCaisse(Ens_GeneralProduct)
      
Dim CatiaAnglais As Integer
'Intérogation du User sur la langue de Catia
CatiaAnglais = MsgBox("Votre catia est il en anglais ?", vbYesNo, "Choix du language")

'Variable temp de traduction des descriptions
Dim LangueQt, LangueRef, LangueDesc As String, LangueGroupeNom As String, LangueRecapPiece As String

'Formatage de la nomenclature en fonction de la langue du Catia
If CatiaAnglais = vbYes Then
    LangueQt = "Quantity"
    LangueRef = "Part Number"
    LangueDesc = "Product Description"
    LangueGroupeNom = "Bill of Material: "
    LangueRecapPiece = "Recapitulation of: "
    
ElseIf CatiaAnglais = vbNo Then
    LangueQt = "Quantité"
    LangueRef = "Référence"
    LangueDesc = "Description du produit"
    LangueGroupeNom = "Nomenclature de "
    LangueRecapPiece = "Récapitulatif sur"
Else
    MsgBox "Erreur dans la détection de la langue paramétré dans Catia."
End If
    
 'verifie si un fichier de nomenclature est déja présent et l'efface
    If Not (EffaceFicNom(CheminDestNomenclature, Ens_GeneralProduct.Name & ".xls")) Then
        End
    End If

'Extraction de la nomenclature du product général et sauvegarde dans un fichier excel
    Dim assemblyConvertor1 As AssemblyConvertor
    Set assemblyConvertor1 = Ens_GeneralProduct.GetItem("BillOfMaterial")
    Dim assemblyConvertor1Variant

    Dim arrayOfVariantOfBSTR1(10)
    arrayOfVariantOfBSTR1(0) = CStr(LangueQt)
    arrayOfVariantOfBSTR1(1) = "NomPulsGSE_Sheet"
    arrayOfVariantOfBSTR1(2) = "NomPulsGSE_ItemNb"
    arrayOfVariantOfBSTR1(3) = CStr(LangueRef)
    arrayOfVariantOfBSTR1(4) = "NomPulsGSE_SupplierRef"
    arrayOfVariantOfBSTR1(5) = CStr(LangueDesc)
    arrayOfVariantOfBSTR1(6) = "NomPulsGSE_Dimension"
    arrayOfVariantOfBSTR1(7) = "NomPulsGSE_Material"
    arrayOfVariantOfBSTR1(8) = "NomPulsGSE_Protect"
    arrayOfVariantOfBSTR1(9) = "NomPulsGSE_Miscellanous"
    arrayOfVariantOfBSTR1(10) = "NomPulsGSE_Weight"

    Set assemblyConvertor1Variant = assemblyConvertor1
    assemblyConvertor1Variant.SetCurrentFormat arrayOfVariantOfBSTR1

    Dim arrayOfVariantOfBSTR2(10)
    arrayOfVariantOfBSTR2(0) = CStr(LangueQt)
    arrayOfVariantOfBSTR2(1) = "NomPulsGSE_Sheet"
    arrayOfVariantOfBSTR2(2) = "NomPulsGSE_ItemNb"
    arrayOfVariantOfBSTR2(3) = CStr(LangueRef)
    arrayOfVariantOfBSTR2(4) = "NomPulsGSE_SupplierRef"
    arrayOfVariantOfBSTR2(5) = CStr(LangueDesc)
    arrayOfVariantOfBSTR2(6) = "NomPulsGSE_Dimension"
    arrayOfVariantOfBSTR2(7) = "NomPulsGSE_Material"
    arrayOfVariantOfBSTR2(8) = "NomPulsGSE_Protect"
    arrayOfVariantOfBSTR2(9) = "NomPulsGSE_Miscellanous"
    arrayOfVariantOfBSTR2(10) = "NomPulsGSE_Weight"

    Set assemblyConvertor1Variant = assemblyConvertor1
    assemblyConvertor1Variant.SetSecondaryFormat arrayOfVariantOfBSTR2
    assemblyConvertor1.[Print] "XLS", CStr(CheminDestNomenclature & Ens_GeneralProduct.Name & ".xls"), Ens_GeneralProduct

'Creation d'un objet eXcel et ouverture de la nomenclature précédement générée
    Dim xls_NomTemp
    Dim WorkBook_NomTemp
    Set xls_NomTemp = CreateObject("EXCEL.APPLICATION")
    Set WorkBook_NomTemp = xls_NomTemp.Workbooks.Open(CStr(CheminDestNomenclature & Ens_GeneralProduct.Name & ".xls"))
    xls_NomTemp.Visible = True
    WorkBook_NomTemp.ActiveSheet.Visible = True
    
'Construction de la liste des pièces et des niveaux
Dim SubToolRef1() As String
Dim Temp_N0_SE_Pere As String, Temp_GSE As String, Temp_PSE As String, Temp_N0_SE As String
Dim LigTabTemp As Long
LigTabTemp = 0
Dim col_SSE As Integer

ReDim Preserve SubToolRef1(13, LigTabTemp)
    'SubToolRef1(0, 0) = Parent
    'SubToolRef1(1, 0) = ToolRef (000)
    'SubToolRef1(2, 0) = Sub tool ref lvl1 (040)
    'SubToolRef1(3, 0) = Sub tool ref lvl2 (100)
    'SubToolRef1(4, 0) = Sub tool ref lvl3 (200, 500, 700)
    'SubToolRef1(5, 0) = Std or Norm or Vendor ref
    'SubToolRef1(6, 0) = Spare
    'SubToolRef1(7, 0) = Mat Group
    'SubToolRef1(8, 0) = Material
    'SubToolRef1(9, 0) = Protection
    'SubToolRef1(10, 0) = Designation
    'SubToolRef1(11, 0) = Provider
    'SubToolRef1(12, 0) = Micellanous

    '1er balayage pour récupérer toutes les lignes dans l'ordre de la nomenclature
    ' Boucle jusqu'a la ligne "Recapitulatif des pieces"
    Do While Not (DetectRecapPieces(WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 1).Value, LangueRecapPiece))
        
        'saute les lignes blanches et les entètes de colonnes
        If DetectLigVide(WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 1).Value) Or Detect_LigEntete(WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 1).Value, CStr(LangueQt)) Then ' Saut de la ligne d'entète du sous ensemble

        Else
        
            If ExtractNomSE(WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 1).Value, LangueGroupeNom) <> "" Then
                'Récupération du Nom du sous ensemble Père
                Temp_N0_SE_Pere = ExtractNomSE(WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 1).Value, LangueGroupeNom)
            Else
                
                ReDim Preserve SubToolRef1(13, LigTabTemp)
                SubToolRef1(0, LigTabTemp) = Temp_N0_SE_Pere
                SubToolRef1(1, LigTabTemp) = ""
                SubToolRef1(2, LigTabTemp) = ""
                SubToolRef1(3, LigTabTemp) = ""
                SubToolRef1(4, LigTabTemp) = WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 4).Value
                'récupération des champs de nomenclature
                SubToolRef1(5, LigTabTemp) = WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 5).Value 'SupplierRef
                SubToolRef1(6, LigTabTemp) = ""
                SubToolRef1(7, LigTabTemp) = ""
                SubToolRef1(8, LigTabTemp) = WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 8).Value 'Material
                SubToolRef1(9, LigTabTemp) = WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 9).Value 'Protection
                SubToolRef1(10, LigTabTemp) = WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 6).Value 'Designation
                SubToolRef1(11, LigTabTemp) = WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 10).Value 'Provider
                SubToolRef1(12, LigTabTemp) = WorkBook_NomTemp.ActiveSheet.cells(Lig_EC, 7).Value 'miscelanous
                LigTabTemp = LigTabTemp + 1
                
            End If
            
        End If
        Lig_EC = Lig_EC + 1
    Loop
    
    'Fermeture du fichier excel de lanomenclature temp
    WorkBook_NomTemp.Close
    
    '2eme balayage pour récupérer le parent N - 1 de chaque pièce
    For i = 0 To UBound(SubToolRef1, 2)

        Select Case TypeElement(SubToolRef1(0, i), TypeNum)
            Case "0" 'Product de tète
                'si c'est une caisse on la garde
                If TypeElement(SubToolRef1(4, i), TypeNum) = "5" Then
                    col_SSE = 3
                Else
                    
                    SubToolRef1(13, i) = "DEL" 'Ligne a supprimer
                End If
            Case "1", "2" 'Outillage et variantes (000)
                col_SSE = 1
            Case "3", "4" ' Grand SE (040)
                col_SSE = 2

            Case "5", "6" ' Petits SE (100)
                col_SSE = 3

            Case "7", "8", "9" 'Pièces (200)
                col_SSE = 4
            Case Else
                col_SSE = 13
        End Select
        'Copie le N° du Parent dans la bonne colonne
        SubToolRef1(col_SSE, i) = SubToolRef1(0, i)
        'Efface le parent
        'SubToolRef1(0, i) = ""
    Next
    
    '3eme balayage pour récuperer les parent de niveau 2 (parent des petits SSE 100)
    'Pour chaque SSE 100, on balaye la table de bas en haut pour rechercher le parent du SSE
     Dim NumOutillage As String
     For i = 0 To UBound(SubToolRef1, 2)
        Select Case TypeElement(SubToolRef1(3, i), TypeNum)
            Case "0" 'Product de tète
            
            'Récupération du 14 Digit pour l'affecter à la caisse
            For j = 0 To UBound(SubToolRef1, 2)
                If Len(SubToolRef1(0, j)) = 14 Then
                    NumOutillage = SubToolRef1(0, j)
                    Exit For
                End If
            Next j

                SubToolRef1(1, i) = NumOutillage
                SubToolRef1(3, i) = ""
            Case "5", "6" ' Petits SE (100)
                For j = i To 0 Step -1
                    If SubToolRef1(4, j) = SubToolRef1(3, i) Then
                        'Copie le N° du Parent dans la bonne colonne
                        SubToolRef1(2, i) = SubToolRef1(2, j)
                        Exit For
                    End If
                Next j
        End Select
    Next
    
    '4eme balayage pour récuperer les parent de niveau 1 (parent des grands SSE 040)
    'Pour chaque SSE 040, on balaye la table de bas en haut pour rechercher le parent du SSE
     For i = 0 To UBound(SubToolRef1, 2)
        Select Case TypeElement(SubToolRef1(2, i), TypeNum)
            Case "3", "4" ' Grand SE (040)
                For j = i To 0 Step -1
                    If SubToolRef1(4, j) = SubToolRef1(2, i) Then
                        'Copie le N° du Parent dans la bonne colonne
                        SubToolRef1(1, i) = SubToolRef1(0, j)
                        Exit For
                    End If
                Next j
        End Select
    Next
    
   '5eme balayage pour nétoyer le fichier
   'suppression des n° de de part dans la colonne "Std or Norm..." pour les 200
   'suppression de lignes en trop "Del" dans colonne 13
    For i = 0 To UBound(SubToolRef1, 2)
        'efface la première colonne
        'SubToolRef1(0, i) = ""
        Select Case TypeElement(SubToolRef1(4, i), TypeNum)
            Case "1", "2", "3", "4", "5", "6" 'Pas de ligne pour les ensembles
                SubToolRef1(13, i) = "DEL"
            Case "1", "2", "3", "4", "5", "6", "7", "8"    'tout sauf Pièces (200)
                SubToolRef1(5, i) = ""
        End Select
     Next
     
     'efface tous les champs d'une ligne déclarée a supprimer (Del dans champs 13)
     For i = 0 To UBound(SubToolRef1, 2)
        If SubToolRef1(13, i) = "DEL" Then
            For j = 0 To 13
                SubToolRef1(j, i) = ""
            Next j
        End If
    Next
    
'Export dans le template Airbus
 'Creation d'un objet eXcel et ouverture de la nomenclature précédement générée
    Dim xls_TemplateAirbus
    Dim WorkBook_TemplateAirbus
    Set xls_TemplateAirbus = CreateObject("EXCEL.APPLICATION")
    Set WorkBook_TemplateAirbus = xls_TemplateAirbus.Workbooks.Open(CStr(CheminSourcesMacro & Nom_TemplateAirbus))
    xls_TemplateAirbus.Visible = True
    WorkBook_TemplateAirbus.ActiveSheet.Visible = True
    Dim DebCol_Template As Integer, DebLig_Template As Integer, FinLig_Template As Long
    DebCol_Template = 1
    DebLig_Template = 5
    FinLig_Template = 5
    LigTabTemp = 0
    Dim ValCell As String
    
    For i = 0 To UBound(SubToolRef1, 2)
        'saut des lignes vides
        If SubToolRef1(4, i) <> "" Then
            For j = 1 To 12
            ValCell = "'" & SuprSautLigne(SubToolRef1(j, i))
                WorkBook_TemplateAirbus.ActiveSheet.cells(LigTabTemp + DebLig_Template, j + DebCol_Template).Value = ValCell
                If Replace(SubToolRef1(j, i), " ", "") = "" Then
                    'Grisage des cellules vides
                    CouleurCell WorkBook_TemplateAirbus.ActiveSheet, j + DebCol_Template, LigTabTemp + DebLig_Template, "gris"
                End If
            Next j
            LigTabTemp = LigTabTemp + 1
            FinLig_Template = FinLig_Template + 1
        End If
    
    Next
    
     'Ligne jaune a la fin
    For i = 1 To 12
        CouleurCell WorkBook_TemplateAirbus.ActiveSheet, i + DebCol_Template, FinLig_Template, "jaune"
    Next i
   
 'Sauvegarde
 WorkBook_TemplateAirbus.SaveAs (CStr(CheminDestNomenclature & "BOM_" & Ens_GeneralProduct.Name & ".xls"))
  
  MsgBox "Penser à modifier la BOM si il y a des standards modifiés ou des variantes", vbInformation
  
    
End Sub



Public Function ExtractNomSE(Val_Cell As String, NomCel_Langue As String) As String
'Detecte si le contenu de la cellule commence par "Nomenclature de" pour les catia farnçais ou "Nomenclature of" pour les catia anglais.
'Si c'est le cas, renvoi le nom du sous ensemble, sinon renvoi une chaine vide
On Error Resume Next
    ExtractNomSE = ""
If Left(Val_Cell, Len(NomCel_Langue)) = NomCel_Langue Then
    If Err.Number <> 0 Then
        Err.Clear
        ExtractNomSE = ""
    Else
        ExtractNomSE = Right(Val_Cell, Len(Val_Cell) - Len(NomCel_Langue))
    End If
End If
On Error GoTo 0
End Function

 Public Function DetectLigVide(Val_Cell As String) As Boolean
 'Detecte si le contenu de la cellue est vide
 DetectLigVide = False
 If Len(Val_Cell) = 0 Then
    DetectLigVide = True
End If
 End Function

Public Function Detect_LigEntete(Val_Cell As String, NomCel_Langue As String) As Boolean
'Detecte si le contenu de la cellule est egal a "Quantité" pour les catia français ou à "Quantity" pour les catia anglais
Detect_LigEntete = False
If Val_Cell = NomCel_Langue Then
    Detect_LigEntete = True
End If

End Function

Public Function DetectRecapPieces(Val_Cell As String, NomCel_Langue As String) As Boolean
'Detecte si le contenu de la cellule commence par "Récapitulatif sur" pour les catia farnçais ou "Recap of" pour les catia anglais.
On Error Resume Next
    DetectRecapPieces = False
If Left(Val_Cell, Len(NomCel_Langue)) = NomCel_Langue Then
    If Err.Number <> 0 Then
        Err.Clear
        DetectRecapPieces = False
    Else
        DetectRecapPieces = True
    End If
End If
On Error GoTo 0
End Function
