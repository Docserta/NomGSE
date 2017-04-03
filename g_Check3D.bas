Attribute VB_Name = "g_Check3D"
Option Explicit
Public ListInst() As String


Sub catmain()
' *****************************************************************
' * Execution d'un check de certains critères sur tous les composants d'un assemblage
' *
' *
' * Création CFR le 19/02/2016
' * Version 4.8
' * Dernière modification le :
' *
' *****************************************************************

Dim MacroLocation As New xMacroLocation
    'Log l'utilisation de la macro
    LogUtilMacro nPath, nFicLog, nMacro, "g_Check3D", VMacro
    
Dim NomRapportCheck As String
Dim LEC_Recp As Long, LEC_Fich As Long, LEC_Arbo As Long, LEC_Body As Long, LEC_Contr As Long, LEC_Attrib As Long
Dim PremLig_Attrib As Long
Dim ErrIsInac  As String, ErrMAJ As String
Dim No_outillage As String
Dim i As Long, j As Long

'Part Body et HybridSahpes
Dim BodyEC As Body
Dim HbodyEC As HybridBody
Dim B_Shapes, H_Shapes
Dim B_Feature, H_Shape
Dim DocEC As Document

Dim NB_Err_Tri As Integer
Dim RadicalEC As String

Dim NivSSE As Integer

Dim SheetSel As String
Dim Col As String
Dim colnum As Long

'Les classes
Dim FicEC As Check3D
Dim resCheck As c_ResCheck3D
Dim ResChecks As c_ResCheck3Ds
Dim mBarre As c_ProgressBarre
Dim ColXlAttribs As c_ColXls

'Objet Excel
Dim objExcelCheck
Dim objWSRecap, objWSFichiers, objWSheetArbo, objWSheetBody, objWSheetAttrib, objWSContr
    
    'Collecte des variables d'environnement
    If Not (MacroLocation.FicIniExist("VarNomenclatureGSE.ini")) Then ' on est pas dans l'environnement GSE_R21
        MsgBox "Vous n'etes pas dans l'environnement GRE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
        Exit Sub
    Else
        MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
        CheminDestNomenclature = MacroLocation.ValVar("CheminDestNomenclature")
        CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
    End If

    Set Coll_Documents = CATIA.Documents

'Check de l'environnement
    On Error Resume Next
    Set ActiveDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "Aucun Catproduct n'est ouvert. Ouvrez le CATProduct d'un outillage avant de lancer la macro.", vbCritical, "Environnement incorect"
        End
    Else
        If Not (CheckProduct(ActiveDoc)) Then
            MsgBox "Le document actif n'est pas un Catproduct. Ouvrez le CATProduct d'un outillage avant de lancer la macro.", vbCritical, "Environnement incorect"
            End
        End If
    End If

    No_outillage = Left(ActiveDoc.Name, InStr(1, ActiveDoc.Name, ".CATProduct") - 1)
    RadicalEC = ActiveDoc.Name
    Set FicEC = New Check3D
    
'Ouverture boite de dialogue
    Load FRM_Check3D
    FRM_Check3D.Tbx_No_Outil = No_outillage
    FRM_Check3D.Show
    If Not (FRM_Check3D.ChB_OkAnnule) Then
        End
    End If
       
    CheminDestRapport = CheminDestNomenclature
    NomRapportCheck = "Check_" & FRM_Check3D.Tbx_No_Outil & Date & ".xlsx"
    NomRapportCheck = Replace(NomRapportCheck, "/", "-", , , vbTextCompare)

 'verifie si un fichier de rapport est déja présent et l'efface
    If Not (EffaceFicNom(CheminDestRapport, NomRapportCheck)) Then
        End
    End If

'Initialisation des resultats des checks
    Set ResChecks = InitResCheck()

'Initialisation de la barre de progression
    Set mBarre = New c_ProgressBarre
    mBarre.ProgressTitre 1, "Check 3D"
    mBarre.Affiche
    
'Création de la trame Excel
    Set objExcelCheck = CreateObject("EXCEL.APPLICATION")
    objExcelCheck.Visible = True
    objExcelCheck.Workbooks.Add
    Set objWSContr = InitWShtContr(objExcelCheck)
    LEC_Contr = 4
    Set objWSFichiers = InitWShtFichiers(objExcelCheck)
    LEC_Fich = 4
    Set objWSheetArbo = InitWShtArbo(objExcelCheck)
    LEC_Arbo = 5
    Set objWSheetBody = InitWShtPartBody(objExcelCheck)
    LEC_Body = 4
    Set ColXlAttribs = InitColAttrib
    Set objWSheetAttrib = InitWShtAttribs(objExcelCheck, ColXlAttribs)
    LEC_Attrib = 4
    PremLig_Attrib = 3
    Set objWSRecap = InitWShtRecap(objExcelCheck, ResChecks)
        objWSRecap.Range("C" & 2) = No_outillage
    LEC_Recp = 7
         
'Active le mode conception
    ActiveDoc.Product.ApplyWorkMode DESIGN_MODE

'#####################
' Nommage des fichiers
'#####################
   
    For i = 1 To Coll_Documents.Count
        mBarre.Progression = ((100 / Coll_Documents.Count) * i)
        Set DocEC = Coll_Documents.Item(i)
        On Error Resume Next
            FicEC.Charge3D = DocEC
            If Err.Number <> 0 Then GoTo Erreur1
        On Error GoTo 0
        
        ErrIsInac = ""
        ErrMAJ = ""
    
        'saute le product de l'outil général
        If Coll_Documents.Item(i).Name <> ActiveDoc.Name Then
            If EstPart(DocEC) Or EstProduct(DocEC) Then
                FicEC.initRad11Digt = ActiveDoc.Name
    
                'ecriture du nom du fichier en cours
                objWSFichiers.Range("A" & LEC_Fich) = FicEC.PN
    '##################################################################################
                'Controles de Nommage des Fichiers
                If FRM_Check3D.ChB_Nommage Then
                    'objWSFichiers.Activate
                'CK01 - Controle de la longueur du nom de fichier
                    Set resCheck = ResChecks.Item("CK01")
                    If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                        resCheck.Result = "Check"
                    End If
                    If FicEC.CK_LgNomFic Then
                        WriteOK objWSFichiers, "B", LEC_Fich, True
                    Else
                        WriteOK objWSFichiers, "B", LEC_Fich, False
                        objWSFichiers.Range("G" & LEC_Fich) = FicEC.NomFic
                        resCheck.Result = "KO"
                    End If
    
                'CK02 - Controle du radical du nom de fichier
                    Set resCheck = ResChecks.Item("CK02")
                    If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                        resCheck.Result = "Check"
                    End If
                    If FicEC.CK_Radical Then
                        WriteOK objWSFichiers, "C", LEC_Fich, True
                    Else
                        WriteOK objWSFichiers, "C", LEC_Fich, False
                        objWSFichiers.Range("F" & LEC_Fich) = FicEC.NomFic
                        resCheck.Result = "KO"
                    End If
    
                'CK03 - Controle que le PN, le nom d'instance et le file name sont identiques
                    Set resCheck = ResChecks.Item("CK03")
                    If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                        resCheck.Result = "Check"
                    End If
                    If FicEC.CK_NumEgal Then
                        WriteOK objWSFichiers, "D", LEC_Fich, True
                    Else
                        WriteOK objWSFichiers, "D", LEC_Fich, False
                        objWSFichiers.Range("F" & LEC_Fich) = FicEC.PN
                        objWSFichiers.Range("G" & LEC_Fich) = FicEC.NomFic
                        resCheck.Result = "KO"
                    End If
                'CK04 - Controle des Numéro Impaire
                    Set resCheck = ResChecks.Item("CK04")
                    If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                        resCheck.Result = "Check"
                    End If
                    If FicEC.CK_Impaire Then
                        WriteVerif objWSFichiers, "E", LEC_Fich, False
                        resCheck.Result = "KO"
                    Else
                        WriteVerif objWSFichiers, "E", LEC_Fich, True
                    End If
                End If
                LEC_Fich = LEC_Fich + 1
    
    '##################################################################################
                'Controles du Part Body
                If FRM_Check3D.ChB_PBody Then
                    'objWSheetBody.Activate
                    If FicEC.EstPart Then
                        'ecriture du nom du fichier en cours
                        objWSheetBody.Range("A" & LEC_Body) = FicEC.PN
    
                    'CK20 - Check le nombre de Corps de pièce pour les pièces fabriquées
                        Set resCheck = ResChecks.Item("CK20")
                        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                            resCheck.Result = "Check"
                        End If
                        If TypeElement(FicEC.PN, FicEC.Val_NomPulsGSE_TypeNum) >= 7 And TypeElement(FicEC.PN, FicEC.Val_NomPulsGSE_TypeNum) <= 8 Then
                            If FicEC.CK_NbBodies Then
                                WriteOK objWSheetBody, "F", LEC_Body, True
                            Else
                                WriteOK objWSheetBody, "F", LEC_Body, False
                                resCheck.Result = "KO"
                            End If
                        Else
                            objWSheetBody.Range("F" & LEC_Body) = "Nocheck"
                            CouleurCell objWSheetBody, "F", LEC_Body, "vert"
                        End If
    
                    'CK21-CK22- Check les fonction inactives ou non a jour dans le corps de pièce
                        For Each BodyEC In FicEC.Coll_bodies
                            Set B_Shapes = BodyEC.Shapes
                            For j = 1 To B_Shapes.Count
                                Set B_Feature = B_Shapes.Item(j)
                                If FicEC.PartEC.IsInactive(B_Feature) Then
                                    ErrIsInac = ErrIsInac & " - " & B_Feature.Name
                                End If
                                If FicEC.PartEC.IsUpToDate(B_Feature) Then
                                Else
                                    ErrMAJ = ErrMAJ & " - " & B_Feature.Name
                                End If
                            Next j
                        Next
                    'CK21-CK22-Check les fonction inactives ou non a jour dans les features
                        For Each HbodyEC In FicEC.Coll_Hbodies
                            Set H_Shapes = HbodyEC.HybridShapes
                            For j = 1 To H_Shapes.Count
                                Set H_Shape = H_Shapes.Item(j)
                                If FicEC.PartEC.IsInactive(H_Shape) Then
                                    ErrIsInac = ErrIsInac & " - " & H_Shape.Name
                                End If
                                If FicEC.PartEC.IsUpToDate(H_Shape) Then
                                Else
                                    ErrMAJ = ErrMAJ & " - " & H_Shape.Name
                                End If
                            Next j
                        Next
                        
                    'CK21 - Vérification des éléments non résolu
                        Set resCheck = ResChecks.Item("CK21")
                        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                            resCheck.Result = "Check"
                        End If
                        'Ecriture du rapport
                        If ErrIsInac <> "" Then
                            WriteOK objWSheetBody, "B", LEC_Body, False
                            objWSheetBody.Range("C" & LEC_Body) = ErrIsInac
                            CouleurCell objWSheetBody, "C", LEC_Body, "rouge"
                            resCheck.Result = "KO"
                        Else
                            WriteOK objWSheetBody, "B", LEC_Body, True
                        End If
                    'CK22 - Vérification des éléments non mis à jour
                        Set resCheck = ResChecks.Item("CK22")
                        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                            resCheck.Result = "Check"
                        End If
                        If ErrMAJ <> "" Then
                            WriteOK objWSheetBody, "D", LEC_Body, False
                            objWSheetBody.Range("E" & LEC_Body) = ErrMAJ
                            CouleurCell objWSheetBody, "E", LEC_Body, "rouge"
                            resCheck.Result = "KO"
                        Else
                            WriteOK objWSheetBody, "B", LEC_Body, True
                        End If
    
                    'CK24 - Check Part Body actif
                        Set resCheck = ResChecks.Item("CK24")
                        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                            resCheck.Result = "Check"
                        End If
                        'If FicEC.EstPart Then
                            If FicEC.CK_InWorkObj Then
                                WriteOK objWSheetBody, "G", LEC_Body, True
                            Else
                                WriteOK objWSheetBody, "G", LEC_Body, False
                                objWSheetBody.Range("G" & LEC_Body) = FicEC.InWorkObj.Name
                                resCheck.Result = "KO"
                            End If
                        'End If
    
                    'CK26 - Check Part hybride
                        Set resCheck = ResChecks.Item("CK26")
                        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                            resCheck.Result = "Check"
                        End If
                        If FicEC.CK_EstHybride Then
                            WriteOK objWSheetBody, "I", LEC_Body, False
                            resCheck.Result = "KO"
                        Else
                            WriteOK objWSheetBody, "I", LEC_Body, True
                        End If
                        
                    'CK25 - Check la présence des matières
                        Set resCheck = ResChecks.Item("CK25")
                        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                            resCheck.Result = "Check"
                        End If
                        objWSheetBody.Range("H" & LEC_Body) = FicEC.partmat
                        Condition_Vide objWSheetBody, "H" & LEC_Body, "H" & LEC_Body
                        
                    'CK27 - Check les formules cassée
                        Set resCheck = ResChecks.Item("CK27")
                        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
                            resCheck.Result = "KO"
                        End If
                        If FicEC.Broken_relations Then
                            WriteOK objWSheetBody, "J", LEC_Body, False
                            resCheck.Result = "KO"
                        Else
                            WriteOK objWSheetBody, "J", LEC_Body, True
                        End If
                        
                        LEC_Body = LEC_Body + 1
                            
                    End If ' Fin FicEC.EstPart
                End If 'Fin FRM_Check3D.ChB_PBody
                
    '##################################################################################
'        'Controle des attributs
                If FRM_Check3D.ChB_AttribGSE Then
                    objWSheetAttrib.Range("A" & LEC_Attrib) = CheckAttibsGSE(objWSheetAttrib, ResChecks, FicEC, LEC_Attrib, ColXlAttribs)
                End If 'Fin FRM_Check3D.ChB_AttribGSE
                
            End If 'Fin de EstPart(DocEC) Or EstProduct(DocEC)
            
        End If 'Exclusion du product général
        
     '##################################################################################
        'Controle des Contraintes
        If FRM_Check3D.ChB_Contraintes Then
            CheckContraintes objWSContr, ResChecks, FicEC, LEC_Contr
        End If
Erreur1:
    Err.Clear
        'le document n'est ni un Part ni un Product on passe au doc suivant
    Next i

    'Mise sous forme de tableau des onglets Attributs GSE et part Body
    SheetSel = "$A$3:$V$" & LEC_Attrib
    objWSheetAttrib.ListObjects.Add(xlSrcRange, objWSheetAttrib.Range(SheetSel), , xlYes).Name = "TablAttrib"
    objWSheetAttrib.ListObjects("TablAttrib").TableStyle = "TableStyleMedium6"
    
    SheetSel = "$A$3:$J$" & LEC_Body
    objWSheetBody.ListObjects.Add(xlSrcRange, objWSheetBody.Range(SheetSel), , xlYes).Name = "TablBody"
    objWSheetBody.ListObjects("TablBody").TableStyle = "TableStyleMedium6"
    
    SheetSel = "$A$3:$G$" & LEC_Body
    objWSFichiers.ListObjects.Add(xlSrcRange, objWSFichiers.Range(SheetSel), , xlYes).Name = "TablFichier"
    objWSFichiers.ListObjects("TablFichier").TableStyle = "TableStyleMedium6"

'##################################################################################
' Noms des instances et tri de l'arbre
Dim ValPrec As String
    If FRM_Check3D.ChB_Nommage Then
        'objWSheetArbo.Activate
        ReDim ListInst(4, 0)
        NivSSE = 1
        'Construit la liste des instances
        ListInstances ActiveDoc.Product.Products, NivSSE
        For i = 0 To UBound(ListInst, 2)
            If ListInst(2, i) = "KO" Then
                objWSheetArbo.Range("A" & LEC_Arbo) = ListInst(0, i)
                objWSheetArbo.Range("B" & LEC_Arbo) = ListInst(1, i)
                objWSheetArbo.Range("C" & LEC_Arbo) = ListInst(3, i)
                LEC_Arbo = LEC_Arbo + 1
            End If
        Next i
        
        LEC_Arbo = 4
   
        For i = 0 To UBound(ListInst, 2)
            objWSheetArbo.Range("F" & LEC_Arbo) = ListInst(0, i)
            If Not (ListInst(0, i) = "Env." Or ListInst(0, i) = "") Then
                'Calcul du décallage de colonne
                If ListInst(4, i) = "" Then
                    Col = "G"
                Else
                    Col = NumCar(ListInst(4, i) + 6)
                    colnum = ListInst(4, i) + 6
                End If
                objWSheetArbo.Range(Col & LEC_Arbo) = ListInst(0, i)
                'Teste si la valeur de la cellule en cours est supérieur à celle de la ligne précédente
                ValPrec = objWSheetArbo.Range(Col & LEC_Arbo - 1)
                If ValPrec <= objWSheetArbo.Range(Col & LEC_Arbo) Then
                    CouleurCell objWSheetArbo, "F", LEC_Arbo, "vert"
                    CouleurCell objWSheetArbo, Col, LEC_Arbo, "vert"
                Else
                    If ValPrec <> "ENV" Then
                        CouleurCell objWSheetArbo, "F", LEC_Arbo, "rouge"
                        CouleurCell objWSheetArbo, Col, LEC_Arbo, "rouge"
                    Else
                        CouleurCell objWSheetArbo, "F", LEC_Arbo, "vert"
                        CouleurCell objWSheetArbo, Col, LEC_Arbo, "vert"
                    End If
                    NB_Err_Tri = NB_Err_Tri + 1
                End If
            End If
            LEC_Arbo = LEC_Arbo + 1
        Next i
        objWSheetArbo.Range("G" & 4) = "Nb d'erreur de tri : " & NB_Err_Tri
        Set resCheck = ResChecks.Item("CK10")
        If NB_Err_Tri > 0 Then
            resCheck.Result = "KO"
        Else
            resCheck.Result = "Check"
        End If
     End If
    

    'Ecriture des resultats des checks dans l'onglet Général
    WritResultChecks objWSRecap, ResChecks
    
'Liberation des classes
Set FicEC = Nothing
Set ColXlAttribs = Nothing
Set mBarre = Nothing
Unload FRM_Check3D

MsgBox "Check 3D terminé", vbInformation, "Opération terminée"

End Sub

Private Function CheckAttibsGSE(objWSAttrib, ResChecks, FicEC, Lig, ColXl) As Long
'Controle des attributs
Dim resCheck As c_ResCheck3D
Dim c_col As c_ColXl
Dim totCheck As Long
Dim iTypGse As Integer
Dim nTypeGse As String

    Set resCheck = ResChecks.Item("CK30")
    If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le laisse a "KO"
        resCheck.Result = "Check"
    End If
    'recherche le type de l'élement
    iTypGse = TypeElement(FicEC.PN, FicEC.Val_NomPulsGSE_TypeNum)
    Select Case iTypGse
        Case 0
            nTypeGse = "Assemblage"
        Case 1, 2, 3, 4, 5, 6, 7, 8
            nTypeGse = "Pièce fab"
        Case 9
            nTypeGse = "Standard"
    End Select
    
    'Ecriture des attributs dans le fichier de rapport
    'objWSAttrib.Activate
        objWSAttrib.Range(ColXl.Item("check").Col & Lig) = totCheck
        objWSAttrib.Range(ColXl.Item("pn").Col & Lig) = FicEC.NomFic
        objWSAttrib.Range(ColXl.Item("type").Col & Lig) = nTypeGse
        objWSAttrib.Range(ColXl.Item("desout").Col & Lig) = FicEC.Val_NomPulsGSE_DesignOutillage
        objWSAttrib.Range(ColXl.Item("noout").Col & Lig) = FicEC.Val_NomPulsGSE_NoOutillage
        objWSAttrib.Range(ColXl.Item("site").Col & Lig) = FicEC.Val_NomPulsGSE_SiteAB
        objWSAttrib.Range(ColXl.Item("chk").Col & Lig) = FicEC.Val_NomPulsGSE_CHK
        objWSAttrib.Range(ColXl.Item("datepl").Col & Lig) = FicEC.Val_NomPulsGSE_DatePlan
        
        objWSAttrib.Range(ColXl.Item("ce").Col & Lig) = FicEC.Val_NomPulsGSE_CE
        objWSAttrib.Range(ColXl.Item("useguide").Col & Lig) = FicEC.Val_NomPulsGSE_PresUserGuide
        objWSAttrib.Range(ColXl.Item("prescais").Col & Lig) = FicEC.Val_NomPulsGSE_PresCaisse
        objWSAttrib.Range(ColXl.Item("nocais").Col & Lig) = FicEC.Val_NomPulsGSE_NoCaisse
        objWSAttrib.Range(ColXl.Item("descr").Col & Lig) = FicEC.Val_Description
        objWSAttrib.Range(ColXl.Item("sheet").Col & Lig) = FicEC.Val_NomPulsGSE_Sheet
        objWSAttrib.Range(ColXl.Item("itemnb").Col & Lig) = FicEC.Val_NomPulsGSE_ItemNb
        
        objWSAttrib.Range(ColXl.Item("dim").Col & Lig) = FicEC.Val_NomPulsGSE_Dimension
        objWSAttrib.Range(ColXl.Item("mat").Col & Lig) = FicEC.Val_NomPulsGSE_Material
        objWSAttrib.Range(ColXl.Item("protect").Col & Lig) = FicEC.Val_NomPulsGSE_Protect
        objWSAttrib.Range(ColXl.Item("misc").Col & Lig) = FicEC.Val_NomPulsGSE_Miscellanous
        
        objWSAttrib.Range(ColXl.Item("supref").Col & Lig) = FicEC.Val_NomPulsGSE_SupplierRef
        objWSAttrib.Range(ColXl.Item("weight").Col & Lig) = FicEC.Val_NomPulsGSE_Weight
        objWSAttrib.Range(ColXl.Item("mecano").Col & Lig) = FicEC.Val_NomPulsGSE_MecanoSoude

    'Mise en evidence des erreurs
    'Champs commun a tous les parts/products
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("pn").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("desout").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("noout").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("site").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("chk").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("datepl").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("sheet").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("itemnb").Col, Lig, "AV")
        
    'Champs unique sur product de tête
    If iTypGse = 0 Then
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("ce").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("nocais").Col, Lig, "AV")
        Mef objWSAttrib, ColXl.Item("useguide").Col & Lig & ":" & ColXl.Item("prescais").Col & Lig, "GM"
        Mef objWSAttrib, ColXl.Item("dim").Col & Lig & ":" & ColXl.Item("mecano").Col & Lig, "GM"
    
    'Le reste des parts/products
    Else
        Mef objWSAttrib, ColXl.Item("ce").Col & Lig & ":" & ColXl.Item("nocais").Col & Lig, "GM"
    End If
    
    'Champs unique sur product outillage et variante
    If iTypGse > 0 And iTypGse < 3 Then
        Mef objWSAttrib, ColXl.Item("dim").Col & Lig & ":" & ColXl.Item("supref").Col & Lig, "GM"
        Mef objWSAttrib, ColXl.Item("mecano").Col & Lig & ":" & ColXl.Item("mecano").Col & Lig, "GM"
    ElseIf iTypGse > 2 And iTypGse < 7 Then
        Mef objWSAttrib, ColXl.Item("dim").Col & Lig & ":" & ColXl.Item("mat").Col & Lig, "GM"
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("protect").Col, Lig, "AV")
        Mef objWSAttrib, ColXl.Item("supref").Col & Lig & ":" & ColXl.Item("supref").Col & Lig, "GM"
        Mef objWSAttrib, ColXl.Item("mecano").Col & Lig & ":" & ColXl.Item("mecano").Col & Lig, "GM"
        'Cas des symetriques
        If iTypGse = 4 Or iTypGse = 6 Then
            totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("misc").Col, Lig, "AV")
        Else
            Mef objWSAttrib, ColXl.Item("misc").Col & Lig & ":" & ColXl.Item("misc").Col & Lig, "GM"
        End If
    'Cas des parts fabriqués
    ElseIf iTypGse > 6 And iTypGse < 9 Then
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("dim").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("mat").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("protect").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("misc").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("mecano").Col, Lig, "A")
        Mef objWSAttrib, ColXl.Item("supref").Col & Lig & ":" & ColXl.Item("supref").Col & Lig, "GM"
        
    'Cas des parts achetés
    ElseIf iTypGse = 9 Then
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("dim").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("mat").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("protect").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("misc").Col, Lig, "AV")
        totCheck = totCheck + verifCheck(objWSAttrib, ColXl.Item("supref").Col, Lig, "AV")
        Mef objWSAttrib, ColXl.Item("mecano").Col & Lig & ":" & ColXl.Item("mecano").Col & Lig, "GM"
    End If

    Lig = Lig + 1
CheckAttibsGSE = totCheck
End Function

Private Sub CheckContraintes(objWSContr, ResChecks, FicEC, Lig As Long)
'Controle des contraintes
Dim ContStatut As Boolean, ContFix As Boolean
Dim resCheck As c_ResCheck3D
Dim contraintes As Constraints
Dim i As Long

    If FicEC.EstProduct Then
        ContStatut = True
        ContFix = False
        Set contraintes = FicEC.col_Consts
        objWSContr.Range("A" & Lig) = FicEC.NomFic
        For i = 1 To contraintes.Count
            If contraintes.Item(i).status <> catCstStatusOK Then
                ContStatut = False
            ElseIf contraintes.Item(i).ReferenceType = catCstRefTypeFixInSpace Then
                ContFix = True
            End If
        Next i
    'CK40    Vérification qu'une contrainte de fixité est présente
        Set resCheck = ResChecks.Item("CK40")
        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
            resCheck.Result = "Check"
        End If
        If ContFix Then
            WriteOK objWSContr, "B", Lig, True
        Else
            WriteOK objWSContr, "B", Lig, False
            resCheck.Result = "KO"
        End If
    'CK41    Vérification des contraintes 'brisées'
        Set resCheck = ResChecks.Item("CK41")
        If resCheck.Result <> "KO" Then 'déclare le check comme vérifié. S'il a echoué sur un précédent 3D on ne le touche pas
            resCheck.Result = "Check"
        End If
        If ContStatut Then
            WriteOK objWSContr, "C", Lig, True
        Else
            WriteOK objWSContr, "C", Lig, False
            resCheck.Result = "KO"
        End If

        Lig = Lig + 1
    End If


End Sub

Private Function InitWShtContr(objExcelCheck)
'Initialise l'onglet contraintes
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Contraintes"

 'Entète onglet Contraintes
    oWShTemp.Range("A" & ligneEC) = "Rapport des erreurs de contraintes"
    ligneEC = ligneEC + 2
    oWShTemp.Range("A" & ligneEC) = "Part Number"
    oWShTemp.Range("B" & ligneEC) = "Fixité"
    oWShTemp.Range("C" & ligneEC) = "Statut des contraintes"
    
    'Mise en forme
    Mef oWShTemp, "A" & 1 & ":" & "C" & 1, "BG"
    Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM"
    LargCol oWShTemp, "A", "I", 25
     'Fige les volets
    oWShTemp.Range("B4").Select
    objExcelCheck.ActiveWindow.FreezePanes = True
    
    Set InitWShtContr = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtFichiers(objExcelCheck)
'Initialise l'onglet Fichiers
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Fichiers"

'Entète onglet Fichiers
    oWShTemp.Range("A" & ligneEC) = "Rapport des erreurs de Nommage"

    ligneEC = ligneEC + 2
    oWShTemp.Range("A" & ligneEC) = "Numero du Fichier"
    oWShTemp.Range("B" & ligneEC) = "14 carractères"
    oWShTemp.Range("C" & ligneEC) = "Radical"
    oWShTemp.Range("D" & ligneEC) = "Nom Part," & Chr(10) & "et fichier identiques"
    oWShTemp.Range("E" & ligneEC) = "N° Impaire" & Chr(10) & "Symétrique ?"
    oWShTemp.Range("F" & ligneEC) = "Part Number"
    oWShTemp.Range("G" & ligneEC) = "Nom du fichier"
    
    'Mise en forme
    Mef oWShTemp, "A" & 1 & ":" & "G" & 1, "BI"
    Mef oWShTemp, "A" & ligneEC & ":" & "G" & ligneEC, "BM"
    LargCol oWShTemp, "A", "A", 25
    LargCol oWShTemp, "B", "G", 14
    'Fige les volets
    oWShTemp.Range("B4").Select
    objExcelCheck.ActiveWindow.FreezePanes = True

    Set InitWShtFichiers = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtArbo(objExcelCheck)
'Initialise l'onglet Arborescence
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Arborescence"

'Entète onglet Arbo
    oWShTemp.Range("A" & ligneEC) = "Rapport des erreurs dans l'arborescence"
    Mef oWShTemp, "A" & 1 & ":" & "G" & 1, "BI"

    ligneEC = ligneEC + 2
    oWShTemp.Range("A" & ligneEC) = "Controle des instances"
    Mef oWShTemp, "A" & ligneEC & ":" & "B" & ligneEC, "BM"
    ligneEC = ligneEC + 1
    oWShTemp.Range("A" & ligneEC) = "Part Number"
    oWShTemp.Range("B" & ligneEC) = "Nom d'instance"
    oWShTemp.Range("C" & ligneEC) = "Product parent"
    oWShTemp.Range("F" & ligneEC) = "Tri de l'arbre"
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC & ":" & "G" & ligneEC, "BM"
    LargCol oWShTemp, "A", "Z", 18
    'Fige les volets
    oWShTemp.Range("D5").Select
    objExcelCheck.ActiveWindow.FreezePanes = True
    
    Set InitWShtArbo = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtPartBody(objExcelCheck)
'Initialise l'onglet Part Body
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Part Body"

'Entète onglet Body
    oWShTemp.Range("A" & ligneEC) = "Rapport des erreurs Dans les Parts Body"
    Mef oWShTemp, "A" & 1 & ":" & "J" & 1, "BI"

    ligneEC = ligneEC + 2
    oWShTemp.Range("A" & ligneEC) = "Part Number"
    oWShTemp.Range("B" & ligneEC) = "Item Non résolu"
    oWShTemp.Range("C" & ligneEC) = "Détail"
    oWShTemp.Range("D" & ligneEC) = "Item Non mis à jour"
    oWShTemp.Range("E" & ligneEC) = "Détail"
    oWShTemp.Range("F" & ligneEC) = "Part Body unique"
    oWShTemp.Range("G" & ligneEC) = "Main Body actif"
    oWShTemp.Range("H" & ligneEC) = "Materiau sur PartBody"
    oWShTemp.Range("I" & ligneEC) = "Body non Hybride"
    oWShTemp.Range("J" & ligneEC) = "Formules cassées"
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC & ":" & "J" & ligneEC, "BM"
    LargCol oWShTemp, "A", "Z", 20
    'Fige les volets
    oWShTemp.Range("B4").Select
    objExcelCheck.ActiveWindow.FreezePanes = True
    
    Set InitWShtPartBody = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtAttribs(objExcelCheck, ColXlAttribs)
'Initialise l'onglet Attributs GSE
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
Dim c_col As c_ColXl

    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Attributs GSE"

'Entète onglet Attributs GSE
    oWShTemp.Range("A" & ligneEC) = "Rapport des erreurs sur les atributs GSE"
    ligneEC = ligneEC + 2
    ' Ajoute les entète et ajuste la largeur de colonnes
    For Each c_col In ColXlAttribs.Items
        oWShTemp.Range(c_col.Col & ligneEC) = c_col.Nom
        LargCol oWShTemp, c_col.Col, c_col.Col, c_col.lCol
    Next
    Set c_col = ColXlAttribs.Item("mecano")
    'Mise en forme
    Mef oWShTemp, "A" & 1 & ":" & c_col.Col & 1, "BG"
    Mef oWShTemp, "A" & ligneEC & ":" & c_col.Col & ligneEC, "BM"
    
    oWShTemp.Range("D4").Select
    objExcelCheck.ActiveWindow.FreezePanes = True
    
    Set InitWShtAttribs = oWShTemp
    Set oWShTemp = Nothing
    Set c_col = Nothing
End Function

Private Function InitWShtRecap(objExcelCheck, ByRef ResChecks)
'Initialise l'onglet Général
Dim oWShTemp
Dim resCheck As c_ResCheck3D
Dim i As Long
Dim ligneEC As Long
    ligneEC = 1
    
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Général"

'Entète onglet recap
    oWShTemp.Range("A" & 1) = "Rapport de controle 3D GSE"
    ligneEC = ligneEC + 1

    oWShTemp.Range("B" & ligneEC) = "Outil analysé :"
    Mef oWShTemp, "A" & ligneEC & ":" & "B" & ligneEC, "BM" 'Mise en forme
    ligneEC = ligneEC + 1
    
    oWShTemp.Range("B" & ligneEC) = "Date :"
    oWShTemp.Range("C" & ligneEC) = Date
    Mef oWShTemp, "A" & ligneEC & ":" & "B" & ligneEC, "BM" 'Mise en forme
    ligneEC = ligneEC + 1

    oWShTemp.Range("B" & ligneEC) = "Heure :"
    oWShTemp.Range("C" & ligneEC) = Time
    Mef oWShTemp, "A" & ligneEC & ":" & "B" & ligneEC, "BM" 'Mise en forme
    ligneEC = ligneEC + 2
    
    oWShTemp.Range("A" & ligneEC) = "Liste des checks effectués"
    Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM" 'Mise en forme
    
    'Mise en forme
    Mef oWShTemp, "A" & 1 & ":" & "C" & 1, "BG"
    Mef oWShTemp, "A" & ligneEC, "BM" 'Mise en forme
    LargCol oWShTemp, "A", "A", 6
    LargCol oWShTemp, "B", "B", 66
    LargCol oWShTemp, "C", "C", 25
    
    'Ecriture de la liste des checks effectués
        ligneEC = ligneEC + 1
        oWShTemp.Range("B" & ligneEC) = "Onglet Fichiers"
        Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM"
        ligneEC = ligneEC + 1
        For i = 1 To 4
            Set resCheck = ResChecks.Item(i)
            oWShTemp.Range("A" & ligneEC) = resCheck.Id
            oWShTemp.Range("B" & ligneEC) = resCheck.Libel
            oWShTemp.Range("C" & ligneEC) = resCheck.Result
            resCheck.Lig = ligneEC
            ligneEC = ligneEC + 1
        Next
        
        ligneEC = ligneEC + 1
        oWShTemp.Range("B" & ligneEC) = "Onglet Arborescence"
        Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM"
        ligneEC = ligneEC + 1
        For i = 5 To 5
            Set resCheck = ResChecks.Item(i)
            oWShTemp.Range("A" & ligneEC) = resCheck.Id
            oWShTemp.Range("B" & ligneEC) = resCheck.Libel
            oWShTemp.Range("C" & ligneEC) = resCheck.Result
            resCheck.Lig = ligneEC
            ligneEC = ligneEC + 1
        Next
    
        ligneEC = ligneEC + 1
        oWShTemp.Range("B" & ligneEC) = "Onglet Part Body"
        Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM"
        ligneEC = ligneEC + 1
        For i = 6 To 12
            Set resCheck = ResChecks.Item(i)
            oWShTemp.Range("A" & ligneEC) = resCheck.Id
            oWShTemp.Range("B" & ligneEC) = resCheck.Libel
            oWShTemp.Range("C" & ligneEC) = resCheck.Result
            resCheck.Lig = ligneEC
            ligneEC = ligneEC + 1
        Next
        ligneEC = ligneEC + 1
        oWShTemp.Range("B" & ligneEC) = "Onglet Attributs GSE"
        Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM"
        ligneEC = ligneEC + 1
        For i = 13 To 13
            Set resCheck = ResChecks.Item(i)
            oWShTemp.Range("A" & ligneEC) = resCheck.Id
            oWShTemp.Range("B" & ligneEC) = resCheck.Libel
            oWShTemp.Range("C" & ligneEC) = resCheck.Result
            resCheck.Lig = ligneEC
            ligneEC = ligneEC + 1
        Next

        ligneEC = ligneEC + 1
        oWShTemp.Range("B" & ligneEC) = "Onglet contraintes"
        Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM"
        ligneEC = ligneEC + 1
        For i = 14 To 15
            Set resCheck = ResChecks.Item(i)
            oWShTemp.Range("A" & ligneEC) = resCheck.Id
            oWShTemp.Range("B" & ligneEC) = resCheck.Libel
            oWShTemp.Range("C" & ligneEC) = resCheck.Result
            resCheck.Lig = ligneEC
            ligneEC = ligneEC + 1
        Next

    
    Set InitWShtRecap = oWShTemp
    Set oWShTemp = Nothing
    
End Function

Private Function InitResCheck() As c_ResCheck3Ds
'Initialise la classe des resultats des Checks

Dim tResChecks As c_ResCheck3Ds

    Set tResChecks = New c_ResCheck3Ds

    tResChecks.Add "CK01", "Vérification de la longueur des noms de fichiers", "No Check", "8"
    tResChecks.Add "CK02", "Vérification du radical du nom de fichier", "No Check", "9"
    tResChecks.Add "CK03", "Vérification que le PN, le nom d'instance et le file name sont identiques", "No Check", "10"
    tResChecks.Add "CK04", "Vérification des Numéros Impaire", "No Check", "11"
    tResChecks.Add "CK10", "Vérification des Noms d'instance et du tri de l'arbre", "No Check", "14"
    tResChecks.Add "CK20", "Vérification qu'il n'y a qu'un seul Part Body pour les pièce fabriquées (200 à 699)", "No Check", "18"
    tResChecks.Add "CK21", "Vérification des éléments non résolu", "No Check", "19"
    tResChecks.Add "CK22", "Vérification des éléments non mis à jour", "No Check", "20"
    'tResChecks.Add "CK23", "Vérification Qu'il n'y a qu'un seul Part Body", "No Check", "21"
    tResChecks.Add "CK24", "Vérification Que le Main Body est actif", "No Check", "22"
    tResChecks.Add "CK25", "Vérification de la présence d'un matériaux sur le Main body", "No Check", "23"
    tResChecks.Add "CK26", "Vérification que le part n'est pas en conception hybride", "No Check", "24"
    tResChecks.Add "CK27", "Vérification que les formules ne soient pas cassées", "No Check", "25"
    tResChecks.Add "CK30", "Vérification de la présence des attributs GSE", "No Check", "29"
    tResChecks.Add "CK40", "Vérification qu'une contrainte de fixité est présente", "No Check", "32"
    tResChecks.Add "CK41", "Vérification des contraintes 'brisées'", "No Check", "33"

    Set InitResCheck = tResChecks

End Function


Private Function DerCelCol(WS_Tri, WS_Col, WS_lig) As String
'remonte dans la colonne du Worksheet de la ligne pasée en argument vers la ligne 0
'Jussqu'a ce qu'on trouve une cellule non vide.
'renvois le contnu de cette cellule sinon renvoi une chaine vide
    Dim i As Long
    DerCelCol = ""
    For i = WS_lig To 1 Step -1
        If WS_Tri.Range(WS_Col & i) <> "" Then DerCelCol = WS_Tri.Range(WS_Col & i)
    Next i
End Function

Private Function InitColAttrib() As c_ColXls
'Initialise la classe des collones du fichier Excel
Dim c_col As c_ColXl
Dim c_cols As c_ColXls
Dim Noms(1 To 22, 1 To 4) As String
Dim i As Long

    'Nom = id, colonne excel, Nom entète colonne, largeur de colonne
    Noms(1, 1) = "check": Noms(1, 2) = "A": Noms(1, 3) = "Erreurs": Noms(1, 4) = "6"
    Noms(2, 1) = "type": Noms(2, 2) = "B": Noms(2, 3) = "Type GSE": Noms(2, 4) = "8"
    Noms(3, 1) = "pn": Noms(3, 2) = "C": Noms(3, 3) = "Part Number": Noms(3, 4) = "25"
    Noms(4, 1) = "desout": Noms(4, 2) = "D": Noms(4, 3) = "DesignOutilage": Noms(4, 4) = "38"
    Noms(5, 1) = "noout": Noms(5, 2) = "E": Noms(5, 3) = "NoOutilage": Noms(5, 4) = "15"
    Noms(6, 1) = "site": Noms(6, 2) = "F": Noms(6, 3) = "SiteAB": Noms(6, 4) = "15"
    Noms(7, 1) = "chk": Noms(7, 2) = "G": Noms(7, 3) = "CHK": Noms(7, 4) = "16"
    Noms(8, 1) = "datepl": Noms(8, 2) = "H": Noms(8, 3) = "DatePlan": Noms(8, 4) = "12"
    Noms(9, 1) = "ce": Noms(9, 2) = "I": Noms(9, 3) = "CE": Noms(9, 4) = "12"
    Noms(10, 1) = "useguide": Noms(10, 2) = "J": Noms(10, 3) = "Presence User Guide": Noms(10, 4) = "12"
    Noms(11, 1) = "prescais": Noms(11, 2) = "K": Noms(11, 3) = "Presence caisse": Noms(11, 4) = "12"
    Noms(12, 1) = "nocais": Noms(12, 2) = "L": Noms(12, 3) = "N° de caisse": Noms(12, 4) = "12"
    Noms(13, 1) = "descr": Noms(13, 2) = "M": Noms(13, 3) = "Description": Noms(13, 4) = "38"
    Noms(14, 1) = "sheet": Noms(14, 2) = "N": Noms(14, 3) = "Sheet": Noms(14, 4) = "12"
    Noms(15, 1) = "itemnb": Noms(15, 2) = "O": Noms(15, 3) = "ItemNB": Noms(15, 4) = "12"
    Noms(16, 1) = "dim": Noms(16, 2) = "P": Noms(16, 3) = "Dimension": Noms(16, 4) = "12"
    Noms(17, 1) = "mat": Noms(17, 2) = "Q": Noms(17, 3) = "Material": Noms(17, 4) = "18"
    Noms(18, 1) = "protect": Noms(18, 2) = "R": Noms(18, 3) = "Protect": Noms(18, 4) = "18"
    Noms(19, 1) = "misc": Noms(19, 2) = "S": Noms(19, 3) = "Micsellanous": Noms(19, 4) = "18"
    Noms(20, 1) = "supref": Noms(20, 2) = "T": Noms(20, 3) = "SupplierRef": Noms(20, 4) = "18"
    Noms(21, 1) = "weight": Noms(21, 2) = "U": Noms(21, 3) = "Weight": Noms(21, 4) = "18"
    Noms(22, 1) = "mecano": Noms(22, 2) = "V": Noms(22, 3) = "MecaoSoude": Noms(22, 4) = "12"

    Set c_col = New c_ColXl
    Set c_cols = New c_ColXls
    For i = 1 To UBound(Noms, 1)
        c_col.Id = Noms(i, 1)
        c_col.Col = Noms(i, 2)
        c_col.Nom = Noms(i, 3)
        c_col.lCol = CDbl(Noms(i, 4))
        c_cols.Add c_col.Id, c_col.Col, c_col.Nom, c_col.lCol
    Next i

    Set InitColAttrib = c_cols
    Set c_cols = Nothing

End Function

Public Sub ListInstances(LI_Products As Products, Niveau As Integer)
' *****************************************************************
' * Construit une liste des product avec leur nom d'instance
' * Procedure récursive
' * Création CFR le 09/02/2016
' * Dernière modification le :
' *****************************************************************
On Error GoTo err_ListInstances
Dim LI_Part As Part
Dim LI_Product As Product
Dim i As Long, LInst_EC As Long
Dim nPatNumber As String
'If Niveau > 3 Then
'    Niveau = Niveau - 1
'Else
'    Niveau = Niveau + 1
'End If
For i = 1 To LI_Products.Count
    nPatNumber = ""
    On Error Resume Next
    Err.Clear
    Set LI_Product = LI_Products.Item(i)
    If Not (Err.Number <> 0) Then
    
        On Error GoTo 0
        If TypeName(LI_Product) = "Product" Then
            'test si le product à un PartNumber (les groupes n'en ont pas)
            On Error Resume Next
            nPatNumber = LI_Product.PartNumber
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
            Else
                On Error GoTo 0
                LInst_EC = UBound(ListInst, 2) + 1
                ReDim Preserve ListInst(4, LInst_EC)
                ListInst(0, LInst_EC) = LI_Product.PartNumber
                ListInst(1, LInst_EC) = LI_Products.Item(i).Name
                ListInst(2, LInst_EC) = Ck_NomInstance(LI_Product.PartNumber, LI_Products.Item(i).Name)
                ListInst(3, LInst_EC) = LI_Products.Item(i).Parent.Parent.Name
                ListInst(4, LInst_EC) = Niveau
                'Appel recursif si le product contient d'autres parts ou products
                If LI_Product.Products.Count > 0 Then
                    Niveau = Niveau + 1
                    ListInstances LI_Product.Products, Niveau
                    Niveau = Niveau + 1
                End If
            End If
        End If
    End If
    On Error Resume Next
    Err.Clear
    Set LI_Part = LI_Products.Item(i)
    If Not (Err.Number <> 0) Then
        On Error GoTo 0
        If TypeName(LI_Part) = "Part" Then
            LInst_EC = UBound(ListInst, 2) + 1
            ReDim Preserve ListInst(4, LInst_EC)
            ListInst(0, LInst_EC) = LI_Part.PartNumber
            ListInst(1, LInst_EC) = LI_Products.Item(i).Name
            ListInst(2, LInst_EC) = Ck_NomInstance(LI_Product.PartNumber, LI_Products.Item(i).Name)
            ListInst(3, LInst_EC) = LI_Products.Item(i).Parent.Parent
            ListInst(4, LInst_EC) = Niveau
         End If
    End If
Next i

GoTo quit_err_ListInstances

err_ListInstances:
MsgBox Err.Number & Err.Description

quit_err_ListInstances:
End Sub

Private Function Ck_NomInstance(pNumber As String, iName As String) As String
'test si le nom d'instance est egal au non de partnumber
'Elimine l'extension ".x" ajouté a la fin du nom d'instance
Dim NomInst As String
Dim Pnum As String

    If InStr(1, iName, ".", vbTextCompare) > 0 Then
        NomInst = Left(iName, InStr(1, iName, ".", vbTextCompare) - 1)
    Else
        NomInst = iName
    End If
    If InStr(1, pNumber, ".", vbTextCompare) > 0 Then
        Pnum = Left(pNumber, InStr(1, pNumber, ".", vbTextCompare) - 1)
    Else
        Pnum = pNumber
    End If
    
    If Pnum = NomInst Then Ck_NomInstance = "OK" Else Ck_NomInstance = "KO"
End Function

Private Function verifCheck(Wsheet, Col, Lig, AV) As Integer
'change la couleur de la cellule  et renvois 1 ou zéro en fonction de son contenu
'Wsheet onglet Excel
'Col, Lig  colonne et ligne de la cellule a vérifier
'AV = "A" ou "V" ou "AV" pour changer la couleur des cellule "vide" ou "Absent" ou les deux
Dim Cell As String
    Cell = Col & Lig

'Format conditionel des cellues.
'Passe en rouge les cellules contenant "Absent" et en orange celle contenant "Vide"

    Dim Formule As String
    Formule = "=""Vide"""
    With Wsheet.Range(Cell)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xLTextString, String:="Absent", TextOperator:=xlContains
        .FormatConditions(1).SetFirstPriority
        .FormatConditions(1).Interior.Color = 255
        .FormatConditions(1).StopIfTrue = False
        If InStr(1, AV, "V", vbTextCompare) > 0 Then
            .FormatConditions.Add Type:=xLTextString, String:="Vide", TextOperator:=xlContains
            .FormatConditions(2).Interior.Color = 49407
        End If

    End With
    
    If Wsheet.Range(Cell) = "Absent" Or Wsheet.Range(Cell) = "Vide" Then
        verifCheck = 1
    Else
        verifCheck = 0
    End If


End Function

Private Sub WriteVerif(Wsheet, Col As String, Lig As Long, Result As Boolean)
'Ecris le résultat "A vérifier" ou "KO" dans la cellule excel et change la couleur de la cellule (rouge ou vert)
'Wsheet = WorkSheet
'col = colonne de la cellule de destination
'Lig = Ligne de la cellule
'Resul = true pour "OK" et false pour "KO"
    
    If Result Then
        Wsheet.Range(Col & CStr(Lig)) = "OK"
        CouleurCell Wsheet, Col, CStr(Lig), "vert"
    Else
        Wsheet.Range(Col & CStr(Lig)) = "A verifier"
        CouleurCell Wsheet, Col, CStr(Lig), "jaune"
    End If


End Sub

Private Sub WriteOK(Wsheet, Col As String, Lig As Long, Result As Boolean)
'Ecris le résultat "OK" ou "KO" dans la cellule excel et change la couleur de la cellule (rouge ou vert)
'Wsheet = WorkSheet
'col = colonne de la cellule de destination
'Lig = Ligne de la cellule
'Resul = true pour "OK" et false pour "KO"
    
    If Result Then
        Wsheet.Range(Col & CStr(Lig)) = "OK"
        CouleurCell Wsheet, Col, CStr(Lig), "vert"
    Else
        Wsheet.Range(Col & CStr(Lig)) = "KO"
        CouleurCell Wsheet, Col, CStr(Lig), "rouge"
    End If


End Sub

Private Sub WritResultChecks(Wsheet, ResChecks)
'Reporte le résultat des checks dans l'onglet "Général
Dim resCheck As c_ResCheck3D

    For Each resCheck In ResChecks.Items
        Wsheet.Range("C" & resCheck.Lig) = resCheck.Result
        If resCheck.Result = "KO" Then
            CouleurCell Wsheet, "C", resCheck.Lig, "rouge"
        End If
    Next

End Sub
