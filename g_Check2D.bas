Attribute VB_Name = "G_Check2D"
Option Explicit
 Public Temp_NB_Planche As String
 

Sub catmain()
' *****************************************************************
' * Execution d'un check de certains critères sur une liasse de plans
' *
' *
' * Création CFR le 02/10/2015
' * Version 3.0
' * Dernière modification le :
' *
' *****************************************************************
'On Error Resume Next
'Chargement des variables

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "g_Check2D", VMacro

Dim checkEC As Check2D
Dim MacroLocation As New xMacroLocation
Dim NomRapportCheck As String
Dim LigneECRecap As Long, LigneECCotes As Long, LigneECParam As Long, LigneECNommage As Long, LigneECNom As Long, LigneECVues As Long
Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long
Dim NB_Ligne As Long
Dim DerPl As Long
Dim XlsFormule As String
Dim nDrawEc As String
Dim Rapport As String
Dim VoyantErr As String
Dim Check_Cartouche As Boolean, Check_Nomenclature As Boolean, Check_Vues As Boolean, Check_Cotes As Boolean
Dim Liens3DOK As Boolean
Dim Planche_KO As Boolean
Dim Temp_PartNb As String, Temp_NoPl As String, Temp_Item_Nb As String
Dim Liste_No_Planche() As String
Dim TypeNum As String ' pour conserver cette valeur apres la fermeture des instances
    TypeNum = "0"
Dim InfoCartRecurentes(8) As String
Dim objExcelCheck
Dim oWShtRecap, oWShtNommage, oWShtParam, oWShtCotes, oWShtVues, oWShtNomencl
Dim SheetSel As String
                    
'Les classes
Dim mBarre As c_ProgressBarre

    If Not (MacroLocation.FicIniExist("VarNomenclatureGSE.ini")) Then ' on est pas dans l'environnement GSE_R21
        MsgBox "Vous n'etes pas dans l'environnement GSE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
        Exit Sub
    Else
        MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
        CheminDestNomenclature = MacroLocation.ValVar("CheminDestNomenclature")
        CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
    End If
            
    Load Frm_Check2D
    Frm_Check2D.Show
    If Not (Frm_Check2D.ChB_OkAnnule) Then
        End
    End If
       
    CheminDestRapport = Frm_Check2D.TB_Chemin
    NomRapportCheck = "Check_" & Frm_Check2D.TB_Chemin & Date & ".xlsx"
    NomRapportCheck = Replace(NomRapportCheck, "/", "-", , , vbTextCompare)

 'verifie si un fichier de rapport est déja présent et l'efface
    If Not (EffaceFicNom(CheminDestRapport, NomRapportCheck)) Then
        End
    End If
    
'Création de la trame Excel
    Set objExcelCheck = CreateObject("EXCEL.APPLICATION")
    objExcelCheck.Visible = True
    objExcelCheck.Workbooks.Add
    
    Set oWShtNomencl = InitWShtNomenclature(objExcelCheck)
    LigneECNom = 3
    Set oWShtParam = InitWShtCartouche(objExcelCheck)
    LigneECParam = 3
    Set oWShtCotes = InitWShtCotes(objExcelCheck)
    LigneECCotes = 4
    Set oWShtVues = InitWShtvues(objExcelCheck)
    LigneECVues = 3
    Set oWShtNommage = InitWShtNommage(objExcelCheck)
    LigneECNommage = 4
    Set oWShtRecap = InitWShtRecap(objExcelCheck)
    LigneECRecap = 7
    oWShtRecap.Activate

'Initialisation de la barre de progression
    Set mBarre = New c_ProgressBarre
    mBarre.ProgressTitre 1, "Check 2D"
    mBarre.Affiche

ReDim Liste_No_Planche(1, Frm_Check2D.LB_Liste_Fichiers_Traites.ListCount - 1)
    'Pour tous les fichiers de la liste
    For i = 0 To Frm_Check2D.LB_Liste_Fichiers_Traites.ListCount - 1
        mBarre.Progression = ((100 / Frm_Check2D.LB_Liste_Fichiers_Traites.ListCount - 1) * i)
        
        nDrawEc = Frm_Check2D.LB_Liste_Fichiers_Traites.List(i, 0)
        Check_Vues = True
        Check_Cotes = True
        Check_Cartouche = True
        Check_Nomenclature = True
        VoyantErr = "OK"
        Liens3DOK = True
        
        'Instanciation de la classe Check2D
        Set checkEC = New Check2D
    
        CheminPlusNomFichier = Frm_Check2D.TB_Chemin & "\" & nDrawEc
        On Error Resume Next
        checkEC.CK_OuvreDraw = CheminPlusNomFichier
        
        'Traitement des erreures de la classe check_2D
        Select Case Err.Number
            Case (vbObjectError + 514)
                oWShtRecap.Range("A" & LigneECRecap) = nDrawEc
                Check_Vues = False
                oWShtVues.Range("B" & LigneECVues) = Err.Description
                Liens3DOK = False
                Err.Clear
                On Error GoTo 0
            Case (vbObjectError + 515)
                oWShtRecap.Range("A" & LigneECRecap) = nDrawEc
                Check_Vues = False
                oWShtVues.Range("B" & LigneECVues) = Err.Description
                Liens3DOK = False
                Err.Clear
                On Error GoTo 0
        End Select
        On Error GoTo 0
        If TypeNum = "0" Then
            TypeNum = checkEC.CK_NomPulsGSE_TypeNum
            'stockage des info du premier plan (planche 1) pour vérifier que toutes les autres planches sont identiques
            InfoCartRecurentes(0) = checkEC.CK_Cart_NumeroPlan  'Numero du plan
            InfoCartRecurentes(1) = checkEC.CK_Cart_Division
            InfoCartRecurentes(2) = checkEC.CK_Cart_DRN
            InfoCartRecurentes(3) = checkEC.CK_Cart_CHK
            InfoCartRecurentes(4) = checkEC.CK_Cart_DOOrig
            InfoCartRecurentes(5) = checkEC.CK_Cart_DWGSys
            InfoCartRecurentes(6) = checkEC.CK_Cart_Process
            InfoCartRecurentes(7) = checkEC.CK_Mod_Date
            InfoCartRecurentes(8) = checkEC.Val_Cart_Title
        End If
        
        'Collecte des N° de planches et des PartNumber
        Liste_No_Planche(0, i) = checkEC.CK_Lien3DName
        Liste_No_Planche(1, i) = checkEC.CK_Cart_Sheet
        
        oWShtRecap.Range("A" & LigneECRecap) = nDrawEc
    

'#####################
' Nommage des fichiers
'#####################
        If Frm_Check2D.ChB_Nommage Then
            oWShtRecap.Range("B" & LigneECRecap) = vNumerotation(oWShtNommage, checkEC, LigneECNommage, nDrawEc)
            ColorCell oWShtRecap, "B", LigneECRecap
        End If
        
'####
'Vues
'####
        If Frm_Check2D.ChB_Vues Then
            oWShtRecap.Range("C" & LigneECRecap) = vVues(oWShtVues, checkEC, LigneECVues, nDrawEc)
            ColorCell oWShtRecap, "C", LigneECRecap
        End If
        
'#####
'Cotes
'#####
        If Frm_Check2D.ChB_Bullage Then
            oWShtRecap.Range("D" & LigneECRecap) = vCotes(oWShtCotes, checkEC, LigneECCotes, nDrawEc)
            ColorCell oWShtRecap, "D", LigneECRecap
        End If
                  
'##########################
'Cartouche et paramètres 3D
'##########################
        
        If Frm_Check2D.ChB_Cartouche Then
            DerPl = Frm_Check2D.LB_Liste_Fichiers_Traites.ListCount - 1

            oWShtRecap.Range("E" & LigneECRecap) = vCartouche(oWShtParam, checkEC, LigneECParam, nDrawEc, DerPl, InfoCartRecurentes)
            ColorCell oWShtRecap, "E", LigneECRecap
        End If
    
'############
'Nomenclature
'############
        
        If Frm_Check2D.ChB_Nomenclature Then
            On Error Resume Next
            If checkEC.Table_Nom_Ens Then
                If Err.Number <> 0 Then
                    Err.Clear
                    oWShtRecap.Range("G" & LigneECRecap) = "Pas de tableau de nomenlature"
                    oWShtRecap.Range("F" & LigneECRecap) = "KO"
                Else
                    'Récapitulatif
                    oWShtRecap.Range("F" & LigneECRecap) = vNomenclature(oWShtNomencl, checkEC, LigneECNom, nDrawEc)
                End If
            ElseIf checkEC.Table_Nom_Det Then
                oWShtRecap.Range("F" & LigneECRecap) = "Plan de détail"
            End If
            ColorCell oWShtRecap, "F", LigneECRecap
        End If
        LigneECRecap = LigneECRecap + 1
    Next i

    'Mise sous forme de tableau de l'onglets Recap
    SheetSel = "$A$6:$G$" & LigneECRecap
    oWShtRecap.ListObjects.Add(xlSrcRange, oWShtRecap.Range(SheetSel), , xlYes).Name = "TablGeneral"
    oWShtRecap.ListObjects("TablAttrib").TableStyle = "TableStyleMedium6"
    
'Verification que les Numéros de planche sont correct
'colle la table des N° de planche dans le fichiers eXcel
    For m = 0 To UBound(Liste_No_Planche, 2)
        oWShtNomencl.Range("Q" & m + 2) = Liste_No_Planche(0, m)
        oWShtNomencl.Range("R" & m + 2) = Liste_No_Planche(1, m)
    Next m
    For m = 1 To LigneECNom
        Planche_KO = True
        Temp_Item_Nb = oWShtNomencl.Range("G" & m)
        Temp_PartNb = oWShtNomencl.Range("H" & m)
        Temp_NoPl = oWShtNomencl.Range("F" & m)
        If TypeElementRep(Temp_Item_Nb, TypeNum) < 9 And TypeElementRep(Temp_Item_Nb, TypeNum) > -1 Then
            For n = 1 To UBound(Liste_No_Planche, 2) + 2
                If oWShtNomencl.Range("Q" & n) = Temp_PartNb And oWShtNomencl.Range("R" & n) = Temp_NoPl Then
                    CouleurCell oWShtNomencl, 6, m, "vert"
                    Planche_KO = False
                End If
            Next n
            If Planche_KO Then
                CouleurCell oWShtNomencl, 6, m, "rouge"
                Check_Nomenclature = False
            End If
        End If
    Next m
    
    'libération des classes
    Set mBarre = Nothing
    Set checkEC = Nothing
    
    Unload Frm_Check2D
    MsgBox "Check 2D terminé", vbInformation, "Traitement terminé"

End Sub

Private Function InitWShtRecap(objExcelCheck)
'Initialise l'onglet récapitulatif
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Général"
    
'Entète onglet recap
    oWShTemp.Range("A" & ligneEC) = "Rapport de controle des 2D GSE"
    'Mise en forme
    Mef oWShTemp, "A" & 1 & ":" & "G" & 1, "BG"

    ligneEC = ligneEC + 1
    oWShTemp.Range("A" & ligneEC) = "Répertoire analysé :"
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC, "BM" 'Mise en forme

    oWShTemp.Range("B" & ligneEC) = Frm_Check2D.TB_Chemin
    ligneEC = ligneEC + 1
    oWShTemp.Range("A" & ligneEC) = "Date :"
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC, "BM" 'Mise en forme

    oWShTemp.Range("B" & ligneEC) = Date
    ligneEC = ligneEC + 1
    oWShTemp.Range("A" & ligneEC) = "Heure :"
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC, "BM" 'Mise en forme
    
    oWShTemp.Range("B" & ligneEC) = Time
    ligneEC = ligneEC + 2
    oWShTemp.Range("A" & ligneEC) = "Fichier traité"
    oWShTemp.Range("B" & ligneEC) = "Nommage des fichiers"
    oWShTemp.Range("C" & ligneEC) = "Vues"
    oWShTemp.Range("D" & ligneEC) = "Cotes et Bullage"
    oWShTemp.Range("E" & ligneEC) = "Cartouche"
    oWShTemp.Range("F" & ligneEC) = "Nomenclature"
    oWShTemp.Range("G" & ligneEC) = "Commentaires"
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC & ":" & "G" & ligneEC, "BM" 'Mise en forme
    
    LargCol oWShTemp, "A", "A", 28
    LargCol oWShTemp, "B", "F", 16
    LargCol oWShTemp, "G", "G", 30
    
    Set InitWShtRecap = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtvues(objExcelCheck)
'Initialise l'onglet Vues
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Vues"
    
    'Entète onglet Vues
    oWShTemp.Range("A" & ligneEC) = "Rapport des erreurs de Vues"
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC & ":" & "A" & ligneEC, "BI" 'Mise en forme
    
    LargCol oWShTemp, "A", "A", 32
    
    Set InitWShtvues = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtNommage(objExcelCheck)
'Initialise l'onglet Nommage
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Nommage des fichiers"
    
    'Entète Onglet Nommage
    oWShTemp.Range("A" & ligneEC) = "Rapport des erreurs de Nommage"
    Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BI" 'Mise en forme

    ligneEC = ligneEC + 2
    oWShTemp.Range("A" & ligneEC) = "Numero du drawing"
    oWShTemp.Range("B" & ligneEC) = "Numero de planche dans cartouche"
    oWShTemp.Range("C" & ligneEC) = "Numero de plan dans cartouche"
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM" 'Mise en forme
   
    LargCol oWShTemp, "A", "C", 32
   
    Set InitWShtNommage = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtCotes(objExcelCheck)
'Initialise l'onglet Cotes
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Cotes et Bullages"
    
    'Entète onglet Cotes
    oWShTemp.Range("A" & ligneEC) = "Rapport des erreurs de Cotes"
    ligneEC = ligneEC + 2
    oWShTemp.Range("A" & ligneEC) = "Numero du drawing"
    oWShTemp.Range("B" & ligneEC) = "Erreurs sur bullage"
    oWShTemp.Range("C" & ligneEC) = "Erreurs sur cotes"
    
    'Mise en forme
    Mef oWShTemp, "A" & 1 & ":" & "C" & ligneEC, "BI" 'Mise en forme
    Mef oWShTemp, "A" & ligneEC & ":" & "C" & ligneEC, "BM" 'Mise en forme

    LargCol oWShTemp, "A", "A", 32
    LargCol oWShTemp, "B", "B", 70
    LargCol oWShTemp, "C", "C", 50

    Set InitWShtCotes = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtCartouche(objExcelCheck)
'Initialise l'onglet Cartouche
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    objExcelCheck.worksheets.Add
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Cartouche"
    
    'Entète onglet Cartouche
    oWShTemp.Range("A" & ligneEC) = "Item checkés"
    oWShTemp.Range("B" & ligneEC) = "Paramètres 3D"
    oWShTemp.Range("C" & ligneEC) = "Valeur dans cartouche"
    oWShTemp.Range("D" & ligneEC) = "Valeur dans l'indice de modif"
    oWShTemp.Range("E" & ligneEC) = "Commentaires"
    
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC & ":" & "E" & ligneEC, "BM" 'Mise en forme
    LargCol oWShTemp, "A", "A", 28
    LargCol oWShTemp, "B", "D", 32
    LargCol oWShTemp, "E", "E", 40
    
    Set InitWShtCartouche = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function InitWShtNomenclature(objExcelCheck)
'Initialise l'onglet Nomenclature
Dim oWShTemp
Dim ligneEC As Long
    ligneEC = 1
    Set oWShTemp = objExcelCheck.worksheets.Item(1)
    oWShTemp.Name = "Nomenclature"
    
    'Entète onglet Nomenclature
    oWShTemp.Range("A" & ligneEC) = "Détail des nomenclatures"
    oWShTemp.Range("Q" & ligneEC) = "N° de détail"
    oWShTemp.Range("R" & ligneEC) = "N° de planche"
    
    'Mise en forme
    Mef oWShTemp, "A" & ligneEC & ":" & "O" & ligneEC, "BG" 'Mise en forme
    Mef oWShTemp, "Q" & ligneEC & ":" & "R" & ligneEC, "BM" 'Mise en forme
    LargCol oWShTemp, "A", "G", 6
    LargCol oWShTemp, "H", "I", 30
    LargCol oWShTemp, "J", "O", 16
    LargCol oWShTemp, "Q", "Q", 20
    LargCol oWShTemp, "R", "R", 25
    
    Set InitWShtNomenclature = oWShTemp
    Set oWShTemp = Nothing
End Function

Private Function vNumerotation(Wsheet, checkEC, LigneECNommage, nDrawEc) As String
'Vérification de la numérotation
Dim tResult As String
    On Error GoTo Erreur
        tResult = checkEC.CK_Numerotation
        If tResult = "KO" Then
            Wsheet.Range("A" & LigneECNommage) = nDrawEc
            Wsheet.Range("B" & LigneECNommage) = "'" & checkEC.CK_Cart_Sheet
            Wsheet.Range("C" & LigneECNommage) = checkEC.CK_Cart_NumeroPlan
            Wsheet.Range("D" & LigneECNommage) = checkEC.Val_NumDrawing
            LigneECNommage = LigneECNommage + 1
        End If
   GoTo Fin
Erreur:
    tResult = Err.Description
Fin:
    vNumerotation = tResult
    On Error GoTo 0
End Function

Private Function vVues(Wsheet, checkEC, LigneECVues, nDrawEc) As String
'Vérification des vues
Dim CheckVues As Boolean
    CheckVues = True
    On Error GoTo Erreur
    
          Wsheet.Range("A" & LigneECVues) = nDrawEc
          CouleurCell Wsheet, 1, LigneECVues, "bleu"
          LigneECVues = LigneECVues + 1
          
          Wsheet.Range("A" & LigneECVues) = "Calque sauvé dans Working View"
          If checkEC.CK_WorkingVueActive Then
              CouleurCell Wsheet, 1, LigneECVues, "vert"
          Else
              CouleurCell Wsheet, 1, LigneECVues, "rouge"
              CheckVues = False
          End If
          LigneECVues = LigneECVues + 1
        
          Wsheet.Range("A" & LigneECVues) = "Cadre Vues caché"
          If checkEC.CK_CadreVue Then
              CouleurCell Wsheet, 1, LigneECVues, "vert"
          Else
              CouleurCell Wsheet, 1, LigneECVues, "rouge"
              CheckVues = False
          End If
          LigneECVues = LigneECVues + 1
          
          Wsheet.Range("A" & LigneECVues) = "Vues vérouillées"
          If checkEC.CK_LockVue Then
              CouleurCell Wsheet, 1, LigneECVues, "vert"
          Else
              CouleurCell Wsheet, 1, LigneECVues, "rouge"
              CheckVues = False
          End If
          LigneECVues = LigneECVues + 1
                 
          Wsheet.Range("A" & LigneECVues) = "Vue ISO présente"
          If checkEC.CK_VueIso Then
              CouleurCell Wsheet, 1, LigneECVues, "vert"
          Else
              CouleurCell Wsheet, 1, LigneECVues, "rouge"
              CheckVues = False
          End If
          LigneECVues = LigneECVues + 2
          GoTo Fin:
Erreur:
    CheckVues = False
Fin:
    If CheckVues Then vVues = "OK" Else vVues = "KO"
            
End Function

Private Function vCartouche(Wsheet, checkEC, LigneECParam, nDrawEc, DerPl, InfoCartRecurentes) As String
'Vérification des cartouches et des paramètres
Dim CheckCart As Boolean
Dim i As Long

On Error GoTo Erreur
            
    Wsheet.Range("A" & LigneECParam) = nDrawEc
    If checkEC.CK_Lien3DName = "" Then
        Wsheet.Range("B" & LigneECParam) = "Lien non trouvé"
    Else
        Wsheet.Range("B" & LigneECParam) = checkEC.CK_Lien3DName
    End If
    
    For i = 1 To 5
        CouleurCell Wsheet, i, LigneECParam, "bleu"
    Next i
    LigneECParam = LigneECParam + 1
            
    'Design Outillage
    Wsheet.Range("A" & LigneECParam) = "Nom outillage"
    Wsheet.Range("B" & LigneECParam) = checkEC.Val_NomPulsGSE_DesignOutillage
    Wsheet.Range("C" & LigneECParam) = checkEC.Val_Cart_Title
                    
    If Len(checkEC.Val_Cart_Title) = 0 Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    ElseIf checkEC.Val_Cart_Title <> InfoCartRecurentes(8) Then
        AjoutComment Wsheet, "A", LigneECParam, "Le titre n'est pas identique aux autres planches"
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    ElseIf Len(checkEC.Val_Cart_Title) > 21 Then
        CouleurCell Wsheet, 1, LigneECParam, "jaune"
        AjoutComment Wsheet, "A", LigneECParam, "Désignation supèrieure à 21 carractères"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
            
    'Num Outillage
    Wsheet.Range("A" & LigneECParam) = "Numéro outillage"
    Wsheet.Range("B" & LigneECParam) = checkEC.Val_NomPulsGSE_NoOutillage
    Wsheet.Range("C" & LigneECParam) = checkEC.CK_Cart_NumeroPlan
    If checkEC.CK_Cart_NumeroPlan <> InfoCartRecurentes(0) Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
            
    'SiteAB
    Wsheet.Range("A" & LigneECParam) = "Site AB"
    Wsheet.Range("B" & LigneECParam) = checkEC.Val_NomPulsGSE_SiteAB
    Wsheet.Range("C" & LigneECParam) = checkEC.CK_Cart_Division
    If checkEC.CK_Cart_Division <> InfoCartRecurentes(1) Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
        
    'DRN dans Cartouche
    Wsheet.Range("A" & LigneECParam) = "DRN"
    Wsheet.Range("B" & LigneECParam) = ""
    Wsheet.Range("C" & LigneECParam) = checkEC.CK_Cart_DRN
    Wsheet.Range("D" & LigneECParam) = checkEC.CK_Mod_Design
    If (checkEC.CK_Cart_DRN <> InfoCartRecurentes(2)) Or (checkEC.CK_Cart_DRN <> checkEC.CK_Mod_Design) Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
            
    'CHK
    Wsheet.Range("A" & LigneECParam) = "CHK"
    Wsheet.Range("B" & LigneECParam) = checkEC.CK_NomPulsGSE_CHK
    Wsheet.Range("C" & LigneECParam) = checkEC.CK_Cart_CHK
    If checkEC.CK_Cart_CHK <> InfoCartRecurentes(3) Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
            
    'D.O.Orig
    Wsheet.Range("A" & LigneECParam) = "D.O.Orig"
    Wsheet.Range("C" & LigneECParam) = checkEC.CK_Cart_DOOrig
    If checkEC.CK_Cart_DOOrig <> InfoCartRecurentes(4) Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
            
    'Dwg.Syst
    Wsheet.Range("A" & LigneECParam) = "Dwg. Syst"
    Wsheet.Range("C" & LigneECParam) = checkEC.CK_Cart_DWGSys
    If checkEC.CK_Cart_DWGSys <> InfoCartRecurentes(5) Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
            
    'Process
    Wsheet.Range("A" & LigneECParam) = "Process"
    Wsheet.Range("C" & LigneECParam) = checkEC.CK_Cart_Process
    If checkEC.CK_Cart_Process <> InfoCartRecurentes(6) Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
            
    'DatePlan
    Wsheet.Range("A" & LigneECParam) = "Date Plan"
    Wsheet.Range("B" & LigneECParam) = "'" & checkEC.CK_NomPulsGSE_DatePlan
    Wsheet.Range("D" & LigneECParam) = "'" & checkEC.CK_Mod_Date
    If checkEC.CK_Mod_Date <> InfoCartRecurentes(7) Then
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    Else
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    End If
    LigneECParam = LigneECParam + 1
        
    'test l'existance d'un tableau de Nomenclature
    If checkEC.CK_Exist_TabNom Then
        'sheet
        Wsheet.Range("A" & LigneECParam) = "N° de planche"
        
        'Cas de l'ensemble général avec le nombre total de planches
        If InStr(checkEC.CK_Cart_Sheet, "/") > 0 Then
            Temp_NB_Planche = Right(checkEC.CK_Cart_Sheet, Len(checkEC.CK_Cart_Sheet) - InStr(checkEC.CK_Cart_Sheet, "/"))
            Wsheet.Range("B" & LigneECParam) = "'" & checkEC.CK_NomPulsGSE_Sheet
            Wsheet.Range("C" & LigneECParam) = "'" & Left(checkEC.CK_Cart_Sheet, Len(checkEC.CK_Cart_Sheet) - Len(Temp_NB_Planche) - 1)
            Wsheet.Range("E" & LigneECParam) = "Nombre total de planche = " & Temp_NB_Planche
            CouleurCell Wsheet, 1, LigneECParam, "vert"
        
        'cas de la dernière planche
        ElseIf i = DerPl Then
            Wsheet.Range("B" & LigneECParam) = "'" & checkEC.CK_NomPulsGSE_Sheet
            Wsheet.Range("C" & LigneECParam) = "'" & checkEC.CK_Cart_Sheet
            If checkEC.CK_Cart_Sheet <> checkEC.CK_NomPulsGSE_Sheet Then
                Wsheet.Range("E" & LigneECParam) = "Le nombre total de planches n'est pas correct"
                CouleurCell Wsheet, 1, LigneECParam, "rouge"
                CheckCart = False
            Else
                CouleurCell Wsheet, 1, LigneECParam, "vert"
            End If
        
        'autres cas
        Else
            Wsheet.Range("B" & LigneECParam) = "'" & checkEC.CK_NomPulsGSE_Sheet
            Wsheet.Range("C" & LigneECParam) = "'" & checkEC.CK_Cart_Sheet
            
            If checkEC.CK_NomPulsGSE_Sheet = checkEC.CK_Cart_Sheet Then
                CouleurCell Wsheet, 1, LigneECParam, "vert"
            'Cas de Part détaillé sur plusieurs planches
            ElseIf InStr(checkEC.CK_NomPulsGSE_Sheet, ",") > 0 Then
                If InStr(checkEC.CK_NomPulsGSE_Sheet, checkEC.CK_Cart_Sheet) > 0 Then
                    CouleurCell Wsheet, 1, LigneECParam, "vert"
                Else
                    CouleurCell Wsheet, 1, LigneECParam, "rouge"
                    CheckCart = False
                End If
            Else
                CouleurCell Wsheet, 1, LigneECParam, "rouge"
                CheckCart = False
            End If
            ' sur les pièces fabriquées, le N° de planche ne doit pas être vide
            If TypeElement(checkEC.Val_Num3D, checkEC.CK_NomPulsGSE_TypeNum) = 5 And checkEC.CK_NomPulsGSE_Sheet = "" Then
                Wsheet.Range("E" & LigneECParam) = "Le N° de planche est vide"
                CouleurCell Wsheet, 14, LigneECParam, "rouge"
                CheckCart = False
            End If
        End If
        LigneECParam = LigneECParam + 1
                
        'Dimension
        Wsheet.Range("A" & LigneECParam) = "Dimension"
        Wsheet.Range("B" & LigneECParam) = checkEC.CK_NomPulsGSE_Dimension
        Wsheet.Range("C" & LigneECParam) = checkEC.CK_Table_Dimension
        'NON vide pour les pièces fabriquées
        If TypeElement(checkEC.Val_Num3D, checkEC.CK_NomPulsGSE_TypeNum) = 8 And checkEC.CK_Table_Dimension = "" Then
            Wsheet.Range("E" & LigneECParam) = "Dimension ne doit pas être vide"
            CouleurCell Wsheet, 1, LigneECParam, "rouge"
            CheckCart = False
        Else
            CouleurCell Wsheet, 1, LigneECParam, "vert"
        End If
        LigneECParam = LigneECParam + 1

        'Material
        Wsheet.Range("A" & LigneECParam) = "Material"
        'Debug.Print CheckEC.Val_NomPulsGSE_Material
        Wsheet.Range("B" & LigneECParam) = "'" & SuprSautLigne(checkEC.Val_NomPulsGSE_Material)
        Wsheet.Range("C" & LigneECParam) = "'" & SupSautLigne(checkEC.CK_Table_Material)
        If SuprSautLigne(checkEC.Val_NomPulsGSE_Material) <> SupSautLigne(checkEC.CK_Table_Material) Then
            CouleurCell Wsheet, 1, LigneECParam, "rouge"
            CheckCart = False
        'NON vide pour les pièces fabriquées
        ElseIf TypeElement(checkEC.Val_Num3D, checkEC.CK_NomPulsGSE_TypeNum) = 8 And checkEC.CK_Table_Material = "" Then
            Wsheet.Range("E" & LigneECParam) = "Material ne doit pas être vide"
            CouleurCell Wsheet, 1, LigneECParam, "rouge"
            CheckCart = False
        Else
            CouleurCell Wsheet, 1, LigneECParam, "vert"
        End If
        LigneECParam = LigneECParam + 1
            
        'Protect
        Wsheet.Range("A" & LigneECParam) = "Protect"
        Wsheet.Range("B" & LigneECParam) = "'" & SuprSautLigne(checkEC.Val_NomPulsGSE_Protect)
        Wsheet.Range("C" & LigneECParam) = "'" & SupSautLigne(checkEC.Val_Table_Protect)
        'debug.print SuprSautLigne(CheckEC.Val_NomPulsGSE_Protect)
        'debug.print SupSautLigne(CheckEC.Val_Table_Protect)
        If SuprSautLigne(checkEC.Val_NomPulsGSE_Protect) = SupSautLigne(checkEC.Val_Table_Protect) Then
            CouleurCell Wsheet, 1, LigneECParam, "vert"
        Else
            CouleurCell Wsheet, 1, LigneECParam, "rouge"
            AjoutComment Wsheet, "A", LigneECParam, "La valeur dans la nomenclature est différent du paramètre 3D"
            CheckCart = False
        End If
        
        'Si Matérial = Acier, alor traitemeant obligatoire
        If checkEC.Val_Matiere_PartLie = "ACIER" Or checkEC.Val_Matiere_PartLie = "STEEL" Then
        'If UCase(CheckEC.CK_Table_Material) = "ACIER" Or UCase(CheckEC.CK_Table_Material) = "STEEL" Then
            If checkEC.Val_Table_Protect = "" Then
                Wsheet.Range("E" & LigneECParam) = "'Material = Acier => Protection obligatoire"
                AjoutComment Wsheet, "A", LigneECParam, "Material = Acier => Protection obligatoire"
                CouleurCell Wsheet, 1, LigneECParam, "rouge"
                CheckCart = False
            End If
        End If
        LigneECParam = LigneECParam + 1
            
        'Weight
        Wsheet.Range("A" & LigneECParam) = "Weight"
        Wsheet.Range("B" & LigneECParam) = checkEC.CK_NomPulsGSE_Weight
        Wsheet.Range("C" & LigneECParam) = checkEC.Val_Table_Weight
        Wsheet.Range("E" & LigneECParam) = "Poids réel de la pièce : " & checkEC.Val_Masse
        If checkEC.CK_NomPulsGSE_Weight = checkEC.Val_Table_Weight Then
            CouleurCell Wsheet, 1, LigneECParam, "vert"
        ElseIf checkEC.CK_Weight Then
            CouleurCell Wsheet, 1, LigneECParam, "jaune"
            AjoutComment Wsheet, "A", LigneECParam, "Le poids suppérieur a +/- 5%"
            CheckCart = False
        Else
            CouleurCell Wsheet, 1, LigneECParam, "rouge"
            AjoutComment Wsheet, "A", LigneECParam, "Le poids est absent"
            CheckCart = False
        End If
        LigneECParam = LigneECParam + 1
        
    Else 'Pas de tableau de nomenclature
        Wsheet.Range("E" & LigneECParam) = "Pas de tableau de nomenclature détecté"
        For i = 13 To 19
            CouleurCell Wsheet, 1, LigneECParam, "rouge"
            CheckCart = False
        Next i
    End If
            
    'Indice de modif
    Wsheet.Range("A" & LigneECParam) = "Issue"
    Wsheet.Range("C" & LigneECParam) = checkEC.Val_Cart_Issue
    Wsheet.Range("D" & LigneECParam) = checkEC.CK_Mod_Issue
    If checkEC.Val_Cart_Issue = checkEC.CK_Mod_Issue And checkEC.Val_Cart_Issue = "1" Then
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    ElseIf checkEC.Val_Cart_Issue = checkEC.CK_Mod_Issue And checkEC.Val_Cart_Issue <> "1" Then
        AjoutComment Wsheet, "A", LigneECParam, "Indice différent de 1"
        CouleurCell Wsheet, 1, LigneECParam, "jaune"
    Else
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    End If
    LigneECParam = LigneECParam + 1
        
    'Format du plan
    Wsheet.Range("A" & LigneECParam) = "Format"
    Wsheet.Range("C" & LigneECParam) = checkEC.CK_Format_Plan
    Wsheet.Range("D" & LigneECParam) = checkEC.CK_Cart_Format
    If InStr(checkEC.CK_Format_Plan, checkEC.CK_Cart_Format) > 0 Then
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    Else
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    End If
    LigneECParam = LigneECParam + 1
        
    'Echelle des vues
    Wsheet.Range("A" & LigneECParam) = "Echelle"
    Wsheet.Range("C" & LigneECParam) = "'" & checkEC.Val_Cart_Echelle
    
    If checkEC.CK_Echelle Then
        CouleurCell Wsheet, 1, LigneECParam, "vert"
    Else
        Wsheet.Range("E" & LigneECParam) = "pas de vue à l'echelle indiquée dans le plan"
        CouleurCell Wsheet, 1, LigneECParam, "rouge"
        CheckCart = False
    End If
        
    'Soulignage
    LigneECParam = LigneECParam + 1
     With Wsheet.Range("A" & LigneECParam & ":E" & LigneECParam).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous '1
        .ColorIndex = xlAutomatic '-4105
        .TintAndShade = 0
        .Weight = xlThin '2
     End With
    LigneECParam = LigneECParam + 1
    If CheckCart Then vCartouche = "OK" Else vCartouche = "KO"
    GoTo Fin

Erreur:
    vCartouche = "Erreur"
Fin:
    
End Function

Private Function vCotes(Wsheet, checkEC, LigneECCotes, nDrawEc) As String
'Vérification des cotes et bullages
Dim CheckCotes As Boolean
Dim Concat As String
Dim listBullageTemp() As String
Dim ListCotesTemp() As String
Dim i As Long

On Error GoTo Erreur

    CheckCotes = True
    listBullageTemp = checkEC.CK_Bullage6
           
    If listBullageTemp(0, 0) = "KO" Then
        CheckCotes = False
        If UBound(listBullageTemp, 1) > 1 Then
            For i = 1 To UBound(listBullageTemp, 2)
                Concat = Concat & "* " & listBullageTemp(0, i) & " - " & listBullageTemp(1, i) & " - " & listBullageTemp(2, i) & Chr(10) & Chr(13)
            Next
        Else
            Concat = ""
        End If
        LigneECCotes = LigneECCotes + 1
        Wsheet.Range("A" & LigneECCotes) = nDrawEc
        Wsheet.Range("B" & LigneECCotes) = Concat
        Concat = ""
    End If

    ListCotesTemp = checkEC.CK_CoteCasse
    If ListCotesTemp(0, 0) = "KO" Then
        CheckCotes = False
        If UBound(ListCotesTemp, 1) > 1 Then
            For i = 1 To UBound(ListCotesTemp, 2)
                Concat = Concat & "* " & ListCotesTemp(0, i) & " - " & ListCotesTemp(1, i) & " - " & ListCotesTemp(2, i) & Chr(10) & Chr(13)
            Next
        Else
            Concat = ""
        End If
        LigneECCotes = LigneECCotes + 1
        Wsheet.Range("A" & LigneECCotes) = nDrawEc
        Wsheet.Range("C" & LigneECCotes) = Concat
        Concat = ""
    End If
    GoTo Fin
        
Erreur:
    CheckCotes = False
    On Error GoTo 0
Fin:
    If CheckCotes Then vCotes = "OK" Else vCotes = "KO"
       
End Function

Private Function vNomenclature(Wsheet, checkEC, LigneECNom, nDrawEc) As String
'Vérification de la nomenclature
Dim Decal_Col As Integer
Dim CheckNom As Boolean
Dim Gimp_Detecte As Boolean
Dim BulleInPlan As Boolean
Dim tmpPartNb As String, tmpItemNb As String
Dim List_Bulles() As String
Dim Rep_Nom As String
Dim i As Long, j As Long
Dim NBCol As Long, NbLig As Long
On Error GoTo Erreur

    CheckNom = True
    'Entète de Nomenclature
    Wsheet.Range("A" & LigneECNom + 1) = nDrawEc
    With Wsheet.Range("A" & LigneECNom + 1 & ":O" & LigneECNom + 1)
        .Font.Size = 11
        .Font.Bold = False
        .Interior.Color = 15917714
    End With
    LigneECNom = LigneECNom + 1

    'Dimension du tableau de nomenclature
    NBCol = checkEC.CK_Table_Nom.NumberOfColumns
    NbLig = checkEC.CK_Table_Nom.NumberOfRows
    Decal_Col = 14 - NBCol
    
    'trace la nomenclature
    For i = NbLig To 1 Step -1
        For j = NBCol + Decal_Col To 1 + Decal_Col Step -1
            Wsheet.Range(XlCol(j) & i + LigneECNom) = checkEC.CK_Table_Nom.GetCellString(i, j - Decal_Col)
        Next j
    Next i
    'Ajout des titre de colonnes suplémentaire
    Wsheet.Range(XlCol(NBCol + Decal_Col + 1) & NbLig + LigneECNom) = "Présence Bullage"
    
    'Mise en forme
    With Wsheet.Range("A" & NbLig + LigneECNom & ":O" & NbLig + LigneECNom)
        .Interior.Color = 15917714
    End With

    'Verif de la présence du poids de l'ensemble
    If checkEC.CK_Prem_Planche Then
        If Wsheet.Range(XlCol(NBCol + Decal_Col) & LigneECNom + NbLig - 1) = "" Then
            AjoutComment Wsheet, XlCol(NBCol + Decal_Col), LigneECNom + NbLig - 1, "Le poids doit être documenté pour l'ensemble général"
            CouleurCell Wsheet, NBCol + Decal_Col, LigneECNom + NbLig - 1, "rouge"
        ElseIf checkEC.CK_Masse5P(ExtractValNum(Wsheet.Range(XlCol(NBCol + Decal_Col) & LigneECNom + NbLig - 1))) Then
            CouleurCell Wsheet, NBCol + Decal_Col, LigneECNom + NbLig - 1, "vert"
        Else
            CouleurCell Wsheet, NBCol + Decal_Col, LigneECNom + NbLig - 1, "jaune"
            CheckNom = False
        End If

    ElseIf checkEC.Val_Masse > 25 Then
        'La masse doit être documentée si > 25 KG
        If Wsheet.Range(XlCol(NBCol + Decal_Col) & LigneECNom + NbLig - 1) = "" Then
                CouleurCell Wsheet, NBCol + Decal_Col, LigneECNom + NbLig - 1, "rouge"
                AjoutComment Wsheet, XlCol(NBCol + Decal_Col), LigneECNom + NbLig - 1, "Le poids est supérieur a 25 KG, il doit être indiqué dans la nomenclature"
        'La masse est différente de + de 5%
        ElseIf checkEC.CK_Masse5P(ExtractValNum(Wsheet.Range(XlCol(NBCol + Decal_Col) & LigneECNom + NbLig - 1))) Then
                CouleurCell Wsheet, NBCol + Decal_Col, LigneECNom + NbLig - 1, "jaune"
                AjoutComment Wsheet, XlCol(NBCol + Decal_Col), LigneECNom + NbLig - 1, checkEC.Val_Masse
        Else
                CouleurCell Wsheet, NBCol + Decal_Col, LigneECNom + NbLig - 1, "vert"
                CheckNom = False
        End If
    End If

    'Vérif de la Gimp sur l'ensemble général
    Gimp_Detecte = False
    If checkEC.CK_Prem_Planche Then
        If checkEC.Val_NomPulsGSE_PresUserGuide = "OUI" Then
            'Le N° de la Gimp doit être = au N° du dossier
            If InStr(checkEC.Val_Gimp, checkEC.Val_NomPulsGSE_NoOutillage) > 0 Then
                For i = 1 To NbLig
                    If Wsheet.Range(XlCol(NBCol + Decal_Col - 5) & LigneECNom + i) = "USER GUIDE" Then
                        CouleurCell Wsheet, NBCol + Decal_Col - 6, LigneECNom + i, "vert"
                        Gimp_Detecte = True
                    End If
                Next i
            'La gimp doit est présente
                If Not (Gimp_Detecte) Then
                    AjoutComment Wsheet, XlCol(NBCol + Decal_Col - 6), LigneECNom + 1, "La Gimp est absente"
                    CouleurCell Wsheet, NBCol + Decal_Col - 6, LigneECNom + 1, "rouge"
                    CheckNom = False
                End If
            Else
                AjoutComment Wsheet, XlCol(NBCol + Decal_Col - 6), LigneECNom + 1, "N° de Gimp incorect : " & checkEC.Val_Gimp
                CouleurCell Wsheet, NBCol + Decal_Col - 6, LigneECNom + 1, "rouge"
                CheckNom = False
            End If
        ElseIf checkEC.Val_NomPulsGSE_PresUserGuide = "" Then
            'recherche si la GIMP a été documentée dans la nom alors qu'il ne devrait pas y en avoir (pas de paramètre NomPulsGSE_PresUserGuide)
            For i = 1 To NbLig
                If Wsheet.Range(XlCol(NBCol + Decal_Col - 5) & LigneECNom + i) = "USER GUIDE" Then
                    CouleurCell Wsheet, NBCol + Decal_Col - 6, LigneECNom + i, "rouge"
                    AjoutComment Wsheet, XlCol(NBCol + Decal_Col - 6), LigneECNom + i, "pas de guimp prévue (NomPulsGSE_PresenceUserGuide absent)"
                End If
            Next i
        End If
    End If
    
    'Vérification que tous les Item Nb on le radical de l'outillage
    For i = 1 To NbLig - 1
        tmpPartNb = Wsheet.Range(XlCol(NBCol + Decal_Col - 6) & LigneECNom + i)
        tmpItemNb = Wsheet.Range(XlCol(NBCol + Decal_Col - 7) & LigneECNom + i)
        If TypeElementRep(tmpItemNb, checkEC.CK_NomPulsGSE_TypeNum) < 9 And TypeElementRep(tmpItemNb, checkEC.CK_NomPulsGSE_TypeNum) > -1 Then
            If InStr(tmpPartNb, checkEC.Val_NomPulsGSE_NoOutillage) > 0 Then
                CouleurCell Wsheet, NBCol + Decal_Col - 6, LigneECNom + i, "vert"
            Else
                CouleurCell Wsheet, NBCol + Decal_Col - 6, LigneECNom + i, "rouge"
                AjoutComment Wsheet, XlCol(NBCol + Decal_Col - 6), LigneECNom + i, "Radical du N° de dossier incorrect"
                CheckNom = False
            End If
        End If
    Next i

    'verification de la présence des Bullages
    List_Bulles = checkEC.Val_List_Bulle

    For i = 1 To NbLig - 3
        BulleInPlan = False
        For j = 1 To UBound(List_Bulles, 1)
            Rep_Nom = Wsheet.Range(XlCol(NBCol + Decal_Col - 7) & LigneECNom + i)
            'Debug.Print Rep_Nom
            If CInt(List_Bulles(j)) = ReplaceBlanck(Rep_Nom) Then
                Wsheet.Range(XlCol(NBCol + Decal_Col + 1) & LigneECNom + i) = "OK"
                CouleurCell Wsheet, NBCol + Decal_Col + 1, LigneECNom + i, "vert"
                BulleInPlan = True
                Exit For
            ElseIf Rep_Nom = "" And Wsheet.Range(XlCol(NBCol + Decal_Col - 6) & LigneECNom + i) = "" Then
                BulleInPlan = True
            End If

            If BulleInPlan Then
            Else
                Wsheet.Range(XlCol(NBCol + Decal_Col + 1) & LigneECNom + i) = "KO"
                CouleurCell Wsheet, NBCol + Decal_Col + 1, LigneECNom + i, "rouge"
                'CheckNom = False
            End If
        Next j
    Next i
    ReDim List_Bulles(0)

    LigneECNom = LigneECNom + NbLig + 2
    GoTo Fin
    
Erreur:
    CheckNom = False
Fin:
    If CheckNom Then vNomenclature = "Voir Onglet" Else vNomenclature = "KO"

End Function
