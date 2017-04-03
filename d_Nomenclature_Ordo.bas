Attribute VB_Name = "d_Nomenclature_Ordo"
Option Explicit

Sub catmain()
' *****************************************************************
' * Genere la nomenclature dans un fichier excel.
' * Exporte ensuite la nomenclature dans un fichier excel formaté pour l'ordonnancement
' *
' * Création CFR le 07/01/16
' * Modification le :
' *
' *
' *****************************************************************

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "d_Nomenclature_Ordo", VMacro

'Chargement des variables
Dim MacroLocation As New xMacroLocation
Dim NomTemplateNomExcel As String

If Not (MacroLocation.FicIniExist("VarNomenclature.ini")) Then ' on est pas dans l'environnement GSE_R21
    MsgBox "Vous n'etes pas dans l'environnement GRE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
    Exit Sub
Else
    MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
    CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
    CheminDestNomenclature = MacroLocation.ValVar("CheminDestNomenclature")
    NomTemplateNomExcel = MacroLocation.ValVar("NomTemplateNomExcel")
End If

'Choix de la langue
Dim ChoixLangue As String
    Load Frm_Langue
    Frm_Langue.Show
    If Frm_Langue.RB_EN Then
        ChoixLangue = "EN"
    Else
        ChoixLangue = "FR"
    End If
    Frm_Langue.Hide
    If Not (Frm_Langue.ChB_OkAnnule) Then
        End
    End If
    Unload Frm_Langue

Dim Barre As New c_ProgressBarre
    Barre.Progression = 2
    Barre.Titre = "Export Ordo"

'Variables
Dim NoSSESwith As Boolean

Barre.Progression = 5
'ouvre le Product liè au Catdrawing
    Set ActiveDoc = CATIA.ActiveDocument
    Dim ProductDoc As Product
    Set ProductDoc = ActiveDoc.Product
    Barre.Progression = 10

'verifie si un fichier de nomenclature est déja présent et l'efface
    If Not (EffaceFicNom(CheminDestNomenclature, ActiveDoc.Name & ".xls")) Then
        End
    End If

'Formatage de la nomenclature en fonction de la langue du Catia
Dim LangueQt, LangueRef, LangueDesc As String
    If ChoixLangue = "EN" Then
        LangueQt = "Quantity"
        LangueRef = "Part Number"
        LangueDesc = "Product Description"
    ElseIf ChoixLangue = "FR" Then
        LangueQt = "Quantité"
        LangueRef = "Référence"
        LangueDesc = "Description du produit"
    Else
        MsgBox "Erreur dans la détection de la langue paramétré dans Catia."
        End
    End If

'Extraction de la nomenclature et sauvegarde dans un fichier excel
    Dim assemblyConvertor1Variant
    Dim assemblyConvertor1 As AssemblyConvertor
    Set assemblyConvertor1 = ProductDoc.GetItem("BillOfMaterial")

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

    assemblyConvertor1.[Print] "XLS", CStr(CheminDestNomenclature & ActiveDoc.Name & ".xls"), ProductDoc
    Barre.Progression = 20
    
'Fermeture du produit lié
'    ActiveDoc.Close

'création d'une table de nomenclature
    Dim TableNomTempo() As String
    Dim LigneNomTempo As String
    Dim Boucle As Long
        Boucle = 0
    
'## Export de la nomenclature dans le fichier template nomenclature PAC destiné a l'Ordo ##
'Creation d'un objet eXcel pour stocker la nomenclature brute
    Dim objexcel
    Dim objWorkBook
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objexcel.Workbooks.Open(CStr(CheminDestNomenclature & ActiveDoc.Name & ".xls"))
    objexcel.Visible = True
    objWorkBook.ActiveSheet.Visible = True

'Récupération du type de numérotation
    Dim TypeNum As String
    Dim ActiveProd As Product
    Set ActiveProd = ActiveDoc.Product
    Dim MesParametres As Parameters
    Set MesParametres = ActiveProd.UserRefProperties
    TypeNum = RecupParam(MesParametres, "NomPulsGSE_TypeNum")

'Creation d'un objet eXcel et ouverture du fichier Template
    Dim objExcelNomOrdo
    Dim objWorkBookOrdo
    Set objExcelNomOrdo = CreateObject("EXCEL.APPLICATION")
    Set objWorkBookOrdo = objExcelNomOrdo.Workbooks.Open(CStr(CheminSourcesMacro & NomTemplateNomExcel))
    objExcelNomOrdo.Visible = True
    objWorkBookOrdo.ActiveSheet.Visible = True

'paramétrage des colonnes du fichier Template
    Dim ParamTemplate(10) As String
    ParamTemplate(0) = "C" 'Repere              - NomPulsGSE_ItemNB
    ParamTemplate(1) = "E" 'Quantité unitaire   - Quantity
    ParamTemplate(2) = "H" 'Designation         - Product_Description ou NomPulsGSE-DesignOutillage
    ParamTemplate(3) = "D" 'Reference           - NomPulsGSE_SupplierRef
    ParamTemplate(4) = "I" 'Marque              - NomPulsGSEMiscellanous
    ParamTemplate(5) = "G" 'Sheet               - NomPulsGSE_Sheet
    ParamTemplate(6) = "A" 'N° Assemblage       - NomPulsGSE_N0Outillage
    ParamTemplate(7) = "B" 'N° Sous Ensemble    - Part_Number
    ParamTemplate(8) = "L" 'Quantité à commander- formule Qté a commander
    ParamTemplate(9) = "M" 'Type Ordo           - en fonction partnumber
    ParamTemplate(10) = "Q" 'Traitement         - NomPulsGSE_Protect
    
    Dim CellQteAss As String
    CellQteAss = "$E$3" 'Cellule contenant la quantité d'assemblage
    Dim FormuleQtCdeSSE As String, FormuleQteCdePiece As String
    Dim ligactiveOrdo As Long, LigActiveNom As Long, LigGrpDeb As Long
    ligactiveOrdo = 3
    LigActiveNom = 5
    
'Construction de la liste des sous-ensembles avec leur localisation (ligne du fichier excel)
    Dim ListeSSE() As String
    Dim NoDerniereLigneNom As Long
        NoDerniereLigneNom = NoDerniereLigne(objWorkBook)
    Dim NoDerniereLigneEns As Long
        NoDerniereLigneEns = NoDebRecap(objWorkBook, ChoixLangue)
    Dim NoLigDebSSE As Long
        NoLigDebSSE = 3
    Dim h As Long, i As Long, j As Integer
    Dim CTNiv As Long
        CTNiv = 0
    
    'Si pas de sous-ensemble on evitera les parties de code inutiles
    NoSSESwith = True
    For i = 3 To NoDerniereLigneEns
        If TestEstSSE(objWorkBook.ActiveSheet.cells(i, 1).Value, ChoixLangue) Then
            NoSSESwith = False
            ReDim Preserve ListeSSE(2, CTNiv)
            ListeSSE(0, CTNiv) = NomSSE(objWorkBook.ActiveSheet.cells(i, 1).Value, ChoixLangue)
            ListeSSE(1, CTNiv) = i 'N° de ligne du départ de la nomenclature du SSE
            CTNiv = CTNiv + 1
        End If
        Barre.Progression = 30
    Next i
If Not (NoSSESwith) Then 'pas de sous-ensemble étecté

'Construction de la liste des sous-ensembles de niveau 1
    Dim ListeSSENiv1() As String
    CTNiv = 0
    For i = ListeSSE(1, 0) + 2 To ListeSSE(1, 1) - 2 'ListeSSE(x, 0) est l'ensemble général
        For j = 0 To UBound(ListeSSE, 2)
            If objWorkBook.ActiveSheet.cells(i, 4).Value = ListeSSE(0, j) Then
                ReDim Preserve ListeSSENiv1(3, CTNiv)
                ListeSSENiv1(0, CTNiv) = ListeSSE(0, 0)
                ListeSSENiv1(1, CTNiv) = 1
                ListeSSENiv1(2, CTNiv) = objWorkBook.ActiveSheet.cells(i, 4).Value ' Nom SSE
                ListeSSENiv1(3, CTNiv) = objWorkBook.ActiveSheet.cells(i, 1).Value ' Qte SSe
                CTNiv = CTNiv + 1
            End If
        Next j
        Barre.Progression = (30 + (10 / SupDivZero(ListeSSE(1, 0) + 2 - (ListeSSE(1, 1) - 2))) * i)
    Next i

'Construction de la liste des sous-ensembles de niveau 2
    Dim ListeSSENiv2() As String
    Dim LigDebSSE As Long, LigFinSSE As Long
    CTNiv = 0
    For h = 0 To UBound(ListeSSENiv1, 2)
        For j = 0 To UBound(ListeSSE, 2)
            If ListeSSE(0, j) = ListeSSENiv1(2, h) Then
                LigDebSSE = ListeSSE(1, j)
                If j = UBound(ListeSSE, 2) Then
                    LigFinSSE = NoDerniereLigneEns
                Else
                    LigFinSSE = ListeSSE(1, j + 1)
                End If
                Exit For
            End If
        Next j
        For i = LigDebSSE + 2 To LigFinSSE - 2
            For j = 0 To UBound(ListeSSE, 2)
                If objWorkBook.ActiveSheet.cells(i, 4).Value = ListeSSE(0, j) Then
                    ReDim Preserve ListeSSENiv2(5, CTNiv)
                    ListeSSENiv2(0, CTNiv) = ListeSSE(0, 0)
                    ListeSSENiv2(1, CTNiv) = 1
                    ListeSSENiv2(2, CTNiv) = ListeSSENiv1(2, h)
                    ListeSSENiv2(3, CTNiv) = ListeSSENiv1(3, h)
                    ListeSSENiv2(4, CTNiv) = objWorkBook.ActiveSheet.cells(i, 4).Value
                    ListeSSENiv2(5, CTNiv) = objWorkBook.ActiveSheet.cells(i, 1).Value
                    CTNiv = CTNiv + 1
                End If
            Next j
            Barre.Progression = (40 + (10 / SupDivZero(UBound(ListeSSENiv1, 2))) * h)
        Next i
    Next h

'Construction de la liste des sous-ensembles de niveau 3
    Dim ListeSSENiv3() As String
    Dim DerSSE As Boolean
    CTNiv = 0
    For h = 0 To UBound(ListeSSENiv2, 2)
        DerSSE = True
        For j = 0 To UBound(ListeSSE, 2)
            If ListeSSE(0, j) = ListeSSENiv2(4, h) Then
                LigDebSSE = ListeSSE(1, j)
                If j = UBound(ListeSSE, 2) Then
                    LigFinSSE = NoDerniereLigneEns
                Else
                    LigFinSSE = ListeSSE(1, j + 1)
                End If
            End If
        Next j
        For i = LigDebSSE + 2 To LigFinSSE - 2
            For j = 0 To UBound(ListeSSE, 2)
                If objWorkBook.ActiveSheet.cells(i, 4).Value = ListeSSE(0, j) Then
                    ReDim Preserve ListeSSENiv3(8, CTNiv)
                    ListeSSENiv3(0, CTNiv) = ListeSSE(0, 0)
                    ListeSSENiv3(1, CTNiv) = 1
                    ListeSSENiv3(2, CTNiv) = ListeSSENiv2(2, h)
                    ListeSSENiv3(3, CTNiv) = ListeSSENiv2(3, h)
                    ListeSSENiv3(4, CTNiv) = ListeSSENiv2(4, h)
                    ListeSSENiv3(5, CTNiv) = ListeSSENiv2(5, h)
                    ListeSSENiv3(6, CTNiv) = objWorkBook.ActiveSheet.cells(i, 4).Value
                    ListeSSENiv3(7, CTNiv) = objWorkBook.ActiveSheet.cells(i, 1).Value
                    ListeSSENiv3(8, CTNiv) = ListeSSENiv3(1, CTNiv) * ListeSSENiv3(3, CTNiv) * ListeSSENiv3(5, CTNiv) * objWorkBook.ActiveSheet.cells(i, 1).Value
                    CTNiv = CTNiv + 1
                    DerSSE = False
                End If
            Next j
            Barre.Progression = (50 + (10 / SupDivZero(UBound(ListeSSENiv2, 2))) * h)
        Next i
        If DerSSE Then
            ReDim Preserve ListeSSENiv3(8, CTNiv)
            ListeSSENiv3(0, CTNiv) = ListeSSE(0, 0)
            ListeSSENiv3(1, CTNiv) = 1
            ListeSSENiv3(2, CTNiv) = ListeSSENiv2(2, h)
            ListeSSENiv3(3, CTNiv) = ListeSSENiv2(3, h)
            ListeSSENiv3(4, CTNiv) = ""
            ListeSSENiv3(5, CTNiv) = 1
            ListeSSENiv3(6, CTNiv) = ListeSSENiv2(4, h)
            ListeSSENiv3(7, CTNiv) = ListeSSENiv2(5, h)
            ListeSSENiv3(8, CTNiv) = ListeSSENiv3(1, CTNiv) * ListeSSENiv3(3, CTNiv) * ListeSSENiv3(5, CTNiv) * ListeSSENiv3(7, CTNiv)
            CTNiv = CTNiv + 1
        End If
    Next h
End If

'## temp ecriture du résultat dans un excel
'    objWorkBook.Sheets.Add
'    For i = 0 To UBound(ListeSSENiv3, 2)
'
'        objWorkBook.ActiveSheet.cells(i + 1, "A") = ListeSSENiv3(0, i)
'        objWorkBook.ActiveSheet.cells(i + 1, "B") = ListeSSENiv3(1, i)
'        objWorkBook.ActiveSheet.cells(i + 1, "C") = ListeSSENiv3(2, i)
'        objWorkBook.ActiveSheet.cells(i + 1, "D") = ListeSSENiv3(3, i)
'        objWorkBook.ActiveSheet.cells(i + 1, "E") = ListeSSENiv3(4, i)
'        objWorkBook.ActiveSheet.cells(i + 1, "F") = ListeSSENiv3(5, i)
'        objWorkBook.ActiveSheet.cells(i + 1, "G") = ListeSSENiv3(6, i)
'        objWorkBook.ActiveSheet.cells(i + 1, "H") = ListeSSENiv3(7, i)
'        objWorkBook.ActiveSheet.cells(i + 1, "I") = ListeSSENiv3(8, i)
'
'    Next i
'    objWorkBook.Sheets("Feuil1").Select

'additionne les qte
Dim TotQte As Integer
For i = 0 To UBound(ListeSSE, 2)
TotQte = 0
'Debug.Print "Addition"
    For j = 0 To UBound(ListeSSENiv3, 2)
        If ListeSSENiv3(6, j) = ListeSSE(0, i) Then
            TotQte = TotQte + ListeSSENiv3(8, j)
        End If
    Next j
    If TotQte = 0 Then
        For j = 0 To UBound(ListeSSENiv2, 2)
            If ListeSSENiv2(4, j) = ListeSSE(0, i) Then
                TotQte = TotQte + ListeSSENiv2(5, j)
            End If
        Next j
        If TotQte = 0 Then
            For j = 0 To UBound(ListeSSENiv1, 2)
                If ListeSSENiv1(2, j) = ListeSSE(0, i) Then
                    TotQte = TotQte + ListeSSENiv1(3, j)
                End If
            Next j
        End If
        If TotQte = 0 Then
            TotQte = TotQte + ListeSSENiv1(1, 0)
        End If
    End If
    
    ListeSSE(2, i) = TotQte
Next i


'## temp ecriture du résultat dans un excel
'objWorkBook.Sheets("Feuil2").Select
'For i = 0 To UBound(ListeSSE, 2)
'    objWorkBook.ActiveSheet.cells(i + 1, "L") = ListeSSE(0, i)
'    objWorkBook.ActiveSheet.cells(i + 1, "M") = ListeSSE(2, i)
'Next i

' Ecriture du nom de l'assemblage et Quantité : 1
    objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(6)).Value = NomMachine(objWorkBook, NoDerniereLigneNom, ChoixLangue)
    objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(1)).Value = 1
    ligactiveOrdo = ligactiveOrdo + 1
    'Mise forme de la ligne de l'assemblage
    objWorkBookOrdo.ActiveSheet.Range("A3:S3").Interior.Color = 16777164
    
    'Pour chaque ligne de composants de premier niveau
    Do While Not objWorkBook.ActiveSheet.cells(LigActiveNom, 1).Value = ""
        'test si c'est un sous ensemble
        Dim EstSSE As Boolean
        EstSSE = False
        If Not (NoSSESwith) Then 'pas de sous-ensemble étecté
            Dim LigneSE
            For Each LigneSE In ListeSSE()
                EstSSE = False
                If objWorkBook.ActiveSheet.cells(LigActiveNom, 3).Value = LigneSE Then 'C'est un SSE
                    EstSSE = True
                    Exit For
                End If
            Next
        End If
        If Not (EstSSE) Then 'C'est une pièce
            ' Ecriture de la ligne pièce
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(0)).Value = objWorkBook.ActiveSheet.cells(LigActiveNom, 3).Value 'Item Nb
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(1)).Value = objWorkBook.ActiveSheet.cells(LigActiveNom, 1).Value 'Qte
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(2)).Value = objWorkBook.ActiveSheet.cells(LigActiveNom, 6).Value 'Designation
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(3)).Value = objWorkBook.ActiveSheet.cells(LigActiveNom, 5).Value ' SupplierRef
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(4)).Value = objWorkBook.ActiveSheet.cells(LigActiveNom, 10).Value 'Miscellanous
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(5)).Value = objWorkBook.ActiveSheet.cells(LigActiveNom, 2).Value 'Sheet
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(9)).Value = TypeElmOrdo(objWorkBook.ActiveSheet.cells(LigActiveNom, 4).Value, TypeNum) 'Type Ordo
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(10)).Value = objWorkBook.ActiveSheet.cells(LigActiveNom, 9).Value 'Protect
            FormuleQteCdePiece = "=IF(((E" & ligactiveOrdo & "*" & CellQteAss & ")-K" & ligactiveOrdo & ")<0,0,(E" & ligactiveOrdo & "*" & CellQteAss & ")-K" & ligactiveOrdo & ")"
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(8)).Formula = FormuleQteCdePiece
            ligactiveOrdo = ligactiveOrdo + 1
        End If
    LigActiveNom = LigActiveNom + 1
    Loop
    Barre.Progression = 70
If Not (NoSSESwith) Then 'pas de sous-ensemble étecté
'Pour chaque SSE
    Dim LignePiece As Integer
    Dim Zone As String
    Dim CellQteSSE As String
    
    For i = 1 To UBound(ListeSSE, 2)   'On saute la ligne 0 du tableau qui contient l'ensemble général
        LigGrpDeb = ligactiveOrdo + 1
    ' Ecriture du nom du  Sous Ensemble et de la Qté
        objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(7)).Value = ListeSSE(0, i)
        objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(1)).Value = ListeSSE(2, i)
        'Formule de calcul de Qté a commander
        FormuleQtCdeSSE = "=IF(((E" & ligactiveOrdo & "-K" & ligactiveOrdo & ")*" & CellQteAss & ")<0,0,(E" & ligactiveOrdo & "-K" & ligactiveOrdo & ")*" & CellQteAss & ")"
        objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(8)).Formula = FormuleQtCdeSSE
        'Mise forme de la ligne du sous Ensemble
        Zone = "A" & ligactiveOrdo & ":S" & ligactiveOrdo
        objWorkBookOrdo.ActiveSheet.Range(CStr(Zone)).Interior.Color = 16751052
        'Stockage de la cellule de la quantité de SSe a commander (pour ligne piece)
        CellQteSSE = "$E$" & ligactiveOrdo
        ligactiveOrdo = ligactiveOrdo + 1
    ' Ecriture des pieces du sous ensemble
        LignePiece = ListeSSE(1, i) + 2
        Do While Not (objWorkBook.ActiveSheet.cells(LignePiece, 1).Value = "")
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(0)).Value = objWorkBook.ActiveSheet.cells(LignePiece, 3).Value 'Item Nb
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(1)).Value = objWorkBook.ActiveSheet.cells(LignePiece, 1).Value 'Qte
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(2)).Value = objWorkBook.ActiveSheet.cells(LignePiece, 6).Value 'Designation
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(3)).Value = objWorkBook.ActiveSheet.cells(LignePiece, 5).Value ' SupplierRef
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(4)).Value = objWorkBook.ActiveSheet.cells(LignePiece, 10).Value 'Miscellanous
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(5)).Value = objWorkBook.ActiveSheet.cells(LignePiece, 2).Value 'Sheet
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(9)).Value = TypeElmOrdo(objWorkBook.ActiveSheet.cells(LignePiece, 4).Value, TypeNum) 'Type Ordo
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(10)).Value = objWorkBook.ActiveSheet.cells(LignePiece, 9).Value 'Protect
            'FormuleQteCdePiece = "=SI(((E5*$E$3*$L$4)-K5)<0;0;(E5*$E$3*$L$4)-K5))"
            FormuleQteCdePiece = "=IF(((E" & ligactiveOrdo & "*" & CellQteAss & "*" & CellQteSSE & ")-K" & ligactiveOrdo & ")<0,0,(E" & ligactiveOrdo & "*" & CellQteAss & "*" & CellQteSSE & ")-K" & ligactiveOrdo & ")"
            objWorkBookOrdo.ActiveSheet.cells(ligactiveOrdo, ParamTemplate(8)).Formula = FormuleQteCdePiece
            ligactiveOrdo = ligactiveOrdo + 1
            LignePiece = LignePiece + 1
        Loop
    Barre.Progression = (70 + (30 / SupDivZero(UBound(ListeSSE, 2))) * i)
    'Regroupement des lignes
    objWorkBookOrdo.ActiveSheet.Rows(CStr(LigGrpDeb & ":" & ligactiveOrdo - 1)).Group
    Next
End If

'Numerotation des lignes
'Ajouté pour remettre la nomenclature dans l'ordre en cas de tri non maitrisé
For i = 1 To ligactiveOrdo
    objWorkBookOrdo.ActiveSheet.cells(i, "Z").Value = i
Next


    Barre.Progression = 100
'Fermeture de l'objet eXcel
    objWorkBookOrdo.SaveAs (CStr(CheminDestNomenclature & ActiveDoc.Name & "-Ordo.xls"))
    'objWorkBookOrdo.Close
    objWorkBook.Close
    
MsgBox "Fin de l'export de la nomenclature ordo !", vbInformation, "Traitement terminé"
    
'Libération des classes
Set Barre = Nothing

End Sub
