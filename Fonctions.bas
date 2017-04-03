Attribute VB_Name = "Fonctions"
' *****************************************************************
' * Version 2.5
' * Création CFR le 17/07/2013
' * modification le : 14/01/15
' *    Ajout fonction PtVig
' *    Modif Fonction FormatNoSheet (remplacement du point par une virgule
' *    Modif de la fonction InsertLigneVide (saut de ligne entre les groupes et non les centaines)
' *****************************************************************

Public Function RecupParam(RP_MesParam As Parameters, RP_NomParam As String) As String
'test si le paramètre passé en argument existe dans le part.
'si oui renvoi sa valeur,
'sinon renvoi une chaine vide mais ne crée pas le paramètre
Dim RP_Param As StrParam
On Error Resume Next
    Set RP_Param = RP_MesParam.Item(RP_NomParam)
If (Err.Number <> 0) Then
    ' Le paramètre n'existe pas
    Err.Clear
    RP_Param.Value = ""
End If
RecupParam = RP_Param.Value
End Function

Public Function TestParamExist(TPE_MesParam As Parameters, TPE_NomParam As String) As String
'test si le paramètre passé en argument existe dans le part.
'si oui renvoi sa valeur,
'sinon la crée et lui affecte une chaine vide
Dim TPE_Param As StrParam
On Error Resume Next
    Set TPE_Param = TPE_MesParam.Item(TPE_NomParam)
If (Err.Number <> 0) Then
    ' Le paramètre n'existe pas, on le crée
    Err.Clear
    Set TPE_Param = TPE_MesParam.CreateString(TPE_NomParam, "")
    TPE_Param.Value = ""
End If
TestParamExist = TPE_Param.Value
End Function

Public Function CreateParamExist(CPE_MesParam As Parameters, CPE_NomParam As String, CPE_ValParam As String) As String
'* Modification du 20/09/14
'* Suppression d'un paramètre "CPE_Param" inutile
'test si le paramètre passé en argument existe dans le part.
'si oui lui affecte la valeur CPE_ValParam
'sinon le crée et lui affecte la valeur CPE_ValParam
'CPE_MesParam Collection des paramètres du 3D
'CPE_NomParam Nom du paramètre
'CPE_ValParam valeur du paramètre
Dim CPE_Param As StrParam
On Error Resume Next
    Set CPE_Param = CPE_MesParam.Item(CPE_NomParam)
If (Err.Number <> 0) Then
    ' Le paramètre n'existe pas, on le crée
    Err.Clear
    Set CPE_Param = CPE_MesParam.CreateString(CPE_NomParam, CPE_ValParam)
End If
CPE_Param.Value = CPE_ValParam
CreateParamExist = CPE_Param.Value
End Function

Public Sub CreateParamExist2(CPE_MesParam As Parameters, CPE_NomParam As String, CPE_ValParam As String)
'test si le paramètre passé en argument existe dans le part.
'si oui lui affecte la valeur CPE_ValParam
'sinon le crée et lui affecte la valeur CPE_ValParam
'CPE_MesParam Collection des paramètres du 3D
'CPE_NomParam Nom du paramètre
'CPE_ValParam valeur du paramètre

Dim CPE_Param As StrParam
On Error Resume Next
    Set CPE_Param = CPE_MesParam.Item(CPE_NomParam)
If (Err.Number <> 0) Then
    ' Le paramètre n'existe pas, on le crée
    Err.Clear
    Set CPE_Param = CPE_MesParam.CreateString(CPE_NomParam, CPE_ValParam)
End If
CPE_Param.Value = CPE_ValParam
End Sub

Public Sub AddCompNom(ACN_Attributs() As String, ACN_NbAttribut)
'Ajoute une ligne avec les attributs passés en argument au tableau de nomenclature
ACN_NbLignes = UBound(TableauPartsParam, 2)
For ACN_i = 0 To ACN_NbAttribut
    TableauPartsParam(ACN_i, ACN_NbLignes) = ACN_Attributs(ACN_i)
Next
ReDim Preserve TableauPartsParam(ACN_NbAttribut, ACN_NbLignes + 1)
End Sub

Public Function ListPartUnique(LPU_Nom As String, LPU_ListeParts() As String) As Boolean
'Vérifie si le nom du part fait partie de la liste
Dim LPU_PartExistinList As Boolean
LPU_PartExistinList = False
For lpu_i = 0 To UBound(LPU_ListeParts, 1)
    If LPU_ListeParts(lpu_i) = LPU_Nom Then LPU_PartExistinList = True
Next
ListPartUnique = Not (LPU_PartExistinList)
End Function

Public Function DetectOutilBase(DO_Product As Product) As String
'Détecte l'outil de base et renvois son N°
'l'outil de base est numéroté 000
Dim DO_Products As Products
Set DO_Products = DO_Product.Products
Dim DO_SousProduct As Product
For Each DO_SousProduct In DO_Products
    If Right(DO_SousProduct.PartNumber, 3) = "000" Then
        DetectOutilBase = DO_SousProduct.PartNumber
        Exit For
    Else
        DetectOutilBase = "NON"
    End If
Next
End Function

Public Function DetectCaisse(DC_Product As Product) As String
'Détecte la présence d'une caisse et renvois son N°
'Une caiise est un détaile en 040 à 199 sous la racine de l'outillage général
'Modif du 09/02/2015 Caisse = 040 ald 100
'Modif du 21/12/2015 Caisse = 040 à 799
Dim DC_Products As Products
Set DC_Products = DC_Product.Products
Dim DC_SousProduct As Product
For Each DC_SousProduct In DC_Products
    If Right(DC_SousProduct.PartNumber, 3) >= "040" And Right(DC_SousProduct.PartNumber, 3) <= "799" Then
        DetectCaisse = DC_SousProduct.PartNumber
        Exit For
    Else
        DetectCaisse = "NON"
    End If
Next
End Function

Public Function DetectVariante(DV_Product As Product) As Boolean
'Détecte la présence de variantes dans le product
'Une variante est numéroté en 001 à 039
Dim DV_Products As Products
Set DV_Products = DV_Product.Products
Dim DV_SousProduct As Product
For Each DV_SousProduct In DV_Products
    If Right(DV_SousProduct.PartNumber, 3) > "001" And Right(DV_SousProduct.PartNumber, 3) < "039" Then
        DetectVariante = True
        Exit For
    Else
        DetectVariante = False
    End If
Next
End Function
Public Function TypeElementFlx(TEF_Number As String) As Boolean
'Renvois vrai ou faux en fonction de la présence ou non de "FLXxx" a la fin du numéro du product
'ou du part passé en argument
    TypeElementFlx = False
    If Mid(TEF_Number, Len(TEF_Number) - 4, 3) = "FLX" Then
        TypeElementFlx = True
    End If
End Function

Public Function TypeElementRep(TER_Number As String, TER_TypeNum) As String
'Renvoi une valeur de 0 à 9 en fonction du No de repère de l'élément passé en argument
'0 = Product de tête -> 11 digit
'1 = Product Outillage -> 000
'2 = Variantes de l'outillage -> 001 to 039
'3 = Grand S-ENS -> 040 to 098 Paire
'4 = Grand S-ENS SYM -> 041 to 099 Impaire
'5 = Petit S-ENS Product Mécano-soudé ->  100 to 198 Paire
'6 = Petit S-ENS Product Mécano - soudé SYM -> 101 to 199 Impaire
'7 = Part Fabriqué -> 200 to 498 Paire pout TE_Type = 1
'7 = Part Fabriqué -> 200 to 698 Paire pout TE_Type = 2
'8 = Part Fabriqué SYM -> 201 to 499 Impaire pout TE_Type = 1
'8 = Part Fabriqué SYM -> 201 to 699 Impaire pout TE_Type = 2
'9 = Part Acheté -> 500 to 999 pout TE_Type = 1
'9 = Part Acheté -> 700 to 998 pout TE_Type = 2
'si TER_TypeNum ="" c'est un ancien dossier dans lequel l'attribut n'existait pas.
'Calibrage de plage de Numéro en fonction du Type de Numérotation
Dim Max_Typ7, Max_Typ8, Min_Type9, Max_Typ9 As Integer
If TER_TypeNum = "1" Or TER_TypeNum = "" Then
    Max_Typ7 = 498
    Max_Typ8 = 499
    Min_Type9 = 500
    Max_Typ9 = 999
ElseIf TER_TypeNum = "2" Then
    Max_Typ7 = 698
    Max_Typ8 = 699
    Min_Type9 = 700
    Max_Typ9 = 998
Else
    MsgBox "Erreur de Type de numérotaion !", vbCritical
    End
End If
'Traitement des TER_Number non numérique
 'On renvoi "-1"
    If IsNumeric(TER_Number) = False Then
        TypeElementRep = -1
    Else
        If TER_Number = 11 Then
            TypeElementRep = 0
        ElseIf TER_Number = "000" Then
            TypeElementRep = 1
        ElseIf TER_Number > 0 And CInt(TER_Number) <= 39 Then
            TypeElementRep = 2
        ElseIf (CInt(TER_Number) >= 40 And CInt(TER_Number) <= 98) And CInt(TER_Number) Mod 2 = 0 Then
            TypeElementRep = 3
        ElseIf (CInt(TER_Number) >= 41 And CInt(TER_Number) <= 99) And CInt(TER_Number) Mod 2 = 1 Then
            TypeElementRep = 4
        ElseIf (CInt(TER_Number) >= 100 And CInt(TER_Number) <= 198) And CInt(TER_Number) Mod 2 = 0 Then
            TypeElementRep = 5
        ElseIf (CInt(TER_Number) >= 101 And CInt(TER_Number) <= 199) And CInt(TER_Number) Mod 2 = 1 Then
            TypeElementRep = 6
        ElseIf (CInt(TER_Number) >= 200 And CInt(TER_Number) <= Max_Typ7) And CInt(TER_Number) Mod 2 = 0 Then
            TypeElementRep = 7
        ElseIf (CInt(TER_Number) >= 201 And CInt(TER_Number) <= Max_Typ8) And CInt(TER_Number) Mod 2 = 1 Then
            TypeElementRep = 8
        ElseIf CInt(TER_Number) >= Min_Type9 And CInt(TER_Number) <= Max_Typ9 Then
            TypeElementRep = 9
        Else
            TypeElementRep = -1
            MsgBox "type inconnu"
        End If
    End If

End Function
Public Function TypeElement(TE_Number As String, TE_TypeNum As String) As String
'Renvoi une valeur de 0 à 9 en fonction du No de l'élément passé en argument
'0 = Product de tête -> 11 digit
'1 = Product Outillage -> 000
'2 = Variantes de l'outillage -> 001 to 039
'3 = Grand S-ENS -> 040 to 098 Paire
'4 = Grand S-ENS SYM -> 041 to 099 Impaire
'5 = Petit S-ENS Product Mécano-soudé ->  100 to 198 Paire
'6 = Petit S-ENS Product Mécano - soudé SYM -> 101 to 199 Impaire
'7 = Part Fabriqué -> 200 to 498 Paire pout TE_Type = 1
'7 = Part Fabriqué -> 200 to 698 Paire pout TE_Type = 2
'8 = Part Fabriqué SYM -> 201 to 499 Impaire pout TE_Type = 1
'8 = Part Fabriqué SYM -> 201 to 699 Impaire pout TE_Type = 2
'9 = Part Acheté -> 500 to 999 pout TE_Type = 1
'9 = Part Acheté -> 700 to 998 pout TE_Type = 2
'si TE_TypeNum ="" c'est un ancien dossier dans lequel l'attribut n'existait pas.

'Calibrage de plage de Numéro en fonction du Type de Numérotation
Dim Max_Typ7, Max_Typ8, Min_Type9, Max_Typ9 As Integer
If TE_TypeNum = "1" Or TE_TypeNum = "" Then
    Max_Typ7 = 498
    Max_Typ8 = 499
    Min_Type9 = 500
    Max_Typ9 = 999
Else: TE_TypeNum = "2"
    Max_Typ7 = 698
    Max_Typ8 = 699
    Min_Type9 = 700
    Max_Typ9 = 998
End If

On Error Resume Next
    'Traitement des Flex. on enlève les 5 derniers carratères
    If TypeElementFlx(TE_Number) Then
        TE_Number = Left(TE_Number, Len(TE_Number) - 5)
    End If
    'Traitement des TE_number non numérique
    'On renvoi "3" (Grand S-Ens)
    TE_temp = CInt(Right(TE_Number, 3))
    If Err.Number <> 0 Then
        Err.Clear
        TypeElement = 3
    Else
        If Len(TE_Number) = 11 Then
            TypeElement = 0
        ElseIf Right(TE_Number, 3) = "000" Then
            TypeElement = 1
        ElseIf CInt(Right(TE_Number, 3)) > 0 And CInt(Right(TE_Number, 3)) <= 39 Then
            TypeElement = 2
        ElseIf (CInt(Right(TE_Number, 3)) >= 40 And CInt(Right(TE_Number, 3)) <= 98) And CInt(Right(TE_Number, 3)) Mod 2 = 0 Then
            TypeElement = 3
        ElseIf (CInt(Right(TE_Number, 3)) >= 41 And CInt(Right(TE_Number, 3)) <= 99) And CInt(Right(TE_Number, 3)) Mod 2 = 1 Then
            TypeElement = 4
        ElseIf (CInt(Right(TE_Number, 3)) >= 100 And CInt(Right(TE_Number, 3)) <= 198) And CInt(Right(TE_Number, 3)) Mod 2 = 0 Then
            TypeElement = 5
        ElseIf (CInt(Right(TE_Number, 3)) >= 101 And CInt(Right(TE_Number, 3)) <= 199) And CInt(Right(TE_Number, 3)) Mod 2 = 1 Then
            TypeElement = 6
        ElseIf (CInt(Right(TE_Number, 3)) >= 200 And CInt(Right(TE_Number, 3)) <= Max_Typ7) And CInt(Right(TE_Number, 3)) Mod 2 = 0 Then
            TypeElement = 7
        ElseIf (CInt(Right(TE_Number, 3)) >= 201 And CInt(Right(TE_Number, 3)) <= Max_Typ8) And CInt(Right(TE_Number, 3)) Mod 2 = 1 Then
            TypeElement = 8
        ElseIf CInt(Right(TE_Number, 3)) >= Min_Type9 And CInt(Right(TE_Number, 3)) <= Max_Typ9 Then
            TypeElement = 9
        Else
            TypeElement = -1
            MsgBox "type inconnu"
        End If
    End If
    TypeElement = TypeElement & TE_Flex
End Function

Public Function TypeElmOrdo(NoElm As String, TNum As String) As String
'Renvois le type d'élément pour la nomenclature Ordo
'Assemblage, STD, FAB en fonction du type de numéro
'0 = Product de tête -> 11 digit                                ""
'1 = Product Outillage -> 000                                   "assemblage"
'2 = Variantes de l'outillage -> 001 to 039                     "assemblage"
'3 = Grand S-ENS -> 040 to 098 Paire                            "assemblage"
'4 = Grand S-ENS SYM -> 041 to 099 Impaire                      "assemblage"
'5 = Petit S-ENS Product Mécano-soudé ->  100 to 198 Paire      "assemblage"
'6 = Petit S-ENS Product Mécano - soudé SYM -> 101 to 199 Impaire"assemblage"
'7 = Part Fabriqué -> 200 to 498 Paire pout TE_Type = 1         "FAB"
'7 = Part Fabriqué -> 200 to 698 Paire pout TE_Type = 2         "FAB"
'8 = Part Fabriqué SYM -> 201 to 499 Impaire pout TE_Type = 1   "FAB"
'8 = Part Fabriqué SYM -> 201 to 699 Impaire pout TE_Type = 2   "FAB"
'9 = Part Acheté -> 500 to 999 pout TE_Type = 1                 "STD"
'9 = Part Acheté -> 700 to 998 pout TE_Type = 2                 "STD"
TypeElmOrdo = ""
Select Case TypeElement(NoElm, TNum)
    Case 0
        TypeElmOrdo = ""
    Case 1 To 6
        TypeElmOrdo = "assemblage"
    Case 7 To 8
        TypeElmOrdo = "FAB"
    Case 9
        TypeElmOrdo = "STD"
End Select

End Function

Public Function ListeVariante(LV_Product As Product) As String()
'Renvoi la liste des variantes dans le product
'Une variante est numéroté en 001 à 009
Dim LV_Products As Products
Set LV_Products = LV_Product.Products
Dim LV_SousProduct As Product
Dim LV_ListeVariantes() As String
i = 0
For Each LV_SousProduct In LV_Products
    If Right(LV_SousProduct.PartNumber, 3) > "000" And Right(LV_SousProduct.PartNumber, 3) < "099" Then
        ReDim Preserve LV_ListeVariantes(i)
        LV_ListeVariantes(i) = LV_SousProduct.PartNumber
    i = i + 1
    End If
Next
ListeVariante = LV_ListeVariantes()
End Function

Public Function ListPartProductOpen(LPPO_Type As Integer) As String()
'Constitue une liste des fichiers ouverts en ne gardant que les parts ou les products
'selon le type "2 pour Part" ou "1 pour Product" passé en argument
'Si le type "2" est choisi, on élimine les parts se terminant par 5xx, 6xx, 7xx, 8xx, 9xx
'Ou par 7xx, 8xx, 9xx en fonction de TypeNum_c, qui correspondent à des éléments du commerce.
    Set LPPO_FichiersOuverts = CATIA.Documents
    Dim LPPO_Nom As String
    Dim LPPO_Liste() As String
    i = 1
    j = 0
    ReDim LPPO_Liste(0)
    If TypeNum_c = "1" Then
        NoMaxPart = "499"
    ElseIf TypeNum_c = "2" Then
        NoMaxPart = "699"
    Else
        NoMaxPart = "499"
    End If
    For i = 1 To LPPO_FichiersOuverts.Count
    LPPO_Nom = LPPO_FichiersOuverts.Item(i).Name
    If InStr(1, LPPO_Nom, ".cgr") > 0 Then
        LPPO_Nom = Left(LPPO_Nom, InStr(1, LPPO_Nom, ".CATPart")) & "CATPart"
    End If
    
        If LPPO_Type = 2 And Right(LPPO_Nom, 7) = "CATPart" And Mid(LPPO_Nom, (Len(LPPO_Nom) - 10), 3) < NoMaxPart Then
            ReDim Preserve LPPO_Liste(j)
            LPPO_Liste(j) = LPPO_Nom
            j = j + 1
        ElseIf LPPO_Type = 1 And Right(LPPO_Nom, 10) = "CATProduct" Then
            'Elimination du 11 digits
            If Len(Left(LPPO_Nom, InStr(LPPO_Nom, ".CATProduct"))) > 11 Then
                ReDim Preserve LPPO_Liste(j)
                LPPO_Liste(j) = LPPO_Nom
                j = j + 1
            End If
        End If
    Next
If UBound(LPPO_Liste(), 1) = 0 Then
    ListPartProductOpen = LPPO_Liste()
Else
    ListPartProductOpen = TriList1D(LPPO_Liste(), True)
End If

End Function

Public Function TriList2D(TL_Liste() As String, TL_Col As Integer, TL_Ordre As Boolean) As String()
'Tri par ordre croissant (TL_Ordre=true) ou décroissant(TL_Ordre=false)
'la liste à 2 dimension passée en argument
On Error GoTo Err_TriList2D
Dim i, j, k As Long
    Dim temp As String
    If TL_Ordre Then    '  croissant
        For i = LBound(TL_Liste, 1) + 1 To UBound(TL_Liste, 1) - 1
            For j = i + 1 To UBound(TL_Liste, 1)
                If CInt(TL_Liste(i, TL_Col)) > CInt(TL_Liste(j, TL_Col)) Then
                    For k = 0 To UBound(TL_Liste, 2)
                        temp = TL_Liste(j, k)
                        TL_Liste(j, k) = TL_Liste(i, k)
                        TL_Liste(i, k) = temp
                    Next k
                End If
            Next j
        Next i
    Else            ' décroissant
        For i = LBound(TL_Liste, 1) + 1 To UBound(TL_Liste, 1) - 1
            For j = i + 1 To UBound(TL_Liste, 1)
                If CInt(TL_Liste(i, TL_Col)) < CInt(TL_Liste(j, TL_Col)) Then
                    For k = 0 To UBound(TL_Liste, 2)
                        temp = TL_Liste(j, k)
                        TL_Liste(j, k) = TL_Liste(i, k)
                        TL_Liste(i, k) = temp
                    Next k
                End If
            Next j
        Next i
    End If
Err_TriList2D:
TriList2D = TL_Liste()
End Function

Public Function TriList1D(TL_Liste() As String, TL_Ordre As Boolean) As String()
'Tri par ordre croissant (TL_Ordre=true) ou décroissant(TL_Ordre=false)
'la liste à une dimension passée en argument
Dim i, j, k As Long
    Dim temp As String
    If TL_Ordre Then    '  croissant
        For i = LBound(TL_Liste, 1) + 1 To UBound(TL_Liste, 1) - 1
            For j = i + 1 To UBound(TL_Liste, 1)
                If TL_Liste(i) > TL_Liste(j) Then
                        temp = TL_Liste(j)
                        TL_Liste(j) = TL_Liste(i)
                        TL_Liste(i) = temp
                End If
            Next j
        Next i
    Else            ' décroissant
        For i = LBound(TL_Liste, 1) + 1 To UBound(TL_Liste, 1) - 1
            For j = i + 1 To UBound(TL_Liste, 1)
                If TL_Liste(i) < TL_Liste(j) Then
                        temp = TL_Liste(j)
                        TL_Liste(j) = TL_Liste(i)
                        TL_Liste(i) = temp
                End If
            Next j
        Next i
    End If
TriList1D = TL_Liste()
End Function

Public Function ProgressBar(PBAvancement As Integer)
If PBAvancement > 100 Then PBAvancement = 100
    Frm_Progression.Bar1.Width = PBAvancement * 3

End Function

Public Function SupDivZero(SPZVal As Long)
'Remplace le Zero par un pour éviter les divisions par zéro
If SPZVal = 0 Then
    SupDivZero = 1
Else
    SupDivZero = SPZVal
End If
End Function

Public Function TranspositionTabl(TT_Table() As String) As String()
'Transposition des lignes et des colonnes du tableau
    Dim TT_TableTemp() As String
    ReDim TT_TableTemp(UBound(TT_Table, 2), UBound(TT_Table, 1))
    For i = 0 To UBound(TT_Table, 2)
        For j = 0 To UBound(TT_Table, 1)
            TT_TableTemp(i, j) = TT_Table(j, i)
        Next
    Next
    TranspositionTabl = TT_TableTemp
End Function

Public Function Txt2Digit(T2D_Txt As String) As String
'Renvois les chiffre de 0 à 9 au format 01 à 09
If Len(T2D_Txt) = 1 Then
    Txt2Digit = "0" & T2D_Txt
Else
    Txt2Digit = T2D_Txt
End If

End Function

Public Function Txt3Digit(T3D_Txt As String) As String
'Renvois les chiffre de 0 à 99 au format 001 à 099
If Len(T3D_Txt) = 1 And T3D_Txt <> " " Then
    Txt3Digit = "00" & T3D_Txt
ElseIf Len(T3D_Txt) = 2 Then
    Txt3Digit = "0" & T3D_Txt
Else
    Txt3Digit = T3D_Txt
End If

End Function

Public Function EstVarianteOut(EVO_Digits As String) As Boolean
'Renvois vrai si les 3 digits passés en arguments sont dans la liste des variantes d'outillage
If (CInt(EVO_Digits) >= 1 And CInt(EVO_Digits) <= 39) Then EstVarianteOut = True
End Function

Public Function EstGrdSSe(EGSS_Digits As String) As Boolean
'Renvois vrai si les 3 digits passés en arguments sont dans la liste des Grand Sous ensembles
If (CInt(EGSS_Digits) >= 40 And CInt(EGSS_Digits) <= 99) Then EstGrdSSe = True
End Function

Public Function EstPttSSe(EPSS_Digits As String) As Boolean
'Renvois vrai si les 3 digits passés en arguments sont dans la liste des Petit Sous ensembles
If (CInt(EPSS_Digits) >= 100 And CInt(EPSS_Digits) <= 199) Then EstPttSSe = True
End Function

Public Function EstPieceCom(EPC_Digits As String) As Boolean
'Renvois vrai si les 3 digits passés en arguments sont dans la liste des Pieces du commerce
If (CInt(EPC_Digits) >= 500 And CInt(EPC_Digits) <= 999) Then EstPieceCom = True
End Function

Public Function FormatNoSheet(NoSheet As String, SiteAB As String) As String
'Formate le numéro de planche (ajoute les zéros non significatifs supprimés par excel en début de chaine
'NoSheet = N° de planche au format 1, 10, 100 ou 1,100
'SiteAB = "Francais", "Allemand", "Anglais", "Espagnol"
'Formate le numéro de planche avec des zéros en tète pour avoir 2 Digit pour les GSE français et 3 Digit pour les GSE allemand
'Remplace les point par des virgules
NoSheet = PtVig(NoSheet)
Dim NB_digit As Integer
Select Case SiteAB
    Case "Allemand"
        NB_digit = 3
    Case Else
        NB_digit = 2
End Select
If InStr(1, NoSheet, ",") > 0 Then
    While Len(Left(NoSheet, InStr(1, NoSheet, ","))) < NB_digit
        NoSheet = "0" & NoSheet
    Wend
Else
    While Len(NoSheet) < NB_digit
        NoSheet = "0" & NoSheet
    Wend
End If
FormatNoSheet = NoSheet
End Function

Public Function SymMiscellanous(SM_ItemNB As String) As String
'Renvois "SYM TO " + N° d'item -1 s'il s'agit d'un symétrique
'Modif du 27/11/15 on ne garde que les 3 dernier chiffre
Dim Temp_ItemNB As Integer
Temp_ItemNB = CInt(Right(SM_ItemNB, 3)) - 1
'Temp_ItemNB = CInt(SM_ItemNB) - 1

SymMiscellanous = "SYM TO " & Txt3Digit(CStr(Temp_ItemNB))
End Function

Public Function InsertLigneVide(ILV_Tabl() As String, ILV_TypNum, DecalCol) As String()
'Insert une ligne blanche entre chaque changement de centaine dans les N° de détails
' * modification le : 14/01/15
' *    Modification du saut de ligne pour les GSE allemnad (plus de saut entre 200/500  ou 500/900

Dim LigEC As Integer, SerieEC As Integer, NoRepEC As Integer
Dim Tabl_temp() As String
'Definition dde la centaine a partir de laquelle on arrète de sauter des lignes.
'700 pour le type 2 et 500 pour le type 1
Dim ILVFinSautLigne As Integer
'1ere Dimention du tableau
Dim Dim1Tabl As Integer
Dim1Tabl = UBound(ILV_Tabl, 1)
'Récupération des premières lignes (Lignes des ensembles + ligne blanche)
    For i = 0 To 1 + DecalCol
        ReDim Preserve Tabl_temp(Dim1Tabl, i)
        For j = 0 To Dim1Tabl
            Tabl_temp(j, i) = ILV_Tabl(j, i)
        Next j
    Next i
    SerieEC = NoRepEnt(ILV_Tabl(2 + DecalCol, 2 + DecalCol))
    For LigEC = 2 + DecalCol To UBound(ILV_Tabl, 2)
        NoRepEC = NoRepEnt(ILV_Tabl(2 + DecalCol, LigEC))
        ReDim Preserve Tabl_temp(Dim1Tabl, UBound(Tabl_temp, 2) + 1)
 
        If (SerieEC <= 99 And NoRepEC >= 100) Or _
            (SerieEC <= 199 And NoRepEC >= 200) Or _
            (ILV_TypNum = 1 And SerieEC <= 499 And NoRepEC >= 500) Or _
            (ILV_TypNum = 2 And SerieEC <= 699 And NoRepEC >= 700) Then
            'Ajout ligne vide
            For i = 0 To UBound(Tabl_temp, 1)
                Tabl_temp(i, UBound(Tabl_temp, 2)) = ""
            Next i
            'récupération du numero de rep pour comparaison avec le prochain
            SerieEC = NoRepEnt(ILV_Tabl(2 + DecalCol, LigEC))
            ReDim Preserve Tabl_temp(Dim1Tabl, UBound(Tabl_temp, 2) + 1)
        End If
        'Ajout Info ligne en cours
        For i = 0 To UBound(Tabl_temp, 1)
            Tabl_temp(i, UBound(Tabl_temp, 2)) = ILV_Tabl(i, LigEC)
        Next i
    Next
InsertLigneVide = Tabl_temp()
End Function

Public Function NoRepEnt(NREC_Num) As Integer
'Renvois sous forme d'entier le Numero du repere passé en string
'catch les erreurs
On Error Resume Next
        NoRepEnt = CInt(NREC_Num)
        If (Err.Number <> 0) Then
        ' Le Champs Numéro de repère est vide ou non numérique. on le met à Zéro
            Err.Clear
            NoRepEnt = 0
        End If
End Function

Public Function SupSautLigne(SSL_chaine As String) As String
'Remplace les saut de ligne 'Chr(10)Chr(13) de la chaine de carractère par un espace
Dim i As Long
Dim Temp_chaine As String
    
    'debug.print SSL_chaine
    For i = 1 To Len(SSL_chaine)
        If Mid(SSL_chaine, i, 1) <> Chr(10) And Mid(SSL_chaine, i, 1) <> Chr(13) Then
            Temp_chaine = Temp_chaine & Mid(SSL_chaine, i, 1)
        Else
            Temp_chaine = Temp_chaine & " "
        End If

    Next i
    SupSautLigne = Temp_chaine
    'debug.print SupSautLigne
End Function

Public Function SautLigne(SL_chaine As String) As String
'Remplace le carratère '¤' par un saut de ligne
    SautLigne = Replace(SL_chaine, "¤", Chr(10) & Chr(13))
End Function

Public Function SuprSautLigne(SL_chaine As String) As String
'Remplace le carratère '¤' par un espace
    SuprSautLigne = Replace(SL_chaine, "¤", " ")
End Function

Public Function PtVig(PV_Chaine As String) As String
'Remplace les chaine avec un point par une chaine avec une virgule
'cas des valeur interprétée comme des valeur numérique par excel
    PtVig = Replace(PV_Chaine, ".", ",")
End Function

Public Function VigPt(VP_Chaine As String) As String
'Remplace les chaine avec une vigule par une chaine avec un point
'cas des valeur interprétée comme des valeur numérique par excel
    VigPt = Replace(VP_Chaine, ",", ".")
End Function

Public Function AjoutQuote(AQ_Chaine As String) As String
'Vérifie si le premier carractère de la cahine est une quote.
' Si oui, renvois la chaine
'Si non ajoute une quote en début de chaine
    If Left(AQ_Chaine, 1) = "'" Then
        AjoutQuote = AQ_Chaine
    Else
        AjoutQuote = "'" & AQ_Chaine
    End If
End Function

Public Function NoSheetSupplier(NoItem As String, NoSheetDet As String, NoSheetSup As String, TypNum As String) As String
'Si l'item est un Supplier Ref, (500 à 999 pour typenum = 1 et 700 à 999 pour typenum = 2)  on renvoi le No de la planche en cours (NoSheetSup)
'le composant peut apparaitre sur plusieur planches
'Sinon on renvoi le Numero de la planche de détail (NoSheetDet)
' NoItem = "xxx"
' NoSheetDet = "xx"
' NoSheetSup = "xx"
' TypNum = "1" ou "2"
Dim SeuilInf As Integer
If TypNum = 1 Then SeuilInf = 5 Else SeuilInf = 7
If Left(NoItem, 1) >= SeuilInf And Left(NoItem, 1) < 10 Then
    NoSheetSupplier = NoSheetSup
Else
    NoSheetSupplier = NoSheetDet
End If
End Function

Public Function EffaceFicNom(EF_Folder, EF_FicNom As String) As Boolean
'Effacement d'un fichier de nomenclature pré-existant
 On Error GoTo Err_EffaceFicNom
    Dim EF_FS, EF_Fold, EF_Files, EF_File
    Set EF_FS = CreateObject("Scripting.FileSystemObject")
    Set EF_Fold = EF_FS.GetFolder(EF_Folder)
    Set EF_Files = EF_Fold.Files
    For Each EF_File In EF_Files
        If EF_File.Name = EF_FicNom Then
            EF_FS.DeleteFile (CStr(EF_Folder & "\" & EF_FicNom))
        End If
    Next
    EffaceFicNom = True
GoTo Quit_EffaceFicNom

Err_EffaceFicNom:
MsgBox "Il est possible que le fichier de nomenclature soit encore ouvert dans Excel. Veuillez le fermer et relancer la macro.", vbCritical, "erreur"
EffaceFicNom = False
Quit_EffaceFicNom:
End Function

Public Sub LargCol(Wsheet, lCol1, lCol2, Lc_Larg As Integer)
'Change la largeur des colonne de la feuille excel
    With Wsheet.Range(lCol1 & 1 & ":" & lCol2 & 1)
        .ColumnWidth = Lc_Larg
    End With
End Sub

Public Sub CouleurCell(CC_Worksheet, CC_Cell1, CC_Cell2, CC_Coul As String)
'Colorie la plage de cellules CC_Cell1:CC_Cell2 dans la couleur passée en argument
'CC_Cell1 = Lettre ou chiffre (Colonne)
'Si cc_Cell1 = chiffre, on le transforme en lettre
'CC_Cell2 = Chiffre (Ligne)
Dim CC_Color As Long
Dim LettreCell As String

If CC_Coul = "gris" Then
    CC_Color = 14277081 'Gris
ElseIf CC_Coul = "jaune" Then
    CC_Color = 49407 'Jaune
ElseIf CC_Coul = "vert" Then
    CC_Color = 5296274 'Vert
ElseIf CC_Coul = "rouge" Then
    CC_Color = 255
ElseIf CC_Coul = "bleu" Then
    CC_Color = 15917714
Else
    CC_Color = 0
End If

If IsNumeric(CC_Cell1) Then
    LettreCell = NumCar(CInt(CC_Cell1))
    'LettreCell = Choose(CC_Cell1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ")
Else
    LettreCell = CC_Cell1
End If
Dim celtemp1 As String, celtemp2 As String
celtemp1 = CStr(LettreCell & CC_Cell2)
celtemp2 = CStr(LettreCell & CC_Cell2)
'With CC_Worksheet.range(celtemp1, celtemp2).Interior
With CC_Worksheet.Range(LettreCell & CC_Cell2).Interior
    '.Pattern = xlSolid
    '.PatternColorIndex = xlAutomatic
    .Color = CC_Color
End With
End Sub

'#####################

Public Sub Condition_Sup0(Wsheet, Col, Lig)
'Format conditionel des cellues.
'Les valeurs suppérieurs à 0 en rouge
Dim Cell As String
    Cell = Col & Lig
    With Wsheet.Range(Cell)
        .FormatConditions.Add Type:=xLCellValue, Operator:=xLGreater, Formula1:="=0"
        .FormatConditions(1).SetFirstPriority
        .FormatConditions(1).Interior.Color = 255
        .FormatConditions(1).PatternColorIndex = xlAutomatic
        .FormatConditions(1).TintAndShade = 0
        .FormatConditions(1).StopIfTrue = False
    End With
    
End Sub

Public Sub Condition_Doublon(Wsheet, Cell1, Cell2)
'Format conditionel des cellues.
'Les doublon en vert et les ref unique en rouge
Dim CD_Cell As String
    CD_Cell = Cell1 & ":" & Cell2

    With Wsheet.Range(CD_Cell)
        .FormatConditions.Delete
        .FormatConditions.AddUniqueValues
        .FormatConditions(1).SetFirstPriority
        .FormatConditions(1).DupeUnique = xlUnique
        .FormatConditions(1).Interior.Color = 255
        .FormatConditions(1).StopIfTrue = False
'        .FormatConditions.AddUniqueValues
        '.FormatConditions(2).SetFirstPriority
'        .FormatConditions(2).DupeUnique = xlDuplicate
'        .FormatConditions(2).Interior.Color = 5296274
'        .FormatConditions(2).StopIfTrue = False
    End With
End Sub

Public Sub Condition_Vide(Wsheet, Cell1, Cell2)
'Format conditionel des cellues.
'Passe en rouge les cellules contenant "Absent" et en orange celle contenant "Vide"
Dim CD_Cell As String
    CD_Cell = Cell1 & ":" & Cell2
    Dim Formule As String
    Formule = "=""Vide"""
    With Wsheet.Range(CD_Cell)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xLTextString, String:="Absent", TextOperator:=xlContains
        .FormatConditions.Add Type:=xLTextString, String:="Vide", TextOperator:=xlContains
        .FormatConditions(1).SetFirstPriority
        .FormatConditions(1).Interior.Color = 255
        .FormatConditions(2).Interior.Color = 49407
        .FormatConditions(1).StopIfTrue = False
    
    End With
End Sub

Public Function XlCol(NoCol) As String
    'Converti les numero de colonne en lettre
    XlCol = Choose(NoCol, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    End Function

Public Sub HauteurCell(Wsheet, Lig)
    'Change la hauteur de ligne
        With Wsheet.Rows(Lig & ":" & Lig).Select
            .RowHeight = 6
        End With
    End Sub

Public Sub Mef(WS As Variant, Zone As String, style As String)
If style = "BG" Then 'Bleu / 14 /Gras
    With WS.Range(Zone)
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = 15917714
        .Rows("1:1").RowHeight = 30
    End With
ElseIf style = "BI" Then 'Bleu / 12 /NonGras
    With WS.Range(Zone)
        .Font.Size = 12
        .Font.Bold = False
        .Interior.Color = 15917714
        .Rows("1:1").RowHeight = 30
    End With
ElseIf style = "BM" Then 'Bleu / 11 /NonGras
    With WS.Range(Zone)
        .Font.Size = 11
        .Font.Bold = False
        .Interior.Color = 15917714
    End With
ElseIf style = "GM" Then 'Gris / 11 /NonGras
    With WS.Range(Zone)
        .Font.Size = 11
        .Font.Bold = False
        .Interior.Color = 14277081
    End With
End If

End Sub

Public Sub ColorCell(Wsheet, Col, Lig)
'Mise en forme conditionnelle.
    With Wsheet.Range(Col & Lig)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xLTextString, String:="KO", TextOperator:=xlContains
        .FormatConditions.Add Type:=xLTextString, String:="OK", TextOperator:=xlContains
        .FormatConditions(1).Font.ColorIndex = 3
        .FormatConditions(2).Font.ColorIndex = 4
    End With
End Sub

Public Sub AjoutComment(Wsheet, Col, Lig, AC_comment)
'Ajoute un commentaire à la cellule Col & Lig de la feuille excel Wsheet
    With Wsheet.Range(Col & Lig)
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:=CStr(AC_comment)
    End With
End Sub

Public Function ReplaceBlanck(RB_String As String) As Integer
'Renvoi une convertion en entier de la chaine passée en argument
'Si ce n'est pas possible, renvois 0
On Error Resume Next
ReplaceBlanck = CInt(RB_String)
If Err.Number <> 0 Then
    Err.Clear
    ReplaceBlanck = 0
    On Error GoTo 0
End If
End Function

Public Function ExtractNameBody(str)
'Extrait le nom du body dans le nom du paramètre passé en argument
'le nom du paramètre a cette forme 'Part.1\Body1\Material'
Dim PosSlash As String

    PosSlash = InStr(1, str, "\", vbTextCompare)
    str = Right(str, Len(str) - PosSlash)
    PosSlash = InStr(1, str, "\", vbTextCompare)
    str = Left(str, PosSlash - 1)
    ExtractNameBody = str

End Function

Public Function ExtractValNum(str)
'Extrai la valeur numérique compride dans une chaine
'ex : renvoi 65 pour la chaine "65Kg"
Dim i As Integer
Dim TempStr As String, OneChar As String
On Error Resume Next
For i = 1 To Len(str)
     OneChar = CInt(Mid(str, i, 1))
    If Err.Number <> 0 Then
        Err.Clear
    Else
        TempStr = TempStr & OneChar
    End If
    
Next i
ExtractValNum = CInt(TempStr)
End Function

Public Function ListPart(ByVal LP_Liste)
'transforme un texte avec des séparateur ";" en liste de textes
Dim Temp_List() As String
Dim i As Integer
If InStr(1, LP_Liste, ";", vbTextCompare) > 0 Then
    While InStr(1, LP_Liste, ";", vbTextCompare) > 0
        ReDim Preserve Temp_List(i)
        Temp_List(i) = Left(LP_Liste, InStr(1, LP_Liste, ";", vbTextCompare) - 1)
        LP_Liste = Right(LP_Liste, Len(LP_Liste) - InStr(1, LP_Liste, ";", vbTextCompare))
        i = i + 1
    Wend
    ReDim Preserve Temp_List(i)
    Temp_List(i) = LP_Liste
Else
    ReDim Temp_List(0)
    Temp_List(0) = LP_Liste
End If
ListPart = Temp_List

End Function

Public Function CageCode(Wsheet, Col As Integer, CC_TypeNum As String) As String()
'Recupère la liste des fourniseur de la nomenclature
'pour chaque fournisseur cherche le Cage Code correspondant dans une table excel
'Wsheet  = Objet Excel contenant la nomenclature
' Col = N° de colonne contenant les noms des fournisseur

'Creation d'un objet eXcel et ouverture de la base des Cages Codes
    Dim objexcel
    Dim objWorkbookCode
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkbookCode = objexcel.Workbooks.Open(CStr(CheminSourcesMacro & Nom_FicCageCodes))
    'objExcel.Visible = True
    'objWorkbookCode.ActiveSheet.Visible = True
    
Dim ListCageCode() As String
Dim ListFourn() As String
Dim i As Long, j As Long

Dim LigActive As Long, LigActiveTab As Long
    LigActive = 0
    
'Construit la liste des Cages codes
    While objWorkbookCode.ActiveSheet.cells(LigActive + 1, 1).Value <> ""
        LigActive = LigActive + 1
        ReDim Preserve ListCageCode(1, LigActive)
        ListCageCode(0, LigActive) = objWorkbookCode.ActiveSheet.cells(LigActive, 1).Value
        If objWorkbookCode.ActiveSheet.cells(LigActive, 2).Value = "" Then
            ListCageCode(1, LigActive) = "XXXXXXXXXXXXXX"
        Else
            ListCageCode(1, LigActive) = objWorkbookCode.ActiveSheet.cells(LigActive, 2).Value
        End If
        
    Wend
    objWorkbookCode.Close
    
'Construit la liste des fournisseurs
    LigActive = 0
    LigActiveTab = 0
    'recherche du récapitulatif des pièces
    While Left(Wsheet.cells(LigActive + 1, 1).Value, 5) <> "Total"
        LigActive = LigActive + 1
    Wend
    LigActive = LigActive + 4 'saut des lignes d'entète
    ReDim ListFourn(1, 0)
        ListFourn(0, 0) = " "
        ListFourn(1, 0) = " "
    While Wsheet.cells(LigActive, 1).Value <> ""
        LigActive = LigActive + 1
        'que pour les élements achetés
        If TypeElement(Wsheet.cells(LigActive, 4).Value, CC_TypeNum) = 9 Then
            'elimination des doublon
            If Not ElimineDbl(ListFourn, Wsheet.cells(LigActive, Col).Value) Then
                ReDim Preserve ListFourn(1, LigActiveTab)
                'Ajout a la liste
                ListFourn(0, LigActiveTab) = Wsheet.cells(LigActive, Col).Value
                LigActiveTab = LigActiveTab + 1
            End If
        End If
    Wend

'Ajout des Cage Codes a la liste des fournisseurs
    For i = 0 To UBound(ListFourn, 2)
        For j = 0 To UBound(ListCageCode, 2)
            If ListFourn(0, i) = ListCageCode(0, j) Then
                ListFourn(1, i) = ListCageCode(1, j)
                Exit For
            End If
        Next j
    Next i
    
CageCode = ListFourn
End Function

Public Function ElimineDbl(TablVal, Valeur) As Boolean
'test si "Valeur" est déja présente dans TablVal
Dim i As Long
ElimineDbl = False
    For i = 0 To UBound(TablVal, 2)
        If TablVal(0, i) = Valeur Then
            ElimineDbl = True
            Exit For
         End If
    Next i
End Function

Public Function NoDerniereLigne(NDL_TablExcel As Variant) As Integer
'recherche la dernière ligne du fichier excel
'On part du principe que 2 lignes vide indiquent la fin du fichier
Dim NoLigne As Integer, Nb_Lig_Vide As Integer
    NoLigne = 1
    Nb_Lig_Vide = 0
    While Nb_Lig_Vide < 2
        If NDL_TablExcel.ActiveSheet.cells(NoLigne, 1).Value = "" Then
            Nb_Lig_Vide = Nb_Lig_Vide + 1
        Else
            Nb_Lig_Vide = 0
        End If
    NoLigne = NoLigne + 1
    Wend
    NoDerniereLigne = NoLigne - 2
End Function

Public Function NoDebRecap(NDR_TablExcel As Variant, langue As String) As Integer
'recherche la 1ere ligne du récapitulatif des pièces
' la ligne commence par "Nomenclature de" ou "Recapitulation of:"
    Dim NomSeparateur As String
    Dim NoLigne As Integer
    NoLigne = 1
    If langue = "EN" Then
        NomSeparateur = "Recapitulation of:"
    ElseIf lange = "FR" Then
        NomSeparateur = "Récapitulatif sur"
    End If
    While Left(NDR_TablExcel.ActiveSheet.cells(NoLigne, 1).Value, Len(NomSeparateur)) <> NomSeparateur
        NoLigne = NoLigne + 1
    Wend
    NoDebRecap = NoLigne
End Function

Public Function TestEstSSE(Ligne As String, langue As String)
'test si la ligne correspond a une entète de sous ensemble
' la ligne commence par "Nomenclature de" ou "Bill of Material"
    Dim NomSeparateur As String
    If langue = "EN" Then
        NomSeparateur = "Bill of Material: "
    ElseIf lange = "FR" Then
        NomSeparateur = "Nomenclature de "
    End If
    On Error Resume Next 'Test si la chaine est vide ou inférieur a len(nomséparateur)
    tmpNomSSE = Left(Ligne, Len(NomSeparateur))
    If Err.Number <> 0 Then
         TestEstSSE = False
    Else
        If Left(Ligne, Len(NomSeparateur)) = NomSeparateur Then
            TestEstSSE = True
        Else
            TestEstSSE = False
        End If
    End If
End Function

Public Function NomSSE(NSSE_Ligne As String, langue As String)
'Récupère le nom du sous ensemble passé en argument
'par suppression des 16 premiers carratères(Nomenclature de )" pour le français
'et des 18 premiers carractères (Bill of Material: ) pour l'anglais
    If langue = "EN" Then
        NomSeparateur = "Bill of Material: "
    ElseIf langue = "FR" Then
        NomSeparateur = "Nomenclature de "
    End If
    On Error Resume Next  'Test si la chaine est vide ou inférieur a len(nomséparateur)
    NomSSE = Right(NSSE_Ligne, Len(NSSE_Ligne) - Len(NomSeparateur))
    If Err.Number <> 0 Then
        NomSSE = ""
    End If
End Function

Public Function NomMachine(NM_TablExcel As Variant, NM_NoDerniereLigne As Long, langue As String)
'Récupère le nom de la machine (Nom de l'ensemble général)
'Sur la ligne du récapitulatif
    Dim NomSeparateur As String
    If langue = "EN" Then
        NomSeparateur = "Recapitulation of: "
    ElseIf lange = "FR" Then
        NomSeparateur = "Récapitulatif sur "
    End If
Dim NM_NoLigne As Integer
    For NM_NoLigne = NM_NoDerniereLigne To 1 Step -1
        If Left(NM_TablExcel.ActiveSheet.cells(NM_NoLigne, 1).Value, Len(NomSeparateur)) = NomSeparateur Then
            NomMachine = Right(NM_TablExcel.ActiveSheet.cells(NM_NoLigne, 1).Value, Len(NM_TablExcel.ActiveSheet.cells(NM_NoLigne, 1).Value) - Len(NomSeparateur))
            Exit Function
        End If
    Next
End Function

Public Function CheckProduct(ActiveDoc) As Boolean
'Test si le document actif est bien un Catproduct
Dim CkProd As Boolean
    CkProd = True
Dim ProdDoc As ProductDocument

    On Error Resume Next
    Set ProdDoc = ActiveDoc
    If Err.Number <> 0 Then
        CkProd = False
    End If
    On Error GoTo 0
    CheckProduct = CkProd
End Function

Public Function EstPart(Obj) As Boolean
'test si le produit passé en argument est un Part
EstPart = False
Dim Prt As Product
    On Error Resume Next
    Err.Clear
    Set Prt = Obj.Product
    If Not (Err.Number <> 0) Then
        On Error GoTo 0
        If TypeName(Prt) = "part" Then
            EstPart = True
        End If
    End If
End Function

Public Function EstProduct(Obj) As Boolean
'test si le produit passé en argument est un Part
EstProduct = False
Dim Prod As Product
    On Error Resume Next
    Err.Clear
    Set Prod = Obj.Product
    If Not (Err.Number <> 0) Then
        On Error GoTo 0
        If TypeName(Prod) = "Product" Then
            EstProduct = True
        End If
    End If
End Function

Public Function NumCar(Num As Integer) As String
'Converti un chiffre en lettre
'1 = A, 2 = B etc
'Attention la numérotation de Array commence à 0 d'ou le double A dans la liste
Dim ListCar
If Num > 78 Then ' a changer si on ajoute des colonnes a la liste Array
    Num = 1
End If
ListCar = Array("A", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ")
NumCar = ListCar(Num)
End Function

Public Sub LogUtilMacro(ByVal mPath As String, ByVal mFic As String, ByVal mMacro As String, ByVal mModule As String, ByVal mVersion As String)
'Log l'utilisation de la macro
'Ecrit une ligne dans un fichier de log sur le serveur
'mPath = localisation du fichier de log ("\\serveur\partage")
'mFic = Nom du fichier de log ("logUtilMacro.txt")
'mMacro = nom de la macro ("NomGSE")
'mVersion = Version de la macro ("version 9.1.4")
'mModule = Nom du module ("_Info_Outillage")

Dim mDate As String
Dim mUser As String
Dim nFicLog As String
Dim LigLog As String
Const ForWriting = 2, ForAppending = 8

    mDate = Date & " " & Time()
    mUser = ReturnUserName()
    nFicLog = mPath & "\" & mFic

    nliglog = mDate & ";" & mUser & ";" & mMacro & ";" & mModule & ";" & mVersion

    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fs.GetFile(nFicLog)
    If Err.Number <> 0 Then
        Set f = fs.opentextfile(nFicLog, ForWriting, 1)
    Else
        Set f = fs.opentextfile(nFicLog, ForAppending, 1)
    End If
    
    f.Writeline nliglog
    f.Close
    On Error GoTo 0

End Sub

Function ReturnUserName() As String 'extrait d'un code de Paul, Dave Peterson Exelabo
'Renvoi le user name de l'utilisateur de la station
'fonctionne avec la fonction GetUserName dans l'entète de déclaration
    Dim Buffer As String * 256
    Dim BuffLen As Long
    BuffLen = 256
    If GetUserName(Buffer, BuffLen) Then _
    ReturnUserName = Left(Buffer, BuffLen - 1)
End Function

