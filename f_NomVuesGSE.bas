Attribute VB_Name = "f_NomVuesGSE"
Option Explicit

Sub catmain()
'Macro de renommage des vue pour liasses GSE Colomiers
'modification du 10/02/11 ==> suppression des textes en Français et changement taille texte (8 ald 7)
'modification du 15/02.01 ==> suppression soulignement et italique. ajout "Section" aux sections
'                             Taille echelles 5 ald 3.5
'modification de 25/11/15 ==> Ajout d'un formulaire avec la liste des echelles
'modification du 18/05/17 ==> Mise à jour due l'échelle dans le cartouche

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "f_NomVuesGSE", VMacro

Dim DocActif As DrawingDocument
Dim Col_Calques As DrawingSheets
Dim CalqueActif As DrawingSheet
Dim NomCalqueActif As String
Dim Col_Vues As DrawingViews
Dim EchelleCalqueActif As Double
Dim VueActive As DrawingView, BackVue As DrawingView
Dim NomVueActive As String, NewViewName As String
Dim ViewName As String, ViewIdent As String, VewSuffix As String
Dim NewTreeViewName As String, NewTreeViewIdent As String, NewTreeVewSuffix As String
Dim Col_Textes As DrawingTexts
Dim TxtNomVue As DrawingText
Dim iScaleVuePrincipale As Double
Dim i As Long, j As Long
Dim Deb As Integer, Fin As Integer
Dim FormatNomVue As Boolean, FormatScaleVue As Boolean, PasDeTexte As Boolean
Dim NotaUnfold_Y As Double, NotaUnfold_X As Double
Dim Notatxt As String
Dim TxtNota As DrawingText
'Nom des type de vues en fonction de la langue
Dim ANameVueF As String, ANameVueD As String, ANameVueG As String, ANameVueH As String, ANameVueB As String
Dim ANameSection As String, ANameCoupe As String, ANameDetail As String, ANameVueAux As String, ANameVueIso As String
Dim ANameVueDep As String
Dim FNameVueF As String, FNameVueD As String, FNameVueG As String, FNameVueH As String, FNameVueB As String
Dim FNameSection As String, FNameCoupe As String, FNameDetail As String, FNameVueAux As String, FNameVueIso As String
Dim FNameVueDep As String

    'Initialisation des variables
    Set DocActif = CATIA.ActiveDocument
    Set Col_Calques = DocActif.Sheets
    Set CalqueActif = Col_Calques.ActiveSheet
    NomCalqueActif = CalqueActif.Name
    Set Col_Vues = CalqueActif.Views
    EchelleCalqueActif = CalqueActif.Scale

    Set BackVue = Col_Vues.Item("Background View") 'Calque de fond

    'Facteur d'échelle de la vue principale
    Load frm_Echelles
    frm_Echelles.Show
    iScaleVuePrincipale = ConvertScale(frm_Echelles.CBL_Echelles.Value)
    'iScaleVuePrincipale = InputBox("Qu'elle est l'échelle de la vue principale ? (ex: 1; 2, 0.5) :", "Echelle vue principale")
    Unload frm_Echelles

'Anglais
    ANameVueF = "Front view"
    ANameVueD = "Right view"
    ANameVueG = "Left view"
    ANameVueH = "Top view"
    ANameVueB = "Bottom view"
    ANameSection = "Section cut"
    ANameCoupe = "Section view"
    ANameVueAux = "Auxiliary view"
    ANameDetail = "Detail"
    ANameVueIso = "Isometric view"
    ANameVueDep = "Unfolded view"
    
'Français
    FNameVueF = "Vue de face"
    FNameVueD = "Vue de droite"
    FNameVueG = "Vue de gauche"
    FNameVueH = "Vue de dessus"
    FNameVueB = "Vue de dessous"
    FNameSection = "Section"
    FNameCoupe = "Coupe"
    FNameVueAux = "Vue auxiliaire"
    FNameDetail = "Détail"
    FNameVueIso = "Vue isométrique"
    FNameVueDep = "Vue dépliée"

For i = 3 To Col_Vues.Count 'On elimine les 2 premières vues (main view et background view)
    Set VueActive = Col_Vues.Item(i)
    Set Col_Textes = VueActive.Texts
    FormatNomVue = False
    FormatScaleVue = False
    PasDeTexte = False

    NomVueActive = VueActive.Name
    NewTreeViewName = NomVueActive

    VueActive.GetViewName ViewName, ViewIdent, VewSuffix
    If Col_Textes.Count = 0 Then
        PasDeTexte = True
    Else
        'Recheche le texte correspondant au nom de la vue
        'For j = 1 To Col_Textes.Count
        '    If InStr(1, UCase(Col_Textes.Item(j).Text), UCase(ViewName), vbTextCompare) > 0 Then
        '        Set TxtNomVue = Col_Textes.Item(j)
        '    End If
        'Next j
        'If TxtNomVue Is Nothing Then
            Set TxtNomVue = Col_Textes.Item(1)
        'End If
    End If
    
    'Vue principales et dérivées Anglais
    If Left(NomVueActive, Len(ANameVueF)) = ANameVueF Or Left(NomVueActive, Len(ANameVueD)) = ANameVueD Or Left(NomVueActive, Len(ANameVueG)) = ANameVueG Or Left(NomVueActive, Len(ANameVueH)) = ANameVueH Or Left(NomVueActive, Len(ANameVueB)) = ANameVueB Then
        If VueActive.Scale = iScaleVuePrincipale Then
            'Col_Textes.Item(1).Activity = False
            NewViewName = ""
        Else
            NewViewName = "SCALE : " & FormatScale(VueActive.Scale)
            'NewViewName = "" & Chr$(10) & "ECHELLE : " & FormatScale(VueActive.Scale) & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        
    'Vue principales et dérivées Français
    ElseIf Left(NomVueActive, Len(FNameVueF)) = FNameVueF Or Left(NomVueActive, Len(FNameVueD)) = FNameVueD Or Left(NomVueActive, Len(FNameVueG)) = FNameVueG Or Left(NomVueActive, Len(FNameVueH)) = FNameVueH Or Left(NomVueActive, Len(FNameVueB)) = FNameVueB Then
        If VueActive.Scale = iScaleVuePrincipale Then
            'Col_Textes.Item(1).Activity = False
            NewViewName = ""
        Else
            NewViewName = "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        
    'Vues Iso Anglais ou Français
    ElseIf Left(NomVueActive, Len(ANameVueIso)) = ANameVueIso Or Left(NomVueActive, Len(FNameVueIso)) = FNameVueIso Then
            NewViewName = "ISOMETRIC VIEW"
            FormatNomVue = True
            NewTreeViewName = "ISOMETRIC VIEW"
            
    'coupes Anglais
    ElseIf Left(NomVueActive, Len(ANameCoupe)) = ANameCoupe Then
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "SECTION " & Mid(NomVueActive, Len(ANameCoupe) + 2, 3)
        Else
            NewViewName = "SECTION " & Mid(NomVueActive, Len(ANameCoupe) + 2, 3) & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        'NewTreeViewName = "SECTION " & Mid(NomVueActive, Len(ANameCoupe) + 2, 3)
        NewTreeViewName = "SECTION "
        
    'coupes Français
    ElseIf Left(NomVueActive, Len(FNameCoupe)) = FNameCoupe Then
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "SECTION " & Mid(NomVueActive, Len(FNameCoupe) + 2, 3)
        Else
            NewViewName = "SECTION " & Mid(NomVueActive, Len(FNameCoupe) + 2, 3) & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        'NewTreeViewName = "SECTION " & Mid(NomVueActive, Len(FNameCoupe) + 2, 3)
        NewTreeViewName = "SECTION "
        
    'Sections Anglais
    ElseIf Left(NomVueActive, Len(ANameSection)) = ANameSection Then
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "SECTION " & Mid(NomVueActive, Len(ANameSection) + 2, 3)
        Else
            NewViewName = "SECTION " & Mid(NomVueActive, Len(ANameSection) + 2, 3) & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        'NewTreeViewName = "SECTION " & Mid(NomVueActive, Len(ANameSection) + 2, 3)
        NewTreeViewName = "SECTION "
        
    'Sections Français
    ElseIf Left(NomVueActive, Len(FNameSection)) = FNameSection Then
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "SECTION " & Mid(NomVueActive, Len(FNameSection) + 2, 3)
        Else
            NewViewName = "SECTION " & Mid(NomVueActive, Len(FNameSection) + 2, 3) & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        'NewTreeViewName = "SECTION " & Mid(NomVueActive, Len(FNameSection) + 2, 3)
        NewTreeViewName = "SECTION "
    
    'Vues Auxiliaires Anglais
    ElseIf Left(NomVueActive, Len(ANameVueAux)) = ANameVueAux Then
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "VIEW " & Mid(NomVueActive, Len(ANameVueAux) + 2, 1)
        Else
            NewViewName = "VIEW " & Mid(NomVueActive, Len(ANameVueAux) + 2, 1) & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        'NewTreeViewName = "VIEW " & Mid(NomVueActive, Len(ANameVueAux) + 2, 1)
        NewTreeViewName = "VIEW "
        
    'Vues Auxiliaires  Français
    ElseIf Left(NomVueActive, Len(FNameVueAux)) = FNameVueAux Then
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "VIEW " & Mid(NomVueActive, Len(FNameVueAux) + 2, 1)
        Else
            NewViewName = "VIEW " & Mid(NomVueActive, Len(FNameVueAux) + 2, 1) & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        'NewTreeViewName = "VIEW " & Mid(NomVueActive, Len(FNameVueAux) + 2, 1)
        NewTreeViewName = "VIEW "
    
    'Vues de détail Anglais ou Français
    ElseIf Left(NomVueActive, 6) = ANameDetail Or Left(NomVueActive, 6) = ANameDetail Then
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "DETAIL " & Mid(NomVueActive, Len(ANameDetail) + 2, 1)
        Else
            NewViewName = "DETAIL " & Mid(NomVueActive, Len(ANameDetail) + 2, 1) & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        'NewTreeViewName = "DETAIL " & Mid(NomVueActive, Len(ANameDetail) + 2, 1)
        NewTreeViewName = "DETAIL "
        
    'Vues Dépliées Anglaise
    ElseIf Left(NomVueActive, Len(ANameVueDep)) = ANameVueDep Or Left(NomVueActive, Len(FNameVueDep)) = FNameVueDep Then
        
        'Recherche du format du calque et paramétrage de la position du tableau
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "UNFOLDED VIEW"
        Else
            NewViewName = "UNFOLDED VIEW " & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        NewTreeViewName = "UNFOLDED VIEW"
        
        'Ajout nota
        NotaUnfold_Y = 160
        NotaUnfold_X = DimCalqueX(CalqueActif.PaperName) - 420
        Notatxt = "NOTE:" & Chr(10) & " BEND ALLOWANCE NOT CALCULATED ON UNFOLDED VIEW"
        Set TxtNota = BackVue.Texts.Add(Notatxt, NotaUnfold_X, NotaUnfold_Y)
        TxtNota.SetFontSize 0, 0, 5
        TxtNota.SetFontSize 1, 5, 8
        TxtNota.Name = "TxtNota"
   
    Else
        PasDeTexte = True
    End If
        
    'Mise en forme du texte
    If Not PasDeTexte Then
        'Col_Textes.Item(1).Text = NewViewName
        TxtNomVue.Text = NewViewName
        Deb = InStr(1, NewViewName, "SCALE :")
        Fin = Len(NewViewName) - InStr(1, NewViewName, "SCALE :") + 1
        'Cour circuite la fonction de justification si le premier texte de la vue est un ballon
        On Error Resume Next
            'Col_Textes.Item(1).TextProperties.Justification = catCenter
            TxtNomVue.TextProperties.Justification = catCenter
        Err.Clear
        On Error GoTo 0
        If FormatNomVue Then
            'With Col_Textes.Item(1)
            With TxtNomVue
                .SetParameterOnSubString catItalic, 0, 0, iFontItalic
                .SetParameterOnSubString catUnderline, 0, 0, iFontUnderline
                .SetParameterOnSubString catBold, 0, 0, iFontBold
                .SetFontName 0, 0, "Monospac821"
                .SetFontSize 0, 0, iFontSize
            End With
            VueActive.SetViewName NewTreeViewName, ViewIdent, ""
                    
        End If
        If FormatScaleVue Then
            'With Col_Textes.Item(1)
            With TxtNomVue
                .SetParameterOnSubString catItalic, Deb, Fin, iFontItalicScale
                .SetParameterOnSubString catUnderline, Deb, Fin, iFontUnderlineScale
                .SetParameterOnSubString catBold, Deb, Fin, iFontBoldScale
                .SetFontName 0, 0, "Monospac821"
                .SetFontSize Deb, Fin, iFontSizeScale
            End With
        End If
    End If
Next

    'Mise à jour de l'echelle dans le cartouche
    'MajEchCart BackVue.Texts, iScaleVuePrincipale
    

End Sub

Private Sub MajEchCart(oVue As DrawingView, strEch As Double)
'Met à jour le texte "Echelle" du cartouche
Dim txtEch As DrawingText
Dim Vuetxts As DrawingTexts

    On Error Resume Next 'au cas ou le texte n'existerait pas
    Set Vuetxts = oVue.Texts
    Set txtEch = Vuetxts.GetItem("Texte.Scale")
    txtEch.Text = FormatScale(strEch)

    On Error GoTo 0
End Sub
Public Function FormatScale(dblEch As Double) As String
'renvoi l'echelle au format x:x
    Select Case dblEch
        Case 0.02
            FormatScale = "1:50"
        Case 0.05
            FormatScale = "1:20"
        Case 0.1
            FormatScale = "1:10"
        Case 0.2
            FormatScale = "1:5"
        Case 0.4
            FormatScale = "2:5"
        Case 0.5
            FormatScale = "1:2"
        Case 0.6
            FormatScale = "3:5"
        Case 1
            FormatScale = "1:1"
        Case 2
            FormatScale = "2:1"
        Case 2.5
            FormatScale = "5:2"
        Case 5
            FormatScale = "5:1"
        Case 10
            FormatScale = "10:1"
        Case 20
            FormatScale = "20:1"
        Case 50
            FormatScale = "50:1"
    End Select
End Function

Public Function ConvertScale(strEch As String) As Double
'Converti les format d'échelle fractionaires en multiplicateur
    Select Case strEch
        Case "1:50"
            ConvertScale = 0.02
        Case "1:20"
            ConvertScale = 0.05
        Case "1:10"
            ConvertScale = 0.1
        Case "1:5"
            ConvertScale = 0.2
        Case "2:5"
            ConvertScale = 0.4
        Case "1:2"
            ConvertScale = 0.5
        Case "3:5"
            ConvertScale = 0.6
        Case "1:1"
            ConvertScale = 1
        Case "2:1"
            ConvertScale = 2
        Case "5:2"
            ConvertScale = 2.5
        Case "5:1"
            ConvertScale = 5
        Case "10:1"
            ConvertScale = 10
        Case "20:1"
            ConvertScale = 20
        Case "50:1"
            ConvertScale = 50
    End Select
End Function

Private Function DimCalqueX(CalqueActifPaperSize As String) As Integer
'Renvoi la dimension en X du calque
    Select Case CalqueActifPaperSize
        Case "A0 ISO"
            DimCalqueX = 1189
        Case "A1 ISO"
            DimCalqueX = 841
        Case "A2 ISO"
            DimCalqueX = 594
        Case "A3 ISO"
            DimCalqueX = 420
    End Select
End Function

Private Function DimCalqueY(CalqueActifPaperSize As String) As Integer
'Renvoi la dimension en Y du calque
    Select Case CalqueActifPaperSize
        Case "A0 ISO"
            DimCalqueY = 841
        Case "A1 ISO"
            DimCalqueY = 594
        Case "A2 ISO"
            DimCalqueY = 420
        Case "A3 ISO"
            DimCalqueY = 297
    End Select
End Function
