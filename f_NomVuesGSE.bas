Attribute VB_Name = "f_NomVuesGSE"
Option Explicit

Sub catmain()
'Macro de renommage des vue pour liasses GSE Colomiers
'modification du 10/02/11 ==> suppression des textes en Français et changement taille texte (8 ald 7)
'modification du 15/02.01 ==> suppression soulignement et italique. ajout "Section" aux sections
'                             Taille echelles 5 ald 3.5
'modification de 25/11/15 ==> Ajout d'un formulaire avec la liste des echelles

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "f_NomVuesGSE", VMacro

Dim DocActif As DrawingDocument
Set DocActif = CATIA.ActiveDocument

Dim Col_Calques As DrawingSheets
Set Col_Calques = DocActif.Sheets

Dim CalqueActif As DrawingSheet
Set CalqueActif = Col_Calques.ActiveSheet
Dim NomCalqueActif As String
NomCalqueActif = CalqueActif.Name

Dim Col_Vues As DrawingViews
Set Col_Vues = CalqueActif.Views
Dim EchelleCalqueActif As Double
EchelleCalqueActif = CalqueActif.Scale

Dim VueActive As DrawingView, BackVue As DrawingView
Dim NomVueActive As String, NewViewName As String
Dim ViewName As String, ViewIdent As String, VewSuffix As String
Dim NewTreeViewName As String, NewTreeViewIdent As String, NewTreeVewSuffix As String

Set BackVue = Col_Vues.Item("Background View")

Dim Col_Textes As DrawingTexts

'paramètres de texte
Const iFontSize As Integer = 8
Const iFontSizeScale As Integer = 5
Const iFontBold As Integer = 0
Const iFontBoldScale As Integer = 0
Const iFontUnderline As Integer = 0
Const iFontUnderlineScale As Integer = 0
Const iFontItalic As Integer = 0
Const iFontItalicScale As Integer = 0

 Dim i As Long
 
'Facteur d'échelle de la vue principale
Dim iScaleVuePrincipale As Double
Load frm_Echelles
frm_Echelles.Show
iScaleVuePrincipale = ConvertScale(frm_Echelles.CBL_Echelles.Value)
'iScaleVuePrincipale = InputBox("Qu'elle est l'échelle de la vue principale ? (ex: 1; 2, 0.5) :", "Echelle vue principale")
Unload frm_Echelles

'Nom des type de vues en fonction de la langue
Dim ANameVueF As String, ANameVueD As String, ANameVueG As String, ANameVueH As String, ANameVueB As String
Dim ANameSection As String, ANameCoupe As String, ANameDetail As String, ANameVueAux As String, ANameVueIso As String
Dim ANameVueDep As String
Dim FNameVueF As String, FNameVueD As String, FNameVueG As String, FNameVueH As String, FNameVueB As String
Dim FNameSection As String, FNameCoupe As String, FNameDetail As String, FNameVueAux As String, FNameVueIso As String
Dim FNameVueDep As String

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

Dim Deb, Fin As Integer
Dim FormatNomVue, FormatScaleVue, PasDeTexte As Boolean

For i = 3 To Col_Vues.Count
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
        
        If VueActive.Scale = iScaleVuePrincipale Then
            NewViewName = "UNFOLDED VIEW"
        Else
            NewViewName = "UNFOLDED VIEW " & Chr$(10) & "SCALE : " & FormatScale(VueActive.Scale)
            FormatScaleVue = True
        End If
        FormatNomVue = True
        NewTreeViewName = "UNFOLDED VIEW"
        'Ajout nota
        
        Dim NotaUnfold_Y As Double
            NotaUnfold_Y = 160
        Dim NotaUnfold_X As Double
            NotaUnfold_X = Dim_Calque_X - 420
        Dim Notatxt As String
            Notatxt = "NOTE:" & Chr(10) & " BEND ALLOWANCE NOT CALCULATED ON UNFOLDED VIEW"
        Dim TxtNota As DrawingText
            Set TxtNota = BackVue.Texts.Add(Notatxt, NotaUnfold_X, NotaUnfold_Y)
            TxtNota.SetFontSize 0, 0, 5
            TxtNota.SetFontSize 1, 5, 8
            TxtNota.Name = "TxtNota"
   
    Else
        PasDeTexte = True
    End If
        
    'Mise en forme du texte
    If Not PasDeTexte Then
        Col_Textes.Item(1).Text = NewViewName
        Deb = InStr(1, NewViewName, "SCALE :")
        Fin = Len(NewViewName) - InStr(1, NewViewName, "SCALE :") + 1
        Col_Textes.Item(1).TextProperties.Justification = catCenter
        If FormatNomVue Then
            With Col_Textes.Item(1)
                .SetParameterOnSubString catItalic, 0, 0, iFontItalic
                .SetParameterOnSubString catUnderline, 0, 0, iFontUnderline
                .SetParameterOnSubString catBold, 0, 0, iFontBold
                .SetFontName 0, 0, "Monospac821"
                .SetFontSize 0, 0, iFontSize
            End With
            VueActive.SetViewName NewTreeViewName, ViewIdent, ""
            
            
        End If
        If FormatScaleVue Then
            With Col_Textes.Item(1)
                .SetParameterOnSubString catItalic, Deb, Fin, iFontItalicScale
                .SetParameterOnSubString catUnderline, Deb, Fin, iFontUnderlineScale
                .SetParameterOnSubString catBold, Deb, Fin, iFontBoldScale
                .SetFontName 0, 0, "Monospac821"
                .SetFontSize Deb, Fin, iFontSizeScale
            End With
        End If
    End If
Next


End Sub

Public Function FormatScale(FSchaine As Double) As String
'renvoi l'echelle au format x/x
If FSchaine = 0.02 Then FormatScale = "1:50"
If FSchaine = 0.05 Then FormatScale = "1:20"
If FSchaine = 0.1 Then FormatScale = "1:10"
If FSchaine = 0.2 Then FormatScale = "1:5"
If FSchaine = 0.4 Then FormatScale = "2:5"
If FSchaine = 0.5 Then FormatScale = "1:2"
If FSchaine = 0.6 Then FormatScale = "3:5"
If FSchaine = 1 Then FormatScale = "1:1"
If FSchaine = 2 Then FormatScale = "2:1"
If FSchaine = 2.5 Then FormatScale = "5:2"
If FSchaine = 5 Then FormatScale = "5:1"
If FSchaine = 10 Then FormatScale = "10:1"
If FSchaine = 20 Then FormatScale = "20:1"
If FSchaine = 50 Then FormatScale = "50:1"

End Function

Public Function ConvertScale(CSChaine As String) As Double
'Converti les format d'échelle fractionaires en multiplicateur
If CSChaine = "1:50" Then ConvertScale = 0.02
If CSChaine = "1:20" Then ConvertScale = 0.05
If CSChaine = "1:10" Then ConvertScale = 0.1
If CSChaine = "1:5" Then ConvertScale = 0.2
If CSChaine = "2:5" Then ConvertScale = 0.4
If CSChaine = "1:2" Then ConvertScale = 0.5
If CSChaine = "3:5" Then ConvertScale = 0.6
If CSChaine = "1:1" Then ConvertScale = 1
If CSChaine = "2:1" Then ConvertScale = 2
If CSChaine = "5:2" Then ConvertScale = 2.5
If CSChaine = "5:1" Then ConvertScale = 5
If CSChaine = "10:1" Then ConvertScale = 10
If CSChaine = "20:1" Then ConvertScale = 20
If CSChaine = "50:1" Then ConvertScale = 50
End Function
