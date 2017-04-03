Attribute VB_Name = "C_Change_Format_Plan"

Option Explicit


Sub catmain()
' *****************************************************************
' * Change le format du plan
' * Efface le cadre de l'ancien format et trace le nouveau cadre
' *
' * Création CFR le 14/08/2012
' *****************************************************************
Dim MacroLocation As New xMacroLocation
Dim ActiveDrawingDoc As DrawingDocument
Dim Col_CalquesDrawActif As DrawingSheets
    
Dim ActiveDrawingVues As DrawingViews
Dim BackgroundView As DrawingView
Dim ActiveDrawingCalque As DrawingSheet

Dim Col_Elem2D  As GeometricElements
Dim Elem2D As GeometricElement
Dim Col_geomElem As GeometricElements
Dim GeomElem As GeometricElement
    
Dim FormatActuel As CatPaperSize
Dim ActualFmt As Fmt, FuturFmt As Fmt

Dim DetailsInstancies As DrawingComponents
Dim DetailEC As DrawingComponent

Dim Coll_Tables As DrawingTables
Dim Tabl As DrawingTable

Dim Col_Textes As DrawingTexts
Dim Txt As DrawingText
    
Dim LigEC As Line2D
Dim CercleEC As Circle2D
    
Dim col_img As DrawingPictures
Dim imgLogo As DrawingPicture

Dim sel As Selection
Dim Move_X As Double, Move_Y As Double

Dim CoordTrace() As String
Dim TiretEC As Double
Dim DemiTiret As Double
Dim LettreRep As String
Dim ListLettreRep As Variant

Dim i As Long
        i = 0
        
'Log l'utilisation de la macro
    LogUtilMacro nPath, nFicLog, nMacro, "C_Change_Format_Plan", VMacro

'Chargement des variables
    If Not (MacroLocation.FicIniExist("VarNomenclatureGSE.ini")) Then ' on est pas dans l'environnement GSE_R21
        MsgBox "Vous n'etes pas dans l'environnement GRE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
        Exit Sub
    Else
        MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
        CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
        NomFicLigneCalque = MacroLocation.ValVar("NomFicLigneCalque")
        NomFicCercleCalque = MacroLocation.ValVar("NomFicCercleCalque")
        NomFicLigneISDE = MacroLocation.ValVar("NomFicLigneISDE")
        NomFicLigneISFR = MacroLocation.ValVar("NomFicLigneISFR")
    End If

'Test si le document actif est un Drawing
    Set ActiveDoc = CATIA.ActiveDocument
    On Error Resume Next
    Set ActiveDrawingDoc = ActiveDoc
    If (Err <> 0) Then
        MsgBox "Le document actif n'est pas un drawing. Activez un drawing avant de lancer cette macro.", vbCritical, "Erreur"
        Exit Sub
    End If
    On Error GoTo 0
    Set Col_CalquesDrawActif = ActiveDrawingDoc.Sheets
    Set ActiveDrawingCalque = ActiveDrawingDoc.Sheets.ActiveSheet
    Set ActiveDrawingVues = ActiveDrawingCalque.Views
    Set BackgroundView = ActiveDrawingVues.Item("Background View")
    Set sel = ActiveDoc.Selection
    sel.Clear
    
'Recupération du format actuel pour calcul du déplacement du cartouche et de la nomenclature
    FormatActuel = ActiveDrawingCalque.PaperSize
    Select Case FormatActuel
        Case catPaperA0
            ActualFmt.Name = "A0"
            ActualFmt.X = 1189
            ActualFmt.Y = 841
        Case catPaperA1
            ActualFmt.Name = "A1"
            ActualFmt.X = 841
            ActualFmt.Y = 594
        Case catPaperA2
            ActualFmt.Name = "A2"
            ActualFmt.X = 594
            ActualFmt.Y = 420
        Case catPaperA3
            ActualFmt.Name = "A3"
            ActualFmt.X = 420
            ActualFmt.Y = 297
    End Select
    
'Choix du nouveau format
Load frm_Formats
    frm_Formats.CBL_Formats = ActualFmt.Name
frm_Formats.Show
    If Not (frm_Formats.ChB_OkAnnule) Then
        Exit Sub
    End If

'Redimensionnement du calque
    Select Case frm_Formats.CBL_Formats
        Case "A0"
            FuturFmt.X = 1189
            FuturFmt.Y = 841
            FuturFmt.Size = catPaperA0
        Case "A1"
            FuturFmt.X = 841
            FuturFmt.Y = 594
            FuturFmt.Size = catPaperA1
        Case "A2"
            FuturFmt.X = 594
            FuturFmt.Y = 420
            FuturFmt.Size = catPaperA2
        Case "A3"
            FuturFmt.X = 420
            FuturFmt.Y = 297
            FuturFmt.Size = catPaperA3
    End Select
    ActiveDrawingCalque.PaperSize = FuturFmt.Size
    
'Recallage des éléments 2D existants (cartouche et nomenclature)
    Set Col_Elem2D = BackgroundView.GeometricElements

'Calcul du décallage
    Move_X = -(ActualFmt.X - FuturFmt.X)
    Move_Y = -(ActualFmt.Y - FuturFmt.Y)

'Déplacement des tableau de nomenclature dans le drawing
    Set Coll_Tables = BackgroundView.Tables
    For Each Tabl In Coll_Tables
        Tabl.X = Tabl.X + Move_X
        'Tabl.Y = Tabl.Y + Move_Y 'pas de déplacement en Y
    Next

'Deplacement des textes
    Set Col_Textes = BackgroundView.Texts
    sel.Clear
    For Each Txt In Col_Textes
        Txt.X = Txt.X + Move_X
        'Deplacement du texte de modification en Y
        If Txt.Name = "Texte.Mod" _
            Or Txt.Name = "Modifications_Issue" _
             Or Txt.Name = "Issue_txt1" _
              Or Txt.Name = "Issue_txt2" _
               Or Txt.Name = "Issue_txt3" _
                Or Txt.Name = "Issue_txt4" _
                 Or Txt.Name = "Texte.DateMod" _
                  Or Txt.Name = "Modification_Designer" _
                   Or Txt.Name = "Modification_Type" _
            Then
            Txt.Y = Txt.Y + Move_Y
        ElseIf Left(Txt.Name, 5) = "Cart_" Then
            sel.Add Txt
        End If
        'Ajout des lettre repères à la selection pour suppression
    Next

'Déplacement des Ditos
    Set DetailsInstancies = BackgroundView.Components
    If DetailsInstancies.Count > 0 Then
        For i = 1 To DetailsInstancies.Count
            Set DetailEC = DetailsInstancies.Item(i)
           DetailEC.X = DetailEC.X + Move_X
           'DetailEC.Y = DetailEC.Y + Move_Y
        Next
        
    End If
'Déplacement du logo Airbus
    Set col_img = BackgroundView.Pictures
    On Error Resume Next
    Set imgLogo = col_img.Item("Picture.1")
        imgLogo.X = imgLogo.X + Move_X
    On Error GoTo 0
    
'Mise a jour du texte de format
    Set Txt = Col_Textes.GetItem("Texte.Size")
    Txt.Text = frm_Formats.CBL_Formats

'Supression du cardre existant
    Set Col_geomElem = BackgroundView.GeometricElements
    On Error Resume Next
    For Each GeomElem In Col_geomElem
        Set LigEC = GeomElem
        If Not (Err.Number <> 0) Then
            sel.Add LigEC
        End If
        Err.Clear
        Set CercleEC = GeomElem
        If Not (Err.Number <> 0) Then
            sel.Add CercleEC
        End If
        Err.Clear
    Next
    sel.Delete
    
'Collecte des coordonnées des lignes du nouveau calque
'trace du calque
    CoordTrace = Rempl_CoordTrace(CheminSourcesMacro & NomFicLigneCalque, 5)
    'Ajout du décallage en X et en Y
    For i = 0 To UBound(CoordTrace, 2)
        If CoordTrace(5, i) = "FFFV" Then 'lignes cadre gauche
            CoordTrace(1, i) = CoordTrace(1, i)
            CoordTrace(2, i) = CoordTrace(2, i)
            CoordTrace(3, i) = CoordTrace(3, i)
            CoordTrace(4, i) = FuturFmt.Y + CoordTrace(4, i)
        ElseIf CoordTrace(5, i) = "FFVF" Then 'lignes cadre bas
            CoordTrace(1, i) = CoordTrace(1, i)
            CoordTrace(2, i) = CoordTrace(2, i)
            CoordTrace(3, i) = FuturFmt.X - CoordTrace(3, i)
            CoordTrace(4, i) = CoordTrace(4, i)
        ElseIf CoordTrace(5, i) = "VFVV" Then 'ligne Cadre droit
            CoordTrace(1, i) = FuturFmt.X - CoordTrace(1, i)
            CoordTrace(2, i) = CoordTrace(2, i)
            CoordTrace(3, i) = FuturFmt.X - CoordTrace(3, i)
            CoordTrace(4, i) = FuturFmt.Y + CoordTrace(4, i)
        ElseIf CoordTrace(5, i) = "FVVV" Then 'ligne Cadre haut
            CoordTrace(1, i) = CoordTrace(1, i)
            CoordTrace(2, i) = FuturFmt.Y + CoordTrace(2, i)
            CoordTrace(3, i) = FuturFmt.X - CoordTrace(3, i)
            CoordTrace(4, i) = FuturFmt.Y + CoordTrace(4, i)
        ElseIf CoordTrace(5, i) = "VFVF" Then 'ligne cartouche
            CoordTrace(1, i) = FuturFmt.X - CoordTrace(1, i)
            CoordTrace(2, i) = CoordTrace(2, i)
            CoordTrace(3, i) = FuturFmt.X - CoordTrace(3, i)
            CoordTrace(4, i) = CoordTrace(4, i)
        ElseIf CoordTrace(5, i) = "VVVV" Then 'ligne cartouche
            CoordTrace(1, i) = FuturFmt.X - CoordTrace(1, i)
            CoordTrace(2, i) = FuturFmt.Y + CoordTrace(2, i)
            CoordTrace(3, i) = FuturFmt.X - CoordTrace(3, i)
            CoordTrace(4, i) = FuturFmt.Y + CoordTrace(4, i)
        
        End If
    Next i
    TraceLig CoordTrace(), BackgroundView
            
'Tracé des marquage
    TiretEC = 148.5
    CoordTrace() = Empty
    ReDim CoordTrace(4, 2)
    i = 1
    'Marques verticales
    While TiretEC < FuturFmt.Y
    'Trace la marque coté droit
        CoordTrace(0, 1) = "Cart_regD" & i
        CoordTrace(1, 1) = FuturFmt.X - 10
        CoordTrace(2, 1) = TiretEC
        CoordTrace(3, 1) = FuturFmt.X - 5
        CoordTrace(4, 1) = TiretEC
        
    'Trace la marque coté gauche
        CoordTrace(0, 2) = "Cart_regG" & i
        CoordTrace(1, 2) = 5
        CoordTrace(2, 2) = TiretEC
        CoordTrace(3, 2) = 10
        CoordTrace(4, 2) = TiretEC
        
        i = i + 1
        TiretEC = TiretEC + Rgl_V
        
        TraceLig CoordTrace(), BackgroundView
    Wend
    'Marques horizontales
    TiretEC = 105
    i = 1
    While TiretEC < FuturFmt.X
    'Trace la marque bas
        CoordTrace(0, 1) = "Cart_regB" & i
        CoordTrace(1, 1) = FuturFmt.X - TiretEC
        CoordTrace(2, 1) = 5
        CoordTrace(3, 1) = FuturFmt.X - TiretEC
        CoordTrace(4, 1) = 10
        
    'Trace la marque haut
        CoordTrace(0, 2) = "Cart_regH" & i
        CoordTrace(1, 2) = FuturFmt.X - TiretEC
        CoordTrace(2, 2) = FuturFmt.Y - 5
        CoordTrace(3, 2) = FuturFmt.X - TiretEC
        CoordTrace(4, 2) = FuturFmt.Y - 10

        i = i + 1
        TiretEC = TiretEC + Rgl_H
        
        TraceLig CoordTrace(), BackgroundView
    Wend
 
'Tracé des cercles
'Collecte des coordonnées des lignes du nouveau calque
    CoordTrace = Rempl_CoordTrace(CheminSourcesMacro & NomFicCercleCalque, 3)
    For i = 0 To UBound(CoordTrace, 2)
       CoordTrace(1, i) = FuturFmt.X - CoordTrace(1, i)
        CoordTrace(2, i) = CoordTrace(2, i)
    Next i
    TraceCer CoordTrace(), BackgroundView

'Tracé de l'indice de modif
    CoordTrace() = Empty
    'recherche si GSE allemand ou autre '#plus de distinctiuon entre les dito Issue 2016/06
    'Set Txt = Col_Textes.GetItem("Texte.DwgSyst")
    'If Txt.Text = "SEMS1" Then
        CoordTrace = Rempl_CoordTrace(CheminSourcesMacro & NomFicLigneISDE, 5)
    'Else
    '    CoordTrace = Rempl_CoordTrace(CheminSourcesMacro & NomFicLigneISFR, 5)
    'End If
    For i = 0 To UBound(CoordTrace, 2)
        CoordTrace(1, i) = FuturFmt.X - CoordTrace(1, i)
        CoordTrace(2, i) = FuturFmt.Y + CoordTrace(2, i)
        CoordTrace(3, i) = FuturFmt.X - CoordTrace(3, i)
        CoordTrace(4, i) = FuturFmt.Y + CoordTrace(4, i)
    Next i
    TraceLig CoordTrace(), BackgroundView

'Change l'épaisseur des lignes
    ChangeEpLig BackgroundView

'tracé des lettre de repère
    ListLettreRep = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")

'Horizontales
    TiretEC = 105
    DemiTiret = TiretEC / 2
    i = 0
    While TiretEC <= FuturFmt.X
        LettreRep = ListLettreRep(i)
        'Trace la Lettre du bas
        TraceLet Col_Textes, LettreRep, FuturFmt.X - TiretEC + (DemiTiret), 8, "Cart_LetB" & LettreRep, 0
        'Trace la Lettre du haut
        TraceLet Col_Textes, LettreRep, FuturFmt.X - TiretEC + (DemiTiret), FuturFmt.Y - 8.2, "Cart_LetH" & LettreRep, 90
        
        i = i + 1
        TiretEC = TiretEC + 105
    Wend
    
'Verticales
    TiretEC = 148.5
    DemiTiret = TiretEC / 2
    i = 0
    While TiretEC <= FuturFmt.Y
        LettreRep = ListLettreRep(i)
        'Trace la Lettre de Droite
        TraceLet Col_Textes, i + 1, FuturFmt.X - 8, TiretEC - (DemiTiret), "Cart_LetD" & i + 1, 0
        'Trace la Lettre de gauche
        TraceLet Col_Textes, i + 1, 6, TiretEC - (DemiTiret), "Cart_LetG" & i + 1, 90
        
        i = i + 1
        TiretEC = TiretEC + 148.5
    Wend

Unload frm_Formats
End Sub

Private Function Rempl_CoordTrace(NomFic As String, NB_Col As Integer) As String()
'Charge les ccordonnées des ligne de tracé du cartouche depuis un fichier texte
'forme du fichier texte
'nom de l'objet;coord X debut;coord Y debut;coord X fin;coord Y fin;V = coord variable (fonction de la dimension du plan) F= Fixe
'Cart_Issue1;90;-20;10;-20;VVVV
'Cart_Issue3;74;-34;10;-34;VVVV
    Dim fs, f
    Set fs = CreateObject("scripting.filesystemobject")
    Set f = fs.opentextfile(NomFic, ForReading, 1)

    Dim Lig_EC() As String
    Dim Temp_TabCoord() As String
    Dim i As Long, j As Long
    
    While Not f.AtEndOfStream
        i = i + 1
        ReDim Preserve Temp_TabCoord(NB_Col, i)
        Lig_EC = DecoupLig(f.ReadLine)
        For j = 0 To UBound(Lig_EC)
            Temp_TabCoord(j, i) = Lig_EC(j)
        Next j
    Wend
    Rempl_CoordTrace = Temp_TabCoord
End Function

Private Function DecoupLig(str As String) As String()
'decoupe la chaine au niveau de chaque ";"
Dim Temp_Str() As String
Dim i As Integer
    i = 0
    While InStr(1, str, ";", vbTextCompare) > 0
        ReDim Preserve Temp_Str(i)
        Temp_Str(i) = Left(str, InStr(1, str, ";", vbTextCompare) - 1)
        i = i + 1
        str = Right(str, Len(str) - InStr(1, str, ";", vbBinaryCompare))
    Wend
    ReDim Preserve Temp_Str(i)
    Temp_Str(i) = str
    DecoupLig = Temp_Str
End Function

Private Sub TraceLig(List() As String, DrawView As DrawingView)
'Trace les lignes dont les coordonées sont passée dans le tableau List

    Dim Fact2D As Factory2D
    Set Fact2D = DrawView.Factory2D
    Dim LineEC As Line2D
    Dim i As Long
    
    For i = 1 To UBound(List, 2)
        Set LineEC = Fact2D.CreateLine(CDbl(List(1, i)), CDbl(List(2, i)), CDbl(List(3, i)), CDbl(List(4, i)))
        LineEC.Name = List(0, i)
    Next i

End Sub

Private Sub TraceCer(List() As String, DrawView As DrawingView)
'Trace les cercles dont les coordonées sont passée dans le tableau List

    Dim Fact2D As Factory2D
    Set Fact2D = DrawView.Factory2D
    Dim CerEC As Circle2D
    Dim i As Long
    
    For i = 1 To UBound(List, 2)
        Set CerEC = Fact2D.CreateClosedCircle(CDbl(List(1, i)), CDbl(List(2, i)), CDbl(List(3, i)))
        CerEC.Name = List(0, i)
    Next i

End Sub

Private Sub ChangeEpLig(DrawView As DrawingView)
'Change l'épaisseur des lignes dont le nom commence par "Cart_"
    Dim SelLine As Selection
    Dim PropLine As VisPropertySet
    Dim i As Long
    
    Set SelLine = ActiveDoc.Selection
    SelLine.Clear
'selection des lignes
    For i = 1 To DrawView.GeometricElements.Count
        If Left(DrawView.GeometricElements.Item(i).Name, 5) = "Cart_" Then
            SelLine.Add DrawView.GeometricElements.Item(i)
        End If
    Next i
'Changement épaisseur ligne
    Set PropLine = SelLine.VisProperties
    PropLine.SetRealWidth 1, 1
End Sub


Private Sub TraceLet(TL_Textes As DrawingTexts, TL_Str As String, TL_X As Double, TL_Y As Double, TL_Name As String, TL_A As Double)
'Trace la lettre passée en argument
    Dim TL_Txt As DrawingText
    Set TL_Txt = TL_Textes.Add(TL_Str, TL_X, TL_Y)
    TL_Txt.Name = TL_Name
    TL_Txt.Angle = TL_A
    TL_Txt.SetFontName 0, 0, "SSS2"
    TL_Txt.SetFontSize 0, 0, 2.5
    TL_Txt.TextProperties.AnchorPoint = catBottomCenter
End Sub





