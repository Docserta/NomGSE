Attribute VB_Name = "c_Creation_Plan"

Option Explicit
Public ListDetails As String

Sub catmain()
' *****************************************************************
' * Cr�e un nouveau Drawing et importe un cartouche en fonction de certains crit�res
' * � Plan d'ensemble ou plan de d�tail
' * � Langue
' * � Nom du client
' * Met ensuite a jours le cartouche avec les infos saisies dans la boite de dialogue "Cartouche"
' * Cr�ation CFR le 14/08/2012
' * modification le : 18/09/14
' *    Ajout module de classe xMacroLocation
' * modification le : 28/10/14
' *    Prise en compte de 2 systemes de num�rotation des achats 500 � 999 ou 700 � 900
' * modification le : 24/11/14
' *    Ajout dans le calque de d�tail d'un texte portant le num�ro du part/product li� au plan
' *    pour initialisation de la macro d_nomenclature2D
' * modification le 21/12/15
' *    Prise en compte des plan multiples (part sym, products avec sym ou variantes)
' *****************************************************************

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "c_Creation_Plan", VMacro

'Chargement des variables
Dim MacroLocation As New xMacroLocation
If Not (MacroLocation.FicIniExist("VarNomenclatureGSE.ini")) Then ' on est pas dans l'environnement GSE_R21
    MsgBox "Vous n'etes pas dans l'environnement GRE_R21. La macro ne peut pas fonctionner!", vbCritical, "erreur d'environneemnt"
    Exit Sub
Else
    MacroLocation.LectureFichierIni = "VarNomenclatureGSE.ini"
    CheminSourcesMacro = MacroLocation.ValVar("CheminSourcesMacro")
    CheminCartouches = MacroLocation.ValVar("CheminCartouches")
End If

Dim i As Long
Dim FichTxt As String

'Test si le document actif est un Part ou un Product
    Dim MsgErr As String
    MsgErr = "Cette macro n�c�ssite qu'un CATPart ou un CATProduct soit ouvert"
    Dim DetectErrProd As Boolean, DetectErrPart As Boolean
    DetectErrProd = False
    DetectErrPart = False
    
    On Error Resume Next
    Set ActiveDoc = CATIA.ActiveDocument
    If (Err <> 0) Then
        DetectErrProd = True
        MsgBox MsgErr, vbCritical, "Environnement incorect"
        End
    End If
    Dim ActiveProductDoc As ProductDocument
    Dim ActivePartDoc As PartDocument

    Set ActiveProductDoc = ActiveDoc
    If (Err <> 0) Then
        DetectErrProd = True
        Err.Clear
    End If
    Set ActivePartDoc = ActiveDoc
    If (Err <> 0) Then
        DetectErrPart = True
        Err.Clear
    End If
    If DetectErrProd And DetectErrPart Then
        MsgBox MsgErr, vbCritical, "Environnement incorect"
        End
    End If
    On Error GoTo 0

 'Cr�ation objet fichier texte
    Dim fs, f
    Set fs = CreateObject("scripting.filesystemobject")

'Chargement de la boite de dialogue Cartouche
    Load Frm_Cartouche
'D�sactivation des controles
    Frm_Cartouche.Cdr_Cartouche.Enabled = False
'Type de plan
    Frm_Cartouche.RBt_TypePlan1 = 1
'Initialisation des listes d�roulantes
    Set fs = CreateObject("scripting.filesystemobject")
'Remplissage de la liste des LANGUES
    'vide la liste des Langues (cas ou on reaffiche le formulaire sans l'avoir d�charg�)
    If Frm_Cartouche.Cbx_Langue.ListCount >= 1 Then
        For i = Frm_Cartouche.Cbx_Langue.ListCount To 1 Step -1
            Frm_Cartouche.Cbx_Langue.RemoveItem (i - 1)
        Next
    End If
    FichTxt = CheminSourcesMacro & List_Lang
    Set f = fs.opentextfile(FichTxt, ForReading, 1)
    
'Remplissage de la liste LIMITS NOT STATED
    Frm_Cartouche.Cbx_LimNotStated = "ABD0001-3"
    
'Remplissage de la liste SURFACE FINISH
    Frm_Cartouche.Cbx_SurfFinish = "ABD0002"
    
'Remplissage de la liste des INDICES
    FichTxt = CheminSourcesMacro & List_Indices
    Set f = fs.opentextfile(FichTxt, ForReading, 1)
    'Choix de la valeur par defaut (Premi�re valeur de la liste)
    Frm_Cartouche.Cbx_Indice = f.ReadLine
    Do While Not f.AtEndOfStream
        Frm_Cartouche.Cbx_Indice.AddItem (f.ReadLine)
    Loop
    
'Remplissage de la liste des ECHELLES
    FichTxt = CheminSourcesMacro & List_Echelles
    Set f = fs.opentextfile(FichTxt, ForReading, 1)
    'Choix de la valeur par defaut (Premi�re valeur de la liste)
    Frm_Cartouche.Cbx_Echelle = f.ReadLine
    Do While Not f.AtEndOfStream
        Frm_Cartouche.Cbx_Echelle.AddItem (f.ReadLine)
    Loop

'Remplissage liste des planches
    Dim List_Sheet As String
    For i = 1 To 40
        If i < 10 Then
            List_Sheet = "0" & CStr(i)
        Else
            List_Sheet = CStr(i)
        End If
        Frm_Cartouche.Cbx_NoSheet.AddItem List_Sheet
    Next
    
'valeur par defaut liste des planches
    Frm_Cartouche.Cbx_NoSheet = "01"
    Frm_Cartouche.Cbx_NbSheet = "XX"
    
'R�cup�ration du type de num�rotation
    Set Coll_Documents = CATIA.Documents
    Set ActiveDoc = CATIA.ActiveDocument
    Dim MonProduct As Product
    Set MonProduct = ActiveDoc.Product
    Dim MesParametres As Parameters
    Set MesParametres = MonProduct.UserRefProperties
    TypeNum = RecupParam(MesParametres, "NomPulsGSE_TypeNum")

'affiche le Num�ro du document actif dans le liste des fichiers ouverts
    'test si c'est un part ou un product
    If Right(ActiveDoc.Name, 8) = ".CATPart" Then
        'Cr�ation de la liste des fichiers Catparts ouverts
        Frm_Cartouche.RBt_TypePlan2.Value = True
        Frm_Cartouche.Cbx_FicOuvert.List = ListPartProductOpen(2)
    ElseIf Right(ActiveDoc.Name, 11) = ".CATProduct" Then
        'Cr�ation de la liste des fichiers Catproduct ouverts
        Frm_Cartouche.RBt_TypePlan1.Value = True
        Frm_Cartouche.Cbx_FicOuvert.List = ListPartProductOpen(1)
    End If
    Frm_Cartouche.Cbx_FicOuvert.Value = ActiveDoc.Name

Frm_Cartouche.Show
Unload Frm_Cartouche


End Sub


