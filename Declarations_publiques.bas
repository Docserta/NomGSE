Attribute VB_Name = "Declarations_publiques"
'Fonction de récupération du username
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Version de la macro
Public Const VMacro As String = "Version 4.1.8 du 18/01/17"
Public Const nMacro As String = "NomGSE"
Public Const nPath As String = "\\srvxsiordo\xLogs\01_CatiaMacros"
Public Const nFicLog As String = "logUtilMacro.txt"

'Collections des documents ouverts
Public Coll_Documents As Documents
     
'Document actif
Public ActiveDoc As Document
Public ProductDoc As ProductDocument

'Nom du fichier ecxel contenant le template de la nomenclature Airbus
Public Nom_TemplateAirbus As String

'Nom du fichier des Cages Codes
Public Nom_FicCageCodes As String

'Nom des fichier textes contenant les coordonnée de traçage du cdre et cartouche
    Public NomFicLigneCalque As String
    Public NomFicCercleCalque  As String
    Public NomFicLigneISDE As String
    Public NomFicLigneISFR As String

'nom du fichier en cours dans la liste
    Public CheminPlusNomFichier As String
    Public CheminSourcesMacro As String

'Chemin de destination des nomenclatures
Public CheminDestNomenclature As String
Public CheminDestRapport As String

'Chemin de stockage des cartouches GSE
Public CheminCartouches As String
Public Const VerCart As String = "_V4"

'type de numérotation de Standards (500 à 900 ou 700 à 900)
Public TypeNum As String

'Liste unique des noms de parts
Public ListPartsNames() As String
Public NbPartsNames As Integer

'tableau des parts et de leurs paramètres
Public TableauPartsParam() As String

'Nombre de Paramètres
Public Const NbParam = 10

'Compteur de sous products pour limite barre de progression
Public CompteurLimiteBarre As Integer

'Fichers des listes pour les Listes déroulantes
    Public Const List_Lang As String = "NomGSE-List-Langues.txt"
    Public Const List_LNS As String = "NomGSE-List-LimitNotStated.txt"
    Public Const List_SurfFinish As String = "NomGSE-List-SurfFinish.txt"
    Public Const List_Indices As String = "NomGSE-List-Ind.txt"
    Public Const List_Echelles As String = "NomGSE-List-Echelle.txt"
    Public Const List_Designation As String = "NomGSE-List-Designation.txt"
    Public Const List_Material As String = "NomGSE-List-Material.txt"
    Public Const List_Protect As String = "NomGSE-List-Protect.txt"
    Public Const List_Miscellanous As String = "NomGSE-List-Miscellanous.txt"
    Public Const List_Catalogue As String = "NomGSE-List-Catalogue.txt"
    
' Constantes de Excel
    Public Const xLCenter As Long = -4108
    Public Const xLHaut As Long = -4160
    Public Const xLDroite As Long = -4152
    Public Const xLMoyen As Long = -4138
    Public Const xLNormal As Long = -4143
    Public Const xLMinimized As Long = -4140
    Public Const xLBetween As Long = 1
    Public Const xLCellValue As Long = 1
    Public Const xLGreater As Long = 5
    Public Const xlSolid As Long = 1
    Public Const xlAutomatic As Long = -4105
    Public Const xLTextString As Long = 9
    Public Const xlContains As Long = 0
    Public Const xlEdgeBottom As Long = 9
    Public Const xlContinuous As Long = 1
    Public Const xlThin As Long = 2
    Public Const xlSrcRange As Long = 1
    Public Const xlYes As Integer = 1

Public Const ForReading As Integer = 1

'Formats de plan
Public Type Fmt 'type format de plan
    Name As String
    X As Integer
    Y As Integer
    Size As CatPaperSize
End Type

'Position des Dito
Public Type Dito 'type position des dito
    Name As String
    X As Integer
    Y As Integer
    Size As Double
    Source As DrawingView
    Cible As DrawingComponent
End Type

Public Const Rgl_V As Double = 148.5
Public Const Rgl_H As Double = 105

'Liste des erreurs
'    No                  Module          Description
'vbObjectError + 513, Check2D , "libre"
'vbObjectError + 514, Check2D , "Liens de la vue non reconnus ou brisés"
'vbObjectError + 515, Check2D , "Impossible d'ouvrir le 3D de référence"
'vbObjectError + 520, Check3D , "Ce document n'est ni un Part ni un Product"
