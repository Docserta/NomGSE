Attribute VB_Name = "e_Macros_GSE"
'Public Liste_Dsk As Variant
Public Liste_Rep As Variant
Public Liste_Fic As Variant
Public Dsk As String
Public Rep As String
Public TypeFic As String
'Public Fic As String


Sub catmain()

On Error GoTo Gestion_Err

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "e_Macros_GSE", VMacro

'Chargement de la boite de dialogue Générale
    Load Lanceur
    Lanceur.Show
    GoTo Fin_Prog

Fin_Prog:
Exit Sub

Gestion_Err:
MsgBox Err.Number
MsgBox Err.Description

End Sub




