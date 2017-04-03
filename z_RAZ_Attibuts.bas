Attribute VB_Name = "z_RAZ_Attibuts"


Sub catmain()
' *****************************************************************
' * Suppression des Attibuts sur tous les parts et Products
' * Création CFR le 29/07/2013
' * Dernière modification le
' *****************************************************************

'NomPulsGSE_DesignOutillage
'NomPulsGSE_NoOutillage
'NomPulsGSE_SiteAB
'NomPulsGSE_CHK
'NomPulsGSE_Client
'NomPulsGSE_DatePlan
'NomPulsGSE_CE
'NomPulsGSE_PresUserGuide
'NomPulsGSE_PresCaisse
'NomPulsGSE_NoCaisse
'NomPulsGSE_Sheet
'NomPulsGSE_ItemNb
'NomPulsGSE_Dimension
'NomPulsGSE_Material
'NomPulsGSE_Protect
'NomPulsGSE_Miscellanous
'NomPulsGSE_SupplierRef
'NomPulsGSE_Weight
'NomPulsGSE_MecanoSoude
'NomPulsGSE_TypeNum

'Log l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "z_RAZ_Attibuts", VMacro

'tous les documents ouverts
Set Coll_Documents = CATIA.Documents
On Error Resume Next

Dim MonProductDoc As Document
Dim MonProduct As Product

Dim NBrPRoducts As Integer
Dim SearchString, NomParam As String
SearchString = "NomPulsGSE"
Dim ParamEC As StrParam
'Paramètres
Dim MesParametres As Parameters
'Nombre d'Item dans Coll_documents
Load Frm_ChoixAttibuts
Frm_ChoixAttibuts.Show
Frm_ChoixAttibuts.Hide
'Avertissement
Dim Msg As String
If Frm_ChoixAttibuts.OB_efface Then
    Msg = "Etes vous sur de vouloir effacer le contenu de tous les paramètres sélectionnés sur tous les parts et product ?"
ElseIf Frm_ChoixAttibuts.OB_Supp Then
    Msg = "Etes vous sur de vouloir supprimer tous les paramètres sélectionnés sur tous les parts et product ?"
End If
If MsgBox(Msg, vbOKCancel, "Nettoyage des paramètres") = vbOK Then

'Chargement de la barre de progression
    Load Frm_Progression
    Frm_Progression.Show vbModeless
    Frm_Progression.Caption = " Suppression des paramètres. Veuillez patienter..."
    ProgressBar (1)

    NBrPRoducts = Coll_Documents.Count
    For i = 1 To NBrPRoducts
        ProgressBar (100 / NBrPRoducts * i)
        Set MonProductDoc = Coll_Documents.Item(i)
        Set MonProduct = MonProductDoc.Product
        Set MesParametres = MonProduct.ReferenceProduct.UserRefProperties
        If Frm_ChoixAttibuts.OB_efface Then
            For Each ParamEC In MesParametres
                NomParam = CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_DesignOutil And InStr(1, ParamEC.Name, "NomPulsGSE_DesignOutillage") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_NoOutil And InStr(1, ParamEC.Name, "NomPulsGSE_NoOutillage") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_SiteAB And InStr(1, ParamEC.Name, "NomPulsGSE_SiteAB") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_CHK And InStr(1, ParamEC.Name, "NomPulsGSE_CHK") <> 0 Then ParamEC.Value = ""
                If frm_choixattributs.ChB_Client And InStr(1, ParamEC.Name, "NomPulsGSE_Client") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_DatePlan And InStr(1, ParamEC.Name, "NomPulsGSE_DatePlan") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_CE And InStr(1, ParamEC.Name, "NomPulsGSE_CE") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_PresUserGuide And InStr(1, ParamEC.Name, "NomPulsGSE_PresUserGuide") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_PresCaisse And InStr(1, ParamEC.Name, "NomPulsGSE_PresCaisse") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_NoCaisse And InStr(1, ParamEC.Name, "NomPulsGSE_NoCaisse") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_Sheet And InStr(1, ParamEC.Name, "NomPulsGSE_Sheet") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_ItemNB And InStr(1, ParamEC.Name, "NomPulsGSE_ItemNb") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_Dimension And InStr(1, ParamEC.Name, "NomPulsGSE_Dimension") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_Material And InStr(1, ParamEC.Name, "NomPulsGSE_Material") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_Protect And InStr(1, ParamEC.Name, "NomPulsGSE_Protect") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_Miscellanous And InStr(1, ParamEC.Name, "NomPulsGSE_Miscellanous") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_SupplierRef And InStr(1, ParamEC.Name, "NomPulsGSE_SupplierRef") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_Weight And InStr(1, ParamEC.Name, "NomPulsGSE_Weight") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_MecanoSoude And InStr(1, ParamEC.Name, "NomPulsGSE_MecanoSoude") <> 0 Then ParamEC.Value = ""
                If Frm_ChoixAttibuts.ChB_TypNum And InStr(1, ParamEC.Name, "NomPulsGSE_TypeNum") <> 0 Then ParamEC.Value = ""
            Next
        ElseIf Frm_ChoixAttibuts.OB_Supp Then
            For Each ParamEC In MesParametres
                NomParam = CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_DesignOutil And InStr(1, ParamEC.Name, "NomPulsGSE_DesignOutillage") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_NoOutil And InStr(1, ParamEC.Name, "NomPulsGSE_NoOutillage") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_SiteAB And InStr(1, ParamEC.Name, "NomPulsGSE_SiteAB") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_CHK And InStr(1, ParamEC.Name, "NomPulsGSE_CHK") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_Client And InStr(1, ParamEC.Name, "NomPulsGSE_Client") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_DatePlan And InStr(1, ParamEC.Name, "NomPulsGSE_DatePlan") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_CE And InStr(1, ParamEC.Name, "NomPulsGSE_CE") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_PresUserGuide And InStr(1, ParamEC.Name, "NomPulsGSE_PresUserGuide") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_PresCaisse And InStr(1, ParamEC.Name, "NomPulsGSE_PresCaisse") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_NoCaisse And InStr(1, ParamEC.Name, "NomPulsGSE_NoCaisse") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_Sheet And InStr(1, ParamEC.Name, "NomPulsGSE_Sheet") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_ItemNB And InStr(1, ParamEC.Name, "NomPulsGSE_ItemNb") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_Dimension And InStr(1, ParamEC.Name, "NomPulsGSE_Dimension") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_Material And InStr(1, ParamEC.Name, "NomPulsGSE_Material") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_Protect And InStr(1, ParamEC.Name, "NomPulsGSE_Protect") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_Miscellanous And InStr(1, ParamEC.Name, "NomPulsGSE_Miscellanous") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_SupplierRef And InStr(1, ParamEC.Name, "NomPulsGSE_SupplierRef") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_Weight And InStr(1, ParamEC.Name, "NomPulsGSE_Weight") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_MecanoSoude And InStr(1, ParamEC.Name, "NomPulsGSE_MecanoSoude") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
                If Frm_ChoixAttibuts.ChB_TypNum And InStr(1, ParamEC.Name, "NomPulsGSE_TypeNum") <> 0 Then MesParametres.Remove CStr(ParamEC.Name)
            Next
        End If
        
    Next
Unload Frm_Progression
Unload Frm_ChoixAttibuts
End If
End Sub
