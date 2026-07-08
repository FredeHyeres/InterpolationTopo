Attribute VB_Name = "InterpolationTopoV2"
'==============================================================================
' InterpolationTopoV2 - Module principal - MicroStation V8i SS3 (VBA)
'
' Interpole l'altitude d'un point situe sur la droite reliant deux textes
' d'altitude existants, avec apercu dynamique, puis cree :
'   - un texte altitude (memes proprietes que le texte P1, pris comme modele)
'   - un cercle centre sur le point projete
'
' Architecture V2 : classes a responsabilites separees + formulaire modeless
'
' INSTALLATION :
'   Dans l'editeur VBA (Alt+F11) : Fichier > Importer les fichiers suivants :
'     InterpolationTopoV2.bas
'     CSymboTexte.cls, CSymboCercle.cls, CMstSettings.cls
'     CPointRef.cls, CInterpolation.cls, CMoteurGraphique.cls
'     CSelectP1.cls, CSnapP1.cls, CSelectP2.cls, CSnapP2.cls, CPlacerPoint.cls
'     frmInterpolation.frm
'
' LANCEMENT :
'   key-in : vba run [InterpolationTopoV2]InterpolerPoint
'==============================================================================
Option Explicit

' --- Instances partagees entre les classes de commande ---
' Chaque classe recoit ces references mais ne les cree pas.
Public g_oSettings  As CMstSettings      ' parametres (symbologie, tolerance)
Public g_oP1        As CPointRef         ' point de reference 1
Public g_oP2        As CPointRef         ' point de reference 2
Public g_oCalc      As CInterpolation    ' moteur de calcul pur
Public g_oMoteur    As CMoteurGraphique  ' moteur de creation graphique
Public g_ptOrigineTexte As Point3d       ' origine reelle du point (tag/cellule/texte)
Public g_sValeurTexte   As String        ' valeur texte brute (pour tags ou textes)
Public g_oTagTrouve     As TagElement    ' tag source (Nothing si texte ou cellule)
Public g_oCellTrouvee   As CellElement   ' cellule source (Nothing si texte ou tag)
Public g_sTagDefName    As String        ' nom definition du tag dans la cellule (vide si texte)

'------------------------------------------------------------------------------
Sub InterpolerPoint()
    ' Verifier qu'un fichier DGN est ouvert
    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, "Interpolation Topo"
        Exit Sub
    End If

    ' Instanciation unique de toutes les classes
    Set g_oSettings = New CMstSettings
    g_oSettings.Init

    Set g_oP1 = New CPointRef
    Set g_oP2 = New CPointRef
    Set g_oCalc = New CInterpolation
    Set g_oMoteur = New CMoteurGraphique

    ' Afficher le formulaire modeless (Tool Settings)
    ' Le formulaire lit g_oSettings pour pre-remplir les champs
    frmInterpolation.Initialiser g_oSettings
    frmInterpolation.Show vbModeless

    ' Demarrer la commande
    CommandState.StartPrimitive New CSelectP1
End Sub

'==============================================================================
' Utilitaires partages (appeles par les classes de commande)
'==============================================================================

'------------------------------------------------------------------------------
' Cherche le texte numerique le plus proche du clic dans le rayon donne.
' Facteur commun aux trois classes de commande.
Function TrouverTexteProche(oPt As Point3d, dRayon As Double) As TextElement
    Set TrouverTexteProche = Nothing

    Dim oScan As New ElementScanCriteria
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeText
    oScan.IncludeType msdElementTypeCellHeader
    oScan.IncludeType msdElementTypeTag

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim dMinDist As Double: dMinDist = dRayon
    Dim oNearest As TextElement
    Dim oTagNearest As TagElement
    Dim oCellNearest As CellElement
    Dim ptOrigine As Point3d
    Dim sValeur As String
    Dim sTagDef As String

    Do While oEnum.MoveNext
        Dim oElem As Element
        Set oElem = oEnum.Current

        ' Ignorer les elements sur un niveau gele (non affiche)
        If EstSurNiveauGele(oElem) Then GoTo SuivantElem

        If oElem.Type = msdElementTypeTag Then
            Dim oTag As TagElement
            Set oTag = oElem
            Dim sTagVal As String
            sTagVal = Trim$(CStr(oTag.Value))
            If g_oCalc.EstNombre(Replace(sTagVal, ",", ".")) Then
                Dim dDT As Double
                dDT = g_oCalc.Dist2D(oPt, oTag.Origin)
                If dDT < dMinDist Then
                    dMinDist = dDT

                    ' Un tag est toujours attache a un element hote (BaseElement).
                    ' Si l'hote est une cellule, on reroute le hit vers le cas
                    ' cellule (chemin ClonerCellule, qui fonctionne) au lieu de
                    ' cloner un tag orphelin. On protege l'acces a BaseElement
                    ' car un tag reellement orphelin peut lever une erreur.
                    Dim oBase As Element
                    Set oBase = Nothing
                    On Error Resume Next
                    Set oBase = oTag.BaseElement
                    On Error GoTo 0

                    If Not oBase Is Nothing Then
                        If oBase.Type = msdElementTypeCellHeader Then
                            ' Cas cellule : la creation passera par ClonerCellule
                            Set oNearest = Nothing
                            Set oTagNearest = Nothing
                            Set oCellNearest = oBase
                            sTagDef = oTag.TagDefinitionName
                            ptOrigine = oTag.Origin
                            sValeur = sTagVal
                            GoTo SuivantElem
                        End If
                    End If

                    ' Cas tag isole (hote non-cellule ou tag orphelin)
                    Set oNearest = Nothing
                    Set oTagNearest = oTag
                    Set oCellNearest = Nothing
                    sTagDef = ""
                    ptOrigine = oTag.Origin
                    sValeur = sTagVal
                End If
            End If

        ElseIf oElem.IsTextElement Then
            Dim oTxt As TextElement
            Set oTxt = oElem
            If g_oCalc.EstNombre(Replace(Trim$(oTxt.Text), ",", ".")) Then
                Dim dD As Double
                dD = g_oCalc.Dist2D(oPt, oTxt.Origin)
                If dD < dMinDist Then
                    dMinDist = dD
                    Set oNearest = oTxt
                    Set oTagNearest = Nothing
                    Set oCellNearest = Nothing
                    ptOrigine = oTxt.Origin
                    sValeur = Trim$(oTxt.Text)
                End If
            End If

        ElseIf oElem.Type = msdElementTypeCellHeader Then
            Dim oCell As CellElement
            Set oCell = oElem
            Dim sCellVal As String
            Dim sCellTagDef As String
            Dim oTxtCell As TextElement
            If ExtraireAltitudeDeCellule(oCell, sCellVal, sCellTagDef, oTxtCell) Then
                Dim dDC As Double
                dDC = g_oCalc.Dist2D(oPt, oCell.Origin)
                If dDC < dMinDist Then
                    dMinDist = dDC
                    Set oNearest = oTxtCell
                    Set oTagNearest = Nothing
                    Set oCellNearest = oCell
                    ptOrigine = oCell.Origin
                    sValeur = sCellVal
                    sTagDef = sCellTagDef
                End If
            End If
        End If
SuivantElem:
    Loop

    g_ptOrigineTexte = ptOrigine
    g_sValeurTexte = sValeur
    Set g_oTagTrouve = oTagNearest
    Set g_oCellTrouvee = oCellNearest
    g_sTagDefName = sTagDef
    Set TrouverTexteProche = oNearest
End Function

'------------------------------------------------------------------------------
' Cherche l'altitude la plus proche (texte, cellule ou tag).
' Renvoie True si trouvee. Les resultats sont dans g_ptOrigineTexte,
' g_sValeurTexte, et oTextOut (Nothing si c'est un tag).
Function TrouverAltitudeProche(oPt As Point3d, dRayon As Double, _
                                oTextOut As TextElement) As Boolean
    Set oTextOut = TrouverTexteProche(oPt, dRayon)
    TrouverAltitudeProche = (Len(g_sValeurTexte) > 0)
End Function

'------------------------------------------------------------------------------
' Teste si l'element est sur un niveau gele. Renvoie False en cas d'erreur
' (element sans niveau valide) pour ne pas bloquer le scan.
Private Function EstSurNiveauGele(oElem As Element) As Boolean
    On Error GoTo Securite
    Dim oLvl As Level
    Set oLvl = oElem.Level
    EstSurNiveauGele = Not oLvl.IsDisplayedInView(ActiveDesignFile.Views(1))
    Exit Function
Securite:
    EstSurNiveauGele = False
End Function

'------------------------------------------------------------------------------
' Parcourt les sous-elements d'une cellule et cherche une altitude numerique.
' Cherche d'abord un tag, puis un texte. Renvoie la valeur trouvee dans sVal,
' le nom de definition du tag dans sDefName (vide si c'est un texte),
' et le TextElement si c'est un texte (Nothing si c'est un tag).
Private Function ExtraireAltitudeDeCellule(oCell As CellElement, _
        sVal As String, sDefName As String, oTxtOut As TextElement) As Boolean
    ExtraireAltitudeDeCellule = False
    sDefName = ""
    Set oTxtOut = Nothing

    Dim oSubEnum As ElementEnumerator
    Set oSubEnum = oCell.GetSubElements
    Do While oSubEnum.MoveNext
        If EstSurNiveauGele(oSubEnum.Current) Then GoTo SuivantSub

        ' Chercher un tag numerique
        If oSubEnum.Current.Type = msdElementTypeTag Then
            Dim oTag As TagElement
            Set oTag = oSubEnum.Current
            Dim sTV As String
            sTV = Trim$(CStr(oTag.Value))
            If g_oCalc.EstNombre(Replace(sTV, ",", ".")) Then
                sVal = sTV
                sDefName = oTag.TagDefinitionName
                ExtraireAltitudeDeCellule = True
                Exit Function
            End If
        End If

        ' Chercher un texte numerique
        If oSubEnum.Current.IsTextElement Then
            Dim oT As TextElement
            Set oT = oSubEnum.Current
            If g_oCalc.EstNombre(Replace(Trim$(oT.Text), ",", ".")) Then
                sVal = Trim$(oT.Text)
                Set oTxtOut = oT
                ExtraireAltitudeDeCellule = True
                Exit Function
            End If
        End If
SuivantSub:
    Loop
End Function
