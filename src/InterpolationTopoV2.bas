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
'     CPointRef.cls, CAltitudeSelection.cls, CInterpolation.cls, CMoteurGraphique.cls
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
Public g_oSelectionCourante As CAltitudeSelection ' dernier texte/tag/cellule trouve
Public g_oSelectionP1       As CAltitudeSelection ' source modele de creation
Public g_oSelectionP2       As CAltitudeSelection ' source altitude P2

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
    Set g_oSelectionCourante = New CAltitudeSelection
    Set g_oSelectionP1 = New CAltitudeSelection
    Set g_oSelectionP2 = New CAltitudeSelection

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

    Dim dMinDist As Double: dMinDist = dRayon
    Dim oBest As CAltitudeSelection
    Set oBest = New CAltitudeSelection

    ScannerAltitudeDansModele ActiveModelReference, Nothing, oScan, oPt, dMinDist, oBest

    Dim oAtt As Object
    On Error Resume Next
    For Each oAtt In ActiveModelReference.Attachments
        If AttachmentAffiche(oAtt) Then
            ScannerAltitudeDansModele oAtt, oAtt, oScan, oPt, dMinDist, oBest
        End If
    Next
    On Error GoTo 0

    MemoriserSelectionCourante oBest
    Set TrouverTexteProche = oBest.oTexte
End Function

'------------------------------------------------------------------------------
' Scanne un modele ou une reference attachee. Les points stockes dans oBest sont
' toujours exprimes dans le repere du modele actif.
Private Sub ScannerAltitudeDansModele(oModel As Object, oAttachment As Object, _
        oScan As ElementScanCriteria, oPtClic As Point3d, _
        dMinDist As Double, oBest As CAltitudeSelection)

    Dim oEnum As ElementEnumerator
    On Error Resume Next
    Set oEnum = oModel.Scan(oScan)
    If Err.Number <> 0 Or oEnum Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Do While oEnum.MoveNext
        Dim oElem As Element
        Set oElem = oEnum.Current

        If EstSurNiveauGele(oElem) Then GoTo SuivantElem

        If oElem.Type = msdElementTypeTag Then
            TraiterTagCandidate oElem, oAttachment, oPtClic, dMinDist, oBest

        ElseIf oElem.IsTextElement Then
            TraiterTexteCandidate oElem, oAttachment, oPtClic, dMinDist, oBest

        ElseIf oElem.Type = msdElementTypeCellHeader Then
            TraiterCelluleCandidate oElem, oAttachment, oPtClic, dMinDist, oBest
        End If
SuivantElem:
    Loop
End Sub

'------------------------------------------------------------------------------
Private Sub TraiterTagCandidate(oElem As Element, oAttachment As Object, _
        oPtClic As Point3d, dMinDist As Double, oBest As CAltitudeSelection)

    Dim oTag As TagElement
    Set oTag = oElem

    Dim sTagVal As String
    sTagVal = Trim$(CStr(oTag.Value))
    If Not g_oCalc.EstNombre(Replace(sTagVal, ",", ".")) Then Exit Sub

    Dim ptMaster As Point3d
    TransformerPointVersMaitre oAttachment, oTag.Origin, ptMaster

    Dim dD As Double
    dD = g_oCalc.Dist2D(oPtClic, ptMaster)
    If dD >= dMinDist Then Exit Sub

    dMinDist = dD
    oBest.Vider
    oBest.ValeurTexte = sTagVal
    oBest.DefinirOrigine ptMaster
    oBest.EstReference = Not (oAttachment Is Nothing)
    Set oBest.oAttachment = oAttachment

    Dim oBase As Element
    Set oBase = Nothing
    On Error Resume Next
    Set oBase = oTag.BaseElement
    On Error GoTo 0

    If Not oBase Is Nothing Then
        If oBase.Type = msdElementTypeCellHeader Then
            Set oBest.oCellule = oBase
            oBest.TagDefName = oTag.TagDefinitionName
            Exit Sub
        End If
    End If

    Set oBest.oTag = oTag
End Sub

'------------------------------------------------------------------------------
Private Sub TraiterTexteCandidate(oElem As Element, oAttachment As Object, _
        oPtClic As Point3d, dMinDist As Double, oBest As CAltitudeSelection)

    Dim oTxt As TextElement
    Set oTxt = oElem

    Dim sTxtVal As String
    sTxtVal = Trim$(oTxt.Text)
    If Not g_oCalc.EstNombre(Replace(sTxtVal, ",", ".")) Then Exit Sub

    Dim ptMaster As Point3d
    TransformerPointVersMaitre oAttachment, oTxt.Origin, ptMaster

    Dim dD As Double
    dD = g_oCalc.Dist2D(oPtClic, ptMaster)
    If dD >= dMinDist Then Exit Sub

    dMinDist = dD
    oBest.Vider
    oBest.ValeurTexte = sTxtVal
    oBest.DefinirOrigine ptMaster
    oBest.EstReference = Not (oAttachment Is Nothing)
    Set oBest.oAttachment = oAttachment
    Set oBest.oTexte = oTxt
End Sub

'------------------------------------------------------------------------------
Private Sub TraiterCelluleCandidate(oElem As Element, oAttachment As Object, _
        oPtClic As Point3d, dMinDist As Double, oBest As CAltitudeSelection)

    Dim oCell As CellElement
    Set oCell = oElem

    Dim sCellVal As String
    Dim sCellTagDef As String
    Dim oTxtCell As TextElement
    If Not ExtraireAltitudeDeCellule(oCell, sCellVal, sCellTagDef, oTxtCell) Then Exit Sub

    Dim ptMaster As Point3d
    TransformerPointVersMaitre oAttachment, oCell.Origin, ptMaster

    Dim dD As Double
    dD = g_oCalc.Dist2D(oPtClic, ptMaster)
    If dD >= dMinDist Then Exit Sub

    dMinDist = dD
    oBest.Vider
    oBest.ValeurTexte = sCellVal
    oBest.TagDefName = sCellTagDef
    oBest.DefinirOrigine ptMaster
    oBest.EstReference = Not (oAttachment Is Nothing)
    Set oBest.oAttachment = oAttachment
    Set oBest.oCellule = oCell
    Set oBest.oTexte = oTxtCell
End Sub

'------------------------------------------------------------------------------
Private Sub MemoriserSelectionCourante(oSel As CAltitudeSelection)
    If g_oSelectionCourante Is Nothing Then Set g_oSelectionCourante = New CAltitudeSelection
    g_oSelectionCourante.CopierDepuis oSel

    g_oSelectionCourante.CopierOrigine g_ptOrigineTexte
    g_sValeurTexte = g_oSelectionCourante.ValeurTexte
    Set g_oTagTrouve = g_oSelectionCourante.oTag
    Set g_oCellTrouvee = g_oSelectionCourante.oCellule
    g_sTagDefName = g_oSelectionCourante.TagDefName
End Sub

'------------------------------------------------------------------------------
Private Function AttachmentAffiche(oAttachment As Object) As Boolean
    AttachmentAffiche = True
    On Error Resume Next
    AttachmentAffiche = CBool(oAttachment.IsDisplayed)
    If Err.Number <> 0 Then
        Err.Clear
        AttachmentAffiche = True
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Point de passage unique pour convertir une origine issue d'une reference vers
' le repere du modele actif. En V8i, selon le type de scan/reference, l'origine
' arrive deja dans le repere visible. On garde donc une version sans appel API
' incertain afin de rester compilable en VBA 6.x ; si une reference deplacee ou
' tournee ne tombe pas juste, c'est ici qu'il faudra brancher la transformation
' Attachment -> master disponible dans l'Object Browser de ton installation.
Private Sub TransformerPointVersMaitre(oAttachment As Object, _
        oPtSource As Point3d, oPtMaster As Point3d)

    oPtMaster = oPtSource
End Sub

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
