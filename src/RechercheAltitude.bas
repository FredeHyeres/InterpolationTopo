Attribute VB_Name = "RechercheAltitude"
'==============================================================================
' RechercheAltitude - Moteur de recherche des altitudes sources
'
' Detecte le texte, tag ou cellule numerique le plus proche d'un clic, dans le
' modele actif et dans les references attachees (coordonnees converties dans le
' repere du modele actif). Le resultat complet est memorise dans
' g_oSelectionCourante (voir InterpolationTopoV2.bas pour les globals).
'
' Fonctions publiques :
'   TrouverAltitudeProche : point d'entree utilise par les classes de commande
'   TrouverTexteProche    : variante renvoyant le TextElement source
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Cherche l'altitude la plus proche (texte, cellule ou tag).
' Renvoie True si trouvee. Le resultat complet est dans g_oSelectionCourante,
' et oTextOut recoit le TextElement source (Nothing si c'est un tag).
Function TrouverAltitudeProche(oPt As Point3d, dRayon As Double, _
                                oTextOut As TextElement) As Boolean
    Set oTextOut = TrouverTexteProche(oPt, dRayon)
    TrouverAltitudeProche = g_oSelectionCourante.Trouvee
End Function

'------------------------------------------------------------------------------
' Cherche le texte numerique le plus proche du clic dans le rayon donne.
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

    If g_oSelectionCourante Is Nothing Then Set g_oSelectionCourante = New CAltitudeSelection
    g_oSelectionCourante.CopierDepuis oBest
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
    Dim ptOrigineAltitude As Point3d
    If Not ExtraireAltitudeDeCellule(oCell, sCellVal, sCellTagDef, _
                                     oTxtCell, ptOrigineAltitude) Then Exit Sub

    Dim ptMaster As Point3d
    TransformerPointVersMaitre oAttachment, ptOrigineAltitude, ptMaster

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
' Convertit une origine issue d'une reference vers le repere du modele actif.
' Sans cette conversion, le clic est compare a des coordonnees locales de la
' reference : on detecte alors un texte numerique proche dans le mauvais repere.
Private Sub TransformerPointVersMaitre(oAttachment As Object, _
        oPtSource As Point3d, oPtMaster As Point3d)

    oPtMaster = oPtSource
    If oAttachment Is Nothing Then Exit Sub

    Dim ptRefOrigin As Point3d
    Dim ptMasterOrigin As Point3d
    Dim dScale As Double

    dScale = 1#
    ptRefOrigin.X = 0#: ptRefOrigin.Y = 0#: ptRefOrigin.Z = 0#
    ptMasterOrigin.X = 0#: ptMasterOrigin.Y = 0#: ptMasterOrigin.Z = 0#

    ' Propriete des attachements V8i : point origine dans la reference.
    On Error Resume Next
    ptRefOrigin = oAttachment.ReferenceOrigin
    If Err.Number <> 0 Then
        Err.Clear
        ptRefOrigin.X = 0#: ptRefOrigin.Y = 0#: ptRefOrigin.Z = 0#
    End If

    ' Propriete des attachements V8i : point correspondant dans le modele actif.
    ptMasterOrigin = oAttachment.MasterOrigin
    If Err.Number <> 0 Then
        Err.Clear
        ptMasterOrigin.X = 0#: ptMasterOrigin.Y = 0#: ptMasterOrigin.Z = 0#
    End If

    dScale = CDbl(oAttachment.ScaleFactor)
    If Err.Number <> 0 Or Abs(dScale) < 0.0000000001 Then
        Err.Clear
        dScale = 1#
    End If
    On Error GoTo 0

    oPtMaster.X = ptMasterOrigin.X + (oPtSource.X - ptRefOrigin.X) * dScale
    oPtMaster.Y = ptMasterOrigin.Y + (oPtSource.Y - ptRefOrigin.Y) * dScale
    oPtMaster.Z = ptMasterOrigin.Z + (oPtSource.Z - ptRefOrigin.Z) * dScale
End Sub

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
        sVal As String, sDefName As String, oTxtOut As TextElement, _
        ptOrigineOut As Point3d) As Boolean
    ExtraireAltitudeDeCellule = False
    sDefName = ""
    Set oTxtOut = Nothing
    ptOrigineOut = oCell.Origin

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
                ptOrigineOut = oTag.Origin
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
                ptOrigineOut = oT.Origin
                ExtraireAltitudeDeCellule = True
                Exit Function
            End If
        End If
SuivantSub:
    Loop
End Function
