VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMoteurGraphique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CMoteurGraphique - Creation des elements MicroStation
'
' Responsabilites :
'   - creer le texte altitude et le cercle pour un point topo valide
'   - dessiner l'apercu dynamique (Redraw) pendant le deplacement curseur
'
' Recoit un resultat deja calcule (point projete + altitude) et des parametres
' de symbologie (CMstSettings). Ne realise aucun calcul.
'==============================================================================
Option Explicit

'==============================================================================
' Creation definitive
'==============================================================================

'------------------------------------------------------------------------------
' Cree le texte altitude et le cercle au point projete, et les ajoute au modele.
Sub CreerPointTopo(oProj As Point3d, dZInterp As Double, oSettings As CMstSettings, _
                   oCalc As CInterpolation)
    Dim sAlt As String
    sAlt = oCalc.FormatAltitude(dZInterp, oSettings.oTexte.Decimales, _
                                oSettings.oTexte.SepDecimal)

    Dim oSrc As CAltitudeSelection
    Set oSrc = g_oSelectionP1

    If oSrc Is Nothing Then
        CreerTexteEtCercle oProj, sAlt, oSettings, oCalc
        Exit Sub
    End If

    If oSrc.EstReference Then
        ' On lit les references, mais on cree toujours dans le fichier actif.
        ' Cloner cellule/tag depuis une reference V8i est fragile : fallback sur
        ' texte + cercle, en gardant le format/symbologie de P1 si disponible.
        CreerTexteEtCercle oProj, sAlt, oSettings, oCalc

    ElseIf Not oSrc.oCellule Is Nothing Then
        ' --- Cas cellule : cloner, deplacer, modifier le texte interne ---
        ClonerCellule oSrc, oProj, sAlt, oCalc

    ElseIf Not oSrc.oTag Is Nothing Then
        ' --- Cas tag : cloner l'hote et reassocier un tag ---
        ClonerTag oSrc, oProj, sAlt, oSettings, oCalc

    Else
        ' --- Cas texte classique : texte + cercle ---
        CreerTexteEtCercle oProj, sAlt, oSettings, oCalc
    End If
End Sub

'------------------------------------------------------------------------------
' Cree uniquement le cercle au point projete.
Sub CreerCercleSeul(oProj As Point3d, oSettings As CMstSettings)
    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oProj, _
                      oSettings.oCercle.Rayon, oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = oSettings.oCercle.Couleur
    Set oCercle.Level = ResolverNiveau(oSettings.oCercle.NomNiveau, oSettings)
    ActiveModelReference.AddElement oCercle
End Sub

'------------------------------------------------------------------------------
' Cree uniquement le texte altitude a la position choisie par l'utilisateur.
Sub CreerTexteSeul(oPtTexte As Point3d, dZ As Double, _
                   oSettings As CMstSettings, oCalc As CInterpolation)
    Dim sAlt As String
    sAlt = oCalc.FormatAltitude(dZ, oSettings.oTexte.Decimales, oSettings.oTexte.SepDecimal)

    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    If Not oSettings.oTextModele Is Nothing Then oRot = oSettings.oTextModele.Rotation
    On Error GoTo 0

    Dim oTxt As TextElement
    Dim bSansModele As Boolean
    On Error Resume Next
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oPtTexte, oRot)
    If oTxt Is Nothing Then
        Err.Clear
        bSansModele = True
        Set oTxt = CreateTextElement1(Nothing, sAlt, oPtTexte, oRot)
    End If
    On Error GoTo 0
    If bSansModele Or oSettings.oTextModele Is Nothing Or Not oSettings.oTexte.CommeModele Then
        oTxt.Color = oSettings.oTexte.Couleur
        Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)
    End If
    ActiveModelReference.AddElement oTxt
End Sub

'------------------------------------------------------------------------------
' Apercu dynamique du texte altitude en placement libre (etape 3).
Sub DessinTexteLibre(ByVal DrawMode As MsdDrawingMode, _
                     oPtTexte As Point3d, dZ As Double, _
                     oSettings As CMstSettings, oCalc As CInterpolation)
    Dim sAlt As String
    sAlt = oCalc.FormatAltitude(dZ, oSettings.oTexte.Decimales, oSettings.oTexte.SepDecimal)

    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    If Not oSettings.oTextModele Is Nothing Then oRot = oSettings.oTextModele.Rotation
    On Error GoTo 0

    Dim oTxt As TextElement
    Dim bSansModele As Boolean
    On Error Resume Next
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oPtTexte, oRot)
    If oTxt Is Nothing Then
        Err.Clear
        bSansModele = True
        Set oTxt = CreateTextElement1(Nothing, sAlt, oPtTexte, oRot)
    End If
    On Error GoTo 0
    If bSansModele Or Not oSettings.oTexte.CommeModele Then oTxt.Color = oSettings.oTexte.Couleur
    oTxt.Redraw DrawMode
End Sub

'------------------------------------------------------------------------------
Private Sub CreerTexteEtCercle(oProj As Point3d, sAlt As String, _
                                oSettings As CMstSettings, oCalc As CInterpolation)
    ' oTextModele peut etre Nothing (cas ou P1 etait un tag) :
    ' CreateTextElement1 accepte Nothing comme modele (attributs actifs)
    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    If Not oSettings.oTextModele Is Nothing Then oRot = oSettings.oTextModele.Rotation
    On Error GoTo 0

    Dim oTxt As TextElement
    Dim bSansModele As Boolean
    On Error Resume Next
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oProj, oRot)
    If oTxt Is Nothing Then
        Err.Clear
        bSansModele = True
        Set oTxt = CreateTextElement1(Nothing, sAlt, oProj, oRot)
    End If
    On Error GoTo 0
    If bSansModele Or oSettings.oTextModele Is Nothing Or Not oSettings.oTexte.CommeModele Then
        oTxt.Color = oSettings.oTexte.Couleur
        Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)
    End If
    ActiveModelReference.AddElement oTxt

    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oProj, _
                      oSettings.oCercle.Rayon, oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = oSettings.oCercle.Couleur
    Set oCercle.Level = ResolverNiveau(oSettings.oCercle.NomNiveau, oSettings)
    ActiveModelReference.AddElement oCercle
End Sub

'------------------------------------------------------------------------------
Private Sub ClonerCellule(oSrc As CAltitudeSelection, oProj As Point3d, _
                          sAlt As String, oCalc As CInterpolation)
    ' Copier la cellule source dans le modele
    Dim oCopy As CellElement
    Set oCopy = ActiveModelReference.CopyElement(oSrc.oCellule)

    ' Deplacer vers le point projete
    Dim oVec As Point3d
    oVec.X = oProj.X - oSrc.oCellule.Origin.X
    oVec.Y = oProj.Y - oSrc.oCellule.Origin.Y
    oVec.Z = 0
    oCopy.Move oVec
    oCopy.Rewrite

    Dim bModifie As Boolean
    bModifie = False

    ' D'abord chercher un tag altitude (cas tag dans cellule)
    If Len(oSrc.TagDefName) > 0 Then
        Dim oSub1 As ElementEnumerator
        Set oSub1 = oCopy.GetSubElements
        Do While oSub1.MoveNext
            If oSub1.Current.Type = msdElementTypeTag Then
                Dim oTag As TagElement
                Set oTag = oSub1.Current
                If oTag.TagDefinitionName = oSrc.TagDefName Then
                    oTag.Value = sAlt
                    oTag.Rewrite
                    bModifie = True
                    Exit Do
                End If
            End If
        Loop
    End If

    ' Sinon chercher un texte altitude (cas cellule simple)
    If Not bModifie Then
        Dim oSub2 As ElementEnumerator
        Set oSub2 = oCopy.GetSubElements
        Do While oSub2.MoveNext
            If oSub2.Current.IsTextElement Then
                Dim oT As TextElement
                Set oT = oSub2.Current
                If oCalc.EstNombre(Replace(Trim$(oT.Text), ",", ".")) Then
                    oT.Text = sAlt
                    oT.Rewrite
                    Exit Do
                End If
            End If
        Loop
    End If
End Sub

'------------------------------------------------------------------------------
Private Sub ClonerTag(oSrc As CAltitudeSelection, oProj As Point3d, sAlt As String, _
                      oSettings As CMstSettings, oCalc As CInterpolation)
    ' Un tag MicroStation est toujours attache a un element hote via son
    ' BaseElement. Cloner le tag seul produit un orphelin (affiche en pointille).
    ' On copie donc l'hote, on le deplace, puis on reassocie un tag a la copie.
    ' Si le tag est reellement orphelin (pas d'hote), on cree un texte + cercle.

    ' 1) Recuperer l'hote via BaseElement (protege : peut lever ou etre Nothing)
    Dim oHost As Element
    Set oHost = Nothing
    On Error Resume Next
    Set oHost = oSrc.oTag.BaseElement
    On Error GoTo 0

    If oHost Is Nothing Then
        ' 3) Aucun hote : ne pas cloner un orphelin, fallback texte + cercle
        CreerTexteEtCercle oProj, sAlt, oSettings, oCalc
        Exit Sub
    End If

    ' Vecteur de deplacement (origine du tag source -> point projete)
    Dim oVec As Point3d
    oVec.X = oProj.X - oSrc.oTag.Origin.X
    oVec.Y = oProj.Y - oSrc.oTag.Origin.Y
    oVec.Z = 0

    ' 2) Copier l'hote et le deplacer
    Dim oCopyHost As Element
    Set oCopyHost = ActiveModelReference.CopyElement(oHost)
    oCopyHost.Move oVec
    oCopyHost.Rewrite

    ' 2a) Approche A : copier le tag source, le deplacer, le reassocier a la copie
    Dim bOK As Boolean
    bOK = False
    On Error Resume Next
    Dim oCopyTag As TagElement
    Set oCopyTag = ActiveModelReference.CopyElement(oSrc.oTag)
    oCopyTag.Move oVec
    Set oCopyTag.BaseElement = oCopyHost
    oCopyTag.Value = sAlt
    oCopyTag.Rewrite
    If Err.Number = 0 And Not oCopyTag Is Nothing Then bOK = True
    On Error GoTo 0
    If bOK Then Exit Sub

    ' Approche A partielle : supprimer le tag copie (potentiellement orphelin)
    ' avant d'essayer l'approche B, pour ne pas laisser d'element parasite.
    On Error Resume Next
    If Not oCopyTag Is Nothing Then ActiveModelReference.RemoveElement oCopyTag
    On Error GoTo 0

    ' 2b) Approche B : recreer le tag via AddTag sur l'hote copie
    On Error Resume Next
    Dim oTagDef As TagDefinition
    Set oTagDef = TrouverTagDefinition(oSrc.oTag.TagDefinitionName)
    Dim oNewTag As TagElement
    If Not oTagDef Is Nothing Then
        Set oNewTag = oCopyHost.AddTag(oTagDef)
    End If
    If Not oNewTag Is Nothing Then
        ' Positionner le tag a l'emplacement du tag source deplace
        Dim oOrig As Point3d
        oOrig = oSrc.oTag.Origin
        Dim oVecTag As Point3d
        oVecTag.X = (oOrig.X + oVec.X) - oNewTag.Origin.X
        oVecTag.Y = (oOrig.Y + oVec.Y) - oNewTag.Origin.Y
        oVecTag.Z = 0
        oNewTag.Move oVecTag

        ' Reprendre la symbologie et les attributs texte du tag source
        ' (AddTag cree le tag avec le style par defaut de la definition)
        CopierAttributsTag oSrc.oTag, oNewTag

        ' Seules la valeur et l'ecriture conditionnent la reussite : un echec
        ' de copie d'attribut ne doit pas declencher le fallback texte+cercle.
        Err.Clear
        oNewTag.Value = sAlt
        oNewTag.Rewrite
        bOK = (Err.Number = 0)
        If Not bOK Then ActiveModelReference.RemoveElement oNewTag
    End If
    On Error GoTo 0
    If bOK Then Exit Sub

    ' Dernier recours : texte + cercle pour ne pas laisser l'utilisateur sans point
    CreerTexteEtCercle oProj, sAlt, oSettings, oCalc
End Sub

'------------------------------------------------------------------------------
' Cherche une TagDefinition par nom en parcourant les TagSets du fichier actif.
' Renvoie Nothing si introuvable.
Private Function TrouverTagDefinition(sDefName As String) As TagDefinition
    Set TrouverTagDefinition = Nothing
    On Error Resume Next
    Dim oSet As TagSet
    Dim oDef As TagDefinition
    For Each oSet In ActiveDesignFile.TagSets
        For Each oDef In oSet.TagDefinitions
            If oDef.Name = sDefName Then
                Set TrouverTagDefinition = oDef
                Exit Function
            End If
        Next
    Next
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Copie la symbologie et les attributs texte (police, taille, justification)
' d'un tag source vers un tag cible. Les proprietes texte sont copiees en
' liaison tardive (Object) : selon la version de l'API certaines ne sont pas
' exposees par TagElement, et un acces en dur ne compilerait pas. L'echec
' d'une copie n'empeche pas les suivantes (Err.Clear apres chaque tentative).
Private Sub CopierAttributsTag(oSrc As TagElement, oDst As TagElement)
    On Error Resume Next

    ' Proprietes sures (exposees par TagElement)
    oDst.Color = oSrc.Color: Err.Clear
    Set oDst.Level = oSrc.Level: Err.Clear
    oDst.LineWeight = oSrc.LineWeight: Err.Clear
    ' oDst.Rotation = oSrc.Rotation: Err.Clear

    ' Proprietes texte incertaines : liaison tardive
    Dim oS As Object, oD As Object
    Set oS = oSrc
    Set oD = oDst
    Set oD.TextStyle = oS.TextStyle: Err.Clear
    Set oD.Font = oS.Font: Err.Clear
    oD.Height = oS.Height: Err.Clear
    oD.Width = oS.Width: Err.Clear
    oD.TextSize = oS.TextSize: Err.Clear
    oD.Justification = oS.Justification: Err.Clear

    On Error GoTo 0
End Sub

'==============================================================================
' Apercu dynamique (appele dans IPrimitiveCommandEvents_Dynamics)
'==============================================================================

'------------------------------------------------------------------------------
' Dessine l'apercu : ligne support, croix P1/P2, cercle et texte provisoires.
Sub DessinDynamique(ByVal DrawMode As MsdDrawingMode, _
                    oP1 As CPointRef, oP2 As CPointRef, _
                    oProj As Point3d, dT As Double, _
                    oSettings As CMstSettings, oCalc As CInterpolation)

    Dim dZInterp As Double
    dZInterp = oCalc.InterpolerZ(oP1.Altitude, oP2.Altitude, dT)

    ' Positions des points de reference en Point3d locaux
    ' Utilise Altitude plutot que Z de position (Z=0 si reference 2D)
    Dim oPt1 As Point3d, oPt2 As Point3d
    oP1.CopierPosition oPt1
    oP2.CopierPosition oPt2
    oPt1.Z = oP1.Altitude
    oPt2.Z = oP2.Altitude

    ' Ligne support P1-P2 etendue de 20% de chaque cote
    Dim dx As Double, dy As Double
    dx = oPt2.X - oPt1.X
    dy = oPt2.Y - oPt1.Y
    Dim oA As Point3d, oB As Point3d
    oA.X = oPt1.X - dx * 0.2: oA.Y = oPt1.Y - dy * 0.2: oA.Z = oPt1.Z
    oB.X = oPt2.X + dx * 0.2: oB.Y = oPt2.Y + dy * 0.2: oB.Z = oPt2.Z
    Dim oLigne As LineElement
    Set oLigne = CreateLineElement2(Nothing, oA, oB)
    oLigne.Color = 3
    oLigne.Redraw DrawMode

    ' Croix sur P1 et P2
    Dim dTailleCroix As Double
    dTailleCroix = oSettings.oCercle.Rayon * 0.8
    DessinCroix DrawMode, oPt1, dTailleCroix
    DessinCroix DrawMode, oPt2, dTailleCroix

    ' Perpendiculaire au point projete
    Dim dLen As Double
    dLen = Sqr(dx * dx + dy * dy)
    If dLen > 0.0000001 Then
        Dim dPerpLen As Double
        dPerpLen = dLen * 0.5
        Dim dPx As Double, dPy As Double
        dPx = -dy / dLen: dPy = dx / dLen
        Dim oPerpA As Point3d, oPerpB As Point3d
        oPerpA.X = oProj.X - dPerpLen * dPx: oPerpA.Y = oProj.Y - dPerpLen * dPy: oPerpA.Z = oProj.Z
        oPerpB.X = oProj.X + dPerpLen * dPx: oPerpB.Y = oProj.Y + dPerpLen * dPy: oPerpB.Z = oProj.Z
        Dim oLPerp As LineElement
        Set oLPerp = CreateLineElement2(Nothing, oPerpA, oPerpB)
        oLPerp.Color = 2
        oLPerp.Redraw DrawMode
    End If

    ' Cercle provisoire au point projete
    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oProj, _
                      oSettings.oCercle.Rayon, oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = oSettings.oCercle.Couleur
    oCercle.Redraw DrawMode

    ' Texte altitude provisoire (meme position que la creation definitive)
    Dim sAlt As String
    sAlt = oCalc.FormatAltitude(dZInterp, oSettings.oTexte.Decimales, _
                                oSettings.oTexte.SepDecimal)
    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    oRot = oSettings.oTextModele.Rotation
    On Error GoTo 0
    Dim oTxt As TextElement
    On Error Resume Next
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oProj, oRot)
    If oTxt Is Nothing Then
        Err.Clear
        Set oTxt = CreateTextElement1(Nothing, sAlt, oProj, oRot)
    End If
    On Error GoTo 0
    If Not oSettings.oTexte.CommeModele Then oTxt.Color = oSettings.oTexte.Couleur
    oTxt.Redraw DrawMode
End Sub

'------------------------------------------------------------------------------
' Dessine l'apercu pour l'etape 2 : ligne P1-P2, croix, point projete fixe,
' ligne perpendiculaire et cercle+texte a la position decalee.
Sub DessinDecalage(ByVal DrawMode As MsdDrawingMode, _
                   oP1 As CPointRef, oP2 As CPointRef, _
                   oProj As Point3d, oPtDecale As Point3d, _
                   dZInterp As Double, _
                   oSettings As CMstSettings, oCalc As CInterpolation)

    Dim oPt1 As Point3d, oPt2 As Point3d
    oP1.CopierPosition oPt1
    oP2.CopierPosition oPt2
    oPt1.Z = oP1.Altitude
    oPt2.Z = oP2.Altitude

    ' Ligne support P1-P2
    Dim dx As Double, dy As Double
    dx = oPt2.X - oPt1.X: dy = oPt2.Y - oPt1.Y
    Dim oA As Point3d, oB As Point3d
    oA.X = oPt1.X - dx * 0.2: oA.Y = oPt1.Y - dy * 0.2: oA.Z = oPt1.Z
    oB.X = oPt2.X + dx * 0.2: oB.Y = oPt2.Y + dy * 0.2: oB.Z = oPt2.Z
    Dim oLigne As LineElement
    Set oLigne = CreateLineElement2(Nothing, oA, oB)
    oLigne.Color = 3
    oLigne.Redraw DrawMode

    ' Croix sur P1, P2 et point projete
    Dim dTailleCroix As Double
    dTailleCroix = oSettings.oCercle.Rayon * 0.8
    DessinCroix DrawMode, oPt1, dTailleCroix
    DessinCroix DrawMode, oPt2, dTailleCroix
    DessinCroix DrawMode, oProj, dTailleCroix

    ' Ligne perpendiculaire depuis le point projete vers le point decale
    Dim oLPerp As LineElement
    Set oLPerp = CreateLineElement2(Nothing, oProj, oPtDecale)
    oLPerp.Color = 2
    oLPerp.Redraw DrawMode

    ' Cercle provisoire au point decale
    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oPtDecale, _
                      oSettings.oCercle.Rayon, oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = oSettings.oCercle.Couleur
    oCercle.Redraw DrawMode

    ' Texte altitude provisoire au point decale
    Dim sAlt As String
    sAlt = oCalc.FormatAltitude(dZInterp, oSettings.oTexte.Decimales, _
                                oSettings.oTexte.SepDecimal)
    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    oRot = oSettings.oTextModele.Rotation
    On Error GoTo 0
    Dim oTxt As TextElement
    On Error Resume Next
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oPtDecale, oRot)
    If oTxt Is Nothing Then
        Err.Clear
        Set oTxt = CreateTextElement1(Nothing, sAlt, oPtDecale, oRot)
    End If
    On Error GoTo 0
    If Not oSettings.oTexte.CommeModele Then oTxt.Color = oSettings.oTexte.Couleur
    oTxt.Redraw DrawMode
End Sub

'==============================================================================
' Indicateur de pente (texte + fleche)
'==============================================================================

'------------------------------------------------------------------------------
' Apercu dynamique du texte pente + fleche au curseur.
Sub DessinPenteDynamique(ByVal DrawMode As MsdDrawingMode, _
                         oPtPlace As Point3d, _
                         oP1 As CPointRef, oP2 As CPointRef, _
                         oSettings As CMstSettings, oCalc As CInterpolation)

    Dim dPente As Double
    dPente = oCalc.PentePct(oP1, oP2)

    Dim sPente As String
    sPente = oCalc.FormatPente(dPente, oSettings.oTexte.SepDecimal)

    Dim dDesX As Double, dDesY As Double
    DirectionP1P2 oP1, oP2, dDesX, dDesY

    ' Angle aligne sur la fleche (direction descente), ajuste pour lisibilite
    Dim dAngleTxt As Double
    dAngleTxt = oCalc.Angle2D(dDesX, dDesY)
    Const PI_D As Double = 3.14159265358979
    If dAngleTxt > PI_D / 2# Then dAngleTxt = dAngleTxt - PI_D
    If dAngleTxt < -PI_D / 2# Then dAngleTxt = dAngleTxt + PI_D

    Dim dTH As Double
    dTH = HauteurTexteModele(oSettings)

    Dim dPerpX As Double, dPerpY As Double
    dPerpX = -Sin(dAngleTxt): dPerpY = Cos(dAngleTxt)

    Dim oPtTxt As Point3d
    oPtTxt.X = oPtPlace.X + dPerpX * dTH * 0.8
    oPtTxt.Y = oPtPlace.Y + dPerpY * dTH * 0.8
    oPtTxt.Z = oPtPlace.Z

    Dim oRot As Matrix3d
    oRot = Matrix3dFromAxisAndRotationAngle(2, dAngleTxt)

    Dim oTxt As TextElement
    On Error Resume Next
    If oSettings.bPentePerso Then
        Set oTxt = CreateTextElement1(Nothing, sPente, oPtTxt, oRot)
        Dim oTSd As TextStyle: Set oTSd = oTxt.TextStyle
        oTSd.Height = oSettings.dPenteHauteur
        Dim dWD As Double: dWD = oSettings.dPenteLargeur
        If dWD > 0 Then oTSd.Width = dWD Else oTSd.Width = oSettings.dPenteHauteur
        oTSd.Justification = msdTextJustificationLeftBottom
        oTxt.TextStyle = oTSd
        oTxt.Color = oSettings.nPenteCouleur
    Else
        Set oTxt = CreateTextElement1(oSettings.oTextModele, sPente, oPtTxt, oRot)
        If oTxt Is Nothing Then
            Err.Clear
            Set oTxt = CreateTextElement1(Nothing, sPente, oPtTxt, oRot)
        End If
        Dim oTSm As TextStyle: Set oTSm = oTxt.TextStyle
        oTSm.Justification = msdTextJustificationLeftBottom
        oTxt.TextStyle = oTSm
        If Not oSettings.oTexte.CommeModele Then
            oTxt.Color = oSettings.oTexte.Couleur
        End If
    End If
    oTxt.Rotation = oRot
    Err.Clear
    On Error GoTo 0
    oTxt.Redraw DrawMode

    ' Fleche
    Dim oPtDeb As Point3d, oPtFin As Point3d
    Dim oPtH1 As Point3d, oPtH2 As Point3d
    GeometrieFleche oPtPlace, dDesX, dDesY, oSettings.dFlecheLongueur, _
                    oPtDeb, oPtFin, oPtH1, oPtH2

    Dim oL As LineElement
    Set oL = CreateLineElement2(Nothing, oPtDeb, oPtFin)
    oL.Color = oSettings.nFlecheCouleur
    oL.Redraw DrawMode

    Dim oLH1 As LineElement, oLH2 As LineElement
    Set oLH1 = CreateLineElement2(Nothing, oPtFin, oPtH1)
    oLH1.Color = oSettings.nFlecheCouleur
    oLH1.Redraw DrawMode
    Set oLH2 = CreateLineElement2(Nothing, oPtFin, oPtH2)
    oLH2.Color = oSettings.nFlecheCouleur
    oLH2.Redraw DrawMode
End Sub

'------------------------------------------------------------------------------
' Creation definitive du texte pente + fleche.
Sub CreerPente(oPtPlace As Point3d, _
               oP1 As CPointRef, oP2 As CPointRef, _
               oSettings As CMstSettings, oCalc As CInterpolation)

    Dim dPente As Double
    dPente = oCalc.PentePct(oP1, oP2)

    Dim sPente As String
    sPente = oCalc.FormatPente(dPente, oSettings.oTexte.SepDecimal)

    Dim dDesX As Double, dDesY As Double
    DirectionP1P2 oP1, oP2, dDesX, dDesY

    Dim dAngleTxt As Double
    dAngleTxt = oCalc.Angle2D(dDesX, dDesY)
    Const PI_C As Double = 3.14159265358979
    If dAngleTxt > PI_C / 2# Then dAngleTxt = dAngleTxt - PI_C
    If dAngleTxt < -PI_C / 2# Then dAngleTxt = dAngleTxt + PI_C

    Dim dTH As Double
    dTH = HauteurTexteModele(oSettings)

    Dim dPerpX As Double, dPerpY As Double
    dPerpX = -Sin(dAngleTxt): dPerpY = Cos(dAngleTxt)

    Dim oPtTxt As Point3d
    oPtTxt.X = oPtPlace.X + dPerpX * dTH * 0.8
    oPtTxt.Y = oPtPlace.Y + dPerpY * dTH * 0.8
    oPtTxt.Z = oPtPlace.Z

    Dim oRot As Matrix3d
    oRot = Matrix3dFromAxisAndRotationAngle(2, dAngleTxt)

    Dim oTxt As TextElement
    On Error Resume Next
    If oSettings.bPentePerso Then
        Set oTxt = CreateTextElement1(Nothing, sPente, oPtTxt, oRot)
        Dim oTSc As TextStyle: Set oTSc = oTxt.TextStyle
        oTSc.Height = oSettings.dPenteHauteur
        Dim dWC As Double: dWC = oSettings.dPenteLargeur
        If dWC > 0 Then oTSc.Width = dWC Else oTSc.Width = oSettings.dPenteHauteur
        oTSc.Justification = msdTextJustificationLeftBottom
        oTxt.TextStyle = oTSc
        oTxt.Color = oSettings.nPenteCouleur
        Set oTxt.Level = ResolverNiveau(oSettings.sPenteNiveau, oSettings)
    Else
        Set oTxt = CreateTextElement1(oSettings.oTextModele, sPente, oPtTxt, oRot)
        If oTxt Is Nothing Then
            Err.Clear
            Set oTxt = CreateTextElement1(Nothing, sPente, oPtTxt, oRot)
        End If
        Dim oTSp As TextStyle: Set oTSp = oTxt.TextStyle
        oTSp.Justification = msdTextJustificationLeftBottom
        oTxt.TextStyle = oTSp
        If oSettings.oTextModele Is Nothing Or Not oSettings.oTexte.CommeModele Then
            oTxt.Color = oSettings.oTexte.Couleur
            Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)
        End If
    End If
    oTxt.Rotation = oRot
    Err.Clear
    On Error GoTo 0
    ActiveModelReference.AddElement oTxt

    ' Fleche permanente
    Dim oPtDeb As Point3d, oPtFin As Point3d
    Dim oPtH1 As Point3d, oPtH2 As Point3d
    GeometrieFleche oPtPlace, dDesX, dDesY, oSettings.dFlecheLongueur, _
                    oPtDeb, oPtFin, oPtH1, oPtH2

    Dim oLvl As Level
    Set oLvl = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)

    Dim oL As LineElement
    Set oL = CreateLineElement2(Nothing, oPtDeb, oPtFin)
    oL.Color = oSettings.nFlecheCouleur
    Set oL.Level = oLvl
    ActiveModelReference.AddElement oL

    Dim oLH1 As LineElement, oLH2 As LineElement
    Set oLH1 = CreateLineElement2(Nothing, oPtFin, oPtH1)
    oLH1.Color = oSettings.nFlecheCouleur
    Set oLH1.Level = oLvl
    ActiveModelReference.AddElement oLH1
    Set oLH2 = CreateLineElement2(Nothing, oPtFin, oPtH2)
    oLH2.Color = oSettings.nFlecheCouleur
    Set oLH2.Level = oLvl
    ActiveModelReference.AddElement oLH2
End Sub

'==============================================================================
' Utilitaires prives
'==============================================================================

Private Sub DessinCroix(ByVal DrawMode As MsdDrawingMode, oPt As Point3d, dTaille As Double)
    Dim oA As Point3d, oB As Point3d, oL As LineElement

    oA.X = oPt.X - dTaille: oA.Y = oPt.Y: oA.Z = oPt.Z
    oB.X = oPt.X + dTaille: oB.Y = oPt.Y: oB.Z = oPt.Z
    Set oL = CreateLineElement2(Nothing, oA, oB)
    oL.Color = 3: oL.Redraw DrawMode

    oA.X = oPt.X: oA.Y = oPt.Y - dTaille
    oB.X = oPt.X: oB.Y = oPt.Y + dTaille
    Set oL = CreateLineElement2(Nothing, oA, oB)
    oL.Color = 3: oL.Redraw DrawMode
End Sub

'------------------------------------------------------------------------------
' Resout un niveau par nom ou numero ; chaine vide ou niveau introuvable =
' niveau du texte-modele P1. Sert au cercle comme au texte personnalise.
Private Function ResolverNiveau(sNiveau As String, oSettings As CMstSettings) As Level
    Set ResolverNiveau = ActiveSettings.Level

    If Not oSettings.oTextModele Is Nothing Then
        On Error Resume Next
        Dim sModNiv As String
        sModNiv = oSettings.oTextModele.Level.Name
        If Err.Number = 0 And Len(sModNiv) > 0 Then
            Dim oLvlMod As Level
            Set oLvlMod = ActiveDesignFile.Levels(sModNiv)
            If Err.Number = 0 And Not oLvlMod Is Nothing Then
                Set ResolverNiveau = oLvlMod
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If

    If sNiveau = "" Then Exit Function

    Dim oLvl As Level
    On Error Resume Next
    Set oLvl = ActiveDesignFile.Levels(sNiveau)
    On Error GoTo 0
    If oLvl Is Nothing And IsNumeric(sNiveau) Then
        Dim oL As Level
        For Each oL In ActiveDesignFile.Levels
            If oL.Number = CLng(Val(sNiveau)) Then
                Set oLvl = oL: Exit For
            End If
        Next
    End If
    If Not oLvl Is Nothing Then Set ResolverNiveau = oLvl
End Function

'------------------------------------------------------------------------------
' Direction unitaire de P1 vers P2.
Private Sub DirectionP1P2(oP1 As CPointRef, oP2 As CPointRef, _
                           dDirX As Double, dDirY As Double)
    Dim dx As Double, dy As Double, dLen As Double
    dx = oP2.X - oP1.X: dy = oP2.Y - oP1.Y
    dLen = Sqr(dx * dx + dy * dy)
    If dLen < 0.0000001 Then
        dDirX = 1#: dDirY = 0#
        Exit Sub
    End If
    dDirX = dx / dLen: dDirY = dy / dLen
End Sub


'------------------------------------------------------------------------------
' Calcule les 4 points d'une fleche centree sur oPtCentre.
' oPtDeb/oPtFin = extremites du corps, oPtH1/oPtH2 = pointe de fleche.
Private Sub GeometrieFleche(oPtCentre As Point3d, _
                             dDirX As Double, dDirY As Double, _
                             dLongueur As Double, _
                             oPtDeb As Point3d, oPtFin As Point3d, _
                             oPtH1 As Point3d, oPtH2 As Point3d)
    Dim dHL As Double
    dHL = dLongueur / 2#

    oPtDeb.X = oPtCentre.X - dDirX * dHL
    oPtDeb.Y = oPtCentre.Y - dDirY * dHL
    oPtDeb.Z = oPtCentre.Z
    oPtFin.X = oPtCentre.X + dDirX * dHL
    oPtFin.Y = oPtCentre.Y + dDirY * dHL
    oPtFin.Z = oPtCentre.Z

    Dim dHeadLen As Double
    dHeadLen = dLongueur * 0.2
    Const HEAD_ANGLE As Double = 0.4363

    oPtH1.X = oPtFin.X - dHeadLen * (dDirX * Cos(HEAD_ANGLE) - dDirY * Sin(HEAD_ANGLE))
    oPtH1.Y = oPtFin.Y - dHeadLen * (dDirX * Sin(HEAD_ANGLE) + dDirY * Cos(HEAD_ANGLE))
    oPtH1.Z = oPtFin.Z

    oPtH2.X = oPtFin.X - dHeadLen * (dDirX * Cos(HEAD_ANGLE) + dDirY * Sin(HEAD_ANGLE))
    oPtH2.Y = oPtFin.Y - dHeadLen * (-dDirX * Sin(HEAD_ANGLE) + dDirY * Cos(HEAD_ANGLE))
    oPtH2.Z = oPtFin.Z
End Sub


'------------------------------------------------------------------------------
' Hauteur du texte modele P1 (fallback 0.1 u.m.).
Private Function HauteurTexteModele(oSettings As CMstSettings) As Double
    If oSettings.bPentePerso Then
        HauteurTexteModele = oSettings.dPenteHauteur
    ElseIf oSettings.dModeleHauteur > 0 Then
        HauteurTexteModele = oSettings.dModeleHauteur
    Else
        HauteurTexteModele = 0.1
    End If
End Function
