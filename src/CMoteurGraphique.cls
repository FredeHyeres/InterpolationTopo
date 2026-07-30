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
    AppliquerRemplissageCercle oCercle, oSettings
    Set oCercle.Level = ResolverNiveau(oSettings.oCercle.NomNiveau, oSettings)
    ActiveModelReference.AddElement oCercle
End Sub

'------------------------------------------------------------------------------
' Applique l'attribut plein/vide du cercle (Couleur sert aussi de couleur de
' remplissage). Reutilise pour la creation definitive et l'apercu dynamique.
Public Sub AppliquerRemplissageCercle(oCercle As EllipseElement, oSettings As CMstSettings)
    On Error Resume Next
    If oSettings.oCercle.Plein Then
        oCercle.FillMode = msdFillModeFilled
        oCercle.FillColor = oSettings.oCercle.Couleur
    Else
        oCercle.FillMode = msdFillModeNotFilled
    End If
    Err.Clear
    On Error GoTo 0
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
        Dim oStyleSeul As TextStyle
        Set oStyleSeul = ResolverStyle(oSettings.oTexte.NomStyle)
        If Not oStyleSeul Is Nothing Then oTxt.TextStyle = oStyleSeul
        Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)
    Else
        On Error Resume Next
        Dim sStyleNom As String
        sStyleNom = oSettings.oTextModele.TextStyle.Name
        If Err.Number = 0 And Len(sStyleNom) > 0 Then
            Dim oStyleLocal As TextStyle
            Set oStyleLocal = ActiveDesignFile.TextStyles(sStyleNom)
            If Not oStyleLocal Is Nothing Then
                oTxt.TextStyle = oStyleLocal
            End If
        End If
        Err.Clear
        Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)
        Err.Clear
        On Error GoTo 0
    End If
    ActiveModelReference.AddElement oTxt
End Sub

'------------------------------------------------------------------------------
' Apercu dynamique du texte altitude en placement libre (etape 3).
Sub DessinTexteLibre(ByVal DrawMode As MsdDrawingMode, _
                     oPtTexte As Point3d, dZ As Double, _
                     oSettings As CMstSettings, oCalc As CInterpolation)
    Dim sAlt As String
    sAlt = oCalc.FormatAltitude(dZ, 3, oSettings.oTexte.SepDecimal)

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
    If bSansModele Or Not oSettings.oTexte.CommeModele Then
        oTxt.Color = oSettings.oTexte.Couleur
        Dim oStyleLibre As TextStyle
        Set oStyleLibre = ResolverStyle(oSettings.oTexte.NomStyle)
        If Not oStyleLibre Is Nothing Then oTxt.TextStyle = oStyleLibre
    Else
        AppliquerStyleModeleTexte oTxt, oSettings
    End If
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
        Dim oStyleEC As TextStyle
        Set oStyleEC = ResolverStyle(oSettings.oTexte.NomStyle)
        If Not oStyleEC Is Nothing Then oTxt.TextStyle = oStyleEC
        Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)
    Else
        On Error Resume Next
        Dim sStyleC As String
        sStyleC = oSettings.oTextModele.TextStyle.Name
        If Err.Number = 0 And Len(sStyleC) > 0 Then
            Dim oStyleC As TextStyle
            Set oStyleC = ActiveDesignFile.TextStyles(sStyleC)
            If Not oStyleC Is Nothing Then
                oTxt.TextStyle = oStyleC
            End If
        End If
        Err.Clear
        Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)
        Err.Clear
        On Error GoTo 0
    End If
    ActiveModelReference.AddElement oTxt

    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oProj, _
                      oSettings.oCercle.Rayon, oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = oSettings.oCercle.Couleur
    AppliquerRemplissageCercle oCercle, oSettings
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
    ' Z aussi : le clone doit etre a l'altitude interpolee, pas a celle du
    ' point d'origine (oProj.Z porte le Z calcule).
    oVec.Z = oProj.Z - oSrc.oCellule.Origin.Z
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
    ' Z aussi : l'hote clone et son tag doivent etre a l'altitude interpolee,
    ' pas a celle du point d'origine (oProj.Z porte le Z calcule).
    oVec.Z = oProj.Z - oSrc.oTag.Origin.Z

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
        oVecTag.Z = (oOrig.Z + oVec.Z) - oNewTag.Origin.Z
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
    AppliquerRemplissageCercle oCercle, oSettings
    oCercle.Redraw DrawMode

    ' Texte altitude provisoire (meme position que la creation definitive)
    Dim sAlt As String
    sAlt = oCalc.FormatAltitude(dZInterp, 3, oSettings.oTexte.SepDecimal)
    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    oRot = oSettings.oTextModele.Rotation
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
    If bSansModele Or Not oSettings.oTexte.CommeModele Then
        oTxt.Color = oSettings.oTexte.Couleur
        Dim oStyleDyn As TextStyle
        Set oStyleDyn = ResolverStyle(oSettings.oTexte.NomStyle)
        If Not oStyleDyn Is Nothing Then oTxt.TextStyle = oStyleDyn
    Else
        AppliquerStyleModeleTexte oTxt, oSettings
    End If
    oTxt.Redraw DrawMode
End Sub

'------------------------------------------------------------------------------
' Dessine l'apercu pour l'etape 2 : ligne P1-P2, croix, point projete fixe,
' ligne perpendiculaire et cercle+texte a la position decalee.
Sub DessinDecalage(ByVal DrawMode As MsdDrawingMode, _
                   oP1 As CPointRef, oP2 As CPointRef, _
                   oProj As Point3d, oPtDecale As Point3d, _
                   dZFinal As Double, dZSurDroite As Double, _
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
    AppliquerRemplissageCercle oCercle, oSettings
    oCercle.Redraw DrawMode

    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    oRot = oSettings.oTextModele.Rotation
    On Error GoTo 0

    ' Texte altitude au point decale (Z final)
    Dim sAltFinal As String
    sAltFinal = oCalc.FormatAltitude(dZFinal, 3, oSettings.oTexte.SepDecimal)
    Dim oTxt As TextElement
    Dim bSansModele As Boolean
    On Error Resume Next
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAltFinal, oPtDecale, oRot)
    If oTxt Is Nothing Then
        Err.Clear
        bSansModele = True
        Set oTxt = CreateTextElement1(Nothing, sAltFinal, oPtDecale, oRot)
    End If
    On Error GoTo 0
    If bSansModele Or Not oSettings.oTexte.CommeModele Then
        oTxt.Color = oSettings.oTexte.Couleur
        Dim oStyleDecale As TextStyle
        Set oStyleDecale = ResolverStyle(oSettings.oTexte.NomStyle)
        If Not oStyleDecale Is Nothing Then oTxt.TextStyle = oStyleDecale
    Else
        AppliquerStyleModeleTexte oTxt, oSettings
    End If
    oTxt.Redraw DrawMode

    ' Texte altitude au point projete sur la droite (Z interpole)
    Dim sAltDroite As String
    sAltDroite = oCalc.FormatAltitude(dZSurDroite, 3, oSettings.oTexte.SepDecimal)
    Dim oTxtDroite As TextElement
    Dim bSansModeleDroite As Boolean
    On Error Resume Next
    Set oTxtDroite = CreateTextElement1(oSettings.oTextModele, sAltDroite, oProj, oRot)
    If oTxtDroite Is Nothing Then
        Err.Clear
        bSansModeleDroite = True
        Set oTxtDroite = CreateTextElement1(Nothing, sAltDroite, oProj, oRot)
    End If
    On Error GoTo 0
    If bSansModeleDroite Or Not oSettings.oTexte.CommeModele Then
        oTxtDroite.Color = oSettings.oTexte.Couleur
        Dim oStyleDroite As TextStyle
        Set oStyleDroite = ResolverStyle(oSettings.oTexte.NomStyle)
        If Not oStyleDroite Is Nothing Then oTxtDroite.TextStyle = oStyleDroite
    Else
        AppliquerStyleModeleTexte oTxtDroite, oSettings
    End If
    oTxtDroite.Redraw DrawMode
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
    sPente = oCalc.FormatPente(dPente, oSettings.oTexte.SepDecimal, _
                               oSettings.oIndicPente.Decimales)

    Dim dDesX As Double, dDesY As Double
    DirectionP1P2 oP1, oP2, dDesX, dDesY

    Dim oRot As Matrix3d
    Dim oPtTxt As Point3d
    PositionnerTextePente oPtPlace, dDesX, dDesY, sPente, oSettings, oCalc, oRot, oPtTxt

    Dim oTxt As TextElement
    On Error Resume Next
    Set oTxt = CreateTextElement1(Nothing, sPente, oPtTxt, oRot)
    Dim oStylePenteD As TextStyle
    Set oStylePenteD = ResolverStyle(oSettings.oIndicPente.NomStyle)
    If Not oStylePenteD Is Nothing Then oTxt.TextStyle = oStylePenteD
    Dim oTSd As TextStyle: Set oTSd = oTxt.TextStyle
    oTSd.Justification = msdTextJustificationLeftCenter
    oTxt.TextStyle = oTSd
    oTxt.Rotation = oRot
    Err.Clear
    On Error GoTo 0
    oTxt.Redraw DrawMode

    ' Fleche (apercu : simple Redraw, pas de niveau, pas de regroupement)
    Dim oPtDeb As Point3d, oPtFin As Point3d
    Dim oPtH1 As Point3d, oPtH2 As Point3d
    GeometrieFleche oPtPlace, dDesX, dDesY, oSettings.oIndicPente.FlecheLongueur, _
                    oSettings.oIndicPente.PointeLongueur, _
                    oPtDeb, oPtFin, oPtH1, oPtH2

    Dim aFlecheD() As Element
    aFlecheD = ElementsFleche(oPtDeb, oPtFin, oPtH1, oPtH2, oSettings, Nothing)
    Dim iFD As Long
    For iFD = LBound(aFlecheD) To UBound(aFlecheD)
        aFlecheD(iFD).Redraw DrawMode
    Next iFD
End Sub

'------------------------------------------------------------------------------
' Creation definitive du texte pente + fleche.
Sub CreerPente(oPtPlace As Point3d, _
               oP1 As CPointRef, oP2 As CPointRef, _
               oSettings As CMstSettings, oCalc As CInterpolation)

    Dim dPente As Double
    dPente = oCalc.PentePct(oP1, oP2)

    Dim sPente As String
    sPente = oCalc.FormatPente(dPente, oSettings.oTexte.SepDecimal, _
                               oSettings.oIndicPente.Decimales)

    Dim dDesX As Double, dDesY As Double
    DirectionP1P2 oP1, oP2, dDesX, dDesY

    Dim oRot As Matrix3d
    Dim oPtTxt As Point3d
    PositionnerTextePente oPtPlace, dDesX, dDesY, sPente, oSettings, oCalc, oRot, oPtTxt

    Dim oTxt As TextElement
    On Error Resume Next
    Set oTxt = CreateTextElement1(Nothing, sPente, oPtTxt, oRot)
    Dim oStylePenteC As TextStyle
    Set oStylePenteC = ResolverStyle(oSettings.oIndicPente.NomStyle)
    If Not oStylePenteC Is Nothing Then oTxt.TextStyle = oStylePenteC
    Dim oTSc As TextStyle: Set oTSc = oTxt.TextStyle
    oTSc.Justification = msdTextJustificationLeftCenter
    oTxt.TextStyle = oTSc
    Set oTxt.Level = ResolverNiveau(oSettings.oIndicPente.NomNiveau, oSettings)
    oTxt.Rotation = oRot
    Err.Clear
    On Error GoTo 0
    ActiveModelReference.AddElement oTxt

    ' Fleche permanente : meme niveau que le texte pente (un seul reglage),
    ' regroupee en une seule cellule orpheline pour etre selectionnable en 1 clic.
    Dim oPtDeb As Point3d, oPtFin As Point3d
    Dim oPtH1 As Point3d, oPtH2 As Point3d
    GeometrieFleche oPtPlace, dDesX, dDesY, oSettings.oIndicPente.FlecheLongueur, _
                    oSettings.oIndicPente.PointeLongueur, _
                    oPtDeb, oPtFin, oPtH1, oPtH2

    Dim oLvl As Level
    Set oLvl = ResolverNiveau(oSettings.oIndicPente.NomNiveau, oSettings)

    Dim aFleche() As Element
    aFleche = ElementsFleche(oPtDeb, oPtFin, oPtH1, oPtH2, oSettings, oLvl)

    Dim oCellFl As CellElement
    Set oCellFl = CreateCellElement1("IndicPente", aFleche, oPtPlace, False)
    Set oCellFl.Level = oLvl
    ActiveModelReference.AddElement oCellFl
End Sub

'------------------------------------------------------------------------------
' Construit les elements de la pointe de fleche selon oSettings.oIndicPente.PointeFermee :
' False = 2 lignes ouvertes (oPtFin->oPtH1, oPtFin->oPtH2) comme avant ;
' True  = triangle plein (oPtFin, oPtH1, oPtH2), corps arrete a sa base.
' oLvl = niveau a affecter aux elements (Nothing = ne pas l'affecter : apercu dynamique).
Private Function ElementsFleche(oPtDeb As Point3d, oPtFin As Point3d, _
                                oPtH1 As Point3d, oPtH2 As Point3d, _
                                oSettings As CMstSettings, oLvl As Level) As Element()
    Dim arr() As Element

    If oSettings.oIndicPente.PointeFermee Then
        ReDim arr(0 To 1)

        Dim oPtBase As Point3d
        oPtBase.X = (oPtH1.X + oPtH2.X) / 2#
        oPtBase.Y = (oPtH1.Y + oPtH2.Y) / 2#
        oPtBase.Z = oPtFin.Z

        Dim oLC As LineElement
        Set oLC = CreateLineElement2(Nothing, oPtDeb, oPtBase)
        oLC.Color = oSettings.oIndicPente.FlecheCouleur
        If Not oLvl Is Nothing Then Set oLC.Level = oLvl
        Set arr(0) = oLC

        Dim vTri(0 To 2) As Point3d
        vTri(0) = oPtFin: vTri(1) = oPtH1: vTri(2) = oPtH2
        Dim oTri As ShapeElement
        Set oTri = CreateShapeElement1(Nothing, vTri, msdFillModeFilled)
        oTri.Color = oSettings.oIndicPente.FlecheCouleur
        On Error Resume Next
        oTri.FillColor = oSettings.oIndicPente.FlecheCouleur
        Err.Clear
        On Error GoTo 0
        If Not oLvl Is Nothing Then Set oTri.Level = oLvl
        Set arr(1) = oTri
    Else
        ReDim arr(0 To 2)

        Dim oL As LineElement
        Set oL = CreateLineElement2(Nothing, oPtDeb, oPtFin)
        oL.Color = oSettings.oIndicPente.FlecheCouleur
        If Not oLvl Is Nothing Then Set oL.Level = oLvl
        Set arr(0) = oL

        Dim oLH1 As LineElement, oLH2 As LineElement
        Set oLH1 = CreateLineElement2(Nothing, oPtFin, oPtH1)
        oLH1.Color = oSettings.oIndicPente.FlecheCouleur
        If Not oLvl Is Nothing Then Set oLH1.Level = oLvl
        Set arr(1) = oLH1

        Set oLH2 = CreateLineElement2(Nothing, oPtFin, oPtH2)
        oLH2.Color = oSettings.oIndicPente.FlecheCouleur
        If Not oLvl Is Nothing Then Set oLH2.Level = oLvl
        Set arr(2) = oLH2
    End If

    ElementsFleche = arr
End Function

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
' Reapplique le TextStyle NOMME du texte-modele P1 sur oTxt. Necessaire car
' CreateTextElement1(oTextModele, ...) ne transfere pas fidelement le style
' nomme quand oTextModele provient d'une reference attachee (fichier actif et
' fichier reference ont chacun leur propre collection TextStyles) : l'apercu
' dynamique doit refaire la meme resolution que la creation definitive
' (cf. CreerTexteEtCercle / CreerTexteSeul / CreerPente).
Private Sub AppliquerStyleModeleTexte(oTxt As TextElement, oSettings As CMstSettings)
    If oSettings.oTextModele Is Nothing Then Exit Sub
    On Error Resume Next
    Dim sStyleNom As String
    sStyleNom = oSettings.oTextModele.TextStyle.Name
    If Err.Number = 0 And Len(sStyleNom) > 0 Then
        Dim oStyleLocal As TextStyle
        Set oStyleLocal = ActiveDesignFile.TextStyles(sStyleNom)
        If Not oStyleLocal Is Nothing Then
            oTxt.TextStyle = oStyleLocal
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Resout un niveau par nom ou numero ; chaine vide ou niveau introuvable =
' niveau du texte-modele P1. Sert au cercle comme au texte personnalise.
' Le niveau source peut venir d'une reference (P1 scanne dans un plan attache) :
' s'il n'existe pas encore dans le fichier actif, il est cree (AddNewLevel),
' meme logique que NiveauOuCreation/ResolverNiveau dans Trans3D (fonction Points).
Private Function ResolverNiveau(sNiveau As String, oSettings As CMstSettings) As Level
    Set ResolverNiveau = ActiveSettings.Level

    If Not oSettings.oTextModele Is Nothing Then
        On Error Resume Next
        Dim sModNiv As String
        sModNiv = oSettings.oTextModele.Level.Name
        If Err.Number = 0 And Len(sModNiv) > 0 Then
            Dim oLvlMod As Level
            Set oLvlMod = ActiveDesignFile.Levels(sModNiv)
            If oLvlMod Is Nothing Then
                Err.Clear
                Set oLvlMod = ActiveDesignFile.AddNewLevel(sModNiv)
            End If
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
    If oLvl Is Nothing Then
        On Error Resume Next
        Set oLvl = ActiveDesignFile.AddNewLevel(sNiveau)
        Err.Clear
        On Error GoTo 0
    End If
    If Not oLvl Is Nothing Then Set ResolverNiveau = oLvl
End Function

'------------------------------------------------------------------------------
' Resout un style de texte nomme dans le fichier actif. Chaine vide ou style
' introuvable = Nothing (l'appelant garde alors le style actif/herite en place).
Private Function ResolverStyle(sNomStyle As String) As TextStyle
    Set ResolverStyle = Nothing
    If Len(Trim$(sNomStyle)) = 0 Then Exit Function
    On Error Resume Next
    Set ResolverStyle = ActiveDesignFile.TextStyles(sNomStyle)
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Calcule la rotation lisible du texte pente et sa position, TOUJOURS en
' retrait derriere la fleche (jamais de chevauchement, quel que soit le sens
' de descente) : le texte est justifie LeftCenter (centre vertical MicroStation
' natif), pose sur la meme ligne que la fleche, decale d'un recul qui inclut
' la moitie de la fleche + un espace + la longueur du texte si son sens de
' lecture (apres redressement pour rester lisible) coincide avec le sens de
' descente -- sinon le texte se lit deja "vers l'arriere" et n'a pas besoin de
' ce recul supplementaire.
Private Sub PositionnerTextePente(oPtPlace As Point3d, dDesX As Double, dDesY As Double, _
                                  sPente As String, oSettings As CMstSettings, _
                                  oCalc As CInterpolation, _
                                  oRot As Matrix3d, oPtTxt As Point3d)
    Const PI_T As Double = 3.14159265358979

    Dim dAngleTxt As Double, bFlip As Boolean
    dAngleTxt = oCalc.Angle2D(dDesX, dDesY)
    bFlip = False
    If dAngleTxt > PI_T / 2# Then
        dAngleTxt = dAngleTxt - PI_T
        bFlip = True
    ElseIf dAngleTxt < -PI_T / 2# Then
        dAngleTxt = dAngleTxt + PI_T
        bFlip = True
    End If
    oRot = Matrix3dFromAxisAndRotationAngle(2, dAngleTxt)

    ' Hauteur/largeur de caractere du style choisi ; repli sur le modele P1
    ' si aucun style nomme n'est selectionne (cf. HauteurTexteModele).
    Dim dTH As Double, dCharW As Double
    Dim oStyleInfo As TextStyle
    Set oStyleInfo = ResolverStyle(oSettings.oIndicPente.NomStyle)
    If Not oStyleInfo Is Nothing Then
        dTH = oStyleInfo.Height
        dCharW = oStyleInfo.Width
    Else
        dTH = HauteurTexteModele(oSettings)
    End If
    If dTH <= 0 Then dTH = 0.1
    If dCharW <= 0 Then dCharW = dTH * 0.6

    Dim dHalfFl As Double: dHalfFl = oSettings.oIndicPente.FlecheLongueur / 2#
    Dim dGap As Double: dGap = dTH * 0.5
    Dim dTxtLen As Double: dTxtLen = dCharW * Len(sPente)

    Dim dRecul As Double
    dRecul = dHalfFl + dGap
    If Not bFlip Then dRecul = dRecul + dTxtLen

    oPtTxt.X = oPtPlace.X - dDesX * dRecul
    oPtTxt.Y = oPtPlace.Y - dDesY * dRecul
    oPtTxt.Z = oPtPlace.Z
End Sub

'------------------------------------------------------------------------------
' Direction unitaire de P1 vers P2.
Private Sub DirectionP1P2(oP1 As CPointRef, oP2 As CPointRef, _
                           dDirX As Double, dDirY As Double)
    Dim dx As Double, dy As Double, dLen As Double
    ' Fleche dans le sens de la descente (du plus haut vers le plus bas)
    If oP1.Altitude >= oP2.Altitude Then
        dx = oP2.X - oP1.X: dy = oP2.Y - oP1.Y
    Else
        dx = oP1.X - oP2.X: dy = oP1.Y - oP2.Y
    End If
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
' dHeadLen = longueur des 2 lignes de la pointe (reglage independant, PointeLongueur).
Private Sub GeometrieFleche(oPtCentre As Point3d, _
                             dDirX As Double, dDirY As Double, _
                             dLongueur As Double, dHeadLen As Double, _
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
    If oSettings.dModeleHauteur > 0 Then
        HauteurTexteModele = oSettings.dModeleHauteur
    Else
        HauteurTexteModele = 0.1
    End If
End Function

