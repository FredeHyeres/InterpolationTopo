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

    If Not g_oCellTrouvee Is Nothing Then
        ' --- Cas cellule : cloner, deplacer, modifier le texte interne ---
        ClonerCellule oProj, sAlt, oCalc

    ElseIf Not g_oTagTrouve Is Nothing Then
        ' --- Cas tag : cloner, deplacer, modifier la valeur ---
        ClonerTag oProj, sAlt

    Else
        ' --- Cas texte classique : texte + cercle ---
        CreerTexteEtCercle oProj, sAlt, oSettings, oCalc
    End If
End Sub

'------------------------------------------------------------------------------
Private Sub CreerTexteEtCercle(oProj As Point3d, sAlt As String, _
                                oSettings As CMstSettings, oCalc As CInterpolation)
    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    oRot = oSettings.oTextModele.Rotation
    On Error GoTo 0

    Dim oTxt As TextElement
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oProj, oRot)
    If Not oSettings.oTexte.CommeModele Then
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
Private Sub ClonerCellule(oProj As Point3d, sAlt As String, oCalc As CInterpolation)
    ' Copier la cellule source dans le modele
    Dim oCopy As CellElement
    Set oCopy = ActiveModelReference.CopyElement(g_oCellTrouvee)

    ' Deplacer vers le point projete
    Dim oVec As Point3d
    oVec.X = oProj.X - g_oCellTrouvee.Origin.X
    oVec.Y = oProj.Y - g_oCellTrouvee.Origin.Y
    oVec.Z = 0
    oCopy.Move oVec
    oCopy.Rewrite

    Dim bModifie As Boolean
    bModifie = False

    ' D'abord chercher un tag altitude (cas tag dans cellule)
    If Len(g_sTagDefName) > 0 Then
        Dim oSub1 As ElementEnumerator
        Set oSub1 = oCopy.GetSubElements
        Do While oSub1.MoveNext
            If oSub1.Current.Type = msdElementTypeTag Then
                Dim oTag As TagElement
                Set oTag = oSub1.Current
                If oTag.TagDefinitionName = g_sTagDefName Then
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
Private Sub ClonerTag(oProj As Point3d, sAlt As String)
    ' Chercher l'element parent du tag (cellule ou element complexe)
    Dim oParent As Element
    On Error Resume Next
    Dim lID As DLong
    lID = g_oTagTrouve.ParentID
    Set oParent = ActiveModelReference.GetElementByID(lID)
    On Error GoTo 0

    If Not oParent Is Nothing Then
        ' Copier l'element parent entier (cellule + tag)
        Dim oCopyParent As Element
        Set oCopyParent = ActiveModelReference.CopyElement(oParent)

        ' Deplacer vers le point projete (depuis l'origine du tag source)
        Dim oVec As Point3d
        oVec.X = oProj.X - g_oTagTrouve.Origin.X
        oVec.Y = oProj.Y - g_oTagTrouve.Origin.Y
        oVec.Z = 0
        oCopyParent.Move oVec
        oCopyParent.Rewrite

        ' Trouver le tag a l'interieur et modifier sa valeur
        If oCopyParent.IsComplexElement Then
            Dim oSub As ElementEnumerator
            Set oSub = oCopyParent.AsComplexElement.GetSubElements
            Do While oSub.MoveNext
                If oSub.Current.Type = msdElementTypeTag Then
                    Dim oTagSub As TagElement
                    Set oTagSub = oSub.Current
                    If oTagSub.TagDefinitionName = g_oTagTrouve.TagDefinitionName Then
                        oTagSub.Value = sAlt
                        oTagSub.Rewrite
                        Exit Do
                    End If
                End If
            Loop
        End If
    Else
        ' Pas de parent trouve : copier le tag seul (fallback)
        Dim oCopy As TagElement
        Set oCopy = ActiveModelReference.CopyElement(g_oTagTrouve)
        Dim oV As Point3d
        oV.X = oProj.X - g_oTagTrouve.Origin.X
        oV.Y = oProj.Y - g_oTagTrouve.Origin.Y
        oV.Z = 0
        oCopy.Move oV
        oCopy.Value = sAlt
        oCopy.Rewrite
    End If
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
    Dim oPt1 As Point3d, oPt2 As Point3d
    oP1.CopierPosition oPt1
    oP2.CopierPosition oPt2

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
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oProj, oRot)
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
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oPtDecale, oRot)
    If Not oSettings.oTexte.CommeModele Then oTxt.Color = oSettings.oTexte.Couleur
    oTxt.Redraw DrawMode
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
    Set ResolverNiveau = oSettings.oTextModele.Level
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
