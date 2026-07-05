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

    ' Texte : meme origin que le cercle, rotation du texte-modele
    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    oRot = oSettings.oTextModele.Rotation
    On Error GoTo 0

    Dim oTxt As TextElement
    Set oTxt = CreateTextElement1(oSettings.oTextModele, sAlt, oProj, oRot)
    ' Attributs personnalises du texte (case decochee dans le formulaire)
    If Not oSettings.oTexte.CommeModele Then
        oTxt.Color = oSettings.oTexte.Couleur
        Set oTxt.Level = ResolverNiveau(oSettings.oTexte.NomNiveau, oSettings)
    End If
    ActiveModelReference.AddElement oTxt

    ' Cercle centre sur oProj
    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oProj, _
                      oSettings.oCercle.Rayon, oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = oSettings.oCercle.Couleur
    Set oCercle.Level = ResolverNiveau(oSettings.oCercle.NomNiveau, oSettings)
    ActiveModelReference.AddElement oCercle
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
