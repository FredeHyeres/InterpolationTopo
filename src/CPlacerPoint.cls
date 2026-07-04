VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlacerPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPlacerPoint - etape 3 : placement dynamique du point interpole sur P1-P2
' Data = creer le point (repetable), Reset = terminer
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt InfoSegment() & "  |  Deplacez le curseur : Data=creer  Reset=nouv. selection"
    ' Indispensable : active l'evenement Dynamics (apercu provisoire)
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    Dim dT As Double, oProj As Point3d
    ProjectSurDroite g_oP1, g_oP2, Point, dT, oProj

    Dim sAlt As String
    sAlt = FormatAltitude(InterpolerZ(g_dZ1, g_dZ2, dT), g_nDecimals, g_sDecSep)

    ' Ligne support P1-P2, etendue de 20 % de part et d'autre
    Dim dx As Double, dy As Double
    dx = g_oP2.X - g_oP1.X: dy = g_oP2.Y - g_oP1.Y
    Dim oA As Point3d, oB As Point3d
    oA.X = g_oP1.X - dx * 0.2: oA.Y = g_oP1.Y - dy * 0.2: oA.Z = g_oP1.Z
    oB.X = g_oP2.X + dx * 0.2: oB.Y = g_oP2.Y + dy * 0.2: oB.Z = g_oP2.Z
    Dim oLigne As LineElement
    Set oLigne = CreateLineElement2(Nothing, oA, oB)
    oLigne.Color = 3
    oLigne.Redraw DrawMode

    ' Croix sur P1 et P2
    DrawCroix DrawMode, g_oP1, g_dCircRayon * 0.8
    DrawCroix DrawMode, g_oP2, g_dCircRayon * 0.8

    ' Cercle au point projete
    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oProj, _
                       g_dCircRayon, g_dCircRayon, Matrix3dIdentity())
    oCercle.Color = g_nCircColor
    oCercle.Redraw DrawMode

    ' Texte altitude (meme position que le texte qui sera cree)
    Dim oPosTxt As Point3d, oRotTxt As Matrix3d
    PositionTexteAltitude oProj, oPosTxt
    RotationTexteModele oRotTxt
    Dim oTxt As TextElement
    Set oTxt = CreateTextElement1(g_oTextTemplate, sAlt, oPosTxt, oRotTxt)
    oTxt.Redraw DrawMode

    If dT < 0# Or dT > 1# Then
        ShowPrompt "Z=" & sAlt & "  t=" & Format$(dT, "0.000") & "  " & InfoSegment() & _
                   "  ATTENTION : hors segment P1-P2 (extrapolation)"
    Else
        ShowPrompt "Z=" & sAlt & "  t=" & Format$(dT, "0.000") & "  " & InfoSegment() & _
                   "  (Data=creer  Reset=nouv. selection)"
    End If
End Sub

Private Sub DrawCroix(ByVal DrawMode As MsdDrawingMode, oPt As Point3d, dTaille As Double)
    Dim oA As Point3d, oB As Point3d, oL As LineElement

    oA.X = oPt.X - dTaille: oA.Y = oPt.Y: oA.Z = oPt.Z
    oB.X = oPt.X + dTaille: oB.Y = oPt.Y: oB.Z = oPt.Z
    Set oL = CreateLineElement2(Nothing, oA, oB)
    oL.Color = 3
    oL.Redraw DrawMode

    oA.X = oPt.X: oA.Y = oPt.Y - dTaille
    oB.X = oPt.X: oB.Y = oPt.Y + dTaille
    Set oL = CreateLineElement2(Nothing, oA, oB)
    oL.Color = 3
    oL.Redraw DrawMode
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim dT As Double, oProj As Point3d
    ProjectSurDroite g_oP1, g_oP2, Point, dT, oProj
    CreateTopoPoint oProj, InterpolerZ(g_dZ1, g_dZ2, dT)
    ' La commande reste active : on peut creer d'autres points sur la meme droite
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset = nouvelle selection de points (memes parametres de cercle).
' Un second Reset, une fois revenu a la selection de P1, quitte la commande.
Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Nouvelle selection : cliquez pres du texte altitude du POINT 1  (Reset=quitter)"
    CommandState.StartPrimitive New CSelectP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
