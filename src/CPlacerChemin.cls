VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlacerChemin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPlacerChemin - Interpolation le long d'un chemin defini par l'utilisateur
'
' Apres selection et snap de P1 et P2, l'utilisateur construit un chemin :
'   Data  = ajouter un point intermediaire au chemin
'   Reset = valider le chemin et creer les elements topo
'           (si aucun point : retour a la selection de P1)
'
' Chemin final : P1 -> points intermediaires -> P2
' Altitude de chaque point interpolee lineairement :
'   Z(Pi) = Z(P1) + (dist_cumulee(Pi) / dist_totale) * (Z(P2) - Z(P1))
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private m_dPtsX() As Double
Private m_dPtsY() As Double
Private m_nPts As Long

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt "CHEMIN : cliquez pour placer le 1er point intermediaire  (Reset=nouvelle selection)"
    m_nPts = 0
    ReDim m_dPtsX(0 To 0)
    ReDim m_dPtsY(0 To 0)
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    Dim oPt1 As Point3d, oPt2 As Point3d
    Dim dTailleCroix As Double
    Dim dTotalDist As Double, dCursorDist As Double
    Dim dPrevX As Double, dPrevY As Double, dPrevZ As Double
    Dim dCumDist As Double, dSegDist As Double
    Dim dT As Double, dZi As Double, dTCursor As Double, dZCursor As Double
    Dim oPtA As Point3d, oPtB As Point3d, oPtC As Point3d, oPtCur As Point3d
    Dim oLigne As LineElement
    Dim oCercle As EllipseElement
    Dim sAltI As String, sAltCur As String
    Dim i As Long

    g_oP1.CopierPosition oPt1
    g_oP2.CopierPosition oPt2
    dTailleCroix = g_oSettings.oCercle.Rayon * 0.8

    ' --- Distance totale du chemin P1 -> pts intermediaires -> curseur -> P2 ---
    dPrevX = oPt1.X: dPrevY = oPt1.Y
    dTotalDist = 0

    For i = 0 To m_nPts - 1
        dTotalDist = dTotalDist + DistXY(dPrevX, dPrevY, m_dPtsX(i), m_dPtsY(i))
        dPrevX = m_dPtsX(i): dPrevY = m_dPtsY(i)
    Next i

    dCursorDist = dTotalDist + DistXY(dPrevX, dPrevY, Point.X, Point.Y)
    dTotalDist = dCursorDist + DistXY(Point.X, Point.Y, oPt2.X, oPt2.Y)

    ' --- Segments places (vert) avec croix et texte altitude ---
    dPrevX = oPt1.X: dPrevY = oPt1.Y: dPrevZ = oPt1.Z
    dCumDist = 0

    For i = 0 To m_nPts - 1
        dSegDist = DistXY(dPrevX, dPrevY, m_dPtsX(i), m_dPtsY(i))
        dCumDist = dCumDist + dSegDist

        If dTotalDist > 0.0001 Then dT = dCumDist / dTotalDist Else dT = 0
        dZi = g_oCalc.InterpolerZ(g_oP1.Altitude, g_oP2.Altitude, dT)
        If g_oSettings.bDecalageFixe Then dZi = dZi + g_oSettings.dDecalageDZ

        oPtA.X = dPrevX: oPtA.Y = dPrevY: oPtA.Z = dPrevZ
        oPtB.X = m_dPtsX(i): oPtB.Y = m_dPtsY(i): oPtB.Z = dZi

        Set oLigne = CreateLineElement2(Nothing, oPtA, oPtB)
        oLigne.Color = 2
        oLigne.Redraw DrawMode

        DessinCroixLocal DrawMode, oPtB, dTailleCroix

        sAltI = g_oCalc.FormatAltitude(dZi, g_oSettings.oTexte.Decimales, _
                                        g_oSettings.oTexte.SepDecimal)
        DessinTexteProvLocal DrawMode, oPtB, sAltI

        dPrevX = m_dPtsX(i): dPrevY = m_dPtsY(i): dPrevZ = dZi
    Next i

    ' --- Z interpole au curseur ---
    If dTotalDist > 0.0001 Then dTCursor = dCursorDist / dTotalDist Else dTCursor = 0
    dZCursor = g_oCalc.InterpolerZ(g_oP1.Altitude, g_oP2.Altitude, dTCursor)
    If g_oSettings.bDecalageFixe Then dZCursor = dZCursor + g_oSettings.dDecalageDZ

    ' --- Segment dynamique : dernier point -> curseur (jaune) ---
    oPtC.X = dPrevX: oPtC.Y = dPrevY: oPtC.Z = dPrevZ
    oPtCur.X = Point.X: oPtCur.Y = Point.Y: oPtCur.Z = dZCursor

    Set oLigne = CreateLineElement2(Nothing, oPtC, oPtCur)
    oLigne.Color = 4
    oLigne.Redraw DrawMode

    ' --- Segment dynamique : curseur -> P2 (jaune) ---
    Set oLigne = CreateLineElement2(Nothing, oPtCur, oPt2)
    oLigne.Color = 4
    oLigne.Redraw DrawMode

    ' --- Croix P1 et P2 ---
    DessinCroixLocal DrawMode, oPt1, dTailleCroix
    DessinCroixLocal DrawMode, oPt2, dTailleCroix

    ' --- Cercle + texte au curseur ---
    Set oCercle = CreateEllipseElement2(Nothing, oPtCur, _
                      g_oSettings.oCercle.Rayon, g_oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = g_oSettings.oCercle.Couleur
    oCercle.Redraw DrawMode

    sAltCur = g_oCalc.FormatAltitude(dZCursor, g_oSettings.oTexte.Decimales, _
                                      g_oSettings.oTexte.SepDecimal)
    DessinTexteProvLocal DrawMode, oPtCur, sAltCur

    ' --- Barre d'etat ---
    If m_nPts = 0 Then
        ShowPrompt "CHEMIN : Z=" & sAltCur & _
                   "  Data=placer 1er point  Reset=nouvelle selection"
    Else
        ShowPrompt "CHEMIN : Z=" & sAltCur & _
                   "  Dist=" & Format$(dCursorDist, "0.00") & "/" & _
                   Format$(dTotalDist, "0.00") & _
                   "  (" & m_nPts & " pts)  Data=ajouter  Reset=valider"
    End If
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    If m_nPts > 0 Then
        ReDim Preserve m_dPtsX(0 To m_nPts)
        ReDim Preserve m_dPtsY(0 To m_nPts)
    End If
    m_dPtsX(m_nPts) = Point.X
    m_dPtsY(m_nPts) = Point.Y
    m_nPts = m_nPts + 1

    ShowPrompt "CHEMIN : point " & m_nPts & " ajoute  |  " & _
               "Data=ajouter  Reset=valider le chemin"
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    If m_nPts = 0 Then
        frmInterpolation.ReinitialiserEtat
        ShowPrompt "Nouvelle selection : cliquez pres du texte altitude du POINT 1  (Reset=quitter)"
        CommandState.StartPrimitive New CSelectP1
        Exit Sub
    End If

    Dim oPt1 As Point3d, oPt2 As Point3d
    Dim dTotalDist As Double
    Dim dPrevX As Double, dPrevY As Double
    Dim dCumDist As Double
    Dim dT As Double, dZ As Double
    Dim oPt As Point3d
    Dim i As Long
    Dim nCreated As Long

    g_oP1.CopierPosition oPt1
    g_oP2.CopierPosition oPt2

    ' Distance totale : P1 -> points -> P2
    dPrevX = oPt1.X: dPrevY = oPt1.Y
    dTotalDist = 0

    For i = 0 To m_nPts - 1
        dTotalDist = dTotalDist + DistXY(dPrevX, dPrevY, m_dPtsX(i), m_dPtsY(i))
        dPrevX = m_dPtsX(i): dPrevY = m_dPtsY(i)
    Next i
    dTotalDist = dTotalDist + DistXY(dPrevX, dPrevY, oPt2.X, oPt2.Y)

    If dTotalDist < 0.0001 Then
        ShowPrompt "Chemin trop court - ajoutez des points plus eloignes"
        Exit Sub
    End If

    ' Creation des elements a chaque point intermediaire
    dPrevX = oPt1.X: dPrevY = oPt1.Y
    dCumDist = 0

    For i = 0 To m_nPts - 1
        dCumDist = dCumDist + DistXY(dPrevX, dPrevY, m_dPtsX(i), m_dPtsY(i))

        dT = dCumDist / dTotalDist
        dZ = g_oCalc.InterpolerZ(g_oP1.Altitude, g_oP2.Altitude, dT)
        If g_oSettings.bDecalageFixe Then dZ = dZ + g_oSettings.dDecalageDZ

        oPt.X = m_dPtsX(i)
        oPt.Y = m_dPtsY(i)
        oPt.Z = dZ

        g_oMoteur.CreerPointTopo oPt, dZ, g_oSettings, g_oCalc

        dPrevX = m_dPtsX(i): dPrevY = m_dPtsY(i)
    Next i

    nCreated = m_nPts

    ' Reinitialiser pour un eventuel nouveau chemin
    m_nPts = 0
    ReDim m_dPtsX(0 To 0)
    ReDim m_dPtsY(0 To 0)

    ShowPrompt "Chemin valide : " & nCreated & " points crees  |  " & _
               "Data=nouveau chemin  Reset=nouvelle selection"
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'==============================================================================
' Utilitaires prives
'==============================================================================

Private Function DistXY(dX1 As Double, dY1 As Double, _
                         dX2 As Double, dY2 As Double) As Double
    DistXY = Sqr((dX2 - dX1) ^ 2 + (dY2 - dY1) ^ 2)
End Function

Private Sub DessinCroixLocal(ByVal DrawMode As MsdDrawingMode, _
                              oPt As Point3d, dTaille As Double)
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

Private Sub DessinTexteProvLocal(ByVal DrawMode As MsdDrawingMode, _
                                  oPt As Point3d, sAlt As String)
    Dim oRot As Matrix3d
    Dim oTxt As TextElement

    oRot = Matrix3dIdentity()
    On Error Resume Next
    If Not g_oSettings.oTextModele Is Nothing Then oRot = g_oSettings.oTextModele.Rotation
    On Error GoTo 0

    On Error Resume Next
    Set oTxt = CreateTextElement1(g_oSettings.oTextModele, sAlt, oPt, oRot)
    If oTxt Is Nothing Then
        Err.Clear
        Set oTxt = CreateTextElement1(Nothing, sAlt, oPt, oRot)
    End If
    On Error GoTo 0
    If Not g_oSettings.oTexte.CommeModele Then oTxt.Color = g_oSettings.oTexte.Couleur
    oTxt.Redraw DrawMode
End Sub
