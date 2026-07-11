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
' CPlacerPoint V2 - Etat commande : placement dynamique du point interpole
'
' Etape 1 : projection sur P1-P2 + decalage perpendiculaire simultane
'           => Data = creer le cercle (et texte si cellule/tag)
' Etape 2 : placement libre du texte altitude => Data = creer le texte
'   (cellule/tag : etape 2 sautee, tout est cree en etape 1)
' Reset = revenir a la selection de P1
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private m_nEtape        As Integer      ' 1=projection+perpendiculaire, 2=texte libre
Private m_dZFinal       As Double       ' altitude finale pour le texte (etape 2)

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt g_oCalc.InfoSegment(g_oP1, g_oP2) & _
               "  |  Deplacez le curseur : Data=placer le point  Reset=nouv. selection"
    m_nEtape = 1
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    Select Case m_nEtape
    Case 1
        ' --- Projection sur P1-P2 + decalage perpendiculaire simultane ---
        Dim dT As Double, oProj As Point3d
        g_oCalc.ProjecterSurSegment g_oP1, g_oP2, Point, dT, oProj

        Dim dZInterp As Double
        dZInterp = g_oCalc.InterpolerZ(g_oP1.Altitude, g_oP2.Altitude, dT)

        Dim dx As Double, dy As Double, dLen As Double
        dx = g_oP2.X - g_oP1.X
        dy = g_oP2.Y - g_oP1.Y
        dLen = Sqr(dx * dx + dy * dy)
        Dim dPerpX As Double, dPerpY As Double
        If dLen > 0.0000001 Then
            dPerpX = -dy / dLen
            dPerpY = dx / dLen
        Else
            dPerpX = 0: dPerpY = 1
        End If

        Dim dOffset As Double
        dOffset = (Point.X - oProj.X) * dPerpX + (Point.Y - oProj.Y) * dPerpY

        Dim oPtDecale As Point3d
        oPtDecale.X = oProj.X + dOffset * dPerpX
        oPtDecale.Y = oProj.Y + dOffset * dPerpY

        Dim dZFinal As Double
        dZFinal = dZInterp
        If g_oSettings.oDecalage.PenteActive Then
            dZFinal = dZFinal + Abs(dOffset) * g_oSettings.oDecalage.Pente / 100#
        End If
        If g_oSettings.oDecalage.DZActive Then dZFinal = dZFinal + g_oSettings.oDecalage.DZ
        oPtDecale.Z = dZFinal

        g_oMoteur.DessinDecalage DrawMode, g_oP1, g_oP2, oProj, oPtDecale, _
                                  dZFinal, dZInterp, g_oSettings, g_oCalc

        Dim sZDroite As String, sZFinal As String
        sZDroite = g_oCalc.FormatAltitude(dZInterp, 3, g_oSettings.oTexte.SepDecimal)
        sZFinal = g_oCalc.FormatAltitude(dZFinal, 3, g_oSettings.oTexte.SepDecimal)

        Dim sPrompt As String
        If Abs(dOffset) < 0.001 Then
            sPrompt = "Z=" & sZDroite & "  t=" & Format$(dT, "0.000")
        Else
            sPrompt = "Z(droite)=" & sZDroite & "  Z(final)=" & sZFinal & _
                       "  dec=" & Format$(dOffset, "0.000") & _
                       "  t=" & Format$(dT, "0.000")
        End If

        If dT < 0# Or dT > 1# Then
            sPrompt = sPrompt & "  HORS SEGMENT"
        End If
        sPrompt = sPrompt & "  (Data=placer  Reset=nouv. selection)"
        ShowPrompt sPrompt

    Case 2
        ' --- Etape 2 : placement libre du texte ---
        g_oMoteur.DessinTexteLibre DrawMode, Point, m_dZFinal, g_oSettings, g_oCalc

        Dim sAlt3 As String
        sAlt3 = g_oCalc.FormatAltitude(m_dZFinal, 3, g_oSettings.oTexte.SepDecimal)
        ShowPrompt "Placez le texte Z=" & sAlt3 & _
                   "  (Data=poser le texte  Reset=nouv. selection)"
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Select Case m_nEtape
    Case 1
        ' --- Creer le point a la position calculee ---
        Dim dT As Double, oProj As Point3d
        g_oCalc.ProjecterSurSegment g_oP1, g_oP2, Point, dT, oProj

        Dim dZInterp As Double
        dZInterp = g_oCalc.InterpolerZ(g_oP1.Altitude, g_oP2.Altitude, dT)

        Dim dx As Double, dy As Double, dLen As Double
        dx = g_oP2.X - g_oP1.X
        dy = g_oP2.Y - g_oP1.Y
        dLen = Sqr(dx * dx + dy * dy)
        Dim dPerpX As Double, dPerpY As Double
        If dLen > 0.0000001 Then
            dPerpX = -dy / dLen
            dPerpY = dx / dLen
        Else
            dPerpX = 0: dPerpY = 1
        End If

        Dim dOffset As Double
        dOffset = (Point.X - oProj.X) * dPerpX + (Point.Y - oProj.Y) * dPerpY

        Dim oPtDecale As Point3d
        oPtDecale.X = oProj.X + dOffset * dPerpX
        oPtDecale.Y = oProj.Y + dOffset * dPerpY

        Dim dZFinal As Double
        dZFinal = dZInterp
        If g_oSettings.oDecalage.PenteActive Then
            dZFinal = dZFinal + Abs(dOffset) * g_oSettings.oDecalage.Pente / 100#
        End If
        If g_oSettings.oDecalage.DZActive Then dZFinal = dZFinal + g_oSettings.oDecalage.DZ
        oPtDecale.Z = dZFinal

        If SourceP1Integree() Then
            g_oMoteur.CreerPointTopo oPtDecale, dZFinal, g_oSettings, g_oCalc

            Dim sAlt1 As String
            sAlt1 = g_oCalc.FormatAltitude(dZFinal, g_oSettings.oTexte.Decimales, _
                                             g_oSettings.oTexte.SepDecimal)
            ShowPrompt "Point cree Z=" & sAlt1 & _
                       "  |  Data = autre point, Reset = nouvelle selection"
        Else
            g_oMoteur.CreerCercleSeul oPtDecale, g_oSettings
            m_dZFinal = dZFinal

            Dim sAlt2 As String
            sAlt2 = g_oCalc.FormatAltitude(dZFinal, g_oSettings.oTexte.Decimales, _
                                             g_oSettings.oTexte.SepDecimal)
            ShowPrompt "Cercle cree Z=" & sAlt2 & _
                       "  |  Placez le texte altitude  (Data=poser  Reset=nouv. selection)"
            m_nEtape = 2
        End If

    Case 2
        ' --- Poser le texte a la position choisie ---
        Point.Z = m_dZFinal
        g_oMoteur.CreerTexteSeul Point, m_dZFinal, g_oSettings, g_oCalc

        Dim sAlt3 As String
        sAlt3 = g_oCalc.FormatAltitude(m_dZFinal, g_oSettings.oTexte.Decimales, _
                                        g_oSettings.oTexte.SepDecimal)
        ShowPrompt "Point cree Z=" & sAlt3 & _
                   "  |  Data = autre point sur la meme droite, Reset = nouvelle selection"
        m_nEtape = 1
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset = nouvelle selection de P1
Private Sub IPrimitiveCommandEvents_Reset()
    frmInterpolation.ReinitialiserEtat
    ShowPrompt "Nouvelle selection : cliquez pres du texte altitude du POINT 1  (Reset=quitter)"
    CommandState.StartPrimitive New CSelectP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'------------------------------------------------------------------------------
Private Function SourceP1Integree() As Boolean
    SourceP1Integree = False
    If g_oSelectionP1 Is Nothing Then Exit Function
    If g_oSelectionP1.EstReference Then Exit Function
    SourceP1Integree = (Not g_oSelectionP1.oCellule Is Nothing) Or _
                       (Not g_oSelectionP1.oTag Is Nothing)
End Function
