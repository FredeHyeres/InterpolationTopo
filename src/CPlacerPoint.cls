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
' Etape 1 : projection sur la droite P1-P2 => Data = figer l'altitude
' Etape 2 : decalage perpendiculaire => Data = creer le cercle
' Etape 3 : placement libre du texte altitude => Data = creer le texte
'   (cellule/tag : etape 3 sautee, tout est cree en etape 2)
' Reset = revenir a la selection de P1
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private m_nEtape        As Integer      ' 1=projection, 2=perpendiculaire, 3=texte libre
Private m_oProj         As Point3d      ' point projete fige (etape 1)
Private m_dZInterp      As Double       ' altitude interpolee figee
Private m_dPerpX        As Double       ' vecteur unitaire perpendiculaire X
Private m_dPerpY        As Double       ' vecteur unitaire perpendiculaire Y
Private m_dZFinal       As Double       ' altitude finale pour le texte (etape 3)

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt g_oCalc.InfoSegment(g_oP1, g_oP2) & _
               "  |  Deplacez le curseur : Data=figer altitude  Reset=nouv. selection"
    m_nEtape = 1
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    Select Case m_nEtape
    Case 1
        ' --- Etape 1 : projection sur la droite P1-P2 ---
        Dim dT As Double, oProj As Point3d
        g_oCalc.ProjecterSurSegment g_oP1, g_oP2, Point, dT, oProj

        g_oMoteur.DessinDynamique DrawMode, g_oP1, g_oP2, oProj, dT, g_oSettings, g_oCalc

        Dim sAlt As String
        sAlt = g_oCalc.FormatAltitude(g_oCalc.InterpolerZ(g_oP1.Altitude, g_oP2.Altitude, dT), _
                                       g_oSettings.oTexte.Decimales, g_oSettings.oTexte.SepDecimal)
        If dT < 0# Or dT > 1# Then
            ShowPrompt "Z=" & sAlt & "  t=" & Format$(dT, "0.000") & "  " & _
                       g_oCalc.InfoSegment(g_oP1, g_oP2) & _
                       "  ATTENTION : hors segment (extrapolation)"
        Else
            ShowPrompt "Z=" & sAlt & "  t=" & Format$(dT, "0.000") & "  " & _
                       g_oCalc.InfoSegment(g_oP1, g_oP2) & _
                       "  (Data=figer altitude)"
        End If

    Case 2
        ' --- Etape 2 : decalage perpendiculaire ---
        Dim dOffset As Double
        Dim oPtDecale As Point3d
        dOffset = (Point.X - m_oProj.X) * m_dPerpX + (Point.Y - m_oProj.Y) * m_dPerpY
        oPtDecale.X = m_oProj.X + dOffset * m_dPerpX
        oPtDecale.Y = m_oProj.Y + dOffset * m_dPerpY
        oPtDecale.Z = m_oProj.Z

        Dim dZAff As Double
        dZAff = m_dZInterp
        If g_oSettings.bAppliquerPente Then
            dZAff = dZAff + Abs(dOffset) * g_oSettings.dPenteDecalage / 100#
        End If
        If g_oSettings.bDecalageFixe Then dZAff = dZAff + g_oSettings.dDecalageDZ

        g_oMoteur.DessinDecalage DrawMode, g_oP1, g_oP2, m_oProj, oPtDecale, _
                                  dZAff, g_oSettings, g_oCalc

        Dim sAlt2 As String
        sAlt2 = g_oCalc.FormatAltitude(dZAff, g_oSettings.oTexte.Decimales, _
                                        g_oSettings.oTexte.SepDecimal)
        ShowPrompt "Z=" & sAlt2 & "  decalage=" & Format$(dOffset, "0.000") & _
                   "  (Data=creer  Reset=nouv. selection)"

    Case 3
        ' --- Etape 3 : placement libre du texte ---
        g_oMoteur.DessinTexteLibre DrawMode, Point, m_dZFinal, g_oSettings, g_oCalc

        Dim sAlt3 As String
        sAlt3 = g_oCalc.FormatAltitude(m_dZFinal, g_oSettings.oTexte.Decimales, _
                                        g_oSettings.oTexte.SepDecimal)
        ShowPrompt "Placez le texte Z=" & sAlt3 & _
                   "  (Data=poser le texte  Reset=nouv. selection)"
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Select Case m_nEtape
    Case 1
        ' --- Etape 1 : figer l'altitude et passer au decalage ---
        Dim dT As Double
        g_oCalc.ProjecterSurSegment g_oP1, g_oP2, Point, dT, m_oProj
        m_dZInterp = g_oCalc.InterpolerZ(g_oP1.Altitude, g_oP2.Altitude, dT)

        Dim dx As Double, dy As Double, dLen As Double
        dx = g_oP2.X - g_oP1.X
        dy = g_oP2.Y - g_oP1.Y
        dLen = Sqr(dx * dx + dy * dy)
        If dLen > 0.0000001 Then
            m_dPerpX = -dy / dLen
            m_dPerpY = dx / dLen
        Else
            m_dPerpX = 0: m_dPerpY = 1
        End If

        m_nEtape = 2

        Dim sAlt As String
        sAlt = g_oCalc.FormatAltitude(m_dZInterp, g_oSettings.oTexte.Decimales, _
                                       g_oSettings.oTexte.SepDecimal)
        ShowPrompt "Altitude figee Z=" & sAlt & _
                   "  |  Deplacez pour decaler. Data=creer  Reset=nouv. selection"

    Case 2
        ' --- Etape 2 : creer le cercle (+ cellule/tag complet) ---
        Dim dOffset As Double
        Dim oPtDecale As Point3d
        dOffset = (Point.X - m_oProj.X) * m_dPerpX + (Point.Y - m_oProj.Y) * m_dPerpY
        oPtDecale.X = m_oProj.X + dOffset * m_dPerpX
        oPtDecale.Y = m_oProj.Y + dOffset * m_dPerpY
        oPtDecale.Z = m_oProj.Z

        Dim dZFinal As Double
        dZFinal = m_dZInterp
        If g_oSettings.bAppliquerPente Then
            dZFinal = dZFinal + Abs(dOffset) * g_oSettings.dPenteDecalage / 100#
        End If
        If g_oSettings.bDecalageFixe Then dZFinal = dZFinal + g_oSettings.dDecalageDZ

        If SourceP1Integree() Then
            ' Cellule/tag : tout creer d'un coup (texte integre)
            g_oMoteur.CreerPointTopo oPtDecale, dZFinal, g_oSettings, g_oCalc

            Dim sAlt2a As String
            sAlt2a = g_oCalc.FormatAltitude(dZFinal, g_oSettings.oTexte.Decimales, _
                                             g_oSettings.oTexte.SepDecimal)
            ShowPrompt "Point cree Z=" & sAlt2a & _
                       "  |  Data = autre point, Reset = nouvelle selection"
            m_nEtape = 1
        Else
            ' Texte : creer le cercle seul, passer au placement libre du texte
            g_oMoteur.CreerCercleSeul oPtDecale, g_oSettings
            m_dZFinal = dZFinal

            Dim sAlt2b As String
            sAlt2b = g_oCalc.FormatAltitude(dZFinal, g_oSettings.oTexte.Decimales, _
                                             g_oSettings.oTexte.SepDecimal)
            ShowPrompt "Cercle cree Z=" & sAlt2b & _
                       "  |  Placez le texte altitude  (Data=poser  Reset=nouv. selection)"
            m_nEtape = 3
        End If

    Case 3
        ' --- Etape 3 : poser le texte a la position choisie ---
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
' True si P1 est une cellule/tag du fichier actif. Si P1 vient d'une reference,
' on lit son altitude mais on cree un nouveau texte+cercle dans le fichier actif.
Private Function SourceP1Integree() As Boolean
    SourceP1Integree = False
    If g_oSelectionP1 Is Nothing Then Exit Function
    If g_oSelectionP1.EstReference Then Exit Function
    SourceP1Integree = (Not g_oSelectionP1.oCellule Is Nothing) Or _
                       (Not g_oSelectionP1.oTag Is Nothing)
End Function
