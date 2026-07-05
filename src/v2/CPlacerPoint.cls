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
' Data = creer le point (repetable sur la meme droite P1-P2)
' Reset = revenir a la selection de P1
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt g_oCalc.InfoSegment(g_oP1, g_oP2) & _
               "  |  Deplacez le curseur : Data=creer  Reset=nouv. selection"
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    Dim dT As Double, oProj As Point3d
    g_oCalc.ProjecterSurSegment g_oP1, g_oP2, Point, dT, oProj

    g_oMoteur.DessinDynamique DrawMode, g_oP1, g_oP2, oProj, dT, g_oSettings, g_oCalc

    ' Prompt : signale l'extrapolation si le curseur est hors segment
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
                   "  (Data=creer  Reset=nouv. selection)"
    End If
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim dT As Double, oProj As Point3d
    g_oCalc.ProjecterSurSegment g_oP1, g_oP2, Point, dT, oProj

    Dim dZInterp As Double
    dZInterp = g_oCalc.InterpolerZ(g_oP1.Altitude, g_oP2.Altitude, dT)

    g_oMoteur.CreerPointTopo oProj, dZInterp, g_oSettings, g_oCalc

    Dim sAlt As String
    sAlt = g_oCalc.FormatAltitude(dZInterp, g_oSettings.oTexte.Decimales, _
                                   g_oSettings.oTexte.SepDecimal)
    ShowPrompt "Point cree Z=" & sAlt & _
               "  |  Data = autre point sur la meme droite, Reset = nouvelle selection"
    ' La commande reste active : on peut creer d'autres points sur la meme droite
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset = nouvelle selection de P1 (memes parametres cercle conserves dans le formulaire)
Private Sub IPrimitiveCommandEvents_Reset()
    frmInterpolation.ReinitialiserEtat
    ShowPrompt "Nouvelle selection : cliquez pres du texte altitude du POINT 1  (Reset=quitter)"
    CommandState.StartPrimitive New CSelectP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
