VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlacerPente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPlacerPente - Placement dynamique du texte pente + fleche
'
' Apres snap P2, l'utilisateur clique pour placer l'indicateur de pente.
'   Data  = placer le texte pente + fleche a la position du curseur
'   Reset = revenir a la selection de P1
' La fleche pointe dans la direction de descente (du plus haut vers le plus bas).
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt "Pente=" & g_oCalc.FormatPente(g_oCalc.PentePct(g_oP1, g_oP2), _
               g_oSettings.oTexte.SepDecimal) & _
               "  |  Cliquez pour placer pente + fleche  (Reset=nouv. selection)"
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    g_oMoteur.DessinPenteDynamique DrawMode, Point, g_oP1, g_oP2, g_oSettings, g_oCalc
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_oMoteur.CreerPente Point, g_oP1, g_oP2, g_oSettings, g_oCalc

    ShowPrompt "Pente placee  |  " & _
               g_oCalc.FormatPente(g_oCalc.PentePct(g_oP1, g_oP2), _
               g_oSettings.oTexte.SepDecimal) & _
               "  Data=placer encore  Reset=nouv. selection"
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    frmInterpolation.ReinitialiserEtat
    ShowPrompt "Nouvelle selection : cliquez pres du texte altitude du POINT 1  (Reset=quitter)"
    CommandState.StartPrimitive New CSelectP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
