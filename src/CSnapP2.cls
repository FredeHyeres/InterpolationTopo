VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSnapP2 - Etat commande : positionnement optionnel du point terrain pour P2
'
' Apres selection du texte altitude (CSelectP2), l'utilisateur peut :
'   Data  = snapper/cliquer le point terrain reel (remplace l'ancrage du texte)
'   Reset = garder le point d'ancrage du texte comme position terrain
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt "P2 = " & g_sValeurTexte & _
               "  |  Snappez le point terrain P2  (Reset = ancrage du texte)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_oP2.DefinirPosition Point

    If g_oCalc.Dist2DRef(g_oP1, g_oP2) < 0.001 Then
        ShowPrompt "P1 et P2 confondus - snappez un autre point terrain (P2)"
        Exit Sub
    End If

    frmInterpolation.AfficherP2 g_oP2
    frmInterpolation.AfficherSegment g_oCalc.InfoSegment(g_oP1, g_oP2)
    frmInterpolation.ActiverBoutonsActions True

    ShowPrompt "P2 positionne  |  " & g_oCalc.InfoSegment(g_oP1, g_oP2) & _
               "  ->  deplacez le curseur. Data=figer altitude  Reset=nouv. selection"
    CommandState.StartPrimitive New CPlacerPoint
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset = garder l'ancrage du texte et passer au placement
Private Sub IPrimitiveCommandEvents_Reset()
    If g_oCalc.Dist2DRef(g_oP1, g_oP2) < 0.001 Then
        ShowPrompt "P1 et P2 confondus - choisissez un autre texte (P2)"
        CommandState.StartPrimitive New CSelectP2
        Exit Sub
    End If

    frmInterpolation.AfficherP2 g_oP2
    frmInterpolation.AfficherSegment g_oCalc.InfoSegment(g_oP1, g_oP2)
    frmInterpolation.ActiverBoutonsActions True

    ShowPrompt "P2 = ancrage texte  |  " & g_oCalc.InfoSegment(g_oP1, g_oP2) & _
               "  ->  deplacez le curseur. Data=figer altitude  Reset=nouv. selection"
    CommandState.StartPrimitive New CPlacerPoint
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
