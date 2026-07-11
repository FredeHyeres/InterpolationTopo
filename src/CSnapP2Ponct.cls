VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapP2Ponct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSnapP2Ponct - Snap terrain pour P2 (Interpolation Ponctuelle)
'
' Apres snap, calcule la pente P1-P2 et la reporte dans les parametres chemin.
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpol. Ponctuelle"
    ShowPrompt "P2 = " & g_sValeurTexte & _
               "  |  Snappez le point terrain P2  (Reset = ancrage du texte)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_oP2.DefinirPosition Point

    If g_oCalc.Dist2DRef(g_oP1, g_oP2) < 0.001 Then
        ShowPrompt "P1 et P2 confondus - snappez un autre point terrain (P2)"
        Exit Sub
    End If

    AppliquerPenteP1P2
    frmInterpolPonct.AfficherP2 g_oP2
    frmInterpolPonct.RafraichirChemin

    ShowPrompt "Pente=" & Format$(g_oSettings.dCheminPente, "0.00") & "%  ->  " & _
               "CHEMIN : Data=placer un point  Reset=nouv. selection"
    CommandState.StartPrimitive New CPlacerPonctuel
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    If g_oCalc.Dist2DRef(g_oP1, g_oP2) < 0.001 Then
        ShowPrompt "P1 et P2 confondus - choisissez un autre texte (P2)"
        CommandState.StartPrimitive New CSelectP2Ponct
        Exit Sub
    End If

    AppliquerPenteP1P2
    frmInterpolPonct.AfficherP2 g_oP2
    frmInterpolPonct.RafraichirChemin

    ShowPrompt "P2 = ancrage texte  Pente=" & Format$(g_oSettings.dCheminPente, "0.00") & _
               "%  ->  CHEMIN : Data=placer un point  Reset=nouv. selection"
    CommandState.StartPrimitive New CPlacerPonctuel
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'------------------------------------------------------------------------------
Private Sub AppliquerPenteP1P2()
    Dim dPente As Double
    dPente = g_oCalc.PentePct(g_oP1, g_oP2)
    g_oSettings.dCheminPente = dPente
    g_oSettings.bCheminPente = True
End Sub
