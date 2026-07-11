VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapP1Ponct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSnapP1Ponct - Snap terrain pour P1 (Interpolation Ponctuelle)
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpol. Ponctuelle"
    ShowPrompt "P1 = " & g_sValeurTexte & _
               "  |  Snappez le point terrain P1  (Reset = ancrage du texte)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_oP1.DefinirPosition Point
    frmInterpolPonct.AfficherP1 g_oP1

    ShowPrompt "P1 positionne  ->  P2 optionnel (Data=selectionner P2  Reset=pente manuelle)"
    CommandState.StartPrimitive New CSelectP2Ponct
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "P1 = ancrage texte  ->  P2 optionnel (Data=selectionner P2  Reset=pente manuelle)"
    CommandState.StartPrimitive New CSelectP2Ponct
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
