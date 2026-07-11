VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelectP2Ponct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSelectP2Ponct - Selection optionnelle de P2 (Interpolation Ponctuelle)
'
' Data  = selectionner un texte altitude P2 (pour calculer la pente)
' Reset = sauter P2, utiliser la pente saisie manuellement dans le formulaire
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpol. Ponctuelle"
    ShowPrompt "P2 optionnel : Data=selectionner P2  Reset=pente manuelle, demarrer le chemin"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim oText As TextElement
    If Not TrouverAltitudeProche(Point, g_oSettings.dTolTexte, oText) Then
        ShowPrompt "Aucun texte altitude a moins de " & g_oSettings.dTolTexte & _
                   " u.m. - recliquez (P2)"
        Exit Sub
    End If

    Dim sSep As String, nDec As Integer
    If Not g_oCalc.LireAltitudeTexte(g_oP2, sSep, nDec) Then
        ShowPrompt "Texte '" & g_sValeurTexte & "' non numerique - recliquez (P2)"
        Exit Sub
    End If
    g_oSelectionP2.CopierDepuis g_oSelectionCourante

    frmInterpolPonct.AfficherP2 g_oP2

    ShowPrompt "P2 = " & g_sValeurTexte & _
               "  |  Snappez le point terrain P2  (Reset = ancrage du texte)"
    CommandState.StartPrimitive New CSnapP2Ponct
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset = sauter P2, demarrer le chemin avec la pente manuelle
Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Pente manuelle  ->  CHEMIN : Data=placer un point  Reset=nouv. selection"
    CommandState.StartPrimitive New CPlacerPonctuel
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
