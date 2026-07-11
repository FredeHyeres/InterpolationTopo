VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelectP1Ponct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSelectP1Ponct - Selection du point de depart (Interpolation Ponctuelle)
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpol. Ponctuelle"
    ShowPrompt "Cliquez pres du texte altitude du POINT 1  (Reset = quitter)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim oText As TextElement
    If Not TrouverAltitudeProche(Point, g_oSettings.dTolTexte, oText) Then
        ShowPrompt "Aucun texte altitude a moins de " & g_oSettings.dTolTexte & _
                   " u.m. - recliquez (P1)"
        Exit Sub
    End If

    Dim sSep As String, nDec As Integer
    If Not g_oCalc.LireAltitudeTexte(g_oP1, sSep, nDec) Then
        ShowPrompt "Texte '" & g_sValeurTexte & "' non numerique - recliquez (P1)"
        Exit Sub
    End If
    g_oSelectionP1.CopierDepuis g_oSelectionCourante

    If Not oText Is Nothing Then
        g_oSettings.ChargerModeleTexte oText, sSep, nDec
    Else
        g_oSettings.ChargerModeleDepuisTag g_oTagTrouve, sSep, nDec
    End If

    frmInterpolPonct.AfficherP1 g_oP1
    frmInterpolPonct.RafraichirTexte

    ShowPrompt "P1 = " & g_sValeurTexte & _
               "  |  Snappez le point terrain P1  (Reset = ancrage du texte)"
    CommandState.StartPrimitive New CSnapP1Ponct
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    frmInterpolPonct.Hide
    ShowPrompt "Interpol. Ponctuelle : termine."
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
