VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelectP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSelectP2 V2 - Etat commande : selection du point de reference P2
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt "Cliquez pres du texte altitude du POINT 2"
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

    If g_oCalc.Dist2DRef(g_oP1, g_oP2) < 0.001 Then
        ShowPrompt "P1 et P2 confondus - choisissez un autre point (P2)"
        Exit Sub
    End If

    ' Mise a jour decimales si P2 en a davantage
    g_oSettings.AjusterDecimales nDec

    ' Mise a jour du formulaire
    frmInterpolation.AfficherP2 g_oP2
    frmInterpolation.AfficherSegment g_oCalc.InfoSegment(g_oP1, g_oP2)

    ShowPrompt "P2 = " & g_sValeurTexte & "  |  " & g_oCalc.InfoSegment(g_oP1, g_oP2) & _
               "  ->  deplacez le curseur. Data=creer  Reset=nouv. selection"
    CommandState.StartPrimitive New CPlacerPoint
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset pendant P2 = revenir a P1
Private Sub IPrimitiveCommandEvents_Reset()
    frmInterpolation.ReinitialiserEtat
    ShowPrompt "Retour au POINT 1 : cliquez pres du texte altitude  (Reset = quitter)"
    CommandState.StartPrimitive New CSelectP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
