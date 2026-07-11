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
' CSelectP2 - Etat commande : selection du point de reference P2
'
' Classe partagee entre les deux commandes via la propriete Mode :
'   modeInterpolation : P2 obligatoire, Reset = revenir a P1
'   modePonctuelle    : P2 optionnel, Reset = pente manuelle, demarrer le chemin
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Public Mode As eModeCommande

Private Sub IPrimitiveCommandEvents_Start()
    If Mode = modePonctuelle Then
        ShowCommand "Interpol. Ponctuelle"
        ShowPrompt "P2 optionnel : Data=selectionner P2  Reset=pente manuelle, demarrer le chemin"
    Else
        ShowCommand "Interpolation Topo"
        ShowPrompt "Cliquez pres du texte altitude du POINT 2"
    End If
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
        ShowPrompt "Texte '" & g_oSelectionCourante.ValeurTexte & _
                   "' non numerique - recliquez (P2)"
        Exit Sub
    End If
    g_oSelectionP2.CopierDepuis g_oSelectionCourante

    If Mode = modePonctuelle Then
        frmInterpolPonct.AfficherP2 g_oP2
    Else
        g_oSettings.AjusterDecimales nDec
        frmInterpolation.AfficherP2 g_oP2
    End If

    ShowPrompt "P2 = " & g_oSelectionCourante.ValeurTexte & _
               "  |  Snappez le point terrain P2  (Reset = ancrage du texte)"
    Dim oSuivant As New CSnapP2
    oSuivant.Mode = Mode
    CommandState.StartPrimitive oSuivant
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset pendant P2 :
'   Interpolation = revenir a P1
'   Ponctuelle    = sauter P2, demarrer le chemin avec la pente manuelle
Private Sub IPrimitiveCommandEvents_Reset()
    If Mode = modePonctuelle Then
        ShowPrompt "Pente manuelle  ->  CHEMIN : Data=placer un point  Reset=nouv. selection"
        CommandState.StartPrimitive New CPlacerPonctuel
    Else
        frmInterpolation.ReinitialiserEtat
        ShowPrompt "Retour au POINT 1 : cliquez pres du texte altitude  (Reset = quitter)"
        CommandState.StartPrimitive New CSelectP1
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
