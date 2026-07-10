VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelectP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSelectP1 V2 - Etat commande : selection du point de reference P1
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt "Cliquez pres du texte altitude du POINT 1  (Reset = quitter)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    ' Cherche le texte/tag/cellule numerique le plus proche du clic
    Dim oText As TextElement
    If Not TrouverAltitudeProche(Point, g_oSettings.dTolTexte, oText) Then
        ShowPrompt "Aucun texte altitude a moins de " & g_oSettings.dTolTexte & _
                   " u.m. - recliquez (P1)"
        Exit Sub
    End If

    ' Parsing de l'altitude
    Dim sSep As String, nDec As Integer
    If Not g_oCalc.LireAltitudeTexte(g_oP1, sSep, nDec) Then
        ShowPrompt "Texte '" & g_sValeurTexte & "' non numerique - recliquez (P1)"
        Exit Sub
    End If
    g_oSelectionP1.CopierDepuis g_oSelectionCourante

    ' Mise a jour des parametres : texte-modele et format
    If Not oText Is Nothing Then
        g_oSettings.ChargerModeleTexte oText, sSep, nDec
    Else
        g_oSettings.ChargerModeleDepuisTag g_oTagTrouve, sSep, nDec
    End If

    ' Si P1 vient d'une reference, convertir les tailles texte des unites
    ' de la reference vers les unites du fichier actif (ScaleFactor)
    If g_oSelectionP1.EstReference And Not g_oSelectionP1.oAttachment Is Nothing Then
        Dim dScale As Double
        dScale = 1#
        On Error Resume Next
        dScale = CDbl(g_oSelectionP1.oAttachment.ScaleFactor)
        If Err.Number <> 0 Or Abs(dScale) < 0.0000000001 Then dScale = 1#
        Err.Clear
        On Error GoTo 0
        If Abs(dScale - 1#) > 0.0000000001 Then
            g_oSettings.dModeleHauteur = g_oSettings.dModeleHauteur * dScale
            g_oSettings.dPenteHauteur = g_oSettings.dPenteHauteur * dScale
            g_oSettings.dPenteLargeur = g_oSettings.dPenteLargeur * dScale
        End If
    End If

    ' Mise a jour du formulaire (etat P1 + attributs herites du texte modele)
    frmInterpolation.AfficherP1 g_oP1
    frmInterpolation.RafraichirTexte
    frmInterpolation.RafraichirPente

    ShowPrompt "P1 = " & g_sValeurTexte & _
               "  |  Snappez le point terrain P1  (Reset = ancrage du texte)"
    CommandState.StartPrimitive New CSnapP1
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset au choix de P1 = quitter la commande
Private Sub IPrimitiveCommandEvents_Reset()
    frmInterpolation.Hide
    ShowPrompt "Interpolation Topo : termine."
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
