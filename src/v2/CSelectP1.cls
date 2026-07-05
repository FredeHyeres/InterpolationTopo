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
    ' Cherche le texte numerique le plus proche du clic
    Dim oText As TextElement
    Set oText = TrouverTexteProche(Point, g_oSettings.dTolTexte)
    If oText Is Nothing Then
        ShowPrompt "Aucun texte altitude a moins de " & g_oSettings.dTolTexte & _
                   " u.m. - recliquez (P1)"
        Exit Sub
    End If

    ' Parsing de l'altitude
    Dim sSep As String, nDec As Integer
    If Not g_oCalc.LireAltitudeTexte(oText, g_oP1, sSep, nDec) Then
        ShowPrompt "Texte '" & oText.Text & "' non numerique - recliquez (P1)"
        Exit Sub
    End If

    ' Mise a jour des parametres : texte-modele et format
    g_oSettings.ChargerModeleTexte oText, sSep, nDec

    ' Mise a jour du formulaire (etat P1 + attributs herites du texte modele)
    frmInterpolation.AfficherP1 g_oP1
    frmInterpolation.RafraichirTexte

    ShowPrompt "P1 = " & oText.Text & "  ->  cliquez pres du texte altitude du POINT 2"
    CommandState.StartPrimitive New CSelectP2
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
