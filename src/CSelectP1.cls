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
' CSelectP1 - Etat commande : selection du point de reference P1
'
' Classe partagee entre les deux commandes via la propriete Mode :
'   modeInterpolation (defaut) : commande Interpolation Topo
'   modePonctuelle             : commande Interpolation Ponctuelle
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Public Mode As eModeCommande

Private Sub IPrimitiveCommandEvents_Start()
    If Mode = modePonctuelle Then
        ShowCommand "Interpol. Ponctuelle"
    Else
        ShowCommand "Interpolation Topo"
    End If
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
        ShowPrompt "Texte '" & g_oSelectionCourante.ValeurTexte & _
                   "' non numerique - recliquez (P1)"
        Exit Sub
    End If
    g_oSelectionP1.CopierDepuis g_oSelectionCourante

    ' Mise a jour des parametres : texte-modele et format
    If Not oText Is Nothing Then
        g_oSettings.ChargerModeleTexte oText, sSep, nDec
    Else
        g_oSettings.ChargerModeleDepuisTag g_oSelectionCourante.oTag, sSep, nDec
    End If

    ' Mise a jour du formulaire (etat P1 + attributs herites du texte modele)
    If Mode = modePonctuelle Then
        frmInterpolPonct.AfficherP1 g_oP1
        frmInterpolPonct.RafraichirTexte
    Else
        frmInterpolation.AfficherP1 g_oP1
        frmInterpolation.RafraichirTexte
        frmInterpolation.RafraichirPente
    End If

    ShowPrompt "P1 = " & g_oSelectionCourante.ValeurTexte & _
               "  |  Snappez le point terrain P1  (Reset = ancrage du texte)"
    Dim oSuivant As New CSnapP1
    oSuivant.Mode = Mode
    CommandState.StartPrimitive oSuivant
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset au choix de P1 = quitter la commande
Private Sub IPrimitiveCommandEvents_Reset()
    If Mode = modePonctuelle Then
        frmInterpolPonct.Hide
        ShowPrompt "Interpol. Ponctuelle : termine."
    Else
        frmInterpolation.Hide
        ShowPrompt "Interpolation Topo : termine."
    End If
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
