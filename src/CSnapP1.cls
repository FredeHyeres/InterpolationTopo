VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSnapP1 - Etat commande : positionnement optionnel du point terrain pour P1
'
' Apres selection du texte altitude (CSelectP1), l'utilisateur peut :
'   Data  = snapper/cliquer le point terrain reel (remplace l'ancrage du texte)
'   Reset = garder le point d'ancrage du texte comme position terrain
'
' Classe partagee entre les deux commandes via la propriete Mode.
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
    ShowPrompt "P1 = " & g_oSelectionCourante.ValeurTexte & _
               "  |  Snappez le point terrain P1  (Reset = ancrage du texte)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_oP1.DefinirPosition Point

    If Mode = modePonctuelle Then
        frmInterpolPonct.AfficherP1 g_oP1
        ShowPrompt "P1 positionne  ->  P2 optionnel (Data=selectionner P2  Reset=pente manuelle)"
    Else
        frmInterpolation.AfficherP1 g_oP1
        ShowPrompt "P1 positionne  ->  cliquez pres du texte altitude du POINT 2"
    End If

    PasserSelectionP2
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset = garder l'ancrage du texte et passer a P2
Private Sub IPrimitiveCommandEvents_Reset()
    If Mode = modePonctuelle Then
        ShowPrompt "P1 = ancrage texte  ->  P2 optionnel (Data=selectionner P2  Reset=pente manuelle)"
    Else
        ShowPrompt "P1 = ancrage texte  ->  cliquez pres du texte altitude du POINT 2"
    End If

    PasserSelectionP2
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'------------------------------------------------------------------------------
Private Sub PasserSelectionP2()
    Dim oSuivant As New CSelectP2
    oSuivant.Mode = Mode
    CommandState.StartPrimitive oSuivant
End Sub
