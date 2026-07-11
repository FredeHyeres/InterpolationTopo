VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSnapP2 - Etat commande : positionnement optionnel du point terrain pour P2
'
' Apres selection du texte altitude (CSelectP2), l'utilisateur peut :
'   Data  = snapper/cliquer le point terrain reel (remplace l'ancrage du texte)
'   Reset = garder le point d'ancrage du texte comme position terrain
'
' Classe partagee entre les deux commandes via la propriete Mode :
'   modeInterpolation : passe au placement du point interpole (CPlacerPoint)
'   modePonctuelle    : calcule la pente P1-P2, la reporte dans les parametres
'                       chemin, puis demarre le chemin (CPlacerPonctuel)
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
    ShowPrompt "P2 = " & g_oSelectionCourante.ValeurTexte & _
               "  |  Snappez le point terrain P2  (Reset = ancrage du texte)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_oP2.DefinirPosition Point

    If g_oCalc.Dist2DRef(g_oP1, g_oP2) < 0.001 Then
        ShowPrompt "P1 et P2 confondus - snappez un autre point terrain (P2)"
        Exit Sub
    End If

    If Mode = modePonctuelle Then
        DemarrerPlacement ""
    Else
        DemarrerPlacement "P2 positionne  |  "
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset = garder l'ancrage du texte et passer au placement
Private Sub IPrimitiveCommandEvents_Reset()
    If g_oCalc.Dist2DRef(g_oP1, g_oP2) < 0.001 Then
        ShowPrompt "P1 et P2 confondus - choisissez un autre texte (P2)"
        Dim oRetour As New CSelectP2
        oRetour.Mode = Mode
        CommandState.StartPrimitive oRetour
        Exit Sub
    End If

    If Mode = modePonctuelle Then
        DemarrerPlacement "P2 = ancrage texte  "
    Else
        DemarrerPlacement "P2 = ancrage texte  |  "
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'------------------------------------------------------------------------------
' Segment P1-P2 valide : met a jour le formulaire et lance l'etat de placement
' selon le mode. sPrefixe complete le message d'invite (origine du point P2).
Private Sub DemarrerPlacement(sPrefixe As String)
    If Mode = modePonctuelle Then
        AppliquerPenteP1P2
        frmInterpolPonct.AfficherP2 g_oP2
        frmInterpolPonct.RafraichirChemin
        ShowPrompt sPrefixe & "Pente=" & Format$(g_oSettings.oChemin.Pente, "0.00") & _
                   "%  ->  CHEMIN : Data=placer un point  Reset=nouv. selection"
        CommandState.StartPrimitive New CPlacerPonctuel
    Else
        frmInterpolation.AfficherP2 g_oP2
        frmInterpolation.AfficherSegment g_oCalc.InfoSegment(g_oP1, g_oP2)
        frmInterpolation.ActiverBoutonsActions True
        ShowPrompt sPrefixe & g_oCalc.InfoSegment(g_oP1, g_oP2) & _
                   "  ->  deplacez le curseur. Data=figer altitude  Reset=nouv. selection"
        CommandState.StartPrimitive New CPlacerPoint
    End If
End Sub

'------------------------------------------------------------------------------
Private Sub AppliquerPenteP1P2()
    Dim dPente As Double
    dPente = g_oCalc.PentePct(g_oP1, g_oP2)
    g_oSettings.oChemin.Pente = dPente
    g_oSettings.oChemin.PenteActive = True
End Sub
