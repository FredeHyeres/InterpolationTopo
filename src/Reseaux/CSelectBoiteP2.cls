VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelectBoiteP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSelectBoiteP2 - Etat 4 : clic explicite sur la cellule Fe de P2
'
' Le snap P2 est deja dans g_ptSnapP2. Ce clic identifie la cellule Fe.
' Si trouve, alimente g_oRefFeP2 et passe a CPlacerPenteReseau. Aucun element
' n'est ajoute au DGN a ce stade : la ligne P1-P2 reste PROVISOIRE (dessinee
' par CPlacerPenteReseau.Dynamics) et disparait au placement pente. Seuls le
' texte pente + fleche seront persistes (via g_oMoteur.CreerPente).
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Pente Reseau EU/EP"
    ShowPrompt "P1 : " & DecrireBoite(g_oBoiteP1) & _
               "  |  cliquez pres de la cellule regard/grille du P2"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim oInfo As New CBoiteEPInfo
    If Not TrouverBoiteEPProche(Point, TolReseau(), oInfo) Then
        ShowPrompt "Aucune cellule regard/grille (avec Fe:) a moins de " & _
                   TolReseau() & " u.m. - recliquez (cellule P2)"
        Exit Sub
    End If

    g_oBoiteP2.CopierDepuis oInfo
    AlimenterPointRefFeSnap g_oRefFeP2, g_ptSnapP2, g_oBoiteP2

    Dim dDist As Double, dPente As Double
    dDist = Sqr((g_ptSnapP2.X - g_ptSnapP1.X) ^ 2 + (g_ptSnapP2.Y - g_ptSnapP1.Y) ^ 2)
    dPente = g_oCalc.PentePct(g_oRefFeP1, g_oRefFeP2)

    frmReseauEP.AfficherBoiteP2 g_oBoiteP2
    frmReseauEP.AfficherDistancePente dDist, dPente

    ShowPrompt "P2 : " & DecrireBoite(g_oBoiteP2) & _
               "  |  Pente=" & Format$(dPente, "0.0") & _
               "%  ->  cliquez pour placer texte+fleche  (Reset=nouv. selection)"
    CommandState.StartPrimitive New CPlacerPenteReseau
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Retour au snap P2  (Reset = re-clic cellule P1)"
    CommandState.StartPrimitive New CSnapBoiteP2
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
