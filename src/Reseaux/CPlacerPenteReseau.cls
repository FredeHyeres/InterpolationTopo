VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlacerPenteReseau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPlacerPenteReseau - Placement du texte pente + fleche entre deux BoiteEP
'
' Reutilise g_oMoteur.DessinPenteDynamique / g_oMoteur.CreerPente pour ne pas
' dupliquer la geometrie de la fleche ni le rendu du texte pente : g_oRefFeP1
' et g_oRefFeP2 (alimentes par les Fe) tiennent lieu de couple P1/P2 topo.
'
'   Dynamics = apercu texte+fleche a la position du curseur
'   Data     = creation definitive + auto-chainage : P2 devient P1 et on
'              retourne au snap du POINT 2 (regard suivant du reseau)
'   Reset    = retour au snap POINT 1 (nouvelle selection independante)
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Pente Reseau EU/EP"
    ShowPrompt "Pente=" & g_oCalc.FormatPente( _
               g_oCalc.PentePct(g_oRefFeP1, g_oRefFeP2), _
               g_oSettings.oTexte.SepDecimal, _
               g_oSettings.oIndicPente.Decimales) & _
               "  |  cliquez pour placer texte+fleche  (Reset=nouv. selection)"
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    On Error Resume Next

    ' Ligne provisoire P1_snap -> P2_snap (visuelle uniquement, ne persiste pas)
    Dim ptL1 As Point3d, ptL2 As Point3d
    ptL1 = g_ptSnapP1: ptL1.Z = Point.Z
    ptL2 = g_ptSnapP2: ptL2.Z = Point.Z
    Dim oL As LineElement
    Set oL = CreateLineElement2(Nothing, ptL1, ptL2)
    oL.Color = 3
    oL.Redraw msdDrawingModeTemporary

    ' Apercu texte pente + fleche
    g_oMoteur.DessinPenteDynamique DrawMode, Point, _
                                   g_oRefFeP1, g_oRefFeP2, g_oSettings, g_oCalc
    On Error GoTo 0
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    ' 1) Creer texte pente + fleche definitifs
    g_oMoteur.CreerPente Point, g_oRefFeP1, g_oRefFeP2, g_oSettings, g_oCalc

    ' 2) Auto-chainage : le regard P2 devient le nouveau P1
    g_oBoiteP1.CopierDepuis g_oBoiteP2
    g_ptSnapP1 = g_ptSnapP2
    AlimenterPointRefFeSnap g_oRefFeP1, g_ptSnapP1, g_oBoiteP1

    ' 3) Preparer un P2 propre
    g_oBoiteP2.Vider
    g_oRefFeP2.Valide = False

    ' 4) Mise a jour du formulaire (P1 = ex-P2, P2 vide)
    frmReseauEP.AfficherBoiteP1 g_oBoiteP1
    frmReseauEP.EffacerP2

    ' 5) On repart directement au snap du regard suivant
    ShowPrompt "Pente placee. P1 = ex-P2 (" & DecrireBoite(g_oBoiteP1) & _
               ")  |  snappez le POINT 2 suivant  (Reset = nouv. selection)"
    CommandState.StartPrimitive New CSnapBoiteP2
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    frmReseauEP.ReinitialiserEtat
    ShowPrompt "Nouvelle selection : snappez le POINT 1  (Reset=quitter)"
    CommandState.StartPrimitive New CSnapBoiteP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
