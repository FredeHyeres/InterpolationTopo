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
'   Data     = creation definitive
'   Reset    = retour a la selection de P1
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Pente Reseau EU/EP"
    ShowPrompt "Pente=" & g_oCalc.FormatPente( _
               g_oCalc.PentePct(g_oRefFeP1, g_oRefFeP2), _
               g_oSettings.oTexte.SepDecimal) & _
               "  |  cliquez pour placer texte+fleche  (Reset=nouv. selection)"
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    On Error Resume Next
    g_oMoteur.DessinPenteDynamique DrawMode, Point, _
                                   g_oRefFeP1, g_oRefFeP2, g_oSettings, g_oCalc
    On Error GoTo 0
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_oMoteur.CreerPente Point, g_oRefFeP1, g_oRefFeP2, g_oSettings, g_oCalc
    ShowPrompt "Pente placee  |  Pente=" & _
               g_oCalc.FormatPente( _
                   g_oCalc.PentePct(g_oRefFeP1, g_oRefFeP2), _
                   g_oSettings.oTexte.SepDecimal) & _
               "  Data=placer encore  Reset=nouv. selection"
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Nouvelle selection : cliquez pres du regard/grille du POINT 1  (Reset=quitter)"
    CommandState.StartPrimitive New CSelectBoiteP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
