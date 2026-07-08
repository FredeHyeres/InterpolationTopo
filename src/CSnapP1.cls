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
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt "P1 = " & g_sValeurTexte & _
               "  |  Snappez le point terrain P1  (Reset = ancrage du texte)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_oP1.DefinirPosition Point
    frmInterpolation.AfficherP1 g_oP1

    ShowPrompt "P1 positionne  ->  cliquez pres du texte altitude du POINT 2"
    CommandState.StartPrimitive New CSelectP2
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset = garder l'ancrage du texte et passer a P2
Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "P1 = ancrage texte  ->  cliquez pres du texte altitude du POINT 2"
    CommandState.StartPrimitive New CSelectP2
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
