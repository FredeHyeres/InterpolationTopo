VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelectBoiteP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSelectBoiteP1 - Etat 2 : clic explicite sur la cellule Fe de P1
'
' Le snap de P1 est deja memorise dans g_ptSnapP1. Ce clic sert uniquement a
' identifier la cellule regard/grille (parfois plusieurs reseaux proches).
'
'   Data  = cherche la BoiteEP la plus proche du clic dans le rayon
'           g_dTolReseau, memorise ses altitudes et alimente g_oRefFeP1
'           avec (g_ptSnapP1.X, g_ptSnapP1.Y, Fe).
'   Reset = revient au snap P1.
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Pente Reseau EU/EP"
    ShowPrompt "Cliquez pres de la cellule regard/grille du P1  (Reset = re-snap P1)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim oInfo As New CBoiteEPInfo
    If Not TrouverBoiteEPProche(Point, TolReseau(), oInfo) Then
        ShowPrompt "Aucune cellule regard/grille (avec Fe:) a moins de " & _
                   TolReseau() & " u.m. - recliquez (cellule P1)"
        Exit Sub
    End If

    g_oBoiteP1.CopierDepuis oInfo
    AlimenterPointRefFeSnap g_oRefFeP1, g_ptSnapP1, g_oBoiteP1
    frmReseauEP.AfficherBoiteP1 g_oBoiteP1

    ShowPrompt "P1 : " & DecrireBoite(g_oBoiteP1) & _
               "  |  snappez le POINT 2"
    CommandState.StartPrimitive New CSnapBoiteP2
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Retour au snap P1 : snappez le POINT 1  (Reset = quitter)"
    CommandState.StartPrimitive New CSnapBoiteP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
