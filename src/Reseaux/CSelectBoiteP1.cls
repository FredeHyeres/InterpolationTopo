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
' CSelectBoiteP1 - Etat commande : selection de la cellule regard/grille P1
'
'   Data  = cherche la cellule BoiteEP la plus proche du clic (rayon TOL_RESEAU)
'           et memorise ses altitudes (T:/Fe:/Prof:) dans g_oBoiteP1.
'   Reset = quitte la commande.
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Pente Reseau EU/EP"
    ShowPrompt "Cliquez pres du regard/grille du POINT 1  (Reset = quitter)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim oInfo As New CBoiteEPInfo
    If Not TrouverBoiteEPProche(Point, TolReseau(), oInfo) Then
        ShowPrompt "Aucune cellule regard/grille (avec Fe:) a moins de " & _
                   TolReseau() & " u.m. - recliquez (P1)"
        Exit Sub
    End If

    g_oBoiteP1.CopierDepuis oInfo
    AlimenterPointRefFe g_oRefFeP1, g_oBoiteP1

    ShowPrompt "P1 : " & DecrireBoite(g_oBoiteP1) & _
               "  |  cliquez pres du regard/grille du POINT 2"
    CommandState.StartPrimitive New CSelectBoiteP2
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Pente Reseau EU/EP : termine."
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
