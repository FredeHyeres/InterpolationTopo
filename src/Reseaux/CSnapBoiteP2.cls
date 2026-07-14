VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapBoiteP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSnapBoiteP2 - Etat 3 : snap du POINT 2
'
'   Dynamics = ligne provisoire g_ptSnapP1 -> curseur
'   Data     = memorise g_ptSnapP2 (verif P1 <> P2) -> passe au clic cellule P2
'   Reset    = revient au clic cellule P1
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Pente Reseau EU/EP"
    ShowPrompt "P1 : " & DecrireBoite(g_oBoiteP1) & _
               "  |  snappez le POINT 2"
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    ' On force msdDrawingModeTemporary : sinon MicroStation peut recevoir un
    ' DrawMode Normal et laisser la derniere frame dessinee dans le DGN a la
    ' transition vers l'etat suivant.
    On Error Resume Next
    Dim ptA As Point3d
    ptA = g_ptSnapP1
    ptA.Z = Point.Z   ' vue a plat pour lisibilite pendant le drag

    Dim oL As LineElement
    Set oL = CreateLineElement2(Nothing, ptA, Point)
    oL.Color = 3
    oL.Redraw msdDrawingModeTemporary
    On Error GoTo 0
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim dd As Double
    dd = Sqr((Point.X - g_ptSnapP1.X) ^ 2 + (Point.Y - g_ptSnapP1.Y) ^ 2)
    If dd < 0.001 Then
        ShowPrompt "P2 confondu avec P1 - resnappez un autre POINT 2"
        Exit Sub
    End If

    g_ptSnapP2 = Point
    ShowPrompt "P2 snappe  ->  cliquez pres de la cellule regard/grille du P2"
    CommandState.StartPrimitive New CSelectBoiteP2
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Retour au clic cellule P1  (Reset = re-snap P1)"
    CommandState.StartPrimitive New CSelectBoiteP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
