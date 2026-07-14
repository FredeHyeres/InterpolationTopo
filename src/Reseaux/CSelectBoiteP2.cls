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
' CSelectBoiteP2 - Etat commande : selection de la cellule regard/grille P2
'
'   Dynamics = ligne provisoire P1 (Fe P1) -> curseur, en rappel du snap en cours.
'   Data     = cherche la BoiteEP la plus proche, memorise dans g_oBoiteP2,
'              trace la ligne definitive P1-P2 puis passe a CPlacerPenteReseau.
'   Reset    = revient a la selection de P1.
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Pente Reseau EU/EP"
    ShowPrompt "P1 : " & DecrireBoite(g_oBoiteP1) & _
               "  |  cliquez pres du regard/grille du POINT 2"
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    On Error Resume Next

    ' Ligne provisoire P1 -> curseur (a plat sur Z du curseur pour lisibilite 2D)
    Dim ptP1 As Point3d
    g_oBoiteP1.CopierOrigine ptP1
    ptP1.Z = Point.Z

    Dim oL As LineElement
    Set oL = CreateLineElement2(Nothing, ptP1, Point)
    oL.Color = 3
    oL.Redraw DrawMode

    On Error GoTo 0
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim oInfo As New CBoiteEPInfo
    If Not TrouverBoiteEPProche(Point, TolReseau(), oInfo) Then
        ShowPrompt "Aucune cellule regard/grille (avec Fe:) a moins de " & _
                   TolReseau() & " u.m. - recliquez (P2)"
        Exit Sub
    End If

    ' Verifier P1 <> P2 (distance 2D)
    Dim ptP1 As Point3d, ptP2 As Point3d
    g_oBoiteP1.CopierOrigine ptP1
    oInfo.CopierOrigine ptP2
    Dim dd As Double
    dd = Sqr((ptP2.X - ptP1.X) ^ 2 + (ptP2.Y - ptP1.Y) ^ 2)
    If dd < 0.001 Then
        ShowPrompt "P1 et P2 confondus - selectionnez une autre cellule (P2)"
        Exit Sub
    End If

    g_oBoiteP2.CopierDepuis oInfo
    AlimenterPointRefFe g_oRefFeP2, g_oBoiteP2

    ' Ligne definitive P1-P2 (au niveau du Fe de chaque cellule = altitude reelle)
    Dim ptL1 As Point3d, ptL2 As Point3d
    ptL1 = ptP1: ptL1.Z = g_oBoiteP1.ZFilEau
    ptL2 = ptP2: ptL2.Z = g_oBoiteP2.ZFilEau
    Dim oLigne As LineElement
    Set oLigne = CreateLineElement2(Nothing, ptL1, ptL2)
    ActiveModelReference.AddElement oLigne

    Dim dPente As Double
    dPente = g_oCalc.PentePct(g_oRefFeP1, g_oRefFeP2)
    ShowPrompt "P2 : " & DecrireBoite(g_oBoiteP2) & _
               "  |  Pente=" & Format$(dPente, "0.00") & _
               "%  ->  cliquez pour placer texte+fleche  (Reset=nouv. selection)"
    CommandState.StartPrimitive New CPlacerPenteReseau
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Retour au POINT 1 : cliquez pres du regard/grille  (Reset = quitter)"
    CommandState.StartPrimitive New CSelectBoiteP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
