VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelectP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSelectP2 - etape 2 : selection du texte altitude du point P2
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt "Cliquez pres du texte altitude du POINT 2"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim oText As TextElement
    Set oText = FindNearestText(Point, TOL_TEXTE)
    If oText Is Nothing Then
        ShowPrompt "Aucun texte altitude a moins de " & TOL_TEXTE & " u.m. - recliquez (P2)"
        Exit Sub
    End If

    Dim oPt As Point3d, dZ As Double, sSep As String, nDec As Integer
    If Not LireTexteAltitude(oText, oPt, dZ, sSep, nDec) Then
        ShowPrompt "Texte '" & oText.Text & "' non numerique - recliquez (P2)"
        Exit Sub
    End If

    If Dist2D(g_oP1, oPt) < 0.001 Then
        ShowPrompt "P1 et P2 confondus - choisissez un autre point (P2)"
        Exit Sub
    End If

    g_oP2 = oPt
    g_dZ2 = dZ
    If nDec > g_nDecimals Then g_nDecimals = nDec

    ShowPrompt "P2 = " & oText.Text & "  |  " & InfoSegment() & _
               "  ->  deplacez le curseur. Data=creer  Reset=nouv. selection"
    CommandState.StartPrimitive New CPlacerPoint
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

' Reset pendant le choix de P2 = revenir au choix de P1
Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Retour au POINT 1 : cliquez pres du texte altitude  (Reset=quitter)"
    CommandState.StartPrimitive New CSelectP1
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
