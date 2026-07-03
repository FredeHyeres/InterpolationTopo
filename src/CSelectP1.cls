VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelectP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSelectP1 - etape 1 : selection du texte altitude du point P1
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpolation Topo"
    ShowPrompt "Cliquez pres du texte altitude du POINT 1"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim oText As TextElement
    Set oText = FindNearestText(Point, TOL_TEXTE)
    If oText Is Nothing Then
        ShowPrompt "Aucun texte altitude a moins de " & TOL_TEXTE & " u.m. - recliquez (P1)"
        Exit Sub
    End If

    Dim oPt As Point3d, dZ As Double, sSep As String, nDec As Integer
    If Not LireTexteAltitude(oText, oPt, dZ, sSep, nDec) Then
        ShowPrompt "Texte '" & oText.Text & "' non numerique - recliquez (P1)"
        Exit Sub
    End If

    g_oP1 = oPt
    g_dZ1 = dZ
    g_sDecSep = sSep
    g_nDecimals = nDec
    Set g_oTextTemplate = oText

    ShowPrompt "P1 = " & oText.Text & "  ->  cliquez pres du texte altitude du POINT 2"
    CommandState.StartPrimitive New CSelectP2
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ShowPrompt "Interpolation Topo : annule."
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
