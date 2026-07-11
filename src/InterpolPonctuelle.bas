Attribute VB_Name = "InterpolPonctuelle"
'==============================================================================
' InterpolPonctuelle - Point d'entree de la commande Interpolation Ponctuelle
'
' LANCEMENT :
'   key-in : vba run [InterpolationTopoV2]InterpolPonctuelle
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
Sub InterpolPonctuelle()
    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, "Interpol. Ponctuelle"
        Exit Sub
    End If

    Set g_oSettings = New CMstSettings
    g_oSettings.Init

    Set g_oP1 = New CPointRef
    Set g_oP2 = New CPointRef
    Set g_oCalc = New CInterpolation
    Set g_oMoteur = New CMoteurGraphique
    Set g_oSelectionCourante = New CAltitudeSelection
    Set g_oSelectionP1 = New CAltitudeSelection
    Set g_oSelectionP2 = New CAltitudeSelection

    frmInterpolPonct.Initialiser g_oSettings
    frmInterpolPonct.StartUpPosition = 0
    frmInterpolPonct.Left = Application.Width * 0.6
    frmInterpolPonct.Top = Application.Height * 0.05
    frmInterpolPonct.Show vbModeless

    CommandState.StartPrimitive New CSelectP1Ponct
End Sub
