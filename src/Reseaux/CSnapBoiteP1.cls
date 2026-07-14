VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapBoiteP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSnapBoiteP1 - Etat 1 : snap du POINT 1 (axe tampon/grille)
'
'   Data  = memorise la position snappee dans g_ptSnapP1 -> passe au clic cellule.
'   Reset = quitte la commande.
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Pente Reseau EU/EP"
    frmReseauEP.ReinitialiserEtat
    ShowPrompt "Snappez le POINT 1 (axe tampon/grille)  (Reset = quitter)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    g_ptSnapP1 = Point
    ShowPrompt "P1 snappe  ->  cliquez pres de la cellule regard/grille du P1"
    CommandState.StartPrimitive New CSelectBoiteP1
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    frmReseauEP.Hide
    ShowPrompt "Pente Reseau EU/EP : termine."
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub
