Attribute VB_Name = "InterpolPonctuelle"
'==============================================================================
' InterpolPonctuelle - Point d'entree de la commande Interpolation Ponctuelle
'
' L'initialisation commune (globals, formulaire) est dans
' InterpolationTopoV2.bas.
'
' LANCEMENT :
'   key-in : vba run [InterpolationTopoV2]InterpolPonctuelle
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
Sub InterpolPonctuelle()
    If Not EnvironnementPret("Interpol. Ponctuelle") Then Exit Sub
    InitialiserContexte
    AfficherFormulaire frmInterpolPonct

    Dim oEtat As New CSelectP1
    oEtat.Mode = modePonctuelle
    CommandState.StartPrimitive oEtat
End Sub
