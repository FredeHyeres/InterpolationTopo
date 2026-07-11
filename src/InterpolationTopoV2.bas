Attribute VB_Name = "InterpolationTopoV2"
'==============================================================================
' InterpolationTopoV2 - Module principal - MicroStation V8i SS3 (VBA)
'
' Interpole l'altitude d'un point situe sur la droite reliant deux textes
' d'altitude existants, avec apercu dynamique, puis cree :
'   - un texte altitude (memes proprietes que le texte P1, pris comme modele)
'   - un cercle centre sur le point projete
'
' Architecture V2 : classes a responsabilites separees + formulaire modeless
'
' Ce module ne contient que les globals partages, l'initialisation commune et
' le point d'entree de la commande Interpolation. La recherche des altitudes
' sources est dans RechercheAltitude.bas ; le point d'entree de la commande
' Interpolation Ponctuelle est dans InterpolPonctuelle.bas.
'
' LANCEMENT :
'   key-in : vba run [InterpolationTopoV2]InterpolerPoint
'==============================================================================
Option Explicit

' --- Mode de commande ---
' Les classes d'etat (CSelectP1/P2, CSnapP1/P2) sont partagees entre les deux
' commandes : leur propriete Mode choisit formulaire, invites et etat suivant.
' modeInterpolation = 0 : valeur par defaut d'une instance creee sans Mode.
Public Enum eModeCommande
    modeInterpolation = 0
    modePonctuelle = 1
End Enum

' --- Instances partagees entre les classes de commande ---
' Chaque classe recoit ces references mais ne les cree pas.
Public g_oSettings  As CMstSettings      ' parametres (symbologie, tolerance)
Public g_oP1        As CPointRef         ' point de reference 1
Public g_oP2        As CPointRef         ' point de reference 2
Public g_oCalc      As CInterpolation    ' moteur de calcul pur
Public g_oMoteur    As CMoteurGraphique  ' moteur de creation graphique
Public g_oSelectionCourante As CAltitudeSelection ' dernier texte/tag/cellule trouve
Public g_oSelectionP1       As CAltitudeSelection ' source modele de creation
Public g_oSelectionP2       As CAltitudeSelection ' source altitude P2

'------------------------------------------------------------------------------
' Point d'entree de la commande Interpolation
Sub InterpolerPoint()
    If Not EnvironnementPret("Interpolation Topo") Then Exit Sub
    InitialiserContexte
    AfficherFormulaire frmInterpolation

    Dim oEtat As New CSelectP1
    oEtat.Mode = modeInterpolation
    CommandState.StartPrimitive oEtat
End Sub

'==============================================================================
' Initialisation commune aux deux commandes
'==============================================================================

'------------------------------------------------------------------------------
' Verifie qu'un fichier DGN est ouvert. Affiche un message sinon.
Public Function EnvironnementPret(sTitre As String) As Boolean
    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, sTitre
        EnvironnementPret = False
    Else
        EnvironnementPret = True
    End If
End Function

'------------------------------------------------------------------------------
' Instanciation unique des objets partages par les classes de commande.
Public Sub InitialiserContexte()
    Set g_oSettings = New CMstSettings
    g_oSettings.Init

    Set g_oP1 = New CPointRef
    Set g_oP2 = New CPointRef
    Set g_oCalc = New CInterpolation
    Set g_oMoteur = New CMoteurGraphique
    Set g_oSelectionCourante = New CAltitudeSelection
    Set g_oSelectionP1 = New CAltitudeSelection
    Set g_oSelectionP2 = New CAltitudeSelection
End Sub

'------------------------------------------------------------------------------
' Initialise, positionne et affiche un formulaire modeless (Tool Settings).
' oFrm en Object : accepte frmInterpolation comme frmInterpolPonct.
Public Sub AfficherFormulaire(oFrm As Object)
    oFrm.Initialiser g_oSettings
    oFrm.StartUpPosition = 0
    oFrm.Left = Application.Width * 0.6
    oFrm.Top = Application.Height * 0.05
    oFrm.Show vbModeless
End Sub
