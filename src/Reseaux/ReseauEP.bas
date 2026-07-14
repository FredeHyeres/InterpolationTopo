Attribute VB_Name = "ReseauEP"
'==============================================================================
' ReseauEP - Module reseau EU/EP : pente entre deux regards/grilles
'
' Flux en 4 clics + 1 (5 etapes) :
'   1) CSnapBoiteP1   : snap sur l'axe P1 (tampon/grille)  -> g_ptSnapP1
'   2) CSelectBoiteP1 : clic pres de la cellule qui donne le Fe de P1
'                       -> g_oBoiteP1 (T:/Fe:/Prof:), g_oRefFeP1 = (P1 snap, Fe)
'   3) CSnapBoiteP2   : snap sur l'axe P2 (dynamics : ligne P1_snap -> curseur)
'                       -> g_ptSnapP2
'   4) CSelectBoiteP2 : clic pres de la cellule Fe de P2
'                       -> ligne definitive P1_snap -> P2_snap au niveau Fe
'   5) CPlacerPenteReseau : texte pente + fleche (comme V2 PlacerPente)
'
' Le point du trace et du calcul est le point SNAPPE, pas l'origine de la
' cellule (des reseaux proches peuvent rendre la cellule la plus proche
' ambigue : le clic explicite sur la cellule leve l'ambiguite).
'
' Le module s'appuie sur g_oCalc/g_oMoteur/g_oSettings deja fournis par
' InterpolationTopoV2 (fonction InitialiserContexte) : on ne duplique ni le
' calcul de pente ni le rendu de la fleche. On complete avec g_oBoiteP1/P2
' (donnees reseau) et deux CPointRef alimentes par les Fe pour reutiliser
' g_oMoteur.CreerPente / DessinPenteDynamique.
'
' LANCEMENT :
'   key-in : vba run [InterpolationTopoV2]PenteReseauEP
'==============================================================================
Option Explicit

' --- Instances partagees entre les classes de commande reseau ---
Public g_oBoiteP1        As CBoiteEPInfo
Public g_oBoiteP2        As CBoiteEPInfo
Public g_oParserBoite    As CBoiteEPParser
Public g_oRefFeP1        As CPointRef   ' alimente avec (snap XY, Fe) pour g_oMoteur
Public g_oRefFeP2        As CPointRef
Public g_dTolReseau      As Double      ' rayon de recherche autour du clic (u.m.),
                                        ' modifiable via frmReseauEP

' --- Points snappes (positions exactes cliquees, servent au trace) ---
' Publics dans un .bas : les UDT Point3d sont autorises hors modules de classe.
Public g_ptSnapP1        As Point3d
Public g_ptSnapP2        As Point3d

'------------------------------------------------------------------------------
' Point d'entree de la commande Pente Reseau EU/EP
Sub PenteReseauEP()
    If Not EnvironnementPret("Pente Reseau EU/EP") Then Exit Sub

    ' Contexte commun (fournit g_oSettings/g_oCalc/g_oMoteur)
    InitialiserContexte

    ' Contexte specifique reseau
    Set g_oParserBoite = New CBoiteEPParser
    Set g_oBoiteP1 = New CBoiteEPInfo
    Set g_oBoiteP2 = New CBoiteEPInfo
    Set g_oRefFeP1 = New CPointRef
    Set g_oRefFeP2 = New CPointRef
    g_dTolReseau = 5#

    ' Pentes reseau : 1 seule decimale (pour V2 le defaut reste 2)
    g_oSettings.oIndicPente.Decimales = 1

    AfficherFormulaire frmReseauEP

    CommandState.StartPrimitive New CSnapBoiteP1
End Sub

'==============================================================================
' Recherche de la cellule regard/grille la plus proche du clic
'==============================================================================

'------------------------------------------------------------------------------
' Scanne les CellHeader du modele actif et retient la plus proche dont le
' parsing renvoie un Fe valide. Renvoie True si trouve, avec oResult peuple.
Public Function TrouverBoiteEPProche(oPt As Point3d, dRayon As Double, _
                                     oResult As CBoiteEPInfo) As Boolean
    TrouverBoiteEPProche = False
    If oResult Is Nothing Then Set oResult = New CBoiteEPInfo
    oResult.Vider

    Dim oScan As New ElementScanCriteria
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeCellHeader

    Dim oEnum As ElementEnumerator
    On Error Resume Next
    Set oEnum = ActiveModelReference.Scan(oScan)
    If Err.Number <> 0 Or oEnum Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim dMin As Double: dMin = dRayon
    Dim oBest As CBoiteEPInfo: Set oBest = Nothing

    Do While oEnum.MoveNext
        Dim oElem As Element
        Set oElem = oEnum.Current
        If oElem.Type <> msdElementTypeCellHeader Then GoTo Suivant

        Dim oInfo As CBoiteEPInfo
        Set oInfo = g_oParserBoite.Lire(oElem)
        If Not oInfo.Valide Then GoTo Suivant

        Dim ptOrig As Point3d
        oInfo.CopierOrigine ptOrig
        Dim dD As Double
        dD = Dist2DLoc(oPt, ptOrig)
        If dD < dMin Then
            dMin = dD
            Set oBest = oInfo
        End If
Suivant:
    Loop

    If oBest Is Nothing Then Exit Function
    oResult.CopierDepuis oBest
    TrouverBoiteEPProche = True
End Function

'------------------------------------------------------------------------------
' Rayon de recherche courant (partage entre les etats, modifiable via UI).
Public Function TolReseau() As Double
    If g_dTolReseau <= 0 Then g_dTolReseau = 5#
    TolReseau = g_dTolReseau
End Function

'------------------------------------------------------------------------------
Private Function Dist2DLoc(oA As Point3d, oB As Point3d) As Double
    Dist2DLoc = Sqr((oB.X - oA.X) ^ 2 + (oB.Y - oA.Y) ^ 2)
End Function

'==============================================================================
' Utilitaires partages par les classes d'etat reseau
'==============================================================================

'------------------------------------------------------------------------------
' Charge un CPointRef avec les coordonnees XY du point snappe et l'altitude Fe
' de la cellule associee. Le point de reference sert au calcul de pente et au
' trace : c'est bien le snap qui compte, pas l'origine de la cellule.
Public Sub AlimenterPointRefFeSnap(oRef As CPointRef, _
                                    ptSnap As Point3d, oInfo As CBoiteEPInfo)
    oRef.DefinirPosition ptSnap
    oRef.Altitude = oInfo.ZFilEau
    oRef.Valide = oInfo.TrouveFilEau
End Sub

'------------------------------------------------------------------------------
' Texte court d'affichage d'une info reseau (utilise dans les prompts).
Public Function DecrireBoite(oInfo As CBoiteEPInfo) As String
    Dim s As String
    s = "Fe=" & Format$(oInfo.ZFilEau, "0.000")
    If oInfo.TrouveTampon Then s = s & "  T=" & Format$(oInfo.ZTampon, "0.000")
    If oInfo.TrouveProfondeur Then s = s & "  Prof=" & Format$(oInfo.Profondeur, "0.00")
    DecrireBoite = s
End Function
