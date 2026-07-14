Attribute VB_Name = "ReseauEP"
'==============================================================================
' ReseauEP - Module reseau EU/EP : pente entre deux cellules regard/grille
'
' Flux :
'   1) L'utilisateur clique/snappe le P1 (axe d'un tampon ou d'une grille)
'      -> scan des CellHeader dans le rayon, on garde la plus proche dont le
'         parsing T:/Fe:/Prof: renvoie au moins un Fe (fil d'eau).
'   2) Ligne provisoire P1 -> curseur pendant le choix de P2 (dynamics).
'   3) Meme detection pour P2. Si P1 et P2 valides :
'      -> ligne definitive P1-P2 ajoutee au modele
'      -> calcul pente entre les Fe :  P = (Fe2 - Fe1) / d2D * 100
'      -> passage a CPlacerPenteReseau qui pose texte pente + fleche
'         (dans le sens de la descente) au point clique.
'
' Le module s'appuie sur g_oCalc/g_oMoteur/g_oSettings deja fournis par
' InterpolationTopoV2 (fonction InitialiserContexte) : on ne duplique pas les
' calculs de pente ni le rendu de la fleche. On complete avec g_oBoiteP1/P2
' (nos donnees reseau) et deux CPointRef alimentes par les Fe pour reutiliser
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
Public g_oRefFeP1        As CPointRef   ' alimente avec Fe pour reutiliser g_oMoteur
Public g_oRefFeP2        As CPointRef

Private Const TOL_RESEAU As Double = 5#   ' rayon de recherche par defaut (u.m.)

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

    CommandState.StartPrimitive New CSelectBoiteP1
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
' Rayon de recherche par defaut (partage entre les etats snap).
Public Function TolReseau() As Double
    TolReseau = TOL_RESEAU
End Function

'------------------------------------------------------------------------------
Private Function Dist2DLoc(oA As Point3d, oB As Point3d) As Double
    Dist2DLoc = Sqr((oB.X - oA.X) ^ 2 + (oB.Y - oA.Y) ^ 2)
End Function

'==============================================================================
' Utilitaires partages par les classes d'etat reseau
'==============================================================================

'------------------------------------------------------------------------------
' Charge un CPointRef depuis une CBoiteEPInfo : XY = origine cellule,
' Altitude = Fe (fil d'eau).
Public Sub AlimenterPointRefFe(oRef As CPointRef, oInfo As CBoiteEPInfo)
    Dim pt As Point3d
    oInfo.CopierOrigine pt
    oRef.DefinirPosition pt
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
