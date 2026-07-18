VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMstSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CMstSettings - Parametres de la commande Interpolation Topo V2
'
' Agregateur unique des parametres, regroupes par domaine :
'   oTexte      : symbologie et format du texte altitude
'   oCercle     : symbologie du cercle
'   oIndicPente : indicateur de pente (texte + fleche)
'   oDecalage   : pente transversale + DZ fixe (interpolation)
'   oChemin     : pente + DZ du chemin principal (ponctuelle)
'   oRayon      : pente + DZ du rayonnement (ponctuelle)
'
' Instancie une seule fois dans le module principal.
' Passe en parametre aux classes qui en ont besoin (CMoteurGraphique, etc.).
'==============================================================================
Option Explicit

Public oTexte           As CSymboTexte      ' symbologie et format du texte
Public oCercle          As CSymboCercle     ' symbologie du cercle
Public oIndicPente      As CSymboPente      ' indicateur de pente (texte + fleche)
Public oDecalage        As CParamPenteDZ    ' interpolation : pente transversale + DZ fixe
Public oChemin          As CParamPenteDZ    ' ponctuelle : chemin principal
Public oRayon           As CParamPenteDZ    ' ponctuelle : rayonnement
Public dTolTexte        As Double           ' rayon de recherche autour du clic (u.m.)
Public oTextModele      As TextElement      ' element P1 : modele pour CreateTextElement1
Public oTagModele       As TagElement       ' tag P1 : modele quand la source est un tag
Public dModeleHauteur   As Double           ' hauteur du texte source (0 = non defini)
Public nPonctDecimales  As Integer          ' ponctuelle : decimales des points crees (chemin + rayon)
Public bRayonIndicPente As Boolean          ' ponctuelle : placer pente+fleche apres chaque rayon
Public bCheminIndicPente As Boolean         ' ponctuelle : placer pente+fleche apres chaque point chemin
Public bDecalIndicPente As Boolean          ' interpolation : placer pente+fleche apres un point decale (pente transversale)

'------------------------------------------------------------------------------
Sub Init()
    Set oTexte = New CSymboTexte
    Set oCercle = New CSymboCercle
    oCercle.InitDefauts
    Set oIndicPente = New CSymboPente
    oIndicPente.InitDefauts
    Set oDecalage = New CParamPenteDZ
    Set oChemin = New CParamPenteDZ
    Set oRayon = New CParamPenteDZ
    dTolTexte = 1#
    Set oTextModele = Nothing
    Set oTagModele = Nothing
    dModeleHauteur = 0#
    nPonctDecimales = 2
End Sub

'------------------------------------------------------------------------------
' Appele apres la selection de P1 : charge le texte-modele et ses attributs.
' L'indicateur de pente herite de la hauteur/largeur/couleur du texte P1.
Sub ChargerModeleTexte(oText As TextElement, sSep As String, nDec As Integer)
    Set oTextModele = oText
    Set oTagModele = Nothing
    oTexte.ChargerDepuisElement oText
    oTexte.SepDecimal = sSep
    oTexte.Decimales = nDec
    On Error Resume Next
    Dim oTemp As TextElement
    Dim ptO As Point3d
    Set oTemp = CreateTextElement1(oText, "0", ptO, Matrix3dIdentity())
    If Not oTemp Is Nothing Then
        dModeleHauteur = CDbl(oTemp.TextStyle.Height)
        oIndicPente.Largeur = CDbl(oTemp.TextStyle.Width)
    Else
        Err.Clear
        dModeleHauteur = CDbl(oText.TextStyle.Height)
        oIndicPente.Largeur = CDbl(oText.TextStyle.Width)
    End If
    oIndicPente.Hauteur = dModeleHauteur
    If oIndicPente.Largeur <= 0 Then oIndicPente.Largeur = dModeleHauteur
    oIndicPente.Couleur = oText.Color
    If Err.Number <> 0 Then
        dModeleHauteur = 0#
        oIndicPente.Hauteur = 0.1
        oIndicPente.Largeur = 0.1
    End If
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Appele quand la source P1 est un tag : estime la hauteur depuis le Range.
Sub ChargerModeleDepuisTag(oTag As TagElement, sSep As String, nDec As Integer)
    Set oTagModele = oTag
    oTexte.SepDecimal = sSep
    oTexte.Decimales = nDec
    dModeleHauteur = 0#
    On Error Resume Next
    Dim oRng As Range3d
    oRng = oTag.Range
    Dim ptL As Point3d, ptH As Point3d
    ptL = oRng.Low: ptH = oRng.High
    Dim dHY As Double, dHX As Double
    dHY = Abs(ptH.Y - ptL.Y)
    dHX = Abs(ptH.X - ptL.X)
    If dHY > 0 And dHX > 0 Then
        If dHX < dHY Then dModeleHauteur = dHX Else dModeleHauteur = dHY
    End If
    If dModeleHauteur > 0 Then
        oIndicPente.Hauteur = dModeleHauteur
        oIndicPente.Largeur = dModeleHauteur
    End If
    oIndicPente.Couleur = oTag.Color
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
Sub AjusterDecimales(nDecP2 As Integer)
    If nDecP2 > oTexte.Decimales Then oTexte.Decimales = nDecP2
End Sub

'------------------------------------------------------------------------------
' Vrai si le texte-modele est disponible (P1 a ete selectionne)
Property Get TextModeleDisponible() As Boolean
    TextModeleDisponible = Not (oTextModele Is Nothing)
End Property
