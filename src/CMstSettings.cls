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
' Agregateur unique des attributs texte, cercle, tolerance et texte-modele.
' Instancie une seule fois dans le module principal.
' Passe en parametre aux classes qui en ont besoin (CMoteurGraphique, etc.).
'==============================================================================
Option Explicit

Public oTexte           As CSymboTexte      ' symbologie et format du texte
Public oCercle          As CSymboCercle     ' symbologie du cercle
Public dTolTexte        As Double           ' rayon de recherche autour du clic (u.m.)
Public oTextModele      As TextElement      ' element P1 : modele pour CreateTextElement1
Public oTagModele       As TagElement       ' tag P1 : modele quand la source est un tag
Public dPenteDecalage   As Double           ' pente transversale en % (appliquee au decalage)
Public bAppliquerPente  As Boolean          ' True = appliquer la pente au decalage
Public dDecalageDZ      As Double           ' decalage fixe en Z (ajoute a l'altitude interpolee)
Public bDecalageFixe    As Boolean          ' True = appliquer le decalage fixe DZ
Public dModeleHauteur   As Double           ' hauteur du texte source (0 = non defini)
Public bPentePerso      As Boolean          ' True = utiliser parametres personnalises
Public dPenteHauteur    As Double           ' hauteur du texte pente
Public dPenteLargeur    As Double           ' largeur du texte pente
Public nPenteCouleur    As Long             ' couleur du texte pente (index 0-255)
Public sPenteNiveau     As String           ' niveau du texte pente
Public dFlecheLongueur  As Double           ' longueur de la fleche indicatrice de pente (u.m.)
Public nFlecheCouleur   As Long             ' couleur de la fleche (index MicroStation 0-255)

' --- Interpolation ponctuelle : chemin principal ---
Public dCheminPente     As Double           ' pente chemin principal (%)
Public bCheminPente     As Boolean          ' True = appliquer pente chemin
Public dCheminDZ        As Double           ' DZ fixe chemin principal
Public bCheminDZ        As Boolean          ' True = appliquer DZ chemin
' --- Interpolation ponctuelle : rayonnement ---
Public dRayonPente      As Double           ' pente rayonnement (%)
Public bRayonPente      As Boolean          ' True = appliquer pente rayonnement
Public dRayonDZ         As Double           ' DZ fixe rayonnement
Public bRayonDZ         As Boolean          ' True = appliquer DZ rayonnement
' --- Interpolation ponctuelle : decimales (commun chemin+rayon) ---
Public nPonctDecimales  As Integer          ' nombre de decimales pour les points crees

'------------------------------------------------------------------------------
Sub Init()
    Set oTexte = New CSymboTexte
    Set oCercle = New CSymboCercle
    oCercle.InitDefauts
    dTolTexte = 1#
    Set oTextModele = Nothing
    Set oTagModele = Nothing
    dModeleHauteur = 0#
    bPentePerso = False
    dPenteHauteur = 0.1
    dPenteLargeur = 0.1
    nPenteCouleur = 0
    sPenteNiveau = ""
    dPenteDecalage = 0#
    bAppliquerPente = False
    dDecalageDZ = 0#
    bDecalageFixe = False
    dFlecheLongueur = 1#
    nFlecheCouleur = 3
    dCheminPente = 0#
    bCheminPente = False
    dCheminDZ = 0#
    bCheminDZ = False
    dRayonPente = 0#
    bRayonPente = False
    dRayonDZ = 0#
    bRayonDZ = False
    nPonctDecimales = 2
End Sub

'------------------------------------------------------------------------------
' Appele apres la selection de P1 : charge le texte-modele et ses attributs.
' nDecP2 permet de mettre a jour les decimales si P2 en a davantage.
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
        dPenteLargeur = CDbl(oTemp.TextStyle.Width)
    Else
        Err.Clear
        dModeleHauteur = CDbl(oText.TextStyle.Height)
        dPenteLargeur = CDbl(oText.TextStyle.Width)
    End If
    dPenteHauteur = dModeleHauteur
    If dPenteLargeur <= 0 Then dPenteLargeur = dModeleHauteur
    nPenteCouleur = oText.Color
    If Err.Number <> 0 Then
        dModeleHauteur = 0#
        dPenteHauteur = 0.1
        dPenteLargeur = 0.1
    End If
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Appele apres la selection de P2 : ajuste les decimales si besoin.
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
        dPenteHauteur = dModeleHauteur
        dPenteLargeur = dModeleHauteur
    End If
    nPenteCouleur = oTag.Color
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
