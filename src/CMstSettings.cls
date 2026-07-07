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
Public dPenteDecalage   As Double           ' pente transversale en % (appliquee au decalage)
Public bAppliquerPente  As Boolean          ' True = appliquer la pente au decalage

'------------------------------------------------------------------------------
Sub Init()
    Set oTexte = New CSymboTexte
    Set oCercle = New CSymboCercle
    oCercle.InitDefauts
    dTolTexte = 1#
    Set oTextModele = Nothing
    dPenteDecalage = 0#
    bAppliquerPente = False
End Sub

'------------------------------------------------------------------------------
' Appele apres la selection de P1 : charge le texte-modele et ses attributs.
' nDecP2 permet de mettre a jour les decimales si P2 en a davantage.
Sub ChargerModeleTexte(oText As TextElement, sSep As String, nDec As Integer)
    Set oTextModele = oText
    oTexte.ChargerDepuisElement oText
    oTexte.SepDecimal = sSep
    oTexte.Decimales = nDec
End Sub

'------------------------------------------------------------------------------
' Appele apres la selection de P2 : ajuste les decimales si besoin.
Sub ChargerModeleDepuisTag(sSep As String, nDec As Integer)
    oTexte.SepDecimal = sSep
    oTexte.Decimales = nDec
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
