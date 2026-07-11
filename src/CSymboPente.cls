VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSymboPente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSymboPente - Attributs de l'indicateur de pente (texte + fleche)
'
' Meme role que CSymboTexte / CSymboCercle : simple sac de proprietes,
' sans dependance a l'API MicroStation.
' Hauteur / Largeur / Couleur sont recalees sur le texte-modele P1 par
' CMstSettings.ChargerModeleTexte, sauf si Perso est actif.
'==============================================================================
Option Explicit

Public Perso            As Boolean   ' True = parametres personnalises (sinon herites de P1)
Public Hauteur          As Double    ' hauteur du texte pente
Public Largeur          As Double    ' largeur du texte pente
Public Couleur          As Long      ' couleur du texte pente (index 0-255)
Public NomNiveau        As String    ' niveau du texte pente ("" = niveau actif)
Public FlecheLongueur   As Double    ' longueur de la fleche (u.m.)
Public FlecheCouleur    As Long      ' couleur de la fleche (index 0-255)

'------------------------------------------------------------------------------
Sub InitDefauts()
    Perso = False
    Hauteur = 0.1
    Largeur = 0.1
    Couleur = 0
    NomNiveau = ""
    FlecheLongueur = 1#
    FlecheCouleur = 3
End Sub
