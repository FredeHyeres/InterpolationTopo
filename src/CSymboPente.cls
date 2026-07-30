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
'
' Le texte pente n'a pas de reglages de taille/couleur dedies : sa taille et
' sa couleur viennent du style de texte nomme (NomStyle), ou du style actif
' si NomStyle est vide. Pour un rendu different, l'utilisateur cree un
' nouveau style de texte plutot que de personnaliser au cas par cas.
'==============================================================================
Option Explicit

Public NomNiveau        As String    ' niveau du texte + fleche pente ("" = niveau actif/P1)
Public NomStyle         As String    ' style de texte nomme du texte pente ("" = style actif)
Public FlecheLongueur   As Double    ' longueur du corps de la fleche (u.m.)
Public FlecheCouleur    As Long      ' couleur de la fleche (index 0-255)
Public Decimales        As Integer   ' decimales du texte pente (defaut 1)
Public PointeLongueur   As Double    ' longueur des 2 lignes de la pointe (u.m.)
Public PointeFermee     As Boolean   ' True = pointe en triangle plein, False = 2 lignes ouvertes

'------------------------------------------------------------------------------
Sub InitDefauts()
    NomNiveau = ""
    NomStyle = ""
    FlecheLongueur = 1#
    FlecheCouleur = 3
    Decimales = 1
    PointeLongueur = 0.2
    PointeFermee = False
End Sub
