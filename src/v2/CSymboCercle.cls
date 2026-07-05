VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSymboCercle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSymboCercle - Attributs de symbologie du cercle topo
'
' Configures par l'utilisateur via frmInterpolation (Tool Settings).
' Persistent pendant toute la duree de la commande.
'==============================================================================
Option Explicit

Public Rayon        As Double   ' demi-diametre en unites maitre
Public Couleur      As Long     ' index couleur MicroStation 0-255
Public NomNiveau    As String   ' "" = meme niveau que le texte P1

'------------------------------------------------------------------------------
' Valeurs par defaut au demarrage de la commande
Sub InitDefauts()
    Rayon = 0.05    ' diametre = 0.10
    Couleur = 2     ' vert
    NomNiveau = ""  ' niveau du texte P1
End Sub

'------------------------------------------------------------------------------
' Diametre (propriete calculee pour l'affichage dans le formulaire)
Property Get Diametre() As Double
    Diametre = Rayon * 2#
End Property

Property Let Diametre(dDiam As Double)
    Rayon = dDiam / 2#
End Property
