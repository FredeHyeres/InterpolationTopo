VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSymboTexte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CSymboTexte - Attributs de symbologie du texte altitude
'
' CommeModele = True (defaut) : le texte cree reprend tous les attributs du
' texte P1 (via CreateTextElement1 avec modele) ; NomNiveau/Couleur refletent
' alors le modele, a titre informatif.
' CommeModele = False : NomNiveau/Couleur sont choisis par l'utilisateur dans
' le formulaire et appliques au texte cree (NomNiveau vide = niveau du P1).
'
' Stocke egalement la convention de formatage numerique (separateur, decimales)
' deduite du texte source.
'==============================================================================
Option Explicit

Public CommeModele  As Boolean  ' True = memes attributs que le texte P1
Public NomNiveau    As String   ' niveau du texte cree (nom ou numero)
Public Couleur      As Long     ' index couleur MicroStation 0-255

Public SepDecimal   As String   ' "." ou "," : separateur lu sur le texte source
Public Decimales    As Integer  ' nombre de decimales (max de P1 et P2)

'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    CommeModele = True
End Sub

'------------------------------------------------------------------------------
' Charge NomNiveau et Couleur depuis le TextElement P1 - uniquement en mode
' modele, pour ne pas ecraser un choix utilisateur.
' SepDecimal et Decimales sont renseignes par CInterpolation lors du parsing.
Sub ChargerDepuisElement(oText As TextElement)
    If Not CommeModele Then Exit Sub
    On Error Resume Next
    NomNiveau = oText.Level.Name
    Couleur = oText.Color
    On Error GoTo 0
End Sub
