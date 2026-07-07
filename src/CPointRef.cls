VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPointRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPointRef - Point de reference topographique (P1 ou P2)
'
' Contient les donnees geometriques et altimetriques d un point connu.
'
' NOTE VBA : un membre Public de type UDT (Point3d) est interdit dans un
' module de classe (erreur de compilation). Les coordonnees sont donc
' stockees en trois Double ; la conversion vers Point3d se fait par des Sub
' avec parametre de sortie (jamais de retour UDT sur ce poste).
'
' Le parsing du texte est delegue a CInterpolation.
'==============================================================================
Option Explicit

Public X            As Double   ' coordonnees du point (origin du texte)
Public Y            As Double
Public Z            As Double
Public Altitude     As Double   ' altitude numerique
Public Valide       As Boolean  ' True si le point a ete correctement charge

'------------------------------------------------------------------------------
Sub Invalider()
    Valide = False
    Altitude = 0#
End Sub

'------------------------------------------------------------------------------
' Charge la position depuis un Point3d (ex. origin du texte)
Sub DefinirPosition(oPt As Point3d)
    X = oPt.X: Y = oPt.Y: Z = oPt.Z
End Sub

'------------------------------------------------------------------------------
' Copie la position dans le Point3d fourni par l'appelant
Sub CopierPosition(oPt As Point3d)
    oPt.X = X: oPt.Y = Y: oPt.Z = Z
End Sub
