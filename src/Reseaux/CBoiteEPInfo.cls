VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBoiteEPInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CBoiteEPInfo - Resultat du parsing d'une cellule regard/grille (BoiteEP)
'
' Contient les trois altitudes attendues :
'   T:    altitude du tampon         -> ZTampon      (flag TrouveTampon)
'   Fe:   altitude du fil d'eau      -> ZFilEau      (flag TrouveFilEau)
'   Prof: profondeur du regard       -> Profondeur   (flag TrouveProfondeur)
'
' Et l'element source + son origine dans le repere du modele actif.
' Compatible VBA 6.x : pas de membre Public de type Point3d (stocke en 3 Double).
'==============================================================================
Option Explicit

Public TrouveTampon         As Boolean
Public ZTampon              As Double

Public TrouveFilEau         As Boolean
Public ZFilEau              As Double

Public TrouveProfondeur     As Boolean
Public Profondeur           As Double

Public oCellule             As CellElement
Public NomCellule           As String
Public EstReference         As Boolean       ' True si la cellule vient d'un fichier attache
Public oAttachment          As Object        ' Attachment source (Nothing si modele actif)

Private m_dX                As Double
Private m_dY                As Double
Private m_dZ                As Double

'------------------------------------------------------------------------------
Sub Vider()
    TrouveTampon = False:      ZTampon = 0#
    TrouveFilEau = False:      ZFilEau = 0#
    TrouveProfondeur = False:  Profondeur = 0#
    Set oCellule = Nothing
    NomCellule = ""
    EstReference = False
    Set oAttachment = Nothing
    m_dX = 0#: m_dY = 0#: m_dZ = 0#
End Sub

'------------------------------------------------------------------------------
' On considere l'info valide si au moins le fil d'eau a ete trouve : c'est
' l'altitude utilisee pour le calcul de pente.
Property Get Valide() As Boolean
    Valide = TrouveFilEau
End Property

'------------------------------------------------------------------------------
Sub DefinirOrigine(oPt As Point3d)
    m_dX = oPt.X: m_dY = oPt.Y: m_dZ = oPt.Z
End Sub

Sub CopierOrigine(oPt As Point3d)
    oPt.X = m_dX: oPt.Y = m_dY: oPt.Z = m_dZ
End Sub

Property Get X() As Double: X = m_dX: End Property
Property Get Y() As Double: Y = m_dY: End Property
Property Get Z() As Double: Z = m_dZ: End Property

'------------------------------------------------------------------------------
Sub CopierDepuis(oSrc As CBoiteEPInfo)
    If oSrc Is Nothing Then Vider: Exit Sub
    TrouveTampon = oSrc.TrouveTampon:         ZTampon = oSrc.ZTampon
    TrouveFilEau = oSrc.TrouveFilEau:         ZFilEau = oSrc.ZFilEau
    TrouveProfondeur = oSrc.TrouveProfondeur: Profondeur = oSrc.Profondeur
    Set oCellule = oSrc.oCellule
    NomCellule = oSrc.NomCellule
    EstReference = oSrc.EstReference
    Set oAttachment = oSrc.oAttachment
    Dim oPt As Point3d
    oSrc.CopierOrigine oPt
    DefinirOrigine oPt
End Sub
