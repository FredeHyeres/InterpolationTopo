VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAltitudeSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CAltitudeSelection - Resultat d'une recherche d'altitude
'
' Memorise la valeur altitude trouvee, son origine dans le repere du modele
' actif, l'element source eventuel, et indique si la source vient d'une
' reference attachee. Compatible VBA 6.x : pas de membre Public de type Point3d.
'==============================================================================
Option Explicit

Public ValeurTexte      As String
Public TagDefName       As String
Public EstReference     As Boolean

Public oTexte           As TextElement
Public oTag             As TagElement
Public oCellule         As CellElement
Public oAttachment      As Object

Private m_dX            As Double
Private m_dY            As Double
Private m_dZ            As Double

'------------------------------------------------------------------------------
Sub Vider()
    ValeurTexte = ""
    TagDefName = ""
    EstReference = False
    Set oTexte = Nothing
    Set oTag = Nothing
    Set oCellule = Nothing
    Set oAttachment = Nothing
    m_dX = 0#: m_dY = 0#: m_dZ = 0#
End Sub

'------------------------------------------------------------------------------
Sub DefinirOrigine(oPt As Point3d)
    m_dX = oPt.X
    m_dY = oPt.Y
    m_dZ = oPt.Z
End Sub

'------------------------------------------------------------------------------
Sub CopierOrigine(oPt As Point3d)
    oPt.X = m_dX
    oPt.Y = m_dY
    oPt.Z = m_dZ
End Sub

'------------------------------------------------------------------------------
Sub CopierDepuis(oSrc As CAltitudeSelection)
    If oSrc Is Nothing Then
        Vider
        Exit Sub
    End If

    ValeurTexte = oSrc.ValeurTexte
    TagDefName = oSrc.TagDefName
    EstReference = oSrc.EstReference
    Set oTexte = oSrc.oTexte
    Set oTag = oSrc.oTag
    Set oCellule = oSrc.oCellule
    Set oAttachment = oSrc.oAttachment

    Dim oPt As Point3d
    oSrc.CopierOrigine oPt
    DefinirOrigine oPt
End Sub

'------------------------------------------------------------------------------
Property Get Trouvee() As Boolean
    Trouvee = (Len(ValeurTexte) > 0)
End Property
