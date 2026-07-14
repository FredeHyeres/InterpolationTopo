Attribute VB_Name = "DebugBoiteEP"
Option Explicit

'======================================================================
' Outils de diagnostic pour valider la structure des cellules BoiteEP
' et le parsing des textes T:/Fe:/Prof: sur un DGN reel.
'
' Utilisation :
'   1) Ouvrir le DGN exemple.
'   2) Selectionner une cellule regard/grille.
'   3) Lancer : vba run [MaCommande]DebugInspecterSelection
'      -> affiche la structure de la premiere cellule selectionnee et
'         le resultat du parser sur cette cellule.
'======================================================================

Sub DebugInspecterSelection()
    Dim oSet As ElementEnumerator
    Set oSet = ActiveModelReference.GetSelectedElements

    If Not oSet.MoveNext Then
        MsgBox "Aucun element selectionne.", vbInformation, "Debug BoiteEP"
        Exit Sub
    End If

    Dim oElem As Element
    Set oElem = oSet.Current

    Dim sRapport As String
    sRapport = "Element de premier niveau :" & vbCrLf
    sRapport = sRapport & DescrireElement(oElem, 0) & vbCrLf

    If oElem.Type = msdElementTypeCellHeader Then
        sRapport = sRapport & vbCrLf & "Enfants :" & vbCrLf
        Dim oCell As CellElement
        Set oCell = oElem
        sRapport = sRapport & DescrireEnfants(oCell.GetSubElements, 1)
    End If

    ' Passer par le parser
    Dim oParser As New CBoiteEPParser
    Dim oInfo As CBoiteEPInfo
    Set oInfo = oParser.Lire(oElem)

    sRapport = sRapport & vbCrLf & "--- Parser ---" & vbCrLf
    sRapport = sRapport & "Fe trouve  : " & oInfo.TrouveFilEau & _
               "  Valeur = " & oInfo.ZFilEau & vbCrLf
    sRapport = sRapport & "T  trouve  : " & oInfo.TrouveTampon & _
               "  Valeur = " & oInfo.ZTampon & vbCrLf
    sRapport = sRapport & "Prof trouve: " & oInfo.TrouveProfondeur & _
               "  Valeur = " & oInfo.Profondeur & vbCrLf
    sRapport = sRapport & "Valide     : " & oInfo.Valide

    MsgBox sRapport, vbInformation, "Debug BoiteEP"
End Sub

'----------------------------------------------------------------------
Private Function DescrireElement(oElem As Element, ByVal nDepth As Long) As String
    Dim sIndent As String
    sIndent = String$(nDepth * 2, " ")

    Dim sType As String
    sType = NomType(oElem.Type)

    Dim sExtra As String
    On Error Resume Next
    If oElem.Type = msdElementTypeText Then
        sExtra = " [Text=""" & oElem.AsTextElement.Text & """]"
    ElseIf oElem.Type = msdElementTypeTag Then
        sExtra = " [TagName=""" & oElem.AsTagElement.TagDefinitionName & _
                 """ Value=""" & CStr(oElem.AsTagElement.Value) & """]"
    ElseIf oElem.Type = msdElementTypeCellHeader Then
        Dim oC As CellElement
        Set oC = oElem
        sExtra = " [Name=""" & oC.Name & """]"
    End If
    On Error GoTo 0

    DescrireElement = sIndent & sType & sExtra
End Function

'----------------------------------------------------------------------
Private Function DescrireEnfants(oEnum As ElementEnumerator, ByVal nDepth As Long) As String
    Dim s As String
    Do While oEnum.MoveNext
        Dim oElem As Element
        Set oElem = oEnum.Current
        s = s & DescrireElement(oElem, nDepth) & vbCrLf

        ' Descendre dans les cellules et les text nodes
        If oElem.Type = msdElementTypeCellHeader Then
            Dim oC As CellElement
            Set oC = oElem
            s = s & DescrireEnfants(oC.GetSubElements, nDepth + 1)
        ElseIf oElem.Type = msdElementTypeTextNode Then
            Dim oN As TextNodeElement
            Set oN = oElem.AsTextNodeElement
            s = s & DescrireEnfants(oN.GetSubElements, nDepth + 1)
        End If
    Loop
    DescrireEnfants = s
End Function

'----------------------------------------------------------------------
Private Function NomType(ByVal t As MsdElementType) As String
    Select Case t
    Case msdElementTypeCellHeader:  NomType = "CellHeader"
    Case msdElementTypeText:        NomType = "Text"
    Case msdElementTypeTextNode:    NomType = "TextNode"
    Case msdElementTypeTag:         NomType = "Tag"
    Case msdElementTypeLine:        NomType = "Line"
    Case msdElementTypeLineString:  NomType = "LineString"
    Case msdElementTypeShape:       NomType = "Shape"
    Case msdElementTypeArc:         NomType = "Arc"
    Case msdElementTypeEllipse:     NomType = "Ellipse"
    Case Else:                      NomType = "Type#" & t
    End Select
End Function
