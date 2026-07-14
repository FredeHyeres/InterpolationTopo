VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBoiteEPParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CBoiteEPParser - Extrait T: / Fe: / Prof: d'une cellule regard/grille
'
' Structure attendue (validee via DebugBoiteEP.bas sur cellule "Boite EU") :
'   CellHeader
'     Shape / Shape / ...
'     TextNode              <- les textes sont sous un TextNode
'       Text "T: 91.98"
'       Text "Fe:90.86"
'       Text "Prof:1.12"
'
' Peut aussi lire des tags si le futur DGN les utilise a la place des textes.
' Robustesse : espaces libres apres le prefixe, separateur decimal , ou .
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Point d'entree unique : lit une cellule et retourne CBoiteEPInfo peuple.
' oElem peut etre un CellHeader ; toute autre chose renvoie un info vide.
Function Lire(oElem As Element) As CBoiteEPInfo
    Dim oInfo As New CBoiteEPInfo
    oInfo.Vider

    If oElem Is Nothing Then Set Lire = oInfo: Exit Function
    If oElem.Type <> msdElementTypeCellHeader Then Set Lire = oInfo: Exit Function

    Dim oCell As CellElement
    Set oCell = oElem
    Set oInfo.oCellule = oCell
    oInfo.NomCellule = oCell.Name
    oInfo.DefinirOrigine oCell.Origin

    ParcourirEnfants oCell.GetSubElements, oInfo

    Set Lire = oInfo
End Function

'------------------------------------------------------------------------------
' Parcours recursif : descend dans les sous-cellules et les text nodes,
' extrait chaque texte/tag rencontre pour tenter d'y lire T:/Fe:/Prof:.
Private Sub ParcourirEnfants(oEnum As ElementEnumerator, oInfo As CBoiteEPInfo)
    Do While oEnum.MoveNext
        Dim oElem As Element
        Set oElem = oEnum.Current

        If oElem.IsTextElement Then
            AnalyserChaine oElem.AsTextElement.Text, oInfo
        ElseIf oElem.Type = msdElementTypeTag Then
            AnalyserChaine CStr(oElem.AsTagElement.Value), oInfo
        ElseIf oElem.Type = msdElementTypeTextNode Then
            ParcourirEnfants oElem.AsTextNodeElement.GetSubElements, oInfo
        ElseIf oElem.Type = msdElementTypeCellHeader Then
            Dim oSub As CellElement
            Set oSub = oElem
            ParcourirEnfants oSub.GetSubElements, oInfo
        End If
    Loop
End Sub

'------------------------------------------------------------------------------
' Reconnait "T:...", "Fe:...", "Prof:..." (insensible a la casse, espaces libres).
' Met a jour les champs correspondants dans oInfo.
Private Sub AnalyserChaine(ByVal sTxt As String, oInfo As CBoiteEPInfo)
    Dim s As String
    s = Trim$(sTxt)
    If Len(s) = 0 Then Exit Sub

    Dim sLow As String
    sLow = LCase$(s)

    Dim dVal As Double

    If EstPrefixe(sLow, "prof") Then
        If ExtraireNombreApres(s, "prof", dVal) Then
            oInfo.Profondeur = dVal
            oInfo.TrouveProfondeur = True
        End If
        Exit Sub
    End If

    If EstPrefixe(sLow, "fe") Then
        If ExtraireNombreApres(s, "fe", dVal) Then
            oInfo.ZFilEau = dVal
            oInfo.TrouveFilEau = True
        End If
        Exit Sub
    End If

    If EstPrefixe(sLow, "t") Then
        If ExtraireNombreApres(s, "t", dVal) Then
            oInfo.ZTampon = dVal
            oInfo.TrouveTampon = True
        End If
        Exit Sub
    End If
End Sub

'------------------------------------------------------------------------------
' Vrai si sLow commence par sPref (deja en minuscule) suivi de : ou espace ou =
Private Function EstPrefixe(ByVal sLow As String, ByVal sPref As String) As Boolean
    EstPrefixe = False
    If Len(sLow) < Len(sPref) + 1 Then Exit Function
    If Left$(sLow, Len(sPref)) <> sPref Then Exit Function
    Dim c As String
    c = Mid$(sLow, Len(sPref) + 1, 1)
    EstPrefixe = (c = ":" Or c = " " Or c = "=")
End Function

'------------------------------------------------------------------------------
' Extrait le nombre situe apres le prefixe (et son separateur : / = / espaces).
' Accepte "." ou "," comme separateur decimal.
Private Function ExtraireNombreApres(ByVal sTxt As String, _
                                     ByVal sPref As String, _
                                     dOut As Double) As Boolean
    ExtraireNombreApres = False
    dOut = 0#

    Dim sReste As String
    sReste = Mid$(sTxt, Len(sPref) + 1)

    ' Retirer les separateurs de tete : ':', '=', espaces
    Do While Len(sReste) > 0
        Dim c As String
        c = Left$(sReste, 1)
        If c = ":" Or c = "=" Or c = " " Or c = vbTab Then
            sReste = Mid$(sReste, 2)
        Else
            Exit Do
        End If
    Loop

    ' Extraire le token nombre (chiffres + . / , + signe eventuel en tete)
    Dim sNb As String
    sNb = ""
    Dim i As Integer
    Dim bDot As Boolean: bDot = False
    For i = 1 To Len(sReste)
        Dim ch As String
        ch = Mid$(sReste, i, 1)
        If i = 1 And (ch = "+" Or ch = "-") Then
            sNb = sNb & ch
        ElseIf ch >= "0" And ch <= "9" Then
            sNb = sNb & ch
        ElseIf (ch = "." Or ch = ",") And Not bDot Then
            sNb = sNb & "."
            bDot = True
        Else
            Exit For
        End If
    Next i

    If Len(sNb) = 0 Then Exit Function
    If sNb = "+" Or sNb = "-" Or sNb = "." Then Exit Function

    dOut = Val(sNb)
    ExtraireNombreApres = True
End Function
