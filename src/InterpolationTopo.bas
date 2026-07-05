Attribute VB_Name = "InterpolationTopo"
'==============================================================================
' InterpolationTopo - MicroStation V8i SS3 (VBA)
'
' Interpole l'altitude d'un point situe sur la droite reliant deux textes
' d'altitude existants (points topo), avec apercu dynamique, puis cree :
'   - un texte altitude (memes proprietes que le texte P1, pris comme modele)
'   - un cercle centre sur le point projete
'
' INSTALLATION :
'   1. MicroStation : Utilitaires > Macros > Gestionnaire de projets VBA
'      > Nouveau projet (MicroStation cree alors le vrai fichier .mvba)
'   2. Dans l'editeur VBA (Alt+F11) : Fichier > Importer un fichier (Ctrl+M)
'      et importer les 4 fichiers :
'        InterpolationTopo.bas, CSelectP1.cls, CSelectP2.cls, CPlacerPoint.cls
'   3. Enregistrer, puis lancer la Sub InterpolerPoint
'      (key-in : vba run [InterpolationTopo]InterpolerPoint)
'
' UTILISATION :
'   1) Dialogues : diametre du cercle, couleur, niveau
'   2) Cliquer pres du texte altitude P1 (il sert de modele de mise en forme)
'   3) Cliquer pres du texte altitude P2
'   4) Deplacer le curseur : la position est projetee sur la droite P1-P2,
'      l'altitude interpolee s'affiche en dynamique
'   5) Data = creer le point (repetable sur la meme droite), Reset = terminer
'==============================================================================
Option Explicit

' --- Contexte partage entre les trois classes de commande ---
Public g_oP1           As Point3d      ' origine du texte P1
Public g_oP2           As Point3d      ' origine du texte P2
Public g_dZ1           As Double       ' altitude lue sur P1
Public g_dZ2           As Double       ' altitude lue sur P2
Public g_sDecSep       As String       ' separateur decimal des textes source ("." ou ",")
Public g_nDecimals     As Integer      ' nb de decimales (max des deux textes source)
Public g_oTextTemplate As TextElement  ' texte P1 : modele de symbologie du texte cree
Public g_dCircRayon    As Double
Public g_nCircColor    As Long
Public g_sCircLevel    As String       ' "" = niveau du texte P1, sinon nom ou numero de niveau

' Rayon de recherche du texte autour du clic (en unites maitre)
Public Const TOL_TEXTE As Double = 1#

'------------------------------------------------------------------------------
Sub InterpolerPoint()
    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, "Interpolation Topo"
        Exit Sub
    End If

    If Not ShowParamDialog() Then
        ShowPrompt "Interpolation Topo : annule."
        Exit Sub
    End If
    CommandState.StartPrimitive New CSelectP1
End Sub

'------------------------------------------------------------------------------
Function ShowParamDialog() As Boolean
    ShowParamDialog = False
    Dim sInput As String

    sInput = InputBox("Diametre du cercle (en unites maitre, ex: 0.50) :", _
                      "Interpolation Topo - Parametres", "0.50")
    If sInput = "" Then Exit Function
    sInput = Replace(Trim$(sInput), ",", ".")
    If Not EstNombre(sInput) Or Val(sInput) <= 0 Then
        MsgBox "Diametre invalide.", vbExclamation: Exit Function
    End If
    g_dCircRayon = Val(sInput) / 2#

    sInput = InputBox("Couleur du cercle (index MicroStation 0-255) :" & vbCrLf & _
                      "0=blanc  1=bleu  2=vert  3=rouge  4=jaune  6=orange", _
                      "Interpolation Topo - Parametres", "2")
    If sInput = "" Then Exit Function
    sInput = Trim$(sInput)
    If Not EstNombre(sInput) Then MsgBox "Couleur invalide.", vbExclamation: Exit Function
    g_nCircColor = CLng(Val(sInput))
    If g_nCircColor < 0 Or g_nCircColor > 255 Then
        MsgBox "Couleur entre 0 et 255.", vbExclamation: Exit Function
    End If

    ' Niveau du cercle : liste des niveaux existants, saisie du numero ou du nom
    ' vide = meme niveau que le texte topo P1
    Do
        g_sCircLevel = Trim$(InputBox( _
            "Niveau (Level) du cercle - tapez le NUMERO ou le NOM :" & vbCrLf & _
            "(laisser vide = meme niveau que le texte topo)" & vbCrLf & vbCrLf & _
            ListeNiveaux(), _
            "Interpolation Topo - Niveau du cercle", ""))
        If g_sCircLevel = "" Then Exit Do                 ' vide = niveau du texte P1
        If Not TrouverNiveau(g_sCircLevel) Is Nothing Then Exit Do
        MsgBox "Niveau '" & g_sCircLevel & "' introuvable dans le fichier." & vbCrLf & _
               "Choisissez un numero ou un nom de la liste.", vbExclamation, _
               "Interpolation Topo"
    Loop

    ShowParamDialog = True
End Function

'------------------------------------------------------------------------------
' Liste "numero : nom" des niveaux du fichier actif (tronquee si trop longue
' pour l'InputBox, mais tous les niveaux restent selectionnables par nom/numero)
Function ListeNiveaux() As String
    Const MAX_CAR As Integer = 700
    Dim oL As Level, s As String, nTotal As Integer, nAffiches As Integer
    For Each oL In ActiveDesignFile.Levels
        nTotal = nTotal + 1
        If Len(s) < MAX_CAR Then
            s = s & oL.Number & " : " & oL.Name & vbCrLf
            nAffiches = nAffiches + 1
        End If
    Next
    If nAffiches < nTotal Then
        s = s & "... (" & (nTotal - nAffiches) & " autres niveaux : tapez leur nom ou numero)"
    End If
    ListeNiveaux = s
End Function

'------------------------------------------------------------------------------
' Cherche un niveau par nom (prioritaire) puis par numero ; Nothing si absent
Function TrouverNiveau(ByVal sNiveau As String) As Level
    Set TrouverNiveau = Nothing

    Dim oLvl As Level
    On Error Resume Next
    Set oLvl = ActiveDesignFile.Levels(sNiveau)
    On Error GoTo 0

    If oLvl Is Nothing And EstNombre(sNiveau) Then
        Dim oL As Level
        For Each oL In ActiveDesignFile.Levels
            If oL.Number = CLng(Val(sNiveau)) Then
                Set oLvl = oL
                Exit For
            End If
        Next
    End If

    If Not oLvl Is Nothing Then Set TrouverNiveau = oLvl
End Function

'------------------------------------------------------------------------------
' Lit l'altitude d'un texte : origine, valeur, separateur decimal, nb decimales
Function LireTexteAltitude(oText As TextElement, oPt As Point3d, dZ As Double, _
                           sSep As String, nDec As Integer) As Boolean
    LireTexteAltitude = False
    oPt = oText.Origin

    Dim sBrut As String
    sBrut = Trim$(oText.Text)
    If InStr(sBrut, ",") > 0 Then
        sSep = ","
        sBrut = Replace(sBrut, ",", ".")
    Else
        sSep = "."
    End If

    If Not EstNombre(sBrut) Then Exit Function

    ' Val() est independant des parametres regionaux Windows (attend toujours ".")
    ' contrairement a CDbl() qui, en locale francaise, lirait "152.43" comme 15243
    dZ = Val(sBrut)

    Dim nDot As Integer
    nDot = InStrRev(sBrut, ".")
    If nDot > 0 Then nDec = Len(sBrut) - nDot Else nDec = 0

    LireTexteAltitude = True
End Function

'------------------------------------------------------------------------------
' Vrai si s est un nombre au format normalise : [+-]chiffres[.chiffres]
Private Function EstNombre(ByVal s As String) As Boolean
    EstNombre = False
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "+" Or Left$(s, 1) = "-" Then s = Mid$(s, 2)

    Dim i As Integer, nDots As Integer, nDigits As Integer, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = "." Then
            nDots = nDots + 1
            If nDots > 1 Then Exit Function
        ElseIf c >= "0" And c <= "9" Then
            nDigits = nDigits + 1
        Else
            Exit Function
        End If
    Next
    EstNombre = (nDigits > 0)
End Function

'------------------------------------------------------------------------------
' Cree le texte altitude + le cercle au point projete
Sub CreateTopoPoint(oProj As Point3d, dZInterp As Double)
    Dim sAltitude As String
    sAltitude = FormatAltitude(dZInterp, g_nDecimals, g_sDecSep)

    ' Le texte P1 sert de modele : niveau, couleur, police, taille, style et
    ' justification sont repris sans copie manuelle des proprietes
    Dim oPosTxt As Point3d, oRotTxt As Matrix3d
    PositionTexteAltitude oProj, oPosTxt
    RotationTexteModele oRotTxt
    Dim oTextElem As TextElement
    Set oTextElem = CreateTextElement1(g_oTextTemplate, sAltitude, oPosTxt, oRotTxt)
    ActiveModelReference.AddElement oTextElem

    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oProj, _
                        g_dCircRayon, g_dCircRayon, Matrix3dIdentity())
    oCercle.Color = g_nCircColor
    oCercle.Level = ResolveCircleLevel()
    ActiveModelReference.AddElement oCercle

    ShowPrompt "Point cree Z=" & sAltitude & _
               "  |  Data = autre point sur la meme droite, Reset = nouvelle selection"
End Sub

'------------------------------------------------------------------------------
' Position du texte altitude par rapport au point projete (idem apercu dynamique)
Sub PositionTexteAltitude(oProj As Point3d, oRes As Point3d)
    oRes = oProj
End Sub

'------------------------------------------------------------------------------
' Rotation du texte modele (identite si indisponible, ex. fichier 2D)
Sub RotationTexteModele(oRes As Matrix3d)
    oRes = Matrix3dIdentity()
    On Error Resume Next
    oRes = g_oTextTemplate.Rotation
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Niveau du cercle : saisie vide = niveau du texte P1, sinon nom puis numero
Function ResolveCircleLevel() As Level
    Set ResolveCircleLevel = g_oTextTemplate.Level
    If g_sCircLevel = "" Then Exit Function

    Dim oLvl As Level
    Set oLvl = TrouverNiveau(g_sCircLevel)
    If oLvl Is Nothing Then
        ShowPrompt "Niveau '" & g_sCircLevel & "' introuvable : niveau du texte utilise."
    Else
        Set ResolveCircleLevel = oLvl
    End If
End Function

'------------------------------------------------------------------------------
' Texte numerique (altitude) le plus proche du clic, dans un rayon dRayon.
' Les textes non numeriques (matricules avec lettres, etc.) sont ignores.
Function FindNearestText(oPt As Point3d, dRayon As Double) As TextElement
    Set FindNearestText = Nothing

    Dim oScan As New ElementScanCriteria
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeText

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim dMinDist As Double: dMinDist = dRayon
    Dim oNearest As TextElement

    Do While oEnum.MoveNext
        If oEnum.Current.IsTextElement Then
            Dim oTxt As TextElement
            Set oTxt = oEnum.Current
            If EstNombre(Replace(Trim$(oTxt.Text), ",", ".")) Then
                Dim dD As Double
                dD = Dist2D(oPt, oTxt.Origin)
                If dD < dMinDist Then
                    dMinDist = dD
                    Set oNearest = oTxt
                End If
            End If
        End If
    Loop

    If Not oNearest Is Nothing Then Set FindNearestText = oNearest
End Function

'------------------------------------------------------------------------------
' Pente du segment P1-P2 en % (positive si Z monte de P1 vers P2)
Function PentePct() As Double
    Dim dL As Double
    dL = Dist2D(g_oP1, g_oP2)
    If dL < 0.0000000001 Then PentePct = 0 Else PentePct = (g_dZ2 - g_dZ1) / dL * 100#
End Function

'------------------------------------------------------------------------------
' Gisement de P1 vers P2 en grades (0 = nord/Y+, sens horaire, 0-400 gon)
Function GisementGon() As Double
    Const PI As Double = 3.14159265358979
    Dim dx As Double, dy As Double, dA As Double
    dx = g_oP2.X - g_oP1.X: dy = g_oP2.Y - g_oP1.Y
    If Abs(dx) < 0.0000000001 And Abs(dy) < 0.0000000001 Then Exit Function
    dA = Atn2(dx, dy)                  ' angle depuis Y+ (nord), sens horaire
    GisementGon = dA * 200# / PI
    If GisementGon < 0 Then GisementGon = GisementGon + 400#
End Function

Private Function Atn2(ByVal dY As Double, ByVal dX As Double) As Double
    Const PI As Double = 3.14159265358979
    If dX > 0 Then
        Atn2 = Atn(dY / dX)
    ElseIf dX < 0 Then
        Atn2 = Atn(dY / dX) + IIf(dY >= 0, PI, -PI)
    Else
        Atn2 = IIf(dY >= 0, PI / 2#, -PI / 2#)
    End If
End Function

'------------------------------------------------------------------------------
' Resume du segment : pente et gisement, pour les prompts
Function InfoSegment() As String
    InfoSegment = "Pente=" & Format$(PentePct(), "0.00") & "%  Gisement=" & _
                  Format$(GisementGon(), "0.00") & " gon"
End Function

'------------------------------------------------------------------------------
Function Dist2D(oA As Point3d, oB As Point3d) As Double
    Dist2D = Sqr((oB.X - oA.X) ^ 2 + (oB.Y - oA.Y) ^ 2)
End Function

'------------------------------------------------------------------------------
' Projette C sur la droite AB (en plan). dT = abscisse parametrique (0=A, 1=B).
' Le Z du point projete est interpole entre les Z de A et B (utile en 3D).
' Resultat renvoye via oRes (parametre de sortie).
Sub ProjectSurDroite(oA As Point3d, oB As Point3d, oC As Point3d, _
                     dT As Double, oRes As Point3d)
    Dim dx As Double, dy As Double, dL2 As Double
    dx = oB.X - oA.X: dy = oB.Y - oA.Y
    dL2 = dx * dx + dy * dy
    If dL2 < 0.0000000001 Then
        dT = 0
        oRes = oA
        Exit Sub
    End If
    dT = ((oC.X - oA.X) * dx + (oC.Y - oA.Y) * dy) / dL2
    oRes.X = oA.X + dT * dx
    oRes.Y = oA.Y + dT * dy
    oRes.Z = oA.Z + dT * (oB.Z - oA.Z)
End Sub

'------------------------------------------------------------------------------
Function FormatAltitude(dZ As Double, nDec As Integer, sSep As String) As String
    Dim sFmt As String
    If nDec <= 0 Then sFmt = "0" Else sFmt = "0." & String$(nDec, "0")
    Dim sRes As String
    sRes = Format$(dZ, sFmt)
    ' Format$ utilise le separateur regional Windows : on normalise vers "."
    ' puis on applique le separateur des textes source
    sRes = Replace(sRes, ",", ".")
    If sSep = "," Then sRes = Replace(sRes, ".", ",")
    FormatAltitude = sRes
End Function

'------------------------------------------------------------------------------
Function InterpolerZ(dZ1 As Double, dZ2 As Double, dT As Double) As Double
    InterpolerZ = dZ1 + (dZ2 - dZ1) * dT
End Function
