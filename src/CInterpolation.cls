VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInterpolation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CInterpolation - Moteur de calcul pur
'
' Responsabilites :
'   - parsing d'un texte altitude (numerique, separateur, decimales)
'   - projection orthogonale d'un point sur une droite
'   - interpolation d'altitude par abscisse parametrique
'   - calcul geometrique : distance 2D, pente, gisement
'   - formatage d'une altitude
'
' Cette classe ne connait pas : symbologie, niveaux, couleurs, elements MST,
' formulaire, parametres de commande. Elle ne manipule que des donnees.
'==============================================================================
Option Explicit

Private Const PI As Double = 3.14159265358979

'==============================================================================
' Parsing
'==============================================================================

'------------------------------------------------------------------------------
' Vrai si s est un nombre au format normalise [+-]chiffres[.chiffres]
Function EstNombre(ByVal s As String) As Boolean
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
' Lit l'altitude d'un TextElement et charge le CPointRef.
' Renseigne aussi sSep ("." ou ",") et nDec (nb de decimales).
' Retourne False si le texte n'est pas numerique.
Function LireAltitudeTexte(oPoint As CPointRef, _
                            sSep As String, nDec As Integer) As Boolean
    LireAltitudeTexte = False
    oPoint.Invalider

    Dim sBrut As String
    sBrut = g_sValeurTexte
    If InStr(sBrut, ",") > 0 Then
        sSep = ","
        sBrut = Replace(sBrut, ",", ".")
    Else
        sSep = "."
    End If

    If Not EstNombre(sBrut) Then Exit Function

    oPoint.Altitude = Val(sBrut)
    oPoint.DefinirPosition g_ptOrigineTexte
    oPoint.Valide = True

    Dim nDot As Integer
    nDot = InStrRev(sBrut, ".")
    If nDot > 0 Then nDec = Len(sBrut) - nDot Else nDec = 0

    LireAltitudeTexte = True
End Function

'==============================================================================
' Geometrie
'==============================================================================

'------------------------------------------------------------------------------
Function Dist2D(oA As Point3d, oB As Point3d) As Double
    Dist2D = Sqr((oB.X - oA.X) ^ 2 + (oB.Y - oA.Y) ^ 2)
End Function

'------------------------------------------------------------------------------
' Variante pour deux points de reference (evite de manipuler des Point3d)
Function Dist2DRef(oP1 As CPointRef, oP2 As CPointRef) As Double
    Dist2DRef = Sqr((oP2.X - oP1.X) ^ 2 + (oP2.Y - oP1.Y) ^ 2)
End Function

'------------------------------------------------------------------------------
' Projette oC sur la droite oA-oB en plan.
' dT : abscisse parametrique (0=A, 1=B, negatif ou >1 = extrapolation)
' oRes : point projete avec Z interpole entre oA.Z et oB.Z
Sub ProjecterSurDroite(oA As Point3d, oB As Point3d, oC As Point3d, _
                       dT As Double, oRes As Point3d)
    Dim dx As Double, dy As Double, dL2 As Double
    dx = oB.X - oA.X: dy = oB.Y - oA.Y
    dL2 = dx * dx + dy * dy
    If dL2 < 0.0000000001 Then
        dT = 0: oRes = oA: Exit Sub
    End If
    dT = ((oC.X - oA.X) * dx + (oC.Y - oA.Y) * dy) / dL2
    oRes.X = oA.X + dT * dx
    oRes.Y = oA.Y + dT * dy
    oRes.Z = oA.Z + dT * (oB.Z - oA.Z)
End Sub

'------------------------------------------------------------------------------
' Variante : projette oC sur la droite definie par deux points de reference.
' Les classes de commande passent directement g_oP1/g_oP2.
Sub ProjecterSurSegment(oP1 As CPointRef, oP2 As CPointRef, oC As Point3d, _
                        dT As Double, oRes As Point3d)
    Dim oA As Point3d, oB As Point3d
    oP1.CopierPosition oA
    oP2.CopierPosition oB
    ProjecterSurDroite oA, oB, oC, dT, oRes
End Sub

'------------------------------------------------------------------------------
Function InterpolerZ(dZ1 As Double, dZ2 As Double, dT As Double) As Double
    InterpolerZ = dZ1 + (dZ2 - dZ1) * dT
End Function

'==============================================================================
' Indicateurs geometriques du segment P1-P2
'==============================================================================

'------------------------------------------------------------------------------
Function PentePct(oP1 As CPointRef, oP2 As CPointRef) As Double
    Dim dL As Double
    dL = Dist2DRef(oP1, oP2)
    If dL < 0.0000000001 Then PentePct = 0 Else _
        PentePct = (oP2.Altitude - oP1.Altitude) / dL * 100#
End Function

'------------------------------------------------------------------------------
Function GisementGon(oP1 As CPointRef, oP2 As CPointRef) As Double
    Dim dx As Double, dy As Double, dA As Double
    dx = oP2.X - oP1.X
    dy = oP2.Y - oP1.Y
    If Abs(dx) < 0.0000000001 And Abs(dy) < 0.0000000001 Then Exit Function
    dA = Atn2Priv(dx, dy)
    GisementGon = dA * 200# / PI
    If GisementGon < 0 Then GisementGon = GisementGon + 400#
End Function

Private Function Atn2Priv(ByVal dY As Double, ByVal dX As Double) As Double
    If dX > 0 Then
        Atn2Priv = Atn(dY / dX)
    ElseIf dX < 0 Then
        Atn2Priv = Atn(dY / dX) + IIf(dY >= 0, PI, -PI)
    Else
        Atn2Priv = IIf(dY >= 0, PI / 2#, -PI / 2#)
    End If
End Function

'------------------------------------------------------------------------------
Function InfoSegment(oP1 As CPointRef, oP2 As CPointRef) As String
    InfoSegment = "Pente=" & Format$(PentePct(oP1, oP2), "0.00") & "%" & _
                  "  Gisement=" & Format$(GisementGon(oP1, oP2), "0.00") & " gon"
End Function

'==============================================================================
' Angles
'==============================================================================

'------------------------------------------------------------------------------
' Angle en radians depuis l'axe X (sens trigonometrique) pour un vecteur (dDX, dDY).
Function Angle2D(dDX As Double, dDY As Double) As Double
    If Abs(dDX) < 0.0000000001 And Abs(dDY) < 0.0000000001 Then
        Angle2D = 0#
        Exit Function
    End If
    Angle2D = Atn2Priv(dDY, dDX)
End Function

'==============================================================================
' Formatage
'==============================================================================

'------------------------------------------------------------------------------
Function FormatAltitude(dZ As Double, nDec As Integer, sSep As String) As String
    Dim sFmt As String
    If nDec <= 0 Then sFmt = "0" Else sFmt = "0." & String$(nDec, "0")
    Dim sRes As String
    sRes = Format$(dZ, sFmt)
    ' Format$ suit la locale Windows : on normalise puis on applique le separateur source
    sRes = Replace(sRes, ",", ".")
    If sSep = "," Then sRes = Replace(sRes, ".", ",")
    FormatAltitude = sRes
End Function

'------------------------------------------------------------------------------
' Formate une pente en pourcentage (valeur absolue, 2 decimales, separateur donne).
Function FormatPente(dPente As Double, sSep As String) As String
    Dim sRes As String
    sRes = Format$(Abs(dPente), "0.00")
    sRes = Replace(sRes, ",", ".")
    If sSep = "," Then sRes = Replace(sRes, ".", ",")
    FormatPente = sRes & "%"
End Function
