VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlacerPonctuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CPlacerPonctuel - Machine a etats : chemin principal + rayonnement
'
' Etape 1 (CHEMIN) : ligne dynamique depuis le dernier point
'   Data  = creer le point (Z = dernier Z + pente*dist + DZ chemin)
'   Reset = nouvelle selection P1
' Etape 2 (TEXTE CHEMIN) : placement libre du texte altitude
'   Data  = poser le texte, passer en mode rayonnement
' Etape 3 (RAYONNEMENT) : ligne dynamique depuis le point chemin
'   Data  = creer un point rayon (Z = centre Z + pente*dist + DZ rayon)
'   Reset = terminer rayonnement, repartir en chemin depuis ce point
' Etape 4 (TEXTE RAYON) : placement libre du texte altitude rayon
'   Data  = poser le texte, revenir en rayonnement
'
' Pour les cellules/tags, les etapes texte (2 et 4) sont sautees.
'==============================================================================
Option Explicit

Implements IPrimitiveCommandEvents

Private m_nEtape        As Integer      ' 1=chemin, 2=texte chemin, 3=rayon, 4=texte rayon
Private m_dLastX        As Double       ' position du dernier point chemin (X)
Private m_dLastY        As Double       ' position du dernier point chemin (Y)
Private m_dLastZ        As Double       ' altitude du dernier point chemin
Private m_dCentreX      As Double       ' centre rayonnement (X) = dernier point chemin
Private m_dCentreY      As Double       ' centre rayonnement (Y)
Private m_dCentreZ      As Double       ' centre rayonnement (Z)
Private m_dZFinal       As Double       ' altitude figee pour le placement texte

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Interpol. Ponctuelle"

    ' Initialiser la position de depart depuis P1
    Dim oPt1 As Point3d
    g_oP1.CopierPosition oPt1
    m_dLastX = oPt1.X
    m_dLastY = oPt1.Y
    m_dLastZ = g_oP1.Altitude

    m_nEtape = 1
    ShowPrompt "CHEMIN depuis P1 (Z=" & Format$(m_dLastZ, "0.000") & _
               ")  Data=placer un point  Reset=nouv. selection"
    CommandState.StartDynamics
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    Select Case m_nEtape
    Case 1
        DessinChemin DrawMode, Point

    Case 2
        g_oMoteur.DessinTexteLibre DrawMode, Point, m_dZFinal, g_oSettings, g_oCalc
        Dim sA2 As String
        sA2 = g_oCalc.FormatAltitude(m_dZFinal, 3, g_oSettings.oTexte.SepDecimal)
        ShowPrompt "Placez le texte Z=" & sA2 & _
                   "  (Data=poser  Reset=nouv. selection)"

    Case 3
        DessinRayon DrawMode, Point

    Case 4
        g_oMoteur.DessinTexteLibre DrawMode, Point, m_dZFinal, g_oSettings, g_oCalc
        Dim sA4 As String
        sA4 = g_oCalc.FormatAltitude(m_dZFinal, 3, g_oSettings.oTexte.SepDecimal)
        ShowPrompt "Placez le texte rayon Z=" & sA4 & _
                   "  (Data=poser  Reset=nouv. selection)"
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Select Case m_nEtape
    Case 1
        ' --- Creer un point chemin ---
        Dim dDist As Double, dZNew As Double
        dDist = DistXYLocal(m_dLastX, m_dLastY, Point.X, Point.Y)
        dZNew = m_dLastZ
        If g_oSettings.bCheminPente Then
            dZNew = dZNew + dDist * g_oSettings.dCheminPente / 100#
        End If
        If g_oSettings.bCheminDZ Then dZNew = dZNew + g_oSettings.dCheminDZ

        Dim oPtChem As Point3d
        oPtChem.X = Point.X: oPtChem.Y = Point.Y: oPtChem.Z = dZNew

        If SourceP1Integree() Then
            g_oMoteur.CreerPointTopo oPtChem, dZNew, g_oSettings, g_oCalc
            AvancerChemin oPtChem.X, oPtChem.Y, dZNew
            EntrerRayonnement dZNew
        Else
            g_oMoteur.CreerCercleSeul oPtChem, g_oSettings
            m_dZFinal = dZNew
            m_dCentreX = oPtChem.X: m_dCentreY = oPtChem.Y: m_dCentreZ = dZNew
            Dim sAC As String
            sAC = g_oCalc.FormatAltitude(dZNew, g_oSettings.nPonctDecimales, _
                                          g_oSettings.oTexte.SepDecimal)
            ShowPrompt "Cercle cree Z=" & sAC & _
                       "  |  Placez le texte altitude (Data=poser  Reset=nouv. selection)"
            m_nEtape = 2
        End If

    Case 2
        ' --- Poser le texte chemin ---
        Point.Z = m_dZFinal
        g_oMoteur.CreerTexteSeul Point, m_dZFinal, g_oSettings, g_oCalc
        AvancerChemin m_dCentreX, m_dCentreY, m_dZFinal
        EntrerRayonnement m_dZFinal

    Case 3
        ' --- Creer un point rayonnement ---
        Dim dDistR As Double, dZRay As Double
        dDistR = DistXYLocal(m_dCentreX, m_dCentreY, Point.X, Point.Y)
        dZRay = m_dCentreZ
        If g_oSettings.bRayonPente Then
            dZRay = dZRay + dDistR * g_oSettings.dRayonPente / 100#
        End If
        If g_oSettings.bRayonDZ Then dZRay = dZRay + g_oSettings.dRayonDZ

        Dim oPtRay As Point3d
        oPtRay.X = Point.X: oPtRay.Y = Point.Y: oPtRay.Z = dZRay

        If SourceP1Integree() Then
            g_oMoteur.CreerPointTopo oPtRay, dZRay, g_oSettings, g_oCalc
            Dim sAR1 As String
            sAR1 = g_oCalc.FormatAltitude(dZRay, g_oSettings.nPonctDecimales, _
                                            g_oSettings.oTexte.SepDecimal)
            ShowPrompt "Rayon cree Z=" & sAR1 & _
                       "  |  Data=autre rayon  Reset=continuer chemin"
        Else
            g_oMoteur.CreerCercleSeul oPtRay, g_oSettings
            m_dZFinal = dZRay
            Dim sAR2 As String
            sAR2 = g_oCalc.FormatAltitude(dZRay, g_oSettings.nPonctDecimales, _
                                            g_oSettings.oTexte.SepDecimal)
            ShowPrompt "Cercle rayon Z=" & sAR2 & _
                       "  |  Placez le texte (Data=poser  Reset=nouv. selection)"
            m_nEtape = 4
        End If

    Case 4
        ' --- Poser le texte rayonnement ---
        Point.Z = m_dZFinal
        g_oMoteur.CreerTexteSeul Point, m_dZFinal, g_oSettings, g_oCalc
        Dim sAR3 As String
        sAR3 = g_oCalc.FormatAltitude(m_dZFinal, g_oSettings.nPonctDecimales, _
                                        g_oSettings.oTexte.SepDecimal)
        ShowPrompt "Rayon cree Z=" & sAR3 & _
                   "  |  Data=autre rayon  Reset=continuer chemin"
        m_nEtape = 3
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    Select Case m_nEtape
    Case 1, 2
        ' Reset en chemin = nouvelle selection P1
        frmInterpolPonct.ReinitialiserEtat
        ShowPrompt "Nouvelle selection : cliquez pres du texte altitude du POINT 1  (Reset=quitter)"
        CommandState.StartPrimitive New CSelectP1Ponct

    Case 3, 4
        ' Reset en rayonnement = repartir en chemin depuis le centre
        m_dLastX = m_dCentreX: m_dLastY = m_dCentreY: m_dLastZ = m_dCentreZ
        m_nEtape = 1
        ShowPrompt "CHEMIN depuis Z=" & Format$(m_dLastZ, "0.000") & _
                   "  Data=placer un point  Reset=nouv. selection"
    End Select
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

'==============================================================================
' Dessin dynamique
'==============================================================================

Private Sub DessinChemin(ByVal DrawMode As MsdDrawingMode, oPtCurseur As Point3d)
    Dim dDist As Double, dZNew As Double
    dDist = DistXYLocal(m_dLastX, m_dLastY, oPtCurseur.X, oPtCurseur.Y)
    dZNew = m_dLastZ
    If g_oSettings.bCheminPente Then
        dZNew = dZNew + dDist * g_oSettings.dCheminPente / 100#
    End If
    If g_oSettings.bCheminDZ Then dZNew = dZNew + g_oSettings.dCheminDZ

    ' Ligne depuis dernier point vers curseur
    Dim oPtDep As Point3d, oPtArr As Point3d
    oPtDep.X = m_dLastX: oPtDep.Y = m_dLastY: oPtDep.Z = m_dLastZ
    oPtArr.X = oPtCurseur.X: oPtArr.Y = oPtCurseur.Y: oPtArr.Z = dZNew

    Dim oLigne As LineElement
    Set oLigne = CreateLineElement2(Nothing, oPtDep, oPtArr)
    oLigne.Color = 2
    oLigne.Redraw DrawMode

    ' Croix sur le point de depart
    Dim dTaille As Double
    dTaille = g_oSettings.oCercle.Rayon * 0.8
    DessinCroixLocal DrawMode, oPtDep, dTaille

    ' Cercle + texte provisoire au curseur
    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oPtArr, _
                      g_oSettings.oCercle.Rayon, g_oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = g_oSettings.oCercle.Couleur
    oCercle.Redraw DrawMode

    Dim sAlt As String
    sAlt = g_oCalc.FormatAltitude(dZNew, 3, g_oSettings.oTexte.SepDecimal)
    DessinTexteProvLocal DrawMode, oPtArr, sAlt

    ShowPrompt "CHEMIN  Z=" & sAlt & "  dist=" & Format$(dDist, "0.000") & _
               "  (Data=placer  Reset=nouv. selection)"
End Sub

Private Sub DessinRayon(ByVal DrawMode As MsdDrawingMode, oPtCurseur As Point3d)
    Dim dDist As Double, dZRay As Double
    dDist = DistXYLocal(m_dCentreX, m_dCentreY, oPtCurseur.X, oPtCurseur.Y)
    dZRay = m_dCentreZ
    If g_oSettings.bRayonPente Then
        dZRay = dZRay + dDist * g_oSettings.dRayonPente / 100#
    End If
    If g_oSettings.bRayonDZ Then dZRay = dZRay + g_oSettings.dRayonDZ

    ' Ligne depuis le centre vers le curseur
    Dim oPtC As Point3d, oPtR As Point3d
    oPtC.X = m_dCentreX: oPtC.Y = m_dCentreY: oPtC.Z = m_dCentreZ
    oPtR.X = oPtCurseur.X: oPtR.Y = oPtCurseur.Y: oPtR.Z = dZRay

    Dim oLigne As LineElement
    Set oLigne = CreateLineElement2(Nothing, oPtC, oPtR)
    oLigne.Color = 4
    oLigne.Redraw DrawMode

    ' Croix sur le centre
    Dim dTaille As Double
    dTaille = g_oSettings.oCercle.Rayon * 0.8
    DessinCroixLocal DrawMode, oPtC, dTaille

    ' Cercle + texte provisoire au curseur
    Dim oCercle As EllipseElement
    Set oCercle = CreateEllipseElement2(Nothing, oPtR, _
                      g_oSettings.oCercle.Rayon, g_oSettings.oCercle.Rayon, Matrix3dIdentity())
    oCercle.Color = g_oSettings.oCercle.Couleur
    oCercle.Redraw DrawMode

    Dim sAlt As String
    sAlt = g_oCalc.FormatAltitude(dZRay, 3, g_oSettings.oTexte.SepDecimal)
    DessinTexteProvLocal DrawMode, oPtR, sAlt

    ShowPrompt "RAYON  Z=" & sAlt & "  dist=" & Format$(dDist, "0.000") & _
               "  (Data=placer  Reset=continuer chemin)"
End Sub

'==============================================================================
' Utilitaires prives
'==============================================================================

Private Sub AvancerChemin(dX As Double, dY As Double, dZ As Double)
    m_dLastX = dX: m_dLastY = dY: m_dLastZ = dZ
End Sub

Private Sub EntrerRayonnement(dZ As Double)
    m_dCentreX = m_dLastX: m_dCentreY = m_dLastY: m_dCentreZ = dZ
    m_nEtape = 3
    Dim sZ As String
    sZ = g_oCalc.FormatAltitude(dZ, g_oSettings.nPonctDecimales, _
                                 g_oSettings.oTexte.SepDecimal)
    ShowPrompt "RAYON depuis Z=" & sZ & _
               "  Data=placer un rayon  Reset=continuer chemin"
End Sub

Private Function DistXYLocal(dX1 As Double, dY1 As Double, _
                              dX2 As Double, dY2 As Double) As Double
    DistXYLocal = Sqr((dX2 - dX1) ^ 2 + (dY2 - dY1) ^ 2)
End Function

Private Sub DessinCroixLocal(ByVal DrawMode As MsdDrawingMode, _
                              oPt As Point3d, dTaille As Double)
    Dim oA As Point3d, oB As Point3d, oL As LineElement
    oA.X = oPt.X - dTaille: oA.Y = oPt.Y: oA.Z = oPt.Z
    oB.X = oPt.X + dTaille: oB.Y = oPt.Y: oB.Z = oPt.Z
    Set oL = CreateLineElement2(Nothing, oA, oB)
    oL.Color = 3: oL.Redraw DrawMode
    oA.X = oPt.X: oA.Y = oPt.Y - dTaille
    oB.X = oPt.X: oB.Y = oPt.Y + dTaille
    Set oL = CreateLineElement2(Nothing, oA, oB)
    oL.Color = 3: oL.Redraw DrawMode
End Sub

Private Sub DessinTexteProvLocal(ByVal DrawMode As MsdDrawingMode, _
                                  oPt As Point3d, sAlt As String)
    Dim oRot As Matrix3d
    oRot = Matrix3dIdentity()
    On Error Resume Next
    If Not g_oSettings.oTextModele Is Nothing Then oRot = g_oSettings.oTextModele.Rotation
    On Error GoTo 0
    Dim oTxt As TextElement
    On Error Resume Next
    Set oTxt = CreateTextElement1(g_oSettings.oTextModele, sAlt, oPt, oRot)
    If oTxt Is Nothing Then
        Err.Clear
        Set oTxt = CreateTextElement1(Nothing, sAlt, oPt, oRot)
    End If
    On Error GoTo 0
    If Not g_oSettings.oTexte.CommeModele Then oTxt.Color = g_oSettings.oTexte.Couleur
    oTxt.Redraw DrawMode
End Sub

Private Function SourceP1Integree() As Boolean
    SourceP1Integree = False
    If g_oSelectionP1 Is Nothing Then Exit Function
    If g_oSelectionP1.EstReference Then Exit Function
    SourceP1Integree = (Not g_oSelectionP1.oCellule Is Nothing) Or _
                       (Not g_oSelectionP1.oTag Is Nothing)
End Function
