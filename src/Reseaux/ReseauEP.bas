Attribute VB_Name = "ReseauEP"
'==============================================================================
' ReseauEP - Module reseau EU/EP : pente entre deux regards/grilles
'
' Flux en 4 clics + 1 (5 etapes) :
'   1) CSnapBoiteP1   : snap sur l'axe P1 (tampon/grille)  -> g_ptSnapP1
'   2) CSelectBoiteP1 : clic pres de la cellule qui donne le Fe de P1
'                       -> g_oBoiteP1 (T:/Fe:/Prof:), g_oRefFeP1 = (P1 snap, Fe)
'   3) CSnapBoiteP2   : snap sur l'axe P2 (dynamics : ligne P1_snap -> curseur)
'                       -> g_ptSnapP2
'   4) CSelectBoiteP2 : clic pres de la cellule Fe de P2
'                       -> ligne definitive P1_snap -> P2_snap au niveau Fe
'   5) CPlacerPenteReseau : texte pente + fleche (comme V2 PlacerPente)
'
' Le point du trace et du calcul est le point SNAPPE, pas l'origine de la
' cellule (des reseaux proches peuvent rendre la cellule la plus proche
' ambigue : le clic explicite sur la cellule leve l'ambiguite).
'
' Le module s'appuie sur g_oCalc/g_oMoteur/g_oSettings deja fournis par
' InterpolationTopoV2 (fonction InitialiserContexte) : on ne duplique ni le
' calcul de pente ni le rendu de la fleche. On complete avec g_oBoiteP1/P2
' (donnees reseau) et deux CPointRef alimentes par les Fe pour reutiliser
' g_oMoteur.CreerPente / DessinPenteDynamique.
'
' LANCEMENT :
'   key-in : vba run [InterpolationTopoV2]PenteReseauEP
'==============================================================================
Option Explicit

' --- Instances partagees entre les classes de commande reseau ---
Public g_oBoiteP1        As CBoiteEPInfo
Public g_oBoiteP2        As CBoiteEPInfo
Public g_oParserBoite    As CBoiteEPParser
Public g_oRefFeP1        As CPointRef   ' alimente avec (snap XY, Fe) pour g_oMoteur
Public g_oRefFeP2        As CPointRef
Public g_dTolReseau      As Double      ' rayon de recherche autour du clic (u.m.),
                                        ' modifiable via frmReseauEP

' --- Points snappes (positions exactes cliquees, servent au trace) ---
' Publics dans un .bas : les UDT Point3d sont autorises hors modules de classe.
Public g_ptSnapP1        As Point3d
Public g_ptSnapP2        As Point3d

' --- Debug : voir Point recu au clic + distances aux cellules ref ---
Public g_bDebugClic      As Boolean

Sub DebugClicOn():  g_bDebugClic = True:  MsgBox "Debug clic ACTIF - lance PenteReseauEP":  End Sub
Sub DebugClicOff(): g_bDebugClic = False: MsgBox "Debug clic desactive":                   End Sub

'------------------------------------------------------------------------------
' Point d'entree de la commande Pente Reseau EU/EP
Sub PenteReseauEP()
    If Not EnvironnementPret("Pente Reseau EU/EP") Then Exit Sub

    ' Contexte commun (fournit g_oSettings/g_oCalc/g_oMoteur)
    InitialiserContexte

    ' Contexte specifique reseau
    Set g_oParserBoite = New CBoiteEPParser
    Set g_oBoiteP1 = New CBoiteEPInfo
    Set g_oBoiteP2 = New CBoiteEPInfo
    Set g_oRefFeP1 = New CPointRef
    Set g_oRefFeP2 = New CPointRef
    g_dTolReseau = 5#

    ' Pentes : 1 decimale (defaut CSymboPente.InitDefauts, ligne explicite pour trace)
    g_oSettings.oIndicPente.Decimales = 1

    AfficherFormulaire frmReseauEP

    CommandState.StartPrimitive New CSnapBoiteP1
End Sub

'==============================================================================
' Recherche de la cellule regard/grille la plus proche du clic
'
' Scanne le modele actif ET les fichiers attaches (references). Les origines
' issues d'une reference sont converties dans le repere du modele actif via
' ReferenceOrigin / MasterOrigin / ScaleFactor (meme approche que
' RechercheAltitude.bas de la commande V2).
'==============================================================================

'------------------------------------------------------------------------------
Public Function TrouverBoiteEPProche(oPt As Point3d, dRayon As Double, _
                                     oResult As CBoiteEPInfo) As Boolean
    TrouverBoiteEPProche = False
    If oResult Is Nothing Then Set oResult = New CBoiteEPInfo
    oResult.Vider

    Dim oScan As New ElementScanCriteria
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeCellHeader

    Dim dMin As Double: dMin = dRayon
    Dim oBest As CBoiteEPInfo: Set oBest = Nothing

    ' 1) Modele actif
    ScannerBoiteDansModele ActiveModelReference, Nothing, oScan, oPt, dMin, oBest

    ' 2) Fichiers attaches : iterer par index (le For Each sur Attachments est
    '    fragile selon la version V8i, il peut echouer sans erreur visible).
    Dim oAttachments As Object
    Set oAttachments = ObtenirAttachments()
    Dim nCount As Long: nCount = 0
    On Error Resume Next
    If Not oAttachments Is Nothing Then nCount = oAttachments.Count
    On Error GoTo 0

    Dim i As Long
    For i = 1 To nCount
        Dim oAtt As Object
        Set oAtt = Nothing
        On Error Resume Next
        Set oAtt = oAttachments(i)
        On Error GoTo 0
        If Not (oAtt Is Nothing) Then
            If AttachmentAffiche(oAtt) Then
                ScannerBoiteDansModele oAtt, oAtt, oScan, oPt, dMin, oBest
            End If
        End If
    Next i

    If oBest Is Nothing Then Exit Function
    oResult.CopierDepuis oBest
    TrouverBoiteEPProche = True
End Function

'------------------------------------------------------------------------------
' Renvoie la collection des attachements du modele actif.
Private Function ObtenirAttachments() As Object
    Set ObtenirAttachments = Nothing
    On Error Resume Next
    Set ObtenirAttachments = ActiveModelReference.Attachments
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Scanne un modele ou une reference attachee. Les origines stockees dans oBest
' sont toujours exprimees dans le repere du modele actif.
Private Sub ScannerBoiteDansModele(oModel As Object, oAttachment As Object, _
        oScan As ElementScanCriteria, oPtClic As Point3d, _
        dMinDist As Double, oBest As CBoiteEPInfo)

    Dim oEnum As ElementEnumerator
    On Error Resume Next
    Set oEnum = oModel.Scan(oScan)
    If Err.Number <> 0 Or oEnum Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Do While oEnum.MoveNext
        Dim oElem As Element
        Set oElem = oEnum.Current
        If oElem.Type <> msdElementTypeCellHeader Then GoTo Suivant

        Dim oInfo As CBoiteEPInfo
        Set oInfo = g_oParserBoite.Lire(oElem)
        If Not oInfo.Valide Then GoTo Suivant

        ' Origine cellule -> repere maitre
        Dim ptSrc As Point3d, ptMaster As Point3d
        oInfo.CopierOrigine ptSrc
        TransformerPointVersMaitre oAttachment, ptSrc, ptMaster

        Dim dD As Double
        dD = Dist2DLoc(oPtClic, ptMaster)
        If dD < dMinDist Then
            dMinDist = dD
            oInfo.DefinirOrigine ptMaster
            oInfo.EstReference = Not (oAttachment Is Nothing)
            Set oInfo.oAttachment = oAttachment
            Set oBest = oInfo
        End If
Suivant:
    Loop
End Sub

'------------------------------------------------------------------------------
Private Function AttachmentAffiche(oAttachment As Object) As Boolean
    AttachmentAffiche = True
    On Error Resume Next
    AttachmentAffiche = CBool(oAttachment.IsDisplayed)
    If Err.Number <> 0 Then
        Err.Clear
        AttachmentAffiche = True
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Convertit une origine issue d'une reference vers le repere du modele actif.
'
' Verification sur DGN reel : MicroStation renvoie oCell.Origin dans un espace
' de coord DEJA compatible avec le master (identite). Le ScaleFactor
' d'attachement est une echelle graphique du CONTENU (taille visuelle), pas
' des positions -> ne pas l'appliquer aux coords.
'
' On conserve simplement la translation MasterOrigin - RefOrigin pour couvrir
' le cas ou le ref est place a un offset du master.
Private Sub TransformerPointVersMaitre(oAttachment As Object, _
        oPtSource As Point3d, oPtMaster As Point3d)

    oPtMaster = oPtSource
    If oAttachment Is Nothing Then Exit Sub

    Dim ptRefOrigin As Point3d, ptMasterOrigin As Point3d
    On Error Resume Next
    ptRefOrigin = oAttachment.ReferenceOrigin
    If Err.Number <> 0 Then Err.Clear
    ptMasterOrigin = oAttachment.MasterOrigin
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    oPtMaster.X = ptMasterOrigin.X + (oPtSource.X - ptRefOrigin.X)
    oPtMaster.Y = ptMasterOrigin.Y + (oPtSource.Y - ptRefOrigin.Y)
    oPtMaster.Z = ptMasterOrigin.Z + (oPtSource.Z - ptRefOrigin.Z)
End Sub

'------------------------------------------------------------------------------
' Calcule le rapport UORsPerMasterUnit(actif) / UORsPerMasterUnit(ref).
' Renvoie 1 si l'un des deux n'est pas accessible.
Private Function RatioUORActifSurRef(oAttachment As Object) As Double
    RatioUORActifSurRef = 1#

    Dim dAct As Double: dAct = 0#
    Dim dRef As Double: dRef = 0#

    Dim oActive As Object
    Set oActive = ActiveModelReference

    On Error Resume Next

    dAct = CDbl(oActive.UORsPerMasterUnit)
    If Err.Number <> 0 Or dAct <= 0 Then Err.Clear: dAct = 0

    ' L'Attachment peut exposer UORsPerMasterUnit directement, ou via son
    ' ModelReference interne selon la version.
    dRef = CDbl(oAttachment.UORsPerMasterUnit)
    If Err.Number <> 0 Or dRef <= 0 Then
        Err.Clear
        dRef = CDbl(oAttachment.ModelReference.UORsPerMasterUnit)
        If Err.Number <> 0 Or dRef <= 0 Then Err.Clear: dRef = 0
    End If

    On Error GoTo 0

    If dAct > 0 And dRef > 0 Then RatioUORActifSurRef = dAct / dRef
End Function

'------------------------------------------------------------------------------
' Rayon de recherche courant (partage entre les etats, modifiable via UI).
Public Function TolReseau() As Double
    If g_dTolReseau <= 0 Then g_dTolReseau = 5#
    TolReseau = g_dTolReseau
End Function

'------------------------------------------------------------------------------
Private Function Dist2DLoc(oA As Point3d, oB As Point3d) As Double
    Dist2DLoc = Sqr((oB.X - oA.X) ^ 2 + (oB.Y - oA.Y) ^ 2)
End Function

'==============================================================================
' Utilitaires partages par les classes d'etat reseau
'==============================================================================

'------------------------------------------------------------------------------
' Charge un CPointRef avec les coordonnees XY du point snappe et l'altitude Fe
' de la cellule associee. Le point de reference sert au calcul de pente et au
' trace : c'est bien le snap qui compte, pas l'origine de la cellule.
Public Sub AlimenterPointRefFeSnap(oRef As CPointRef, _
                                    ptSnap As Point3d, oInfo As CBoiteEPInfo)
    oRef.DefinirPosition ptSnap
    oRef.Altitude = oInfo.ZFilEau
    oRef.Valide = oInfo.TrouveFilEau
End Sub

'------------------------------------------------------------------------------
' Texte court d'affichage d'une info reseau (utilise dans les prompts).
Public Function DecrireBoite(oInfo As CBoiteEPInfo) As String
    Dim s As String
    s = "Fe=" & Format$(oInfo.ZFilEau, "0.000")
    If oInfo.TrouveTampon Then s = s & "  T=" & Format$(oInfo.ZTampon, "0.000")
    If oInfo.TrouveProfondeur Then s = s & "  Prof=" & Format$(oInfo.Profondeur, "0.00")
    If oInfo.EstReference Then s = s & "  [ref]"
    DecrireBoite = s
End Function

'------------------------------------------------------------------------------
' Debug : affiche Point recu + distance aux 5 cellules ref les plus proches.
' Appele depuis CSelectBoiteP1.DataPoint quand g_bDebugClic est actif.
Public Sub DebugAfficherClicEtDistances(oPtClic As Point3d)
    Dim s As String
    s = "=== Click recu ===" & vbCrLf & _
        "Point.X = " & Format$(oPtClic.X, "0.000000") & vbCrLf & _
        "Point.Y = " & Format$(oPtClic.Y, "0.000000") & vbCrLf & vbCrLf & _
        "TolReseau = " & TolReseau() & vbCrLf & vbCrLf & _
        "=== Distances aux cellules ==="

    ' Scanner comme TrouverBoiteEPProche mais tout logger
    Dim oScan As New ElementScanCriteria
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeCellHeader

    Dim oAttachments As Object: Set oAttachments = ObtenirAttachments()
    Dim nCount As Long: nCount = 0
    On Error Resume Next
    nCount = oAttachments.Count
    On Error GoTo 0

    s = s & vbCrLf & AfficherDistancesModele(ActiveModelReference, Nothing, oPtClic, oScan, "actif")

    Dim i As Long
    For i = 1 To nCount
        Dim oAtt As Object: Set oAtt = Nothing
        On Error Resume Next
        Set oAtt = oAttachments(i)
        On Error GoTo 0
        If Not oAtt Is Nothing Then
            s = s & AfficherDistancesModele(oAtt, oAtt, oPtClic, oScan, "ref#" & i)
        End If
    Next i

    MsgBox s, vbInformation, "Debug clic BoiteEP"
End Sub

Private Function AfficherDistancesModele(oModel As Object, oAtt As Object, _
        oPtClic As Point3d, oScan As ElementScanCriteria, sTag As String) As String

    Dim oEnum As ElementEnumerator
    On Error Resume Next
    Set oEnum = oModel.Scan(oScan)
    On Error GoTo 0
    If oEnum Is Nothing Then AfficherDistancesModele = "": Exit Function

    Dim s As String: s = vbCrLf & "[" & sTag & "]"
    Dim n As Long: n = 0
    Do While oEnum.MoveNext
        Dim oElem As Element: Set oElem = oEnum.Current
        If oElem.Type = msdElementTypeCellHeader Then
            Dim oInfo As CBoiteEPInfo
            Set oInfo = g_oParserBoite.Lire(oElem)
            If oInfo.Valide Then
                Dim pt As Point3d, ptM As Point3d
                oInfo.CopierOrigine pt
                TransformerPointVersMaitre oAtt, pt, ptM
                Dim dD As Double
                dD = Sqr((oPtClic.X - ptM.X) ^ 2 + (oPtClic.Y - ptM.Y) ^ 2)
                If n < 8 Then
                    s = s & vbCrLf & "  Fe=" & Format$(oInfo.ZFilEau, "0.00") & _
                        "  master=(" & Format$(ptM.X, "0.00") & "," & _
                                       Format$(ptM.Y, "0.00") & ")" & _
                        "  dist=" & Format$(dD, "0.00")
                    n = n + 1
                End If
            End If
        End If
    Loop
    AfficherDistancesModele = s
End Function

'==============================================================================
' Diagnostic references (a lancer via : vba run [InterpolationTopoV2]DebugScanRefsBoiteEP)
'==============================================================================

'------------------------------------------------------------------------------
' Ouvre une boite d'info listant, pour chaque attachement :
'   - son nom / son chemin logique
'   - IsDisplayed
'   - nb total de CellHeader scannees
'   - nb de cellules BoiteEP valides (avec Fe:)
' Utile pour verifier que TrouverBoiteEPProche voit bien les cellules du ref.
Sub DebugScanRefsBoiteEP()
    If Not EnvironnementPret("Debug Scan Refs") Then Exit Sub
    If g_oParserBoite Is Nothing Then Set g_oParserBoite = New CBoiteEPParser

    Dim sRap As String
    sRap = "=== Modele actif ===" & vbCrLf
    sRap = sRap & ResumerScan(ActiveModelReference, Nothing) & vbCrLf

    Dim oAttachments As Object
    Set oAttachments = ObtenirAttachments()
    Dim nCount As Long: nCount = 0
    On Error Resume Next
    If Not oAttachments Is Nothing Then nCount = oAttachments.Count
    On Error GoTo 0

    sRap = sRap & vbCrLf & "=== Attachements : " & nCount & " ===" & vbCrLf

    Dim i As Long
    For i = 1 To nCount
        Dim oAtt As Object: Set oAtt = Nothing
        On Error Resume Next
        Set oAtt = oAttachments(i)
        On Error GoTo 0
        If oAtt Is Nothing Then
            sRap = sRap & vbCrLf & "[" & i & "] (nul)" & vbCrLf
        Else
            sRap = sRap & vbCrLf & "[" & i & "] " & DecrireAttachment(oAtt) & vbCrLf
            sRap = sRap & ResumerScan(oAtt, oAtt) & vbCrLf
        End If
    Next i

    MsgBox sRap, vbInformation, "Debug Scan Refs BoiteEP"
End Sub

Private Function DecrireAttachment(oAtt As Object) As String
    Dim s As String
    On Error Resume Next
    s = "IsDisplayed=" & CStr(oAtt.IsDisplayed)
    Err.Clear
    Dim sName As String
    sName = CStr(oAtt.LogicalName)
    If Err.Number <> 0 Then Err.Clear: sName = "(sans nom logique)"
    s = s & "  Name=""" & sName & """"
    Err.Clear
    Dim sFile As String
    sFile = CStr(oAtt.AttachFileName)
    If Err.Number = 0 Then s = s & "  File=""" & sFile & """"
    Err.Clear
    Dim dScale As Double
    dScale = oAtt.ScaleFactor
    If Err.Number = 0 Then s = s & vbCrLf & "    ScaleFactor=" & Format$(dScale, "0.000000")
    Err.Clear
    Dim dRatio As Double
    dRatio = RatioUORActifSurRef(oAtt)
    s = s & "  UOR ratio=" & Format$(dRatio, "0.000000")
    Dim ptMO As Point3d, ptRO As Point3d
    ptMO = oAtt.MasterOrigin
    If Err.Number = 0 Then s = s & vbCrLf & "    MasterOrigin=(" & _
        Format$(ptMO.X, "0.00") & ", " & Format$(ptMO.Y, "0.00") & ")"
    Err.Clear
    ptRO = oAtt.ReferenceOrigin
    If Err.Number = 0 Then s = s & "  RefOrigin=(" & _
        Format$(ptRO.X, "0.00") & ", " & Format$(ptRO.Y, "0.00") & ")"
    Err.Clear
    Dim tf As Transform3d
    tf = oAtt.Transform
    If Err.Number = 0 Then
        s = s & vbCrLf & "    Matrix RowX=(" & _
            Format$(tf.RowX.X, "0.000000") & ", " & _
            Format$(tf.RowX.Y, "0.000000") & ", " & _
            Format$(tf.RowX.Z, "0.000000") & ")"
    End If
    Err.Clear
    On Error GoTo 0
    DecrireAttachment = s
End Function

Private Function ResumerScan(oModel As Object, oAtt As Object) As String
    Dim oScan As New ElementScanCriteria
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeCellHeader

    Dim oEnum As ElementEnumerator
    On Error Resume Next
    Set oEnum = oModel.Scan(oScan)
    If Err.Number <> 0 Or oEnum Is Nothing Then
        ResumerScan = "  -> Scan impossible (err " & Err.Number & ")"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim nCells As Long: nCells = 0
    Dim nValides As Long: nValides = 0
    Dim sExemples As String

    Do While oEnum.MoveNext
        Dim oElem As Element
        Set oElem = oEnum.Current
        If oElem.Type = msdElementTypeCellHeader Then
            nCells = nCells + 1
            Dim oInfo As CBoiteEPInfo
            Set oInfo = g_oParserBoite.Lire(oElem)
            If oInfo.Valide Then
                nValides = nValides + 1
                If Len(sExemples) < 600 Then
                    Dim pt As Point3d, ptM As Point3d
                    oInfo.CopierOrigine pt
                    TransformerPointVersMaitre oAtt, pt, ptM
                    sExemples = sExemples & _
                        "     Fe=" & Format$(oInfo.ZFilEau, "0.00") & _
                        "  ref=(" & Format$(pt.X, "0") & _
                                "," & Format$(pt.Y, "0") & ")" & _
                        "  master=(" & Format$(ptM.X, "0.00") & _
                                    "," & Format$(ptM.Y, "0.00") & ")" & vbCrLf
                End If
            End If
        End If
    Loop

    ResumerScan = "  -> CellHeader = " & nCells & _
                  " ; BoiteEP valides = " & nValides
    If Len(sExemples) > 0 Then ResumerScan = ResumerScan & vbCrLf & sExemples
End Function
