Attribute VB_Name = "InterpolationTopoV2"
'==============================================================================
' InterpolationTopoV2 - Module principal - MicroStation V8i SS3 (VBA)
'
' Interpole l'altitude d'un point situe sur la droite reliant deux textes
' d'altitude existants, avec apercu dynamique, puis cree :
'   - un texte altitude (memes proprietes que le texte P1, pris comme modele)
'   - un cercle centre sur le point projete
'
' Architecture V2 : classes a responsabilites separees + formulaire modeless
'
' INSTALLATION :
'   Dans l'editeur VBA (Alt+F11) : Fichier > Importer les fichiers suivants :
'     InterpolationTopoV2.bas
'     CSymboTexte.cls, CSymboCercle.cls, CMstSettings.cls
'     CPointRef.cls, CInterpolation.cls, CMoteurGraphique.cls
'     CSelectP1.cls, CSelectP2.cls, CPlacerPoint.cls
'     frmInterpolation.frm
'
' LANCEMENT :
'   key-in : vba run [InterpolationTopoV2]InterpolerPoint
'==============================================================================
Option Explicit

' --- Instances partagees entre les classes de commande ---
' Chaque classe recoit ces references mais ne les cree pas.
Public g_oSettings  As CMstSettings      ' parametres (symbologie, tolerance)
Public g_oP1        As CPointRef         ' point de reference 1
Public g_oP2        As CPointRef         ' point de reference 2
Public g_oCalc      As CInterpolation    ' moteur de calcul pur
Public g_oMoteur    As CMoteurGraphique  ' moteur de creation graphique

'------------------------------------------------------------------------------
Sub InterpolerPoint()
    ' Verifier qu'un fichier DGN est ouvert
    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, "Interpolation Topo"
        Exit Sub
    End If

    ' Instanciation unique de toutes les classes
    Set g_oSettings = New CMstSettings
    g_oSettings.Init

    Set g_oP1 = New CPointRef
    Set g_oP2 = New CPointRef
    Set g_oCalc = New CInterpolation
    Set g_oMoteur = New CMoteurGraphique

    ' Afficher le formulaire modeless (Tool Settings)
    ' Le formulaire lit g_oSettings pour pre-remplir les champs
    frmInterpolation.Initialiser g_oSettings
    frmInterpolation.Show vbModeless

    ' Demarrer la commande
    CommandState.StartPrimitive New CSelectP1
End Sub

'==============================================================================
' Utilitaires partages (appeles par les classes de commande)
'==============================================================================

'------------------------------------------------------------------------------
' Cherche le texte numerique le plus proche du clic dans le rayon donne.
' Facteur commun aux trois classes de commande.
Function TrouverTexteProche(oPt As Point3d, dRayon As Double) As TextElement
    Set TrouverTexteProche = Nothing

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
            If g_oCalc.EstNombre(Replace(Trim$(oTxt.Text), ",", ".")) Then
                Dim dD As Double
                dD = g_oCalc.Dist2D(oPt, oTxt.Origin)
                If dD < dMinDist Then
                    dMinDist = dD
                    Set oNearest = oTxt
                End If
            End If
        End If
    Loop

    If Not oNearest Is Nothing Then Set TrouverTexteProche = oNearest
End Function
