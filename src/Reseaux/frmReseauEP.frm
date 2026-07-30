VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReseauEP
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmReseauEP.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "frmReseauEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmReseauEP - Tool Settings de la commande Pente Reseau EU/EP
'
' Cadres :
'   Etat        : lecture P1 et P2 (Fe / T / Prof), distance 2D, pente
'   Recherche   : rayon de recherche autour du clic (u.m.)
'   Indicateur  : niveau, style de texte, decimales, couleur/longueur du corps
'                 de fleche, longueur de la pointe et choix pointe ouverte
'                 (2 lignes) / fermee (triangle plein) -- memes reglages
'                 partages (g_oSettings.oIndicPente) et meme rendu que
'                 frmInterpolation / frmInterpolPonct : la fleche (corps +
'                 pointe) est creee par CMoteurGraphique.CreerPente comme une
'                 seule cellule orpheline "IndicPente", selectionnable en 1 clic.
'                 (reglages toujours actifs, pas de mode "personnalise" : la
'                 taille/couleur du texte pente vient du style de texte nomme)
'
' Tous les controles sont crees au runtime dans ConstruireControles (le .frx
' n'est qu'un blob de formulaire vide, cf. CLAUDE.md).
'==============================================================================
Option Explicit

Private m_oSettings  As CMstSettings
Private m_bInit      As Boolean
Private m_bConstruit As Boolean

' --- Etat ---
Private lblP1        As MSForms.Label
Private lblP2        As MSForms.Label
Private lblDist      As MSForms.Label
Private lblPente     As MSForms.Label

' --- Recherche ---
Private WithEvents txtTol As MSForms.TextBox
Attribute txtTol.VB_VarHelpID = -1

' --- Indicateur pente ---
Private WithEvents cmbNiveau As MSForms.ComboBox
Attribute cmbNiveau.VB_VarHelpID = -1
Private WithEvents cmbStyle As MSForms.ComboBox
Attribute cmbStyle.VB_VarHelpID = -1
Private WithEvents txtDec As MSForms.TextBox
Attribute txtDec.VB_VarHelpID = -1
Private WithEvents txtLongFleche As MSForms.TextBox
Attribute txtLongFleche.VB_VarHelpID = -1
Private WithEvents txtCoulFleche As MSForms.TextBox
Attribute txtCoulFleche.VB_VarHelpID = -1
Private WithEvents txtLongPointe As MSForms.TextBox
Attribute txtLongPointe.VB_VarHelpID = -1
Private WithEvents chkFermee As MSForms.CheckBox
Attribute chkFermee.VB_VarHelpID = -1

'==============================================================================
' Construction des controles
'==============================================================================

Private Sub UserForm_Initialize()
    ConstruireControles
End Sub

Private Sub ConstruireControles()
    If m_bConstruit Then Exit Sub
    m_bConstruit = True

    Me.Caption = "Pente Reseau EU/EP"
    Me.Width = 212
    Me.Height = 338

    Dim dY As Double
    dY = 6

    ' --- Cadre Etat -----------------------------------------------------------
    Dim fraEtat As MSForms.Frame
    Set fraEtat = Me.Controls.Add("Forms.Frame.1", "fraEtat")
    fraEtat.Caption = "Etat"
    fraEtat.Left = 6: fraEtat.Top = dY
    fraEtat.Width = 192: fraEtat.Height = 68

    Set lblP1 = CreerLabel(fraEtat, "lblP1", "P1 : -", 6, 12, 180)
    Set lblP2 = CreerLabel(fraEtat, "lblP2", "P2 : -", 6, 24, 180)
    Set lblDist = CreerLabel(fraEtat, "lblDist", "Distance : -", 6, 38, 180)
    Set lblPente = CreerLabel(fraEtat, "lblPente", "Pente : -", 6, 50, 180)

    dY = dY + 74

    ' --- Cadre Recherche ------------------------------------------------------
    Dim fraRech As MSForms.Frame
    Set fraRech = Me.Controls.Add("Forms.Frame.1", "fraRech")
    fraRech.Caption = "Recherche"
    fraRech.Left = 6: fraRech.Top = dY
    fraRech.Width = 192: fraRech.Height = 38

    CreerLabel fraRech, "lblTol", "Rayon (u.m.) :", 6, 14, 70
    Set txtTol = fraRech.Controls.Add("Forms.TextBox.1", "txtTol")
    txtTol.Left = 82: txtTol.Top = 12
    txtTol.Width = 48: txtTol.Height = 16
    txtTol.Text = "5.00"

    dY = dY + 44

    ' --- Cadre Indicateur pente ----------------------------------------------
    ' Reglages toujours actifs (pas de mode "personnalise") : la taille et la
    ' couleur du texte pente viennent du style de texte nomme choisi.
    Dim fraInd As MSForms.Frame
    Set fraInd = Me.Controls.Add("Forms.Frame.1", "fraInd")
    fraInd.Caption = "Indicateur pente (texte + fleche)"
    fraInd.Left = 6: fraInd.Top = dY
    fraInd.Width = 192: fraInd.Height = 128

    CreerLabel fraInd, "lblNiv", "Niveau (vide = niveau actif) :", 6, 10, 178
    Set cmbNiveau = fraInd.Controls.Add("Forms.ComboBox.1", "cmbNiveau")
    cmbNiveau.Left = 6: cmbNiveau.Top = 24
    cmbNiveau.Width = 180: cmbNiveau.Height = 16

    CreerLabel fraInd, "lblStyle", "Style de texte (vide = style actif) :", 6, 46, 178
    Set cmbStyle = fraInd.Controls.Add("Forms.ComboBox.1", "cmbStyle")
    cmbStyle.Left = 6: cmbStyle.Top = 60
    cmbStyle.Width = 180: cmbStyle.Height = 16

    CreerLabel fraInd, "lblDec", "Dec:", 6, 82, 24
    Set txtDec = fraInd.Controls.Add("Forms.TextBox.1", "txtDec")
    txtDec.Left = 32: txtDec.Top = 80: txtDec.Width = 28: txtDec.Height = 16

    CreerLabel fraInd, "lblCF", "C.fl:", 68, 82, 24
    Set txtCoulFleche = fraInd.Controls.Add("Forms.TextBox.1", "txtCoulFleche")
    txtCoulFleche.Left = 94: txtCoulFleche.Top = 80
    txtCoulFleche.Width = 24: txtCoulFleche.Height = 16

    CreerLabel fraInd, "lblLF", "Long:", 126, 82, 24
    Set txtLongFleche = fraInd.Controls.Add("Forms.TextBox.1", "txtLongFleche")
    txtLongFleche.Left = 152: txtLongFleche.Top = 80
    txtLongFleche.Width = 34: txtLongFleche.Height = 16

    CreerLabel fraInd, "lblLP", "Pointe:", 6, 104, 40
    Set txtLongPointe = fraInd.Controls.Add("Forms.TextBox.1", "txtLongPointe")
    txtLongPointe.Left = 46: txtLongPointe.Top = 102
    txtLongPointe.Width = 30: txtLongPointe.Height = 16

    Set chkFermee = fraInd.Controls.Add("Forms.CheckBox.1", "chkFermee")
    chkFermee.Caption = "Fermee (triangle)"
    chkFermee.Left = 82: chkFermee.Top = 104
    chkFermee.Width = 104: chkFermee.Height = 14
End Sub

'------------------------------------------------------------------------------
Private Function CreerLabel(oParent As MSForms.Frame, sNom As String, _
                            sCaption As String, dLeft As Double, dTop As Double, _
                            dWidth As Double) As MSForms.Label
    Set CreerLabel = oParent.Controls.Add("Forms.Label.1", sNom)
    CreerLabel.Caption = sCaption
    CreerLabel.Left = dLeft: CreerLabel.Top = dTop
    CreerLabel.Width = dWidth: CreerLabel.Height = 12
End Function

'==============================================================================
' Initialisation
'==============================================================================

Sub Initialiser(oSettings As CMstSettings)
    ConstruireControles
    Set m_oSettings = oSettings
    m_bInit = True

    txtTol.Text = Format$(g_dTolReseau, "0.00")

    RemplirNiveaux cmbNiveau
    PositionnerNiveau cmbNiveau, m_oSettings.oIndicPente.NomNiveau
    RemplirStyles cmbStyle
    PositionnerStyle cmbStyle, m_oSettings.oIndicPente.NomStyle
    txtDec.Text = CStr(m_oSettings.oIndicPente.Decimales)
    txtLongFleche.Text = Format$(m_oSettings.oIndicPente.FlecheLongueur, "0.00")
    txtCoulFleche.Text = CStr(m_oSettings.oIndicPente.FlecheCouleur)
    txtLongPointe.Text = Format$(m_oSettings.oIndicPente.PointeLongueur, "0.00")
    chkFermee.Value = m_oSettings.oIndicPente.PointeFermee

    ReinitialiserEtat

    m_bInit = False
End Sub

'------------------------------------------------------------------------------
Private Sub RemplirNiveaux(cmb As MSForms.ComboBox)
    cmb.Clear
    cmb.AddItem ""   ' premier element vide = niveau actif
    Dim oLvl As Level
    For Each oLvl In ActiveDesignFile.Levels
        cmb.AddItem oLvl.Number & " : " & oLvl.Name
    Next
    cmb.ListIndex = 0
End Sub

'------------------------------------------------------------------------------
Private Sub PositionnerNiveau(cmb As MSForms.ComboBox, sNom As String)
    If Len(sNom) = 0 Then cmb.ListIndex = 0: Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If ExtraireNiveau(cmb.List(i)) = sNom Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next
    cmb.ListIndex = 0
End Sub

'------------------------------------------------------------------------------
Private Sub RemplirStyles(cmb As MSForms.ComboBox)
    cmb.Clear
    cmb.AddItem ""   ' premier element vide = style actif
    Dim oTS As TextStyle
    For Each oTS In ActiveDesignFile.TextStyles
        cmb.AddItem oTS.Name
    Next
    cmb.ListIndex = 0
End Sub

'------------------------------------------------------------------------------
Private Sub PositionnerStyle(cmb As MSForms.ComboBox, sNom As String)
    If Len(sNom) = 0 Then cmb.ListIndex = 0: Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = sNom Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next
    cmb.ListIndex = 0
End Sub

'------------------------------------------------------------------------------
' Extrait le niveau d'un item de combo "42 : Altimetrie" -> "42".
Private Function ExtraireNiveau(ByVal sItem As String) As String
    sItem = Trim$(sItem)
    If InStr(sItem, " : ") > 0 Then
        ExtraireNiveau = Trim$(Left$(sItem, InStr(sItem, " : ") - 1))
    Else
        ExtraireNiveau = sItem
    End If
End Function

'==============================================================================
' Mise a jour depuis les classes de commande
'==============================================================================

Sub AfficherBoiteP1(oInfo As CBoiteEPInfo)
    lblP1.Caption = "P1 : " & DecrireBoite(oInfo)
End Sub

Sub AfficherBoiteP2(oInfo As CBoiteEPInfo)
    lblP2.Caption = "P2 : " & DecrireBoite(oInfo)
End Sub

Sub AfficherDistancePente(dDist As Double, dPente As Double)
    lblDist.Caption = "Distance : " & Format$(dDist, "0.00") & " u.m."
    lblPente.Caption = "Pente : " & Format$(dPente, "0.0") & " %"
End Sub

Sub ReinitialiserEtat()
    lblP1.Caption = "P1 : -"
    lblP2.Caption = "P2 : -"
    lblDist.Caption = "Distance : -"
    lblPente.Caption = "Pente : -"
End Sub

' Efface seulement les infos de P2 (pente + distance), pour l'auto-chainage
' apres placement pente (P2 devient P1 et on repart snap P2).
Sub EffacerP2()
    lblP2.Caption = "P2 : -"
    lblDist.Caption = "Distance : -"
    lblPente.Caption = "Pente : -"
End Sub

'==============================================================================
' Evenements Recherche
'==============================================================================

Private Sub txtTol_Change()
    If m_bInit Then Exit Sub
    Dim d As Double
    d = Val(Replace(Trim$(txtTol.Text), ",", "."))
    If d > 0 Then g_dTolReseau = d
End Sub

Private Sub txtTol_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                            ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then _
        txtTol.Text = Format$(g_dTolReseau, "0.00")
End Sub

'==============================================================================
' Evenements Indicateur pente
'==============================================================================

Private Sub cmbNiveau_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oIndicPente.NomNiveau = ExtraireNiveau(cmbNiveau.Text)
End Sub

Private Sub cmbStyle_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oIndicPente.NomStyle = Trim$(cmbStyle.Text)
End Sub

Private Sub txtDec_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim nDec As Integer
    nDec = CInt(Val(Trim$(txtDec.Text)))
    If nDec >= 0 And nDec <= 6 Then m_oSettings.oIndicPente.Decimales = nDec
End Sub

Private Sub txtDec_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                            ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtDec.Text = CStr(m_oSettings.oIndicPente.Decimales)
End Sub

Private Sub txtLongFleche_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim d As Double
    d = Val(Replace(Trim$(txtLongFleche.Text), ",", "."))
    If d > 0 Then m_oSettings.oIndicPente.FlecheLongueur = d
End Sub

Private Sub txtCoulFleche_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtCoulFleche.Text)
    If sVal = "" Then Exit Sub
    Dim n As Long: n = CLng(Val(sVal))
    If n >= 0 And n <= 255 Then m_oSettings.oIndicPente.FlecheCouleur = n
End Sub

Private Sub txtLongPointe_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim d As Double
    d = Val(Replace(Trim$(txtLongPointe.Text), ",", "."))
    If d > 0 Then m_oSettings.oIndicPente.PointeLongueur = d
End Sub

Private Sub txtLongPointe_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                    ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtLongPointe.Text = Format$(m_oSettings.oIndicPente.PointeLongueur, "0.00")
End Sub

Private Sub chkFermee_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oIndicPente.PointeFermee = (chkFermee.Value = True)
End Sub

'==============================================================================
' Fermeture par la croix : quitter la commande
'==============================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
        CommandState.StartDefaultCommand
    End If
End Sub
