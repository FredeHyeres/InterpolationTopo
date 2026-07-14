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
'   Indicateur  : parametres perso hauteur/couleur du texte pente et de la fleche,
'                 independants des reglages Interpolation V2
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
Private WithEvents chkPerso As MSForms.CheckBox
Attribute chkPerso.VB_VarHelpID = -1
Private WithEvents txtHauteur As MSForms.TextBox
Attribute txtHauteur.VB_VarHelpID = -1
Private WithEvents txtCoulTxt As MSForms.TextBox
Attribute txtCoulTxt.VB_VarHelpID = -1
Private WithEvents txtLongFleche As MSForms.TextBox
Attribute txtLongFleche.VB_VarHelpID = -1
Private WithEvents txtCoulFleche As MSForms.TextBox
Attribute txtCoulFleche.VB_VarHelpID = -1

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
    Me.Height = 340

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
    Dim fraInd As MSForms.Frame
    Set fraInd = Me.Controls.Add("Forms.Frame.1", "fraInd")
    fraInd.Caption = "Indicateur pente (texte + fleche)"
    fraInd.Left = 6: fraInd.Top = dY
    fraInd.Width = 192: fraInd.Height = 130

    Set chkPerso = fraInd.Controls.Add("Forms.CheckBox.1", "chkPerso")
    chkPerso.Caption = "Reglages personnalises"
    chkPerso.Left = 6: chkPerso.Top = 10
    chkPerso.Width = 180: chkPerso.Height = 14
    chkPerso.Value = True

    CreerLabel fraInd, "lblHt", "Hauteur texte :", 6, 32, 78
    Set txtHauteur = fraInd.Controls.Add("Forms.TextBox.1", "txtHauteur")
    txtHauteur.Left = 108: txtHauteur.Top = 30
    txtHauteur.Width = 48: txtHauteur.Height = 16

    CreerLabel fraInd, "lblCTxt", "Couleur texte :", 6, 52, 78
    Set txtCoulTxt = fraInd.Controls.Add("Forms.TextBox.1", "txtCoulTxt")
    txtCoulTxt.Left = 108: txtCoulTxt.Top = 50
    txtCoulTxt.Width = 30: txtCoulTxt.Height = 16

    CreerLabel fraInd, "lblLF", "Longueur fleche :", 6, 72, 90
    Set txtLongFleche = fraInd.Controls.Add("Forms.TextBox.1", "txtLongFleche")
    txtLongFleche.Left = 108: txtLongFleche.Top = 70
    txtLongFleche.Width = 48: txtLongFleche.Height = 16

    CreerLabel fraInd, "lblCF", "Couleur fleche :", 6, 92, 90
    Set txtCoulFleche = fraInd.Controls.Add("Forms.TextBox.1", "txtCoulFleche")
    txtCoulFleche.Left = 108: txtCoulFleche.Top = 90
    txtCoulFleche.Width = 30: txtCoulFleche.Height = 16

    CreerLabel fraInd, "lblAide", "(decoche = defauts V2)", 6, 112, 180
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

    ' Par defaut, on active les reglages perso (sinon la V2 n'ayant pas de
    ' texte-modele P1, les tailles heritees restent a la valeur par defaut 0.1).
    m_oSettings.oIndicPente.Perso = (chkPerso.Value = True)

    txtTol.Text = Format$(g_dTolReseau, "0.00")
    txtHauteur.Text = Format$(m_oSettings.oIndicPente.Hauteur, "0.000")
    txtCoulTxt.Text = CStr(m_oSettings.oIndicPente.Couleur)
    txtLongFleche.Text = Format$(m_oSettings.oIndicPente.FlecheLongueur, "0.00")
    txtCoulFleche.Text = CStr(m_oSettings.oIndicPente.FlecheCouleur)

    ActiverChampsPerso
    ReinitialiserEtat

    m_bInit = False
End Sub

Private Sub ActiverChampsPerso()
    Dim b As Boolean
    b = (chkPerso.Value = True)
    txtHauteur.Enabled = b
    txtCoulTxt.Enabled = b
    txtLongFleche.Enabled = b
    txtCoulFleche.Enabled = b
End Sub

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

Private Sub chkPerso_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oIndicPente.Perso = (chkPerso.Value = True)
    ActiverChampsPerso
End Sub

Private Sub txtHauteur_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim d As Double
    d = Val(Replace(Trim$(txtHauteur.Text), ",", "."))
    If d > 0 Then
        m_oSettings.oIndicPente.Hauteur = d
        m_oSettings.oIndicPente.Largeur = d
    End If
End Sub

Private Sub txtCoulTxt_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtCoulTxt.Text)
    If sVal = "" Then Exit Sub
    Dim n As Long: n = CLng(Val(sVal))
    If n >= 0 And n <= 255 Then m_oSettings.oIndicPente.Couleur = n
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
