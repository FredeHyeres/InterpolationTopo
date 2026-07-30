VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInterpolPonct
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmInterpolPonct.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "frmInterpolPonct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmInterpolPonct - Formulaire de la commande Interpolation Ponctuelle
'
' Cadres :
'   Chemin principal : pente (%), DZ, checkbox
'   Rayonnement      : pente (%), DZ, checkbox
'   Indicateur pente : niveau, style de texte, decimales, couleur/longueur
'                      du corps de la fleche, longueur de la pointe et choix
'                      pointe ouverte (2 lignes) / fermee (triangle plein)
'                      (reglages toujours actifs, partages entre "Pente +
'                      fleche" du chemin et du rayonnement) -- meme cadre que
'                      frmInterpolation (fraIndPente). La fleche (corps +
'                      pointe) est creee comme une seule cellule orpheline
'                      ("IndicPente") pour etre selectionnable en un clic.
'   Decimales        : commun chemin + rayon
'   Cercle           : diametre, couleur, niveau, plein/vide (defaut plein)
'   Texte            : comme modele, couleur, niveau, style de texte
'   Etat             : P1, P2 (optionnel)
'
' Tous les controles sont crees au runtime dans ConstruireControles.
'==============================================================================
Option Explicit

Private m_oSettings As CMstSettings
Private m_bInit     As Boolean
Private m_bConstruit As Boolean

' --- Chemin principal ---
Private WithEvents chkCheminPente As MSForms.CheckBox
Attribute chkCheminPente.VB_VarHelpID = -1
Private WithEvents txtCheminPente As MSForms.TextBox
Attribute txtCheminPente.VB_VarHelpID = -1
Private WithEvents btnInverserCheminPente As MSForms.CommandButton
Attribute btnInverserCheminPente.VB_VarHelpID = -1
Private WithEvents chkCheminDZ As MSForms.CheckBox
Attribute chkCheminDZ.VB_VarHelpID = -1
Private WithEvents txtCheminDZ As MSForms.TextBox
Attribute txtCheminDZ.VB_VarHelpID = -1
Private WithEvents chkCheminIndPente As MSForms.CheckBox
Attribute chkCheminIndPente.VB_VarHelpID = -1
' --- Rayonnement ---
Private WithEvents chkRayonPente As MSForms.CheckBox
Attribute chkRayonPente.VB_VarHelpID = -1
Private WithEvents txtRayonPente As MSForms.TextBox
Attribute txtRayonPente.VB_VarHelpID = -1
Private WithEvents btnInverserRayonPente As MSForms.CommandButton
Attribute btnInverserRayonPente.VB_VarHelpID = -1
Private WithEvents chkRayonDZ As MSForms.CheckBox
Attribute chkRayonDZ.VB_VarHelpID = -1
Private WithEvents txtRayonDZ As MSForms.TextBox
Attribute txtRayonDZ.VB_VarHelpID = -1
' --- Rayonnement indicateur pente ---
Private WithEvents chkRayonIndPente As MSForms.CheckBox
Attribute chkRayonIndPente.VB_VarHelpID = -1
' --- Indicateur pente (niveau/style/fleche, partage chemin + rayon) ---
Private WithEvents cmbPenteNiveau As MSForms.ComboBox
Attribute cmbPenteNiveau.VB_VarHelpID = -1
Private WithEvents cmbPenteStyle As MSForms.ComboBox
Attribute cmbPenteStyle.VB_VarHelpID = -1
Private WithEvents txtPenteCoulFl As MSForms.TextBox
Attribute txtPenteCoulFl.VB_VarHelpID = -1
Private WithEvents txtPenteFlLong As MSForms.TextBox
Attribute txtPenteFlLong.VB_VarHelpID = -1
Private WithEvents txtPenteDec As MSForms.TextBox
Attribute txtPenteDec.VB_VarHelpID = -1
Private WithEvents txtPentePteLong As MSForms.TextBox
Attribute txtPentePteLong.VB_VarHelpID = -1
Private WithEvents chkPenteFermee As MSForms.CheckBox
Attribute chkPenteFermee.VB_VarHelpID = -1
' --- Decimales ---
Private WithEvents txtDecimales As MSForms.TextBox
Attribute txtDecimales.VB_VarHelpID = -1
' --- Cercle ---
Private WithEvents txtDiametre As MSForms.TextBox
Attribute txtDiametre.VB_VarHelpID = -1
Private WithEvents txtCouleur As MSForms.TextBox
Attribute txtCouleur.VB_VarHelpID = -1
Private WithEvents cmbNiveau As MSForms.ComboBox
Attribute cmbNiveau.VB_VarHelpID = -1
Private WithEvents chkCerclePlein As MSForms.CheckBox
Attribute chkCerclePlein.VB_VarHelpID = -1
' --- Texte ---
Private WithEvents chkTexteModele As MSForms.CheckBox
Attribute chkTexteModele.VB_VarHelpID = -1
Private WithEvents txtCouleurTexte As MSForms.TextBox
Attribute txtCouleurTexte.VB_VarHelpID = -1
Private WithEvents cmbNiveauTexte As MSForms.ComboBox
Attribute cmbNiveauTexte.VB_VarHelpID = -1
Private WithEvents cmbStyleTexte As MSForms.ComboBox
Attribute cmbStyleTexte.VB_VarHelpID = -1
' --- Etat ---
Private lblP1       As MSForms.Label
Private lblP2       As MSForms.Label

'==============================================================================
' Construction des controles
'==============================================================================

Private Sub UserForm_Initialize()
    ConstruireControles
End Sub

Private Sub ConstruireControles()
    If m_bConstruit Then Exit Sub
    m_bConstruit = True

    Me.Caption = "Interpol. Ponctuelle"
    Me.Width = 212
    Me.Height = 684

    Dim dY As Double
    dY = 6

    ' --- Cadre Chemin principal -----------------------------------------------
    Dim fraChemin As MSForms.Frame
    Set fraChemin = Me.Controls.Add("Forms.Frame.1", "fraChemin")
    fraChemin.Caption = "Chemin principal (Vert)"
    fraChemin.Left = 6: fraChemin.Top = dY
    fraChemin.Width = 192: fraChemin.Height = 94

    Set chkCheminPente = fraChemin.Controls.Add("Forms.CheckBox.1", "chkCheminPente")
    chkCheminPente.Caption = "Pente (%)"
    chkCheminPente.Left = 6: chkCheminPente.Top = 10
    chkCheminPente.Width = 70: chkCheminPente.Height = 14

    Set txtCheminPente = fraChemin.Controls.Add("Forms.TextBox.1", "txtCheminPente")
    txtCheminPente.Left = 78: txtCheminPente.Top = 9
    txtCheminPente.Width = 48: txtCheminPente.Height = 16
    txtCheminPente.Text = "0": txtCheminPente.Enabled = False

    Set btnInverserCheminPente = fraChemin.Controls.Add("Forms.CommandButton.1", "btnInvChP")
    btnInverserCheminPente.Caption = "+/-"
    btnInverserCheminPente.Left = 130: btnInverserCheminPente.Top = 9
    btnInverserCheminPente.Width = 28: btnInverserCheminPente.Height = 16
    btnInverserCheminPente.Enabled = False

    Set chkCheminDZ = fraChemin.Controls.Add("Forms.CheckBox.1", "chkCheminDZ")
    chkCheminDZ.Caption = "DZ"
    chkCheminDZ.Left = 6: chkCheminDZ.Top = 32
    chkCheminDZ.Width = 36: chkCheminDZ.Height = 14

    Set txtCheminDZ = fraChemin.Controls.Add("Forms.TextBox.1", "txtCheminDZ")
    txtCheminDZ.Left = 44: txtCheminDZ.Top = 31
    txtCheminDZ.Width = 48: txtCheminDZ.Height = 16
    txtCheminDZ.Text = "0": txtCheminDZ.Enabled = False

    Set chkCheminIndPente = fraChemin.Controls.Add("Forms.CheckBox.1", "chkCheminIndPente")
    chkCheminIndPente.Caption = "Pente + fleche"
    chkCheminIndPente.Left = 6: chkCheminIndPente.Top = 52
    chkCheminIndPente.Width = 100: chkCheminIndPente.Height = 14

    CreerLabel fraChemin, "lblInfoChem", _
        "Pente auto si P2 selectionne", 6, 72, 180

    dY = dY + 100

    ' --- Cadre Rayonnement ----------------------------------------------------
    Dim fraRayon As MSForms.Frame
    Set fraRayon = Me.Controls.Add("Forms.Frame.1", "fraRayon")
    fraRayon.Caption = "Rayonnement (Jaune)"
    fraRayon.Left = 6: fraRayon.Top = dY
    fraRayon.Width = 192: fraRayon.Height = 72

    Set chkRayonPente = fraRayon.Controls.Add("Forms.CheckBox.1", "chkRayonPente")
    chkRayonPente.Caption = "Pente (%)"
    chkRayonPente.Left = 6: chkRayonPente.Top = 10
    chkRayonPente.Width = 70: chkRayonPente.Height = 14

    Set txtRayonPente = fraRayon.Controls.Add("Forms.TextBox.1", "txtRayonPente")
    txtRayonPente.Left = 78: txtRayonPente.Top = 9
    txtRayonPente.Width = 48: txtRayonPente.Height = 16
    txtRayonPente.Text = "0": txtRayonPente.Enabled = False

    Set btnInverserRayonPente = fraRayon.Controls.Add("Forms.CommandButton.1", "btnInvRP")
    btnInverserRayonPente.Caption = "+/-"
    btnInverserRayonPente.Left = 130: btnInverserRayonPente.Top = 9
    btnInverserRayonPente.Width = 28: btnInverserRayonPente.Height = 16
    btnInverserRayonPente.Enabled = False

    Set chkRayonDZ = fraRayon.Controls.Add("Forms.CheckBox.1", "chkRayonDZ")
    chkRayonDZ.Caption = "DZ"
    chkRayonDZ.Left = 6: chkRayonDZ.Top = 32
    chkRayonDZ.Width = 36: chkRayonDZ.Height = 14

    Set txtRayonDZ = fraRayon.Controls.Add("Forms.TextBox.1", "txtRayonDZ")
    txtRayonDZ.Left = 44: txtRayonDZ.Top = 31
    txtRayonDZ.Width = 48: txtRayonDZ.Height = 16
    txtRayonDZ.Text = "0": txtRayonDZ.Enabled = False

    Set chkRayonIndPente = fraRayon.Controls.Add("Forms.CheckBox.1", "chkRayonIndPente")
    chkRayonIndPente.Caption = "Pente + fleche"
    chkRayonIndPente.Left = 6: chkRayonIndPente.Top = 52
    chkRayonIndPente.Width = 100: chkRayonIndPente.Height = 14

    dY = dY + 78

    ' --- Cadre Indicateur pente (niveau/style/fleche, chemin + rayon) ---------
    ' Pas de mode "personnalise" : ces reglages sont toujours actifs. La
    ' taille/couleur du texte pente vient du style de texte nomme choisi.
    Dim fraIndPente As MSForms.Frame
    Set fraIndPente = Me.Controls.Add("Forms.Frame.1", "fraIndPente")
    fraIndPente.Caption = "Indicateur pente"
    fraIndPente.Left = 6: fraIndPente.Top = dY
    fraIndPente.Width = 192: fraIndPente.Height = 128

    CreerLabel fraIndPente, "lblPenteNiv", "Niveau (vide = niveau actif) :", 6, 10, 178
    Set cmbPenteNiveau = fraIndPente.Controls.Add("Forms.ComboBox.1", "cmbPenteNiveau")
    cmbPenteNiveau.Left = 6: cmbPenteNiveau.Top = 24
    cmbPenteNiveau.Width = 180: cmbPenteNiveau.Height = 16

    CreerLabel fraIndPente, "lblPenteStyle", "Style de texte (vide = style actif) :", 6, 46, 178
    Set cmbPenteStyle = fraIndPente.Controls.Add("Forms.ComboBox.1", "cmbPenteStyle")
    cmbPenteStyle.Left = 6: cmbPenteStyle.Top = 60
    cmbPenteStyle.Width = 180: cmbPenteStyle.Height = 16

    CreerLabel fraIndPente, "lblPenteDec", "Dec:", 6, 82, 24
    Set txtPenteDec = fraIndPente.Controls.Add("Forms.TextBox.1", "txtPenteDec")
    txtPenteDec.Left = 32: txtPenteDec.Top = 80: txtPenteDec.Width = 28: txtPenteDec.Height = 16

    CreerLabel fraIndPente, "lblPCFl", "C.fl:", 68, 82, 24
    Set txtPenteCoulFl = fraIndPente.Controls.Add("Forms.TextBox.1", "txtPenteCoulFl")
    txtPenteCoulFl.Left = 94: txtPenteCoulFl.Top = 80
    txtPenteCoulFl.Width = 24: txtPenteCoulFl.Height = 16

    CreerLabel fraIndPente, "lblPFlLong", "Long:", 126, 82, 24
    Set txtPenteFlLong = fraIndPente.Controls.Add("Forms.TextBox.1", "txtPenteFlLong")
    txtPenteFlLong.Left = 152: txtPenteFlLong.Top = 80
    txtPenteFlLong.Width = 34: txtPenteFlLong.Height = 16

    CreerLabel fraIndPente, "lblPPteLong", "Pointe:", 6, 104, 40
    Set txtPentePteLong = fraIndPente.Controls.Add("Forms.TextBox.1", "txtPentePteLong")
    txtPentePteLong.Left = 46: txtPentePteLong.Top = 102
    txtPentePteLong.Width = 30: txtPentePteLong.Height = 16

    Set chkPenteFermee = fraIndPente.Controls.Add("Forms.CheckBox.1", "chkPenteFermee")
    chkPenteFermee.Caption = "Fermee (triangle)"
    chkPenteFermee.Left = 82: chkPenteFermee.Top = 104
    chkPenteFermee.Width = 104: chkPenteFermee.Height = 14

    dY = dY + 134

    ' --- Decimales ------------------------------------------------------------
    Dim fraDec As MSForms.Frame
    Set fraDec = Me.Controls.Add("Forms.Frame.1", "fraDec")
    fraDec.Caption = "Decimales (points crees)"
    fraDec.Left = 6: fraDec.Top = dY
    fraDec.Width = 192: fraDec.Height = 38

    Set txtDecimales = fraDec.Controls.Add("Forms.TextBox.1", "txtDecimales")
    txtDecimales.Left = 6: txtDecimales.Top = 12
    txtDecimales.Width = 24: txtDecimales.Height = 16
    txtDecimales.Text = "2"

    dY = dY + 44

    ' --- Cadre Cercle ---------------------------------------------------------
    Dim fraCercle As MSForms.Frame
    Set fraCercle = Me.Controls.Add("Forms.Frame.1", "fraCercle")
    fraCercle.Caption = "Cercle"
    fraCercle.Left = 6: fraCercle.Top = dY
    fraCercle.Width = 192: fraCercle.Height = 96

    CreerLabel fraCercle, "lblDiam", "Diametre :", 6, 10, 48
    Set txtDiametre = fraCercle.Controls.Add("Forms.TextBox.1", "txtDiametre")
    txtDiametre.Left = 56: txtDiametre.Top = 8
    txtDiametre.Width = 48: txtDiametre.Height = 16

    CreerLabel fraCercle, "lblCoul", "Couleur :", 6, 30, 42
    Set txtCouleur = fraCercle.Controls.Add("Forms.TextBox.1", "txtCouleur")
    txtCouleur.Left = 56: txtCouleur.Top = 28
    txtCouleur.Width = 30: txtCouleur.Height = 16

    CreerLabel fraCercle, "lblNiv", "Niveau :", 6, 50, 42
    Set cmbNiveau = fraCercle.Controls.Add("Forms.ComboBox.1", "cmbNiveau")
    cmbNiveau.Left = 56: cmbNiveau.Top = 48
    cmbNiveau.Width = 130: cmbNiveau.Height = 16

    Set chkCerclePlein = fraCercle.Controls.Add("Forms.CheckBox.1", "chkCerclePlein")
    chkCerclePlein.Caption = "Plein (rempli)"
    chkCerclePlein.Left = 6: chkCerclePlein.Top = 68
    chkCerclePlein.Width = 140: chkCerclePlein.Height = 14
    chkCerclePlein.Value = True

    dY = dY + 102

    ' --- Cadre Texte ----------------------------------------------------------
    Dim fraTexte As MSForms.Frame
    Set fraTexte = Me.Controls.Add("Forms.Frame.1", "fraTexte")
    fraTexte.Caption = "Texte altitude"
    fraTexte.Left = 6: fraTexte.Top = dY
    fraTexte.Width = 192: fraTexte.Height = 92

    Set chkTexteModele = fraTexte.Controls.Add("Forms.CheckBox.1", "chkTexteModele")
    chkTexteModele.Caption = "Memes attributs que P1"
    chkTexteModele.Left = 6: chkTexteModele.Top = 10
    chkTexteModele.Width = 180: chkTexteModele.Height = 14
    chkTexteModele.Value = True

    CreerLabel fraTexte, "lblCoulTxt", "Couleur :", 6, 30, 42
    Set txtCouleurTexte = fraTexte.Controls.Add("Forms.TextBox.1", "txtCouleurTexte")
    txtCouleurTexte.Left = 56: txtCouleurTexte.Top = 28
    txtCouleurTexte.Width = 30: txtCouleurTexte.Height = 16

    CreerLabel fraTexte, "lblNivTxt", "Niveau :", 6, 50, 42
    Set cmbNiveauTexte = fraTexte.Controls.Add("Forms.ComboBox.1", "cmbNiveauTexte")
    cmbNiveauTexte.Left = 56: cmbNiveauTexte.Top = 48
    cmbNiveauTexte.Width = 130: cmbNiveauTexte.Height = 16

    CreerLabel fraTexte, "lblStyleTxt", "Style :", 6, 70, 42
    Set cmbStyleTexte = fraTexte.Controls.Add("Forms.ComboBox.1", "cmbStyleTexte")
    cmbStyleTexte.Left = 56: cmbStyleTexte.Top = 68
    cmbStyleTexte.Width = 130: cmbStyleTexte.Height = 16

    dY = dY + 98

    ' --- Cadre Etat -----------------------------------------------------------
    Dim fraEtat As MSForms.Frame
    Set fraEtat = Me.Controls.Add("Forms.Frame.1", "fraEtat")
    fraEtat.Caption = "Etat"
    fraEtat.Left = 6: fraEtat.Top = dY
    fraEtat.Width = 192: fraEtat.Height = 44
    

    Set lblP1 = CreerLabel(fraEtat, "lblP1", "P1 : -", 6, 12, 180)
    Set lblP2 = CreerLabel(fraEtat, "lblP2", "P2 : -", 6, 24, 180)
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

    ' Chemin principal
    chkCheminPente.Value = m_oSettings.oChemin.PenteActive
    txtCheminPente.Text = Format$(m_oSettings.oChemin.Pente, "0.00")
    txtCheminPente.Enabled = m_oSettings.oChemin.PenteActive
    btnInverserCheminPente.Enabled = m_oSettings.oChemin.PenteActive
    chkCheminDZ.Value = m_oSettings.oChemin.DZActive
    txtCheminDZ.Text = Format$(m_oSettings.oChemin.DZ, "0.00")
    txtCheminDZ.Enabled = m_oSettings.oChemin.DZActive
    chkCheminIndPente.Value = m_oSettings.bCheminIndicPente

    ' Rayonnement
    chkRayonPente.Value = m_oSettings.oRayon.PenteActive
    txtRayonPente.Text = Format$(m_oSettings.oRayon.Pente, "0.00")
    txtRayonPente.Enabled = m_oSettings.oRayon.PenteActive
    btnInverserRayonPente.Enabled = m_oSettings.oRayon.PenteActive
    chkRayonDZ.Value = m_oSettings.oRayon.DZActive
    txtRayonDZ.Text = Format$(m_oSettings.oRayon.DZ, "0.00")
    txtRayonDZ.Enabled = m_oSettings.oRayon.DZActive
    chkRayonIndPente.Value = m_oSettings.bRayonIndicPente

    ' Indicateur pente (reglages toujours actifs, partages chemin + rayon)
    RemplirNiveaux cmbPenteNiveau
    PositionnerNiveau cmbPenteNiveau, m_oSettings.oIndicPente.NomNiveau
    RemplirStyles cmbPenteStyle
    PositionnerStyle cmbPenteStyle, m_oSettings.oIndicPente.NomStyle
    txtPenteCoulFl.Text = CStr(m_oSettings.oIndicPente.FlecheCouleur)
    txtPenteFlLong.Text = Format$(m_oSettings.oIndicPente.FlecheLongueur, "0.00")
    txtPenteDec.Text = CStr(m_oSettings.oIndicPente.Decimales)
    txtPentePteLong.Text = Format$(m_oSettings.oIndicPente.PointeLongueur, "0.00")
    chkPenteFermee.Value = m_oSettings.oIndicPente.PointeFermee

    ' Decimales
    txtDecimales.Text = CStr(m_oSettings.nPonctDecimales)

    ' Cercle
    txtDiametre.Text = Format$(m_oSettings.oCercle.Diametre, "0.00")
    txtCouleur.Text = CStr(m_oSettings.oCercle.Couleur)
    chkCerclePlein.Value = m_oSettings.oCercle.Plein
    RemplirNiveaux cmbNiveau
    PositionnerNiveau cmbNiveau, m_oSettings.oCercle.NomNiveau

    ' Texte
    chkTexteModele.Value = m_oSettings.oTexte.CommeModele
    txtCouleurTexte.Text = CStr(m_oSettings.oTexte.Couleur)
    RemplirNiveaux cmbNiveauTexte
    PositionnerNiveau cmbNiveauTexte, m_oSettings.oTexte.NomNiveau
    RemplirStyles cmbStyleTexte
    PositionnerStyle cmbStyleTexte, m_oSettings.oTexte.NomStyle
    ActiverChampsTexte

    ReinitialiserEtat
    m_bInit = False
End Sub

'------------------------------------------------------------------------------
Private Sub RemplirNiveaux(cmb As MSForms.ComboBox)
    cmb.Clear
    cmb.AddItem ""
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

Private Sub ActiverChampsTexte()
    Dim bLibre As Boolean
    bLibre = Not m_oSettings.oTexte.CommeModele
    txtCouleurTexte.Enabled = bLibre
    cmbNiveauTexte.Enabled = bLibre
    cmbStyleTexte.Enabled = bLibre
End Sub

'------------------------------------------------------------------------------
Private Sub RemplirStyles(cmb As MSForms.ComboBox)
    cmb.Clear
    cmb.AddItem ""
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

'==============================================================================
' Mise a jour par les classes de commande
'==============================================================================

Sub AfficherP1(oP As CPointRef)
    lblP1.Caption = "P1 : " & oP.Altitude
End Sub

Sub AfficherP2(oP As CPointRef)
    lblP2.Caption = "P2 : " & oP.Altitude
End Sub

Sub ReinitialiserEtat()
    lblP1.Caption = "P1 : -"
    lblP2.Caption = "P2 : -"
End Sub

Sub RafraichirTexte()
    If m_oSettings Is Nothing Then Exit Sub
    m_bInit = True
    txtCouleurTexte.Text = CStr(m_oSettings.oTexte.Couleur)
    cmbNiveauTexte.Text = m_oSettings.oTexte.NomNiveau
    cmbStyleTexte.Text = m_oSettings.oTexte.NomStyle
    m_bInit = False
End Sub

Sub RafraichirChemin()
    If m_oSettings Is Nothing Then Exit Sub
    m_bInit = True
    chkCheminPente.Value = m_oSettings.oChemin.PenteActive
    txtCheminPente.Text = Format$(m_oSettings.oChemin.Pente, "0.00")
    txtCheminPente.Enabled = m_oSettings.oChemin.PenteActive
    btnInverserCheminPente.Enabled = m_oSettings.oChemin.PenteActive
    m_bInit = False
End Sub

'==============================================================================
' Evenements Chemin principal
'==============================================================================

Private Sub chkCheminPente_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oChemin.PenteActive = (chkCheminPente.Value = True)
    txtCheminPente.Enabled = m_oSettings.oChemin.PenteActive
    btnInverserCheminPente.Enabled = m_oSettings.oChemin.PenteActive
End Sub

Private Sub txtCheminPente_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oChemin.Pente = Val(Replace(Trim$(txtCheminPente.Text), ",", "."))
End Sub

Private Sub txtCheminPente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                    ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtCheminPente.Text = Format$(m_oSettings.oChemin.Pente, "0.00")
End Sub

Private Sub btnInverserCheminPente_Click()
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oChemin.Pente = -m_oSettings.oChemin.Pente
    txtCheminPente.Text = Format$(m_oSettings.oChemin.Pente, "0.00")
End Sub

Private Sub chkCheminDZ_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oChemin.DZActive = (chkCheminDZ.Value = True)
    txtCheminDZ.Enabled = m_oSettings.oChemin.DZActive
End Sub

Private Sub chkCheminIndPente_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bCheminIndicPente = (chkCheminIndPente.Value = True)
End Sub

Private Sub txtCheminDZ_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oChemin.DZ = Val(Replace(Trim$(txtCheminDZ.Text), ",", "."))
End Sub

Private Sub txtCheminDZ_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                 ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtCheminDZ.Text = Format$(m_oSettings.oChemin.DZ, "0.00")
End Sub

'==============================================================================
' Evenements Rayonnement
'==============================================================================

Private Sub chkRayonPente_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oRayon.PenteActive = (chkRayonPente.Value = True)
    txtRayonPente.Enabled = m_oSettings.oRayon.PenteActive
    btnInverserRayonPente.Enabled = m_oSettings.oRayon.PenteActive
End Sub

Private Sub txtRayonPente_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oRayon.Pente = Val(Replace(Trim$(txtRayonPente.Text), ",", "."))
End Sub

Private Sub txtRayonPente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                   ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtRayonPente.Text = Format$(m_oSettings.oRayon.Pente, "0.00")
End Sub

Private Sub btnInverserRayonPente_Click()
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oRayon.Pente = -m_oSettings.oRayon.Pente
    txtRayonPente.Text = Format$(m_oSettings.oRayon.Pente, "0.00")
End Sub

Private Sub chkRayonDZ_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oRayon.DZActive = (chkRayonDZ.Value = True)
    txtRayonDZ.Enabled = m_oSettings.oRayon.DZActive
End Sub

Private Sub txtRayonDZ_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oRayon.DZ = Val(Replace(Trim$(txtRayonDZ.Text), ",", "."))
End Sub

Private Sub txtRayonDZ_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtRayonDZ.Text = Format$(m_oSettings.oRayon.DZ, "0.00")
End Sub

Private Sub chkRayonIndPente_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bRayonIndicPente = (chkRayonIndPente.Value = True)
End Sub

Private Sub txtPenteDec_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim nDec As Integer
    nDec = CInt(Val(Trim$(txtPenteDec.Text)))
    If nDec >= 0 And nDec <= 6 Then m_oSettings.oIndicPente.Decimales = nDec
End Sub

Private Sub txtPenteDec_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                 ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtPenteDec.Text = CStr(m_oSettings.oIndicPente.Decimales)
End Sub

Private Sub cmbPenteNiveau_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oIndicPente.NomNiveau = ExtraireNiveau(cmbPenteNiveau.Text)
End Sub

Private Sub cmbPenteStyle_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oIndicPente.NomStyle = Trim$(cmbPenteStyle.Text)
End Sub

Private Sub txtPenteCoulFl_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtPenteCoulFl.Text)
    If sVal = "" Then Exit Sub
    Dim nCoul As Long: nCoul = CLng(Val(sVal))
    If nCoul >= 0 And nCoul <= 255 Then m_oSettings.oIndicPente.FlecheCouleur = nCoul
End Sub

Private Sub txtPenteCoulFl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                    ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtPenteCoulFl.Text = CStr(m_oSettings.oIndicPente.FlecheCouleur)
End Sub

Private Sub txtPenteFlLong_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dVal As Double
    dVal = Val(Replace(Trim$(txtPenteFlLong.Text), ",", "."))
    If dVal > 0 Then m_oSettings.oIndicPente.FlecheLongueur = dVal
End Sub

Private Sub txtPenteFlLong_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                    ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtPenteFlLong.Text = Format$(m_oSettings.oIndicPente.FlecheLongueur, "0.00")
End Sub

Private Sub txtPentePteLong_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dVal As Double
    dVal = Val(Replace(Trim$(txtPentePteLong.Text), ",", "."))
    If dVal > 0 Then m_oSettings.oIndicPente.PointeLongueur = dVal
End Sub

Private Sub txtPentePteLong_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                     ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtPentePteLong.Text = Format$(m_oSettings.oIndicPente.PointeLongueur, "0.00")
End Sub

Private Sub chkPenteFermee_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oIndicPente.PointeFermee = (chkPenteFermee.Value = True)
End Sub

'==============================================================================
' Evenements Decimales
'==============================================================================

Private Sub txtDecimales_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim nDec As Integer
    nDec = CInt(Val(Trim$(txtDecimales.Text)))
    If nDec >= 0 And nDec <= 6 Then m_oSettings.nPonctDecimales = nDec
End Sub

Private Sub txtDecimales_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                  ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtDecimales.Text = CStr(m_oSettings.nPonctDecimales)
End Sub

'==============================================================================
' Evenements Cercle
'==============================================================================

Private Sub txtDiametre_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dDiam As Double
    dDiam = Val(Replace(Trim$(txtDiametre.Text), ",", "."))
    If dDiam > 0 Then m_oSettings.oCercle.Diametre = dDiam
End Sub

Private Sub txtCouleur_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtCouleur.Text)
    If sVal = "" Then Exit Sub
    Dim nCoul As Long: nCoul = CLng(Val(sVal))
    If nCoul >= 0 And nCoul <= 255 Then m_oSettings.oCercle.Couleur = nCoul
End Sub

Private Sub cmbNiveau_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oCercle.NomNiveau = ExtraireNiveau(cmbNiveau.Text)
End Sub

Private Sub chkCerclePlein_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oCercle.Plein = (chkCerclePlein.Value = True)
End Sub

'==============================================================================
' Evenements Texte
'==============================================================================

Private Sub chkTexteModele_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.CommeModele = (chkTexteModele.Value = True)
    ActiverChampsTexte
    If m_oSettings.oTexte.CommeModele And m_oSettings.TextModeleDisponible Then
        m_oSettings.oTexte.ChargerDepuisElement m_oSettings.oTextModele
        RafraichirTexte
    End If
End Sub

Private Sub txtCouleurTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String: sVal = Trim$(txtCouleurTexte.Text)
    If sVal = "" Then Exit Sub
    Dim nCoul As Long: nCoul = CLng(Val(sVal))
    If nCoul >= 0 And nCoul <= 255 Then m_oSettings.oTexte.Couleur = nCoul
End Sub

Private Sub cmbNiveauTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.NomNiveau = ExtraireNiveau(cmbNiveauTexte.Text)
End Sub

Private Sub cmbStyleTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.NomStyle = Trim$(cmbStyleTexte.Text)
End Sub

'------------------------------------------------------------------------------
Private Function ExtraireNiveau(ByVal sItem As String) As String
    sItem = Trim$(sItem)
    If InStr(sItem, " : ") > 0 Then
        ExtraireNiveau = Trim$(Left$(sItem, InStr(sItem, " : ") - 1))
    Else
        ExtraireNiveau = sItem
    End If
End Function

'==============================================================================
' Fermeture
'==============================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
        CommandState.StartDefaultCommand
    End If
End Sub
