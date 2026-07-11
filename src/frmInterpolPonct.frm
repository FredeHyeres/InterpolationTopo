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
'   Decimales        : commun chemin + rayon
'   Cercle           : diametre, couleur, niveau
'   Texte            : comme modele, couleur, niveau
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
' --- Texte ---
Private WithEvents chkTexteModele As MSForms.CheckBox
Attribute chkTexteModele.VB_VarHelpID = -1
Private WithEvents txtCouleurTexte As MSForms.TextBox
Attribute txtCouleurTexte.VB_VarHelpID = -1
Private WithEvents cmbNiveauTexte As MSForms.ComboBox
Attribute cmbNiveauTexte.VB_VarHelpID = -1
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
    Me.Height = 480

    Dim dY As Double
    dY = 6

    ' --- Cadre Chemin principal -----------------------------------------------
    Dim fraChemin As MSForms.Frame
    Set fraChemin = Me.Controls.Add("Forms.Frame.1", "fraChemin")
    fraChemin.Caption = "Chemin principal (Vert)"
    fraChemin.Left = 6: fraChemin.Top = dY
    fraChemin.Width = 192: fraChemin.Height = 76

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

    CreerLabel fraChemin, "lblInfoChem", _
        "Pente auto si P2 selectionne", 6, 54, 180

    dY = dY + 82

    ' --- Cadre Rayonnement ----------------------------------------------------
    Dim fraRayon As MSForms.Frame
    Set fraRayon = Me.Controls.Add("Forms.Frame.1", "fraRayon")
    fraRayon.Caption = "Rayonnement (Jaune)"
    fraRayon.Left = 6: fraRayon.Top = dY
    fraRayon.Width = 192: fraRayon.Height = 56

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

    dY = dY + 62

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
    fraCercle.Width = 192: fraCercle.Height = 76

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

    dY = dY + 82

    ' --- Cadre Texte ----------------------------------------------------------
    Dim fraTexte As MSForms.Frame
    Set fraTexte = Me.Controls.Add("Forms.Frame.1", "fraTexte")
    fraTexte.Caption = "Texte altitude"
    fraTexte.Left = 6: fraTexte.Top = dY
    fraTexte.Width = 192: fraTexte.Height = 76

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

    dY = dY + 82

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

    ' Rayonnement
    chkRayonPente.Value = m_oSettings.oRayon.PenteActive
    txtRayonPente.Text = Format$(m_oSettings.oRayon.Pente, "0.00")
    txtRayonPente.Enabled = m_oSettings.oRayon.PenteActive
    btnInverserRayonPente.Enabled = m_oSettings.oRayon.PenteActive
    chkRayonDZ.Value = m_oSettings.oRayon.DZActive
    txtRayonDZ.Text = Format$(m_oSettings.oRayon.DZ, "0.00")
    txtRayonDZ.Enabled = m_oSettings.oRayon.DZActive

    ' Decimales
    txtDecimales.Text = CStr(m_oSettings.nPonctDecimales)

    ' Cercle
    txtDiametre.Text = Format$(m_oSettings.oCercle.Diametre, "0.00")
    txtCouleur.Text = CStr(m_oSettings.oCercle.Couleur)
    RemplirNiveaux cmbNiveau

    ' Texte
    chkTexteModele.Value = m_oSettings.oTexte.CommeModele
    txtCouleurTexte.Text = CStr(m_oSettings.oTexte.Couleur)
    RemplirNiveaux cmbNiveauTexte
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

Private Sub ActiverChampsTexte()
    Dim bLibre As Boolean
    bLibre = Not m_oSettings.oTexte.CommeModele
    txtCouleurTexte.Enabled = bLibre
    cmbNiveauTexte.Enabled = bLibre
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
