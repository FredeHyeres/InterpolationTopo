VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInterpolation 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmInterpolation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInterpolation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmInterpolation - Formulaire "Tool Settings" de la commande Interpolation Topo V2
'
' Fonctionne en mode modeless : reste visible pendant toute la commande.
' Lit et ecrit g_oSettings (CMstSettings).
' Mise a jour de l'etat (P1, P2, segment) par les classes de commande.
'
' IMPORTANT - CONSTRUCTION AU RUNTIME :
'   Tous les controles sont crees par code dans ConstruireControles (appele par
'   UserForm_Initialize). Le .frm ne decrit donc que le formulaire vide : aucun
'   fichier .frx n'est necessaire et le fichier reste versionnable en texte.
'
'   Si l'import du .frm echoue malgre tout : dans l'editeur VBA, Insertion >
'   UserForm, nommer le formulaire frmInterpolation, puis coller tout le code
'   ci-dessous (a partir de Option Explicit) dans son module.
'
' Controles (crees au runtime) :
'   fraCercle
'     txtDiametre    TextBox  - diametre du cercle (Double > 0)
'     txtCouleur     TextBox  - index couleur MicroStation 0-255
'     cmbNiveau      ComboBox - niveaux du fichier DGN (vide = niveau P1)
'   fraTexte
'     chkTexteModele CheckBox - True = memes attributs que le texte P1 (defaut)
'     txtCouleurTexte TextBox - couleur du texte cree (si case decochee)
'     cmbNiveauTexte ComboBox - niveau du texte cree (si case decochee)
'   fraEtat
'     lblP1          Label    - affiche "P1 : <altitude>" apres selection
'     lblP2          Label    - affiche "P2 : <altitude>" apres selection
'     lblSegment     Label    - affiche "Pente=x%  Gisement=x gon"
'
' Les saisies sont validees en continu (evenement Change) : la derniere valeur
' valide est conservee dans les parametres, une saisie invalide est ignoree.
' Entree dans un champ reformate l'affichage depuis la valeur courante.
'==============================================================================
Option Explicit

Private m_oSettings As CMstSettings  ' reference aux parametres de la commande
Private m_bInit     As Boolean       ' True pendant l'initialisation (bloque les evenements Change)
Private m_bConstruit As Boolean      ' True quand les controles ont ete crees

' Controles crees au runtime. WithEvents sur les champs de saisie :
' les procedures txtDiametre_Change etc. recoivent ainsi leurs evenements.
Private WithEvents txtDiametre As MSForms.TextBox
Attribute txtDiametre.VB_VarHelpID = -1
Private WithEvents txtCouleur  As MSForms.TextBox
Attribute txtCouleur.VB_VarHelpID = -1
Private WithEvents cmbNiveau   As MSForms.ComboBox
Attribute cmbNiveau.VB_VarHelpID = -1
Private WithEvents chkTexteModele As MSForms.CheckBox
Attribute chkTexteModele.VB_VarHelpID = -1
Private WithEvents txtCouleurTexte As MSForms.TextBox
Attribute txtCouleurTexte.VB_VarHelpID = -1
Private WithEvents cmbNiveauTexte As MSForms.ComboBox
Attribute cmbNiveauTexte.VB_VarHelpID = -1
Private WithEvents chkPente As MSForms.CheckBox
Attribute chkPente.VB_VarHelpID = -1
Private WithEvents txtPente As MSForms.TextBox
Attribute txtPente.VB_VarHelpID = -1
Private WithEvents btnInverserPente As MSForms.CommandButton
Attribute btnInverserPente.VB_VarHelpID = -1
Private WithEvents chkDecalageFixe As MSForms.CheckBox
Attribute chkDecalageFixe.VB_VarHelpID = -1
Private WithEvents txtDecalageDZ As MSForms.TextBox
Attribute txtDecalageDZ.VB_VarHelpID = -1
Private WithEvents btnRetourInterp As MSForms.CommandButton
Attribute btnRetourInterp.VB_VarHelpID = -1
Private WithEvents btnPlacerPente As MSForms.CommandButton
Attribute btnPlacerPente.VB_VarHelpID = -1
Private lblP1       As MSForms.Label
Private lblP2       As MSForms.Label
Private lblSegment  As MSForms.Label

'==============================================================================
' Construction des controles
'==============================================================================

Private Sub UserForm_Initialize()
    ConstruireControles
End Sub

'------------------------------------------------------------------------------
' Cree tous les controles par code (dimensions en points).
Private Sub ConstruireControles()
    If m_bConstruit Then Exit Sub
    m_bConstruit = True

    Me.Caption = "Interpolation Topo"
    Me.Width = 212
    Me.Height = 490

    ' --- Cadre Cercle -------------------------------------------------------
    Dim fraCercle As MSForms.Frame
    Set fraCercle = Me.Controls.Add("Forms.Frame.1", "fraCercle")
    fraCercle.Caption = "Cercle si Alti est un texte"
    fraCercle.Left = 6: fraCercle.Top = 6
    fraCercle.Width = 192: fraCercle.Height = 126

    CreerLabel fraCercle, "lblDiametre", "Diametre (unites maitre) :", 6, 10, 178
    Set txtDiametre = fraCercle.Controls.Add("Forms.TextBox.1", "txtDiametre")
    txtDiametre.Left = 6: txtDiametre.Top = 24
    txtDiametre.Width = 48: txtDiametre.Height = 16

    CreerLabel fraCercle, "lblCouleur", "Couleur (index MicroStation 0-255) :", 6, 46, 178
    Set txtCouleur = fraCercle.Controls.Add("Forms.TextBox.1", "txtCouleur")
    txtCouleur.Left = 6: txtCouleur.Top = 60
    txtCouleur.Width = 30: txtCouleur.Height = 16

    Dim lblInfo As MSForms.Label
    Set lblInfo = CreerLabel(fraCercle, "lblInfoCouleur", _
        "0=blanc 1=bleu 2=vert 3=rouge 4=jaune 6=orange", 42, 62, 146)
    lblInfo.Font.Size = 7

    CreerLabel fraCercle, "lblNiveau", "Niveau (vide = niveau du texte P1) :", 6, 82, 178
    Set cmbNiveau = fraCercle.Controls.Add("Forms.ComboBox.1", "cmbNiveau")
    cmbNiveau.Left = 6: cmbNiveau.Top = 96
    cmbNiveau.Width = 180: cmbNiveau.Height = 16

    ' --- Cadre Texte altitude ------------------------------------------------
    Dim fraTexte As MSForms.Frame
    Set fraTexte = Me.Controls.Add("Forms.Frame.1", "fraTexte")
    fraTexte.Caption = "Texte altitude si alti est un texte"
    fraTexte.Left = 6: fraTexte.Top = 138
    fraTexte.Width = 192: fraTexte.Height = 104

    Set chkTexteModele = fraTexte.Controls.Add("Forms.CheckBox.1", "chkTexteModele")
    chkTexteModele.Caption = "Memes attributs que le texte P1"
    chkTexteModele.Left = 6: chkTexteModele.Top = 10
    chkTexteModele.Width = 180: chkTexteModele.Height = 14
    chkTexteModele.Value = True

    CreerLabel fraTexte, "lblCouleurTexte", "Couleur (index MicroStation 0-255) :", 6, 28, 178
    Set txtCouleurTexte = fraTexte.Controls.Add("Forms.TextBox.1", "txtCouleurTexte")
    txtCouleurTexte.Left = 6: txtCouleurTexte.Top = 42
    txtCouleurTexte.Width = 30: txtCouleurTexte.Height = 16

    CreerLabel fraTexte, "lblNiveauTexte", "Niveau (vide = niveau du texte P1) :", 6, 64, 178
    Set cmbNiveauTexte = fraTexte.Controls.Add("Forms.ComboBox.1", "cmbNiveauTexte")
    cmbNiveauTexte.Left = 6: cmbNiveauTexte.Top = 78
    cmbNiveauTexte.Width = 180: cmbNiveauTexte.Height = 16

    ' --- Cadre Pente decalage -----------------------------------------------
    Dim fraPente As MSForms.Frame
    Set fraPente = Me.Controls.Add("Forms.Frame.1", "fraPente")
    fraPente.Caption = "Pente decalage"
    fraPente.Left = 6: fraPente.Top = 248
    fraPente.Width = 192: fraPente.Height = 96

    Set chkPente = fraPente.Controls.Add("Forms.CheckBox.1", "chkPente")
    chkPente.Caption = "Appliquer pente transversale"
    chkPente.Left = 6: chkPente.Top = 10
    chkPente.Width = 180: chkPente.Height = 14
    chkPente.Value = False

    CreerLabel fraPente, "lblPente", "Pente (%) :", 6, 30, 60
    Set txtPente = fraPente.Controls.Add("Forms.TextBox.1", "txtPente")
    txtPente.Left = 68: txtPente.Top = 28
    txtPente.Width = 48: txtPente.Height = 16
    txtPente.Text = "0"
    txtPente.Enabled = False

    Set btnInverserPente = fraPente.Controls.Add("Forms.CommandButton.1", "btnInverserPente")
    btnInverserPente.Caption = "+/-"
    btnInverserPente.Left = 122: btnInverserPente.Top = 28
    btnInverserPente.Width = 28: btnInverserPente.Height = 16
    btnInverserPente.Enabled = False

    Set chkDecalageFixe = fraPente.Controls.Add("Forms.CheckBox.1", "chkDecalageFixe")
    chkDecalageFixe.Caption = "Decalage fixe DZ"
    chkDecalageFixe.Left = 6: chkDecalageFixe.Top = 50
    chkDecalageFixe.Width = 180: chkDecalageFixe.Height = 14
    chkDecalageFixe.Value = False

    CreerLabel fraPente, "lblDZ", "DZ :", 6, 70, 24
    Set txtDecalageDZ = fraPente.Controls.Add("Forms.TextBox.1", "txtDecalageDZ")
    txtDecalageDZ.Left = 32: txtDecalageDZ.Top = 68
    txtDecalageDZ.Width = 48: txtDecalageDZ.Height = 16
    txtDecalageDZ.Text = "0"
    txtDecalageDZ.Enabled = False

    ' --- Cadre Etat ---------------------------------------------------------
    Dim fraEtat As MSForms.Frame
    Set fraEtat = Me.Controls.Add("Forms.Frame.1", "fraEtat")
    fraEtat.Caption = "Etat"
    fraEtat.Left = 6: fraEtat.Top = 350
    fraEtat.Width = 192: fraEtat.Height = 64

    Set lblP1 = CreerLabel(fraEtat, "lblP1", "P1 : -", 6, 12, 180)
    Set lblP2 = CreerLabel(fraEtat, "lblP2", "P2 : -", 6, 26, 180)
    Set lblSegment = CreerLabel(fraEtat, "lblSegment", "-", 6, 40, 180)

    ' --- Cadre Actions (disponible apres snap P2) ---------------------------
    Dim fraActions As MSForms.Frame
    Set fraActions = Me.Controls.Add("Forms.Frame.1", "fraActions")
    fraActions.Caption = "Actions"
    fraActions.Left = 6: fraActions.Top = 420
    fraActions.Width = 192: fraActions.Height = 46

    Set btnRetourInterp = fraActions.Controls.Add("Forms.CommandButton.1", "btnRetourInterp")
    btnRetourInterp.Caption = "Interpolation"
    btnRetourInterp.Left = 6: btnRetourInterp.Top = 14
    btnRetourInterp.Width = 86: btnRetourInterp.Height = 22
    btnRetourInterp.Enabled = False

    Set btnPlacerPente = fraActions.Controls.Add("Forms.CommandButton.1", "btnPlacerPente")
    btnPlacerPente.Caption = "Pente + fleche"
    btnPlacerPente.Left = 98: btnPlacerPente.Top = 14
    btnPlacerPente.Width = 88: btnPlacerPente.Height = 22
    btnPlacerPente.Enabled = False
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

'------------------------------------------------------------------------------
' Appele par InterpolationTopoV2 avant Show vbModeless.
Sub Initialiser(oSettings As CMstSettings)
    ConstruireControles   ' securite si Initialize n'a pas encore ete declenche
    Set m_oSettings = oSettings
    m_bInit = True

    ' Cercle : pre-remplir depuis les parametres
    txtDiametre.Text = Format$(m_oSettings.oCercle.Diametre, "0.00")
    txtCouleur.Text = CStr(m_oSettings.oCercle.Couleur)

    ' Niveaux : remplir les listes depuis le fichier DGN
    RemplirNiveaux cmbNiveau
    RemplirNiveaux cmbNiveauTexte

    ' Texte altitude : case + champs selon le mode
    chkTexteModele.Value = m_oSettings.oTexte.CommeModele
    txtCouleurTexte.Text = CStr(m_oSettings.oTexte.Couleur)
    ActiverChampsTexte

    ' Pente decalage
    chkPente.Value = m_oSettings.bAppliquerPente
    txtPente.Text = Format$(m_oSettings.dPenteDecalage, "0.00")
    txtPente.Enabled = m_oSettings.bAppliquerPente
    btnInverserPente.Enabled = m_oSettings.bAppliquerPente

    ' Decalage fixe DZ
    chkDecalageFixe.Value = m_oSettings.bDecalageFixe
    txtDecalageDZ.Text = Format$(m_oSettings.dDecalageDZ, "0.00")
    txtDecalageDZ.Enabled = m_oSettings.bDecalageFixe

    ' Etat : reinitialiser
    ReinitialiserEtat

    m_bInit = False
End Sub

'------------------------------------------------------------------------------
Private Sub RemplirNiveaux(cmb As MSForms.ComboBox)
    cmb.Clear
    cmb.AddItem ""   ' premier element vide = niveau du texte P1
    Dim oLvl As Level
    For Each oLvl In ActiveDesignFile.Levels
        cmb.AddItem oLvl.Number & " : " & oLvl.Name
    Next
    cmb.ListIndex = 0  ' vide selectionne par defaut
End Sub

'------------------------------------------------------------------------------
' Active/desactive les champs du texte selon la case a cocher
Private Sub ActiverChampsTexte()
    Dim bLibre As Boolean
    bLibre = Not m_oSettings.oTexte.CommeModele
    txtCouleurTexte.Enabled = bLibre
    cmbNiveauTexte.Enabled = bLibre
End Sub

'------------------------------------------------------------------------------
' Appele par CSelectP1 apres la selection de P1 : en mode modele, affiche
' la couleur et le niveau herites du texte selectionne (champs grises).
Sub RafraichirTexte()
    If m_oSettings Is Nothing Then Exit Sub
    m_bInit = True
    txtCouleurTexte.Text = CStr(m_oSettings.oTexte.Couleur)
    cmbNiveauTexte.Text = m_oSettings.oTexte.NomNiveau
    m_bInit = False
End Sub

'==============================================================================
' Mise a jour de l'etat par les classes de commande
'==============================================================================

Sub AfficherP1(oP As CPointRef)
    lblP1.Caption = "P1 : " & oP.Altitude
End Sub

Sub AfficherP2(oP As CPointRef)
    lblP2.Caption = "P2 : " & oP.Altitude
End Sub

Sub AfficherSegment(sInfo As String)
    lblSegment.Caption = sInfo
End Sub

Sub ReinitialiserEtat()
    lblP1.Caption = "P1 : -"
    lblP2.Caption = "P2 : -"
    lblSegment.Caption = "-"
    ActiverBoutonsActions False
End Sub

'------------------------------------------------------------------------------
' Active ou desactive les boutons Interpolation/Pente (disponibles apres snap P2).
Sub ActiverBoutonsActions(bActif As Boolean)
    btnRetourInterp.Enabled = bActif
    btnPlacerPente.Enabled = bActif
End Sub

'==============================================================================
' Evenements des controles : validation et ecriture dans m_oSettings
'
' Les controles etant crees au runtime, l'evenement Exit (extender) n'est pas
' disponible via WithEvents : la validation se fait sur Change. Une saisie
' invalide ou incomplete est simplement ignoree (les parametres gardent la
' derniere valeur valide) ; Entree reformate le champ depuis cette valeur.
'==============================================================================

Private Sub txtDiametre_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dDiam As Double
    dDiam = Val(Replace(Trim$(txtDiametre.Text), ",", "."))
    If dDiam > 0 Then m_oSettings.oCercle.Diametre = dDiam
End Sub

Private Sub txtDiametre_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtDiametre.Text = Format$(m_oSettings.oCercle.Diametre, "0.00")
End Sub

Private Sub txtCouleur_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String
    sVal = Trim$(txtCouleur.Text)
    If sVal = "" Then Exit Sub
    Dim nCoul As Long
    nCoul = CLng(Val(sVal))
    If nCoul >= 0 And nCoul <= 255 Then m_oSettings.oCercle.Couleur = nCoul
End Sub

Private Sub txtCouleur_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                               ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtCouleur.Text = CStr(m_oSettings.oCercle.Couleur)
End Sub

Private Sub cmbNiveau_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oCercle.NomNiveau = ExtraireNiveau(cmbNiveau.Text)
End Sub

'------------------------------------------------------------------------------
' Case a cocher : True = attributs du texte P1, False = choix libre
Private Sub chkTexteModele_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.CommeModele = (chkTexteModele.Value = True)
    ActiverChampsTexte
    ' Retour au mode modele : recharger les attributs du texte P1 si connu
    If m_oSettings.oTexte.CommeModele And m_oSettings.TextModeleDisponible Then
        m_oSettings.oTexte.ChargerDepuisElement m_oSettings.oTextModele
        RafraichirTexte
    End If
End Sub

Private Sub txtCouleurTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim sVal As String
    sVal = Trim$(txtCouleurTexte.Text)
    If sVal = "" Then Exit Sub
    Dim nCoul As Long
    nCoul = CLng(Val(sVal))
    If nCoul >= 0 And nCoul <= 255 Then m_oSettings.oTexte.Couleur = nCoul
End Sub

Private Sub txtCouleurTexte_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                    ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtCouleurTexte.Text = CStr(m_oSettings.oTexte.Couleur)
End Sub

Private Sub cmbNiveauTexte_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.oTexte.NomNiveau = ExtraireNiveau(cmbNiveauTexte.Text)
End Sub

'------------------------------------------------------------------------------
' Extrait le niveau d'un item de combo "42 : Altimetrie" -> "42".
' ResolverNiveau (CMoteurGraphique) accepte nom ou numero ; vide = niveau P1.
Private Function ExtraireNiveau(ByVal sItem As String) As String
    sItem = Trim$(sItem)
    If InStr(sItem, " : ") > 0 Then
        ExtraireNiveau = Trim$(Left$(sItem, InStr(sItem, " : ") - 1))
    Else
        ExtraireNiveau = sItem
    End If
End Function

'==============================================================================
' Evenements Pente decalage
'==============================================================================

Private Sub chkPente_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bAppliquerPente = (chkPente.Value = True)
    txtPente.Enabled = m_oSettings.bAppliquerPente
    btnInverserPente.Enabled = m_oSettings.bAppliquerPente
End Sub

Private Sub btnInverserPente_Click()
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.dPenteDecalage = -m_oSettings.dPenteDecalage
    txtPente.Text = Format$(m_oSettings.dPenteDecalage, "0.00")
End Sub

Private Sub txtPente_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dP As Double
    dP = Val(Replace(Trim$(txtPente.Text), ",", "."))
    m_oSettings.dPenteDecalage = dP
End Sub

Private Sub txtPente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtPente.Text = Format$(m_oSettings.dPenteDecalage, "0.00")
End Sub

'==============================================================================
' Evenements Decalage fixe DZ
'==============================================================================

Private Sub chkDecalageFixe_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    m_oSettings.bDecalageFixe = (chkDecalageFixe.Value = True)
    txtDecalageDZ.Enabled = m_oSettings.bDecalageFixe
End Sub

Private Sub txtDecalageDZ_Change()
    If m_bInit Then Exit Sub
    If m_oSettings Is Nothing Then Exit Sub
    Dim dDZ As Double
    dDZ = Val(Replace(Trim$(txtDecalageDZ.Text), ",", "."))
    m_oSettings.dDecalageDZ = dDZ
End Sub

Private Sub txtDecalageDZ_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                   ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Not m_oSettings Is Nothing Then _
        txtDecalageDZ.Text = Format$(m_oSettings.dDecalageDZ, "0.00")
End Sub

'==============================================================================
' Boutons Actions (apres snap P2)
'==============================================================================

Private Sub btnPlacerPente_Click()
    CommandState.StartPrimitive New CPlacerPente
End Sub

Private Sub btnRetourInterp_Click()
    CommandState.StartPrimitive New CPlacerPoint
End Sub

'==============================================================================
' Fermeture
'==============================================================================

'------------------------------------------------------------------------------
' Croix de la fenetre : terminer la commande SANS decharger le formulaire.
' Les classes de commande referencent l'instance par defaut : la decharger
' ici en recreerait une vierge (m_oSettings = Nothing) au prochain acces.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
        CommandState.StartDefaultCommand
    End If
End Sub







