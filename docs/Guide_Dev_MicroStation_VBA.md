# Guide de developpement VBA pour MicroStation V8i

*Reference pour la creation de nouvelles commandes VBA MicroStation V8i SS3,
basee sur l'architecture V2 d'Interpolation Topo (juillet 2026).*

---

## 1. Architecture cible : classes a responsabilite unique

Chaque commande suit la meme decomposition. Le module principal ne contient
que le point d'entree, les globals partages et les utilitaires de scan.
Toute la logique est repartie dans des classes.

```
MaCommande()                          ' Module .bas : point d'entree
      |
      +--- CSettings (g_oSettings)    ' Agregateur de parametres
      |         +-- CSymboXxx         '   sous-objets de symbologie
      |         +-- CSymboYyy
      |
      +--- CDonnees (g_oDonnees)      ' Donnees metier (pas d'API MST)
      +--- CCalcul (g_oCalc)          ' Calcul pur (pas d'API MST)
      +--- CGraphique (g_oMoteur)     ' Creation/apercu (API MST)
      |
      +--- frmMaCommande              ' Formulaire modeless (Tool Settings)
      |         <--> g_oSettings      '   lecture/ecriture bidirectionnelle
      |
      +--- CommandState
               +-- CSelect1 --> CSelect2 --> CPlacer
               '   (Reset remonte d'un cran, second Reset quitte)
```

### Repartition des responsabilites

| Domaine | Classe(s) | Connait l'API MST ? |
|---|---|---|
| Parametres | `CSettings` + sous-objets `CSymboXxx` | Non |
| Donnees metier | `CDonnees` | Non (sauf `Point3d` UDT) |
| Calcul | `CCalcul` | Non |
| Graphique | `CGraphique` | Oui |
| Commande | `CSelectXxx`, `CPlacer` | Oui (`IPrimitiveCommandEvents`) |
| Interface | `frmXxx` | Oui (lecture niveaux DGN, etc.) |

**Regle :** une classe de calcul ou de donnees ne doit jamais importer
l'API MicroStation. Le moteur graphique et les classes de commande sont
les seuls a manipuler des `Element`, `View`, `DesignFile`, etc.

---

## 2. Structure de fichiers

```
MonProjet/
  src/
    MaCommandeV1.bas            ' Module principal
    CSettings.cls               ' Agregateur parametres
    CSymboTexte.cls             ' Symbologie texte (sous-objet)
    CSymboCercle.cls            ' Symbologie cercle (sous-objet)
    CDonnees.cls                ' Donnees metier
    CCalcul.cls                 ' Moteur de calcul pur
    CGraphique.cls              ' Creation graphique + apercu dynamique
    CSelect1.cls                ' Etat commande : selection 1
    CSelect2.cls                ' Etat commande : selection 2
    CPlacer.cls                 ' Etat commande : placement
    frmMaCommande.frm           ' Formulaire (texte + code)
    frmMaCommande.frx           ' Blob designer (binaire, ne pas editer)
  scripts/
    export_frm_via_excel.ps1    ' Generation du couple .frm/.frx
  docs/
    Guide_Dev_MicroStation_VBA.md
  install.cmd                   ' Installeur (appelle install.ps1)
  install.ps1                   ' Script d'installation PowerShell
  .gitattributes                ' Encodage CRLF + .frx binaire
```

---

## 3. Squelette du module principal (.bas)

```vba
Attribute VB_Name = "MaCommandeV1"
Option Explicit

' --- Globals partages ---
Public g_oSettings  As CSettings
Public g_oCalc      As CCalcul
Public g_oMoteur    As CGraphique

'----------------------------------------------------------------------
Sub MaCommande()
    ' Verifier qu'un DGN est ouvert
    Dim oDgn As DesignFile
    On Error Resume Next
    Set oDgn = ActiveDesignFile
    On Error GoTo 0
    If oDgn Is Nothing Then
        MsgBox "Ouvrez d'abord un fichier DGN.", vbExclamation, "Ma Commande"
        Exit Sub
    End If

    ' Instanciation unique
    Set g_oSettings = New CSettings
    g_oSettings.Init

    Set g_oCalc = New CCalcul
    Set g_oMoteur = New CGraphique

    ' Formulaire modeless
    frmMaCommande.Initialiser g_oSettings
    frmMaCommande.Show vbModeless

    ' Demarrer la chaine de commande
    CommandState.StartPrimitive New CSelect1
End Sub
```

**Lancement :**
```
vba run [MaCommandeV1]MaCommande
```

---

## 4. Classes de commande (IPrimitiveCommandEvents)

Chaque etat de la commande est une classe implementant
`IPrimitiveCommandEvents`. Les cinq methodes a implementer :

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSelect1"
' ... autres attributs ...
Option Explicit
Implements IPrimitiveCommandEvents

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Ma Commande"
    ShowPrompt "Instruction pour l'utilisateur (Reset = quitter)"
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    ' Traitement du clic gauche
    ' ...
    ' Passer a l'etat suivant :
    CommandState.StartPrimitive New CSelect2
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    ' Reset = annuler / remonter d'un cran / quitter
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    ' Apercu dynamique (si active par CommandState.StartDynamics)
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
    ' Liberation des ressources
End Sub
```

### Chaine d'etats typique

```
CSelect1 --Data--> CSelect2 --Data--> CPlacer --Data--> (repete)
    |                  |                  |
  Reset              Reset              Reset
    |                  |                  |
  Quitter         CSelect1           CSelect1
```

### Apercu dynamique

Pour activer l'apercu dans un etat de placement :

```vba
Private Sub IPrimitiveCommandEvents_Start()
    ' ... prompts ...
    CommandState.StartDynamics     ' <-- indispensable
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, _
                                             ByVal DrawMode As MsdDrawingMode)
    ' Dessiner les elements temporaires avec DrawMode
    ' Ex: tracer une ligne, un cercle, un texte provisoire
End Sub
```

`CommandState.StartDynamics` doit etre appele dans le `Start` de la classe
de placement, pas avant. Sans cet appel, l'evenement `Dynamics` ne se
declenche jamais.

---

## 5. Classe donnees metier (sans API MST)

Pour les donnees contenant des coordonnees, les `Point3d` (UDT) ne peuvent
pas etre membres `Public` d'une classe. Contournement :

```vba
' CPointRef - stocke les coordonnees en Double
Option Explicit

Public X        As Double
Public Y        As Double
Public Z        As Double
Public Altitude As Double
Public Valide   As Boolean

Sub DefinirPosition(oPt As Point3d)
    X = oPt.X: Y = oPt.Y: Z = oPt.Z
End Sub

Sub CopierPosition(oPt As Point3d)
    oPt.X = X: oPt.Y = Y: oPt.Z = Z
End Sub
```

**Regle VBA :** un membre `Public` de type UDT (`Point3d`, `Matrix3d`...)
dans un module de classe provoque l'erreur *"Les types definis par
l'utilisateur ne sont pas autorises comme membre public d'un module
d'objet."* Stocker les composantes en `Double` et convertir via des
`Sub` avec parametre de sortie.

---

## 6. Classe de calcul pur

Ne depend d'aucune API MicroStation. Ne manipule que des nombres et des
chaines. Peut etre testee en isolation.

```vba
' CCalcul - moteur de calcul
Option Explicit

Function Calculer(dValA As Double, dValB As Double) As Double
    Calculer = (dValA + dValB) / 2#
End Function

Function FormatResultat(dVal As Double, nDec As Integer, sSep As String) As String
    FormatResultat = Format$(dVal, "0." & String$(nDec, "0"))
    If sSep = "," Then FormatResultat = Replace$(FormatResultat, ".", ",")
End Function
```

---

## 7. Classe graphique (API MST)

Recoit des resultats deja calcules et des parametres de symbologie.
Ne realise aucun calcul metier.

```vba
' CGraphique - creation des elements MicroStation
Option Explicit

Sub CreerElement(oProj As Point3d, sValeur As String, oSettings As CSettings)
    Dim oTexte As TextElement
    ' ...creation via CreateTextElement1...
    ActiveModelReference.AddElement oTexte

    Dim oCercle As EllipseElement
    ' ...creation...
    ActiveModelReference.AddElement oCercle
End Sub

Sub DessinDynamique(DrawMode As MsdDrawingMode, ...)
    ' Dessiner l'apercu temporaire (lignes, cercles provisoires)
    ' Utiliser .Redraw DrawMode sur chaque element temporaire
End Sub
```

---

## 8. Classe settings (agregateur de parametres)

```vba
' CSettings
Option Explicit

Public oTexte       As CSymboTexte
Public oCercle      As CSymboCercle
Public dTolerance   As Double
Public oModele      As TextElement      ' element modele (P1)

Sub Init()
    Set oTexte = New CSymboTexte
    Set oCercle = New CSymboCercle
    oCercle.InitDefauts
    dTolerance = 2#     ' rayon de recherche par defaut
End Sub
```

Les sous-objets de symbologie (`CSymboTexte`, `CSymboCercle`) sont de
simples classes conteneurs de proprietes, sans logique metier.

---

## 9. Formulaire modeless (.frm)

### 9.1 Principe : formulaire vide + controles au runtime

On ne dessine **rien** dans le designer VBA. Le `.frx` reste un blob
vide de ~2,5 Ko et ne change jamais. Tout le contenu visible est cree
par code dans `UserForm_Initialize`, ce qui rend le `.frm` entierement
versionnable en texte.

```vba
Private WithEvents txtDiametre As MSForms.TextBox

Private Sub UserForm_Initialize()
    ConstruireControles
End Sub

Private Sub ConstruireControles()
    Dim fra As MSForms.Frame
    Set fra = Me.Controls.Add("Forms.Frame.1", "fraCercle")
    fra.Caption = "Cercle"
    fra.Left = 6: fra.Top = 6: fra.Width = 192: fra.Height = 126

    Set txtDiametre = fra.Controls.Add("Forms.TextBox.1", "txtDiametre")
    txtDiametre.Left = 6: txtDiametre.Top = 24
    txtDiametre.Width = 48: txtDiametre.Height = 16
End Sub
```

### 9.2 ProgID des controles

| Controle | ProgID |
|---|---|
| Label | `Forms.Label.1` |
| TextBox | `Forms.TextBox.1` |
| ComboBox | `Forms.ComboBox.1` |
| CheckBox | `Forms.CheckBox.1` |
| OptionButton | `Forms.OptionButton.1` |
| CommandButton | `Forms.CommandButton.1` |
| Frame | `Forms.Frame.1` |
| ListBox | `Forms.ListBox.1` |

Les dimensions (`Left`, `Top`, `Width`, `Height`) sont en **points** (1 pt = 1/72 pouce).
Les controles enfants d'un Frame s'ajoutent via `leFrame.Controls.Add`.

### 9.3 Evenements : limitation des controles runtime

Les evenements *extender* (`Enter`, `Exit`, `BeforeUpdate`, `AfterUpdate`)
ne sont **pas disponibles** pour les controles crees au runtime. Strategie :

- **Validation continue sur `Change`** : si la saisie est valide,
  l'enregistrer dans `g_oSettings` ; sinon l'ignorer (garder la derniere
  valeur valide)
- **Entree (`KeyDown`, `KeyCode = vbKeyReturn`)** : reformater le champ
  depuis la valeur enregistree
- **Drapeau `m_bInit`** : empecher les gestionnaires de reagir pendant le
  pre-remplissage initial

```vba
Private m_bInit As Boolean

Sub Initialiser(oSettings As CSettings)
    m_bInit = True
    txtDiametre.Text = Format$(oSettings.oCercle.Diametre, "0.00")
    m_bInit = False
End Sub

Private Sub txtDiametre_Change()
    If m_bInit Then Exit Sub
    Dim d As Double
    If IsNumeric(txtDiametre.Text) Then
        d = CDbl(txtDiametre.Text)
        If d > 0 Then g_oSettings.oCercle.Diametre = d
    End If
End Sub
```

### 9.4 Fermeture par la croix

Un formulaire modeless reference par d'autres classes ne doit pas etre
decharge (sinon VBA recree une instance vierge => erreur 91) :

```vba
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1          ' ne pas decharger
        Me.Hide
        CommandState.StartDefaultCommand    ' terminer la commande
    End If
End Sub
```

### 9.5 Communication formulaire <-> commande

Le formulaire lit et ecrit `g_oSettings` directement. Les classes de
commande appellent des methodes publiques du formulaire pour mettre a jour
l'etat affiche :

```vba
' Dans CSelect1, apres selection de P1 :
frmMaCommande.AfficherP1 g_oP1

' Dans frmMaCommande :
Sub AfficherP1(oP1 As CPointRef)
    lblP1.Caption = "P1 : " & Format$(oP1.Altitude, "0.000")
End Sub
```

---

## 10. Generation du couple .frm/.frx

### Pourquoi c'est necessaire

Un `.frm` ne peut **pas** etre ecrit a la main. L'import VBA exige :
- le GUID designer MSForms `{C62A69F0-16DC-11CE-9E98-00AA00574A4F}`
- une ligne `OleObjectBlob` pointant vers un `.frx` valide

Seul un editeur VBA sait generer le `.frx`. Le script du projet utilise
**Excel** en COM pour le faire sans intervention manuelle.

### Utilisation du script

```powershell
.\scripts\export_frm_via_excel.ps1 `
    -SrcFrm  "src\frmMaCommande.frm" `
    -OutDir  "src" `
    -FormName "frmMaCommande"
```

Le script :
1. Active temporairement `AccessVBOM` dans le registre Excel (restaure a la fin)
2. Cree un classeur, ajoute un UserForm vide, injecte le code lu dans `SrcFrm`
3. Exporte le couple `<FormName>_export.frm` + `<FormName>_export.frx`

**Apres generation :** renommer le couple et corriger la ligne
`OleObjectBlob` du `.frm` pour qu'elle pointe vers le bon `.frx` :

```
OleObjectBlob   =   "frmMaCommande.frx":0000
```

### Anatomie d'un .frm valide

```
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMaCommande
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmMaCommande.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMaCommande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ... tout le code du formulaire ...
```

---

## 11. Encodage et git

### Regles imperatives

| Regle | Pourquoi |
|---|---|
| **ANSI sans BOM UTF-8** | Un BOM (`i»¿`) devant `Attribute VB_Name` casse le parsing |
| **Fins de ligne CRLF** | En LF, les `.cls` sont importes comme modules standard |
| **Pas d'accents dans le code** | Les sources doivent rester en ASCII pur |
| **`.frx` = binaire, ne jamais editer** | Toute conversion de fins de ligne corrompt le blob |

### .gitattributes obligatoire

```
*.bas text eol=crlf
*.cls text eol=crlf
*.frm text eol=crlf
*.frx binary
*.ps1 text eol=crlf
*.dgn binary
*.dgnlib binary
*.mvba binary
```

### Normalisation apres edition externe

```powershell
Get-ChildItem src\* -Include *.bas,*.cls,*.frm | ForEach-Object {
    $t = [IO.File]::ReadAllText($_.FullName)
    $t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($_.FullName, $t, [Text.Encoding]::ASCII)
}
```

---

## 12. Compatibilite MicroStation V8i

### API a utiliser

| Bonne pratique | A eviter |
|---|---|
| `CommandState.StartPrimitive` | `CommandState.StartPrimitiveCommand` (absent sur certaines installations) |
| `Sub` avec parametre de sortie (`ByRef`) | `Function` retournant `Point3d` / `Matrix3d` (refuse par certains environnements VBA) |
| `CommandState.StartDynamics` dans le `Start` | L'appeler ailleurs (Dynamics ne se declenchera pas) |
| `On Error Resume Next` autour des acces API optionnels | Supposer que toute propriete existe |

### Types UDT dans les classes

Un membre `Public` de type UDT (`Point3d`, `Matrix3d`...) dans un module
de classe provoque une erreur de compilation. Solutions :

- Stocker les composantes en `Double` (X, Y, Z) et convertir via des `Sub`
- Passer les `Point3d` en parametre des methodes (pas en membre)

### Scan d'elements

Utiliser `ElementScanCriteria` pour filtrer par type :

```vba
Dim oScan As New ElementScanCriteria
oScan.ExcludeAllTypes
oScan.IncludeType msdElementTypeText
oScan.IncludeType msdElementTypeCellHeader
oScan.IncludeType msdElementTypeTag

Dim oEnum As ElementEnumerator
Set oEnum = ActiveModelReference.Scan(oScan)
Do While oEnum.MoveNext
    Dim oElem As Element
    Set oElem = oEnum.Current
    ' ... traitement selon oElem.Type ...
Loop
```

### Niveaux geles

Filtrer les elements sur un niveau non affiche dans la vue active :

```vba
Function EstSurNiveauGele(oElem As Element) As Boolean
    On Error GoTo Erreur
    EstSurNiveauGele = Not oElem.Level.IsDisplayedInView(ActiveDesignFile.Views(1))
    Exit Function
Erreur:
    EstSurNiveauGele = False
End Function
```

---

## 13. Check-list nouveau projet

1. **Creer la structure de fichiers** (cf. section 2)
2. **Module principal** : point d'entree, globals, utilitaires de scan
3. **Classes metier** : donnees, calcul (sans API MST)
4. **Classes MST** : graphique, etats de commande (`IPrimitiveCommandEvents`)
5. **Formulaire** :
   - ecrire le code dans un `.frm` de travail (controles au runtime)
   - generer le couple `.frm/.frx` avec `scripts/export_frm_via_excel.ps1`
   - renommer + corriger `OleObjectBlob`
6. **Verifier l'encodage** : ANSI, CRLF, `.gitattributes` en place
7. **Importer dans MicroStation** : les `.frm` et `.frx` doivent etre dans
   le meme dossier au moment du `Ctrl+M`
8. **Compiler** : `Debogage > Compiler` dans l'editeur VBA
9. **Tester** : key-in `vba run [MonProjet]MaCommande`

---

## 14. Depannage express

| Symptome | Cause probable |
|---|---|
| « Impossible de charger ...frm » | `.frx` absent ou `OleObjectBlob` incorrect, ou GUID designer invalide (fichier ecrit a la main) |
| `.cls` importe comme module standard | Fins de ligne LF au lieu de CRLF |
| Caracteres `i»¿` en tete de fichier | BOM UTF-8 (reecrire en ASCII) |
| Evenements du controle ne se declenchent pas | Variable non declaree `WithEvents` ou `Set` oublie lors du `Controls.Add` |
| Erreur sur `Sub xxx_Exit(...)` | Evenement extender indisponible au runtime : utiliser `Change` + `KeyDown` |
| Erreur 91 apres fermeture de la fenetre | `QueryClose` sans `Cancel = 1` (formulaire decharge puis recree vierge) |
| UDT interdit dans une classe | Membre `Public` de type `Point3d` : stocker en `Double` + `Sub` de conversion |
| `Dynamics` ne s'appelle jamais | `CommandState.StartDynamics` absent du `Start` de la classe de placement |
| `StartPrimitiveCommand` introuvable | Utiliser `StartPrimitive` (compatible toutes installations) |
