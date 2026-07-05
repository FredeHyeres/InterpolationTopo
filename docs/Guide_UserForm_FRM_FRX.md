# Guide : créer un UserForm VBA distribuable (.frm + .frx)

*Pour MicroStation V8i (VBA6/MSForms), valable pour tout hôte VBA. Rédigé
d'après l'expérience du formulaire `frmInterpolation` de la V2 (juillet 2026).*

---

## 1. Le problème

Un UserForm VBA est exporté par l'éditeur en **deux fichiers indissociables** :

| Fichier | Contenu |
|---|---|
| `.frm` | Texte : en-tête designer + attributs + tout le code VBA |
| `.frx` | Binaire : le « blob designer » (le dessin du formulaire et de ses contrôles) |

**On ne peut pas écrire un `.frm` à la main** et l'importer. Deux pièges vécus :

1. **Syntaxe VB6 ≠ VBA.** Déclarer les contrôles en texte dans le `.frm`
   (`Begin MSForms.Frame ... End`) est la syntaxe de Visual Basic 6, pas celle
   de VBA. Un UserForm VBA ne sérialise JAMAIS ses contrôles en texte : ils
   vivent dans le `.frx`.
2. **GUID et blob obligatoires.** L'import exige le GUID designer MSForms
   exact et une ligne `OleObjectBlob` pointant vers un `.frx` valide. Sinon :
   **« Impossible de charger &lt;fichier&gt;.frm »**.

Anatomie d'un `.frm` valide (généré, jamais écrit à la main) :

```
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMonForm
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmMonForm.frx":0000    <-- reference au binaire
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ... tout le code VBA du formulaire ...
```

Le GUID `{C62A69F0-16DC-11CE-9E98-00AA00574A4F}` est celui de la classe
UserForm MSForms — il ne s'invente pas.

---

## 2. La stratégie retenue : formulaire vide + contrôles au runtime

Plutôt que de dessiner les contrôles dans le designer (ce qui grossit le
`.frx`, illisible et non diffable dans git), on garde le **formulaire vide**
et on crée tout par code :

```vba
Private WithEvents txtDiametre As MSForms.TextBox   ' declaration module

Private Sub UserForm_Initialize()
    Dim fra As MSForms.Frame
    Set fra = Me.Controls.Add("Forms.Frame.1", "fraCercle")
    fra.Caption = "Cercle"
    fra.Left = 6: fra.Top = 6: fra.Width = 192: fra.Height = 126

    Set txtDiametre = fra.Controls.Add("Forms.TextBox.1", "txtDiametre")
    txtDiametre.Left = 6: txtDiametre.Top = 24
    txtDiametre.Width = 48: txtDiametre.Height = 16
End Sub
```

**Avantages :** le `.frx` reste minuscule (~2,5 Ko, formulaire vide) et ne
change plus jamais ; toute évolution du formulaire = édition du code texte du
`.frm`, versionnable et relisible.

### ProgID des contrôles courants

| Contrôle | ProgID |
|---|---|
| Label | `Forms.Label.1` |
| TextBox | `Forms.TextBox.1` |
| ComboBox | `Forms.ComboBox.1` |
| CheckBox | `Forms.CheckBox.1` |
| OptionButton | `Forms.OptionButton.1` |
| CommandButton | `Forms.CommandButton.1` |
| Frame | `Forms.Frame.1` |
| ListBox | `Forms.ListBox.1` |

Les dimensions (`Left`, `Top`, `Width`, `Height`) sont en **points**.
Les contrôles enfants d'un Frame s'ajoutent via `leFrame.Controls.Add`.

### Événements : la limitation à connaître

Pour recevoir les événements d'un contrôle créé au runtime, il faut une
variable de module déclarée `WithEvents` (voir ci-dessus). Les procédures
`txtDiametre_Change()` etc. fonctionnent alors normalement, **MAIS** les
événements *extender* (fournis par le conteneur au design) ne sont **pas
disponibles** : `Enter`, `Exit`, `BeforeUpdate`, `AfterUpdate`.

Conséquence pratique : pas de validation « à la sortie du champ ». À la place :

- valider en continu sur `Change` : si la saisie est valide, l'enregistrer ;
  sinon l'ignorer (on garde la dernière valeur valide) ;
- offrir la touche **Entrée** (`KeyDown`, `KeyCode = vbKeyReturn`) pour
  reformater l'affichage depuis la valeur enregistrée ;
- protéger les gestionnaires avec un drapeau `m_bInit` pendant le
  pré-remplissage initial.

### Autre garde-fou : la fermeture par la croix

Un formulaire modeless référencé par d'autres classes (via son instance par
défaut) ne doit pas être déchargé par l'utilisateur, sinon VBA en recrée une
instance vierge au prochain accès (variables à `Nothing` → erreur 91) :

```vba
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1          ' ne pas decharger
        Me.Hide
        ' ... terminer proprement la commande ici ...
    End If
End Sub
```

---

## 3. Générer le couple .frm/.frx : `scripts/export_frm_via_excel.ps1`

Seul un éditeur VBA sait produire un `.frx` valide. Le script du projet pilote
**Excel** en COM pour le faire sans intervention manuelle :

1. active temporairement « Accès approuvé au modèle d'objet VBA »
   (`AccessVBOM = 1` dans `HKCU:\Software\Microsoft\Office\<ver>\Excel\Security`,
   **restauré à la fin**) ;
2. crée un classeur, ajoute un composant UserForm (`VBComponents.Add(3)`),
   le nomme ;
3. injecte le code lu depuis votre fichier source (les lignes `Attribute`
   sont filtrées : illégales dans un module de code, régénérées à l'export) ;
4. exporte : Excel écrit le `.frm` (avec GUID et `OleObjectBlob` corrects)
   et le `.frx`.

**Usage pour un nouveau formulaire :**

```powershell
.\scripts\export_frm_via_excel.ps1 `
    -SrcFrm  "src\v3\frmMonNouveauForm.frm" `
    -OutDir  "src\v3" `
    -FormName "frmMonNouveauForm"
```

`SrcFrm` = votre fichier de travail contenant le code (l'en-tête designer y
est ignoré, seul le code après `Attribute VB_Exposed` est lu). Le script
produit `<FormName>_export.frm/.frx` dans `OutDir` : renommer ensuite le
couple, **en corrigeant la ligne `OleObjectBlob` du `.frm`** pour qu'elle
pointe vers le nouveau nom du `.frx` (c'est du texte, éditable).

Point de compatibilité : Excel 2007 (VBA6, 32 bits, MSForms 2.0) = même
environnement que MicroStation V8i, le couple exporté s'importe sans souci.

Le code n'a **pas besoin de compiler dans Excel** (types MicroStation
inconnus) : l'export n'exige aucune compilation.

---

## 4. Règles d'encodage (valables pour TOUT fichier VBA du projet)

L'import VBA (Ctrl+M) attend :

- **ANSI sans BOM UTF-8** : un BOM (`ï»¿`) devant `Attribute VB_Name` casse le
  parsing des attributs. Écrire les sources en ASCII pur — pas d'accents dans
  le code ni les commentaires ;
- **CRLF** : en LF, les `.cls` sont importés comme modules standard au lieu de
  modules de classe ;
- normalisation après toute édition externe :

```powershell
Get-ChildItem src\v2\* -Include *.bas,*.cls,*.frm | ForEach-Object {
    $t = [IO.File]::ReadAllText($_.FullName)
    $t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($_.FullName, $t, [Text.Encoding]::ASCII)
}
```

Côté git, le `.gitattributes` du projet impose déjà :

```
*.bas text eol=crlf
*.cls text eol=crlf
*.frm text eol=crlf
*.frx binary        <-- CRITIQUE : une conversion de fins de ligne corrompt le blob
```

---

## 5. Check-list nouveau formulaire

1. Écrire le code du formulaire dans un `.frm` de travail (en-tête quelconque,
   tout le contenu visible construit au runtime, cf. §2).
2. Générer le couple avec `scripts/export_frm_via_excel.ps1` (§3).
3. Renommer le couple + corriger `OleObjectBlob`.
4. Vérifier l'encodage (§4) — le `.frx` ne se touche JAMAIS.
5. Importer les DEUX fichiers ensemble (le `.frx` doit être dans le même
   dossier que le `.frm` au moment du Ctrl+M).
6. Pour toute évolution ultérieure : éditer la partie code du `.frm` au-dessous
   des lignes `Attribute` — le `.frx` reste valide tant que le formulaire
   designer reste vide.

## 6. Dépannage express

| Symptôme | Cause probable |
|---|---|
| « Impossible de charger ...frm » | `.frx` absent/renommé sans corriger `OleObjectBlob`, ou GUID designer invalide (fichier écrit à la main) |
| `.cls`/`.frm` importé comme module standard, en-tête visible en commentaire | Fins de ligne LF |
| Caractères `ï»¿` ou erreur sur la 1re ligne | BOM UTF-8 |
| Les événements du contrôle ne se déclenchent pas | Variable non déclarée `WithEvents`, ou `Set` oublié lors du `Controls.Add` |
| Erreur de compilation sur `Sub xxx_Exit(...)` | Événement extender indisponible au runtime : utiliser `Change` + `KeyDown` |
| Erreur 91 après fermeture de la fenêtre | `QueryClose` sans `Cancel = 1` (formulaire déchargé puis recréé vierge) |
| « ...types définis par l'utilisateur ne sont pas autorisés comme membre public de module d'objet » | Membre `Public` de type UDT (`Point3d`...) dans une classe : stocker en `Double` + Subs de conversion |
