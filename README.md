# Interpolation Topo — MicroStation V8i (VBA)

Macro VBA pour **MicroStation V8i SS3** qui interpole l'altitude d'un point situé sur la
droite reliant deux textes d'altitude existants (points topo), avec **aperçu dynamique**,
puis crée le texte altitude et un cercle au point choisi.

Pensée pour les levés topographiques : gisement en **gon** (0 = nord, sens horaire),
pente en **%**, séparateur décimal et nombre de décimales repris automatiquement des
textes source, cercles plaçables sur un niveau distinct des textes (pratique pour
imprimer les altitudes sans les symboles de points).

## Fonctionnalités

- Sélection de deux textes d'altitude par simple clic à proximité (les textes non
  numériques — matricules, etc. — sont ignorés)
- Le texte du point 1 sert de **modèle de mise en forme** : police, taille, niveau,
  couleur et rotation sont repris pour le texte créé
- **Aperçu dynamique** pendant le placement :
  - ligne support P1–P2 (étendue de 20 % de part et d'autre)
  - croix sur P1 et P2, cercle et texte provisoires suivant le curseur
  - barre d'état : altitude interpolée, abscisse `t` (0 = P1, 1 = P2), pente %, gisement gon
  - avertissement en cas d'extrapolation hors du segment
- Création **répétable** : plusieurs points sur la même droite (Data = créer, Reset = terminer)
- Paramètres au lancement : diamètre et couleur du cercle, niveau choisi dans la
  **liste des calques existants** du fichier

## Installation automatique (recommandée)

1. Télécharger le dépôt (**Code > Download ZIP**) et le décompresser
2. Double-cliquer sur **`install.cmd`**

Le script copie le projet VBA (`InterpolationTopo.mvba`) dans le workspace, installe
la boîte à outils (`MesMacros.dgnlib`) et configure le chargement automatique au
démarrage de MicroStation. Relançable sans risque. Au prochain démarrage de
MicroStation, le bouton et le key-in sont opérationnels.

> Le script suppose le workspace dans `Documents\MicroStV8i\WorkSpace`
> (sinon, modifier la variable `$Workspace` en tête de `install.ps1`).

## Installation manuelle

1. MicroStation : *Utilitaires > Macros > Gestionnaire de projets VBA* > **Nouveau projet**
   (MicroStation crée le fichier `.mvba`)
2. Éditeur VBA (`Alt+F11`) : *Fichier > Importer un fichier* (`Ctrl+M`) et importer les
   4 fichiers du dossier [`src/`](src) :

   | Fichier | Type après import |
   |---|---|
   | `InterpolationTopo.bas` | Module |
   | `CSelectP1.cls` | Module de classe |
   | `CSelectP2.cls` | Module de classe |
   | `CPlacerPoint.cls` | Module de classe |

3. *Débogage > Compiler*, puis enregistrer

> ⚠️ Les fichiers doivent conserver des fins de ligne **Windows (CRLF)** — le
> `.gitattributes` du dépôt s'en charge. Avec des fins de ligne Unix (LF), l'éditeur
> VBA importe les `.cls` comme modules standard et la compilation échoue.
> Éviter aussi le copier-coller du code (risque d'espaces insécables) : toujours
> passer par l'import de fichiers.

## Lancement

Key-in :

```
vba run [InterpolationTopo]InterpolerPoint
```

Le key-in peut être associé à une **touche de fonction** (*Utilitaires > Touches de
fonction*) ou à un **bouton de boîte à outils** (*Espace de travail > Personnaliser*).

📖 Guide détaillé pas à pas (installation, utilisation, touche de fonction, boîte à
outils, dépannage) : **[Mode_Emploi_InterpolationTopo.html](Mode_Emploi_InterpolationTopo.html)**
(à ouvrir dans un navigateur).

## Utilisation en bref

1. Lancer la commande → 3 dialogues (diamètre, couleur, niveau du cercle)
2. Cliquer près du texte altitude du **point 1** (modèle de mise en forme)
3. Cliquer près du texte altitude du **point 2** → pente et gisement s'affichent
4. Déplacer le curseur : la position est projetée sur la droite P1–P2, l'altitude
   interpolée s'affiche en dynamique
5. **Data** = créer le point (répétable) · **Reset** = nouvelle sélection de points
   (retour à l'étape 2, mêmes paramètres) · **Reset** au choix du point 1 = quitter

## Notes de compatibilité

Testé sur MicroStation V8i SS3 (08.11.09.578). Le code utilise volontairement :

- `CommandState.StartPrimitive` (certaines installations n'exposent pas
  `StartPrimitiveCommand`)
- des `Sub` avec paramètre de sortie plutôt que des `Function` retournant `Point3d` /
  `Matrix3d` (refusées par certains environnements VBA)
- `CommandState.StartDynamics` dans le `Start` de la commande de placement,
  indispensable pour que l'événement `Dynamics` (aperçu) se déclenche

## Licence

Utilisation et modification libres. Fourni tel quel, sans garantie.
