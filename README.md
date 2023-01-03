# CoCo-SAlma - Contrôle des correspondances des cotes Sudoc-Alma

[![Abandonned](https://img.shields.io/badge/Maintenance%20Level-Abandoned-orange.svg)](https://gist.github.com/cheerfulstoic/d107229326a01ff0f333a1d3476e068d)

Coco-Salma est un outil visant spécifiquement à contrôler la correspondance des cotes pour un PPN entre les données d'exemplaire du Sudoc en utilisant [le webservice MARCXML de l'Abes](http://documentation.abes.fr/sudoc/manuels/administration/aidewebservices/index.html#SudocMarcXML) et les holdings d'Alma. N'étant personnellement pas fan de posséder un document dans plusieurs localisations, elle signale également la présence de plusieurs cotes.

**Évitez d'avoir d'autres fichiers Excel ouverts pendant l'analyse (dans le cas où une erreur de programmation pourrait faire intéragir Constance avec des fichiers non prévus).**

_Version du 14/10/2021. En cours de réécriture._

## Initialisation

Enregistrez le document `CoCo-SAlma2.xlsm` (Coco-Salma dans le reste du document) dans un dossier de travail. __Attention, Coco-SAlma créera un dossier à l'emplacement où vous sauvegarderez ce fichier.__

Allumez Coco-Salma, remettez à zéro les données grâce bouton dans la feuille `Introduction` puis rendez-vous dans la feuille `Données` :
* choisissez un nom pour le fonds en `G2` (le dossier créé portera le nom `CoCo-SAlma_NomDuFonds`) ;
* entrez en `I2` le RCR de votre bibliothèque (manuellement ou en appuyant sur Alt + flèche du bas pour afficher une liste déroulante), ce qui remplira automatiquement le nom de la bibliothèque.

### Ajouter une bibliothèque

La liste des bibliothèques se trouvent sur la feuille `Introduction` en colonnes `S:T`. Pour le nom de la bibliothèque dans Alma, il doit correspondre exactement à son intitulé dans Alma. La liste déroulante des RCR se mettra automatiquement à jour après un ajout (jusqu'à 297 bibliothèques. Au-delà, il faudra modifier la formule).

À noter : la sélection de bibliothèque se fait via le RCR mais une modification pour se faire via l'intitulé de la bibliothèque est tout à fait possible avec peu de changements.

## Renseigner la liste des PPN

### À partir d'une liste de PPN d'Alma

Exportez d'Alma une liste de Titres physiques, renommez-la `export_alma.xlsx` et placez-la dans le même dossier que Coco-Salma. Dans Coco-Salma, purifiez la liste de PPN via le bouton de la feuille `Introduction`. __Veillez à ne pas ouvrir `export_alma.xlsx` avant de lancer la purification.__

### À partir d'une liste de PPN déjà établie

Collez votre liste de PPN dans la `Liste de PPN originale`. Coco-Salma prend en compte les 9 derniers caractères de la cellule, si votre liste se présente sous la forme `PPN 123456789` ou `(PPN)123456789` ce n'est pas la peine de la retoucher, ni de rajouter des 0 en début de PPN, elle les ajoutera automatiquement.

## Générer la liste pour l'export et importer les données

Une fois la liste de PPN renseignée, générez la liste des PPN pour l'export via le bouton de la feuille `Introduction`. Cette étape créera le dossier, qui contiendra deux fichiers :
* `Fichier_ppal.xlsm`, qui est la version de travail de Coco-Salma pour le traitement de ce fonds ;
* `import_alma.xlsx`, qui est le document à utiliser pour exporter la liste des cotes d'Alma.

Note : plus la liste est longue, plus cela peut prendre du temps.

### Cas n° 1 : vous avez importé la liste de PPN originale depuis Alma

Dans ce cas, déplacez le fichier `export_alma.xlsx` dans le dossier `CoCo-SAlma_NomDuFonds`. Vous pouvez également supprimer le fichier `import_alma.xlsx`.

### Cas n° 2 : votre liste de PPN originale ne vient pas d'Alma

Dans ce cas, il va falloir générer l'export. Connectez-vous à Alma, puis rendez-vous dans `Admin` puis `Gérer les jeux`. Sélectionnez `Ajouter un jeu` puis `exemplarisé`. Remplissez alors le formulaire en :
* nommant le jeu ;
* ajoutant un description si vous le souhaitez ;
* choisissant `Titres physiques` comme type de contenu ;
* sélectionnant `Privé` ;
* sélectionnant `À partir d'un fichier`.

Cliquez ensuite sur l'icône de dossier et sélectionner le fichier `import_alma.xlsx`. Finissez cette étape en cliquant sur `Enregistrer`.

Attendez ensuite que le jeu soit importé, puis faites un clic droit sur son nom et sélectionnez `Membres`. Exportez enfin la liste de résultats avec `Excel (tous les champs)`, puis enregistrez le fichier sous le nom `export_alma.xlsx` et placez-le dans le dossier `CoCo-SAlma_NomDuFonds`.

## Lancement de l'analyse

À ce stade, votre dossier doit contenir trois (parfois deux pour le cas n° 1) fichiers :
* `export_alma.xlsx` ;
* `Fichier_ppal.xlsm` ;
* `import_alma.xlsx` (si vous ne l'avez pas supprimé. À ce stade, il ne sera plus utilisé).

Si vous possédez bien les deux premiers, lancez l'analyse via le bouton de la feuille `Introduction`.

Celle-ci peut prendre du temps selon la taille de la liste (par exemple, 18:17 pour 5507 titres). Durant l'analyse, évitez de toucher aux fichiers du dossier `CoCo-SAlma_NomDuFonds`, au cas où.

Une fois qu'elle est terminée, une pop-up vous indiquera que l'analyse est terminée, vous indiquant par ailleurs le temps de traitement.

## Fonctionnement

_En cours d'écriture_

## Résultats

À ce stade vous pouvez si vous le souhaitez exporter les résultats dans un autre classeur (clic droit sur la feuille `Résultats`, `Déplacer ou copier`, `Nouveau classeur`, cocher `Créer une copie`, `OK`).

Vous pouvez également filtrer les données pour en faciliter la lecture. Par ailleurs, un code couleur est utilisé pour la liste des résultats :
* colonne `Correspondance ?` :
  * bleu : un problème a été détecté dans Alma (prioritaire sur le Sudoc) ;
  * rouge : un problème a été détecté dans le Sudoc ;
  * orange/jaune : un problème a été détecté, sans plus de précision.
* colonne `Cotes` (Alma et Sudoc) :
  * rouge : absence de cote dans ce logiciel ;
  * vert : plusieurs cotes détectées.

### Quelques précisions sur les résultats

* Si plusieurs cotes sont détectées dans le Sudoc ou dans Alma pour votre RCR / bibliothèque, `Correspondance ?` affichera `Double localisation` et les cotes dans les colonnes en question seront séparées par la chaîne de caractères `;_;`.
* S'il n'y a pas correspondance ou que le PPN dans le Sudoc n'a pas son équivalent dans Alma (ou inversement), `Correspondance ?` affichera `Non` et sera coloré en orange (ou bleu ou rouge si l’erreur est précisée).
* Par ailleurs, il peut y avoir de bonnes correspondances tout de même signalées si plusieurs localisations sont utilisées.
* Attention, il est possible que certaines cotes soient des "faux positifs", il n'y aura pas de problème mais la manière de comparer aura échoué (ex : il y a un espace de trop, une minuscule à la place d’une majuscule, etc.).
* Enfin, Coco-Salma désigne comme `PPN incorrect` toute entrée pour laquelle elle n’a pas pu récupérer le PPN via [le webservice MARCXML de l'Abes](http://documentation.abes.fr/sudoc/manuels/administration/aidewebservices/index.html#SudocMarcXML). Cela ne veut pas dire que le PPN est forcément incorrect. Si vous constatez un nombre anormal de PPN incorrects, relancez une analyse sur cette liste et/ou sur la liste des résultats `Pas de cote Sudoc` pour lesquels aucun PPN n’est associé dans la colonne Sudoc. (Plus d’informations à ce sujet dans [la partie sur le fonctionnement](https://github.com/Alban-Peyrat/CoCo-SAlma#fonctionnement))

## Nettoyer une fois achevé

Une fois le traitement fini ou la feuille de résultats exportée :
* supprimez si vous le souhaitez les fichiers inutiles (`export_alma.xlsx`, `import_alma.xlsx` ou tout le dossier `Coco-Salma_NomDuFonds` si vous avez exporté les résultats) ;
*	supprimez le jeu dans Alma si vous l’avez créé et que vous ne comptez pas le réutiliser.
