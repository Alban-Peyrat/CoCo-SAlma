# CoCo-SAlma

CoCo-SAlma est un outil visant à comparer les cotes renseignées dans le Sudoc avec les cotes associées aux holdings dans Alma pour une liste de PPN donnée. Elle signale également si elle détecte plusieurs localisations dans le Sudoc ou Alma, ou si aucune cote n'a été trouvée.

À l'heure actuelle, peu d'utilisations pratiques ont été effectuées, n'hésitez pas à me partager des problèmes que vous auriez pu rencontrer en l'utilisant.

Les documents suivants se trouvent dans ce dépôt :

* CoCo-SAlma2.xlsm : l'outil, version du script du 27 septembre 2021 à 17h00
* vba_CoCo-SAlma2.vbs : le code VBA dans un fichier à part pour permettre son visionnage dans GitHub. __Inutile de le télécharger pour utiliser de CoCo-SAlma__
* Fctnmt_CoCo-SAlma2.docx / .pdf : explication sur le fonctionnement de CoCo-SAlma2
* Utiliser_CoCo-SAlma2.pptx / .pdf (du diaporama) / .docx : guide pour utiliser CoCo-SAlma2, la version texte n'est qu'un copier-coller du PowerPoint
* Ressources [dossier] qui contient des fichiers préfaits pour effectuer des tests (ces listes proviennent d'export Alma pour ma bibliothèque) :

   * List_PPN_Double_odonto.xlsx : liste originale des PPN pour les doublons de localisations en odontologie ;
   * export_alma_DOUBLON_SUDOC.txt  : le résultat de l'export Alma de la liste sus-mentionnée ;
   * export_alma_FONDS_ODONTO.xlsx : export de la requête Alma BIB=BUSVS && Holding=Odontologie Salle Ronde
   * export_alma_FONDS_MEDECINE.xlsx : export de la requête Alma BIB=BUSVS && Holding=Médecine Salle Ronde
   * List_PPN_Fds_His_Med_Bx.xlsx : liste originale des PPN du fonds Histoire de la Médecine de Bordeaux
