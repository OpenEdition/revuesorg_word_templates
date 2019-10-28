# revuesorg_word_templates

**revuesorg_word_templates** est une suite d'outils pour automatiser la traduction et l'installation des modèles pour Microsoft Word. Ce projet a été spécifiquement développé pour générer les modèles de stylage pour Revues.org.

## Avertissement 
Les modèles de documents produits avec cet outil sont compatibles avec Microsoft Word jusqu'à la version 2016 pour Windows et jusqu'à la version 2011 pour MacOS.

Pour les versions ultérieurs de Microsoft Word, se reporter au dépôt Gitub suivant : https://github.com/OpenEdition/templates.openedition

## Description

**revuesorg_word_templates** permet de :

* Générer automatiquement des modèles Word traduits à partir d'un fichier `base.dot` (qui contient les styles et les menus du modèle) et d'un fichier `translations.ini` (qui contient les traductions des styles et menus en chacune des langues de destination).
* Créer un exécutable qui automatise l'installation de ces modèles sur le poste utilisateurv (actuellement seul Windows est supporté).

## Documentation

La documentation complète se trouve dans [le répertoire `docs`](https://github.com/OpenEdition/revuesorg_word_templates/tree/master/docs).

## Développement

Les macros en VBA existent en double dans le dépot : une version `.bas` (code brut versionné) et une version `.dot` (binaire Word). Les deux versions sont (et doivent rester) identiques.

Dans la mesure du possible, il est recommandé de conserver les sources du générateur de modèle à jour sur ce dépôt.

## Licence

**revuesorg\_word\_templates**  
Copyright (C) 2015 Thomas Brouard (OpenEdition)

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
