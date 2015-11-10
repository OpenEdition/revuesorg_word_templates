# revuesorg_word_templates : documentation

**Avertissement : cette page contient la documentation pour la création de modèles pour Revues.org. Pour de la documentation sur l'utilisation des modèles pour Revues.org, consultez [la Maison des revues](http://maisondesrevues.org).**

## Description générale

**revuesorg_word_templates** s'utilise en plusieurs étapes :

1. Modification des sources,
2. Génération des modèles à partir des sources modifiées,
3. Génération de l'installateur des modèles pour Windows.

Le projet contient un répertoire `src` qui où placer les sources utilisées. Tous les contenus produits (modèles et installateur) se trouvent dans le répertoire `build` après les traitements.

## Installation

### Pour la génération de modèles

Télécharger le zip du projet et le décompresser dans le répertoire des modèles de Word. 

**Attention :** l'ensemble des fichiers du projet doivent impérativement être contenus dans un répertoire `revuesorg_word_templates`. L'arborescence attendue est donc la suivante :

```
[user]\AppData\Roaming\Microsoft\Templates\
└── revuesorg_word_templates\
    ├── docs
    ├── src
    ├── (...)
    ├── template_generator.dot
    └── win_setup.iss
```

### Pour la compilation de l'installateur

La compilation de l'installateur requiert Inno Setup :

* Télécharger et installer la dernière version "Unicode" d'Inno Setup ici : http://www.jrsoftware.org/isdl.php
* Lors de l'installation, cocher l'option "Install Inno Setup Preprocessor".

## Modification des sources

[Page de documentation consacrée](template_generator.md)

Le générateur de modèles facilite la production à grande échelle de modèles pour Word : les styles, les menus, leurs traductions et les raccourcis clavier sont définis une seule fois dans des fichiers nommées "sources". Lors de l'exécution le générateur créée les modèles à partir des éléments donnés en source.

Toutes les sources sont contenues dans le dossier `src` :

```
src
├── base.dot
├── translations.ini
├── macros
│   ├── macros_revuesorg_mac.dot
│   └── macros_revuesorg_win.dot
└── startup
    └── revuesorg_startup.dot
```

Les sources suivantes sont utilisées par le générateur de modèles :

* `base.dot` : modèle Word qui sert de base à la création des modèles. Contient les menus et les styles a utiliser, nommés avec des identifiants du type `$nom`.
* `translations.ini` : fichier contenant la configuration générale ainsi qu'une configuration relative à chacun des styles.

Pour connaître les règles de modification de `base.dot` et `translations.ini`, consulter [la page de documentation consacrée](template_generator.md).

Les deux dossiers suivants contiennent des ressources utilisées par l'installateur :

* `macros` et `startup` : ces répertoires contiennent les modèles copiés (respectivement dans les dossiers `templates` et `startup` de Word) lors de l'installation.

## Utilisation du générateur de modèles

La macro de génération de modèles est compatible avec toutes les versions de Word depuis 2003 **sur Windows uniquement**.

Pour générer des modèles à partir des fichiers contenus dans `src` :

* Ouvrir Word.
* Dans la fenêtre "Modèles et compléments", ajouter et cocher le modèle `revuesorg_word_templates/template_generator.dot` en tant que "Modèles globaux et compléments".
* Le modèles ajoute deux boutons dans l'onget "Compléments" de Word. Le premier bouton "Générer les modèles traduits" lance la génération des modèles. 
* À la fin du traitement les modèles traduits se trouvent dans le répertoire `build/templates` et le journal des erreurs dans `build/log.txt`.

### Journal des erreurs

Le journal `build\log.txt` affiché à la fin du traitement contient les éventuelles erreurs rencontrées lors de la dernière exécution du générateur de modèles. Le plus souvent il s'agit d'une traduction manquante dans `translation.ini`. Le journal des erreurs permet donc de détecter les oublis de traductions ou les styles inexistants.

Lorsqu'une erreur est rencontrée, le style et/ou le menu associé est supprimé du modèle généré. Dans certains cas ce comportement peut permettre de réserver l'inclusion d'un style, d'un menu ou d'un bouton à une langue en particulier. Voir la [documentation sur `translations.ini`](template_generator.md) pour plus d'informations sur ce comportement.


## Créer un installateur à partir des modèles générés

Après la génération des modèles, il est possible de créer un exécutable qui installera automatiquement les modèles sur le poste de l'utilisateur. La compilation de l'installateur requiert Inno Setup (voir plus haut, partie "Installation").

Pour compiler l'installateur :

* Dans Word, cliquer sur le deuxième bouton de la barre d'outil "Ouvrir le répertoire de la macro" pour explorer le répertoire `revuesorg_word_templates`.
* Faire un clic droit sur le fichier `win_setup.iss` et sélectionner l'option "Compile".
* Lorsque la compilation est terminée, l'installateur est disponible dans le répertoire `build/win_setup`.

[Voir les informations de compatiblité de l'installateur pour Windows.](setup_win.md)

## En cas de problème

* Vérifier la [FAQ](faq.md).
* Sinon [c'est qu'il y a un bug...](https://github.com/OpenEdition/revuesorg_word_templates/issues)


