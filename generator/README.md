# generator

Macro Word pour générer des versions linguisitiques des modèles à partir d'un modèle de base.

## Utilisation

1. Placer le dossier generator dans le répertoire des modèles de Word.
2. Editer les styles et les menus de generator/src/base.dot. Les noms utilisés pour les styles et les menus doivent correspondre aux identifiants déclarés dans le fichier translations.ini.
3. Editer les traduction dans generator/src/translations.ini. Attention : ce fichier doit impérativement être encodé en ANSI ou UTF-16-LE.
4. Lancer la macro.

## Todolist de développement

1. Copier generator/src/base.dot > generator/tmp/styles.dot.
2. Ouvrir generator/tmp/styles.dot.
3. Pour chaque style dans styles.dot :
    1. Traduire dans le nom (propriete `style` de la section INI) par défaut s'il existe.
    2. Chercher les traductions dans toutes les langues et ajouter un nouveau style pour chacune s'il n'existe pas déjà. Les styles créés de la sorte ou pour propriété `BaseStyle` le style d'origine.
4. Sauvegarder et fermer generator/tmp/styles.dot.
5. Pour chaque langue déclarée dans generator/src/translations.ini :
    1. Copier generator/tmp/styles.dot > generator/build/revuesorg_[langue].dot
    2. Ouvrir generator/build/revuesorg_[langue].dot
    3. ~~Copier la macro d'application des styles dans generator/build/revuesorg_[langue].dot~~ La copie de macro ne fonctionne pas à tous les coups (probablement à cause des autorisations dans Word). Il faut que cette macro soit initialement présente dans base.dot.
    4. Traduire la barre d'outil de generator/build/revuesorg_[langue].dot d'après le fichier generator/src/translations.ini et éditer les attributs `OnAction`, `Tag` et/ou `Parameter`.
    5. Assigner les raccourcis clavier (?).
    6. Sauvegarder et fermer generator/build/revuesorg_[langue].dot.
6. Afficher un message de fin de traitement.

TODO: générer un log des traductions opérées par la macro.
