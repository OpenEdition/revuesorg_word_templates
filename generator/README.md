# generator

Macro Word pour générer des versions linguisitiques des modèles à partir d'un modèle de base.

## Utilisation

1. Placer le dossier generator dans le répertoire des modèles de Word.
2. Editer les styles et les menus de `generator/src/base.dot`. Les noms utilisés pour les styles et les menus doivent correspondre aux identifiants déclarés dans le fichier `translations.ini`.
3. Editer les traduction dans `generator/src/translations.ini`.
4. Lancer la macro.

### Menu et macro de lancement : `base.dot`

Le modèle `generator/src/base.dot` doit contenir la macro d'application des styles (voir `base.dot.vb`) et le menu d'application des styles. Leurs noms doivent respecter les noms déclarés dans la macro `template_generator` soit par défaut :

```vba
    Const TOOLBARNAME As String = "LodelStyles"
    Const MACRONAME As String = "ApplyLodelStyle"
```

Les modifications de l'arborescence du menu ne posent pas de problème. L'action (ou le style) associée aux boutons des menus n'a aucune incidence dans la mesure où elle sera modifiée par la macro `template_generator`. On peut donc par exemple choisir d'associer la macro `ApplyLodelStyle` à tous les boutons.

Remarque : l'action associée aux boutons qui renvoient vers un lien hypertexte ne sera pas traitée. Cela permet par exemple d'ajouter un bouton qui renvoit vers les métadonnées. **TODO: Renvoyer vers une page dans la langue du modèle. Il faut pour cela ajouter informatiquement le lien et donc utiliser un attribut du type `[lang].link` dans `translations.ini`.**

### Traductions des styles et des menus : `translations.ini`

Le fichier `generator/src/translations.ini` contient les traductions des styles et des menus.

Dans la première section `[_configuration]`, la propriété `translateTo` contient la liste des code de langues de destination séparées par des virgules.

Les sections qui suivent se présentent ainsi :

```ini
    [identifiant_du_style_et_du_menu_dans_base.dot]
    fr.style="Nom du style en français"
    fr.menu="Nom du menu en français"
    en.style="Nom du style en anglais"
    en.menu="Nom du menu en anglais"
    ; etc
    style="Nom du style par défaut. C'est le nom qui sera utilisé si aucune traduction n'est trouvée dans la langue en cours. La cas échéant aucune erreur n'est produite dans log.txt. Cette option ne doit donc être complétée que pour les styles qui ne doivent pas être traduits."
    menu="Nom du menu par défaut. Idem."
    wordId="Identifiant numérique pour les styles natifs de Word, documenté ici: https://msdn.microsoft.com/en-us/library/bb237495%28v=office.12%29.aspx"
    key="Raccourci clavier" ; TODO: A documenter
```

Remarque : les sections des sous-menus doivent être préfixés par `menu_`. Exemple : `[menu_texte]`

## Todolist de développement

1. Copier `generator/src/base.dot` > `generator/tmp/styles.dot`.
2. Ouvrir `generator/tmp/styles.dot`.
3. Pour chaque style dans styles.dot :
    1. Traduire dans le nom (propriete `style` de la section INI) par défaut s'il existe.
    2. Chercher les traductions dans toutes les langues et ajouter un nouveau style pour chacune s'il n'existe pas déjà. Les styles créés de la sorte ou pour propriété `BaseStyle` le style d'origine.
4. Sauvegarder et fermer `generator/tmp/styles.dot`.
5. Pour chaque langue déclarée dans `generator/src/translations.ini` :
    1. Copier `generator/tmp/styles.dot` > `generator/build/revuesorg_[langue].dot`
    2. Ouvrir `generator/build/revuesorg_[langue].dot`
    3. ~~Copier la macro d'application des styles dans `generator/build/revuesorg_[langue].dot`~~ La copie de macro ne fonctionne pas à tous les coups (probablement à cause des autorisations dans Word). Il faut que cette macro soit initialement présente dans base.dot.
    4. Traduire la barre d'outil de `generator/build/revuesorg_[langue].dot` d'après le fichier `generator/src/translations.ini` et éditer les attributs `OnAction`, `Tag` et/ou `Parameter`.
    5. Assigner les raccourcis clavier (?).
    6. Sauvegarder et fermer `generator/build/revuesorg_[langue].dot`.
6. Afficher un message de fin de traitement.

Remarque pour intégrer le VBA dans Word : le code doit être inséré dans le document en tant que module : Alt+F11, clic droit sur le projet correspondant au document qui doit contenir la macro, Insertion > Module, coller le code. 
