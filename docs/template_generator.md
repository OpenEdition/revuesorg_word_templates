# Générateur de templates :  `base.dot` et `translations.ini`

Lors de l'exécution, la macro de génération des modèles copie un modèle de base (nommé `base.dot`) et produit une version traduite pour chaque langue déclarée dans le fichier de configuration (nommé `translations.ini`). Plus précisément, les styles et les menus du modèle sont traduits d'après les traductions contenues dans `translations.ini`.

Cette page de documentation détaille les règles de modification de `base.dot` et `translations.ini` qui se trouvent dans le répertoire `src` du projet.

## Préparation de `base.dot`

`base.dot` contient l'intégralité des styles et des menus/boutons (modèle complet compris).

Les opérations suivantes peuvent être réalisées dans `src/base.dot` :

* ajouter ou supprimer des styles,
* modifier l'apparence des styles,
* ajouter, supprimer ou réorganiser des menus ou des boutons,
* modifier les icônes des menus ou des boutons.

Les opérations concernant le nom des éléments sont réalisées dans `translations.ini` (voir partie consacrée plus bas).

Lors de la création d'un nouveau bouton dans un menu, le bouton créé doit être un lien vers un style (peu importe lequel).

**Remarque :** il est recommandé d'utiliser Microsoft Word 2003 pour éditer les barres d'outils. 

### Identifiants

Dans `base.dot`, **tous les styles, menus et boutons à traduire doivent avoir pour nom un identifiant** respectant la syntaxe suivante :

* Tous les identifiants des boutons et des styles doivent commencer par le caractère `$`.
* Les identifiants des menus et sous-menus doivent commencer par la chaîne `$menu_`.
* Les identifiants ne doivent contenir que des caractère alphabétiques non accentués, les caractères `_` et `$`. Pas d'espaces, de capitales, d'accents ou de ponctuation.
* Les identifiants utilisés dans `base.dot` doivent correspondre aux identifiants déclarés dans le fichier `translations.ini`.
* Lorsqu'un bouton permet l'application d'un style, son identifiant doit être strictement identique à celui du style. Par exemple pour le style du résumé en français il faut un bouton `$resumefr` ainsi qu'un style `$resumefr`.

### Condition de fonctionnement de `base.dot`

Le menu des styles de `base.dot` doit impérativement avoir pour nom : `LodelStyles`.

## Préparation de `translations.ini`

Le fichier `src/translations.ini` contient la configuration de la macro de génération des modèles ainsi que les traductions des styles et des menus.

Il s'édite avec n'importe quel éditeur de texte, en respectant [la syntaxe des fichiers INI](https://fr.wikipedia.org/wiki/Fichier_INI "Fichier INI sur Wikipedia"). Il doit être encodé en `utf-8`.

### Section `[_configuration]`

`translations.ini` s'ouvre avec la section `[_configuration]` qui contient la configuration utilisée lors de l'exécution de la macro.

#### Clé `translateTo`

La clé `translateTo` contient [les identifiants 639-1](https://fr.wikipedia.org/wiki/Liste_des_codes_ISO_639-1 "Codes ISO 639-1 sur Wikipedia") des langues de destination, séparés par des virgules.

Exemple :

```ini
translateTo="fr, en, es, pt"
```

Seules les langues référencées seront traitées par la macro.

#### Clé `version`

La clé `version` contient la version du modèle qui sera insérée un tant que corps de texte de chaque modèle. Forme: `[majeur].[mineur].[correctif]`. Pour des question de suivi des révisions, la version des modèles doit être incrémentée à chaque modification de `base.dot` ou de `translations.ini`.

Exemple :

```ini
version="4.0.2"
```

Cette clé est également reprise par l'installateur lors de sa compilation.

### Autres sections

Les autres sections correspondent chacune à un identifiant présent dans `base.dot` (voir plus haut). Les clés de la section déterminent le comportement pour l'objet (menu, style et/ou bouton) qui porte l'identifiant.

Par convention, les identifiants des menus et sous-menus sont composés de la façon suivante : `[$menu_nomdumenu]`. Pour le menu complet : `[$menu_completnomdumenu]`.

#### Clés de traduction : `[lang].menu` et `[lang].style`

La clé `[lang].menu` contient la traduction dans la langue `[lang]` du bouton associé.

La clé `[lang].style` contient la traduction dans la langue `[lang]` du style associé. Cette clé n'est pas utilisée pour les menus et sous-menus.

Exemple :

```ini
[$personnescitees]
fr.menu="Personnes citées"
fr.style="personnescitees"
```

**Remarque concernant les styles natifs.** Le nom des styles natifs de Word est automatiquement pris en charge par Word selon la langue d'installation d'Office. Dans le cas de styles natifs il est donc indispensable de renseigner une valeur `[lang].style` qui corresponde exactement au nom du style natif dans Word dans la langue cible.

#### Clés de traduction par défaut : `menu` et `style`

Les clés `menu` et `style` permettent d'attribuer des traductions par défaut qui seront appliquées si aucune traduction n'est trouvée dans la langue traitée.

**Attention : l'utilisation d'une traduction par défaut ne provoque pas d'enregistrement dans le journal `log.txt`, il est donc recommandé de limiter son utilisation aux éléments qui ne sont jamais traduits (index linguistiques par exemple) afin de conserver un journal pertinent.**

Exemple :

```ini
[$titreen]
style="Title (en)"
menu="Title (en)"
```

Comme la clé `[lang].style`, la clé `style` n'a aucun effet pour les sections qui ne sont pas associées à un style (menus, sous-menus).

#### Clé d'assignation de raccourci clavier : `key` et `[lang].key`

La clé `key` détermine le raccourci clavier du style associé.

Exemple :

```ini
key="Ctrl+Alt+A"
```

Comme avec les clés de traduction, il est possible d'appliquer des raccourcis clavier différents selon la langue.

Exemple :

```ini
[$periode]
fr.style="Periode"
fr.menu="Période"
fr.key="Ctrl+Alt+P"
en.style="Chronology"
en.menu="Chronological index"
en.key="Ctrl+Alt+C"
```

Les touches acceptées sont : `0`, `1`, `2`, `3`, `4`, `5`, `6`, `7`, `8`, `9`, `A`, `B`, `BackSingleQuote`, `BackSlash`, `Backspace`, `C`, `CloseSquareBrace`, `Comma`, `Command`, `D`, `Delete`, `E`, `End`, `Equals`, `Esc`, `F`, `F1`, `F10`, `F11`, `F12`, `F13`, `F14`, `F15`, `F16`, `F2`, `F3`, `F4`, `F5`, `F6`, `F7`, `F8`, `F9`, `G`, `H`, `Home`, `Hyphen`, `I`, `Insert`, `J`, `K`, `L`, `M`, `N`, `Numeric0`, `Numeric1`, `Numeric2`, `Numeric3`, `Numeric4`, `Numeric5`, `Numeric5Special`, `Numeric6`, `Numeric7`, `Numeric8`, `Numeric9`, `NumericAdd`, `NumericDecimal`, `NumericDivide`, `NumericMultiply`, `NumericSubtract`, `O`, `OpenSquareBrace`, `Option`, `P`, `PageDown`, `PageUp`, `Pause`, `Period`, `Q`, `R`, `Return`, `S`, `ScrollLock`, `SemiColon`, `SingleQuote`, `Slash`, `Spacebar`, `T`, `Tab`, `U`, `V`, `W`, `X`, `Y`, `Z`

Elle peuvent être combinées aux modificateurs : `Alt`, `Control` (ou `Ctrl`), `Shift` (ou `Maj`). Le séparateur est le signe `+` (sans espace).

**Remarque :** Word se réserve l'utilisation de certaines combinaisons de touches. Le cas échéant, l'option `key` n'est pas appliquée. Il faut alors changer de raccourci clavier.

#### Traitement des styles natifs de Word avec la clé `builtIn`

Les styles nativement pris en charge par Word ne sont pas traduits par la macro. Afin d'éviter les erreurs, il convient de les distinguer dans `translations.ini` en leur ajoutant la clé `builtIn` avec la valeur `true`.

Exemple :

```ini
[$titre]
builtIn="true"
fr.menu="Titre"
fr.style="Titre"
en.menu="Title"
en.style="Title"
es.menu="Título"
es.style="Título"
pt.menu="Título"
pt.style="Título"
key="Ctrl+T"
```

#### Lien hypertexte : clé `link` (et `[lang].link`)

La clé `link` permet d'attribuer un redirection hypertexte à un bouton. Comme avec `style`, `menu` et `key`, il est possible de contextualiser la valeur à la langue du modèle.

Voici par exemple un bouton qui ouvre le navigateur de l'utilisateur sur une page web différente selon la langue :

```ini
[$ordremetadonnees]
fr.menu="Ordre des métadonnées"
fr.hyperlink="http://maisondesrevues.org/108"
en.menu="Metadata's order"
en.hyperlink="http://maisondesrevues.org/404"
```

#### Distinction du modèle complet : la clé `complet`

La clé `complet` permet, lorsque sa valeur est égale à `true`, de limiter l'insertion de l'élément associé au modèle complet. Attention, cette clé doit être appliquée à tous les types d'éléments concernés : menu et styles.

Exemple :

```ini
[$resumeja]
style="resumeja"
menu="レジュメー (ja)"
complet="true"
```

### Conditionner l'insertion d'un élément à la langue

Lorsqu'une erreur est rencontrée, le style et/ou le menu associé est supprimé du modèle généré. Dans certains cas ce comportement peut permettre de réserver l'inclusion d'un style, d'un menu ou d'un bouton à une langue en particulier. Par exemple ici on n'affiche pas le bouton `$ordremetadonnees` dans les modèles autres que français, ce bouton n'ayant pas d'utilité dans les autres langue.

```ini
[$ordremetadonnees]
fr.menu="Ordre des métadonnées"
fr.hyperlink="http://maisondesrevues.org/108"
```

## Exemple d'utilisation

Exemple d'insertion d'un style "Accroche" dans le modèle complet.

Dans `base.dot` :

* Créer un style nommé `$accroche` dans `base.dot` et définir ses attributs (police, couleur, taille...).
* Créer un menu nommé `$accroche` dans le menu "Texte" (complet).

Dans `translations.ini` :

* Incrémenter la version dans la clé `[_configuration]`.
* Ajouter la section `$accroche` :

```ini
[$accroche]
fr.menu="Accroche"
fr.style="Accroche"
en.menu="Introductory text"
en.style="IntroText"
es.menu="Postítulo"
es.style="Postitulo"
pt.menu="Apresentação sumária"
pt.style="ApresentacaoSumaria"
complet="true"
```

* Enregistrer ensuite les deux fichiers puis lancer la macro de génération des modèles, etc.
* Enfin, penser à commiter les nouvelles sources (`base.dot` et `translations.ini`) afin de centraliser les modifications.
