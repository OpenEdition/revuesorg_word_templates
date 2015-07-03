# generator

Macro Word pour générer des versions linguistiques des modèles à partir d'un modèle de base.

## Description

Lors de son exécution, `generator` copie un modèle de base (nommé `base.dot`) et produit une version traduite pour chaque langue déclarée dans le fichier de configuration (nommé `translations.ini`). Plus précisément, les styles et les menus du modèle sont traduits d'après les traductions contenues dans `translations.ini`.

## Compatibilité

`generator` est compatible avec toutes les versions de Word depuis 2003 **sur Windows uniquement**.

Les modèles produits par `generator` fonctionnent quelle que soit la langue d'installation de Word. Par exemple, il est possible d'utiliser un modèle anglophone avec une version française de Word. Toutefois cela requiert VBA, les modèles générés ne sont donc pas compatibles avec Word 2008 (Mac) qui ne supporte pas VBA.

Selon les partis pris durant la génération des modèles il est également possible dans certains cas de rencontrer lors de l'application des styles natifs dans Word 2003 (lire plus bas la documentation concernant l'utilisation du paramètre `wordId`).

## Utilisation

## Installation 

Désarchiver `generator` dans le répertoire des modèles de Word afin d'obtenir l'arborescence suivante : 

	[user]\AppData\Roaming\Microsoft\Templates\
		|- generator\
			|- build
			|- src
			|- tmp
			|- utils
			|- generator.dot

**Remarque :** Pour Word 2003, le dossier n'est pas "Templates" mais "Modèles".

Le sous-dossier `src` contient les éléments modifiables avant le lancement du `generator`. 

Le sous-dossier `build` contient les élément produits après l'exécution : les modèles traduits et `log.txt`.

## Préparation de `base.dot`

`base.dot` contient l'intégralité des styles et des menus/boutons (modèle complet compris). 

Les opérations suivantes peuvent être réalisées dans `generator/src/base.dot` :

* ajouter ou supprimer des styles 
* modifier l'apparence des styles
* ajouter, supprimer ou réorganiser des menus ou des boutons
* modifier les icônes des menus ou des boutons

Les opérations concernant le nom des éléments sont réalisées dans `translations.ini` (voir partie consacrée plus bas). 

De même l'action (ou le style) associée aux boutons des menus n'a aucune importance dans la mesure où elle sera modifiée lors du traitement à partir des informations contenues dans `translations.ini`. On peut donc par exemple choisir d'associer la macro `ApplyLodelStyle` à tous les boutons.

### Identifiants

Dans `base.dot`, **tous les styles, menus et boutons à traduire doivent avoir pour nom un identifiant** respectant la syntaxe suivante :

* Tous les identifiants des boutons et des styles commencent par le caractère `$`.
* Les identifiants des menus et sous-menus commencent par la chaîne `$menu_`.
* Les identifiants ne contiennent que des caractère alphabétiques non accentués, les caractères `_` et `$`. Pas d'espaces, de capitales, d'accents ou de ponctuation.
* Les identifiants utilisés dans `base.dot` doivent correspondre aux identifiants déclarés dans le fichier `translations.ini`.
* Lorsqu'un bouton permet l'application d'un style, son identifiant doit être strictement identique à celui du style. Par exemple pour le style du résumé en français il faut un bouton `$resumefr` ainsi qu'un style `$resumefr`.  

### Conditions de fonctionnement de `base.dot`

La construction de `base.dot` doit respecter plusieurs contraintes. Le modèle `base.dot` respectant déjà ces conditions, il n'est normalement plus utile de toucher à ces éléments.

**Condition 1.** Le modèle `generator/src/base.dot` doit impérativement contenir la macro d'application des styles suivante :

```vba
' Macro d'application de style
' Cette macro doit impérativement être présente dans base.dot
Sub ApplyLodelStyle()
    Dim ctlCBarControl  As CommandBarControl
    Dim parameter As String
    Set ctlCBarControl = CommandBars.ActionControl
    If ctlCBarControl Is Nothing Then Exit Sub
    parameter = ctlCBarControl.parameter
    If parameter <> "" Then
		If IsNumeric(parameter) Then
			' BuiltIn Word style
			Selection.Range.Style = CInt(parameter)
		Else
			' User defined style
			Selection.Range.Style = parameter
        End If
    End If
End Sub
```

**Condition 2.** Le menu des styles de `base.dot` doit impérativement avoir pour nom : `LodelStyles`.

## Préparation de `translations.ini`

Le fichier `generator/src/translations.ini` contient la configuration du `generator` ainsi que les traductions des styles et des menus.

Il s'édite avec n'importe quel éditeur de texte, en respectant [la syntaxe des fichiers INI](https://fr.wikipedia.org/wiki/Fichier_INI "Fichier INI sur Wikipedia"). Il doit être encodé en utf-8.

### Section `[_configuration]`

`translations.ini` s'ouvre avec la section `[_configuration]` qui contient la configuration utilisée lors de l'exécution du `generator`.

#### Clé `translateTo`

La clé `translateTo` contient [les identifiants 639-1](https://fr.wikipedia.org/wiki/Liste_des_codes_ISO_639-1 "Codes ISO 639-1 sur Wikipedia") des langues de destination, séparés par des virgules.

Exemple :

```ini
translateTo="fr, en, es, pt"
```

#### Clé `version`

La clé `version` contient la version du modèle qui sera insérée un tant que corps de texte de chaque modèle. Forme: `[majeur].[mineur].[correctif]`. Pour des question de suivi des révisions, la version des modèles doit être incrémentée à chaque modification de `base.dot` ou de `translations.ini`.

Exemple :

```ini
version="4.0.1"
```

### Autres sections

Les autres sections correspondent chacune à un identifiant présent dans `base.dot` (voir plus haut). Les clés de la section déterminent le comportement pour l'objet (menu, style et/ou bouton) qui porte cet identifiant lors de l'exécution.  

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
	
#### Clés de traduction par défaut : `menu` et `style`

Les clés `menu` et `style` permettent d'attribuer des traductions par défaut qui seront appliquées si aucune traduction n'est trouvée dans la langue traitée. 

**Attention : l'utilisation d'une traduction par défaut ne provoque pas d'enregistrement dans le `log.txt`, il est donc recommandé de limiter son utilisation aux éléments qui ne sont jamais traduits (index linguistiques par exemple) afin de conserver un log pertinent.**

Exemple :

```ini
[$titreen]
style="Title (en)"
menu="Title (en)"
```

Comme la clé `[lang].style`, la clé `style` n'a aucun effet pour les sections qui ne sont pas associées à un style (menu, sous-menu).

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

#### Traitement des styles natifs de Word avec la clé `wordId`

Quand un style est nativement pris en charge par Word (style natif), il est préférable de ne pas le traduire et de le traiter en utilisant son identifiant Word. Word affichera automatiquement ce style dans la langue de l'utilisateur et le bouton d'application du style sera utilisable quelle que soit la version linguistique de Word (c'est donc cette option qui permet une utilisation universelle du modèle, quelle que soit la langue de Word). 

On utilise pour cela la clé `wordId` qui prend la valeur `WdBuiltinStyle` correspondante, à retrouver ici : https://msdn.microsoft.com/en-us/library/bb237495%28v=office.12%29.aspx

Exemple (ici pas besoin d'ajouter de traductions) :

```ini
[$titre]
wordId="-63"
```

**Remarque :** certains styles natifs (dont les citations) ne sont natifs que depuis Word 2007. Utiliser la clé `wordId ` avec ces styles produira donc une erreur lors de l'utilisation du modèle dans Word 2003.

Cet exemple fonctionnera donc sur toutes les versions linguistiques de Word mais uniquement à partir de Word 2007 : 

```ini
[$citation]
wordId="-181"
```

Tandis que celui-ci fonctionnera avec Word 2003, mais uniquement en français et en anglais :

```ini
[$citation]
fr.style="Citation"
fr.menu="Citation"
en.style="Quotation"
en.menu="Quotation"
```

#### Lien hypertexte : clé `link` (et `[lang].link`)

La clé `link` permet d'attribuer un redirection hypertexte à un bouton. Comme avec `style`, `menu` et `key`, il est possible de contextualiser la valeur à la langue du modèle. 

Voici par exemple un bouton qui ouvre le navigateur de l'utilisateur sur une page web différente selon la langue :

```ini
[$ordremetadonnees]
fr.menu="Ordre des métadonnées"
fr.hyperlink="http://maisondesrevues.org/108"
en.menu="Metadata's order"
en.hyperlink="http://maisondesrevues.org/404notfound"
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

## Exécution de la macro

La macro `generator.dot` doit être attachée dans Word en tant que "Modèles globaux et compléments" dans la fenêtre "Modèles et compléments". Deux boutons sont alors ajoutés à l'onglet "Compléments" de Word.

* Le bouton "Générer les modèles traduits" permet de lancer `generator`. Le journal des erreurs `log.txt` est affiché à la fin du traitement.
* Le bouton "Ouvrir le répertoire des modèles générés" affiche le répertoire `generator\build\` dans lequel se trouve les modèles produits par la macro.

## Journal des erreurs 

Le journal `build\log.txt` contient les éventuelles erreurs rencontrées lors de la dernière exécution du `generator`. Le plus souvent il s'agit d'une traduction manquante dans `translation.ini`. Le journal des erreurs permet donc de détecter les oublis de traductions ou les styles inexistants.

### Conditionner l'insertion d'un élément à la langue

Lorsqu'une erreur est rencontrée, le style et/ou le menu associé est supprimé du modèle généré. Dans certains cas ce comportement peut permettre de réserver l'inclusion d'un style, d'un menu ou d'un bouton à une langue en particulier. Par exemple ici on n'affiche pas le bouton `$ordremetadonnees` dans les modèles autres que français, ce bouton n'ayant pas d'utilité dans les autres langue.

```ini
[$ordremetadonnees]
fr.menu="Ordre des métadonnées"
fr.hyperlink="http://maisondesrevues.org/108"
```

## FAQ

### Un style a changé d'apparence au cours de la conversion

Il est probablement basé sur un autre style qui a été supprimé. Vérifier dans `log.txt` que tous les styles ont correctement été traduits ou changer le style de base du style.

### `log.txt` mentionne des styles indésirables

Tous les styles supplémentaires présents dans `base.dot` seront copiés dans les modèles, il est donc important de nettoyer correctement ce modèle. Si nécessaire (notamment pour nettoyer les styles " Car Car") on pourra utiliser la macro "style management.dot" : http://h2fooko.free.fr/spip.php?article19

**TODO:** la macro pourrait filtrer les styles afin de ne laisser passer que ceux qui commencent par $

### Les caractères accentués, les idéogrammes, etc. sont remplacés par des points d'interrogation 

Vérifier que `translations.ini` est bien encodé en utf-8.

### Quand le modèle généré est utilisé avec une version linguistique de Word qui ne correspond pas à la langue du modèle, il produit des erreurs lors de l'application d'un ou plusieurs styles

* Vérifier qu'aucune erreur du type "Style inexistant" n'a été signalée dans `log.txt`.
* Vérifier que le style qui pose problème n'est pas un style natif de Word (le cas échéant il faut utiliser la clé `wordId` comme indiqué ci-dessus).

### Tel bouton fonctionne mais pas son raccourci clavier

Word se réserve l'utilisation de certaines combinaisons de touches. Le cas échéant, l'option `key` n'est pas appliquée. Il faut alors essayer un autre raccourci clavier. 

### Comment faire pour qu'un bouton/menu/style ne soit supprimer dans telle langue

Voir plus haut "Conditionner l'insertion d'un élément à la langue".

### Word produit une erreur lors de l'exécution

Fermer toutes les instances de Word, relancer et réessayer la macro.

[Sinon c'est qu'il y a un bug...](https://github.com/brrd/revuesorg_word_templates/issues)