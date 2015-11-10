# FAQ

## Générateur de modèles

### Un style a changé d'apparence au cours de la conversion

Il est probablement basé sur un autre style qui a été supprimé. Vérifier dans `log.txt` que tous les styles ont correctement été traduits ou changer le style de base du style.

### `log.txt` mentionne des styles indésirables

Tous les styles supplémentaires présents dans `base.dot` seront copiés dans les modèles, il est donc important de nettoyer correctement ce modèle. Si nécessaire (notamment pour nettoyer les styles " Car Car") on pourra utiliser la macro "style management.dot" : http://h2fooko.free.fr/spip.php?article19

### Les caractères accentués, les idéogrammes, etc. sont remplacés par des points d'interrogation

Vérifier que `translations.ini` est bien encodé en utf-8.

### Tel bouton fonctionne mais pas son raccourci clavier

Word se réserve l'utilisation de certaines combinaisons de touches. Le cas échéant, l'option `key` n'est pas appliquée. Il faut alors essayer un autre raccourci clavier.

### Le bouton d'un style ne fonctionne pas

S'il s'agit d'un style natif, vérifier que son attribut `[lang].style` de `translations.ini` correspond exactement au nom du style dans la langue cible.

### Comment faire pour qu'un bouton/menu/style soit supprimé dans une langue définie

Voir dans [la documentation sur translations.ini](template_generator.md) "Conditionner l'insertion d'un élément à la langue".

### Word produit une erreur lors de l'exécution

Fermer toutes les instances de Word, relancer et réessayer la macro.
