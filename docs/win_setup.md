# Installateur de modèles pour Revues.org

Installateur des modèles Word pour Windows.

## Compatibilité

* Word 2003 (version française uniquement) et versions ultérieures (toutes versions)
* Windows XP, Vista, 7, 8 et 10

## Limitations 

* Les modèles et macros sont installés dans les répertoires par défaut de  MS Office. Ils ne fonctionneront pas si l'utilisateur a modifié le chemin de ces répertoires dans la configuration de Word.
* L'installation pour Word 2003 ne fonctionnera pas avec les versions non francophones de Word.

## Développement

### Compilation

L'installateur est développé avec Inno Setup : http://jrsoftware.org/isinfo.php

Voir la documentation générale pour les instructions de compilation.

### Versions 

[Version des modèles].[Révision de l'installateur]

La version est automatiquement mise à jour à partir du fichier `translations.ini`.

## Changelog

**4.0.2.2**

* Utilisation intégrée avec le génrateur de modèles. 
* La version ets automatiquement retrouvée à partir de `translations.ini` (il n'est plus nécessaire de modifier le code source de l'installateur lors de la mise à jour des modèles).

**4.0.2.1**

* Installateur revue pour installer les nouveaux modèles (version 4).
* Suppression de l'installation personnalisée : tous les modèles du répertoire source sont copiés dans Templates et la macro de démarrage est copiée dans Startup.
* Test de l'existence des dossiers de destination (en plus du test du registre).

**3.1.4.3**

* L'installation ne nécessite plus les droits d'administrateur sur le poste. Le programme de désinstallation est stocké dans  le répertoire C:\Users\xxx\AppData\Local\Programs\RevuesOrgForWord.

**3.1.4.2**

* Test de l'existence de Word dans le registre (l'installation est annulée dans le cas contraire). Les versions proposées dans l'installateur sont celles détectées sur le poste.
* Ajout du menu de démarrage rapide pour le modèle standard fr.
