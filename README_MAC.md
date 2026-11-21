# GED2DOCX pour Mac — installation simple

Ce guide est pour vous si vous voulez **juste utiliser l'application**.
Vous n'avez **rien à installer** au préalable (ni Python, ni autre logiciel) : tout est inclus dans l'application.

GED2DOCX transforme un fichier de généalogie **.ged** (exporté depuis Geneanet, Heredis, MyHeritage, etc.) en un document **Word (.docx)** mis en page.

---

## 1. Télécharger l'application

1. Ouvrez cette page : **https://github.com/mohawk2409/genopresse-project/releases/latest**
2. Dans la section **Assets** en bas de la page, cliquez sur le fichier qui correspond à votre Mac :

   | Votre Mac | Fichier à télécharger |
   |---|---|
   | Puce **Apple** (M1, M2, M3, M4… — Mac achetés depuis fin 2020) | `GED2DOCX-mac-apple-silicon.zip` |
   | Processeur **Intel** (Mac plus anciens) | `GED2DOCX-mac-intel.zip` |

   > **Comment savoir ?** Cliquez sur le menu  (en haut à gauche de l'écran) → **À propos de ce Mac**.
   > Si vous voyez « Puce Apple M… », prenez *apple-silicon*. Si vous voyez « Processeur … Intel », prenez *intel*.

3. Le fichier arrive dans votre dossier **Téléchargements**.

## 2. Installer

1. Ouvrez le dossier **Téléchargements**.
2. Double-cliquez sur le fichier `.zip` : une icône **GED2DOCX** apparaît à côté.
3. Faites glisser **GED2DOCX** dans le dossier **Applications** (dans la barre latérale du Finder).

## 3. Premier lancement (une seule fois)

L'application n'est pas distribuée par l'App Store, donc macOS demande une confirmation la première fois. C'est normal.

1. Ouvrez le dossier **Applications** et double-cliquez sur **GED2DOCX**.
2. Un message indique que macOS n'a pas pu vérifier l'application. Cliquez sur **Terminé** (ou **OK**).
3. Ouvrez **Réglages Système** (menu  → Réglages Système…) → **Confidentialité et sécurité**.
4. Faites défiler vers le bas jusqu'au message *« GED2DOCX a été bloqué… »* et cliquez sur **Ouvrir quand même**.
5. Entrez le mot de passe de votre Mac, puis cliquez à nouveau sur **Ouvrir quand même**.

Les fois suivantes, un simple double-clic suffit.

> **Sur un Mac plus ancien (macOS 14 Sonoma ou avant)**, il y a plus simple : faites un **clic droit** (ou Ctrl + clic) sur GED2DOCX → **Ouvrir** → **Ouvrir**.

## 4. Utiliser l'application

1. **Fichier GED** : cliquez sur **Browse GED** et choisissez votre fichier `.ged`.
2. **Nom du JSON** : laissez `report` ou tapez un autre nom, puis cliquez sur **Convert GED → JSON**.
3. **Sortie DOCX** : tapez le nom du document Word à créer (par exemple `ma-famille`).
4. *(Facultatif)* Cliquez sur **Index** pour choisir les index à ajouter (noms, lieux, professions), puis sur **← Retour**.
5. Cliquez sur **Generate**.

📁 **Vos fichiers sont enregistrés dans : Documents → GED2DOCX.**
Ouvrez le fichier `.docx` avec Word, Pages ou LibreOffice.

## 5. Mettre à jour

Téléchargez la nouvelle version (étape 1), puis faites glisser la nouvelle icône dans **Applications** et choisissez **Remplacer**. Vos documents dans *Documents → GED2DOCX* ne sont pas touchés.

---

## En cas de problème

**« GED2DOCX est endommagé et ne peut pas être ouvert »** ou le bouton *Ouvrir quand même* n'apparaît pas :
1. Ouvrez l'application **Terminal** (Applications → Utilitaires → Terminal).
2. Copiez-collez cette ligne puis appuyez sur Entrée :
   ```
   xattr -dr com.apple.quarantine /Applications/GED2DOCX.app
   ```
3. Relancez GED2DOCX.

**L'application ne s'ouvre pas du tout** : vérifiez que vous avez téléchargé la bonne version (Apple Silicon ou Intel, voir étape 1).

**Le bouton « Generate » affiche une erreur** : vérifiez que vous avez bien cliqué sur *Convert GED → JSON* avant, et que le champ *Fichier JSON* n'est pas vide.

Pour toute autre question, contactez la personne qui vous a envoyé ce lien.
