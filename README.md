# CORBE49 — Transfert Outlook vers Teams

Complément Outlook pour transférer un e-mail vers une conversation Microsoft Teams existante ou nouvelle. Le nom du projet et les fichiers sont prévus pour rester les noms définitifs.

## Objectifs

- Utilisation prioritaire dans Outlook mobile Android.
- Sélection d'une conversation individuelle ou de groupe existante.
- Création d'une nouvelle conversation individuelle ou de groupe.
- Préservation autant que possible de l'objet, de l'expéditeur, du corps du message et des pièces jointes.
- Commentaire facultatif avant l'envoi.

## Fichiers

- `manifest.xml` : manifeste Office XML avec point d'entrée mobile.
- `taskpane.html` : interface du complément.
- `taskpane.js` : lecture du message sélectionné et logique de l'interface.
- `taskpane.css` : styles.
- `index.html` : page d'accueil du site GitHub Pages.

## État actuel

La base du complément est créée. L'envoi vers Teams reste désactivé tant que l'authentification Microsoft Entra ID, les autorisations Microsoft Graph, la sélection des conversations et le transfert des pièces jointes n'ont pas été configurés et validés.

Avant de valider ou de déployer le manifeste, il reste notamment à ajouter les icônes référencées dans `manifest.xml`, activer GitHub Pages et valider le manifeste avec les outils Microsoft. Les URL GitHub Pages ne fonctionneront qu'après activation de la publication.

## Sécurité

- Ne jamais placer de secret client, mot de passe ou jeton d'accès dans ce dépôt.
- Utiliser une authentification déléguée Microsoft pour agir au nom de l'utilisateur.
- Limiter les autorisations Graph à celles strictement nécessaires.
- Valider les autorisations et le manifeste avant le déploiement aux collaborateurs.
