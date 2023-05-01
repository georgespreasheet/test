pour utiliser l'API Google Voice dans Google Apps Script, vous devez d'abord activer l'API Google Voice et créer une clé d'API pour votre projet Google Cloud. Voici les étapes à suivre pour configurer l'API Google Voice dans votre projet :

    Ouvrez le "Google Cloud Console" à l'adresse https://console.cloud.google.com/.

    Si vous n'avez pas déjà créé un projet, créez-en un en cliquant sur le bouton "Sélectionner un projet" en haut de la page, puis en cliquant sur le bouton "Nouveau projet".

    Activez l'API Google Voice en cliquant sur le bouton "Activer les API et les services" en haut de la page, puis en recherchant "Google Voice".

    Cliquez sur l'API "Google Voice API" pour la sélectionner, puis cliquez sur le bouton "Activer".

    Créez une clé d'API pour votre projet en cliquant sur le lien "Créer des identifiants" en haut de la page, puis en sélectionnant "Clé d'API".

    Sélectionnez "RESTRICTED KEY" pour restreindre votre clé d'API aux appels Google Voice, puis cliquez sur le bouton "Créer".

    Copiez la clé d'API et enregistrez-la pour une utilisation ultérieure dans votre script.

    Dans votre script Google Apps, accédez au menu "Ressources > Bibliothèques" et ajoutez la bibliothèque "Google API Client". Entrez "MjRmMmUzN2MtOWY5ZC00OWU5LWE2NWUtODU5NjU1MzU2MGM2" comme identifiant de bibliothèque.

    Dans votre script, utilisez la méthode getGoogleService() pour configurer le service OAuth2 avec votre clé d'API et vos informations de connexion Google.
