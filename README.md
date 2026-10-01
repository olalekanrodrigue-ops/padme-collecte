# Collecte PADME Pobè

Questionnaire de 7 questions, base SQLite, export Excel (DONNEES, CODEBOOK, README).

## Variables d'environnement
- ADMIN_PASSWORD : mot de passe de /admin (obligatoire en production, identifiant libre)
- DB_PATH : chemin du fichier SQLite (défaut reponses.db). À placer sur un disque persistant.
- PORT : port d'écoute (défaut 5000)

## Lancer en local
    pip install -r requirements.txt
    ADMIN_PASSWORD=monmotdepasse python app.py
Formulaire : http://localhost:5000  |  Admin : http://localhost:5000/admin

## Tester
    python test_app.py

## Déployer
- Render ou Railway : créer un service web Python depuis ce dépôt, commande de démarrage
  `gunicorn -b 0.0.0.0:$PORT -w 2 app:app`, ajouter un disque persistant monté sur /data,
  définir DB_PATH=/data/reponses.db et ADMIN_PASSWORD.
- PythonAnywhere : téléverser les fichiers, créer une web app Flask pointant sur app.py
  (le disque est persistant par défaut).
- Docker : `docker build -t padme-collecte .` puis
  `docker run -p 5000:5000 -v padme-data:/data -e ADMIN_PASSWORD=monmotdepasse padme-collecte`

## Remarques
- Question 7 facultative : mettre APPRECIATION_OBLIGATOIRE = False dans app.py avant la première utilisation.
- Ne pas modifier les codes après le début de la collecte.
- Sauvegarder régulièrement le fichier SQLite ou télécharger l'export Excel.
