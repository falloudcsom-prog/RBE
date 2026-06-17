# Registre National des Bénéficiaires Effectifs
## Ministère de la Justice - République du Sénégal

---

## Lancement rapide (local)

```bash
pip install -r requirements.txt
python app.py
```

Ouvrir : http://localhost:5000

---

## Déploiement sur Render

1. Pousser ce dossier sur un dépôt GitHub ou GitLab.
2. Créer un nouveau service Web sur Render.
3. Paramètres :
   - **Build command** : `pip install -r requirements.txt`
   - **Start command** : `gunicorn wsgi:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
   - **Environment** : Python 3
4. Ajouter la variable d'environnement `SECRET_KEY` avec une valeur aléatoire longue.
5. La base SQLite est créée automatiquement au premier démarrage.

Pour PostgreSQL (recommandé en production), ajouter la variable :
```
DATABASE_URL=postgresql://user:password@host/dbname
```
Et modifier dans app.py :
```python
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///declarations.db')
```

---

## Compte administrateur par défaut

| Email | Mot de passe |
|-------|-------------|
| admin@rne.justice.sn | changez_ce_mot_de_passe |

**Changer le mot de passe dès le premier accès.**

---

## Structure

```
├── app.py            Application Flask principale
├── wsgi.py           Point d'entrée gunicorn
├── Procfile          Commande de démarrage (Render / Heroku)
├── render.yaml       Configuration Render
├── requirements.txt  Dépendances Python
├── templates/        Pages HTML (login, dashboard, déclaration, etc.)
├── static/           CSS et JavaScript
├── uploads/          Fichiers joints aux déclarations
└── exports/          Fichiers Excel générés
```

---

## Fonctionnalités

- Authentification sécurisée, rôles Admin et Utilisateur
- Formulaire complet de déclaration en 5 sections (A à E)
- Import en masse via fichier Excel avec prévisualisation
- Export Excel et CSV
- Tableau de bord avec statistiques et graphiques
- Consultation et recherche dans la base nationale
- Gestion des fichiers joints consultables directement
