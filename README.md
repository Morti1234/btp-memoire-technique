# 🏗️ Agent IA – Générateur de Mémoire Technique BTP

Application Streamlit qui automatise la rédaction de Mémoires Techniques pour les appels d'offres du BTP.

Uploadez un CCTP (PDF), l'IA analyse les exigences et génère un document Word complet, personnalisé avec votre profil entreprise.

## 🚀 Déployer sur Streamlit Cloud (Snowflake)

### Étape 1 : Pousser le code sur GitHub

```bash
cd btp-memoire-technique
git init
git add .
git commit -m "Initial commit - Agent IA BTP"
git remote add origin https://github.com/VOTRE_USERNAME/btp-memoire-technique.git
git push -u origin main
```

### Étape 2 : Déployer sur Streamlit Cloud

1. Allez sur [share.streamlit.io](https://share.streamlit.io/)
2. Connectez-vous avec votre compte GitHub
3. Cliquez **« New app »**
4. Sélectionnez votre repo `btp-memoire-technique`
5. **Main file path** : `app.py`
6. Cliquez **« Deploy! »**

### Étape 3 : Configurer vos secrets

1. Dans le dashboard Streamlit Cloud, ouvrez votre app
2. Cliquez **⋮** → **Settings** → **Secrets**
3. Copiez-collez le contenu de `.streamlit/secrets.toml.example`
4. **Remplacez les valeurs** par vos vraies informations (clé API, profil entreprise)
5. Cliquez **Save**

> ⚠️ Les secrets sont **chiffrés** côté Streamlit Cloud. Ils ne sont jamais exposés publiquement.

## 🖥️ Lancer en local (développement)

```bash
pip install -r requirements.txt
streamlit run app.py
```

En local, la configuration est stockée dans `config.json` (créé automatiquement).

## 📁 Structure du projet

```
btp-memoire-technique/
├── app.py                            # Application principale
├── requirements.txt                  # Dépendances Python
├── .gitignore                        # Protège secrets.toml et config.json
├── .streamlit/
│   ├── config.toml                   # Thème et config Streamlit
│   └── secrets.toml.example          # Template des secrets (à ne PAS committer)
└── README.md                         # Ce fichier
```

## 🔧 Stack

- **Streamlit** – Interface web
- **pdfplumber** – Extraction PDF
- **OpenAI / Gemini** – Analyse IA + rédaction
- **python-docx** – Export Word
