# Liasse mails et factures

Application Streamlit (Capeos) : on dépose des mails Outlook (.msg) et on récupère un seul PDF à imprimer.

- Chaque mail est imprimé sur sa propre page, suivi de ses pièces jointes (PDF, images), par date d'envoi.
- La facture et le RIB donnés en lien dans le mail (« Facture : », « RIB : ») sont téléchargés automatiquement, uniquement depuis les sites autorisés (`sngtv.org` par défaut). Les liens « Justificatif » sont ignorés.
- Facture Orange : 4 premières pages, une page d'avis « Document tronqué », puis la dernière page.
- Au-delà de 10 pages, la même troncature est proposée (case cochée par défaut).
- Les mails en double sont écartés. Une pièce identique dans deux mails (un même RIB, par exemple) est imprimée et signalée.
- Une rubrique « À vérifier » liste ce qui n'a pas pu être récupéré.

## Contenu

| Fichier | Rôle |
|---|---|
| `app.py` | Interface Streamlit |
| `liasse_core.py` | Traitement (lecture des mails, téléchargements, PDF). Réglages en tête de fichier. |
| `requirements.txt` | Modules Python, versions figées |
| `.streamlit/config.toml` | Couleurs Capeos, statistiques d'usage et détails d'erreur désactivés |
| `polices/` | Police Carlito (métriques de Calibri), licence SIL OFL 1.1 |
| `.gitignore` | Empêche de publier des .msg ou des PDF de liasse |

## Déploiement sur Streamlit Community Cloud

1. Créer un dépôt GitHub **privé** et y déposer le contenu de ce dossier (dossiers `.streamlit` et `polices` compris). Ne jamais y mettre de mails ni de PDF clients.
2. Sur share.streamlit.io, se connecter avec le compte GitHub, puis créer une app : choisir le dépôt, la branche et le fichier `app.py`. Dans les paramètres avancés, choisir Python 3.12.
3. Une fois l'app en ligne, la rendre privée dans ses paramètres de partage et inviter les collègues par adresse e-mail.

## Réglages (`liasse_core.py`)

- `SEUIL_PAGES = 10` : au-delà, troncature proposée.
- `PAGES_DEBUT = 4` : pages gardées au début d'un document tronqué.
- `DOMAINES_AUTORISES = ("sngtv.org",)` : seuls sites d'où l'app télécharge. Ajouter un domaine ici si un autre fournisseur envoie ses factures en lien.

## Données

L'app ne conserve rien : les fichiers sont traités en mémoire le temps de la session, sans écriture sur disque ni journalisation de leur contenu. Les mails, factures et RIB transitent toutefois par les serveurs de Streamlit Community Cloud, situés aux États-Unis (région non modifiable selon la documentation Streamlit). Cet usage est à faire valider par le référent RGPD de Capeos avant de traiter des données clients réelles.

## Limites connues

- Un scan sans texte ne peut pas être reconnu comme facture Orange.
- Un PDF protégé par mot de passe n'est pas repris (signalé « à imprimer à part »).
- Si un lien a expiré ou si le site demande une connexion, le document est signalé dans « À vérifier ».
