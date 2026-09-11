# Mon Menu IA — serveur relais

Petit serveur relais (Cloudflare Worker) pour le fichier Excel « Mon Menu
Familial ». Il reçoit les préférences saisies dans Excel (nombre de
personnes, allergies, budget...), appelle l'API Gemini (Google) pour
générer un menu de 4 semaines structuré en JSON, et renvoie ce menu à
Excel via VBA.

La clé API Gemini n'est jamais présente dans ce dépôt ni dans le fichier
Excel : elle est stockée comme secret côté Cloudflare (voir plus bas).

## Déploiement

1. Connecter ce dépôt dans Cloudflare (Workers & Pages > Create > Import
   a repository), ou copier `worker.js` directement dans l'éditeur en
   ligne de Cloudflare.
2. Dans Settings > Variables and Secrets du Worker, ajouter deux
   secrets :
   - `GEMINI_API_KEY` — clé obtenue sur https://aistudio.google.com
   - `APP_SHARED_SECRET` — un mot de passe choisi par toi, à reporter
     aussi dans `modIA.bas` (constante `APP_SHARED_SECRET`) côté Excel
3. Déployer. Noter l'URL du Worker (ex.
   `https://mon-menu-ia.tonpseudo.workers.dev`) et la reporter dans
   `modIA.bas` (constante `WORKER_URL`).

## Fichiers

- `worker.js` — le code du serveur relais
- `wrangler.toml` — configuration Cloudflare (nom, point d'entrée)
