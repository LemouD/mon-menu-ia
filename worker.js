// =====================================================================
//  worker.js — petit serveur relais Cloudflare Worker
//  ---------------------------------------------------------------
//  Rôle : reçoit les paramètres depuis Excel (VBA), appelle l'API
//  Gemini (Google) avec la clé secrète (jamais visible dans Excel),
//  et renvoie un menu structuré en JSON, prêt à être lu par VBA.
//
//  Variables à configurer dans Cloudflare (Settings > Variables and
//  Secrets), voir les instructions fournies à côté de ce fichier :
//    - GEMINI_API_KEY   : ta clé API Google AI Studio (secret)
//    - APP_SHARED_SECRET: un mot de passe que toi seule choisis, pour
//                          éviter que n'importe qui sur internet
//                          utilise ton serveur (secret)
//    - PEXELS_API_KEY   : ta clé gratuite de l'API Pexels (secret),
//                          utilisée pour trouver une photo par plat sur
//                          la feuille "Détail des menus". Facultatif :
//                          si elle n'est pas configurée, les menus se
//                          génèrent quand même, simplement sans photos.
// =====================================================================

const MODEL_NAME = "gemini-3.6-flash";
// Si Google retire ce modèle du plan gratuit, regarde la liste à jour
// sur https://aistudio.google.com/ et remplace la valeur ci-dessus.

export default {
  async fetch(request, env) {
    if (request.method !== "POST") {
      return jsonResponse({ error: "Utilise une requête POST." }, 405);
    }

    // --- Vérification du mot de passe partagé -------------------------
    const secretRecu = request.headers.get("X-App-Secret") || "";
    if (!env.APP_SHARED_SECRET || secretRecu !== env.APP_SHARED_SECRET) {
      return jsonResponse({ error: "Non autorisé." }, 401);
    }

    let params;
    try {
      params = await request.json();
    } catch (e) {
      return jsonResponse({ error: "Corps de requête JSON invalide." }, 400);
    }

    // --- Recherche d'une photo de plat (feuille "Détail des menus") ---
    if (params.type === "photo") {
      return handlePhoto(params, env);
    }

    const nbPersonnes = Number(params.nbPersonnes) || 5;
    const nbSemaines = Math.min(Math.max(Number(params.nbSemaines) || 4, 1), 8);
    const preferences = String(params.preferences || "Aucune restriction");
    const allergies = String(params.allergies || "Aucune");
    const budgetMensuel = Math.max(Number(params.budgetMensuel) || 0, 0);
    const tempsPrep = String(params.tempsPrep || "30 min");
    const alimentsAimes = String(params.alimentsAimes || "");
    const alimentsDetestes = String(params.alimentsDetestes || "");

    const prompt = construirePrompt({
      nbPersonnes, nbSemaines, preferences, allergies, budgetMensuel, tempsPrep,
      alimentsAimes, alimentsDetestes,
    });

    const schema = construireSchema();

    let geminiResponse;
    try {
      geminiResponse = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${env.GEMINI_API_KEY}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: {
              temperature: 0.9,
              responseMimeType: "application/json",
              responseSchema: schema,
            },
          }),
        }
      );
    } catch (e) {
      return jsonResponse({ error: "Impossible de joindre l'IA : " + e.message }, 502);
    }

    if (!geminiResponse.ok) {
      const detail = await geminiResponse.text();
      return jsonResponse(
        { error: "L'IA a renvoyé une erreur (" + geminiResponse.status + ") : " + detail.slice(0, 500) },
        502
      );
    }

    const geminiData = await geminiResponse.json();
    const texte = geminiData &&
      geminiData.candidates &&
      geminiData.candidates[0] &&
      geminiData.candidates[0].content &&
      geminiData.candidates[0].content.parts &&
      geminiData.candidates[0].content.parts[0] &&
      geminiData.candidates[0].content.parts[0].text;

    if (!texte) {
      return jsonResponse({ error: "Réponse de l'IA vide ou inattendue." }, 502);
    }

    // texte est déjà le JSON du menu (grâce à responseSchema) : on le
    // renvoie tel quel à Excel, avec le bon Content-Type.
    return new Response(texte, {
      status: 200,
      headers: { "Content-Type": "application/json; charset=utf-8" },
    });
  },
};

function jsonResponse(obj, status) {
  return new Response(JSON.stringify(obj), {
    status: status || 200,
    headers: { "Content-Type": "application/json; charset=utf-8" },
  });
}

// ---------------------------------------------------------------------
//  Recherche d'une photo libre de droits pour un plat, via l'API
//  Pexels. On ne renvoie que l'URL publique de l'image (pas besoin de
//  la clé Pexels pour la télécharger ensuite depuis Excel) : la clé
//  API reste donc uniquement côté serveur, jamais dans Excel/VBA.
// ---------------------------------------------------------------------
async function handlePhoto(params, env) {
  const query = String(params.query || "").trim().slice(0, 100);
  if (!query) {
    return jsonResponse({ error: "Requête de recherche vide." }, 400);
  }
  if (!env.PEXELS_API_KEY) {
    return jsonResponse({ error: "PEXELS_API_KEY n'est pas configurée sur le Worker." }, 500);
  }

  let pexelsResponse;
  try {
    pexelsResponse = await fetch(
      "https://api.pexels.com/v1/search?query=" + encodeURIComponent(query) +
        "&per_page=1&orientation=square",
      { headers: { Authorization: env.PEXELS_API_KEY } }
    );
  } catch (e) {
    return jsonResponse({ error: "Impossible de joindre Pexels : " + e.message }, 502);
  }

  if (!pexelsResponse.ok) {
    const detail = await pexelsResponse.text();
    return jsonResponse(
      { error: "Pexels a renvoyé une erreur (" + pexelsResponse.status + ") : " + detail.slice(0, 300) },
      502
    );
  }

  const data = await pexelsResponse.json();
  const photo = data && data.photos && data.photos[0];
  const url = (photo && photo.src && (photo.src.medium || photo.src.small)) || null;

  return jsonResponse({ url: url }, 200);
}

function construirePrompt(p) {
  const consigneBudget = p.budgetMensuel > 0
    ? `- Budget total pour les courses de tout le mois : environ ${p.budgetMensuel} euros. ` +
      `Choisis des ingrédients et des quantités raisonnables (privilégie les produits de saison, ` +
      `les légumineuses, les féculents, et limite les protéines animales les plus chères) pour ` +
      `rester autant que possible dans ce budget sur l'ensemble du mois.`
    : `- Budget : non précisé, reste raisonnable.`;

  return `Tu es un(e) nutritionniste qui prépare un plan de repas familial pour une maman en France.

Génère un menu complet pour ${p.nbSemaines} semaine(s), pour ${p.nbPersonnes} personne(s).

Contraintes à respecter strictement :
- Préférences alimentaires : ${p.preferences}
- Allergies / aliments à éviter absolument : ${p.allergies}
${consigneBudget}
- Temps de préparation maximum par repas : ${p.tempsPrep}
- Aliments aimés à privilégier si possible : ${p.alimentsAimes || "(aucune préférence particulière)"}
- Aliments détestés à éviter si possible : ${p.alimentsDetestes || "(aucun)"}

Pour chaque jour (Lundi à Dimanche) de chaque semaine, propose 4 repas : un petit-déjeuner,
un jus/smoothie maison, un déjeuner et un dîner. Varie les plats d'un jour à l'autre autant
que possible, avec des légumes variés et des repas équilibrés. Les recettes doivent être
simples et réalistes pour une famille.

Pour chaque plat, donne son nom et la liste de ses ingrédients avec la quantité totale
nécessaire pour ${p.nbPersonnes} personne(s) (pas par personne), et l'unité (g, kg, ml, L,
pièce(s), cuillère à soupe, etc.). Les noms d'ingrédients doivent être simples et génériques
(ex: "carottes", "poulet", "riz"), pas de quantité dans le nom.

Réponds avec un tableau plat "jours" contenant EXACTEMENT ${p.nbSemaines * 7} éléments (un par
jour), dans cet ordre précis : d'abord les 7 jours de la semaine 1 (Lundi, Mardi, Mercredi,
Jeudi, Vendredi, Samedi, Dimanche dans cet ordre), puis les 7 jours de la semaine 2, et ainsi
de suite jusqu'à la semaine ${p.nbSemaines}. Pour chaque élément, indique aussi le numéro de
semaine (NumeroSemaine, de 1 à ${p.nbSemaines}) et le nom du jour (NomJour).

Réponds uniquement avec les données demandées, dans la langue française, en respectant
strictement le format demandé.`;
}

function construireSchema() {
  const ingredientSchema = {
    type: "object",
    properties: {
      Nom: { type: "string" },
      Quantite: { type: "number" },
      Unite: { type: "string" },
    },
    required: ["Nom", "Quantite", "Unite"],
  };

  const platSchema = {
    type: "object",
    properties: {
      Nom: { type: "string" },
      Ingredients: { type: "array", items: ingredientSchema },
    },
    required: ["Nom", "Ingredients"],
  };

  const jourSchema = {
    type: "object",
    properties: {
      NumeroSemaine: { type: "number" },
      NomJour: { type: "string" },
      PetitDej: platSchema,
      Jus: platSchema,
      Dejeuner: platSchema,
      Diner: platSchema,
    },
    required: ["NumeroSemaine", "NomJour", "PetitDej", "Jus", "Dejeuner", "Diner"],
  };

  return {
    type: "object",
    properties: {
      jours: {
        type: "array",
        items: jourSchema,
      },
    },
    required: ["jours"],
  };
}