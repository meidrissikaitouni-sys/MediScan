// Backend MediScan - Analyse Gemini
// Utilise Google Generative AI et accepte les deux types de clés

const path = require('path');
const express = require('express');
const cors = require('cors');
const { GoogleGenerativeAI } = require('@google/generative-ai');

const app = express();
const PORT = process.env.PORT || 8787;

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

function sendPage(fileName) {
  return (req, res) => res.sendFile(path.join(__dirname, fileName));
}

app.use((err, req, res, next) => {
  if (err && err instanceof SyntaxError && 'body' in err) {
    return res.status(400).json({
      error: 'Requête JSON invalide',
      details: err.message
    });
  }

  return next(err);
});

app.get('/', sendPage('index.html'));
app.get('/mobile', sendPage('mobile.html'));
app.get('/test', sendPage('Test.html'));
app.get('/tester', sendPage('Testeur.html'));

const localDiseaseLibrary = [
  {
    name: 'Rhume',
    symptoms: ['toux', 'fatigue', 'mal de gorge', 'écoulement nasal', 'mal de tête'],
    description: 'Infection virale bénigne des voies respiratoires supérieures.'
  },
  {
    name: 'Grippe',
    symptoms: ['fièvre', 'fatigue', 'toux', 'mal de tête', 'douleur musculaire', 'mal de gorge'],
    description: 'Infection virale saisonnière avec symptômes généraux marqués.'
  },
  {
    name: 'Migraine',
    symptoms: ['mal de tête', 'nausée', 'vertiges', 'sensibilité à la lumière'],
    description: 'Céphalée intense pouvant être accompagnée de nausées et de sensibilité.'
  },
  {
    name: 'Gastro-entérite',
    symptoms: ['nausée', 'diarrhée', 'fatigue', 'fièvre', 'douleur abdominale'],
    description: 'Infection digestive provoquant troubles digestifs et fatigue.'
  },
  {
    name: 'Sinusite',
    symptoms: ['mal de tête', 'écoulement nasal', 'mal de gorge', 'pression faciale'],
    description: 'Inflammation des sinus avec pression faciale et congestion.'
  },
  {
    name: 'Allergie',
    symptoms: ['écoulement nasal', 'éruption cutanée', 'démangeaisons', 'mal de tête'],
    description: 'Réaction allergique pouvant provoquer des symptômes respiratoires ou cutanés.'
  }
];

function normalizeText(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function buildLocalResults(symptomsText) {
  const userSymptoms = normalizeText(symptomsText)
    .split(',')
    .map((symptom) => symptom.trim())
    .filter(Boolean);

  const scored = localDiseaseLibrary.map((disease) => {
    const matchedSymptoms = disease.symptoms.filter((symptom) =>
      userSymptoms.some((userSymptom) => userSymptom.includes(normalizeText(symptom)) || normalizeText(symptom).includes(userSymptom))
    );

    const coverage = matchedSymptoms.length / Math.max(disease.symptoms.length, userSymptoms.length || 1);
    const riskPercentage = Math.max(10, Math.min(95, Math.round(coverage * 100)));

    return {
      name: disease.name,
      riskPercentage,
      description: disease.description,
      matchedSymptoms
    };
  });

  return scored
    .filter((disease) => disease.riskPercentage >= 20)
    .sort((a, b) => b.riskPercentage - a.riskPercentage)
    .slice(0, 5);
}

// Endpoint pour analyser les symptômes
app.post('/api/analyze', async (req, res) => {
  try {
    const { apiKey, symptoms, duration, severity, history } = req.body;

    if (!apiKey) {
      return res.status(400).json({ error: 'Clé API manquante' });
    }
    if (!symptoms) {
      return res.status(400).json({ error: 'Symptômes manquants' });
    }

    // Détecter le type de clé
    const isGoogleKey = apiKey.startsWith('AIzaSy');
    const isVertexKey = apiKey.startsWith('AQ.');

    if (!isGoogleKey && !isVertexKey) {
      return res.status(400).json({ 
        error: 'Format de clé API invalide. Doit commencer par AIzaSy ou AQ.' 
      });
    }

    if (isVertexKey) {
      const fallbackData = buildLocalResults(req.body.symptoms);
      return res.status(200).json({
        success: true,
        data: { diseases: fallbackData, source: 'local-fallback' },
        fallbackUsed: true,
        fallbackMessage: 'La clé AQ.* correspond à Vertex AI et ne peut pas être utilisée directement ici. Résultats générés localement à partir des symptômes.'
      });
    }

    // Préparer le prompt
    const prompt = `Tu es un assistant médical expert. Analyse les symptômes suivants et fournis une liste des maladies les plus probables avec un taux de correspondance.

SYMPTÔMES: ${symptoms}
DURÉE: ${duration || 'Non spécifiée'}
SÉVÉRITÉ: ${severity || 'Non spécifiée'}
ANTÉCÉDENTS: ${history || 'Aucun spécifié'}

INSTRUCTIONS:
1. Identifie les 5-7 maladies les plus probables
2. Pour chaque maladie, fournis:
   - Nom de la maladie
   - Taux de risque (0-100%)
   - Description médicale brève
   - Symptômes qui correspondent
3. Format ta réponse en JSON valide UNIQUEMENT:
{
  "diseases": [
    {
      "name": "Nom de la maladie",
      "riskPercentage": 85,
      "description": "Description courte",
      "matchedSymptoms": ["symptôme1", "symptôme2"]
    }
  ]
}

IMPORTANT: Retourne UNIQUEMENT le JSON valide, sans texte supplémentaire.`;

    let result;
    let fallbackUsed = false;
    let fallbackMessage = '';

    if (isGoogleKey) {
      // Utiliser Google Generative AI pour les clés AIzaSy
      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash' });
      const response = await model.generateContent(prompt);
      const responseText = response.response.text();
      
      // Parser le JSON
      const jsonMatch = responseText.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error('Format de réponse invalide');
      }
      result = JSON.parse(jsonMatch[0]);
    }

    res.json({ success: true, data: result, fallbackUsed, fallbackMessage });
  } catch (error) {
    console.error('Erreur:', error);
    const fallbackData = buildLocalResults(req.body.symptoms);
    return res.status(200).json({
      success: true,
      data: { diseases: fallbackData, source: 'local-fallback' },
      fallbackUsed: true,
      fallbackMessage: `L'API distante a échoué (${error.message || 'erreur inconnue'}). Résultats générés localement à partir des symptômes.`
    });
  }
});

// Health check
app.get('/health', (req, res) => {
  res.json({ status: 'ok', message: 'MediScan Backend is running' });
});

module.exports = app;

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`🏥 MediScan Backend running on http://localhost:${PORT}`);
    console.log(`📝 POST http://localhost:${PORT}/api/analyze`);
    console.log(`✅ Health check: http://localhost:${PORT}/health`);
  });
}
