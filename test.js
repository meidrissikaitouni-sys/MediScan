const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaStethoscope, FaRobot, FaDatabase, FaCode, FaGlobe, FaExclamationTriangle, FaCheckCircle, FaArrowRight, FaServer, FaDesktop, FaBrain } = require("react-icons/fa");

// ─── Color Palette ─────────────────────────────────────────────────
// Medical / Tech theme: deep navy + teal + white
const C = {
  navyDark:   "0A1628",
  navy:       "1A2E4A",
  navyMid:    "1E3A5F",
  teal:       "0D9488",
  tealLight:  "14B8A6",
  tealPale:   "CCFBF1",
  white:      "FFFFFF",
  offWhite:   "F0F4F8",
  grayLight:  "E2E8F0",
  grayMid:    "94A3B8",
  grayDark:   "475569",
  red:        "EF4444",
  orange:     "F97316",
  green:      "22C55E",
  accent:     "38BDF8",
};

// ─── Icon helper ───────────────────────────────────────────────────
async function iconPng(IconComponent, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const makeShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity: 0.15 });

// ─── Build Presentation ────────────────────────────────────────────
async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title  = "MediScan – Détecteur de Maladies par Symptômes";

  // ══════════════════════════════════════════════════════════════════
  // SLIDE 1 — TITRE
  // ══════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };

    // Gradient overlay strip (top)
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }, line: { color: C.teal } });

    // Decorative teal circle (right side)
    s.addShape(pres.shapes.OVAL, { x: 7.2, y: -0.8, w: 4.5, h: 4.5, fill: { color: C.teal, transparency: 88 }, line: { color: C.teal, transparency: 70, width: 1 } });
    s.addShape(pres.shapes.OVAL, { x: 7.8, y: -0.3, w: 3.0, h: 3.0, fill: { color: C.tealLight, transparency: 92 }, line: { color: C.tealLight, transparency: 80, width: 1 } });

    // Tag: ESISA
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.3, w: 1.8, h: 0.38, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.05 });
    s.addText("🏫 ESISA", { x: 0.5, y: 0.3, w: 1.8, h: 0.38, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });

    // Main title
    s.addText("MediScan", {
      x: 0.5, y: 1.1, w: 7, h: 1.1,
      fontSize: 56, bold: true, color: C.white,
      fontFace: "Cambria", align: "left", margin: 0,
    });

    // Subtitle
    s.addText("Détecteur Intelligent de Maladies par Symptômes", {
      x: 0.5, y: 2.15, w: 7, h: 0.65,
      fontSize: 18, color: C.tealLight, align: "left", margin: 0,
      fontFace: "Calibri",
    });

    // Divider line
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.88, w: 3.5, h: 0.04, fill: { color: C.teal }, line: { color: C.teal } });

    // Team members
    const members = [
      "Mehdi Idrissi Kaitouni",
      "Mosaab Zerrad",
      "Youssef El Bouazzaoui",
    ];
    members.forEach((m, i) => {
      s.addText(`👤 ${m}`, {
        x: 0.5, y: 3.15 + i * 0.42, w: 6, h: 0.38,
        fontSize: 14, color: C.grayLight, align: "left", margin: 0, fontFace: "Calibri",
      });
    });

    // Bottom info
    s.addText("Encadrant : Prof. Chafik Boulealam  •  Projet de Fin de Semestre  •  Juin 2025", {
      x: 0.5, y: 5.15, w: 9, h: 0.3,
      fontSize: 9, color: C.grayMid, align: "left", margin: 0, fontFace: "Calibri",
    });

    // Bottom teal strip
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.55, w: 10, h: 0.075, fill: { color: C.teal }, line: { color: C.teal } });

    // Notes
    s.addNotes("Bonjour, nous sommes [PRÉNOM] et notre groupe présente MediScan, un détecteur intelligent de maladies par symptômes. Notre équipe est composée de Mehdi, Mosaab et Youssef, encadrés par le Professeur Boulealam.");
  }

  // ══════════════════════════════════════════════════════════════════
  // SLIDE 2 — PROBLÉMATIQUE & OBJECTIFS
  // ══════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }, line: { color: C.teal } });

    // Section label
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fill: { color: C.navyDark }, line: { color: C.navyDark }, rectRadius: 0.05 });
    s.addText("01  PROBLÉMATIQUE", { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fontSize: 8, bold: true, color: C.tealLight, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    // Title
    s.addText("Pourquoi MediScan ?", {
      x: 0.5, y: 0.72, w: 9, h: 0.7,
      fontSize: 32, bold: true, color: C.navyDark, fontFace: "Cambria", align: "left", margin: 0,
    });

    // 3 problem cards
    const problems = [
      { icon: "❓", title: "L'incertitude médicale", desc: "Face à des symptômes courants, les gens ne savent pas quelle maladie peut être en cause ni quel médecin consulter." },
      { icon: "⏳", title: "Le temps d'attente", desc: "Les délais de consultation médicale peuvent être longs. Un premier aiguillage rapide est précieux." },
      { icon: "📊", title: "Le manque d'information", desc: "Il n'existe pas d'outil simple permettant d'obtenir un diagnostic préliminaire structuré avec niveau de risque." },
    ];

    problems.forEach((p, i) => {
      const x = 0.45 + i * 3.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.65, w: 2.9, h: 2.55,
        fill: { color: C.white }, line: { color: C.grayLight, width: 0.5 }, rectRadius: 0.1,
        shadow: makeShadow(),
      });
      s.addText(p.icon, { x, y: 1.78, w: 2.9, h: 0.5, fontSize: 26, align: "center", margin: 0 });
      s.addText(p.title, { x: x + 0.12, y: 2.3, w: 2.65, h: 0.45, fontSize: 13, bold: true, color: C.navyDark, align: "center", fontFace: "Calibri", margin: 0 });
      s.addText(p.desc, { x: x + 0.12, y: 2.78, w: 2.65, h: 1.25, fontSize: 10.5, color: C.grayDark, align: "center", fontFace: "Calibri", margin: 0 });
    });

    // Solution banner
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.4, w: 9, h: 0.9, fill: { color: C.navyDark }, line: { color: C.navyDark }, rectRadius: 0.08, shadow: makeShadow() });
    s.addText([
      { text: "Notre solution : ", options: { bold: false, color: C.tealLight } },
      { text: "MediScan", options: { bold: true, color: C.white } },
      { text: "  —  L'utilisateur décrit ses symptômes en langage naturel et obtient instantanément les maladies possibles avec un niveau de risque (faible / modéré / élevé) et des recommandations.", options: { bold: false, color: C.grayLight } },
    ], { x: 0.7, y: 4.4, w: 8.6, h: 0.9, fontSize: 11.5, valign: "middle", fontFace: "Calibri", margin: 0 });

    s.addNotes("Concrètement, le problème est simple : quand on a de la fièvre et qu'on tousse, on ne sait pas si c'est une grippe, une bronchite ou quelque chose de plus sérieux. MediScan répond à ce besoin en donnant un premier aiguillage médical intelligent, rapide, et structuré.");
  }

  // ══════════════════════════════════════════════════════════════════
  // SLIDE 3 — ARCHITECTURE TECHNIQUE
  // ══════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }, line: { color: C.teal } });

    // Section label
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.05 });
    s.addText("02  ARCHITECTURE", { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fontSize: 8, bold: true, color: C.navyDark, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    s.addText("Architecture Technique", {
      x: 0.5, y: 0.72, w: 9, h: 0.6,
      fontSize: 30, bold: true, color: C.white, fontFace: "Cambria", align: "left", margin: 0,
    });

    // ── 3 Architecture Layers ──
    const layers = [
      { label: "FRONTEND", icon: "🖥️", tech: "HTML / CSS / JS", desc: "Interface web utilisateur\nSaisie des symptômes\nAffichage des résultats", color: C.accent },
      { label: "BACKEND API", icon: "⚙️", tech: "Node.js + Express", desc: "Reçoit les symptômes\nConstruit le prompt Gemini\nRetourne une réponse JSON", color: C.teal },
      { label: "IA", icon: "🧠", tech: "Google Gemini", desc: "Analyse des symptômes\nRéponse structurée\nFallback local si besoin", color: C.orange },
    ];

    layers.forEach((l, i) => {
      const x = 0.38 + i * 3.12;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 1.52, w: 2.9, h: 2.8,
        fill: { color: C.navyMid }, line: { color: l.color, width: 1.5 }, rectRadius: 0.1,
        shadow: makeShadow(),
      });
      // Color top bar
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.52, w: 2.9, h: 0.38, fill: { color: l.color }, line: { color: l.color }, rectRadius: 0.1 });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.72, w: 2.9, h: 0.18, fill: { color: l.color }, line: { color: l.color } });

      s.addText(l.label, { x: x + 0.08, y: 1.54, w: 2.74, h: 0.36, fontSize: 10, bold: true, color: C.navyDark, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });
      s.addText(l.icon, { x: x + 0.08, y: 1.98, w: 2.74, h: 0.42, fontSize: 22, align: "center", margin: 0 });
      s.addText(l.tech, { x: x + 0.1, y: 2.42, w: 2.7, h: 0.38, fontSize: 10, bold: true, color: l.color, align: "center", margin: 0, fontFace: "Calibri" });
      s.addText(l.desc, { x: x + 0.12, y: 2.82, w: 2.66, h: 1.3, fontSize: 10, color: C.grayLight, align: "center", margin: 0, fontFace: "Calibri" });
    });

    // Arrows between layers
    const arrowOpts = { color: C.tealLight, width: 2, endArrowType: "arrow" };
    s.addShape(pres.shapes.LINE, { x: 3.28, y: 2.92, w: 0.62, h: 0, line: arrowOpts });
    s.addShape(pres.shapes.LINE, { x: 6.4, y: 2.92, w: 0.62, h: 0, line: arrowOpts });

    // Vercel badge
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.38, y: 4.52, w: 2.9, h: 0.72, fill: { color: C.navy }, line: { color: C.accent, width: 1 }, rectRadius: 0.08 });
    s.addText("📦 Version locale\nDémo exécutée en local", { x: 0.38, y: 4.52, w: 2.9, h: 0.72, fontSize: 9.5, color: C.accent, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    // CI/CD badge
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.5, y: 4.52, w: 2.9, h: 0.72, fill: { color: C.navy }, line: { color: C.tealLight, width: 1 }, rectRadius: 0.08 });
    s.addText("⚡ Démo locale\nBackend et frontend séparés", { x: 3.5, y: 4.52, w: 2.9, h: 0.72, fontSize: 9.5, color: C.tealLight, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    // Security badge
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.62, y: 4.52, w: 2.9, h: 0.72, fill: { color: C.navy }, line: { color: C.orange, width: 1 }, rectRadius: 0.08 });
    s.addText("🔐 Clé API en variable\nd'environnement", { x: 6.62, y: 4.52, w: 2.9, h: 0.72, fontSize: 9.5, color: C.orange, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    s.addNotes("Notre architecture est séparée en 3 couches. Le frontend, développé en HTML/CSS/JS, capture les symptômes. Il envoie ces données au backend Node.js/Express, qui construit le prompt et appelle Gemini via le SDK Google Generative AI. La réponse est renvoyée en JSON au frontend. Cette version du projet est pensée comme une démo locale, sans base de données ni déploiement public finalisés dans le dépôt actuel.");
  }

  // ══════════════════════════════════════════════════════════════════
  // SLIDE 4 — INTÉGRATION IA (GEMINI)
  // ══════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }, line: { color: C.teal } });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fill: { color: C.navyDark }, line: { color: C.navyDark }, rectRadius: 0.05 });
    s.addText("03  INTÉGRATION IA", { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fontSize: 8, bold: true, color: C.tealLight, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    s.addText("Pipeline IA — Google Gemini", {
      x: 0.5, y: 0.72, w: 9, h: 0.6,
      fontSize: 30, bold: true, color: C.navyDark, fontFace: "Cambria", align: "left", margin: 0,
    });

    // Pipeline steps
    const steps = [
      { num: "1", label: "Saisie utilisateur", desc: "\"J'ai de la fièvre et je tousse\"", color: C.navyDark, icon: "💬" },
      { num: "2", label: "Prompt Engineering", desc: "Backend Node.js construit le prompt structuré pour Gemini", color: C.teal, icon: "⚙️" },
      { num: "3", label: "Appel API Gemini", desc: "Le SDK Google Generative AI envoie la requête", color: C.navyMid, icon: "🌐" },
      { num: "4", label: "Parsing & Analyse", desc: "La réponse JSON est extraite et exploitée", color: C.teal, icon: "🔍" },
      { num: "5", label: "Résultat affiché", desc: "Maladies possibles + niveau de risque coloré", color: C.navyDark, icon: "📋" },
    ];

    steps.forEach((step, i) => {
      const x = 0.35 + i * 1.86;
      // Circle
      s.addShape(pres.shapes.OVAL, { x: x + 0.56, y: 1.52, w: 0.74, h: 0.74, fill: { color: step.color }, line: { color: step.color } });
      s.addText(step.icon, { x: x + 0.56, y: 1.52, w: 0.74, h: 0.74, fontSize: 16, align: "center", valign: "middle", margin: 0 });

      // Step number
      s.addText(step.num, { x, y: 1.54, w: 0.62, h: 0.38, fontSize: 10, bold: true, color: step.color, align: "left", margin: 0, fontFace: "Calibri" });

      // Label & desc
      s.addText(step.label, { x, y: 2.38, w: 1.75, h: 0.38, fontSize: 9.5, bold: true, color: C.navyDark, align: "center", margin: 0, fontFace: "Calibri" });
      s.addText(step.desc, { x, y: 2.76, w: 1.75, h: 0.7, fontSize: 8.5, color: C.grayDark, align: "center", margin: 0, fontFace: "Calibri" });

      // Arrow between steps
      if (i < 4) {
        s.addShape(pres.shapes.LINE, { x: x + 1.72, y: 1.89, w: 0.18, h: 0, line: { color: C.teal, width: 1.5, endArrowType: "arrow" } });
      }
    });

    // Gestion des erreurs card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.62, w: 4.4, h: 1.62, fill: { color: C.white }, line: { color: C.grayLight, width: 0.5 }, rectRadius: 0.08, shadow: makeShadow() });
    s.addText("🛡️  Robustesse & Gestion des erreurs", { x: 0.65, y: 3.72, w: 4.1, h: 0.4, fontSize: 12, bold: true, color: C.navyDark, margin: 0, fontFace: "Calibri" });
    s.addText([
      { text: "✅  Validation de la requête utilisateur\n", options: { breakLine: false } },
      { text: "✅  Gestion des erreurs JSON côté backend\n", options: { breakLine: false } },
      { text: "✅  Retour de repli local si l'API échoue", options: { breakLine: false } },
    ], { x: 0.65, y: 4.15, w: 4.1, h: 1.0, fontSize: 10, color: C.grayDark, margin: 0, fontFace: "Calibri" });

    // Risk levels card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.1, y: 3.62, w: 4.4, h: 1.62, fill: { color: C.white }, line: { color: C.grayLight, width: 0.5 }, rectRadius: 0.08, shadow: makeShadow() });
    s.addText("📊  Niveaux de risque retournés", { x: 5.25, y: 3.72, w: 4.1, h: 0.4, fontSize: 12, bold: true, color: C.navyDark, margin: 0, fontFace: "Calibri" });
    const risks = [
      { label: "FAIBLE", color: C.green }, { label: "MODÉRÉ", color: C.orange }, { label: "ÉLEVÉ", color: C.red },
    ];
    risks.forEach((r, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.25 + i * 1.35, y: 4.18, w: 1.22, h: 0.88, fill: { color: r.color }, line: { color: r.color }, rectRadius: 0.06 });
      s.addText(r.label, { x: 5.25 + i * 1.35, y: 4.18, w: 1.22, h: 0.88, fontSize: 9.5, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });
    });

    s.addNotes("Pour l'intégration IA, voici le pipeline complet. L'utilisateur saisit ses symptômes en langage naturel. Notre backend Node.js/Express construit un prompt structuré, puis appelle l'API Gemini via le SDK Google Generative AI. La réponse est extraite en JSON puis renvoyée au frontend. Si l'appel distant échoue, le backend peut retourner un résultat local de repli.");
  }

  // ══════════════════════════════════════════════════════════════════
  // SLIDE 5 — DÉMONSTRATION
  // ══════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }, line: { color: C.teal } });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.05 });
    s.addText("04  DÉMONSTRATION", { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fontSize: 8, bold: true, color: C.navyDark, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    s.addText("Demo Live — MediScan en action", {
      x: 0.5, y: 0.72, w: 9, h: 0.6,
      fontSize: 28, bold: true, color: C.white, fontFace: "Cambria", align: "left", margin: 0,
    });

    // Mock UI
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 1.45, w: 5.5, h: 3.9, fill: { color: C.navyMid }, line: { color: C.tealLight, width: 1 }, rectRadius: 0.12, shadow: makeShadow() });
    s.addText("🩺 MediScan", { x: 0.55, y: 1.58, w: 5.2, h: 0.38, fontSize: 14, bold: true, color: C.tealLight, margin: 0, fontFace: "Calibri" });

    // Input box
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.08, w: 5.1, h: 0.72, fill: { color: C.navy }, line: { color: C.grayMid, width: 0.5 }, rectRadius: 0.06 });
    s.addText("\"J'ai de la fièvre depuis 2 jours et je tousse...\"", { x: 0.7, y: 2.12, w: 4.9, h: 0.62, fontSize: 10.5, color: C.grayLight, valign: "middle", margin: 0, fontFace: "Calibri", italic: true });

    // Button
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.9, w: 5.1, h: 0.42, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.06 });
    s.addText("🔍  Analyser mes symptômes", { x: 0.6, y: 2.9, w: 5.1, h: 0.42, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    // Results
    const diseases = [
      { name: "Grippe (Influenza)", risk: "ÉLEVÉ", riskColor: C.red, prob: "78%" },
      { name: "Bronchite aiguë", risk: "MODÉRÉ", riskColor: C.orange, prob: "55%" },
      { name: "Rhinopharyngite", risk: "FAIBLE", riskColor: C.green, prob: "34%" },
    ];
    s.addText("Résultats :", { x: 0.65, y: 3.42, w: 5.0, h: 0.32, fontSize: 10, bold: true, color: C.tealLight, margin: 0, fontFace: "Calibri" });
    diseases.forEach((d, i) => {
      const y = 3.75 + i * 0.46;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 5.1, h: 0.4, fill: { color: C.navy }, line: { color: C.grayMid, width: 0.3 }, rectRadius: 0.05 });
      s.addText(d.name, { x: 0.72, y, w: 2.9, h: 0.4, fontSize: 10, color: C.white, valign: "middle", margin: 0, fontFace: "Calibri" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.7, y: y + 0.06, w: 1.0, h: 0.28, fill: { color: d.riskColor }, line: { color: d.riskColor }, rectRadius: 0.04 });
      s.addText(d.risk, { x: 3.7, y: y + 0.06, w: 1.0, h: 0.28, fontSize: 7.5, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });
      s.addText(d.prob, { x: 4.8, y, w: 0.75, h: 0.4, fontSize: 10.5, bold: true, color: C.tealLight, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });
    });

    // Right side: code snippet
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.1, y: 1.45, w: 3.5, h: 3.9, fill: { color: "0D1117" }, line: { color: C.teal, width: 1 }, rectRadius: 0.1 });
    s.addText("// Backend Node.js — Appel Gemini", { x: 6.2, y: 1.6, w: 3.3, h: 0.3, fontSize: 8.5, color: "6A9955", margin: 0, fontFace: "Courier New" });
    const codeLines = [
      { t: "const genAI =", c: C.accent },
      { t: "  new GoogleGenerativeAI(apiKey);", c: C.white },
      { t: "", c: C.white },
      { t: "const model =", c: C.tealLight },
      { t: "  genAI.getGenerativeModel({", c: C.orange },
      { t: "    model: 'gemini-2.0-flash'", c: C.white },
      { t: "", c: C.white },
      { t: "const response =", c: C.tealLight },
      { t: "  await model.generateContent(prompt);", c: C.orange },
      { t: "", c: C.white },
      { t: "// Parse JSON response", c: "6A9955" },
      { t: "const jsonMatch =", c: C.accent },
      { t: "  responseText.match(/\\{[\\s\\S]*\\}/);", c: C.white },
      { t: "const result =", c: C.tealLight },
      { t: "  JSON.parse(jsonMatch[0]);", c: C.white },
    ];
    codeLines.forEach((l, i) => {
      s.addText(l.t, { x: 6.2, y: 1.98 + i * 0.27, w: 3.3, h: 0.27, fontSize: 8, color: l.c, margin: 0, fontFace: "Courier New" });
    });

    s.addNotes("Voici la démonstration de MediScan. L'utilisateur décrit ses symptômes en langage naturel. L'application affiche les maladies possibles avec leur probabilité et leur niveau de risque coloré. À droite, vous voyez un extrait du backend Node.js qui appelle Gemini avec le SDK Google Generative AI et parse la réponse JSON.");
  }

  // ══════════════════════════════════════════════════════════════════
  // SLIDE 6 — RÉSULTATS, LIMITES & PERSPECTIVES
  // ══════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }, line: { color: C.teal } });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fill: { color: C.navyDark }, line: { color: C.navyDark }, rectRadius: 0.05 });
    s.addText("05  BILAN", { x: 0.5, y: 0.25, w: 2.2, h: 0.35, fontSize: 8, bold: true, color: C.tealLight, align: "center", valign: "middle", margin: 0, fontFace: "Calibri" });

    s.addText("Résultats, Limites & Perspectives", {
      x: 0.5, y: 0.72, w: 9, h: 0.6,
      fontSize: 28, bold: true, color: C.navyDark, fontFace: "Cambria", align: "left", margin: 0,
    });

    // Results
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 1.5, w: 2.9, h: 3.72, fill: { color: C.white }, line: { color: C.green, width: 1.5 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addText("✅  Ce qui fonctionne", { x: 0.55, y: 1.62, w: 2.6, h: 0.4, fontSize: 11, bold: true, color: C.green, margin: 0, fontFace: "Calibri" });
    const results = ["Saisie symptômes en langage naturel", "Analyse IA via Gemini", "3 niveaux de risque colorés", "Frontend et backend séparés", "Version locale démontrable", "Gestion d'erreurs basique"];
    results.forEach((r, i) => {
      s.addText(`✓  ${r}`, { x: 0.55, y: 2.1 + i * 0.52, w: 2.6, h: 0.45, fontSize: 9.5, color: C.grayDark, margin: 0, fontFace: "Calibri" });
    });

    // Limits
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.55, y: 1.5, w: 2.9, h: 3.72, fill: { color: C.white }, line: { color: C.orange, width: 1.5 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addText("⚠️  Limites actuelles", { x: 3.7, y: 1.62, w: 2.6, h: 0.4, fontSize: 11, bold: true, color: C.orange, margin: 0, fontFace: "Calibri" });
    const limits = ["Backend en JavaScript, pas en C", "Pas de base de données intégrée", "Pas de déploiement public finalisé", "Symptômes en français uniquement"];
    limits.forEach((l, i) => {
      s.addText(`!  ${l}`, { x: 3.7, y: 2.1 + i * 0.62, w: 2.6, h: 0.55, fontSize: 9.5, color: C.grayDark, margin: 0, fontFace: "Calibri" });
    });

    // Perspectives
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.7, y: 1.5, w: 2.9, h: 3.72, fill: { color: C.white }, line: { color: C.teal, width: 1.5 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addText("🚀  Perspectives", { x: 6.85, y: 1.62, w: 2.6, h: 0.4, fontSize: 11, bold: true, color: C.teal, margin: 0, fontFace: "Calibri" });
    const perspectives = ["Refonte du backend en C si exigé par le cahier des charges", "Ajout d'une base de données", "Déploiement public sur Vercel ou équivalent", "Version mobile dédiée"];
    perspectives.forEach((p, i) => {
      s.addText(`→  ${p}`, { x: 6.85, y: 2.1 + i * 0.62, w: 2.6, h: 0.55, fontSize: 9.5, color: C.grayDark, margin: 0, fontFace: "Calibri" });
    });

    s.addNotes("Notre application fonctionne comme une démonstration locale. Elle inclut un frontend web, un backend Node.js/Express et une intégration Gemini, mais le dépôt actuel ne contient pas encore un backend C, une base de données ni un déploiement public finalisé. Les perspectives consistent à ajouter ces couches si elles sont exigées pour la soutenance.");
  }

  // ══════════════════════════════════════════════════════════════════
  // SLIDE 7 — Q&A
  // ══════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navyDark };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }, line: { color: C.teal } });

    // Big decorative circle
    s.addShape(pres.shapes.OVAL, { x: 6.5, y: 0.5, w: 4.5, h: 4.5, fill: { color: C.teal, transparency: 90 }, line: { color: C.teal, transparency: 80, width: 1 } });

    s.addText("Questions ?", {
      x: 0.5, y: 0.9, w: 8, h: 1.4,
      fontSize: 58, bold: true, color: C.white, fontFace: "Cambria", align: "left", margin: 0,
    });
    s.addText("Nous sommes prêts à répondre 🎤", {
      x: 0.5, y: 2.28, w: 7, h: 0.5,
      fontSize: 16, color: C.tealLight, fontFace: "Calibri", align: "left", margin: 0,
    });

    // Anticipated Q&A
    const qas = [
      { q: "Pourquoi le backend n'est-il pas en C ?", a: "La version actuelle du dépôt est implémentée en Node.js/Express pour la démonstration. Si le cahier des charges impose le C, il faut remplacer cette couche." },
      { q: "Comment sécurisez-vous les credentials ?", a: "La clé API doit rester en variable d'environnement et ne pas être commise dans le dépôt." },
      { q: "Que se passe-t-il si l'API est hors ligne ?", a: "Le backend renvoie un fallback local basé sur les symptômes saisis pour garder une réponse exploitable." },
    ];

    qas.forEach((qa, i) => {
      const y = 2.98 + i * 0.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.82, fill: { color: C.navyMid }, line: { color: C.teal, width: 0.5 }, rectRadius: 0.08 });
      s.addText(`❓ ${qa.q}`, { x: 0.55, y: y + 0.04, w: 8.9, h: 0.32, fontSize: 10, bold: true, color: C.tealLight, margin: 0, fontFace: "Calibri" });
      s.addText(`    ${qa.a}`, { x: 0.55, y: y + 0.36, w: 8.9, h: 0.42, fontSize: 9.5, color: C.grayLight, margin: 0, fontFace: "Calibri" });
    });

    // Footer with names
    s.addText("Mehdi Idrissi Kaitouni  •  Mosaab Zerrad  •  Youssef El Bouazzaoui  •  ESISA — Juin 2025", {
      x: 0.5, y: 5.25, w: 9, h: 0.28,
      fontSize: 8.5, color: C.grayMid, align: "center", margin: 0, fontFace: "Calibri",
    });

    s.addNotes("Nous avons préparé des réponses aux questions les plus probables du jury. Restez calmes, répondez avec assurance. Si une question dépasse votre préparation, il est acceptable de dire : 'Excellente question, nous n'avons pas encore exploré cet aspect mais c'est une piste intéressante pour la suite.'");
  }

  await pres.writeFile({ fileName: "./MediScan_Presentation.pptx" });
  console.log("✅ Présentation générée avec succès !");
}

build().catch(console.error);