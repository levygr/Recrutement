const fs = require("fs");
const D = require("docx");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, HeadingLevel,
  PageBreak, Header, Footer, PageNumber, LevelFormat, VerticalAlign
} = D;
const { PROBES, SITUATIONS, FICHES } = require("./data.js");

const ACCENT = "1F4B7A", INK = "15181D", INK2 = "4B535E", INK3 = "6E7784";
const RULE = "C9CFD8", CLAY = "8F5238", SOFT = "EDF2F8", CLAYSOFT = "F7F0EC", ZEBRA = "F5F7F9";
const SERIF = "Cambria", SANS = "Calibri";
const W = 9638; // largeur utile en DXA (A4, marges 2 cm)

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = c => ({ style: BorderStyle.SINGLE, size: 4, color: c });

// ---------- helpers de paragraphe ----------
const p = (text, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, before: o.before ?? 0, line: 276 },
  alignment: o.align,
  indent: o.indent,
  border: o.border,
  shading: o.shading,
  children: [new TextRun({
    text, font: o.font ?? SANS, size: o.size ?? 20,
    color: o.color ?? INK, bold: o.bold, italics: o.italics,
    allCaps: o.caps, characterSpacing: o.caps ? 20 : undefined
  })]
});

const h1 = t => new Paragraph({
  heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: INK } },
  children: [new TextRun({ text: t, font: SERIF, size: 32, bold: true, color: INK })]
});
const h2 = t => new Paragraph({
  heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 140 },
  children: [new TextRun({ text: t, font: SERIF, size: 26, bold: true, color: ACCENT })]
});
const h3 = t => new Paragraph({
  heading: HeadingLevel.HEADING_3, spacing: { before: 220, after: 100 },
  children: [new TextRun({ text: t, font: SANS, size: 21, bold: true, color: INK })]
});
const kicker = t => p(t, { caps: true, size: 15, color: INK3, bold: true, after: 80, before: 160 });
const lead = t => p(t, { color: INK2, size: 20, after: 160 });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

const bullets = arr => arr.map(t => new Paragraph({
  numbering: { reference: "puces", level: 0 },
  spacing: { after: 70, line: 264 },
  children: [new TextRun({ text: t, font: SANS, size: 20, color: INK })]
}));
const numbered = (arr, ref = "questions") => arr.map(t => new Paragraph({
  numbering: { reference: ref, level: 0 },
  spacing: { after: 90, line: 264 },
  children: [new TextRun({ text: t, font: SANS, size: 20, color: INK })]
}));

// lignes d'écriture manuscrite
const lines = n => Array.from({ length: n }, () => new Paragraph({
  spacing: { after: 0, before: 180 },
  border: { bottom: thin(RULE) },
  children: [new TextRun({ text: "", size: 20 })]
}));

const callout = (label, text, color = ACCENT, bg = SOFT) => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [W],
  borders: { top: noBorder, bottom: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder,
             left: { style: BorderStyle.SINGLE, size: 18, color } },
  rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: bg, color: "auto" },
    margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [new Paragraph({
      spacing: { after: 0, line: 264 },
      children: [
        ...(label ? [new TextRun({ text: label + " ", font: SANS, size: 20, bold: true, color: INK })] : []),
        new TextRun({ text, font: SANS, size: 20, color: INK2 })
      ]
    })]
  })] })]
});

// ---------- tableaux ----------
const cellP = (t, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 40, line: 252 },
  alignment: o.align,
  children: [new TextRun({ text: t, font: SANS, size: o.size ?? 18, bold: o.bold, color: o.color ?? INK, allCaps: o.caps })]
});

function table(widths, headers, rows, opts = {}) {
  const mk = (children, w, o = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    margins: { top: 90, bottom: 90, left: 110, right: 110 },
    verticalAlign: VerticalAlign.TOP,
    children
  });
  const rs = [];
  if (headers) rs.push(new TableRow({ tableHeader: true, children: headers.map((h, i) =>
    mk(String(h).split("\n").map(part => cellP(part, { bold: true, size: 15, caps: true, color: INK3,
         align: opts.center?.includes(i) ? AlignmentType.CENTER : undefined })),
       widths[i], { fill: ZEBRA })) }));
  rows.forEach((r, ri) => rs.push(new TableRow({ children: r.map((c, i) =>
    mk([cellP(c, { align: opts.center?.includes(i) ? AlignmentType.CENTER : undefined,
                   bold: opts.boldCol?.includes(i), color: opts.boldCol?.includes(i) ? INK : undefined })],
       widths[i], { fill: opts.zebra && ri % 2 === 1 ? ZEBRA : undefined })) })));
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    borders: { top: thin(RULE), bottom: thin(RULE), left: thin(RULE), right: thin(RULE),
               insideHorizontal: thin(RULE), insideVertical: thin(RULE) },
    rows: rs
  });
}

// tableau 3 colonnes de comportements
function levels(f) {
  const w = [3213, 3213, 3212];
  const head = [
    { t: "Doué pour la compétence", c: ACCENT, fill: SOFT },
    { t: "Compétence maîtrisée", c: INK2, fill: ZEBRA },
    { t: "Moins bien maîtrisée · posez des questions lorsque la personne", c: CLAY, fill: CLAYSOFT }
  ];
  const cols = [f.high, f.mid, f.low];
  const cell = (i) => new TableCell({
    width: { size: w[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: head[i].fill, color: "auto" },
    margins: { top: 100, bottom: 100, left: 110, right: 110 },
    children: [new Paragraph({
      spacing: { after: 100 },
      children: [new TextRun({ text: head[i].t, font: SANS, size: 16, bold: true, color: head[i].c })]
    })]
  });
  const body = (i) => new TableCell({
    width: { size: w[i], type: WidthType.DXA },
    margins: { top: 100, bottom: 100, left: 110, right: 110 },
    children: cols[i].map(t => new Paragraph({
      numbering: { reference: "puces", level: 0 },
      spacing: { after: 60, line: 240 },
      children: [new TextRun({ text: t, font: SANS, size: 17, color: INK2 })]
    }))
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: w,
    borders: { top: thin(RULE), bottom: thin(RULE), left: thin(RULE), right: thin(RULE),
               insideHorizontal: thin(RULE), insideVertical: thin(RULE) },
    rows: [new TableRow({ tableHeader: true, children: [cell(0), cell(1), cell(2)] }),
           new TableRow({ children: [body(0), body(1), body(2)] })]
  });
}

function probeTable(pairs) {
  const w = [3213, 3213, 3212];
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: w,
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
               insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [new TableRow({ children: pairs.map(([k, v], i) => new TableCell({
      width: { size: w[i], type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 140, right: 100 },
      borders: { left: { style: BorderStyle.SINGLE, size: 10, color: RULE },
                 top: noBorder, bottom: noBorder, right: noBorder },
      children: [
        p(k, { caps: true, size: 14, bold: true, color: ACCENT, after: 50 }),
        p(v, { size: 17, color: INK2, after: 0 })
      ]
    })) })]
  });
}

const scoreLine = () => table([4819, 4819],
  null,
  [["Candidat : ______________________________", "Note (1 à 4) : _______     1 ne répond pas · 2 partiellement · 3 répond · 4 dépasse"]],
  {});

// ---------- fiche compétence ----------
function fiche(f) {
  const out = [];
  out.push(new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    borders: { top: thin(RULE), bottom: thin(RULE), left: thin(RULE), right: thin(RULE),
               insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: SOFT, color: "auto" },
      margins: { top: 140, bottom: 140, left: 160, right: 160 },
      children: [
        p(f.tag, { caps: true, size: 14, bold: true, color: ACCENT, after: 60 }),
        new Paragraph({ spacing: { after: 70 }, children: [new TextRun({ text: f.name, font: SERIF, size: 28, bold: true, color: INK })] }),
        p(f.def, { size: 18, italics: true, color: INK2, after: 0 })
      ]
    })] })]
  }));
  out.push(p("", { after: 120 }));
  if (f.quals) { out.push(kicker("Ce que le poste exige")); out.push(...bullets(f.quals)); }
  if (f.high) { out.push(kicker("Comportements attendus")); out.push(levels(f)); out.push(p("", { after: 120 })); }
  out.push(kicker(f.qintro || "Posez l'une des quatre questions ci-dessous, puis approfondissez"));
  out.push(...numbered(f.q));
  out.push(kicker("Questions d'approfondissement"));
  out.push(probeTable(f.probes || PROBES));
  out.push(p("", { after: 100 }));
  if (f.tip) { out.push(callout("À l'intervieweur.", f.tip)); out.push(p("", { after: 100 })); }
  if (f.kill) { out.push(callout("", f.kill, CLAY, CLAYSOFT)); out.push(p("", { after: 100 })); }
  out.push(kicker("Notes prises au cours de l'entretien"));
  out.push(...lines(5));
  out.push(p("", { after: 160 }));
  out.push(scoreLine());
  out.push(pageBreak());
  return out;
}

const e3 = FICHES.filter(f => f.step === 3).sort((a, b) => a.order - b.order);
const e4 = FICHES.filter(f => f.step === 4).sort((a, b) => a.order - b.order);

// ---------- corps du document ----------
const body = [];

// Couverture
body.push(new Paragraph({ spacing: { before: 1200, after: 140 },
  children: [new TextRun({ text: "VICTIMES & PRÉJUDICES AVOCATS · GRENOBLE", font: SANS, size: 17, bold: true, color: ACCENT, characterSpacing: 30 })] }));
body.push(new Paragraph({ spacing: { after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: INK } },
  children: [new TextRun({ text: "Guide pour les entretiens", font: SERIF, size: 56, bold: true, color: INK })] }));
body.push(p("Avocat futur associé, droit du dommage corporel", { font: SERIF, size: 28, color: INK2, after: 300 }));
body.push(lead("Quatre étapes. Sept compétences du référentiel Korn Ferry Leadership Architect. Les définitions reprennent celles du référentiel. Les comportements observables et les questions ont été rédigés pour le contexte du cabinet."));
body.push(p("", { after: 600 }));
body.push(...lines(1).map(() => new Paragraph({ spacing: { before: 300, after: 0 }, border: { bottom: thin(RULE) },
  children: [new TextRun({ text: "Nom du candidat", font: SANS, size: 18, color: INK3 })] })));
body.push(new Paragraph({ spacing: { before: 300, after: 0 }, border: { bottom: thin(RULE) },
  children: [new TextRun({ text: "Intervieweur", font: SANS, size: 18, color: INK3 })] }));
body.push(new Paragraph({ spacing: { before: 300, after: 0 }, border: { bottom: thin(RULE) },
  children: [new TextRun({ text: "Date", font: SANS, size: 18, color: INK3 })] }));
body.push(p("", { after: 900 }));
body.push(callout("Document interne.", "Ne pas remettre au candidat."));
body.push(pageBreak());

// 1. Préparation
body.push(h1("1. Préparation et règles"));
body.push(h2("Qui conduit quoi"));
body.push(table([1900, 3100, 900, 3738],
  ["Étape", "Intervieweur", "Durée", "Évalué"],
  [["1 · téléphone", "Gregory Levy, directeur du développement", "30 min", "Pré-qualification et qualification technique consignée"],
   ["2 · valeurs", "Gregory Levy, directeur du développement", "60 min", "Expertise, humanité, justice, réparation"],
   ["3 · compétences métier", "Milène Baud, office manager", "80 min", "Gérer la complexité, Être orienté client, Être résilient, Développer les talents"],
   ["4 · compétences et reprise", "Hervé Gerbi, avocat fondateur", "90 min", "Expertise fonctionnelle, Courage, Réseaux, Esprit stratégique, projet de reprise"]],
  { center: [2], boldCol: [0], zebra: true }));
body.push(p("", { after: 160 }));

body.push(h2("Deux règles de notation"));
body.push(lead("Le dispositif ne compte qu'un intervieweur par étape. Un entretien noté par une seule personne mesure autant l'intervieweur que le candidat, et rien ne permet de faire la part des deux. Ces deux règles font partie du dispositif au même titre que la grille."));
body.push(h3("Règle 1 · calibration préalable"));
body.push(p("Une heure, les trois intervieweurs ensemble, en présence, avant le premier entretien d'étape 2."));
body.push(...bullets([
  "Prenez une compétence et ses trois colonnes de comportements. Chacun note un même cas fictif, seul, sur la grille. Comparez. Recommencez sur une seconde compétence.",
  "L'objectif n'est pas de tomber d'accord sur la note. Il est de découvrir où vos échelles divergent et pourquoi. Sans cette heure, un 3 de l'un ne vaut pas un 3 de l'autre, et la note globale additionne des unités différentes.",
  "Produisez un relevé d'une page : sur quelles compétences vos lectures divergeaient, et quel repère commun vous retenez. Ce relevé se relit avant chaque entretien."]));
body.push(h3("Règle 2 · second observateur sur les finalistes"));
body.push(p("À partir du moment où il ne reste que deux ou trois candidats. Pas avant, le coût ne le justifie pas."));
body.push(...bullets([
  "Gregory Levy observe l'étape 3 conduite par Milène Baud. Milène Baud observe l'étape 4 conduite par Hervé Gerbi. Les deux étapes se tenant le même jour, les deux personnes sont déjà au cabinet.",
  "L'observateur ne pose aucune question et n'intervient pas, y compris quand il aurait quelque chose à ajouter. Il note sur la même grille, seul.",
  "Les deux grilles se comparent après l'entretien, jamais pendant."]));
body.push(p("", { after: 120 }));
body.push(callout("Ce qu'on fait des écarts.", "Un écart d'un point se note et se range. Un écart de deux points ou plus ne se moyenne pas : il remonte tel quel à la réunion de décision, comme point de discussion. Deux personnes qui ont vu la même heure et en tirent 2 et 4 ont vu deux choses différentes, et c'est cette différence qui porte l'information."));
body.push(p("", { after: 120 }));
body.push(callout("Prévenez le candidat.", "Annoncez la présence d'un second intervieweur dans le message de convocation, avec son nom et sa fonction. Un observateur qui apparaît sans avoir été annoncé met le candidat en défiance pour une heure, et l'article L1221-8 du code du travail impose de toute façon d'informer le candidat des méthodes utilisées à son égard."));

body.push(h2("Avant l'entretien"));
body.push(...bullets([
  "Lisez le bloc que vous allez conduire. Familiarisez-vous avec les comportements observables, pas seulement avec les questions.",
  "Relisez le CV, la fiche de scoring établie au tri, et les notes de l'étape précédente.",
  "Relisez le relevé de calibration.",
  "Vérifiez que vous n'allez pas poser les mêmes questions que l'intervieweur précédent. Une compétence, un intervieweur."]));
body.push(h2("Pendant l'entretien"));
body.push(...bullets([
  "Respectez les durées du tableau ci-dessus. Une étape 4 écourtée est une décision d'association prise à l'aveugle.",
  "Ne reportez pas un entretien. Soyez à l'heure. Le candidat que vous voulez recruter a d'autres options.",
  "Consacrez les 5 à 10 premières minutes aux présentations. Votre rôle, votre parcours, le déroulé du processus, ce qui se passe après cet entretien. Demandez au candidat s'il a des questions avant de commencer.",
  "Si vous avez connaissance d'une situation de handicap nécessitant un aménagement, du processus ou du poste, abordez-la directement et proposez l'aménagement.",
  "Prenez des notes pendant, pas après. Notez ce que le candidat dit, pas votre interprétation. La note vient ensuite."]));
body.push(h3("Questions d'ouverture, au choix"));
body.push(...numbered([
  "Selon vous, en quoi consiste ce poste ?",
  "Qu'est-ce qui vous a fait répondre ?",
  "Au vu de ce que vous savez du poste, de quoi auriez-vous besoin pour y réussir ?"], "ouverture"));
body.push(h2("Après l'entretien"));
body.push(...bullets([
  "Demandez au candidat s'il souhaite revenir sur un point et s'il a des questions.",
  "Rappelez-lui quand il aura un retour. L'engagement du cabinet est de cinq jours ouvrés.",
  "Notez seul, avant d'en parler aux autres intervieweurs. Une note influencée ne vaut rien."]));
body.push(p("", { after: 140 }));
body.push(callout("Cherchez le comportement, pas l'opinion.", "Si le candidat répond « en général, je fais... », ramenez-le à un fait : donnez-moi une situation précise. Une question comportementale sans approfondissement produit un récit préparé. C'est l'approfondissement qui produit le signal."));
body.push(p("", { after: 120 }));
body.push(callout("Restez attentif aux biais.", "Vous avez tendance à mieux noter les personnes qui démontrent une compétence de la même façon que vous. Vous excluez alors des personnes qui communiquent, travaillent ou décident autrement, sans être moins compétentes."));
body.push(pageBreak());

// 2. Matrice
body.push(h1("2. Compétences par étape"));
body.push(table([3638, 1500, 1500, 1500, 1500],
  ["Ce qui est évalué", "E1 téléphone\nG. Levy · 30 min", "E2 valeurs\nG. Levy · 60 min", "E3 métier\nM. Baud · 80 min", "E4 association\nH. Gerbi · 90 min"],
  [["Pré-qualification", "×", "", "", ""],
   ["Qualification technique, réponses consignées", "×", "", "", ""],
   ["Adéquation aux valeurs du cabinet", "", "×", "", ""],
   ["Gérer la complexité", "", "", "×", ""],
   ["Être orienté client", "", "", "×", ""],
   ["Être résilient", "", "", "×", ""],
   ["Développer les talents", "", "", "×", ""],
   ["Expertise fonctionnelle", "", "", "", "×"],
   ["Faire preuve de courage", "", "", "", "×"],
   ["Créer des réseaux", "", "", "", "×"],
   ["Faire preuve d'esprit stratégique", "", "", "", "×"],
   ["Projet de reprise", "", "", "", "×"]],
  { center: [1, 2, 3, 4], zebra: true }));
body.push(p("", { after: 160 }));
body.push(callout("Pourquoi l'expertise fonctionnelle est évaluée à l'étape 4.", "C'est le seul bloc du guide qui exige de l'intervieweur qu'il maîtrise lui-même la matière. Savoir si un candidat lit correctement un rapport d'expertise médicale, s'il chiffre un poste de préjudice selon la bonne méthode, si son arbitrage entre transaction et contentieux tient, suppose de connaître la réponse. Aucune grille ne compense cette absence. Une note produite par quelqu'un qui ne peut pas juger du fond n'est pas une évaluation prudente, c'est une évaluation fausse habillée en méthode."));
body.push(p("", { after: 160 }));
body.push(p("Les étapes 3 et 4 se tiennent le même jour.", { color: INK2 }));
body.push(pageBreak());

// 3. Synthèse
body.push(h1("3. Grille de synthèse"));
body.push(h2("Système de notation"));
body.push(table([700, 2400, 6538], ["Note", "Libellé", "Définition"],
  [["4", "Dépasse les critères", "A montré des connaissances, compétences ou capacités qui dépassent les critères définis pour le poste."],
   ["3", "Répond aux critères", "A montré des connaissances, compétences ou capacités qui répondent entièrement aux critères définis."],
   ["2", "Répond partiellement", "A montré des connaissances, compétences ou capacités qui répondent en partie mais non en totalité."],
   ["1", "Ne répond pas", "A montré des connaissances, compétences ou capacités clairement en dessous des critères définis."]],
  { center: [0], boldCol: [1], zebra: true }));
body.push(h2("Report des notes"));
body.push(p("Candidat : _________________________________________________", { after: 160 }));
body.push(table([700, 2900, 1450, 900, 1750, 950, 988],
  ["Ét.", "Évalué", "Intervieweur", "Note", "Observateur (finalistes)", "Note", "Écart"],
  [["2", "Adéquation aux valeurs", "G. Levy", "___/4", "—", "—", "—"],
   ["3", "Gérer la complexité", "M. Baud", "___/4", "G. Levy", "___/4", "___"],
   ["3", "Être orienté client", "M. Baud", "___/4", "G. Levy", "___/4", "___"],
   ["3", "Être résilient", "M. Baud", "___/4", "G. Levy", "___/4", "___"],
   ["3", "Développer les talents", "M. Baud", "___/4", "G. Levy", "___/4", "___"],
   ["4", "Expertise fonctionnelle", "H. Gerbi", "___/4", "M. Baud", "___/4", "___"],
   ["4", "Faire preuve de courage", "H. Gerbi", "___/4", "M. Baud", "___/4", "___"],
   ["4", "Créer des réseaux", "H. Gerbi", "___/4", "M. Baud", "___/4", "___"],
   ["4", "Esprit stratégique", "H. Gerbi", "___/4", "M. Baud", "___/4", "___"]],
  { center: [0, 3, 5, 6], zebra: true }));
body.push(p("", { after: 140 }));
body.push(p("Note globale : ______ / 4", { bold: true, after: 140 }));
body.push(p("Un écart de deux points ou plus entre intervieweur et observateur ne se moyenne pas. Il remonte tel quel à la réunion de décision.", { color: INK2, size: 18 }));
body.push(h2("Règles de décision"));
body.push(...bullets([
  "Une note de 1 sur Expertise fonctionnelle, Être orienté client ou Créer des réseaux arrête le processus. Ces trois dimensions ne se rattrapent pas par la formation.",
  "Une note de 1 sur Développer les talents arrête également le processus, pour une autre raison : ce recrutement est une succession, et l'équipe restera après le départ du fondateur ou partira avec lui.",
  "Une note de 2 sur Faire preuve de courage est un signal sérieux pour une association. Une personne qui ne vous contredit pas pendant le recrutement ne vous contredira pas une fois associée, et c'est exactement ce dont vous aurez besoin.",
  "Une note de 2 sur Gérer la complexité ou Être résilient se discute si le reste est solide, et se traite par un plan d'accompagnement écrit."]));
body.push(h2("Recommandation de l'intervieweur"));
body.push(table([2409, 2409, 2410, 2410], null,
  [["Je ne soutiens pas", "Je soutiens partiellement", "Je soutiens", "Je soutiens sans réserve"]],
  { center: [0, 1, 2, 3] }));
body.push(p("", { after: 140 }));
body.push(kicker("Motif en une phrase"));
body.push(...lines(2));
body.push(pageBreak());

// ÉTAPE 1
body.push(h1("Étape 1 · entretien téléphonique"));
body.push(p("30 minutes · Gregory Levy, directeur du développement", { color: ACCENT, bold: true, size: 19 }));
body.push(lead("Vérifier les faits, poser les conditions, donner envie. Un candidat qui passe cette étape doit repartir avec l'envie de venir au cabinet."));
body.push(h2("Qualification"));
body.push(table([2700, 4438, 2500], ["Point à vérifier", "Question", "Réponse"],
  [["Situation actuelle", "Où exercez-vous, sous quel statut, depuis quand ?", ""],
   ["Pratique réelle", "Quelle part du dommage corporel dans votre activité ? En volume, pas en intérêt", ""],
   ["Contentieux", "Combien de fois avez-vous plaidé sur les douze derniers mois ?", ""],
   ["Motif du mouvement", "Qu'est-ce qui vous fait regarder ailleurs aujourd'hui ?", ""],
   ["Projet d'association", "Vous êtes-vous déjà projeté dans une association ? Qu'est-ce qui vous a arrêté ou fait avancer ?", ""],
   ["Statut souhaité", "Collaboration libérale ou salariat, une préférence, et pourquoi ?", ""],
   ["Mobilité", "Où vivez-vous, qu'implique concrètement une installation à Grenoble ?", ""],
   ["Disponibilité", "Quel préavis ? À partir de quand ?", ""],
   ["Rémunération", "Où vous situez-vous aujourd'hui, qu'attendez-vous ?", ""],
   ["Autres processus", "Engagé dans d'autres discussions ? À quel stade ?", ""]],
  { boldCol: [0], zebra: true }));
body.push(p("", { after: 160 }));
body.push(callout("Notez ce que le candidat demande.", "C'est le premier indicateur exploitable du processus. Un candidat qui interroge sur les termes de l'association, sur l'origine des dossiers ou sur l'organisation du cabinet se projette. Un candidat qui interroge d'abord sur les congés et le télétravail se projette aussi, sur autre chose."));
body.push(pageBreak());
body.push(h2("Qualification technique, réponses consignées"));
body.push(callout("Vous ne notez pas ces réponses, vous les consignez.", "L'expertise fonctionnelle n'est évaluée qu'à l'étape 4, par le seul intervieweur en mesure de la juger. Attendre la quatrième étape pour découvrir qu'un candidat ne sait pas chiffrer un préjudice coûte trois entretiens à trois personnes. Notez la réponse mot à mot, sans juger, sans reformuler, sans compléter les silences. Hervé Gerbi les lit et rend un avis avant l'étape 2. Dix minutes par candidat."));
body.push(p("", { after: 200 }));
[["1. Sur quels référentiels vous appuyez-vous pour chiffrer un préjudice ?", 3],
 ["2. Racontez-moi votre dernière expertise médicale. Quel était le point en discussion, et qu'avez-vous fait ?", 4],
 ["3. Sur votre dernier dossier transigé, qu'est-ce qui vous a fait conclure qu'il fallait transiger plutôt qu'aller au contentieux ?", 4]
].forEach(([q, n]) => { body.push(p(q, { bold: true, after: 60, before: 160 })); body.push(...lines(n)); });
body.push(p("", { after: 220 }));
body.push(p("Avis d'Hervé Gerbi :          poursuivre          /          poursuivre avec réserve          /          arrêter", { bold: true, after: 140 }));
body.push(kicker("Réserve ou motif"));
body.push(...lines(2));
body.push(h2("Points d'attention"));
body.push(...bullets([
  "Si le candidat n'a aucune pratique du dommage corporel ni de matière connexe, dites-le franchement et évaluez sa réaction. Vous cherchez quelqu'un qui mesure ce qu'il ignore.",
  "Vouloir s'associer sans avoir réfléchi aux conditions financières n'est pas disqualifiant. C'est un sujet de l'étape 4.",
  "Si la fourchette de rémunération est incompatible, dites-le maintenant. Pas au quatrième entretien."]));
body.push(h2("Décision"));
body.push(p("Poursuivre          /          Ne pas poursuivre", { bold: true, after: 140 }));
body.push(kicker("Motif, et points à transmettre aux intervieweurs suivants"));
body.push(...lines(3));
body.push(pageBreak());

// ÉTAPE 2
body.push(h1("Étape 2 · entretien valeurs"));
body.push(p("60 minutes · Gregory Levy, directeur du développement", { color: ACCENT, bold: true, size: 19 }));
body.push(lead("Expertise, humanité, justice, réparation. Choisissez quatre des cinq situations. Vous évaluez la façon dont le candidat arbitre, pas l'exactitude juridique de sa réponse."));
body.push(callout("Pourquoi cet entretien n'interroge pas les valeurs directement.", "Demander à un candidat ce que l'humanité signifie pour lui ne trie personne. Il a lu votre site. Il dira ce qu'il faut dire, et il le pensera sans doute. Ce qui produit du signal, c'est le conflit : une valeur ne se révèle que lorsqu'elle entre en concurrence avec une autre et qu'il faut trancher. Il n'y a pas de bonne réponse. Vous observez comment il pense, ce qu'il met en premier, et s'il tranche ou s'il esquive."));
body.push(p("", { after: 160 }));
body.push(kicker("Consigne au candidat"));
body.push(p("« Je vais vous soumettre quatre situations. Elles n'ont pas de bonne réponse. Ce sont des arbitrages que nous avons à faire ici, et je veux comprendre comment vous les faites. Prenez le temps. Vous pouvez me poser des questions sur la situation. »", { italics: true, color: INK2 }));
SITUATIONS.forEach(s => {
  body.push(p("Situation " + s.n, { caps: true, size: 14, bold: true, color: CLAY, before: 260, after: 50 }));
  body.push(new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: s.t, font: SERIF, size: 24, bold: true, color: INK })] }));
  body.push(p(s.c));
  body.push(p(s.a, { bold: true, color: ACCENT, font: SERIF, size: 21, after: 120 }));
  body.push(kicker("Ce que vous observez"));
  body.push(...bullets(s.o));
  body.push(...lines(3));
});
body.push(pageBreak());
body.push(h2("Notation de l'étape 2"));
body.push(table([2200, 6200, 1238], ["Valeur", "Ce qui a été observé", "1 à 4"],
  [["Expertise", "", ""], ["Humanité", "", ""], ["Justice", "", ""], ["Réparation", "", ""]],
  { boldCol: [0], center: [2], zebra: true }));
body.push(p("", { after: 160 }));
body.push(p("Adéquation globale aux valeurs du cabinet : ______ / 4", { bold: true, after: 160 }));
body.push(callout("", "Une question à se poser avant de noter : est-ce que je confierais à cette personne la cliente de la situation 1 ?          Oui  ·  Non  ·  Pas encore"));
body.push(pageBreak());

// ÉTAPE 3
body.push(h1("Étape 3 · compétences métier"));
body.push(p("80 minutes · Milène Baud, office manager", { color: ACCENT, bold: true, size: 19 }));
body.push(lead("Quatre compétences comportementales. La méthode y note l'approche décrite par le candidat, les étapes qu'il a suivies et ce qu'il en a tiré, pas l'exactitude juridique de sa conclusion."));
body.push(pageBreak());
e3.forEach(f => body.push(...fiche(f)));

// ÉTAPE 4
body.push(h1("Étape 4 · compétences et projet de reprise"));
body.push(p("90 minutes · Hervé Gerbi, avocat fondateur", { color: ACCENT, bold: true, size: 19 }));
body.push(lead("L'étape la plus lourde du dispositif. Elle ouvre sur l'expertise fonctionnelle, seul bloc du guide qui exige de l'intervieweur qu'il maîtrise la matière. Aucun document n'est remis au candidat au cours de cet entretien."));
body.push(pageBreak());
e4.forEach(f => body.push(...fiche(f)));

// Projet de reprise
body.push(h2("Projet de reprise"));
body.push(p("Séquence non notée · 20 minutes · Hervé Gerbi", { color: ACCENT, bold: true, size: 19 }));
body.push(kicker("Cadre de la séquence"));
body.push(p("Aucun document n'est remis au candidat à ce stade. Le protocole d'association sera rédigé plus tard. Cette séquence est donc une conversation, et c'est vous qui portez la crédibilité du projet. Dites-le au candidat en ouverture, plutôt que de le laisser le déduire : une chose annoncée n'est pas la même qu'une chose découverte."));
body.push(kicker("Questions"));
body.push(...numbered([
  "Qu'est-ce qui vous attire dans une reprise plutôt que dans une association ?",
  "Avez-vous déjà dirigé quelque chose ? Quoi, sur quelle durée, avec combien de personnes ?",
  "Comment envisagez-vous de financer une reprise ?",
  "Le cabinet porte le nom de son fondateur. Qu'en feriez-vous ?",
  "Qu'attendez-vous de moi pendant la transition, concrètement ?",
  "Qu'est-ce qui vous ferait partir d'ici au bout de deux ans ?"], "assoc"));
body.push(kicker("Ce que vous observez"));
body.push(...bullets([
  "Question 1. Elle sépare deux populations : ceux qui veulent une bonne pratique dans une bonne structure, et ceux qui veulent diriger. Les deux sont respectables, une seule vous intéresse.",
  "Question 3. Elle n'attend pas un plan de financement. Un candidat qui n'a jamais pensé à l'argent ne s'est pas projeté. Un candidat qui pose des questions précises s'est projeté.",
  "Question 4. Celle que personne ne prépare. La réponse dit s'il compte exister en son nom ou vivre dans une ombre. Les deux réponses sont recevables, l'absence de réponse ne l'est pas.",
  "Question 5. La plus utile de l'entretien. Elle oblige le candidat à dire ce dont il a besoin de vous. Un candidat qui ne demande rien n'a pas réfléchi à la transition, ou n'ose pas vous le dire.",
  "Question 6. Elle vous dit ce que vous aurez à tenir. Une réponse vague signale que la personne n'a pas de critère, et donc rien qui la retienne."]));
body.push(p("", { after: 160 }));
body.push(kicker("Réserves exprimées par le candidat, à traiter avant toute proposition"));
body.push(...lines(4));
body.push(p("", { after: 160 }));
body.push(p("Le candidat se projette dans la reprise :          oui          /          avec réserves          /          non", { bold: true }));
body.push(pageBreak());

// 5. Références
body.push(h1("5. Prise de références"));
body.push(lead("Deux références minimum, avec l'accord écrit du candidat. Privilégiez un associé ou employeur précédent et un contact professionnel extérieur, confrère, expert ou magistrat. N'interrogez jamais l'employeur actuel sans accord explicite et écrit."));
body.push(...numbered([
  "Dans quel cadre avez-vous travaillé ensemble, et sur quelle période ?",
  "Quels types de dossiers lui confiiez-vous, et lesquels ne lui confiiez-vous pas ?",
  "Comment se comportait-il avec les clients difficiles ?",
  "Racontez-moi une fois où il vous a contredit.",
  "Comment a-t-il réagi à un échec professionnel ?",
  "Sur quoi avait-il besoin de progresser quand vous avez cessé de travailler ensemble ?",
  "Le reprendriez-vous ? Sur quel poste ?"], "refs"));
body.push(p("", { after: 140 }));
body.push(callout("", "La question 4 est la plus discriminante. Si la personne interrogée ne trouve aucun exemple, c'est une information sur le candidat."));
body.push(p("", { after: 160 }));
body.push(kicker("Notes"));
body.push(...lines(6));
body.push(pageBreak());

// 6. Décision
body.push(h1("6. Décision finale"));
body.push(lead("Réunion de décision avec tous les intervieweurs. Chacun arrive avec ses notes écrites."));
body.push(callout("Règle de conduite.", "Chacun donne sa note et sa recommandation avant toute discussion. On discute ensuite. Une réunion qui commence par l'avis du plus ancien ne produit que son avis."));
body.push(p("", { after: 160 }));
body.push(table([4000, 5638], ["Point", "Constat"],
  [["Notes reportées", ""],
   ["Écarts de 2 points ou plus entre intervieweur et observateur, et ce que chacun a vu", ""],
   ["Règles de décision déclenchées, notes de 1 ou de 2", ""],
   ["Références concordantes", ""],
   ["Réserves exprimées par le candidat", ""],
   ["Plan d'accompagnement nécessaire et temps à y consacrer", ""]],
  { boldCol: [0], zebra: true }));
body.push(p("", { after: 200 }));
body.push(p("Décision :          proposition          /          deuxième candidat prioritaire          /          pas de suite", { bold: true, after: 160 }));
body.push(kicker("Si proposition, engagements pris auprès du candidat à formaliser par écrit"));
body.push(...lines(4));

// ---------- assemblage ----------
const doc = new Document({
  creator: "Victimes & Préjudices Avocats",
  title: "Guide pour les entretiens — Avocat futur associé",
  description: "Guide d'entretien structuré, quatre étapes, six compétences Korn Ferry Leadership Architect.",
  numbering: { config: [
    { reference: "puces", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 340, hanging: 200 } }, run: { color: ACCENT } } }] },
    ...["questions", "ouverture", "assoc", "refs"].map(ref => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 280 } }, run: { color: ACCENT, bold: true } } }] }))
  ] },
  styles: { default: { document: { run: { font: SANS, size: 20, color: INK } } } },
  sections: [{
    properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT, spacing: { after: 0 },
      border: { bottom: thin(RULE) },
      children: [new TextRun({ text: "Guide pour les entretiens · Avocat futur associé · Document interne", font: SANS, size: 15, color: INK3 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT, spacing: { before: 60 },
      children: [new TextRun({ children: ["Page ", PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES], font: SANS, size: 15, color: INK3 })] })] }) },
    children: body
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(process.argv[2], b);
  console.log("Écrit :", process.argv[2], (b.length / 1024).toFixed(0) + " Ko");
});
