const pptxgen = require("pptxgenjs");
const P = new pptxgen();
P.layout = "LAYOUT_WIDE";            // 13.333 x 7.5
P.author = "Gregory Levy"; P.company = "Pinede";
P.title = "Plan de recrutement - Avocat futur associe";

// ---- palette ----
const INK="12212E", BLUE="1F4B7A", CLAY="B05F3A", MUTED="5A6673";
const S1="123E6B", S2="2E6DA8", S3="7FA9D0";       // rampe sequentielle, luminance croissante
const TINT="F2F5F8", WHITE="FFFFFF", ONDARK="D6DEE7", CLAYL="E0956B", RULE="D9E0E7";
const SERIF="Cambria", SANS="Calibri";

const W=13.333, H=7.5, M=0.65, CW=W-2*M;   // 12.033

// ---- helpers ----
function titled(s, title, kicker){
  if(kicker) s.addText(kicker.toUpperCase(), {x:M,y:0.44,w:CW,h:0.26,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:11,bold:true,color:CLAY,charSpacing:2});
  s.addText(title, {x:M,y:kicker?0.72:0.55,w:CW,h:0.85,isTextBox:true,margin:0,valign:"top",
    fontFace:SERIF,fontSize:33,bold:true,color:INK});
}
function darkSlide(){ const s=P.addSlide(); s.background={color:INK}; return s; }
function lightSlide(){ const s=P.addSlide(); s.background={color:WHITE}; return s; }
function card(s,{x,y,w,h,fill=TINT,line=null}){ s.addShape(P.ShapeType.roundRect,{x,y,w,h,
  rectRadius:0.06,fill:{color:fill},line:line?{color:line,width:1}:{type:"none"}}); }
function chip(s,x,y,d,label,fill=BLUE,color=WHITE){
  s.addShape(P.ShapeType.ellipse,{x,y,w:d,h:d,fill:{color:fill},line:{type:"none"}});
  s.addText(label,{x,y,w:d,h:d,isTextBox:true,margin:0,align:"center",valign:"middle",
    fontFace:SANS,fontSize:13,bold:true,color});
}
function foot(s,t,dark=false){ s.addText(t,{x:M,y:6.98,w:CW,h:0.3,isTextBox:true,margin:0,
  fontFace:SANS,fontSize:9.5,color:dark?"7C8896":MUTED}); }
const body=(s,t,o)=>s.addText(t,Object.assign({isTextBox:true,margin:0,fontFace:SANS,
  fontSize:14,color:INK,valign:"top"},o));

// =========================================================== 1. TITRE
{ const s=darkSlide();
  s.addShape(P.ShapeType.rect,{x:0,y:0,w:0.0001,h:0.0001,fill:{color:INK},line:{type:"none"}});
  s.addText("VICTIMES & PRÉJUDICES AVOCATS", {x:M,y:1.55,w:CW,h:0.3,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:12,bold:true,color:CLAYL,charSpacing:3});
  s.addText("Recruter un futur associé",{x:M,y:2.0,w:11.2,h:1.5,isTextBox:true,margin:0,valign:"top",
    fontFace:SERIF,fontSize:52,bold:true,color:WHITE});
  s.addText("Plan de recrutement · La succession d'Hervé Gerbi · Grenoble",
    {x:M,y:3.5,w:11.2,h:0.45,isTextBox:true,margin:0,fontFace:SANS,fontSize:17,color:ONDARK});
  s.addShape(P.ShapeType.line,{x:M,y:4.3,w:2.4,h:0,line:{color:CLAYL,width:2}});
  s.addText([{text:"Pour Hervé Gerbi et Milène Baud",options:{breakLine:true,bold:true,color:WHITE}},
             {text:"Présenté par Gregory Levy · Septembre 2026",options:{color:ONDARK}}],
    {x:M,y:4.6,w:6,h:0.9,isTextBox:true,margin:0,fontFace:SANS,fontSize:14,lineSpacing:22});
  s.addNotes("Objectif de la séance : que vous compreniez pourquoi ce recrutement est difficile, ce que le dispositif prévoit, et ce que chacun de vous a à faire. Trois décisions sont à prendre aujourd'hui, elles sont en dernière slide.");
}

// =========================================================== 2. LE POSTE
{ const s=lightSlide(); titled(s,"Nous ne recrutons pas un associé","Le poste");
  card(s,{x:M,y:1.78,w:CW,h:1.15,fill:INK});
  s.addText("Nous recrutons la personne qui remplacera Hervé Gerbi et reprendra le cabinet.",
    {x:M+0.45,y:1.98,w:CW-0.9,h:0.86,isTextBox:true,margin:0,valign:"top",
     fontFace:SERIF,fontSize:23,bold:true,color:WHITE,lineSpacing:29});
  const items=[["Une succession, pas une association","Statut salarié à l'entrée, puis association, puis reprise. C'est l'objet du recrutement, pas une option."],
    ["6 à 15 ans d'exercice","Un avocat de quarante ans qui veut enfin diriger sa structure est un candidat parfaitement cohérent."],
    ["Dommage corporel, ou matière proche","La pratique de la matière ne change pas seulement le niveau. Elle change la date de reprise possible."],
    ["Capable d'apporter et de diriger","Un successeur reprend une clientèle et une équipe, pas seulement un portefeuille de dossiers."]];
  let y=3.1;
  items.forEach((it,i)=>{ card(s,{x:M,y,w:CW,h:0.86});
    chip(s,M+0.28,y+0.19,0.48,String(i+1));
    body(s,it[0],{x:M+0.95,y:y+0.1,w:4.2,h:0.32,fontSize:14.5,bold:true,color:BLUE});
    body(s,it[1],{x:M+5.3,y:y+0.1,w:CW-5.6,h:0.66,fontSize:12,color:MUTED,lineSpacing:16});
    y+=0.98; });
  foot(s,"Détail : document 02, offre d'emploi.");
  s.addNotes("Tout le dispositif découle de là. C'est aussi notre meilleur argument sur le marché : « devenez associé » est une promesse banale, « reprenez un cabinet » ne l'est pas.");
}

// =========================================================== 2bis. CERTIFICAT DE SPÉCIALISATION
{ const s=lightSlide(); titled(s,"Le certificat fixe la date de reprise","Une contrainte de calendrier");
  body(s,"Le positionnement du cabinet repose en partie sur les certificats de spécialisation d'Hervé Gerbi. Le jour où il cesse d'exercer, le cabinet perd le droit d'utiliser ce titre si personne d'autre ne le détient.",
    {x:M,y:1.8,w:CW,h:0.62,fontSize:14.5,color:MUTED,lineSpacing:21});
  const cw=(CW-0.4)/2;
  card(s,{x:M,y:2.55,w:cw,h:1.7});
  s.addText("Ce qu'il faut pour l'obtenir",{x:M+0.3,y:2.72,w:cw-0.6,h:0.3,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:13.5,bold:true,color:BLUE});
  s.addText([{text:"Quatre années de pratique professionnelle dans le domaine de spécialisation.",options:{breakLine:true}},
    {text:"Puis un entretien de validation devant un jury, quarante minutes, en séance publique, avec mise en situation.",options:{}}],
    {x:M+0.3,y:3.08,w:cw-0.6,h:1.05,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:MUTED,lineSpacing:18});
  card(s,{x:M+cw+0.4,y:2.55,w:cw,h:1.7,fill:"EAF1E9"});
  s.addText("La bonne nouvelle",{x:M+cw+0.7,y:2.72,w:cw-0.6,h:0.3,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:13.5,bold:true,color:"2C5F2D"});
  s.addText("Les années passées chez nous comme avocat salarié comptent dans les quatre ans. Le compteur tourne dès la prise de poste. Le salariat ne retarde rien.",
    {x:M+cw+0.7,y:3.08,w:cw-0.6,h:1.05,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:INK,lineSpacing:18});
  body(s,"Ce que cela donne selon le profil recruté",{x:M,y:4.42,w:CW,h:0.28,fontSize:13,bold:true,color:BLUE});
  const rows=[["Déjà titulaire du certificat","Immédiat","Sans contrainte",S1],
    ["4 ans ou plus de dommage corporel","Dès la 1re année","24 à 36 mois",S1],
    ["2 ans de dommage corporel","Après 2 ans chez nous","36 à 48 mois",S2],
    ["Aucune pratique de la matière","Après 4 ans chez nous","60 mois et plus",CLAY]];
  let yy=4.74;
  s.addText("PROFIL À L'ENTRÉE",{x:M+0.2,y:yy,w:4.2,h:0.24,isTextBox:true,margin:0,fontFace:SANS,fontSize:9.5,bold:true,color:MUTED,charSpacing:1.2});
  s.addText("CERTIFICAT PRÉSENTABLE",{x:M+4.7,y:yy,w:3.2,h:0.24,isTextBox:true,margin:0,fontFace:SANS,fontSize:9.5,bold:true,color:MUTED,charSpacing:1.2});
  s.addText("REPRISE ENVISAGEABLE",{x:M+8.2,y:yy,w:3.5,h:0.24,isTextBox:true,margin:0,fontFace:SANS,fontSize:9.5,bold:true,color:MUTED,charSpacing:1.2});
  yy+=0.28;
  rows.forEach(r=>{ card(s,{x:M,y:yy,w:CW,h:0.40,fill:r[3]===CLAY?"FBF1EC":TINT});
    s.addText(r[0],{x:M+0.2,y:yy,w:4.4,h:0.40,isTextBox:true,margin:0,valign:"middle",fontFace:SANS,fontSize:12,color:INK});
    s.addText(r[1],{x:M+4.7,y:yy,w:3.4,h:0.40,isTextBox:true,margin:0,valign:"middle",fontFace:SANS,fontSize:12,color:MUTED});
    s.addText(r[2],{x:M+8.2,y:yy,w:3.6,h:0.40,isTextBox:true,margin:0,valign:"middle",fontFace:SANS,fontSize:12,bold:true,color:r[3]===CLAY?CLAY:BLUE});
    yy+=0.46; });
  foot(s,"Décret du 27 novembre 1991 et arrêté du 15 décembre 2023. Détail : document 01, point 0.1.");
  s.addNotes("Hervé, c'est ce tableau qui doit guider votre arbitrage de sélection. Un candidat sans pratique de la matière n'est pas disqualifié, mais il décale la reprise de deux à trois ans. Autant le décider en connaissance de cause plutôt que de le découvrir après.");
}

// =========================================================== 3. POURQUOI C'EST DIFFICILE
{ const s=lightSlide(); titled(s,"Pourquoi ce recrutement est difficile","Le marché, en quatre chiffres");
  const stats=[["77 600","avocats en France","Chiffres du ministère de la Justice"],
    ["2,5 %","exercent sous statut salarié","Le statut que nous proposons aujourd'hui"],
    ["79","spécialistes du dommage corporel au barreau de Paris","Sur plus de 33 000 avocats inscrits"],
    ["~600","avocats au barreau de Grenoble","Toutes matières confondues"]];
  const w=(CW-3*0.3)/4;
  stats.forEach((st,i)=>{ const x=M+i*(w+0.3);
    const hot=(i===1||i===2);
    card(s,{x,y:1.9,w,h:2.55,fill:hot?"FBF1EC":TINT});
    s.addText(st[0],{x:x+0.25,y:2.05,w:w-0.5,h:0.95,isTextBox:true,margin:0,valign:"middle",
      fontFace:SERIF,fontSize:44,bold:true,color:hot?CLAY:BLUE});
    s.addText(st[1],{x:x+0.25,y:3.02,w:w-0.5,h:0.8,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:13,bold:true,color:INK});
    s.addText(st[2],{x:x+0.25,y:3.82,w:w-0.5,h:0.5,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:10.5,color:MUTED});
  });
  card(s,{x:M,y:4.75,w:CW,h:1.55,fill:INK});
  s.addText("Croisez les deux chiffres en orange.",{x:M+0.4,y:4.98,w:CW-0.8,h:0.38,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:15,bold:true,color:CLAYL});
  s.addText("Le nombre de spécialistes du dommage corporel, disponibles, prêts au salariat et intéressés par Grenoble tient sur une main. Ce vivier ne lit pas les annonces, parce qu'il n'est pas en recherche.",
    {x:M+0.4,y:5.4,w:CW-0.8,h:0.75,isTextBox:true,margin:0,fontFace:SANS,fontSize:14,color:ONDARK,lineSpacing:20});
  foot(s,"Sources détaillées : fichier SOURCES.md du dossier de recrutement.");
  s.addNotes("C'est la slide la plus importante. Si le vivier est aussi étroit, publier une annonce et attendre ne produira rien. Deux leviers : élargir le statut proposé, et aller chercher les gens un par un.");
}

// =========================================================== 4. CE QUE CES CHIFFRES IMPOSENT
{ const s=lightSlide(); titled(s,"Ce que ces chiffres imposent","La stratégie");
  body(s,"L'annonce sert à ne pas passer à côté d'un candidat déjà en mouvement, et à donner un point d'atterrissage à notre approche directe. Elle ne produira pas notre futur associé.",
    {x:M,y:1.8,w:CW,h:0.7,fontSize:15,color:MUTED,lineSpacing:22});
  const parts=[[70,"Approche directe","Nous écrivons nous-mêmes, nominativement, à des avocats identifiés",S1],
               [20,"Réseau et prescription","Barreaux, confrères, médecins de recours, associations",S2],
               [10,"Diffusion d'annonce","Job boards spécialisés et généralistes",S3]];
  let x=M; const total=CW;
  parts.forEach(pt=>{ const w=total*pt[0]/100;
    s.addShape(P.ShapeType.rect,{x,y:2.75,w:w-0.06,h:1.0,fill:{color:pt[3]},line:{type:"none"}});
    s.addText(pt[0]+" %",{x:x+0.16,y:2.9,w:w-0.28,h:0.7,isTextBox:true,margin:0,valign:"middle",
      fontFace:SERIF,fontSize:w<1.8?22:30,bold:true,color:WHITE});
    s.addText(pt[1],{x,y:3.86,w:w-0.1,h:0.52,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:14,bold:true,color:INK,lineSpacing:18});
    s.addText(pt[2],{x,y:4.42,w:w-0.2,h:0.85,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:11.5,color:MUTED,lineSpacing:16});
    x+=w; });
  card(s,{x:M,y:5.42,w:CW,h:1.2,fill:"FBF1EC"});
  s.addText("Ce qui compense le statut salarié",{x:M+0.35,y:5.58,w:6,h:0.3,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:13,bold:true,color:CLAY});
  s.addText("Le salariat représente 2,5 % de la profession : nous demandons au candidat d'accepter un statut minoritaire. La perspective de reprise est ce qui le compense, à condition d'être dite clairement dès l'annonce.",
    {x:M+0.35,y:5.9,w:CW-0.7,h:0.6,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:13,color:INK,lineSpacing:18});
  foot(s,"Détail : document 05, plan de diffusion.");
  s.addNotes("Répartition de l'effort, pas du budget. L'approche directe coûte du temps, pas d'argent.");
}

// =========================================================== 5. LE FICHIER DE CHASSE
{ const s=lightSlide(); titled(s,"Notre meilleur fichier est déjà ici","Approche directe");
  card(s,{x:M,y:1.8,w:CW,h:1.75,fill:INK});
  s.addText("Les avocats qu'Hervé Gerbi a affrontés en expertise et à l'audience depuis dix ans.",
    {x:M+0.45,y:1.98,w:CW-0.9,h:0.78,isTextBox:true,margin:0,valign:"top",
     fontFace:SERIF,fontSize:21,bold:true,color:WHITE,lineSpacing:27});
  s.addText("Il sait déjà lesquels sont bons, lesquels tiennent un dossier, lesquels parlent correctement aux victimes d'en face. Une heure de son temps pour dresser cette liste vaut plus que tout le budget de diffusion.",
    {x:M+0.45,y:2.74,w:CW-0.9,h:0.68,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:13.5,color:ONDARK,lineSpacing:19});
  body(s,"Les autres sources, toutes publiques et gratuites",{x:M,y:3.75,w:CW,h:0.3,fontSize:13,bold:true,color:BLUE});
  const src=[["Plateforme de spécialisation du CNB","L'annuaire officiel des titulaires du certificat en dommage corporel, filtrable par région"],
    ["ANADAVI","L'association ne liste que les titulaires de la spécialité. Le vivier le plus dense du quart sud-est"],
    ["Annuaires des barreaux","Lyon, Chambéry, Annecy, Valence, Saint-Étienne, Clermont-Ferrand"],
    ["LinkedIn","Le seul canal qui montre l'ancienneté au poste, donc le moment favorable"]];
  const cw=(CW-0.3)/2; let yy=4.15;
  src.forEach((sc,i)=>{ const x=M+(i%2)*(cw+0.3), y=yy+Math.floor(i/2)*1.2;
    card(s,{x,y,w:cw,h:1.08});
    s.addText(sc[0],{x:x+0.25,y:y+0.16,w:cw-0.5,h:0.3,isTextBox:true,margin:0,fontFace:SANS,fontSize:13.5,bold:true,color:INK});
    s.addText(sc[1],{x:x+0.25,y:y+0.5,w:cw-0.5,h:0.52,isTextBox:true,margin:0,fontFace:SANS,fontSize:11.5,color:MUTED,lineSpacing:16});
  });
  foot(s,"Objectif : 60 à 80 noms qualifiés. Deux à trois jours de travail.");
  s.addNotes("Hervé : c'est la demande la plus importante que je vous ferai. Une heure, une liste de noms.");
}

// =========================================================== 6. SECTION
{ const s=darkSlide();
  s.addText("01",{x:M,y:2.3,w:2,h:1.1,isTextBox:true,margin:0,fontFace:SERIF,fontSize:64,bold:true,color:CLAYL});
  s.addText("Ce qui bloque aujourd'hui",{x:M,y:3.45,w:10.5,h:0.9,isTextBox:true,margin:0,
    fontFace:SERIF,fontSize:40,bold:true,color:WHITE});
  s.addText("Rien ne part en diffusion tant que ces points ne sont pas tranchés.",
    {x:M,y:4.4,w:9.5,h:0.5,isTextBox:true,margin:0,fontFace:SANS,fontSize:16,color:ONDARK});
  s.addNotes("Une annonce qui promet une association sans en dire les termes attire des candidats qui partiront au troisième entretien, quand ils poseront les vraies questions.");
}

// =========================================================== 7. DECISIONS
{ const s=lightSlide(); titled(s,"Ce qui est tranché, ce qui reste","Arbitrages");
  card(s,{x:M,y:1.85,w:CW,h:1.35,fill:"FBF1EC"});
  chip(s,M+0.35,1.98,0.55,"!",CLAY);
  s.addText("Une seule décision bloque la diffusion",{x:M+1.1,y:2.02,w:7.5,h:0.36,isTextBox:true,margin:0,
    fontFace:SERIF,fontSize:21,bold:true,color:INK});
  s.addText("La fourchette de rémunération à afficher dans l'annonce. Sur un marché de candidats passifs, son absence est une friction gratuite : elle fait perdre les candidats qui n'osent pas demander, et fait perdre du temps sur ceux qui sont hors budget.",
    {x:M+1.1,y:2.45,w:CW-1.5,h:0.65,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:13,color:MUTED,lineSpacing:18});
  const cw=(CW-0.4)/2;
  card(s,{x:M,y:3.45,w:cw,h:2.65,fill:TINT});
  s.addText("Déjà tranché",{x:M+0.3,y:3.65,w:cw-0.6,h:0.32,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:14,bold:true,color:BLUE});
  s.addText([{text:"Le statut : avocat salarié à l'entrée.",options:{breakLine:true}},
    {text:"Le protocole d'association : rédigé plus tard par Hervé et Milène.",options:{breakLine:true}},
    {text:"Les intervieweurs et la répartition des étapes.",options:{breakLine:true}},
    {text:"Les deux règles de notation.",options:{}}],
    {x:M+0.3,y:4.05,w:cw-0.6,h:1.9,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:INK,lineSpacing:24});
  card(s,{x:M+cw+0.4,y:3.45,w:cw,h:2.65,fill:TINT});
  s.addText("Ouvert, non bloquant",{x:M+cw+0.7,y:3.65,w:cw-0.6,h:0.32,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:14,bold:true,color:MUTED});
  s.addText([{text:"L'avenir de l'enseigne, qui porte le nom du fondateur.",options:{breakLine:true}},
    {text:"Le périmètre : Grenoble seul ou Grenoble et Annecy.",options:{breakLine:true}},
    {text:"La place du droit du travail.",options:{breakLine:true}},
    {text:"L'ordre des entretiens, la date de prise de poste, la vidéo.",options:{}}],
    {x:M+cw+0.7,y:4.05,w:cw-0.6,h:1.9,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:MUTED,lineSpacing:24});
  foot(s,"Détail : document 01, point 2.");
  s.addNotes("Une seule décision retient la diffusion. Les quatre autres peuvent se traiter pendant que le sourcing tourne.");
}

// =========================================================== 8. PROTOCOLE
// =========================================================== 9. SECTION
{ const s=darkSlide();
  s.addText("02",{x:M,y:2.3,w:2,h:1.1,isTextBox:true,margin:0,fontFace:SERIF,fontSize:64,bold:true,color:CLAYL});
  s.addText("Le dispositif d'évaluation",{x:M,y:3.45,w:10.5,h:0.9,isTextBox:true,margin:0,
    fontFace:SERIF,fontSize:40,bold:true,color:WHITE});
  s.addText("Quatre étapes, six compétences, et le rôle de chacun.",
    {x:M,y:4.4,w:9.5,h:0.5,isTextBox:true,margin:0,fontFace:SANS,fontSize:16,color:ONDARK});
  s.addNotes("À partir d'ici, c'est votre partie à tous les deux.");
}

// =========================================================== 10. LES 4 ETAPES
{ const s=lightSlide(); titled(s,"Quatre étapes, un intervieweur par étape","Le parcours candidat");
  body(s,"Tri des CV sur 100 points, seuil à 70. Puis :",{x:M,y:1.78,w:CW,h:0.3,fontSize:13,color:MUTED});
  const st=[["1","Téléphone","Gregory Levy","30 min","Pré-qualification. Trois questions techniques consignées mot à mot, relues par Hervé Gerbi."],
    ["2","Valeurs","Gregory Levy","55 min","Expertise, humanité, justice, réparation. Quatre séquences de douze minutes, sur des situations générales sans prérequis sur la matière."],
    ["3","Compétences métier","Milène Baud","80 min","Gérer la complexité. Être orienté client. Être résilient. Développer les talents."],
    ["4","Compétences et reprise","Hervé Gerbi","90 min","Expertise fonctionnelle. Courage. Réseaux. Esprit stratégique. Projet de reprise."]];
  const cw=(CW-3*0.25)/4;
  st.forEach((e,i)=>{ const x=M+i*(cw+0.25);
    card(s,{x,y:2.2,w:cw,h:3.5,fill:i===3?INK:TINT});
    chip(s,x+0.25,y=2.42,0.5,e[0],i===3?CLAY:BLUE);
    s.addText(e[1],{x:x+0.25,y:3.1,w:cw-0.5,h:0.62,isTextBox:true,margin:0,valign:"top",
      fontFace:SERIF,fontSize:17,bold:true,color:i===3?WHITE:INK});
    s.addText(e[2],{x:x+0.25,y:3.78,w:cw-0.5,h:0.28,isTextBox:true,margin:0,
      fontFace:SANS,fontSize:12,bold:true,color:i===3?CLAYL:BLUE});
    s.addText(e[3],{x:x+0.25,y:4.06,w:cw-0.5,h:0.26,isTextBox:true,margin:0,
      fontFace:SANS,fontSize:11,color:i===3?"9AA7B4":MUTED});
    s.addText(e[4],{x:x+0.25,y:4.42,w:cw-0.5,h:1.15,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:11,color:i===3?ONDARK:MUTED,lineSpacing:15});
  });
  card(s,{x:M,y:5.9,w:CW,h:0.75,fill:"FBF1EC"});
  s.addText("Les étapes 3 et 4 se tiennent le même jour. Nous répondons au candidat sous cinq jours ouvrés après chaque étape.",
    {x:M+0.35,y:6.08,w:CW-0.7,h:0.42,isTextBox:true,margin:0,valign:"middle",fontFace:SANS,fontSize:13,color:INK});
  foot(s,"Détail : document 04, guide pour les entretiens.");
  s.addNotes("Le candidat consacre environ quatre heures au processus. C'est défendable pour une association, à condition de tenir les délais de réponse.");
}

// =========================================================== 11. LES 6 COMPETENCES
{ const s=lightSlide(); titled(s,"Sept compétences, et ce que chacune protège","Référentiel Korn Ferry Leadership Architect");
  const e3=[["Gérer la complexité","Le cœur technique du dossier corporel : pièces médicales, expertises contradictoires, chiffrage",false],
    ["Être orienté client","La relation avec la victime et ses proches, qui est notre positionnement même",true],
    ["Être résilient","La durée des dossiers et la charge émotionnelle, première cause de décrochage",false],
    ["Développer les talents","L'équipe, au moment où le fondateur se retire",true]];
  const e4=[["Faire preuve de courage","La capacité à contredire un client, un expert, et son futur associé",false],
    ["Créer des réseaux","L'apport d'affaires, condition économique réelle de la reprise",true],
    ["Faire preuve d'esprit stratégique","La direction du cabinet, pas seulement son propre portefeuille",false]];
  s.addText("ÉTAPE 3 · MILÈNE BAUD",{x:M,y:1.82,w:5,h:0.24,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:9.5,bold:true,color:BLUE,charSpacing:1.5});
  const w4=(CW-3*0.24)/4;
  e3.forEach((it,i)=>{ const x=M+i*(w4+0.24), y=2.12;
    card(s,{x,y,w:w4,h:1.85});
    s.addText(it[0],{x:x+0.22,y:y+0.18,w:w4-0.44,h:0.72,isTextBox:true,margin:0,valign:"top",
      fontFace:SERIF,fontSize:16,bold:true,color:INK});
    s.addText(it[1],{x:x+0.22,y:y+0.95,w:w4-0.44,h:0.78,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:11,color:MUTED,lineSpacing:14});
    if(it[2]) s.addShape(P.ShapeType.ellipse,{x:x+w4-0.4,y:y+0.2,w:0.18,h:0.18,fill:{color:CLAY},line:{type:"none"}});
  });
  s.addText("ÉTAPE 4 · HERVÉ GERBI",{x:M,y:4.18,w:5,h:0.24,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:9.5,bold:true,color:BLUE,charSpacing:1.5});
  const w3=(CW-2*0.28)/3;
  e4.forEach((it,i)=>{ const x=M+i*(w3+0.28), y=4.48;
    card(s,{x,y,w:w3,h:1.55});
    s.addText(it[0],{x:x+0.25,y:y+0.18,w:w3-0.5,h:0.62,isTextBox:true,margin:0,valign:"top",
      fontFace:SERIF,fontSize:16.5,bold:true,color:INK});
    s.addText(it[1],{x:x+0.25,y:y+0.82,w:w3-0.5,h:0.62,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:11,color:MUTED,lineSpacing:14});
    if(it[2]) s.addShape(P.ShapeType.ellipse,{x:x+w3-0.45,y:y+0.2,w:0.18,h:0.18,fill:{color:CLAY},line:{type:"none"}});
  });
  s.addShape(P.ShapeType.ellipse,{x:M,y:6.26,w:0.16,h:0.16,fill:{color:CLAY},line:{type:"none"}});
  s.addText("Une note de 1 sur ces compétences, ou sur le bloc Expertise fonctionnelle, arrête le processus.",
    {x:M+0.3,y:6.18,w:CW-0.3,h:0.3,isTextBox:true,margin:0,fontFace:SANS,fontSize:12.5,bold:true,color:INK});
  foot(s,"Sept compétences retenues sur les trente-huit du référentiel. Développer les talents a été ajoutée après l'arbitrage sur la succession.");
  s.addNotes("La septième a été ajoutée parce qu'une succession n'est pas une association. Le risque propre à une succession n'est pas que le successeur soit mauvais juriste. C'est que l'équipe parte dans les dix-huit mois qui suivent le départ du fondateur.");
}

// =========================================================== 12. EXPERTISE A L'ETAPE 4
{ const s=lightSlide(); titled(s,"Pourquoi l'expertise technique est évaluée par Hervé","Une correction du dispositif");
  const cw=(CW-0.4)/2;
  card(s,{x:M,y:1.9,w:cw,h:2.5,fill:TINT});
  s.addText("Ce que Milène peut évaluer",{x:M+0.3,y:2.12,w:cw-0.6,h:0.34,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:14,bold:true,color:BLUE});
  s.addText([{text:"Les trois compétences de l'étape 3 sont comportementales.",options:{breakLine:true}},
    {text:"La méthode note l'approche décrite par le candidat, les étapes qu'il a suivies et ce qu'il en a tiré. Pas l'exactitude juridique de sa conclusion.",options:{breakLine:true}},
    {text:"Sur Être orienté client, l'office manager est même le meilleur observateur du cabinet : elle voit chaque jour comment les clients sont réellement traités.",options:{}}],
    {x:M+0.3,y:2.52,w:cw-0.6,h:1.75,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:MUTED,lineSpacing:18});
  card(s,{x:M+cw+0.4,y:1.9,w:cw,h:2.5,fill:"FBF1EC"});
  s.addText("Ce qu'elle ne peut pas évaluer",{x:M+cw+0.7,y:2.12,w:cw-0.6,h:0.34,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:14,bold:true,color:CLAY});
  s.addText([{text:"L'expertise fonctionnelle est le seul bloc qui exige de l'intervieweur qu'il maîtrise la matière.",options:{breakLine:true}},
    {text:"Savoir si un candidat lit correctement un rapport d'expertise médicale, s'il chiffre un poste de préjudice selon la bonne méthode, si son arbitrage entre transaction et contentieux tient : cela suppose de connaître la réponse.",options:{}}],
    {x:M+cw+0.7,y:2.52,w:cw-0.6,h:1.75,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:MUTED,lineSpacing:18});
  card(s,{x:M,y:4.65,w:CW,h:1.15,fill:INK});
  s.addText("Une note produite par quelqu'un qui ne peut pas juger du fond n'est pas une évaluation prudente. C'est une évaluation fausse habillée en méthode, et c'est plus dangereux qu'une absence de note.",
    {x:M+0.45,y:4.9,w:CW-0.9,h:0.7,isTextBox:true,margin:0,valign:"top",fontFace:SERIF,fontSize:16,bold:true,color:WHITE,lineSpacing:24});
  body(s,"Le bloc est donc passé à l'étape 4. L'étape 4 monte à 90 minutes, l'étape 3 descend à 60. Et l'étape 1 gagne trois questions techniques consignées mot à mot, qu'Hervé relit en dix minutes, pour ne pas découvrir un problème technique au quatrième entretien.",
    {x:M,y:6.0,w:CW,h:0.7,fontSize:12.5,color:MUTED,lineSpacing:18});
  s.addNotes("Milène, ce n'est pas un jugement sur vous. C'est une règle de méthode : on n'évalue pas une compétence qu'on ne possède pas soi-même.");
}

// =========================================================== 13. HERVE
{ const s=lightSlide(); titled(s,"Hervé, ce que le plan attend de vous","Rôles");
  const t=[["Avant le lancement","Trancher la fourchette de rémunération à afficher dans l'annonce.","30 minutes"],
    ["Semaine 1","Dresser la liste des avocats que vous avez affrontés et que vous jugez bons. Signer les courriers aux bâtonniers.","1 heure + signatures"],
    ["À chaque candidat","Relire les trois réponses techniques consignées à l'étape 1 et rendre un avis.","10 minutes"],
    ["Semaine 4","Une heure de calibration avec Gregory et Milène sur la grille de notation.","1 heure"],
    ["Étape 4","Conduire l'entretien de 90 minutes. Contredire délibérément le candidat une fois, sur un point où il a raison, et regarder ce qu'il fait.","90 min par finaliste"]];
  let y=1.88;
  t.forEach(it=>{ card(s,{x:M,y,w:CW,h:0.74});
    s.addText(it[0],{x:M+0.3,y:y+0.1,w:2.6,h:0.3,isTextBox:true,margin:0,fontFace:SANS,fontSize:12.5,bold:true,color:BLUE});
    s.addText(it[1],{x:M+3.1,y:y+0.1,w:CW-5.1,h:0.56,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:11.5,color:INK,lineSpacing:15});
    s.addText(it[2],{x:M+CW-1.95,y:y+0.1,w:1.7,h:0.3,isTextBox:true,margin:0,align:"right",
      fontFace:SANS,fontSize:10.5,bold:true,color:CLAY});
    y+=0.85; });
  foot(s,"Vous ne rencontrez le candidat qu'une fois, avant de décider de lui confier le cabinet. Les notes des trois étapes précédentes vous parviennent avant, par écrit.");
  s.addNotes("Total de votre temps hors entretiens : environ quatre heures sur dix semaines. La liste des confrères que vous avez affrontés est de loin la plus utile.");
}

// =========================================================== 14. MILENE
{ const s=lightSlide(); titled(s,"Milène, ce que le plan attend de vous","Rôles");
  const t=[["Semaine 4","Une heure de calibration avec Gregory et Hervé sur la grille de notation.","1 heure"],
    ["Étape 3","Conduire l'entretien de 80 minutes sur quatre compétences : gérer la complexité, être orienté client, être résilient, développer les talents.","80 min par candidat"],
    ["Étape 4, finalistes","Assister en observatrice silencieuse à l'entretien d'Hervé. Ne poser aucune question. Noter seule.","90 min par finaliste"],
    ["Réunion de décision","Donner votre note et votre recommandation avant toute discussion.","1 heure"],
    ["Plus tard","Rédiger le protocole d'association détaillé avec Hervé Gerbi.","À planifier"]];
  let y=1.85;
  t.forEach(it=>{ card(s,{x:M,y,w:CW,h:0.74});
    s.addText(it[0],{x:M+0.3,y:y+0.1,w:2.6,h:0.3,isTextBox:true,margin:0,fontFace:SANS,fontSize:12.5,bold:true,color:BLUE});
    s.addText(it[1],{x:M+3.1,y:y+0.1,w:CW-5.1,h:0.56,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:11.5,color:INK,lineSpacing:15});
    s.addText(it[2],{x:M+CW-1.95,y:y+0.1,w:1.7,h:0.3,isTextBox:true,margin:0,align:"right",
      fontFace:SANS,fontSize:10.5,bold:true,color:CLAY});
    y+=0.82; });
  card(s,{x:M,y:6.0,w:CW,h:0.9,fill:INK});
  s.addText("Sur Développer les talents, la question que vous avez réellement à trancher n'est pas « cette personne sait-elle faire grandir ses collaborateurs ». C'est « est-ce que je resterais si cette personne dirigeait le cabinet ». Vous êtes la seule à pouvoir y répondre, et votre réponse vaut plus que toutes les autres notes.",
    {x:M+0.4,y:6.14,w:CW-0.8,h:0.66,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:ONDARK,lineSpacing:17});
  foot(s,"Le guide d'entretien vous est remis imprimé avant le premier candidat.");
  s.addNotes("Milène, vous n'avez pas à juger le droit. Vous notez ce que le candidat a fait, comment il s'y est pris, et ce qu'il en a tiré. Sur Être orienté client et sur Développer les talents, vous êtes la mieux placée du cabinet.");
}

// =========================================================== 15. REGLES DE NOTATION
{ const s=lightSlide(); titled(s,"Deux règles qui rendent la grille fiable","Notation");
  body(s,"Le dispositif ne compte qu'un intervieweur par étape. Un entretien noté par une seule personne mesure autant l'intervieweur que le candidat, et rien ne permet de faire la part des deux.",
    {x:M,y:1.8,w:CW,h:0.6,fontSize:14,color:MUTED,lineSpacing:20});
  const cw=(CW-0.4)/2;
  const r=[["Calibration préalable","Une heure, tous les trois ensemble, avant le premier entretien d'étape 3. Chacun note un même cas fictif, seul, puis on compare.",
    "L'objectif n'est pas de tomber d'accord sur la note. Il est de voir où nos échelles divergent. Sans cette heure, un 3 de Milène ne vaut pas un 3 d'Hervé.","Semaine 4"],
   ["Second observateur","Sur les deux ou trois finalistes seulement. Gregory observe l'étape 3, Milène observe l'étape 4. Aucune question, notation indépendante.",
    "Un écart d'un point se range. Un écart de deux points ou plus ne se moyenne pas : il remonte tel quel à la réunion de décision.","Semaines 7 et 8"]];
  r.forEach((it,i)=>{ const x=M+i*(cw+0.4);
    card(s,{x,y:2.5,w:cw,h:3.3});
    chip(s,x+0.3,2.75,0.5,String(i+1));
    s.addText(it[0],{x:x+0.3,y:3.42,w:cw-0.6,h:0.38,isTextBox:true,margin:0,fontFace:SERIF,fontSize:20,bold:true,color:INK});
    s.addText(it[1],{x:x+0.3,y:3.88,w:cw-0.6,h:0.85,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:INK,lineSpacing:17});
    s.addText(it[2],{x:x+0.3,y:4.75,w:cw-0.6,h:0.85,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:11.5,color:MUTED,lineSpacing:16});
    s.addText(it[3],{x:x+cw-1.85,y:2.8,w:1.6,h:0.3,isTextBox:true,margin:0,align:"right",fontFace:SANS,fontSize:11,bold:true,color:CLAY});
  });
  card(s,{x:M,y:6.0,w:CW,h:0.8,fill:"FBF1EC"});
  s.addText("Le second intervieweur est annoncé au candidat dans le message de convocation, avec son nom et sa fonction. C'est une obligation de l'article L1221-8 du code du travail, et c'est aussi la correction élémentaire.",
    {x:M+0.35,y:6.18,w:CW-0.7,h:0.48,isTextBox:true,margin:0,valign:"middle",fontFace:SANS,fontSize:12.5,color:INK});
  foot(s,"Détail : document 04, section 1.");
  s.addNotes("Ces deux règles coûtent une heure chacun, plus deux heures par finaliste. C'est le prix d'une grille qui veut dire quelque chose.");
}

// =========================================================== 16. CALENDRIER
{ const s=lightSlide(); titled(s,"Dix à quatorze semaines jusqu'à la signature","Calendrier");
  const ph=[["S0","Arbitrages et protocole d'association","Hervé, Gregory"],
    ["S1","Publication, courriers aux bâtonniers, liste de chasse","Gregory"],
    ["S2-S4","Trois cycles d'approche directe, 60 à 80 messages","Gregory"],
    ["S4","Heure de calibration sur la grille","Tous les trois"],
    ["S3-S6","Tri des CV et entretiens téléphoniques au fil de l'eau","Gregory"],
    ["S5-S8","Entretiens valeurs, métier et reprise","Tous les trois"],
    ["S9","Prise de références et décision","Hervé, Gregory"],
    ["S10","Proposition et signature","Hervé"]];
  const bw=(CW-7*0.12)/8;
  ph.forEach((pp,i)=>{ const x=M+i*(bw+0.12);
    const on=(i===3||i===5);
    s.addShape(P.ShapeType.rect,{x,y:2.0,w:bw,h:0.42,fill:{color:on?CLAY:S1},line:{type:"none"}});
    s.addText(pp[0],{x,y:2.0,w:bw,h:0.42,isTextBox:true,margin:0,align:"center",valign:"middle",
      fontFace:SANS,fontSize:13,bold:true,color:WHITE});
    s.addText(pp[1],{x,y:2.58,w:bw,h:1.5,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:11,color:INK,lineSpacing:15});
    s.addText(pp[2],{x,y:4.1,w:bw,h:0.5,isTextBox:true,margin:0,valign:"top",
      fontFace:SANS,fontSize:10,bold:true,color:on?CLAY:MUTED});
  });
  card(s,{x:M,y:4.85,w:CW,h:1.5,fill:TINT});
  s.addText("Puis le préavis du candidat",{x:M+0.35,y:5.05,w:5,h:0.3,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:13,bold:true,color:BLUE});
  s.addText("Trois mois en général pour un avocat salarié, selon contrat pour un collaborateur libéral. Prise de poste réaliste : quatre à six mois après le lancement. L'annonce actuelle dit « dès que possible », ce qui ne veut rien dire. Nous la remplacerons par une date.",
    {x:M+0.35,y:5.4,w:CW-0.7,h:0.8,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:12.5,color:MUTED,lineSpacing:18});
  foot(s,"Les deux jalons en orange vous concernent tous les deux.");
  s.addNotes("Si nous décidons aujourd'hui, la prise de poste se situe autour de février ou mars 2027.");
}

// =========================================================== 17. TUNNEL
{ const s=lightSlide(); titled(s,"Ce que nous attendons, et quand s'inquiéter","Pilotage");
  const fn=[["CV reçus","25 à 40",1.0],["CV au-dessus du seuil de 70","8 à 12",0.55],
    ["Entretiens téléphoniques","12 à 18",0.62],["Candidats en étape 3","4 à 6",0.34],
    ["Candidats en étape 4","2 à 3",0.22],["Propositions émises","1 à 2",0.14]];
  let y=1.95; const maxw=7.0;
  fn.forEach((f)=>{
    s.addText(f[0],{x:M,y:y+0.03,w:3.1,h:0.3,isTextBox:true,margin:0,fontFace:SANS,fontSize:12.5,color:INK});
    s.addShape(P.ShapeType.rect,{x:M+3.3,y,w:maxw*f[2],h:0.36,fill:{color:S2},line:{type:"none"}});
    s.addText(f[1],{x:M+3.45+maxw*f[2],y,w:1.5,h:0.36,isTextBox:true,margin:0,valign:"middle",
      fontFace:SANS,fontSize:12.5,bold:true,color:BLUE});
    y+=0.5; });
  const cw=(CW-0.4)/2;
  [["Moins de 5 CV au-dessus du seuil en semaine 4","Le problème vient de l'offre, pas de la diffusion. Nous revoyons le statut et la fourchette."],
   ["Moins de 10 % de réponse à l'approche directe","Le problème vient du message. Nous le réécrivons avant d'envoyer la suite."]]
  .forEach((a,i)=>{ const x=M+i*(cw+0.4);
    card(s,{x,y:5.2,w:cw,h:1.35,fill:"FBF1EC"});
    s.addText("ALERTE",{x:x+0.3,y:5.38,w:2,h:0.24,isTextBox:true,margin:0,fontFace:SANS,fontSize:9.5,bold:true,color:CLAY,charSpacing:1.5});
    s.addText(a[0],{x:x+0.3,y:5.63,w:cw-0.6,h:0.32,isTextBox:true,margin:0,fontFace:SANS,fontSize:13,bold:true,color:INK});
    s.addText(a[1],{x:x+0.3,y:5.97,w:cw-0.6,h:0.5,isTextBox:true,margin:0,valign:"top",fontFace:SANS,fontSize:11.5,color:MUTED,lineSpacing:16});
  });
  foot(s,"Suivi hebdomadaire par Gregory Levy.");
  s.addNotes("Le tunnel se lit de haut en bas. Si une étape s'écarte de la cible, on corrige à cet endroit précis plutôt que de tout relancer.");
}

// =========================================================== 18. CLOTURE
{ const s=darkSlide();
  s.addText("CE QUE NOUS DÉCIDONS AUJOURD'HUI",{x:M,y:1.3,w:CW,h:0.32,isTextBox:true,margin:0,
    fontFace:SANS,fontSize:12,bold:true,color:CLAYL,charSpacing:3});
  const d=[["La fourchette de rémunération","Le montant que nous affichons dans l'annonce. C'est la seule décision qui retient la diffusion"],
    ["La liste des confrères","Hervé dresse la liste des avocats qu'il a affrontés et qu'il juge bons. Une heure"],
    ["La date de la séance de calibration","Une heure, tous les trois ensemble, avant le premier entretien d'étape 3"]];
  let y=1.9;
  d.forEach((it,i)=>{
    s.addShape(P.ShapeType.roundRect,{x:M,y,w:CW,h:1.15,rectRadius:0.06,
      fill:{color:"1B2E40"},line:{type:"none"}});
    chip(s,M+0.35,y+0.33,0.5,String(i+1),CLAY);
    s.addText(it[0],{x:M+1.1,y:y+0.22,w:CW-1.5,h:0.36,isTextBox:true,margin:0,fontFace:SERIF,fontSize:20,bold:true,color:WHITE});
    s.addText(it[1],{x:M+1.1,y:y+0.63,w:CW-1.5,h:0.35,isTextBox:true,margin:0,fontFace:SANS,fontSize:13,color:ONDARK});
    y+=1.3; });
  s.addText("Tout le reste est écrit et prêt. Le dossier complet compte huit documents : plan, annonce, grille de tri des CV, guide d'entretien, plan de diffusion, concept vidéo, sources. Le protocole d'association détaillé sera rédigé plus tard par Hervé et Milène.",
    {x:M,y:6.0,w:CW,h:0.7,isTextBox:true,margin:0,fontFace:SANS,fontSize:13,color:"9AA7B4",lineSpacing:19});
  s.addNotes("Sans la fourchette, rien ne part. Les deux autres se calent dans la foulée.");
}

P.writeFile({ fileName: process.argv[2] }).then(f=>console.log("Écrit :", f));
