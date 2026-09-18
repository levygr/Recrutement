# Régénération des fichiers Office

Scripts de génération du guide Word et de la présentation PowerPoint.
Les fichiers produits sont à la racine du dépôt. Toute correction se fait ici,
jamais dans le `.docx` ou le `.pptx` directement, sinon elle est perdue à la
prochaine génération.

```bash
npm install docx pptxgenjs
node make-docx.js ../Guide-entretiens.docx
node make-pptx.js ../Plan-de-recrutement-presentation.pptx
```

`data.js` porte le contenu des sept blocs de compétences et des cinq mises en
situation de l'entretien valeurs. C'est le seul fichier à modifier pour changer
une question, un comportement observable ou une définition.

La source de vérité du contenu reste `04-guide-entretiens.md`. Les scripts en
sont une mise en forme, pas une version parallèle : toute modification de fond
se répercute dans les deux.
