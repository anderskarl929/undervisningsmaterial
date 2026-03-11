const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require('docx');

// --- Shared config ---
const FONT = "Arial";
const PAGE_WIDTH = 11906; // A4
const PAGE_HEIGHT = 16838;
const MARGIN = 1440; // 1 inch
const CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN; // 9026 DXA

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

const headerBg = { fill: "1B4F72", type: ShadingType.CLEAR };
const altRowBg = { fill: "F2F7FA", type: ShadingType.CLEAR };

function bulletConfig() {
  return {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  };
}

function styles() {
  return {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: FONT, color: "1B4F72" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: "2C3E50" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 }
      },
    ]
  };
}

function footerSection(kurs, moment) {
  return new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: `${kurs} \u2014 ${moment}`, font: FONT, size: 18, color: "999999" }),
        new TextRun({ text: "\tSida ", font: FONT, size: 18, color: "999999" }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "999999" }),
      ],
      tabStops: [{ type: "right", position: CONTENT_WIDTH }],
    })]
  });
}

function headerPara(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: FONT, size: 24, ...opts })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: FONT, size: 24 })]
  });
}

function boldBullet(label, text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [
      new TextRun({ text: label, font: FONT, size: 24, bold: true }),
      new TextRun({ text: ` ${text}`, font: FONT, size: 24 }),
    ]
  });
}

function timeTable(rows, colWidths) {
  const cw = colWidths || [900, 1200, 1800, 5126];
  const headerTexts = ["Tid", "Fas", "Aktivitet", "Beskrivning"];

  const headerRow = new TableRow({
    children: headerTexts.map((t, i) => new TableCell({
      borders, shading: headerBg, width: { size: cw[i], type: WidthType.DXA }, margins: cellMargins,
      children: [new Paragraph({ children: [new TextRun({ text: t, font: FONT, size: 22, bold: true, color: "FFFFFF" })] })]
    }))
  });

  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders, width: { size: cw[ci], type: WidthType.DXA }, margins: cellMargins,
      shading: ri % 2 === 1 ? altRowBg : undefined,
      children: [new Paragraph({ children: [new TextRun({ text: cell, font: FONT, size: 22 })] })]
    }))
  }));

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: cw,
    rows: [headerRow, ...dataRows]
  });
}

// ============================================================
// LEKTION 2
// ============================================================
function lektion2() {
  return new Document({
    numbering: bulletConfig(),
    styles: styles(),
    sections: [{
      properties: {
        page: { size: { width: PAGE_WIDTH, height: PAGE_HEIGHT }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } }
      },
      footers: { default: footerSection("Samh\u00e4llskunskap 1a1", "Ungas ekonomi") },
      children: [
        headerPara('Lektion 2: "Vem f\u00e5r jobben?" \u2014 Arbetsmarknaden f\u00f6r unga'),
        para("Kurs: Samh\u00e4llskunskap 1a1 | Moment: Ungas ekonomi | L\u00e4ngd: 60 minuter", { italics: true, color: "666666" }),

        h2("L\u00e4randem\u00e5l f\u00f6r lektionen"),
        bullet("M\u00e5l 2: Analysera arbetsmarknadens villkor f\u00f6r unga \u2014 arbetsmarknadens f\u00f6r\u00e4ndringar och villkor med fokus p\u00e5 ungas situation"),

        h2("F\u00f6rberedelse"),
        bullet("F\u00f6rbered kort statistikblad: ungdomsarbetsl\u00f6shet i Sverige (SCB-data) [VERIFIERA aktuella siffror]"),
        bullet("F\u00f6rbered j\u00e4mf\u00f6relsetabell: anst\u00e4llningsformer (fast, visstid, gig/frilans)"),
        bullet("Skriv ut Alex tre jobbdilemma \u2014 ett per par"),
        bullet("Presentation om arbetsmarknaden"),

        h2("Tidsplanering"),
        timeTable([
          ["0\u20135 min", "Uppstart", "Retrieval practice", '"Skriv ner tre saker du minns om det ekonomiska kretsloppet." 2 min enskilt, 2 min i par, 1 min genomg\u00e5ng.'],
          ["5\u201318 min", "Instruktion", "Arbetsmarknaden f\u00f6r unga", "Utbud och efterfr\u00e5gan p\u00e5 arbetskraft. Anst\u00e4llningsformer (fast, visstid, tim-, gig). Ungas situation: h\u00f6gre arbetsl\u00f6shet, fler os\u00e4kra anst\u00e4llningar. Alex s\u00f6ker jobb."],
          ["18\u201348 min", "Bearbetning", "J\u00e4mf\u00f6relse + EPA", "Del 1 (12 min): J\u00e4mf\u00f6relsetabell i par. Del 2 (18 min): EPA \u2014 Alex tre jobbdilemma."],
          ["48\u201355 min", "Summering", "Snabbskrivning", '"Alex kompis s\u00e4ger: Gig-jobb \u00e4r framtiden. Alex mamma s\u00e4ger: Ta lagerjobbet. Vem har r\u00e4tt?"'],
          ["55\u201360 min", "Koppling fram\u00e5t", "N\u00e4sta g\u00e5ng", '"N\u00e4sta g\u00e5ng ska Alex flytta hemifr\u00e5n \u2014 vi r\u00e4knar p\u00e5 om pengarna r\u00e4cker."'],
        ]),

        new Paragraph({ children: [new PageBreak()] }),

        h2("Alex tre jobbdilemma"),
        para("Alex har f\u00e5tt erbjudande om tre jobb. Inget \u00e4r perfekt \u2014 alla har f\u00f6rdelar och nackdelar:"),
        boldBullet("1. Fast anst\u00e4llning p\u00e5 lager \u2014", "25 500 kr/m\u00e5n brutto, nattskift, 40 min pendling, heltid. Tryggt men d\u00e5liga arbetstider och tungt fysiskt."),
        boldBullet("2. Timanst\u00e4llning i butik \u2014", "ca 18 000 kr/m\u00e5n (om Alex f\u00e5r tillr\u00e4ckligt med timmar), n\u00e4ra hemmet, trevlig arbetsmilj\u00f6, men inga garanterade timmar. Chefen sms:ar kv\u00e4llen innan."),
        boldBullet("3. Frilansare via plattform (Foodora/Uber) \u2014", "helt flexibelt, potentiellt 20 000+ kr/m\u00e5n, men ingen semester, ingen sjukpenning, egen skatteinbetalning, slitage p\u00e5 egen cykel/bil."),
        para("EPA-fr\u00e5ga: Vilket jobb ska Alex v\u00e4lja \u2014 och varf\u00f6r? T\u00e4nk p\u00e5 trygghet, l\u00f6n, h\u00e4lsa, frihet och framtid.", { bold: true }),

        h2("L\u00e4rarinstruktioner"),
        boldBullet("Uppstart:", "Retrieval practice kopplat till lektion 1. Bekr\u00e4fta att eleverna minns akt\u00f6rerna och fl\u00f6dena."),
        boldBullet("Instruktion:", "Anv\u00e4nd exempel som eleverna k\u00e4nner igen (McDonald\u2019s, Foodora, butiksjobb). Var tydlig med att anst\u00e4llningsformer har olika konsekvenser."),
        boldBullet("Bearbetning:", "Pusha elever att motivera sitt val med kopplingar till trygghet, ekonomi OCH livskvalitet. Ingen av de tre anst\u00e4llningarna ska framst\u00e5 som r\u00e4tt svar."),
        boldBullet("Summering:", "Perspektivfr\u00e5gan (kompisen vs. mamman) \u00e4r medvetet polariserande \u2014 l\u00e5t eleverna argumentera."),

        h2("Differentiering"),
        boldBullet("St\u00f6d (mot E):", 'J\u00e4mf\u00f6relsetabellen har en kolumn redan ifylld som modell. St\u00f6dfr\u00e5gor vid jobbvalet: "Lista tre saker som \u00e4r bra med varje jobb. Vilket v\u00e4ger tyngst f\u00f6r dig?"'),
        boldBullet("Utmaning (mot A):", '"Diskutera hur gig-ekonomins tillv\u00e4xt p\u00e5verkar samh\u00e4llsekonomin i stort \u2014 vad h\u00e4nder med skatteint\u00e4kterna och v\u00e4lf\u00e4rden om fler jobbar som frilansare? Vem tj\u00e4nar p\u00e5 gig-ekonomin?"'),

        h2("Koppling till kunskapskrav"),
        boldBullet("E-niv\u00e5:", "Beskriva olika anst\u00e4llningsformer och ge enkla resonemang om f\u00f6rdelar/nackdelar."),
        boldBullet("C-niv\u00e5:", "J\u00e4mf\u00f6ra anst\u00e4llningsformer utf\u00f6rligt och ge v\u00e4lgrundade resonemang om konsekvenser."),
        boldBullet("A-niv\u00e5:", "Analysera arbetsmarknadens villkor ur flera perspektiv (individ, arbetsgivare, samh\u00e4lle) och visa p\u00e5 komplexa samband med samh\u00e4llsekonomin."),
      ]
    }]
  });
}

// ============================================================
// LEKTION 3
// ============================================================
function lektion3() {
  return new Document({
    numbering: bulletConfig(),
    styles: styles(),
    sections: [{
      properties: {
        page: { size: { width: PAGE_WIDTH, height: PAGE_HEIGHT }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } }
      },
      footers: { default: footerSection("Samh\u00e4llskunskap 1a1", "Ungas ekonomi") },
      children: [
        headerPara('Lektion 3: "R\u00e4cker pengarna?" \u2014 Privatekonomi och budget'),
        para("Kurs: Samh\u00e4llskunskap 1a1 | Moment: Ungas ekonomi | L\u00e4ngd: 60 minuter", { italics: true, color: "666666" }),

        h2("L\u00e4randem\u00e5l f\u00f6r lektionen"),
        bullet("M\u00e5l 3: Resonera om privatekonomiska val och deras konsekvenser"),

        h2("F\u00f6rberedelse"),
        bullet("F\u00f6rbered presentation om privatekonomiska grunder (brutto/netto, skatt, fasta/r\u00f6rliga utgifter)"),
        bullet("Skriv ut de tre livsstilsbeskrivningarna med utgiftsposter \u2014 tv\u00e5 per par"),
        bullet("Miniwhiteboard eller papper f\u00f6r utr\u00e4kningar"),

        h2("Tidsplanering"),
        timeTable([
          ["0\u20135 min", "Uppstart", "Retrieval practice", '"Vilka tre anst\u00e4llningsformer pratade vi om f\u00f6rra g\u00e5ngen? N\u00e4mn en f\u00f6rdel och en nackdel med varje."'],
          ["5\u201318 min", "Instruktion", "Privatekonomi-grunder", "Bruttol\u00f6n vs. nettol\u00f6n, skatt, arbetsgivaravgift. Fasta vs. r\u00f6rliga utgifter. Sparande och buffertsparande. Visa konkret: Alex f\u00e5r 22 000 kr i bruttol\u00f6n \u2014 hur mycket blir kvar?"],
          ["18\u201350 min", "Bearbetning", "Tre livsstilar", "I par: budgetera f\u00f6r TV\u00c5 av tre livsstilar. J\u00e4mf\u00f6r och argumentera: vilken \u00e4r smartast \u2014 och f\u00f6r vem?"],
          ["50\u201357 min", "Summering", "EPA \u2014 ansvarsfr\u00e5gan", '"\u00c4r det Alex eget fel om hen hamnar i ekonomiska problem \u2014 eller handlar det om vilken l\u00f6n arbetsmarknaden erbjuder unga?"'],
          ["57\u201360 min", "Koppling fram\u00e5t", "N\u00e4sta g\u00e5ng", '"N\u00e4sta g\u00e5ng st\u00e4ller vi Alex inf\u00f6r sv\u00e5ra val \u2014 vad h\u00e4nder n\u00e4r ov\u00e4ntade utgifter dyker upp?"'],
        ]),

        new Paragraph({ children: [new PageBreak()] }),

        h2("Alex tre livsstilar"),
        para("Alex nettoinkomst: ca 17 600 kr/m\u00e5n. Tre m\u00f6jliga s\u00e4tt att leva:"),

        para('1. "Sparsam Alex"', { bold: true }),
        bullet("Delar l\u00e4genhet med en kompis (4 200 kr var)"),
        bullet("Lagar all mat hemma (2 800 kr)"),
        bullet("Ingen prenumeration, inget gym, begr\u00e4nsat n\u00f6jesliv (500 kr)"),
        bullet("Sparar 2 500 kr/m\u00e5n"),

        para('2. "H\u00e4r-och-nu-Alex"', { bold: true }),
        bullet("Egen etta (7 500 kr)"),
        bullet("\u00c4ter ute 2\u20133 ggr/vecka (4 500 kr mat totalt)"),
        bullet("Gym (399 kr), Spotify+streaming (250 kr), kl\u00e4der (800 kr), n\u00f6je/uteliv (1 500 kr)"),
        bullet("Sparar 0 kr"),

        para('3. "Investerar-Alex"', { bold: true }),
        bullet("Delar l\u00e4genhet (4 200 kr)"),
        bullet("Sn\u00e5lar p\u00e5 n\u00f6je (300 kr), lagar mat hemma (2 800 kr)"),
        bullet("L\u00e4gger 3 000 kr/m\u00e5n p\u00e5 k\u00f6rkort + sparar till utlandsstudier"),
        bullet("Lever stramt nu f\u00f6r framtida m\u00f6jligheter"),

        para("Gemensamma fasta utgifter: telefon (399 kr), hemf\u00f6rs\u00e4kring (150 kr), kollektivtrafik (950 kr), el (300\u2013500 kr).", { italics: true }),

        h2("L\u00e4rarinstruktioner"),
        boldBullet("Uppstart:", "Retrieval practice fr\u00e5n lektion 2. Koppla anst\u00e4llningsformer till inkomst."),
        boldBullet("Instruktion:", "R\u00e4kna konkret p\u00e5 tavlan: 22 000 brutto \u2192 ca 17 600 netto. Koppla till kretsloppet."),
        boldBullet("Bearbetning:", "Pusha j\u00e4mf\u00f6relsen: Vad offrar Sparsam Alex som H\u00e4r-och-nu-Alex f\u00e5r? Se till att paren argumenterar, inte bara r\u00e4knar."),
        boldBullet("Summering:", "EPA-fr\u00e5gan \u00e4r medvetet polariserande. Lyft att det finns strukturella begr\u00e4nsningar OCH individuella val."),

        h2("Differentiering"),
        boldBullet("St\u00f6d (mot E):", '"R\u00e4kna ihop alla fasta utgifter f\u00f6rst. Hur mycket blir kvar? J\u00e4mf\u00f6r de tv\u00e5 livsstilarna \u2014 vilken ger mest pengar \u00f6ver?"'),
        boldBullet("Utmaning (mot A):", '"Vad h\u00e4nder om Alex bara jobbar halvtid? Vilka samh\u00e4llseffekter f\u00e5r det om en hel generation unga lever som Sparsam Alex \u2014 vad h\u00e4nder med konsumtionen och kretsloppet?"'),

        h2("Koppling till kunskapskrav"),
        boldBullet("E-niv\u00e5:", "Redogöra översiktligt för privatekonomiska begrepp och göra en enkel budget."),
        boldBullet("C-niv\u00e5:", "Redogöra utförligt och resonera välgrundat om konsekvenser av olika prioriteringar."),
        boldBullet("A-niv\u00e5:", "Resonera välgrundat och nyanserat om privatekonomiska val, koppla till komplexa samband i samhällsekonomin."),
      ]
    }]
  });
}

// ============================================================
// LEKTION 4
// ============================================================
function lektion4() {
  return new Document({
    numbering: bulletConfig(),
    styles: styles(),
    sections: [{
      properties: {
        page: { size: { width: PAGE_WIDTH, height: PAGE_HEIGHT }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } }
      },
      footers: { default: footerSection("Samh\u00e4llskunskap 1a1", "Ungas ekonomi") },
      children: [
        headerPara('Lektion 4: "Vad h\u00e4nder om...?" \u2014 Risker, val och konsekvenser'),
        para("Kurs: Samh\u00e4llskunskap 1a1 | Moment: Ungas ekonomi | L\u00e4ngd: 60 minuter", { italics: true, color: "666666" }),

        h2("L\u00e4randem\u00e5l f\u00f6r lektionen"),
        bullet("M\u00e5l 3: Resonera om privatekonomiska val och deras konsekvenser"),
        bullet("M\u00e5l 4: Diskutera sambandet individ\u2013samh\u00e4lle"),

        h2("F\u00f6rberedelse"),
        bullet("F\u00f6rbered fyra scenariokort (utskrift, ett set per par)"),
        bullet("Kort presentation om l\u00e5neformer och r\u00e4nta"),
        bullet("Valfritt: Klarna-reklam och snabbl\u00e5ne-reklam f\u00f6r diskussion"),

        h2("Tidsplanering"),
        timeTable([
          ["0\u20135 min", "Uppstart", "Retrieval practice", '"Hur mycket hade Alex kvar varje m\u00e5nad i er budget? Vad var den sv\u00e5raste prioriteringen?"'],
          ["5\u201315 min", "Instruktion", "Risker och skuldf\u00e4llan", "Ov\u00e4ntade utgifter. L\u00e5neformer: snabbl\u00e5n, avbetalning, Klarna. R\u00e4nta-p\u00e5-r\u00e4nta. Konkret: Alex l\u00e5nar 10 000 kr till 35% r\u00e4nta."],
          ["15\u201348 min", "Bearbetning", "Scenarioanalys", "I par: fyra scenariokort med moraliska och strukturella dimensioner. Minst 3 av 4 kort."],
          ["48\u201356 min", "Summering", "Helklassdiskussion", "Vilket scenario var sv\u00e5rast att enas om? Vems ansvar \u00e4r det att unga inte hamnar i skuldf\u00e4llan?"],
          ["56\u201360 min", "Koppling fram\u00e5t", "N\u00e4sta g\u00e5ng", '"N\u00e4sta lektion kopplar vi ihop allt: kretsloppet, arbetsmarknaden, privatekonomin."'],
        ]),

        new Paragraph({ children: [new PageBreak()] }),

        h2("Scenariokort"),

        para("Scenario 1: Tandl\u00e4karen", { bold: true, size: 26 }),
        para("Alex har ont i en tand. Akuttandv\u00e5rd kostar 4 200 kr. Alex har 1 800 kr p\u00e5 kontot. N\u00e4sta l\u00f6n kommer om 18 dagar."),
        bullet("(A) Ta ett snabbl\u00e5n och fixa det direkt."),
        bullet("(B) V\u00e4nta till l\u00f6nen \u2014 men riskera att det blir v\u00e4rre och dyrare."),
        bullet("(C) Fr\u00e5ga f\u00f6r\u00e4ldrarna om pengar."),
        bullet("(D) G\u00e5 till en v\u00e5rdcentral och se om det r\u00e4knas som akutv\u00e5rd."),
        para("Diskussionsfr\u00e5ga: Alla har inte f\u00f6r\u00e4ldrar som kan hj\u00e4lpa till. Hur p\u00e5verkar det Alex valm\u00f6jligheter? \u00c4r det r\u00e4ttvist?", { italics: true, bold: true }),

        para("Scenario 2: Kompisen som beh\u00f6ver hj\u00e4lp", { bold: true, size: 26 }),
        para("Alex kompis Dani har blivit av med sitt jobb och kan inte betala hyran. Dani fr\u00e5gar Alex om att l\u00e5na 5 000 kr \u2014 \u201Cbara till n\u00e4sta m\u00e5nad.\u201D Alex har pengarna p\u00e5 sitt sparkonto, men det \u00e4r hela buffertsparandet."),
        para("Diskussionsfr\u00e5ga: Vad \u00e4r skillnaden mellan att vara en bra v\u00e4n och att fatta kloka ekonomiska beslut? Kan man vara b\u00e5da?", { italics: true, bold: true }),

        para('Scenario 3: "Alla andra har det"', { bold: true, size: 26 }),
        para("Alex jobbarkompisar \u00e5ker p\u00e5 en weekend till K\u00f6penhamn. Det kostar ca 4 000 kr. Alex har egentligen inte r\u00e5d, men vill inte vara den som alltid s\u00e4ger nej. En kompis s\u00e4ger \u201Cl\u00e4gg det p\u00e5 Klarna, betala senare.\u201D"),
        para("Diskussionsfr\u00e5ga: Hur p\u00e5verkar sociala f\u00f6rv\u00e4ntningar ekonomiska beslut? \u00c4r det ett individuellt problem eller ett samh\u00e4llsproblem att f\u00f6retag som Klarna g\u00f6r det l\u00e4tt att skuldsa\u0308tta sig?", { italics: true, bold: true }),

        para("Scenario 4: Systemfr\u00e5gan", { bold: true, size: 26 }),
        para("Alex kollega Fatima jobbar lika m\u00e5nga timmar men f\u00e5r 2 000 kr mindre i m\u00e5naden \u2014 hon \u00e4r inhyrd via bemanningsf\u00f6retag medan Alex \u00e4r direktanst\u00e4lld. Fatima har inga semesterdagar och ingen sjukpenning."),
        para("Diskussionsfr\u00e5ga: \u00c4r det r\u00e4ttvist? Vems ansvar \u00e4r det att l\u00f6sa det \u2014 Fatimas, f\u00f6retagets, fackets eller politikernas?", { italics: true, bold: true }),

        new Paragraph({ children: [new PageBreak()] }),

        h2("L\u00e4rarinstruktioner"),
        boldBullet("Uppstart:", "Retrieval practice fr\u00e5n lektion 3 \u2014 koppla direkt till budgetarbetet."),
        boldBullet("Instruktion:", "R\u00e4kna konkret: snabbl\u00e5n 10 000 kr \u00d7 35% r\u00e4nta. L\u00e5t eleverna reagera."),
        boldBullet("Bearbetning:", "Scenarierna \u00e4r designade att skapa oenighet \u2014 uppmuntra det! Pusha: \u201CSpelar det roll att Alex inte har f\u00f6r\u00e4ldrar som kan hj\u00e4lpa?\u201D (scenario 1, mot C/A). \u201C\u00c4r Klarna problemet \u2014 eller \u00e4r det sociala f\u00f6rv\u00e4ntningar?\u201D (scenario 3, mot A)."),
        boldBullet("Summering:", "Lyft det scenario som skapade mest debatt. Ekonomiska beslut \u00e4r aldrig bara rationella \u2014 de p\u00e5verkas av relationer, normer och vilka resurser man har."),

        h2("Differentiering"),
        boldBullet("St\u00f6d (mot E):", 'Varje scenariokort har st\u00f6dfr\u00e5gor: "Vilka alternativ har Alex?" och "Vad h\u00e4nder om Alex v\u00e4ljer alternativ A? Alternativ B?"'),
        boldBullet("Utmaning (mot A):", '"Diskutera fr\u00e5gan ur b\u00e5de individens och samh\u00e4llets perspektiv. F\u00f6resl\u00e5 en politisk \u00e5tg\u00e4rd som skulle kunna f\u00f6rebygga situationen."'),

        h2("Koppling till kunskapskrav"),
        boldBullet("E-niv\u00e5:", "Identifiera alternativ och ge enkla resonemang om konsekvenser av ekonomiska val."),
        boldBullet("C-niv\u00e5:", "F\u00f6ra v\u00e4lgrundade resonemang om kort- och l\u00e5ngsiktiga konsekvenser och visa p\u00e5 samband."),
        boldBullet("A-niv\u00e5:", "F\u00f6ra v\u00e4lgrundade och nyanserade resonemang ur flera perspektiv och visa p\u00e5 komplexa samband mellan privatekonomi och samh\u00e4llsekonomi."),
      ]
    }]
  });
}

// ============================================================
// GENERATE
// ============================================================
async function main() {
  const outputDir = __dirname;

  const docs = [
    { doc: lektion2(), file: "lektion-2.docx" },
    { doc: lektion3(), file: "lektion-3.docx" },
    { doc: lektion4(), file: "lektion-4.docx" },
  ];

  for (const { doc, file } of docs) {
    const buffer = await Packer.toBuffer(doc);
    const path = `${outputDir}/${file}`;
    fs.writeFileSync(path, buffer);
    console.log(`Generated: ${path}`);
  }
}

main().catch(err => { console.error(err); process.exit(1); });
