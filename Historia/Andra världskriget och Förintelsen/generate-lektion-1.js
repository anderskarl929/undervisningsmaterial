// Generate lektion-1.docx for "Andra världskriget och Förintelsen"
// Historia 1a1 designsystem: Calibri 13pt, Georgia rubriker, navy/guld

const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber,
} = require("docx");

const NAVY = "1B3A5C";
const GOLD = "C8A951";
const GREY_BORDER = "CCCCCC";
const CREME_BG = "FBF8F1";

const A4_WIDTH = 11906;
const A4_HEIGHT = 16838;
const MARGIN = 1440;
const CONTENT_WIDTH = A4_WIDTH - 2 * MARGIN; // 9026

const border = { style: BorderStyle.SINGLE, size: 4, color: GREY_BORDER };
const cellBorders = { top: border, bottom: border, left: border, right: border };

const calibri = (text, opts = {}) =>
  new TextRun({ text, font: "Calibri", size: 26, ...opts });

const para = (children, opts = {}) =>
  new Paragraph({ children, spacing: { after: 120 }, ...opts });

const heading1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 200 },
    children: [new TextRun({ text, font: "Georgia", size: 36, bold: true, color: NAVY })],
  });

const heading2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 140 },
    children: [new TextRun({ text, font: "Georgia", size: 28, bold: true, color: NAVY })],
  });

const heading3 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 100 },
    children: [new TextRun({ text, font: "Georgia", size: 24, bold: true, color: NAVY })],
  });

const meta = (label, value) =>
  new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({ text: `${label}: `, font: "Calibri", size: 26, bold: true, color: NAVY }),
      calibri(value),
    ],
  });

const bullet = (text) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [calibri(text)],
  });

const bulletRich = (parts) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: parts,
  });

// Inline parser: turns markdown-style **bold** and *italic* into TextRuns
function inline(text) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0;
  let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(calibri(text.slice(last, m.index)));
    const tok = m[0];
    if (tok.startsWith("**")) out.push(calibri(tok.slice(2, -2), { bold: true }));
    else out.push(calibri(tok.slice(1, -1), { italics: true }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(calibri(text.slice(last)));
  return out;
}

const richPara = (text) =>
  new Paragraph({ spacing: { after: 120 }, children: inline(text) });

const richBullet = (text) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: inline(text),
  });

const tableCell = (text, opts = {}) =>
  new TableCell({
    borders: cellBorders,
    width: { size: opts.width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [
      new Paragraph({
        spacing: { after: 0 },
        children: inline(text).map((tr) => {
          if (opts.bold) tr.options = { ...tr.options, bold: true };
          return tr;
        }),
      }),
    ],
  });

// Build a row with custom widths
const makeRow = (cells, widths, headerOpts = {}) =>
  new TableRow({
    tableHeader: !!headerOpts.header,
    children: cells.map((text, i) =>
      new TableCell({
        borders: cellBorders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: headerOpts.header ? { fill: NAVY, type: ShadingType.CLEAR } : undefined,
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [
          new Paragraph({
            spacing: { after: 0 },
            children: inline(text).map((tr) => {
              if (headerOpts.header) {
                tr.options = { ...tr.options, bold: true, color: "FFFFFF" };
              }
              return tr;
            }),
          }),
        ],
      })
    ),
  });

// === Lektion 1 content ===

const tidsplaneringHeader = ["Tid", "Fas", "Aktivitet", "Beskrivning"];
const tidsplaneringWidths = [900, 1500, 1700, 4926]; // sum = 9026

const tidsplaneringRows = [
  ["0–5 min", "1. Introduktion", "Retrieval + krok",
    "(a) Tre snabba retrieval-frågor muntligt (no-hands-up + miniwhiteboard): \"Vem var Hitler?\", \"Vad var Versaillesfördraget?\", \"Vad menas med nationalism?\" — 30 sek per fråga. (b) Krok: visa Hédis ungdomsfoto från Sighet och säg: \"Det här är Hédi Fried. Hon föddes 1924. Idag följer vi henne genom hela momentet. Just den här bilden är från en helt vanlig dag i hennes hemstad. Men medan hon firade födelsedag i Sighet började något hända i Tyskland som skulle förändra hennes liv för alltid.\""],
  ["5–20 min", "2. Huvudinnehåll", "Lärarledd genomgång",
    "Tre block: **(A) Hédis Sighet** (3 min) — judiskt liv i Sighet 1933, en av flera minoriteter i Transsylvanien (då Rumänien). **(B) Tre begrepp att äga** (8 min) — bygg Frayer-modell på tavlan tillsammans med eleverna för *antisemitism*, *rasideologi*, *eugenik*. **(C) Tidslinjen 1933–1938** (4 min) — peka på fyra hållpunkter: Hitler vid makten (1933), Nürnberglagarna (1935), Anschluss (1938), Kristallnatten (1938). Avsluta med spegelfrågan: \"Hédi var fortfarande ett barn — märkte hon något av detta?\""],
  ["20–35 min", "3. Elevaktivitet", "Tidslinjeanalys i par",
    "Eleverna får dubbelsidigt papper: vänster sida tom tidslinje 1933–1938, höger sida ett kort utdrag från Hédi (barndom i Sighet). **Uppgift:** (1) placera fyra händelser i tidslinjen, (2) skriv vid varje: \"Detta är ett exempel på antisemitism / rasideologi / eugenik — för att…\", (3) avsluta med en mening om Hédis perspektiv. Läraren modellerar Nürnberglagarna först. Sedan parvis arbete; läraren cirkulerar."],
  ["35–40 min", "4. Summering", "Exit ticket + preview",
    "**Exit ticket:** \"Förklara med egna ord vad antisemitism är. Ge ett exempel från Tyskland 1933–1938.\" **Preview:** \"Nästa gång: 1939 bryter kriget ut. Hédi blir tonåring. Vi följer henne när lagarna gradvis kryper närmare hennes vardag — också i Ungern, dit Sighet snart hör.\""],
];

const kkavHeader = ["Nivå", "Mål 1 (redogöra)", "Mål 2 (begrepp)"];
const kkavWidths = [1100, 3963, 3963];
const kkavRows = [
  ["**E**", "Översiktligt redogör för minst en händelse 1933–1938 (t.ex. *\"Nürnberglagarna gjorde att judar förlorade medborgarskap\"*).",
    "Använder *antisemitism* korrekt med visst stöd; ger ett rimligt exempel."],
  ["**C**", "Utförligt redogör för flera händelser och beskriver hur de hänger ihop (t.ex. *\"Nürnberglagarna förberedde marken för Kristallnatten\"*).",
    "Använder två eller tre begrepp korrekt och visar att hen ser skillnad mellan dem."],
  ["**A**", "Utförligt och nyanserat redogör — visar att nazistisk antijudisk politik var stegvis och systematisk, inte plötslig.",
    "Använder begreppen säkert, kan nyansera (t.ex. *\"Antisemitism var grunden, men rasideologin gjorde den 'vetenskaplig' och eugeniken kopplade ihop den med statlig politik\"*)."],
];

const tidsplaneringTable = new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: tidsplaneringWidths,
  rows: [
    makeRow(tidsplaneringHeader, tidsplaneringWidths, { header: true }),
    ...tidsplaneringRows.map((r) => makeRow(r, tidsplaneringWidths)),
  ],
});

const kkavTable = new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: kkavWidths,
  rows: [
    makeRow(kkavHeader, kkavWidths, { header: true }),
    ...kkavRows.map((r) => makeRow(r, kkavWidths)),
  ],
});

const children = [
  heading1("Lektion 1: Hédis vardag — och hotet hon inte ser"),
  meta("Kurs", "Historia 1a1 (50 p)"),
  meta("Moment", "Andra världskriget och Förintelsen"),
  meta("Lektionslängd", "40 minuter"),

  heading2("Lärandemål för lektionen"),
  bulletRich([
    calibri("LM1", { bold: true }),
    calibri(" (Bloom 2): Eleven ska kunna översiktligt redogöra för nazismens framväxt 1933–1939 och hur antijudisk politik byggdes upp i Tyskland."),
  ]),
  bulletRich([
    calibri("LM2", { bold: true }),
    calibri(" (Bloom 2): Eleven ska kunna använda begreppen "),
    calibri("antisemitism", { italics: true }),
    calibri(", "),
    calibri("rasideologi", { italics: true }),
    calibri(" och "),
    calibri("eugenik", { italics: true }),
    calibri(" i enkla sammanhang."),
  ]),

  heading2("Förberedelse"),
  richBullet("Familjefoto / porträtt av Hédi Fried som ung (Sighet 1930-tal). [VERIFIERA bildkälla — ev. Forum för levande historia eller hennes egna böcker]"),
  richBullet("Karta över Europa 1933 med Tyskland, Ungern, Rumänien och Sighet (Transsylvanien) markerade."),
  richBullet("Tidslinje 1933–1938 med fyra hållpunkter: Hitler vid makten (jan 1933), Nürnberglagarna (sept 1935), Anschluss (mars 1938), Kristallnatten (nov 1938)."),
  richBullet("Begreppskort/begreppsruta på tavlan eller utskrivet: *antisemitism*, *rasideologi*, *eugenik* (förenklade definitioner)."),
  richBullet("Exit ticket-lappar."),
  richBullet("Standard- och förenklad version av elevuppgiften."),

  heading2("Retrieval review-koppling"),
  richPara("Detta är momentets första lektion. Retrieval review kopplar tillbaka till **förkunskaper från Mellankrigstiden / Världskrigen-momentet**: nazismens framväxt i Tyskland, Versaillesfördraget och ekonomisk kris på 1930-talet. Eleverna har stött på *nationalism* tidigare i kursen — det begreppet återbrukas idag och länkas till *antisemitism* och *rasideologi*."),

  heading2("Tidsplanering"),
  tidsplaneringTable,

  heading2("Lärarinstruktioner"),

  heading3("Fas 1 (introduktion)"),
  richBullet("Retrieval-frågorna ska kännas snabba och låga risk — använd 30 sekunders tystnad, ta sedan svar från 2 elever per fråga (no-hands-up). Korrigera kort om nationalism är otydligt."),
  richBullet("Krok-momentet är **det affektiva ankaret för hela momentet** — använd Hédis foto i stillhet i 5–10 sekunder innan du börjar berätta. Säg hennes namn långsamt. Detta är medvetet narrativt."),

  heading3("Fas 2 (huvudinnehåll)"),
  richBullet("**Block A:** Var noggrann med *Sighet 1933 var Rumänien, inte Ungern*. Det är en vanlig missuppfattning. Visa kartan."),
  richBullet("**Block B (modellering):** Bygg Frayer-modellen på tavlan **tillsammans** med eleverna — ställ frågor istället för att skriva färdig text. *\"Vilket ord är motsatsen till antisemitism — vad skulle vara icke-exempel?\"* (T.ex. religionsfrihet, jämlikhet inför lagen.)"),
  richBullet("**Nyckelformulering antisemitism:** *\"Hat eller fientlighet riktad specifikt mot judar — som grupp, oavsett vad enskilda judar gör eller tror.\"*"),
  richBullet("**Nyckelformulering rasideologi:** *\"En lögn om att människor kan delas in i 'raser' med olika värde — och att den egna 'rasen' är överlägsen. Inom nazismen var detta en kärnidé, inte en bisak.\"*"),
  richBullet("**Nyckelformulering eugenik:** *\"Tanken att samhället kan 'förbättras' genom att kontrollera vilka som får skaffa barn — i värsta fall genom tvångssterilisering eller mord. Naziregimen tillämpade detta först på funktionsnedsatta tyskar (T4-programmet) [VERIFIERA — T4 startade 1939, gränsfall för denna lektion], innan dödandet utvidgades till judar och andra grupper.\"*"),
  richBullet("**Block C:** Stanna upp efter Kristallnatten. Fråga: *\"Om ni var judiska i Tyskland efter 9 november 1938 — skulle ni stanna eller fly?\"* Notera (utan att fördjupa nu): *\"Hédis familj kunde inte fly — vi förstår varför nästa lektion.\"*"),
  richBullet("**Kontrollfråga vid slutet av fas 2:** *\"Vad är skillnaden mellan antisemitism och rasideologi? Är de samma sak?\"* Ta in 2–3 svar."),

  heading3("Fas 3 (elevaktivitet)"),
  richBullet("Modellera Nürnberglagarna fullt ut innan paren börjar. Skriv på tavlan: *\"1935 — Nürnberglagarna. Detta är ett exempel på **antisemitism** — för att lagen specifikt riktade sig mot judar och tog ifrån dem rättigheter. Det är också **rasideologi** — för att lagen byggde på lögnen om 'raser'.\"*"),
  richBullet("Cirkulera till stödpar först. Ställ inte ledande frågor — ställ Frayer-frågor: *\"Vilket av begreppen passar bäst här? Varför inte de andra två?\"*"),
  richBullet("Fördjupningsfråga till A-elever: *\"Kan en händelse vara både antisemitism OCH rasideologi OCH eugenik samtidigt? Hitta en sådan.\"*"),

  heading3("Fas 4 (summering)"),
  richBullet("Samla in exit tickets. Sortera **idag eller senast inför L2** i tre högar: *förstod / osäker / missade*. Använd som retrieval-data för L2:s introduktion."),
  richBullet("Preview ska peka framåt utan att avslöja för mycket — håll spänningen i den narrativa bågen."),

  heading2("Elevaktiviteter"),
  richBullet("**Fas 1:** Svara enskilt på tre korta retrievalfrågor (miniwhiteboard eller papper)."),
  richBullet("**Fas 2:** Aktivt deltagande i Frayer-modellbygge på tavlan (svara på frågor från läraren, bidra med exempel/icke-exempel)."),
  richBullet("**Fas 3:** I par: (1) placera fyra händelser på en tidslinje, (2) skriva vid varje vilket av begreppen *antisemitism / rasideologi / eugenik* som det är ett exempel på + motivering, (3) skriva en avslutande mening om Hédis perspektiv."),
  richBullet("**Fas 4:** Skriva exit ticket enskilt — kort definition + ett konkret exempel från perioden 1933–1938."),

  heading2("Differentiering"),

  heading3("Stöd (mot E)"),
  richBullet("Förenklad elevuppgift med kortare meningar och bilder vid varje händelse på tidslinjen."),
  richBullet("Förifylld begreppsruta (definitioner finns redan; eleven hittar exempel)."),
  richBullet("Stödblad med en exempelmening: *\"Nürnberglagarna var ett exempel på antisemitism eftersom...\"*"),
  richBullet("Läraren cirkulerar till stödpar tidigt under fas 3."),
  richBullet("Uppgift 3 i fas 3 (avslutningsmeningen om Hédi) är frivillig för stödelever om tiden är knapp."),

  heading3("Utmaning (mot A)"),
  richBullet("Tilläggsfråga skriftligt: *\"Diskutera: kan vanligt folk i Tyskland 1938 sägas ha 'inte vetat'? Motivera utifrån vad du nu vet om Nürnberglagarna och Kristallnatten.\"*"),
  richBullet("Fördjupning: *\"Vilket av de tre begreppen (antisemitism, rasideologi, eugenik) anser du var den 'farligaste' grunden för det som kom att hända? Argumentera kort.\"*"),
  richBullet("Möjligt extramaterial: kort utdrag (1 sida) ur kapitel 30 i *A History of World Societies* om nazismens uppgång [VERIFIERA — finns i NotebookLM]."),

  heading2("Material"),
  richBullet("Hédis ungdomsfoto (Sighet 1930-tal) [VERIFIERA bildkälla]"),
  richBullet("Karta Europa 1933 med Tyskland, Ungern, Rumänien, Sighet markerade"),
  richBullet("Tidslinje 1933–1938 (på tavla eller utskrift)"),
  richBullet("Begreppskort: antisemitism, rasideologi, eugenik (för Frayer-modell)"),
  richBullet("Elevuppgift L1 standard + förenklad (genereras separat)"),
  richBullet("Exit ticket-lappar"),
  richBullet("Stödblad för elever som behöver"),

  heading2("Exit ticket"),
  richPara("**Fråga:** *\"Förklara med egna ord vad antisemitism är. Ge ett exempel från Tyskland 1933–1938.\"*"),
  richPara("**Användning:**"),
  richBullet("Sortera i tre högar: förstod / osäker / missade."),
  richBullet("Om många blandar ihop *antisemitism* med *rasideologi* — börja L2 med en snabb klargörande retrieval (*\"Är allt antisemitism också rasideologi? Är allt rasideologi också antisemitism?\"*)."),
  richBullet("Om många missar Nürnberglagarna eller Kristallnatten som exempel — repetera den händelsen kort i L2:s intro."),
  richBullet("Spara exit tickets — de blir formativ data för L6 där samband mellan ideologi och politiska beslut analyseras."),

  heading2("Koppling till kunskapskrav (Historia 1a1)"),
  kkavTable,

  heading2("Kvalitetskontroll"),
  richBullet("Retrieval-koppling till förkunskaper finns (fas 1)."),
  richBullet("Elevaktiv tid: 20 min av 40 = 50% (15 min fas 3 + 5 min fas 1) — godkänt."),
  richBullet("Differentiering konkret (förifylld begreppsruta, exempelmening, fördjupningsfrågor markerade)."),
  richBullet("[VERIFIERA]-taggar vid osäkra fakta (Sighet 1933 = Rumänien; T4-programmet; bildkälla; citat)."),
  richBullet("Kopplar framåt till L2 (preview om kriget och Ungern)."),
  richBullet("4-fasstrukturen följs."),

  heading2("Etiska anmärkningar för läraren"),
  richBullet("**Yad Vashem-principen:** Hédis foto visas som *person* (en flicka i Sighet), inte som offer. Den affektiva ingången bygger empati före tragedi."),
  richBullet("**Inga grafiska bilder** denna lektion. Ingen bild från Kristallnatten med våld — beskriv händelsen i ord."),
  richBullet("Om någon elev har judisk, romsk, polsk eller annan direkt familjekoppling: kolla med läraren i förväg. Ge utrymme för det före lektionsstart om relevant."),
  richBullet("L1 är pedagogiskt \"mjuk\" — den affektiva tyngden växer mot L4–L5. Var medveten om bågen."),
];

const doc = new Document({
  creator: "Cornelius v03.26",
  title: "Lektion 1 — Hédis vardag",
  description: "Andra världskriget och Förintelsen, Historia 1a1",
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 26 } }, // 13pt brödtext
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Georgia", color: NAVY },
        paragraph: { spacing: { before: 240, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Georgia", color: NAVY },
        paragraph: { spacing: { before: 220, after: 140 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Georgia", color: NAVY },
        paragraph: { spacing: { before: 160, after: 100 }, outlineLevel: 2 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: A4_WIDTH, height: A4_HEIGHT },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "Historia 1a1 · Andra världskriget och Förintelsen", font: "Calibri", size: 20, color: NAVY })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Lektion 1 · sida ", font: "Calibri", size: 20, color: "666666" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Calibri", size: 20, color: "666666" }),
          ],
        })],
      }),
    },
    children,
  }],
});

const outPath = "/home/anders/undervisningsmaterial/Historia/Andra världskriget och Förintelsen/lektion-1.docx";
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("Wrote:", outPath, "(" + buf.length + " bytes)");
});
