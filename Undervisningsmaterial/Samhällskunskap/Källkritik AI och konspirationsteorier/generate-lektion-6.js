const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageBreak
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "2C3E50", type: ShadingType.CLEAR },
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 22 })] })],
  });
}

function cell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 22 })] })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 24 })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true })],
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true })],
  });
}

function bodyText(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 24 })],
  });
}

function italicText(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 24, italics: true })],
  });
}

function boldBodyText(label, text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({ text: label, font: "Arial", size: 24, bold: true }),
      new TextRun({ text, font: "Arial", size: 24 }),
    ],
  });
}

function boldItalicBody(label, text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({ text: label, font: "Arial", size: 24, bold: true }),
      new TextRun({ text, font: "Arial", size: 24, italics: true }),
    ],
  });
}

function spacer() {
  return new Paragraph({ spacing: { after: 60 }, children: [] });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1A1A2E" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2C3E50" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "34495E" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Samh\u00E4llskunskap 3 / Int. relationer \u2014 K\u00E4llkritik: AI och konspirationsteorier", font: "Arial", size: 18, color: "888888" })],
        })],
      }),
    },
    children: [
      // TITEL
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "Lektion 6: Summativ skrivuppgift \u2014 K\u00E4llkritisk analys", font: "Arial" })],
      }),
      bodyText("Perspektivanalys av konspirationsteori och AI-genererat inneh\u00E5ll"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska visa sin f\u00F6rm\u00E5ga att granska och v\u00E4rdera k\u00E4llor med hj\u00E4lp av k\u00E4llkritiska verktyg, resonera om varf\u00F6r konspirationsteorier och AI-genererat inneh\u00E5ll kan upplevas som trov\u00E4rdiga, samt referera till k\u00E4llor med Harvard-referering."),
      spacer(),
      bullet("Till\u00E4mpa k\u00E4llkritiska verktyg p\u00E5 digitalt inneh\u00E5ll (m\u00E5l 1)"),
      bullet("Analysera varf\u00F6r konspirationsteorier upplevs som \u00F6vertygande (m\u00E5l 2)"),
      bullet("Referera till k\u00E4llor i skriftliga uppgifter (m\u00E5l 4)"),
      spacer(),
      boldBodyText("E: ", "Eleven g\u00F6r en \u00F6versiktlig analys av det k\u00E4llkritiska materialet och f\u00F6r enkla resonemang om varf\u00F6r konspirationsteorier och AI-genererat inneh\u00E5ll kan upplevas som trov\u00E4rdiga. Eleven visar p\u00E5 enkla samband mellan k\u00E4llkritiska begrepp och det analyserade materialet. Eleven anv\u00E4nder i n\u00E5gon m\u00E5n relevanta k\u00E4llkritiska verktyg (t.ex. SIFT-metoden eller de fyra grundfr\u00E5gorna) f\u00F6r att granska p\u00E5st\u00E5enden. Eleven refererar till k\u00E4llor p\u00E5 ett i huvudsak fungerande s\u00E4tt."),
      boldBodyText("C: ", "Eleven g\u00F6r en utf\u00F6rlig analys av det k\u00E4llkritiska materialet och f\u00F6r v\u00E4lgrundade resonemang om varf\u00F6r konspirationsteorier och AI-genererat inneh\u00E5ll kan upplevas som trov\u00E4rdiga, med visst kritiskt perspektiv p\u00E5 de psykologiska och sociala mekanismer som driver spridning. Eleven visar p\u00E5 samband mellan k\u00E4llkritiska begrepp, konspirationsteorins struktur och AI:s roll i informationslandskapet. Eleven anv\u00E4nder relevanta k\u00E4llkritiska verktyg systematiskt och motiverar sina slutsatser med st\u00F6d i det analyserade materialet. Eleven refererar till k\u00E4llor p\u00E5 ett fungerande s\u00E4tt med Harvard-referering."),
      boldBodyText("A: ", "Eleven g\u00F6r en utf\u00F6rlig och nyanserad analys av det k\u00E4llkritiska materialet och f\u00F6r v\u00E4lgrundade och nyanserade resonemang om varf\u00F6r konspirationsteorier och AI-genererat inneh\u00E5ll kan upplevas som trov\u00E4rdiga, ur flera perspektiv (t.ex. psykologiska, sociala och teknologiska). Eleven visar p\u00E5 komplexa samband mellan k\u00E4llkritiska begrepp, konspirationsteorins struktur, AI:s roll och bredare samh\u00E4llskonsekvenser. Eleven anv\u00E4nder k\u00E4llkritiska verktyg med s\u00E4kerhet och v\u00E4ger olika tolkningar mot varandra. Eleven refererar till k\u00E4llor p\u00E5 ett v\u00E4l fungerande s\u00E4tt med konsekvent Harvard-referering."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),
      bullet("K\u00E4llh\u00E4nvisning enligt vanliga system."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("V\u00E4lj EN av de tre rubrikf\u00F6rslagen (se bilaga 1) \u2014 rubriken avsl\u00F6jas F\u00D6RST p\u00E5 examinationsdagen"),
      bullet("Skriv ut examinationsmaterialet (se bilaga 2) \u2014 ett exemplar per elev"),
      bullet("Skriv ut bed\u00F6mningsmatrisen (se bilaga 3) \u2014 ett exemplar per elev"),
      bullet("Skriv ut exit ticket L6 (se bilaga 4) \u2014 ett exemplar per elev"),
      bullet("Ha l\u00E4randem\u00E5let och E/C/A-kriterierna synliga p\u00E5 tavlan eller projicerade"),
      bullet("F\u00F6rbered retrieval review-fr\u00E5gor baserat p\u00E5 exit tickets fr\u00E5n L5"),
      bullet("S\u00E4kerst\u00E4ll att inga datorer/telefoner anv\u00E4nds under examinationen (ingen internet)"),

      // TIDSPLANERING
      heading2("Tidsplanering"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1000, 1300, 2200, 4526],
        rows: [
          new TableRow({
            children: [
              headerCell("Tid", 1000),
              headerCell("Fas", 1300),
              headerCell("Aktivitet", 2200),
              headerCell("Beskrivning", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("0\u20138 min", 1000),
              cell("Retrieval review", 1300),
              cell("Repetition av centrala begrepp", 2200),
              cell("Snabb parvning: eleverna testar varandra p\u00E5 k\u00E4llkritiska grundfr\u00E5gor, SIFT-metoden och mekanismer bakom konspirationsteorier. L\u00E4raren samlar 3\u20134 svar och fyller i luckor. Fokus p\u00E5 begrepp som \u00E4r relevanta f\u00F6r skrivuppgiften.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201313 min", 1000),
              cell("Instruktion", 1300),
              cell("Uppgiftsinstruktion och rubrikutdelning", 2200),
              cell("Presentera rubriken (v\u00E4ld i f\u00F6rv\u00E4g fr\u00E5n bilaga 1). Dela ut analysmaterialet (bilaga 2) och bed\u00F6mningsmatrisen (bilaga 3). G\u00E5 igenom praktiska ramar: 500\u2013800 ord, egna anteckningar + utdelat material, ingen internet. Svara p\u00E5 fr\u00E5gor om uppgiftsformuleringen.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("13\u201318 min", 1000),
              cell("Gemensam l\u00E4sning", 1300),
              cell("Genomg\u00E5ng av materialet", 2200),
              cell("L\u00E4raren l\u00E4ser igenom materialet med klassen. Klarg\u00F6r oklarheter i texten (ordf\u00F6rklaringar, kontext). VIKTIGT: Ingen analys h\u00E4r \u2014 bara f\u00F6rst\u00E5else av vad texten s\u00E4ger. Eleverna kan st\u00E4lla fr\u00E5gor om inneh\u00E5llet.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("18\u201373 min", 1000),
              cell("Sj\u00E4lvst\u00E4ndigt skrivande", 1300),
              cell("Examination (55 min)", 2200),
              cell("55 minuter oavbrutet skrivande. Eleverna skriver sin perspektivanalys (500\u2013800 ord). L\u00E4raren \u00E4r tillg\u00E4nglig f\u00F6r fr\u00E5gor om uppgiftsformuleringen men ger INTE inneh\u00E5llshj\u00E4lp. Hj\u00E4lpmedel: egna anteckningar fr\u00E5n L1\u2013L5 + utdelat material.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("73\u201380 min", 1000),
              cell("Insamling + exit ticket", 1300),
              cell("Avslut", 2200),
              cell("Samla in skrivuppgifterna. Exit ticket (bilaga 4): \u201CVilken del av skrivuppgiften var sv\u00E5rast? Beskriv vad som var sv\u00E5rt och vilken strategi du anv\u00E4nde f\u00F6r att hantera det.\u201D", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Fas 1: Retrieval review (0\u20138 min)"),
      boldBodyText("Parrepetition (0\u20135 min): ", ""),
      italicText("S\u00E4g: \u201CInnan vi b\u00F6rjar med dagens examination \u2014 en snabb repetition. V\u00E4nd er till den som sitter bredvid och testa varandra. En av er fr\u00E5gar, den andra svarar. Byt efter 2 minuter.\u201D"),
      spacer(),
      bodyText("Skriv fr\u00E5gor p\u00E5 tavlan:"),
      bullet("Vilka \u00E4r de fyra k\u00E4llkritiska grundfr\u00E5gorna?"),
      bullet("Vad st\u00E5r SIFT f\u00F6r?"),
      bullet("N\u00E4mn tv\u00E5 mekanismer som g\u00F6r konspirationsteorier \u00F6vertygande."),
      bullet("Hur kan AI f\u00F6rst\u00E4rka spridningen av konspirationsteorier?"),
      spacer(),
      boldBodyText("Gemensam genomg\u00E5ng (5\u20138 min): ", ""),
      bodyText("Samla 3\u20134 svar muntligt. Fyll i eventuella luckor. Om exit tickets fr\u00E5n L5 visade specifika sv\u00E5righeter, l\u00E4gg extra tid p\u00E5 dessa begrepp."),
      italicText("\u201CBra. Nu har ni friskat upp minnet. Dessa begrepp \u00E4r era verktyg f\u00F6r dagens skrivuppgift.\u201D"),

      heading3("Fas 2: Uppgiftsinstruktion och rubrikutdelning (8\u201313 min)"),
      italicText("S\u00E4g: \u201CIdag ska ni visa vad ni l\u00E4rt er under momentet. Ni ska skriva en perspektivanalys d\u00E4r ni granskar ett material med hj\u00E4lp av k\u00E4llkritiska verktyg.\u201D"),
      spacer(),
      bodyText("Dela ut:"),
      bullet("Analysmaterialet (bilaga 2) \u2014 samma material till alla elever"),
      bullet("Bed\u00F6mningsmatrisen (bilaga 3) \u2014 s\u00E5 eleverna vet vad som bed\u00F6ms"),
      spacer(),
      bodyText("Skriv rubriken p\u00E5 tavlan (v\u00E4ld EN fr\u00E5n bilaga 1). G\u00E5 igenom ramarna:"),
      bullet("Omfattning: 500\u2013800 ord"),
      bullet("Hj\u00E4lpmedel: egna anteckningar fr\u00E5n L1\u2013L5 + utdelat material"),
      bullet("Ingen internet, inga datorer/telefoner"),
      bullet("Tid: 55 minuter"),
      bullet("Skriv f\u00F6r hand p\u00E5 papper"),
      spacer(),
      boldBodyText("VIKTIGT: ", "Svara p\u00E5 fr\u00E5gor om uppgiftsformuleringen nu. Under skrivtiden svarar du BARA p\u00E5 fr\u00E5gor om uppgiften, INTE p\u00E5 inneh\u00E5llsfr\u00E5gor."),
      spacer(),
      italicText("\u201CKom ih\u00E5g tredjepersons-framing: ni analyserar varf\u00F6r n\u00E5gon KAN finna materialet \u00F6vertygande \u2014 inte om NI tror p\u00E5 det.\u201D"),

      heading3("Fas 3: Gemensam genomg\u00E5ng av materialet (13\u201318 min)"),
      italicText("S\u00E4g: \u201CInnan ni b\u00F6rjar skriva l\u00E4ser vi igenom materialet tillsammans. Det h\u00E4r \u00E4r f\u00F6r att s\u00E4kerst\u00E4lla att alla f\u00F6rst\u00E5r vad texten s\u00E4ger \u2014 sj\u00E4lva analysen g\u00F6r ni sj\u00E4lva.\u201D"),
      spacer(),
      bodyText("L\u00E4s igenom materialet h\u00F6gt f\u00F6r klassen. Stanna vid:"),
      bullet("Sv\u00E5ra ord eller begrepp \u2014 f\u00F6rklara kort"),
      bullet("Kontextuell information som kan vara ok\u00E4nd \u2014 ge bakgrund"),
      spacer(),
      boldBodyText("VIKTIGT: ", "Ge INGEN analys eller ledtr\u00E5dar. Fr\u00E5gor som \u201C\u00C4r detta trov\u00E4rdigt?\u201D besvaras med: \u201CDet \u00E4r precis det ni ska analysera i er text.\u201D"),
      spacer(),
      italicText("\u201COk, nu f\u00F6rst\u00E5r alla vad materialet handlar om. D\u00E5 \u00E4r det dags att b\u00F6rja skriva. Ni har 55 minuter.\u201D"),

      heading3("Fas 4+5: Sj\u00E4lvst\u00E4ndigt skrivande (18\u201373 min)"),
      bodyText("55 minuter oavbrutet skrivande. L\u00E4raren:"),
      bullet("Cirkulerar tyst i klassrummet"),
      bullet("Svarar p\u00E5 fr\u00E5gor om uppgiftsformuleringen (\u201CVad menar ni med \u2018mekanismer\u2019?\u201D \u2014 ok att f\u00F6rtydliga)"),
      bullet("Ger INTE inneh\u00E5llshj\u00E4lp (\u201C\u00C4r detta ett bra argument?\u201D \u2014 svarar: \u201CDet \u00E4r din bed\u00F6mning att g\u00F6ra\u201D)"),
      bullet("H\u00E5ller tiden synlig p\u00E5 tavlan eller projicerad"),
      spacer(),
      boldBodyText("Tidsp\u00E5minnelser: ", ""),
      bullet("\u201C30 minuter kvar.\u201D (vid 43 min)"),
      bullet("\u201C15 minuter kvar. B\u00F6rja avrunda om ni inte redan gjort det.\u201D (vid 58 min)"),
      bullet("\u201C5 minuter kvar. Avsluta meningen ni h\u00E5ller p\u00E5 med.\u201D (vid 68 min)"),
      spacer(),
      bodyText("Om en elev blir klar tidigt: uppmuntra att l\u00E4sa igenom och f\u00F6rb\u00E4ttra sin text. \u201CL\u00E4s igenom din text. Har du anv\u00E4nt k\u00E4llkritiska begrepp? Har du k\u00E4llh\u00E4nvisat?\u201D"),

      heading3("Fas 6: Insamling + exit ticket (73\u201380 min)"),
      italicText("S\u00E4g: \u201CPennorna ner. L\u00E4mna in era texter.\u201D"),
      spacer(),
      bodyText("Samla in alla skrivuppgifter."),
      spacer(),
      boldBodyText("Exit ticket (75\u201380 min): ", ""),
      bodyText("Dela ut bilaga 4 eller skriv p\u00E5 tavlan:"),
      italicText("\u201CVilken del av skrivuppgiften var sv\u00E5rast? Beskriv vad som var sv\u00E5rt och vilken strategi du anv\u00E4nde f\u00F6r att hantera det.\u201D"),
      spacer(),
      bodyText("Samla in. Anv\u00E4nd svaren f\u00F6r retrieval review i L7 (metakognitiv reflektion)."),
      spacer(),
      boldBodyText("Avslutning: ", ""),
      italicText("\u201CBra jobbat. N\u00E4sta lektion reflekterar vi \u00F6ver hela momentet \u2014 vad ni l\u00E4rt er och hur ert t\u00E4nkande om k\u00E4llkritik har f\u00F6r\u00E4ndrats. Lektion 8 \u00E4r seminarie-examinationen.\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Parrepetition: testa varandra p\u00E5 centrala begrepp (par, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Gemensam l\u00E4sning av analysmaterialet (helklass, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Sj\u00E4lvst\u00E4ndigt skrivande: perspektivanalys 500\u2013800 ord (enskilt, 55 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Exit ticket: reflektion \u00F6ver sv\u00E5raste delen (enskilt, 5 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 70 av 80 minuter (88%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      bodyText("OBS: Detta \u00E4r en summativ examination. Differentiering sker genom uppgiftens design, inte genom att ge olika st\u00F6d under skrivtiden."),
      spacer(),
      boldBodyText("Inbyggd differentiering: ", "Rubriken till\u00E5ter E/C/A-differentiering. E-elever kan g\u00F6ra en \u00F6versiktlig analys med enkla resonemang. A-elever kan visa nyanserade resonemang fr\u00E5n flera perspektiv och komplexa samband."),
      boldBodyText("Tillg\u00E4nglighet: ", "Elever med dokumenterat behov av extra tid f\u00E5r f\u00F6rl\u00E4ngd skrivtid enligt skolans rutiner. Materialet l\u00E4ses igenom gemensamt f\u00F6r att s\u00E4kerst\u00E4lla f\u00F6rst\u00E5else."),

      // MATERIAL
      heading2("Material"),
      bullet("Analysmaterial \u2014 konspirationsteori-text + AI-genererat inneh\u00E5ll (se bilaga 2)"),
      bullet("Bed\u00F6mningsmatris (se bilaga 3)"),
      bullet("Exit ticket L6 (se bilaga 4)"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),
      bullet("Elevernas egna anteckningar fr\u00E5n L1\u2013L5"),
      bullet("Klocka/tidvisning f\u00F6r eleverna"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("Skrivuppgiften pr\u00F6var f\u00F6rm\u00E5gan att granska och v\u00E4rdera k\u00E4llor med k\u00E4llkritiska verktyg (m\u00E5l 1)"),
      bullet("Perspektivanalysen pr\u00F6var resonemang om varf\u00F6r konspirationsteorier och AI-genererat inneh\u00E5ll upplevs som \u00F6vertygande (m\u00E5l 2)"),
      bullet("K\u00E4llh\u00E4nvisning i texten pr\u00F6var referenshantering med Harvard-systemet (m\u00E5l 4)"),
      spacer(),
      bodyText("Uppgiftsdesignen s\u00E4kerst\u00E4ller att alla elever har samma utg\u00E5ngspunkt (givet material) f\u00F6r j\u00E4mf\u00F6rbar bed\u00F6mning. E/C/A-differentiering sker naturligt genom analysens djup och nyansering."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: L1\u2013L5 (alla verktyg och \u00F6vningar under momentet)"),
      bullet("\u00C5terkommande case fr\u00E5n L1 anv\u00E4nds i analysmaterialet"),
      bullet("Bed\u00F6ms mot E/C/A-kriterier fr\u00E5n bed\u00F6mningskriterier.md (skriftlig examination)"),
      bullet("Exit ticket-data anv\u00E4nds i L7 (metakognitiv reflektion)"),
      bullet("L5 (dress rehearsal) f\u00F6rberedde eleverna f\u00F6r denna uppgift"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: RUBRIKFÖRSLAG
      heading2("Bilaga 1: Rubrikf\u00F6rslag f\u00F6r skrivuppgiften"),
      bodyText("V\u00E4lj ETT av f\u00F6ljande tre alternativ. Rubriken avsl\u00F6jas f\u00F6rst p\u00E5 examinationsdagen."),
      spacer(),

      boldBodyText("Alternativ 1:", ""),
      italicText("\u201CAnalysera det bifogade materialet med hj\u00E4lp av k\u00E4llkritiska verktyg. Varf\u00F6r kan detta inneh\u00E5ll upplevas som trov\u00E4rdigt, och vilka mekanismer ligger bakom?\u201D"),
      spacer(),

      boldBodyText("Alternativ 2:", ""),
      italicText("\u201CMed utg\u00E5ngspunkt i det bifogade materialet: Varf\u00F6r kan n\u00E5gon som l\u00E4ser detta finna det \u00F6vertygande? Analysera med hj\u00E4lp av k\u00E4llkritiska begrepp och resonera om AI:s roll.\u201D"),
      spacer(),

      boldBodyText("Alternativ 3:", ""),
      italicText("\u201CGranska det bifogade materialet k\u00E4llkritiskt. Analysera vilka strategier som anv\u00E4nds f\u00F6r att g\u00F6ra inneh\u00E5llet trov\u00E4rdigt och diskutera hur AI p\u00E5verkar spridningen av denna typ av inneh\u00E5ll.\u201D"),
      spacer(),

      bodyText("Alla tre alternativ kr\u00E4ver analys (inte beskrivning), anv\u00E4nder tredjepersons-framing och till\u00E5ter E/C/A-differentiering. V\u00E4lj det alternativ som b\u00E4st matchar klassens niv\u00E5 \u2014 alternativ 1 \u00E4r mest \u00F6ppet, alternativ 3 \u00E4r mest strukturerat."),

      // BILAGA 2: EXEMPELMATERIAL
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Analysmaterial f\u00F6r examination"),
      bodyText("Materialet best\u00E5r av tv\u00E5 delar: en konspirationsteori-text och ett AI-genererat socialt medie-inl\u00E4gg. B\u00E5da anknyter till de \u00E5terkommande casen fr\u00E5n momentet."),
      spacer(),

      heading3("Del A: Konspirationsteori-text"),
      bodyText("\u201CSanningen om AI-kontona p\u00E5 TikTok\u201D"),
      spacer(),
      bodyText("Under 2025 avsl\u00F6jades att fyra svenska TikTok-konton drevs av en enda person med hj\u00E4lp av AI. Men vad medierna inte ber\u00E4ttar \u00E4r att detta bara \u00E4r toppen p\u00E5 isberget. K\u00E4llor inom teknikindustrin uppger att minst 40% av alla svenska influencer-konton p\u00E5 TikTok och Instagram egentligen styrs av AI-system. De stora techf\u00F6retagen vet om detta men agerar inte \u2014 eftersom dessa konton genererar enorma reklamint\u00E4kter."),
      spacer(),
      bodyText("En anonym visselbl\u00E5sare p\u00E5 ett stort socialt medie-f\u00F6retag har l\u00E4ckt interna dokument som visar att algoritmen medvetet prioriterar AI-genererat inneh\u00E5ll framf\u00F6r \u00E4kta inneh\u00E5ll, eftersom AI-inneh\u00E5llet \u00E4r optimerat f\u00F6r maximalt engagemang. \u201CDe vill inte att ni ska veta detta\u201D, skriver visselbl\u00E5saren, \u201Cf\u00F6r om m\u00E4nniskor slutade lita p\u00E5 inneh\u00E5llet skulle hela aff\u00E4rsmodellen kollapsa.\u201D"),
      spacer(),
      bodyText("Det \u00E4r ingen slump att detta sker just nu. Inf\u00F6r det svenska valet 2026 har intresset f\u00F6r att p\u00E5verka svensk opinion aldrig varit st\u00F6rre. Enligt oberoende forskare \u00E4r det sannolikt att minst tre utl\u00E4ndska stater anv\u00E4nder AI-genererade konton f\u00F6r att sprida desinformation riktad mot svenska v\u00E4ljare."),
      spacer(),

      heading3("Del B: AI-genererat socialt medie-inl\u00E4gg"),
      bodyText("(Simulerat Instagram-inl\u00E4gg)"),
      spacer(),
      boldBodyText("@nyhetsgraansen ", "(31 200 f\u00F6ljare)"),
      bodyText("\u201CAVSL\u00D6JAT: Forskare vid KTH bekr\u00E4ftar \u2014 73% av nyskapat inneh\u00E5ll p\u00E5 svenska TikTok genereras av AI. \u2018Vi var chockade \u00F6ver resultaten\u2019, s\u00E4ger professor Anna Lindstr\u00F6m. Samtidigt v\u00E4grar plattformarna att agera. L\u00E4nk i bio f\u00F6r hela rapporten. #AIsvenskpolitik #deepfake #val2026 #k\u00E4llkritik\u201D"),
      spacer(),
      bodyText("[Bilden visar en professionellt designad nyhetsgrafik med KTH:s logotyp och rubriken \u201CFORSKNINGSRAPPORT: AI p\u00E5 svenska sociala medier\u201D]"),
      spacer(),
      italicText("OBS f\u00F6r l\u00E4raren: Inl\u00E4gget \u00E4r konstruerat f\u00F6r examinationen. \u201C@nyhetsgraansen\u201D existerar inte. Siffran 73% \u00E4r fabricerad. \u201CProfessor Anna Lindstr\u00F6m\u201D \u00E4r en p\u00E5hittad person. KTH:s logotyp anv\u00E4nds utan tillst\u00E5nd i detta fiktiva exempel."),

      // BILAGA 3: BEDÖMNINGSMATRIS
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Bed\u00F6mningsmatris \u2014 Skriftlig examination L6"),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 2509, 2508, 2509],
        rows: [
          new TableRow({
            children: [
              headerCell("Kriterium", 1500),
              headerCell("E-niv\u00E5", 2509),
              headerCell("C-niv\u00E5", 2508),
              headerCell("A-niv\u00E5", 2509),
            ],
          }),
          new TableRow({
            children: [
              cell("Analys", 1500),
              cell("\u00D6versiktlig analys. Enkla resonemang om trov\u00E4rdighet. Enkla samband mellan begrepp och material.", 2509),
              cell("Utf\u00F6rlig analys. V\u00E4lgrundade resonemang med visst kritiskt perspektiv p\u00E5 psykologiska/sociala mekanismer. Samband mellan begrepp, struktur och AI:s roll.", 2508),
              cell("Utf\u00F6rlig och nyanserad analys. V\u00E4lgrundade och nyanserade resonemang ur flera perspektiv. Komplexa samband inkl. samh\u00E4llskonsekvenser.", 2509),
            ],
          }),
          new TableRow({
            children: [
              cell("K\u00E4llkritiska verktyg", 1500),
              cell("Anv\u00E4nder i n\u00E5gon m\u00E5n relevanta verktyg (t.ex. SIFT, grundfr\u00E5gorna).", 2509),
              cell("Anv\u00E4nder verktyg systematiskt. Motiverar slutsatser med st\u00F6d i materialet.", 2508),
              cell("Anv\u00E4nder verktyg med s\u00E4kerhet. V\u00E4ger olika tolkningar mot varandra.", 2509),
            ],
          }),
          new TableRow({
            children: [
              cell("K\u00E4llh\u00E4nvisning", 1500),
              cell("Refererar till k\u00E4llor p\u00E5 ett i huvudsak fungerande s\u00E4tt.", 2509),
              cell("Refererar p\u00E5 ett fungerande s\u00E4tt med Harvard-referering.", 2508),
              cell("Refererar p\u00E5 ett v\u00E4l fungerande s\u00E4tt med konsekvent Harvard-referering.", 2509),
            ],
          }),
        ],
      }),

      // BILAGA 4: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: Exit ticket \u2014 Lektion 6"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "Vilken del av skrivuppgiften var sv\u00E5rast? Beskriv vad som var sv\u00E5rt och vilken strategi du anv\u00E4nde f\u00F6r att hantera det."),
      spacer(),
      bodyText("Denna exit ticket \u00E4r en metakognitiv reflektion \u2014 det finns inget \u201Cr\u00E4tt svar\u201D. Syftet \u00E4r att du t\u00E4nker \u00F6ver din egen l\u00E4rprocess."),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-6.docx", buffer);
  console.log("lektion-6.docx skapad!");
});
