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
        children: [new TextRun({ text: "Lektion 8: Seminarium \u2014 K\u00E4llkritik i praktiken", font: "Arial" })],
      }),
      bodyText("Summativ seminarie-examination med tilldelade perspektiv"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska visa sin f\u00F6rm\u00E5ga att resonera om varf\u00F6r konspirationsteorier och AI-genererat inneh\u00E5ll kan upplevas som \u00F6vertygande, samt delta i muntlig diskussion med analytisk distans och tredjepersons-framing."),
      spacer(),
      bullet("Analysera varf\u00F6r konspirationsteorier upplevs som \u00F6vertygande (m\u00E5l 2)"),
      bullet("Diskutera k\u00E4llkritiska fr\u00E5gor i analytisk form (m\u00E5l 3)"),
      spacer(),
      boldBodyText("E: ", "Eleven f\u00F6r enkla resonemang utifr\u00E5n sitt tilldelade perspektiv och visar \u00F6versiktligt p\u00E5 kopplingar mellan k\u00E4llkritiska begrepp och konspirationsteoriers spridning. Eleven lyssnar p\u00E5 andra deltagare och ger enkla kommentarer som anknyter till diskussionen. Eleven anv\u00E4nder tredjepersons-framing i n\u00E5gon m\u00E5n \u2014 det vill s\u00E4ga analyserar varf\u00F6r n\u00E5gon kan finna ett p\u00E5st\u00E5ende \u00F6vertygande snarare \u00E4n att uttrycka personliga \u00E5sikter. Eleven uttrycker sig med viss spr\u00E5klig s\u00E4kerhet i den muntliga framst\u00E4llningen."),
      boldBodyText("C: ", "Eleven f\u00F6r v\u00E4lgrundade resonemang utifr\u00E5n sitt tilldelade perspektiv och g\u00F6r en utf\u00F6rlig analys av kopplingar mellan k\u00E4llkritiska begrepp, konspirationsteoriers psykologi och AI:s roll i spridningen, med visst kritiskt perspektiv. Eleven lyssnar aktivt p\u00E5 andra deltagare, st\u00E4ller f\u00F6rdjupande fr\u00E5gor och bygger vidare p\u00E5 andras resonemang med egna v\u00E4lgrundade argument. Eleven anv\u00E4nder tredjepersons-framing konsekvent genom hela diskussionen och visar p\u00E5 samband mellan olika perspektiv. Eleven uttrycker sig med relativt god spr\u00E5klig s\u00E4kerhet och anpassar sitt spr\u00E5k till den akademiska diskussionsformen."),
      boldBodyText("A: ", "Eleven f\u00F6r v\u00E4lgrundade och nyanserade resonemang utifr\u00E5n sitt tilldelade perspektiv och g\u00F6r en utf\u00F6rlig och nyanserad analys av kopplingar mellan k\u00E4llkritiska begrepp, konspirationsteoriers psykologi och AI:s roll i spridningen, ur flera perspektiv och med ett kritiskt f\u00F6rh\u00E5llningss\u00E4tt. Eleven lyssnar aktivt, st\u00E4ller f\u00F6rdjupande fr\u00E5gor som driver diskussionen fram\u00E5t och syntetiserar andras resonemang med egna komplexa kopplingar. Eleven anv\u00E4nder tredjepersons-framing genomg\u00E5ende och kan reflektera \u00F6ver varf\u00F6r denna analytiska h\u00E5llning \u00E4r viktig. Eleven uttrycker sig med god spr\u00E5klig s\u00E4kerhet och bidrar till att h\u00F6ja den analytiska niv\u00E5n i gruppen."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),
      bullet("Hur propaganda, desinformation och konspirationsteorier fungerar som p\u00E5verkansformer i det digitala medielandskapet."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("Dela in eleverna i sm\u00E5grupper om 4\u20135 elever i f\u00F6rv\u00E4g"),
      bullet("Tilldela varje elev ett perspektiv/roll (se bilaga 2) \u2014 INGEN elev v\u00E4ljer sj\u00E4lv"),
      bullet("V\u00E4lj EN av seminariefr\u00E5gest\u00E4llningarna (se bilaga 1)"),
      bullet("Skriv ut rollkort (bilaga 2) \u2014 ett per elev"),
      bullet("Skriv ut bed\u00F6mningsmatrisen (bilaga 3) \u2014 ett per elev"),
      bullet("F\u00F6rbered observationsprotokoll (bilaga 4) \u2014 ett per grupp"),
      bullet("S\u00E4kerst\u00E4ll inspelningsutrustning (ljud) \u2014 en enhet per grupp"),
      bullet("Informera eleverna i f\u00F6rv\u00E4g att seminariet spelas in (ljud) f\u00F6r bed\u00F6mning"),
      bullet("S\u00E4kerst\u00E4ll samtycke f\u00F6r inspelning enligt skolans rutiner"),
      bullet("Ha l\u00E4randem\u00E5let synligt p\u00E5 tavlan under hela lektionen"),

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
              cell("Repetition inf\u00F6r seminarium", 2200),
              cell("Repetition av tredjepersons-framing: l\u00E4raren visar exempelformuleringar fr\u00E5n tredjepersons-framing.md. Kort p\u00E5minnelse om analytiska perspektiv. P\u00E5minnelse om seminariets norm: vi analyserar varf\u00F6r n\u00E5gon KAN finna n\u00E5got \u00F6vertygande.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201313 min", 1000),
              cell("Instruktion", 1300),
              cell("Fr\u00E5gest\u00E4llning, roller och praktisk info", 2200),
              cell("Presentera seminariefr\u00E5gest\u00E4llningen (bilaga 1). Dela ut rollkort (bilaga 2). P\u00E5minn: tredjepersons-framing \u00E4r obligatorisk. Praktisk info: seminariet spelas in (ljud) f\u00F6r efterbed\u00F6mning. Dela ut bed\u00F6mningsmatrisen (bilaga 3).", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("13\u201368 min", 1000),
              cell("Seminarium", 1300),
              cell("Sm\u00E5grupps-seminarium (55 min)", 2200),
              cell("Seminarium i sm\u00E5grupper om 4\u20135 elever. Varje elev analyserar utifr\u00E5n sitt tilldelade perspektiv. L\u00E4raren observerar, noterar och spelar in. Facilitering vid behov: \u201CNu vill jag h\u00F6ra fr\u00E5n [perspektiv]-experten.\u201D Equity-safeguards aktiva genom hela seminariet.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("68\u201380 min", 1000),
              cell("Avslut", 1300),
              cell("Sammanfattning och reflektion", 2200),
              cell("L\u00E4raren sammanfattar viktiga insikter fr\u00E5n seminarierna. INGEN exit ticket (summativ bed\u00F6mning). Kort reflektion: \u201CVad tar ni med er fr\u00E5n detta moment?\u201D Avslutning av momentet.", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Fas 1: Retrieval review inf\u00F6r seminarium (0\u20138 min)"),
      boldBodyText("Tredjepersons-framing-repetition (0\u20135 min): ", ""),
      italicText("S\u00E4g: \u201CInnan vi b\u00F6rjar seminariet \u2014 en viktig p\u00E5minnelse. Under hela momentet har vi anv\u00E4nt tredjepersons-framing. L\u00E5t oss repetera vad det inneb\u00E4r.\u201D"),
      spacer(),
      bodyText("Skriv p\u00E5 tavlan:"),
      italicText("\u201CVi fr\u00E5gar inte: Tror du p\u00E5 detta?\u201D"),
      italicText("\u201CVi fr\u00E5gar: Varf\u00F6r KAN n\u00E5gon finna detta \u00F6vertygande?\u201D"),
      spacer(),
      bodyText("Visa 2\u20133 exempelformuleringar:"),
      bullet("\u201CUtifr\u00E5n ett psykologiskt perspektiv kan man f\u00F6rst\u00E5 att...\u201D"),
      bullet("\u201CDen som m\u00F6ter denna information utan f\u00F6rkunskap kan finna den \u00F6vertygande eftersom...\u201D"),
      bullet("\u201CEn m\u00F6jlig f\u00F6rklaring till att m\u00E4nniskor attraheras av denna typ av ber\u00E4ttelser \u00E4r att...\u201D"),
      spacer(),
      boldBodyText("Analytiska perspektiv (5\u20138 min): ", ""),
      italicText("S\u00E4g: \u201CNi har arbetat med k\u00E4llkritiska verktyg, konspirationsteoriers psykologi och AI:s roll under hela momentet. Idag ska ni anv\u00E4nda allt detta i en analytisk diskussion. Kom ih\u00E5g \u2014 det handlar inte om att ha r\u00E4tt, utan om att resonera v\u00E4l.\u201D"),

      heading3("Fas 2: Fr\u00E5gest\u00E4llning, roller och praktisk info (8\u201313 min)"),
      italicText("S\u00E4g: \u201CNu till dagens examination. Ni ska delta i ett seminarium i sm\u00E5grupper. Varje grupp har samma fr\u00E5gest\u00E4llning, men ni har olika perspektiv.\u201D"),
      spacer(),
      bodyText("Skriv seminariefr\u00E5gest\u00E4llningen p\u00E5 tavlan (v\u00E4ld EN fr\u00E5n bilaga 1)."),
      spacer(),
      bodyText("Dela ut rollkort (bilaga 2) och bed\u00F6mningsmatrisen (bilaga 3)."),
      spacer(),
      boldBodyText("G\u00E5 igenom praktiska ramar: ", ""),
      bullet("Ni diskuterar i 55 minuter"),
      bullet("Alla ska bidra utifr\u00E5n sitt tilldelade perspektiv"),
      bullet("Tredjepersons-framing \u00E4r obligatorisk \u2014 analysera, uttryck inte personliga \u00E5sikter"),
      bullet("Seminariet spelas in (ljud) s\u00E5 att jag kan bed\u00F6ma i efterhand"),
      bullet("Lyssna p\u00E5 varandra, st\u00E4ll fr\u00E5gor, bygg vidare p\u00E5 andras resonemang"),
      spacer(),
      italicText("\u201CL\u00E4s igenom ert rollkort. Ni har 1 minut att t\u00E4nka igenom ert perspektiv innan vi b\u00F6rjar.\u201D"),

      heading3("Fas 3+4+5: Seminariets genomf\u00F6rande (13\u201368 min)"),
      bodyText("55 minuter seminarium i sm\u00E5grupper. L\u00E4raren:"),
      spacer(),
      boldBodyText("Observation och notering: ", ""),
      bullet("Cirkulera mellan grupperna"),
      bullet("Anv\u00E4nd observationsprotokollet (bilaga 4) f\u00F6r att notera elevers bidrag"),
      bullet("S\u00E4kerst\u00E4ll att inspelningen fungerar i varje grupp"),
      spacer(),
      boldBodyText("Facilitering vid behov: ", ""),
      bullet("Om en elev inte h\u00F6rs: \u201CNu vill jag h\u00F6ra fr\u00E5n [perspektiv]-experten. Vad ser du utifr\u00E5n ditt perspektiv?\u201D"),
      bullet("Om diskussionen stannar: \u201CFinns det n\u00E5got perspektiv vi inte h\u00F6rt fr\u00E5n \u00E4n?\u201D"),
      bullet("Om samtalet blir ytligt: \u201CKan ni f\u00F6rdjupa det \u2014 VARF\u00D6R \u00E4r det s\u00E5?\u201D"),
      bullet("Om elev glider mot personliga \u00E5sikter: \u201CL\u00E5t oss ist\u00E4llet fr\u00E5ga \u2014 varf\u00F6r kan n\u00E5gon som ser denna information finna den \u00F6vertygande?\u201D"),
      spacer(),
      boldBodyText("Equity-safeguards under seminariet: ", ""),
      spacer(),
      boldBodyText("1. Tilldelade roller och perspektiv: ", "Perspektiven \u00E4r f\u00F6rdelade av l\u00E4raren. Ingen elev v\u00E4ljer sj\u00E4lv. Eleverna analyserar utifr\u00E5n sitt perspektiv \u2014 de \u201C\u00E4ger\u201D inte \u00E5sikten."),
      boldBodyText("2. Tredjepersons-framing obligatorisk: ", "L\u00E4raren har modellerat framing i fas 1. P\u00E5minn vid behov under diskussionen. Om en elev glider mot personliga \u00E5sikter, omformulera fr\u00E5gan: \u201CL\u00E5t oss ist\u00E4llet fr\u00E5ga \u2014 varf\u00F6r kan n\u00E5gon finna detta \u00F6vertygande?\u201D"),
      boldBodyText("3. Privat uppf\u00F6ljningsplan: ", "Observera under seminariet. Notera elever som visar tecken p\u00E5 obehag, stark personlig koppling till \u00E4mnet eller drar sig undan. F\u00F6lj upp enskilt efter lektionen. Samtalet fokuserar p\u00E5 elevens upplevelse. Vid behov h\u00E4nvisas till skolkurator."),
      spacer(),
      boldBodyText("Tidsh\u00E5llning: ", ""),
      bullet("\u201C30 minuter kvar.\u201D (vid 38 min)"),
      bullet("\u201C10 minuter kvar. B\u00F6rja avrunda era resonemang.\u201D (vid 58 min)"),
      bullet("\u201C2 minuter kvar. Sammanfatta era viktigaste insikter.\u201D (vid 66 min)"),

      heading3("Fas 6: Avslut och sammanfattning (68\u201380 min)"),
      italicText("S\u00E4g: \u201CTack f\u00F6r era diskussioner. L\u00E5t oss samla de viktigaste insikterna.\u201D"),
      spacer(),
      bodyText("Sammanfatta kort de centrala temana du observerade:"),
      bullet("Vilka perspektiv gav mest insikt?"),
      bullet("Vilka kopplingar gjordes mellan perspektiven?"),
      bullet("Var uppstod de mest intressanta meningsskiljaktigheterna?"),
      spacer(),
      boldBodyText("INGEN exit ticket ", "\u2014 detta \u00E4r en summativ bed\u00F6mning."),
      spacer(),
      boldBodyText("Avslutande reflektion (75\u201380 min): ", ""),
      italicText("S\u00E4g: \u201CVi har nu avslutat hela momentet om k\u00E4llkritik, AI och konspirationsteorier. Jag vill att ni tar en minut och t\u00E4nker \u2014 vad tar ni med er? Vad vet ni nu som ni inte visste n\u00E4r vi b\u00F6rjade?\u201D"),
      spacer(),
      bodyText("L\u00E5t 3\u20134 elever dela sina reflektioner muntligt. Ingen bed\u00F6mning \u2014 detta \u00E4r en avslutning."),
      spacer(),
      italicText("\u201CTack f\u00F6r ert arbete under momentet. Ni har byggt verktyg som ni kommer ha nytta av varje g\u00E5ng ni m\u00F6ter information online.\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Repetition av tredjepersons-framing (helklass, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "L\u00E4sa rollkort och f\u00F6rbereda perspektiv (enskilt, 1 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Seminarium: diskussion i sm\u00E5grupper (sm\u00E5grupp 4\u20135, 55 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Avslutande reflektion: vad tar ni med er? (helklass, 5 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 66 av 80 minuter (83%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      bodyText("OBS: Detta \u00E4r en summativ examination. Differentiering sker genom seminariets design, inte genom olika st\u00F6dniv\u00E5er."),
      spacer(),
      boldBodyText("Inbyggd differentiering: ", "Rollkorten ger alla elever en utg\u00E5ngspunkt f\u00F6r sina resonemang. E/C/A-differentiering sker naturligt genom djupet i resonemangen, f\u00F6rm\u00E5gan att bygga vidare p\u00E5 andras argument och anv\u00E4ndningen av tredjepersons-framing."),
      boldBodyText("Faciliteringsst\u00F6d: ", "L\u00E4raren faciliterar vid behov f\u00F6r att s\u00E4kerst\u00E4lla att alla h\u00F6rs. \u201CNu vill jag h\u00F6ra fr\u00E5n [perspektiv]-experten\u201D ger tysta elever en naturlig ing\u00E5ng utan att det upplevs som utpekande."),

      // MATERIAL
      heading2("Material"),
      bullet("Seminariefr\u00E5gest\u00E4llning (se bilaga 1)"),
      bullet("Rollkort \u2014 5 perspektiv (se bilaga 2)"),
      bullet("Bed\u00F6mningsmatris f\u00F6r seminarium (se bilaga 3)"),
      bullet("L\u00E4rarens observationsprotokoll (se bilaga 4)"),
      bullet("Inspelningsutrustning (ljud) \u2014 en enhet per grupp"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("Seminariet pr\u00F6var f\u00F6rm\u00E5gan att resonera om varf\u00F6r konspirationsteorier och AI-genererat inneh\u00E5ll upplevs som \u00F6vertygande (m\u00E5l 2)"),
      bullet("Seminariet pr\u00F6var f\u00F6rm\u00E5gan att diskutera k\u00E4llkritiska fr\u00E5gor med analytisk distans och bygga vidare p\u00E5 andras resonemang (m\u00E5l 3)"),
      bullet("Tredjepersons-framing \u00E4r ett explicit bed\u00F6mningskriterium (E: i n\u00E5gon m\u00E5n, C: konsekvent, A: genomg\u00E5ende med reflektion)"),
      spacer(),
      bodyText("Seminariets design (tilldelade perspektiv, facilitering) s\u00E4kerst\u00E4ller att alla elever har en utg\u00E5ngspunkt f\u00F6r sina resonemang. E/C/A-differentiering sker genom analysdjup, f\u00F6rm\u00E5ga att bygga vidare p\u00E5 andras argument och anv\u00E4ndning av tredjepersons-framing."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: HELA momentet (L1\u2013L7)"),
      bullet("\u00C5terkommande case fr\u00E5n L1 anv\u00E4nds som seminariets utg\u00E5ngspunkt"),
      bullet("Bed\u00F6ms mot E/C/A-kriterier fr\u00E5n bed\u00F6mningskriterier.md (muntlig examination)"),
      bullet("Equity-safeguards fr\u00E5n bed\u00F6mningskriterier.md till\u00E4mpas fullt ut"),
      bullet("L4 (formativt seminarium) f\u00F6rberedde eleverna f\u00F6r denna examination"),
      bullet("Sista lektionen i momentet \u2014 avslutar med reflektion"),

      // BEDÖMNING
      heading2("Bed\u00F6mning"),
      bodyText("Seminariet bed\u00F6ms via inspelning + efterbed\u00F6mning. L\u00E4raren lyssnar igenom inspelningarna och bed\u00F6mer varje elev mot E/C/A-kriterierna i bilaga 3."),
      spacer(),
      bodyText("Observationsprotokollet (bilaga 4) anv\u00E4nds som st\u00F6d under seminariet f\u00F6r att notera:"),
      bullet("Vilka elever som bidrar och i vilken omfattning"),
      bullet("Kvaliteten p\u00E5 resonemang (ytligt vs djupt)"),
      bullet("Anv\u00E4ndning av tredjepersons-framing"),
      bullet("F\u00F6rm\u00E5ga att bygga vidare p\u00E5 andras argument"),
      bullet("Elever som kan beh\u00F6va privat uppf\u00F6ljning"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: SEMINARIFRÅGESTÄLLNING
      heading2("Bilaga 1: Seminariefr\u00E5gest\u00E4llning"),
      bodyText("V\u00E4lj EN av f\u00F6ljande tre alternativ. Alla baseras p\u00E5 de \u00E5terkommande casen fr\u00E5n momentet."),
      spacer(),

      boldBodyText("Alternativ 1:", ""),
      italicText("\u201CMed utg\u00E5ngspunkt i AI-genererade TikTok-konton i Sverige (2025): Analysera fr\u00E5n ert tilldelade perspektiv varf\u00F6r denna typ av inneh\u00E5ll kan upplevas som trov\u00E4rdigt och vilka konsekvenser det kan f\u00E5 f\u00F6r samh\u00E4llet.\u201D"),
      spacer(),

      boldBodyText("Alternativ 2:", ""),
      italicText("\u201CMed utg\u00E5ngspunkt i deepfakes inf\u00F6r det svenska valet 2026: Analysera fr\u00E5n ert tilldelade perspektiv hur manipulerat inneh\u00E5ll p\u00E5verkar v\u00E4ljares f\u00F6rm\u00E5ga att g\u00F6ra informerade val och vilka motmedel som finns.\u201D"),
      spacer(),

      boldBodyText("Alternativ 3:", ""),
      italicText("\u201CMed utg\u00E5ngspunkt i hur AI-modeller (som DeepSeek) kan generera \u00F6vertygande konspirationsteorier: Analysera fr\u00E5n ert tilldelade perspektiv vad detta inneb\u00E4r f\u00F6r det k\u00E4llkritiska landskapet och vilka utmaningar samh\u00E4llet st\u00E5r inf\u00F6r.\u201D"),
      spacer(),

      bodyText("Alla tre alternativ anv\u00E4nder tredjepersons-framing och \u00E5terkommande case. V\u00E4lj det alternativ som b\u00E4st matchar vad klassen arbetat mest med under momentet."),

      // BILAGA 2: ROLLKORT
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Rollkort \u2014 Tilldelade perspektiv"),
      bodyText("Varje elev f\u00E5r ETT rollkort. L\u00E4raren f\u00F6rdelar perspektiven \u2014 ingen elev v\u00E4ljer sj\u00E4lv."),
      spacer(),

      heading3("Perspektiv 1: Medieforskaren"),
      bodyText("Du analyserar fr\u00E5gest\u00E4llningen utifr\u00E5n medielandskapet och algoritmernas roll."),
      spacer(),
      bodyText("T\u00E4nk p\u00E5:"),
      bullet("Hur p\u00E5verkar algoritmerna vad m\u00E4nniskor ser och tror p\u00E5?"),
      bullet("Vilken roll spelar ekkammare och filterbubblor?"),
      bullet("Hur har sociala medier f\u00F6r\u00E4ndrat informationsspridningen?"),
      bullet("Vad kan plattformarna g\u00F6ra \u2014 och varf\u00F6r g\u00F6r de det inte?"),
      spacer(),

      heading3("Perspektiv 2: AI-utvecklaren"),
      bodyText("Du analyserar fr\u00E5gest\u00E4llningen utifr\u00E5n teknologins m\u00F6jligheter och begr\u00E4nsningar."),
      spacer(),
      bodyText("T\u00E4nk p\u00E5:"),
      bullet("Hur fungerar AI-generering av text, bild och video?"),
      bullet("Vad g\u00F6r AI-genererat inneh\u00E5ll sv\u00E5rt att skilja fr\u00E5n \u00E4kta?"),
      bullet("Vilka tekniska motmedel finns (t.ex. AI-detektion, vattenm\u00E4rkning)?"),
      bullet("Var g\u00E5r gr\u00E4nsen f\u00F6r vad teknologin kan och b\u00F6r anv\u00E4ndas till?"),
      spacer(),

      heading3("Perspektiv 3: Psykologen"),
      bodyText("Du analyserar fr\u00E5gest\u00E4llningen utifr\u00E5n kognitiva biases och grupppsykologi."),
      spacer(),
      bodyText("T\u00E4nk p\u00E5:"),
      bullet("Vilka psykologiska mekanismer g\u00F6r att m\u00E4nniskor tror p\u00E5 konspirationsteorier?"),
      bullet("Hur p\u00E5verkar confirmation bias, proportionality bias och grupptillh\u00F6righet?"),
      bullet("Varf\u00F6r \u00E4r det sv\u00E5rt att \u00E4ndra uppfattning n\u00E4r man v\u00E4l trott p\u00E5 n\u00E5got?"),
      bullet("Hur kan f\u00F6rst\u00E5else f\u00F6r dessa mekanismer hj\u00E4lpa oss att granska information?"),
      spacer(),

      heading3("Perspektiv 4: Samh\u00E4llsdebatt\u00F6ren"),
      bodyText("Du analyserar fr\u00E5gest\u00E4llningen utifr\u00E5n demokrati och yttrandefrihet."),
      spacer(),
      bodyText("T\u00E4nk p\u00E5:"),
      bullet("Hur p\u00E5verkas demokratin n\u00E4r v\u00E4ljare inte kan lita p\u00E5 information?"),
      bullet("Var g\u00E5r gr\u00E4nsen mellan yttrandefrihet och skydd mot desinformation?"),
      bullet("Vem b\u00E4r ansvaret \u2014 individen, plattformarna eller staten?"),
      bullet("Vilka lagliga och samh\u00E4lleliga \u00E5tg\u00E4rder \u00E4r rimliga?"),
      spacer(),

      heading3("Perspektiv 5: Ungdomsrepresentanten"),
      bodyText("Du analyserar fr\u00E5gest\u00E4llningen utifr\u00E5n ungdomars medievanor och s\u00E5rbarhet."),
      spacer(),
      bodyText("T\u00E4nk p\u00E5:"),
      bullet("Hur konsumerar unga information j\u00E4mf\u00F6rt med \u00E4ldre generationer?"),
      bullet("Varf\u00F6r kan unga vara s\u00E4rskilt s\u00E5rbara f\u00F6r AI-genererad desinformation?"),
      bullet("Vilken roll spelar k\u00E4llkritisk utbildning i skolan?"),
      bullet("Vad kan unga sj\u00E4lva g\u00F6ra f\u00F6r att skydda sig?"),

      // BILAGA 3: BEDÖMNINGSMATRIS
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Bed\u00F6mningsmatris \u2014 Muntlig examination L8"),
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
              cell("Resonemang", 1500),
              cell("Enkla resonemang utifr\u00E5n tilldelat perspektiv. \u00D6versiktliga kopplingar.", 2509),
              cell("V\u00E4lgrundade resonemang. Utf\u00F6rlig analys med visst kritiskt perspektiv. Samband mellan perspektiv.", 2508),
              cell("V\u00E4lgrundade och nyanserade resonemang. Utf\u00F6rlig och nyanserad analys ur flera perspektiv med kritiskt f\u00F6rh\u00E5llningss\u00E4tt.", 2509),
            ],
          }),
          new TableRow({
            children: [
              cell("Diskussion", 1500),
              cell("Lyssnar. Ger enkla kommentarer som anknyter.", 2509),
              cell("Lyssnar aktivt. St\u00E4ller f\u00F6rdjupande fr\u00E5gor. Bygger vidare p\u00E5 andras resonemang.", 2508),
              cell("Lyssnar aktivt. Driver diskussionen fram\u00E5t. Syntetiserar andras resonemang med egna komplexa kopplingar.", 2509),
            ],
          }),
          new TableRow({
            children: [
              cell("Tredjepersons-framing", 1500),
              cell("Anv\u00E4nder i n\u00E5gon m\u00E5n.", 2509),
              cell("Anv\u00E4nder konsekvent genom hela diskussionen.", 2508),
              cell("Anv\u00E4nder genomg\u00E5ende. Reflekterar \u00F6ver varf\u00F6r h\u00E5llningen \u00E4r viktig.", 2509),
            ],
          }),
          new TableRow({
            children: [
              cell("Spr\u00E5klig s\u00E4kerhet", 1500),
              cell("Uttrycker sig med viss spr\u00E5klig s\u00E4kerhet.", 2509),
              cell("Uttrycker sig med relativt god spr\u00E5klig s\u00E4kerhet. Anpassar spr\u00E5ket till akademisk diskussion.", 2508),
              cell("Uttrycker sig med god spr\u00E5klig s\u00E4kerhet. Bidrar till att h\u00F6ja analytisk niv\u00E5.", 2509),
            ],
          }),
        ],
      }),

      // BILAGA 4: OBSERVATIONSPROTOKOLL
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: L\u00E4rarens observationsprotokoll"),
      bodyText("Anv\u00E4nd detta protokoll under seminariet f\u00F6r att notera elevers bidrag. Protokollet \u00E4r ett st\u00F6d \u2014 den slutgiltiga bed\u00F6mningen g\u00F6rs via inspelningen."),
      spacer(),
      boldBodyText("Grupp: ", "_______________"),
      boldBodyText("Datum: ", "_______________"),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 1500, 1500, 1500, 1500, 1526],
        rows: [
          new TableRow({
            children: [
              headerCell("Elevnamn", 1500),
              headerCell("Perspektiv", 1500),
              headerCell("Bidrar aktivt?", 1500),
              headerCell("Tredjepersons-framing?", 1500),
              headerCell("Bygger vidare?", 1500),
              headerCell("Anteckningar", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1500),
              cell("", 1526),
            ],
          }),
        ],
      }),

      spacer(),
      boldBodyText("Elever som kan beh\u00F6va privat uppf\u00F6ljning: ", ""),
      bodyText("_______________________________________________________________"),
      spacer(),
      boldBodyText("\u00D6vriga observationer: ", ""),
      bodyText("_______________________________________________________________"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-8.docx", buffer);
  console.log("lektion-8.docx skapad!");
});
