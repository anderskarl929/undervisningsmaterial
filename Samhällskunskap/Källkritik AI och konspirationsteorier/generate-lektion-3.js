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
        children: [new TextRun({ text: "Lektion 3: Verktygsl\u00E5dan i praktiken", font: "Arial" })],
      }),
      bodyText("F\u00F6rdjupad k\u00E4llkritisk analys \u2014 n\u00E4r AI m\u00F6ter konspirationsteorier"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "75 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska kunna analysera hur AI kan anv\u00E4ndas f\u00F6r att skapa och sprida konspirationsteorier, samt granska k\u00E4llor med hj\u00E4lp av tidigare verktyg."),
      spacer(),
      bullet("Granska och v\u00E4rdera k\u00E4llor med hj\u00E4lp av de fyra k\u00E4llkritiska grundfr\u00E5gorna (m\u00E5l 1)"),
      bullet("Analysera hur AI-genererat inneh\u00E5ll p\u00E5verkar informationslandskapet (m\u00E5l 2)"),
      bullet("Analysera hur konspirationsteorier skapas, sprids och f\u00F6rst\u00E4rks med hj\u00E4lp av AI (m\u00E5l 3)"),
      spacer(),
      boldBodyText("E: ", "Eleven granskar k\u00E4llor med enkla omd\u00F6men och f\u00F6r enkla resonemang om hur AI kan anv\u00E4ndas f\u00F6r att sprida konspirationsteorier."),
      boldBodyText("C: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade omd\u00F6men och f\u00F6r v\u00E4lgrundade resonemang om sambandet mellan AI och konspirationsteorier."),
      boldBodyText("A: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade och nyanserade omd\u00F6men och f\u00F6r v\u00E4lgrundade och nyanserade resonemang om komplexa samband mellan AI, konspirationsteorier och demokrati."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),
      bullet("Presentation i olika former, till exempel debatter och unders\u00F6kande samtal, och med anv\u00E4ndning av digitala verktyg."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("Analysera exit tickets fr\u00E5n lektion 2 \u2014 notera vilka spridningsmekanismer eleverna beh\u00E4rskar och vilka som beh\u00F6ver repeteras"),
      bullet("F\u00F6rbered retrieval practice-fr\u00E5gor baserat p\u00E5 lektion 2:s exit tickets"),
      bullet("F\u00F6rbered tre verkliga fall d\u00E4r AI anv\u00E4nts f\u00F6r att skapa/sprida desinformation (se bilaga 1 f\u00F6r f\u00E4rdiga fall)"),
      bullet("Skriv ut analysblad \u201CAI + konspirationsteori \u2014 kombinerad analys\u201D (en per elev, se bilaga 2)"),
      bullet("Ha l\u00E4randem\u00E5len synliga p\u00E5 tavlan under hela lektionen"),
      bullet("Testa att alla l\u00E4nkar i bilaga 4 fungerar"),

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
              cell("Uppstart", 1300),
              cell("Retrieval practice + hook", 2200),
              cell("Snabbskrivning (3 min): \u201CAnge tv\u00E5 anledningar till att konspirationsteorier sprids.\u201D Gemensam genomg\u00E5ng. Hook: \u201C2 089 fejkade nyhetssajter drivna av AI. Idag sl\u00E5r vi ihop era tv\u00E5 verktygsl\u00E5dor.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201322 min", 1000),
              cell("Instruktion", 1300),
              cell("Tre verkliga fall: AI + konspirationsteorier", 2200),
              cell("Presentera tre fall d\u00E4r AI anv\u00E4nts f\u00F6r att skapa/sprida konspirationsteorier. Analysera varje fall med b\u00E5de grundfr\u00E5gorna OCH konspirationsteorins anatomi. EPA-stopp vid minut 16.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("22\u201358 min", 1000),
              cell("Bearbetning", 1300),
              cell("Kombinerad analys i grupp + presentation", 2200),
              cell("Grupper om 3\u20134 f\u00E5r varsitt case (AI-driven konspirationsteori). Analysera med kombinerat analysblad: k\u00E4llkritiska grundfr\u00E5gor + konspirationsteorins anatomi + AI-rollen (28 min). Kort presentation f\u00F6r klassen (8 min).", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("58\u201372 min", 1000),
              cell("Summering", 1300),
              cell("Helklassdiskussion + exit ticket", 2200),
              cell("Diskussion: Vad g\u00F6r AI-drivna konspirationsteorier farligare? Koppla till demokrati. Exit ticket: \u201CF\u00F6rklara hur AI kan anv\u00E4ndas f\u00F6r att sprida en konspirationsteori. Ange minst en k\u00E4lla med korrekt k\u00E4llh\u00E4nvisning.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("72\u201375 min", 1000),
              cell("Fram\u00E5tkoppling", 1300),
              cell("N\u00E4sta lektion", 2200),
              cell("\u201CNu har ni alla verktygen. N\u00E4sta lektion anv\u00E4nder vi dem p\u00E5 riktigt: vi f\u00F6rbereder oss f\u00F6r ett seminarium om informationskrig, demokrati och AI.\u201D", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Uppstart (0\u20138 min)"),
      boldBodyText("Retrieval practice (0\u20135 min): ", ""),
      italicText("S\u00E4g: \u201CF\u00F6rra g\u00E5ngen pratade vi om konspirationsteorier \u2014 hur de uppst\u00E5r och sprids. Nu vill jag se vad som sitter kvar. Skriv ner tv\u00E5 anledningar till att konspirationsteorier sprids. Ni har 2 minuter. Skriv fritt, inga hj\u00E4lpmedel.\u201D"),
      bodyText("Samla svar muntligt. Fyll i eventuella luckor. Om exit tickets fr\u00E5n lektion 2 visade att en specifik mekanism var sv\u00E5r (t.ex. confirmation bias, ekkammare), l\u00E4gg extra tid h\u00E4r."),
      spacer(),
      boldBodyText("Hook (5\u20138 min): ", ""),
      bodyText("Visa p\u00E5 sk\u00E4rmen:"),
      italicText("\u201CNewsGuard har identifierat \u00F6ver 2 000 fejkade nyhetssajter som drivs helt av AI. Sajterna ser ut som riktig lokaljournalistik \u2014 men ingen m\u00E4nniska har skrivit en enda artikel.\u201D"),
      spacer(),
      italicText("S\u00E4g: \u201CF\u00F6rra lektionen l\u00E4rde vi oss hur konspirationsteorier fungerar \u2014 och ni fick en f\u00F6rsta introduktion till k\u00E4llh\u00E4nvisning. Lektionen innan dess fick ni verktyg f\u00F6r att granska k\u00E4llor. Idag sl\u00E5r vi ihop allt: Vad h\u00E4nder n\u00E4r AI anv\u00E4nds f\u00F6r att skapa och sprida konspirationsteorier?\u201D"),
      bodyText("Visa dagens l\u00E4randem\u00E5l p\u00E5 tavlan."),

      heading3("Instruktion: Tre verkliga fall (8\u201322 min)"),
      bodyText("Presentera tre verkliga fall d\u00E4r AI har anv\u00E4nts f\u00F6r att skapa eller sprida konspirationsteorier/desinformation. Max 14 minuter. F\u00F6r varje fall: visa, analysera med b\u00E5de grundfr\u00E5gorna och konspirationsteorins anatomi."),
      spacer(),
      bodyText("Skriv p\u00E5 tavlan:"),
      bullet("K\u00E4llkritiska grundfr\u00E5gor: Vem? Varf\u00F6r? Hur? N\u00E4r?"),
      bullet("Konspirationsteorins anatomi: Hemlig grupp? Stor plan? \u201CBevis\u201D som alltid bekr\u00E4ftar?"),
      bullet("NY FR\u00C5GA: Vilken roll spelar AI?"),
      spacer(),

      boldBodyText("Fall 1: Southport-kravallerna och AI-bilder (4 min)", ""),
      bodyText("Sommaren 2024: Tre flickor d\u00F6dades i en knivattack i Southport, England. Inom timmar spreds falska p\u00E5st\u00E5enden om att g\u00E4rningsmannen var en asylans\u00F6kare \u2014 n\u00E5got som var helt felaktigt."),
      bullet("AI-rollen: 39 AI-genererade bilder spreds p\u00E5 sociala medier och f\u00F6rest\u00E4llde rasistiska konspirationsteorier. Trots att de bara utgjorde 10% av inneh\u00E5llet fick de i genomsnitt 1,14 miljoner visningar per inl\u00E4gg."),
      bullet("Konsekvens: Kravaller i flera brittiska st\u00E4der \u2014 den v\u00E4rsta v\u00E5gen av h\u00F6gerextremistiskt v\u00E5ld i Storbritannien sedan andra v\u00E4rldskriget."),
      bullet("Analys: Vem? Anonyma konton. Varf\u00F6r? Sprida hat och mobilisera. Hur? AI-bilder som v\u00E4cker starka k\u00E4nslor. N\u00E4r? Omedelbart efter en kris, n\u00E4r k\u00E4nslorna \u00E4r som starkast."),
      spacer(),

      boldBodyText("Fall 2: AI-r\u00F6stsamtal inf\u00F6r USA:s prim\u00E4rval (3 min)", ""),
      bodyText("Januari 2024: Upp till 25 000 v\u00E4ljare i New Hampshire fick robotsamtal med en AI-genererad r\u00F6st som l\u00E4t som president Biden. R\u00F6sten uppmanade dem att INTE r\u00F6sta."),
      bullet("AI-rollen: Deepfake-r\u00F6st som var n\u00E4stan om\u00F6jlig att skilja fr\u00E5n den riktiga Biden."),
      bullet("Konsekvens: Direkt f\u00F6rs\u00F6k att p\u00E5verka ett demokratiskt val."),
      bullet("Analys: Vem? Ok\u00E4nd uppdragsgivare. Varf\u00F6r? P\u00E5verka valresultatet. Hur? AI-genererad r\u00F6st via telefon \u2014 personlig och direkt. N\u00E4r? Dagarna f\u00F6re valet."),
      spacer(),

      boldBodyText("Fall 3: AI-drivna nyhetssajter (3 min)", ""),
      bodyText("NewsGuard har identifierat \u00F6ver 2 000 sajter som ser ut som seri\u00F6sa nyhetssajter men d\u00E4r allt inneh\u00E5ll \u00E4r AI-genererat. Ingen redaktion, inga riktiga journalister. Bland dessa finns 167 sajter med ryska kopplingar som publicerar falska narrativ om Ukrainakriget."),
      bullet("AI-rollen: Skapa trov\u00E4rdiga nyhetssajter i stor skala, automatiskt och billigt."),
      bullet("Konsekvens: N\u00E4r lokaljournalistiken d\u00F6r ers\u00E4tts den av AI-fejk \u2014 och l\u00E4saren m\u00E4rker det inte."),
      bullet("Analys: Vem? Ok\u00E4nda akt\u00F6rer, ibland statliga. Varf\u00F6r? P\u00E5verka opinion, tj\u00E4na pengar p\u00E5 annonser. Hur? AI genererar tusentals artiklar automatiskt. N\u00E4r? P\u00E5g\u00E5r just nu."),
      spacer(),

      boldItalicBody("EPA-stopp (4 min): ", "\u201CT\u00E4nk enskilt i 1 minut: Vad g\u00F6r AI-driven desinformation farligare \u00E4n \u2018vanlig\u2019 desinformation? Diskutera med din granne i 2 minuter.\u201D Samla 2\u20133 svar. Skriv nyckelord p\u00E5 tavlan (t.ex. skala, hastighet, trov\u00E4rdighet, kostnad, personalisering)."),

      heading3("Bearbetning: Kombinerad analys (22\u201358 min)"),
      bodyText("Dela in eleverna i grupper om 3\u20134. Varje grupp f\u00E5r ETT case (se bilaga 1) och det kombinerade analysbladet (se bilaga 2)."),
      spacer(),
      italicText("Instruktion: \u201CNi f\u00E5r ett verkligt fall d\u00E4r AI och konspirationsteorier m\u00F6ts. Ert jobb \u00E4r att analysera fallet med ALLA verktyg ni l\u00E4rt er hittills: de k\u00E4llkritiska grundfr\u00E5gorna, konspirationsteorins anatomi OCH AI:s roll. Fyll i analysbladet tillsammans. Ni har 28 minuter.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 skriv p\u00E5 tavlan:"),
      bullet("Steg 1: L\u00E4s igenom fallet tillsammans (5 min)"),
      bullet("Steg 2: Besvara de k\u00E4llkritiska grundfr\u00E5gorna (7 min)"),
      bullet("Steg 3: Identifiera konspirationsteorins k\u00E4nnetecken (8 min)"),
      bullet("Steg 4: Analysera AI:s specifika roll (8 min)"),
      spacer(),
      bodyText("Cirkulera mellan grupperna. St\u00E4ll f\u00F6rdjupande fr\u00E5gor:"),
      bullet("Till grupper som fastnar: \u201CB\u00F6rja med grundfr\u00E5gorna \u2014 vem ligger bakom?\u201D"),
      bullet("Till snabba grupper: \u201CKunde den h\u00E4r teorin ha spridits lika snabbt utan AI? Varf\u00F6r/varf\u00F6r inte?\u201D"),
      bullet("Utmaning: \u201CVad beh\u00F6ver samh\u00E4llet g\u00F6ra f\u00F6r att skydda sig mot den h\u00E4r typen av p\u00E5verkan?\u201D"),
      spacer(),
      boldBodyText("Presentation (8 min): ", "Varje grupp presenterar sitt fall f\u00F6r klassen (max 2 minuter per grupp). \u00D6vriga elever f\u00E5r st\u00E4lla EN f\u00F6ljdfr\u00E5ga."),
      spacer(),
      bodyText("Sammanfatta p\u00E5 tavlan efter presentationerna:"),
      italicText("\u201CVilka m\u00F6nster ser vi? AI g\u00F6r konspirationsteorier [snabbare/billigare/mer trov\u00E4rdiga/sv\u00E5rare att avsl\u00F6ja/mer personliga]. Vilka av era verktyg fungerade b\u00E4st?\u201D"),

      heading3("Summering (58\u201372 min)"),
      bodyText("Helklassdiskussion (5 min):"),
      bullet("\u201CVad g\u00F6r AI-drivna konspirationsteorier farligare \u00E4n \u2018vanliga\u2019 konspirationsteorier?\u201D"),
      bullet("\u201CVilket av era verktyg \u2014 grundfr\u00E5gorna eller konspirationsteorins anatomi \u2014 k\u00E4nns viktigast?\u201D"),
      bullet("\u201CVad inneb\u00E4r det h\u00E4r f\u00F6r demokratin?\u201D"),
      spacer(),
      bodyText("Koppla tillbaka till l\u00E4randem\u00E5len p\u00E5 tavlan."),
      spacer(),
      boldBodyText("Exit ticket (7 min):", ""),
      italicText("\u201CF\u00F6rklara hur AI kan anv\u00E4ndas f\u00F6r att skapa eller sprida en konspirationsteori. Ge ett konkret exempel. Ange minst en k\u00E4lla med korrekt k\u00E4llh\u00E4nvisning (Harvardsystemet).\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 meningsstartare p\u00E5 tavlan:"),
      bullet("\u201CAI kan anv\u00E4ndas f\u00F6r att sprida konspirationsteorier genom att...\u201D"),
      bullet("\u201CEtt exempel p\u00E5 detta \u00E4r n\u00E4r...\u201D"),
      bullet("\u201CDetta \u00E4r problematiskt f\u00F6r demokratin eftersom...\u201D"),
      bullet("\u201CEnligt [k\u00E4lla] (Efternamn, \u00E5r) s\u00E5...\u201D"),
      spacer(),
      bodyText("Samla in. Anv\u00E4nd svaren f\u00F6r att bed\u00F6ma: (a) f\u00F6rst\u00E5r eleven kopplingen AI + konspirationsteorier? (b) kan eleven k\u00E4llh\u00E4nvisa? Planera seminariegrupperna utifr\u00E5n detta."),

      heading3("Fram\u00E5tkoppling (72\u201375 min)"),
      italicText("\u201CNu har ni alla verktygen: grundfr\u00E5gorna, konspirationsteorins anatomi och f\u00F6rst\u00E5else f\u00F6r AI:s roll. N\u00E4sta lektion \u00E4r det dags att anv\u00E4nda allt p\u00E5 riktigt: vi f\u00F6rbereder oss f\u00F6r seminariet \u2018Informationskriget\u2019 \u2014 d\u00E4r ni ska diskutera vad AI-driven desinformation inneb\u00E4r f\u00F6r demokratin. B\u00F6rja fundera p\u00E5: Vilken sida av fr\u00E5gan vill ni argumentera f\u00F6r?\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval practice: skriva ner spridningsorsaker (enskilt, 3 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "EPA: vad g\u00F6r AI-driven desinformation farligare? (enskilt + par, 4 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Kombinerad analys med analysblad (grupp 3\u20134, 28 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Gruppresentation f\u00F6r klassen (grupp, 2 min per grupp)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Exit ticket: AI + konspirationsteori + k\u00E4llh\u00E4nvisning (enskilt, 7 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 52 av 75 minuter (69%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Analysbladet har ifyllda exempelf\u00E4lt f\u00F6r varje del (grundfr\u00E5gor, anatomi, AI-roll). Stegen f\u00F6r analysen \u00E4r synliga p\u00E5 tavlan med tids\u00E5tg\u00E5ng. Exit ticket har meningsstartare och en k\u00E4llh\u00E4nvisningsmall (eleverna har sett Harvardsystemet i lektion 2). L\u00E4raren prioriterar dessa grupper vid cirkulering och ger konkreta ledtr\u00E5dar: \u201CTitta p\u00E5 fall 1 som vi gick igenom \u2014 hur liknar ert fall det?\u201D"),
      boldBodyText("Utmaning (mot A): ", "Analysera utan analysblad \u2014 fritt resonemang. Till\u00E4ggsfr\u00E5ga vid grupparbetet: \u201CKunna detta f\u00F6rebyggas med lagstiftning, eller \u00E4r det en fr\u00E5ga om utbildning och medvetenhet? Argumentera f\u00F6r b\u00E5da sidor.\u201D Vid presentationen: l\u00E4gg till en reflektion om demokratiska konsekvenser. Exit ticket utan meningsstartare, med till\u00E4gget: \u201CResonera om avv\u00E4gningen mellan yttrandefrihet och skydd mot AI-driven desinformation. Anv\u00E4nd minst tv\u00E5 k\u00E4llor.\u201D"),

      // MATERIAL
      heading2("Material"),
      bullet("Exit ticket-data fr\u00E5n lektion 2 (f\u00F6r retrieval practice)"),
      bullet("Presentation med tre verkliga fall (se l\u00E4rarinstruktioner ovan, detaljer i bilaga 1)"),
      bullet("Analysblad \u201CAI + konspirationsteori \u2014 kombinerad analys\u201D \u2014 en per elev (se bilaga 2)"),
      bullet("L\u00E4randem\u00E5len utskrivna/projicerade"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("Retrieval practice repeterar konspirationsteoriernas spridningsmekanismer och f\u00F6rst\u00E4rker l\u00E5ngtidsminnet (m\u00E5l 3)"),
      bullet("De tre verkliga fallen och EPA-stoppet tr\u00E4nar resonemang om hur AI f\u00F6r\u00E4ndrar informationslandskapet (m\u00E5l 2, 3)"),
      bullet("Den kombinerade analysen tr\u00E4nar f\u00F6rm\u00E5gan att granska k\u00E4llor med alla verktyg samtidigt (m\u00E5l 1, 2, 3: E\u2013A)"),
      bullet("Exit ticket m\u00E4ter syntes-f\u00F6rm\u00E5ga (AI + konspirationsteori) och \u00F6var k\u00E4llh\u00E4nvisning fr\u00E5n lektion 2 (m\u00E5l 2, 3)"),
      spacer(),
      bodyText("Analysens scaffolding (ifyllda exempel, steg p\u00E5 tavlan) s\u00E4kerst\u00E4ller att alla elever kan n\u00E5 E-niv\u00E5. Den ut\u00F6kade analystiden (28 min ist\u00E4llet f\u00F6r 18) ger alla grupper m\u00F6jlighet att n\u00E5 djupare. De \u00F6ppna fr\u00E5gorna om lagstiftning, yttrandefrihet och demokrati ger A-elever m\u00F6jlighet att visa v\u00E4lgrundade och nyanserade resonemang med komplexa samband."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: Lektion 1 (k\u00E4llkritiska grundfr\u00E5gor) + Lektion 2 (konspirationsteorier)"),
      bullet("Sista \u201Cverktygslektionen\u201D \u2014 fr\u00E5n lektion 4 till\u00E4mpas allt i seminarieform"),
      bullet("N\u00E4sta lektion: \u201CInformationskriget\u201D \u2014 seminarium om AI, desinformation och demokrati"),
      bullet("Exit ticket-data anv\u00E4nds f\u00F6r att planera seminariegrupper i lektion 4"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: CASE-KORT
      heading2("Bilaga 1: Case-kort \u2014 AI-drivna konspirationsteorier"),
      bodyText("Dela ut ett case per grupp. Varje kort beskriver ett verkligt fall d\u00E4r AI anv\u00E4nts f\u00F6r att skapa eller sprida en konspirationsteori eller desinformation."),
      spacer(),

      boldBodyText("Case A: \u201CSouthport-kravallerna \u2014 AI-bilder som v\u00E5ld\u201D", ""),
      bodyText("Bakgrund: Den 29 juli 2024 d\u00F6dades tre flickor i en knivattack i Southport, England. Inom n\u00E5gra timmar spreds falska p\u00E5st\u00E5enden om att g\u00E4rningsmannen var en muslimsk asylans\u00F6kare \u2014 vilket var helt felaktigt."),
      bodyText("AI:s roll: 39 AI-genererade bilder skapades och spreds p\u00E5 sociala medier. Bilderna f\u00F6rest\u00E4llde rasistiska scenarier kopplade till den falska ber\u00E4ttelsen. Trots att de bara utgjorde 10% av det spridda inneh\u00E5llet fick de i genomsnitt 1,14 miljoner visningar per inl\u00E4gg."),
      bodyText("Konsekvens: Kravaller i flera brittiska st\u00E4der. Asylboenden attackerades. Det beskrivs som den v\u00E4rsta v\u00E5gen av h\u00F6gerextremistiskt v\u00E5ld i efterkrigstidens Storbritannien."),
      bodyText("K\u00E4lla: GNET Research, \u201CThe Weaponisation of AI: Visual storytelling of the Great Replacement Conspiracy Theory Amid the Southport Riots\u201D (2025)."),
      spacer(),

      boldBodyText("Case B: \u201CDeepfake-Biden: r\u00F6sta inte\u201D", ""),
      bodyText("Bakgrund: I januari 2024, dagarna f\u00F6re prim\u00E4rvalet i New Hampshire (USA), fick upp till 25 000 v\u00E4ljare robotsamtal d\u00E4r en r\u00F6st som l\u00E4t exakt som president Joe Biden uppmanade dem att inte r\u00F6sta i prim\u00E4rvalet."),
      bodyText("AI:s roll: R\u00F6sten var en AI-genererad deepfake \u2014 skapad med r\u00F6stklonings-teknologi. Den var s\u00E5 \u00F6vertygande att m\u00E5nga lyssnare trodde att det var den riktiga presidenten."),
      bodyText("Konsekvens: Direkt f\u00F6rs\u00F6k att undertrycka r\u00F6stning i ett demokratiskt val. Ledde till ny lagstiftning mot AI-genererade robocalls i USA."),
      bodyText("K\u00E4lla: Munich Security Conference, \u201CAI-pocalypse Now? Disinformation, AI, and the Super Election Year\u201D (2024)."),
      spacer(),

      boldBodyText("Case C: \u201C2 000 AI-nyhetssajter \u2014 journalistik utan journalister\u201D", ""),
      bodyText("Bakgrund: Sedan 2023 har NewsGuard identifierat \u00F6ver 2 000 webbsajter som ser ut som seri\u00F6sa nyhetssajter men d\u00E4r allt inneh\u00E5ll \u00E4r AI-genererat. Sajtena publicerar artiklar utan m\u00E4nsklig \u00F6versikt, stjal inneh\u00E5ll fr\u00E5n riktiga medier och presenterar det som eget."),
      bodyText("AI:s roll: Hela produktionskedjan \u00E4r automatiserad \u2014 fr\u00E5n artikelgenerering till publicering. Bland sajtena finns 167 med ryska kopplingar som sprider falska narrativ om Ukrainakriget."),
      bodyText("Konsekvens: N\u00E4r lokaltidningar l\u00E4ggs ner ers\u00E4tts de av AI-fejk. Risken \u00E4r st\u00F6rre \u00E4n 50% att en nyhetssajt som p\u00E5st\u00E5r sig vara lokal i verkligheten \u00E4r fejk."),
      bodyText("K\u00E4lla: NewsGuard, \u201CTracking AI-enabled Misinformation\u201D (2025). newsguardtech.com/special-reports/ai-tracking-center/"),
      spacer(),

      boldBodyText("Case D: \u201CKatastrofbilder som aldrig h\u00E4nde \u2014 AI-bilder under orkanen Helene\u201D", ""),
      bodyText("Bakgrund: Under orkanen Helene i oktober 2024 spreds AI-genererade bilder p\u00E5 sociala medier som f\u00F6rest\u00E4llde dramatiska scener \u2014 \u00F6versv\u00E4mmade st\u00E4der, barn som r\u00E4ddas ur vattnet \u2014 som aldrig hade \u00E4gt rum. Bilderna anv\u00E4ndes f\u00F6r att st\u00F6dja konspirationsteorier om att myndigheter medvetet l\u00E4t katastrofen ske."),
      bodyText("AI:s roll: AI-genererade bilder skapades och spreds inom timmar. De var sv\u00E5ra att skilja fr\u00E5n verkliga katastrofbilder och anv\u00E4ndes som \u201Cbevis\u201D i konspirationsteorier om att FEMA (USA:s krishanteringsmyndighet) saboterade hj\u00E4lpinsatser."),
      bodyText("Konsekvens: Misstro mot myndigheter under en p\u00E5g\u00E5ende kris, risk att m\u00E4nniskor inte f\u00F6ljer evakueringsorder."),
      bodyText("K\u00E4lla: NPR, \u201CAI images of hurricanes and disasters are going viral. They can also become propaganda\u201D (2024). npr.org/2024/10/18/nx-s1-5153741/"),

      // BILAGA 2: KOMBINERAT ANALYSBLAD
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Analysblad \u2014 AI + konspirationsteori, kombinerad analys"),
      bodyText("Fyll i alla tre delar f\u00F6r ert case."),
      spacer(),

      heading3("Del 1: K\u00E4llkritiska grundfr\u00E5gor"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 7026],
        rows: [
          new TableRow({
            children: [
              headerCell("Grundfr\u00E5ga", 2000),
              headerCell("Ert svar", 7026),
            ],
          }),
          new TableRow({
            children: [
              cell("Vem? (avs\u00E4ndare)", 2000),
              cell("Exempel: Anonyma konton p\u00E5 X/Twitter som delade AI-genererade bilder", 7026),
            ],
          }),
          new TableRow({
            children: [
              cell("Varf\u00F6r? (syfte)", 2000),
              cell("", 7026),
            ],
          }),
          new TableRow({
            children: [
              cell("Hur? (presentation)", 2000),
              cell("", 7026),
            ],
          }),
          new TableRow({
            children: [
              cell("N\u00E4r? (aktualitet)", 2000),
              cell("", 7026),
            ],
          }),
        ],
      }),

      spacer(),
      heading3("Del 2: Konspirationsteorins anatomi"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2800, 6226],
        rows: [
          new TableRow({
            children: [
              headerCell("K\u00E4nnetecken", 2800),
              headerCell("Ert svar", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Hemlig grupp?", 2800),
              cell("Exempel: \u201CEliten/staten d\u00F6ljer sanningen om g\u00E4rningsmannen\u201D", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Stor, d\u00F6ljd plan?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("\u201CBevis\u201D som alltid bekr\u00E4ftar?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vilka k\u00E4nslor spelas det p\u00E5?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vilka spridningsmekanismer?", 2800),
              cell("", 6226),
            ],
          }),
        ],
      }),

      spacer(),
      heading3("Del 3: AI:s roll"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2800, 6226],
        rows: [
          new TableRow({
            children: [
              headerCell("Fr\u00E5ga", 2800),
              headerCell("Ert svar", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vilken AI-teknik anv\u00E4ndes?", 2800),
              cell("Exempel: AI-bildgenerering (typ Midjourney/DALL-E)", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vad gjorde AI m\u00F6jligt som inte gick f\u00F6rut?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Hur p\u00E5verkade AI trov\u00E4rdigheten?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Hur p\u00E5verkade AI spridningshastigheten?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Kunde detta ha spridits utan AI? Varf\u00F6r/varf\u00F6r inte?", 2800),
              cell("", 6226),
            ],
          }),
        ],
      }),

      spacer(),
      boldBodyText("Ordbank: ", "deepfake, AI-genererad bild, r\u00F6stkloning, textgenerering, automatisering, skalbarhet, algoritm, ekkammare, confirmation bias, informationsp\u00E5verkan, demokrati, yttrandefrihet, k\u00E4llkritik, desinformation, propaganda"),

      // BILAGA 3: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Exit ticket \u2014 Lektion 3"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "F\u00F6rklara hur AI kan anv\u00E4ndas f\u00F6r att skapa eller sprida en konspirationsteori. Ge ett konkret exempel. Ange minst en k\u00E4lla med korrekt k\u00E4llh\u00E4nvisning (Harvardsystemet)."),
      spacer(),
      bodyText("Meningsstartare:"),
      bullet("\u201CAI kan anv\u00E4ndas f\u00F6r att sprida konspirationsteorier genom att...\u201D"),
      bullet("\u201CEtt exempel p\u00E5 detta \u00E4r n\u00E4r...\u201D"),
      bullet("\u201CDetta \u00E4r problematiskt f\u00F6r demokratin eftersom...\u201D"),
      bullet("\u201CEnligt [k\u00E4lla] (Efternamn, \u00E5r) s\u00E5...\u201D"),
      spacer(),
      bodyText("K\u00E4llh\u00E4nvisningsmall:"),
      boldBodyText("I texten: ", "(Efternamn/Organisation, \u00E5r)"),
      boldBodyText("I referenslistan: ", "Efternamn/Organisation (\u00E5r). Titel. K\u00E4lla. URL"),

      // BILAGA 4: KÄLLMATERIAL OCH LÄNKAR
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: K\u00E4llmaterial och l\u00E4nkar"),
      bodyText("Samlade l\u00E4nkar f\u00F6r case-korten, l\u00E4rarinstruktionerna och f\u00F6rdjupning."),

      heading3("Verkliga fall: AI och desinformation"),
      boldBodyText("Southport-kravallerna \u2014 AI-bilder och rasistiska konspirationsteorier: ", "gnet-research.org/2025/07/16/the-weaponisation-of-ai-visual-storytelling-of-the-great-replacement-conspiracy-theory-amid-the-southport-riots/"),
      boldBodyText("Biden deepfake-r\u00F6stsamtal \u2014 Munich Security Conference: ", "securityconference.org/en/publications/analyses/ai-pocalypse-disinformation-super-election-year/"),
      boldBodyText("AI-genererade nyhetssajter \u2014 NewsGuard: ", "newsguardtech.com/special-reports/ai-tracking-center/"),
      boldBodyText("AI-katastrofbilder under orkanen Helene \u2014 NPR: ", "npr.org/2024/10/18/nx-s1-5153741/ai-images-hurricanes-disasters-propaganda"),
      boldBodyText("2024 Election Misinformation \u2014 DISA: ", "disa.org/2024-election-misinformation-and-ai-generated-hoaxes-a-review/"),
      boldBodyText("AI-driven desinformation \u2014 \u00F6versikt (ADL): ", "adl.org/resources/article/mis-and-disinformation-trends-and-tactics-watch-2025"),

      heading3("Svenska resurser: AI, k\u00E4llkritik och desinformation"),
      boldBodyText("MPF \u2014 AI och informationsp\u00E5verkan: ", "mpf.se/kunskap-och-stod/temabibliotek/ai-och-informationspaverkan"),
      boldBodyText("MPF \u2014 K\u00E4llkritikens dag: ", "mpf.se/om-oss/nyheter/2025/2025-03-13-kallkritikens-dag-13-mars"),
      boldBodyText("Internetkunskap \u2014 s\u00E5 p\u00E5verkar AI k\u00E4llkritiken: ", "internetkunskap.se/artiklar/grundkurs-i-ai/sa-paverkar-ai-kallkritiken-pa-natet/"),
      boldBodyText("Internetstiftelsen \u2014 AI i skolan: ", "internetstiftelsen.se/nyheter/ai-i-skolan/"),
      boldBodyText("S\u00E4kerhetskollen \u2014 vad \u00E4r deepfake?: ", "sakerhetskollen.se/bli-trygg-pa-internet/sakerhet-pa-sociala-medier/vad-ar-deepfake"),
      boldBodyText("Digitala lektioner \u2014 k\u00E4llkritik: ", "digitalalektioner.se/amnesomrade/digital-kallkritik/"),
      boldBodyText("K\u00E4llkritikbyr\u00E5n (Sveriges faktagranskare): ", "kallkritikbyran.se"),
      boldBodyText("Krisinformation.se \u2014 k\u00E4llkritik: ", "krisinformation.se/detta-gor-samhallet/kallkritik/"),

      heading3("K\u00E4llh\u00E4nvisning \u2014 Harvardsystemet"),
      boldBodyText("Guide till Harvardsystemet (H\u00F6gskolan i Bor\u00E5s): ", "hb.se/biblioteket/skriva-och-referera/referera-till-kallor/guide-till-harvardsystemet/"),
      boldBodyText("Studentapan \u2014 Harvard referens: ", "studentapan.se/studieteknik/harvard-referens-kallhanvisningar-en-guide-for-harvardsystemet"),
      boldBodyText("Uppsala universitet \u2014 k\u00E4llh\u00E4nvisningar i l\u00F6pande text: ", "libguides.ub.uu.se/harvard/kallhanvisning"),
      boldBodyText("Enskilda gymnasiet \u2014 referera: ", "enskildagymnasiet.se/referera/"),
      boldBodyText("Slottegymnasiet \u2014 k\u00E4llh\u00E4nvisning: ", "slottegymnasiet.se/vartbibliotek/kallhanvisningochreferenser"),

      heading3("Faktagranskning \u2014 \u00F6vriga resurser"),
      boldBodyText("Snopes \u2014 faktagranskning: ", "snopes.com"),
      boldBodyText("SO-rummet \u2014 k\u00E4llkritik: ", "so-rummet.se/content/kallkritik-fyra-viktiga-kriterier"),
      boldBodyText("EU-kommissionens faktakoll (Sverige): ", "sweden.representation.ec.europa.eu/nyheter-och-evenemang/faktakoll_sv"),
      boldBodyText("Riksdagsmotion om AI och deepfakes: ", "riksdagen.se/sv/dokument-och-lagar/dokument/motion/markning-av-ai-genererat-material-for-att-motverka_hd021815/"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-3.docx", buffer);
  console.log("lektion-3.docx skapad!");
});
