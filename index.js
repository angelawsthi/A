function showMessage() {
    alert("GO STUDY TORTS RIGHT NOW 😭");
}
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, PageBreak, VerticalAlign
} = require('docx');
const fs = require('fs');

const NAVY   = "1A2C5B";
const GOLD   = "B8860B";
const LIGHT  = "EAF0FB";
const ACCENT = "2E4A8E";
const RULE   = "C0C8DC";
const GREEN  = "0A4A2E";
const LGREEN = "E6F4EC";

const b1 = { style: BorderStyle.SINGLE, size: 1, color: RULE };
const borders = { top: b1, bottom: b1, left: b1, right: b1 };
const noB = { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };
const noBorders = { top: noB, bottom: noB, left: noB, right: noB };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, color: NAVY, font: "Georgia", size: 36, bold: true })],
    spacing: { before: 480, after: 200 },
    border: { bottom: { style: BorderStyle.THICK, size: 8, color: GOLD, space: 6 } }
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, color: ACCENT, font: "Georgia", size: 26, bold: true })],
    spacing: { before: 320, after: 120 }
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, color: GOLD, font: "Georgia", size: 22, bold: true, italics: true })],
    spacing: { before: 200, after: 80 }
  });
}
function body(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Palatino Linotype", size: 22 })],
    spacing: { before: 60, after: 80 },
    alignment: AlignmentType.JUSTIFIED,
  });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, font: "Palatino Linotype", size: 22 })],
    spacing: { before: 40, after: 60 },
  });
}
function sub(label, text) {
  return new Paragraph({
    children: [
      new TextRun({ text: label + ": ", font: "Palatino Linotype", size: 22, bold: true, color: ACCENT }),
      new TextRun({ text, font: "Palatino Linotype", size: 22 }),
    ],
    spacing: { before: 60, after: 80 },
    alignment: AlignmentType.JUSTIFIED,
  });
}
function cn(name, ratio) {
  return new Paragraph({
    children: [
      new TextRun({ text: name + " — ", font: "Palatino Linotype", size: 21, bold: true, italics: true, color: NAVY }),
      new TextRun({ text: ratio, font: "Palatino Linotype", size: 21, italics: true }),
    ],
    spacing: { before: 50, after: 70 },
    alignment: AlignmentType.JUSTIFIED,
    indent: { left: 360 },
  });
}
function src(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Calibri", size: 18, color: "888888", bold: true })],
    spacing: { before: 0, after: 60 },
    alignment: AlignmentType.RIGHT,
  });
}
function rule() {
  return new Paragraph({
    children: [new TextRun({ text: "" })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: RULE } },
    spacing: { before: 80, after: 80 },
  });
}
function sp() { return new Paragraph({ children: [new TextRun({ text: "" })], spacing: { before: 30, after: 30 } }); }

function tCell(text, fill, width = 4680) {
  return new TableCell({
    borders, shading: { fill, type: ShadingType.CLEAR },
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: [new Paragraph({ children: [new TextRun({ text, font: "Palatino Linotype", size: 20 })], alignment: AlignmentType.JUSTIFIED })]
  });
}
function hCell(text, width = 4680) {
  return new TableCell({
    borders, shading: { fill: NAVY, type: ShadingType.CLEAR },
    width: { size: width, type: WidthType.DXA },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({ children: [new TextRun({ text, color: "FFFFFF", bold: true, font: "Georgia", size: 20 })], alignment: AlignmentType.CENTER })]
  });
}

// ─── Title Page ──────────────────────────────────────────────────────────────
function titlePage() {
  return [
    sp(), sp(), sp(),
    new Paragraph({ children: [new TextRun({ text: "CONSTITUTIONAL LAW", font: "Georgia", size: 64, bold: true, color: NAVY })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
    new Paragraph({ children: [new TextRun({ text: "Examination Notes — Volume II", font: "Georgia", size: 36, color: GOLD, bold: true })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
    new Paragraph({ children: [new TextRun({ text: "──────────────────────────────", font: "Georgia", size: 28, color: RULE })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
    new Paragraph({ children: [new TextRun({ text: "Topics Covered:", font: "Georgia", size: 26, color: ACCENT, bold: true })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
    new Paragraph({ children: [new TextRun({ text: "B. Executive Power: President & Governor  |  C. Judiciary: SC & HC Jurisdiction  |  Appointment & Independence", font: "Palatino Linotype", size: 24, color: NAVY })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
    new Paragraph({ children: [new TextRun({ text: "Sources: D.D. Basu  |  M.P. Jain  |  P.M. Bakshi", font: "Calibri", size: 22, color: "888888" })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
    new Paragraph({ children: [new TextRun({ text: "Articles 53, 74, 77, 123, 124, 131–136, 143, 154, 163, 213, 214–217, 226, 227", font: "Calibri", size: 20, color: "888888" })], alignment: AlignmentType.CENTER }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ═══════════════════════════════════════════════════════════════════════════
// SECTION B — EXECUTIVE POWER
// ═══════════════════════════════════════════════════════════════════════════
function sectionB() {
  return [
    h1("B. Executive Power: Position and Power of President and Governor"),
    rule(),

    h2("I. The President — Constitutional Framework"),
    body("Under Article 53(1), the executive power of the Union is vested in the President and shall be exercised by her either directly or through officers subordinate to her in accordance with the Constitution. The supreme command of the Defence Forces is also vested in the President [Article 53(2)], to be regulated by law."),
    src("M.P. Jain | D.D. Basu | P.M. Bakshi"),

    h3("A. Parliamentary Form — President as Constitutional Head"),
    body("India has adopted the British parliamentary model. The President is a formal or constitutional head — not a real executive. The real executive power vests in the Prime Minister and the Council of Ministers (Cabinet)."),
    cn("Ram Jawaya Kapoor v State of Punjab, AIR 1955 SC 549", "The President is only 'a formal or constitutional head of the executive' and the real executive powers are vested in the Ministers or the Cabinet. The Cabinet is 'a hyphen which joins, a buckle which fastens' the legislative part of the State to the executive part."),
    cn("Samsher Singh v State of Punjab, AIR 1974 SC 2192 (7-judge bench)", "The President is not to be personally satisfied in exercising executive power. Whenever the Constitution requires the 'satisfaction' of the President for any power, it is not her personal satisfaction but the satisfaction of the Council of Ministers in the constitutional sense."),
    cn("UNR Rao v Indira Gandhi, AIR 1971 SC 1002", "The Constituent Assembly did not choose the Presidential system. British constitutional conventions governing the Crown–Ministers relationship are pertinent to the Indian Constitution."),
    src("M.P. Jain"),

    h3("B. Article 74 — Council of Ministers to Aid and Advise"),
    body("Article 74(1) (as amended by the 42nd Amendment, 1976) provides: 'There shall be a Council of Ministers with the Prime Minister at the head to AID AND ADVISE the President who SHALL, in the exercise of his functions, ACT IN ACCORDANCE with such advice.' The word 'shall' makes it mandatory. The President is bound to act on ministerial advice."),
    body("Article 74(2) makes the advice tendered by Ministers to the President non-justiciable — no court can inquire into what advice was given."),
    sub("Pre-1976 position", "Before the 42nd Amendment, Article 74(1) only said the Council was to 'aid and advise'. There was no express provision making advice binding, but this was the convention and the Supreme Court had already recognised the President as a titular head."),
    sub("Post-42nd Amendment", "The word 'shall' was inserted and the President was expressly bound to act on ministerial advice. The 44th Amendment (1978) added a proviso allowing the President to ask the Council to reconsider advice once, but if retendered, the President is bound to act on it."),
    src("M.P. Jain | D.D. Basu"),

    h3("C. Article 77 — Authentication of Orders"),
    body("All executive action of the Union Government is to be expressed in the name of the President [Art. 77(1)]. This is directory, not mandatory — non-compliance does not render an order void; it can be proved by other evidence. When an order is duly authenticated under Article 77(2), it carries an irrebuttable presumption that it was made by the President. The personal application of mind of the President is not necessary for authenticated orders."),
    body("'The decision of any Minister or officer under the Rules of Business [made under Article 77(3)] is the decision of the President.' [Samsher Singh]"),
    src("M.P. Jain"),

    h3("D. Constituent Assembly Debates — Intent Clear"),
    body("Ambedkar: 'The President is merely a nominal figure-head… He represents the nation but does not rule the nation. He has no discretion and no powers of administration at all… He occupies the same position as the King under the English Constitution.'"),
    body("Nehru: 'Power really resided in the Ministry and in the Legislature and not in the President as such. We did not give him any real power, but we have made his position one of authority and dignity.'"),
    src("M.P. Jain"),
    sp(),

    h2("II. Powers and Functions of the President"),

    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 6960],
      rows: [
        new TableRow({ children: [hCell("Category", 2400), hCell("Key Powers", 6960)] }),
        ...[ 
          ["Executive", "Appointment of PM, Ministers, Governors, CJI, Judges, AG, CAG, election commissioners; Makes Rules of Business [Art. 77(3)]; Deploys armed forces."],
          ["Legislative", "Summons, prorogues and dissolves Parliament [Arts. 85, 86]; Addresses Parliament; Gives/withholds assent to Bills; Promulgates Ordinances [Art. 123]; Nominates 12 RS members."],
          ["Financial", "Annual Financial Statement (Budget) laid before Parliament; Recommends Money Bills; Contingency Fund of India."],
          ["Judicial / Clemency", "Power of pardon, reprieve, respite, remission or commutation of sentences [Art. 72]; Pardons include death sentences and military court sentences."],
          ["Emergency", "Proclamation of National Emergency [Art. 352]; State Emergency/President's Rule [Art. 356]; Financial Emergency [Art. 360]."],
          ["Diplomatic", "Receives ambassadors; Ratifies treaties; Represents India in international fora."],
        ].map((r, i) => new TableRow({ children: [tCell(r[0], i%2===0 ? LIGHT : "FFFFFF", 2400), tCell(r[1], i%2===0 ? LIGHT : "FFFFFF", 6960)] }))
      ]
    }),
    sp(),
    src("D.D. Basu | M.P. Jain"),
    sp(),

    h2("III. The Governor — Constitutional Framework"),
    body("Each State has a Governor [Article 153]. The Governor is appointed by the President [Article 155], formally but actually on the advice of the Prime Minister. The Governor may serve one or more States and holds office during the pleasure of the President [Article 156(1)]."),
    body("The Governor has a dual capacity: (1) Head of the State Government; and (2) Representative of the Centre in the State. The Governor's term is 5 years but continues until a successor is appointed."),
    src("M.P. Jain"),

    h3("A. Appointment and Qualifications"),
    bullet("Must be a citizen of India, at least 35 years old [Article 157]."),
    bullet("Cannot be a member of a House of Parliament or State Legislature."),
    bullet("Cannot hold any other office of profit."),
    bullet("Emoluments cannot be reduced during term [Article 158]."),
    body("The Sarkaria Commission recommended that the Centre consult the Chief Minister before appointing the Governor and suggested that the person appointed should be from outside the State, detached from local politics, and not have taken too great a part in recent politics."),
    src("M.P. Jain"),

    h3("B. Article 163 — Council of Ministers to Aid Governor"),
    body("Article 163(1): There shall be a Council of Ministers with the Chief Minister as the head to aid and advise the Governor in the exercise of his functions, EXCEPT in so far as he is by or under the Constitution required to act in his discretion."),
    body("Article 163(2): Where any question arises whether a matter is or is not a matter as respects which the Governor is required to act in his discretion, the decision of the Governor in his discretion shall be FINAL."),
    body("Article 163(3): No court can inquire into what advice was tendered by Ministers to the Governor."),
    sub("Unlike the President", "The Governor has discretionary powers; the President has none. In the non-discretionary area the Governor is bound to act on ministerial advice, just as the President is under Article 74."),
    src("M.P. Jain"),

    h3("C. Governor's Discretionary Powers"),
    body("The circumstances in which discretion may be exercised cannot be exhaustively enumerated [MP Special Police Establishment v State of MP, (2004) 8 SCC 788]. Key discretionary powers include:"),

    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3000, 6360],
      rows: [
        new TableRow({ children: [hCell("Situation", 3000), hCell("Governor's Discretion", 6360)] }),
        ...[
          ["No clear majority in Assembly", "Governor decides whom to invite to form government. If one party has a clear majority, the task is mechanical and non-discretionary."],
          ["Hung assembly / Fluid situation", "After appropriate enquiry, Governor can invite the person most likely to command a majority. [Committee of Governors Guidelines, 1971]"],
          ["Dissolution of Assembly", "May refuse dissolution if an alternative stable government can be formed. Must accept if the incumbent Ministry clearly has majority support."],
          ["Dismissal of Ministry", "Non-controversial: dismissal of a Minister who has lost confidence of CM, or a Ministry that has lost majority. Governor must not act on whim; ascertainment only on floor of the House. [Samsher Singh; Subhash Desai v Governor of Maharashtra, 2023]"],
          ["Report under Art. 356", "Governor's report on breakdown of constitutional machinery. Must be based on objective material — not 'ipse dixit, suspicion or whims'. [Rameshwar Prasad (VI) v UOI, (2006) 2 SCC 1]"],
          ["Reserving Bill for President", "Power to reserve a Bill for the President's consideration is discretionary."],
          ["Sanction to prosecute CM/Minister", "May act in own discretion where bias is inherent or manifest in Cabinet advice, or Cabinet disables itself. [M.P. Jain]"],
        ].map((r, i) => new TableRow({ children: [tCell(r[0], i%2===0 ? LIGHT : "FFFFFF", 3000), tCell(r[1], i%2===0 ? LIGHT : "FFFFFF", 6360)] }))
      ]
    }),
    sp(),
    src("M.P. Jain"),

    h3("D. Governor's Immunity [Article 361]"),
    body("The Governor enjoys personal immunity under Article 361 — no criminal proceedings can be instituted against her during her term of office, and she is not answerable to any court for acts done in exercise of her powers. This immunity is 'complete'. Any mala fide actions may be gone into by the President. [Rameshwar Prasad (VI) v UOI]"),
    sp(),

    h3("E. Key Cases on Governor"),
    cn("Samsher Singh v State of Punjab, AIR 1974 SC 2192", "Definitive ruling (7-judge) that the Governor, like the President, acts in the constitutional sense on the aid and advice of the Council of Ministers. The decision of a Minister or officer under Rules of Business is the decision of the Governor."),
    cn("Nabam Rebia v Dy Speaker, Arunachal Pradesh, (2016) 8 SCC 1", "Governor's power to act without Council of Ministers is of extraordinary nature. Discretion to call for a floor test must be exercised with circumspection, not arbitrarily."),
    cn("Subhash Desai v Governor of Maharashtra, 2023 SCC OnLine SC 607", "Governor had no objective material to call for a floor test; his exercise of discretion 'was not in accordance with Law'. Courts have power to scrutinise Governor's discretionary actions."),
    cn("Rameshwar Prasad (VI) v UOI, (2006) 2 SCC 1", "Dissolution of Bihar Assembly recommended by Governor declared unconstitutional — Governor's report was 'ipse dixit' without verified material. Castigated practice of appointing politically active persons as Governors without cooling period."),
    sp(),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ═══════════════════════════════════════════════════════════════════════════
// SECTION C — JUDICIARY
// ═══════════════════════════════════════════════════════════════════════════
function sectionC1() {
  return [
    h1("C(i). Jurisdiction of the Supreme Court"),
    rule(),

    h2("Overview of Supreme Court Jurisdiction"),
    body("India has a unified judicial system — a single hierarchy culminating in the Supreme Court. The Supreme Court is the guardian of the Constitution, the protector of Fundamental Rights, the final arbiter of constitutional disputes, and the highest court of appeal. Its jurisdictions may be classified as follows:"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 6960],
      rows: [
        new TableRow({ children: [hCell("Jurisdiction", 2400), hCell("Article / Basis", 6960)] }),
        ...[
          ["Original — Fundamental Rights", "Article 32 — Writ jurisdiction for enforcement of Fundamental Rights"],
          ["Extraordinary Original — Inter-Governmental Disputes", "Article 131 — Disputes between Centre & States, or between States"],
          ["Appellate — Constitutional", "Article 132 — Substantial constitutional questions from HC"],
          ["Appellate — Civil", "Article 133 — Substantial questions of law of general importance (HC certificate)"],
          ["Appellate — Criminal", "Article 134 — Death sentence confirmed/reversed, or HC certificate"],
          ["Special Leave", "Article 136 — Discretionary leave from any court/tribunal in India"],
          ["Transfer", "Article 139-A — Transfer of cases between High Courts"],
          ["Advisory", "Article 143 — President's reference of law/fact questions"],
          ["Complete Justice", "Article 142 — Plenary power to do complete justice"],
          ["Court of Record", "Article 129 — Power to punish for contempt of itself"],
        ].map((r, i) => new TableRow({ children: [tCell(r[0], i%2===0 ? LIGHT : "FFFFFF", 2400), tCell(r[1], i%2===0 ? LIGHT : "FFFFFF", 6960)] }))
      ]
    }),
    sp(),

    h2("1. Original Jurisdiction — Article 32 (Writ Jurisdiction)"),
    body("Article 32 is the 'right to constitutional remedies' — a Fundamental Right itself [Part III]. It empowers the Supreme Court to issue writs (habeas corpus, mandamus, prohibition, certiorari, quo warranto) for enforcement of Fundamental Rights. Dr Ambedkar called it 'the very soul of the Constitution and the very heart of it.'"),
    body("Scope: The Supreme Court acts as a 'sentinel on the qui vive' to enforce Fundamental Rights. Article 32(1) cannot be restricted by legislation [Article 32(4) requires a constitutional amendment to suspend it, only during emergency]."),
    body("Doctrine of Horizontality: In Kaushal Kishore v State of UP (2023), a 5-judge bench (3:2) held that some Fundamental Rights (e.g., Art. 15(2), 17, 21) can be enforced even against non-State actors. The dissent (Nagarathna J) cautioned this collapsed distinct doctrines. The matter remains contested."),
    src("M.P. Jain"),

    h2("2. Extraordinary Original Jurisdiction — Article 131"),
    body("Article 131 gives the Supreme Court EXCLUSIVE original jurisdiction in disputes between:"),
    bullet("The Centre and one or more States"),
    bullet("The Centre and a State(s) on one side, and another State(s) on the other"),
    bullet("Two or more States inter se"),
    body("This jurisdiction applies only to disputes involving questions of law or fact on which the existence or extent of a legal right depends. Disputes are justiciable only if they assert a legal right — not mere political or administrative disagreements."),
    body("EXCLUSION: Article 131 jurisdiction is excluded for disputes arising out of pre-constitutional treaties/agreements [State of West Bengal v UOI, AIR 1963 SC 1241]. Disputes involving private parties are also excluded."),
    src("M.P. Jain"),

    h2("3. Appellate Jurisdiction — Articles 132, 133, 134"),

    h3("A. Constitutional Matters [Article 132]"),
    body("An appeal lies to the Supreme Court from any judgment, decree or final order of a High Court (in civil, criminal or other proceedings) if the High Court certifies that the case involves a SUBSTANTIAL QUESTION OF LAW as to the INTERPRETATION OF THE CONSTITUTION."),
    body("The principle underlying Article 132: 'The final authority of interpreting the Constitution must rest with the Supreme Court. With that object the Article is freed from other limitations imposed under Articles 133 and 134 and the right of the wildest amplitude is allowed irrespective of the nature of the proceedings in a case involving only a substantial question of law as to the interpretation of the Constitution.'"),
    body("A 'substantial question' of constitutional law means one on which there is a real difference of opinion — not one already authoritatively settled by the Supreme Court."),
    src("M.P. Jain"),

    h3("B. Civil Matters [Article 133]"),
    body("An appeal lies to the Supreme Court from a judgment, decree or final order in a civil proceeding of a High Court if the High Court certifies: (a) the case involves a substantial question of law of GENERAL IMPORTANCE; AND (b) in the High Court's opinion, the question NEEDS TO BE DECIDED BY THE SUPREME COURT."),
    body("No appeal as of right — only on certificate. The test for 'substantial question': whether it is fairly arguable, or open to argument, or calls for discussion of alternative views. If practically covered by a Supreme Court decision, it is NOT substantial."),
    body("Article 133 discards the distinction between original and appellate jurisdiction of the High Court — it covers all judgments regardless of the HC's capacity."),
    src("M.P. Jain"),

    h3("C. Criminal Matters [Article 134]"),
    body("An appeal to the Supreme Court as a matter of right in criminal cases where a HC: (a) on appeal reversed acquittal and sentenced to death; or (b) has withdrawn a case from a subordinate court and sentenced the accused to death. In other cases, only on a HC certificate that the case is fit for appeal to the Supreme Court."),
    src("M.P. Jain"),

    h2("4. Special Leave Petition — Article 136"),
    body("Article 136(1): 'Notwithstanding anything in this Chapter, the Supreme Court may, in its discretion, grant SPECIAL LEAVE TO APPEAL from any judgment, decree, determination, sentence or order in any cause or matter passed or made by any COURT OR TRIBUNAL in the territory of India.'"),
    body("Exclusion: Judgments of armed forces tribunals [Article 136(2)]."),

    h3("Nature of Article 136 Power"),
    body("Article 136 confers a residuary, plenary, discretionary jurisdiction — a 'sweeping power exercisable outside the purview of ordinary law to meet pressing demands of justice.' It opens with a non-obstante clause — it is unaffected by Articles 132, 133, 134. The Supreme Court has characterised its Article 136 power as 'an untrammelled reservoir of power incapable of being confined to definitional bounds.'"),
    body("But it is DISCRETIONARY — no party has a right to appeal under Article 136. The Court grants leave only where: (i) substantial question of law of general or public importance is involved; OR (ii) manifest injustice results from the impugned order."),
    body("The power is 'exceptional — to be exercised sparingly, with caution and care and to remedy extraordinary situations or situations occasioning gross failure of justice.' [Kunhayammed v State of Kerala]"),
    src("M.P. Jain"),

    h2("5. Advisory Jurisdiction — Article 143"),
    body("Article 143(1): When it appears to the President that a question of law or fact has arisen (or is likely to arise) which is of such a nature and public importance that it is EXPEDIENT to obtain the opinion of the Supreme Court, she may refer it to the Court for consideration. The Court MAY, after such hearing as it thinks fit, report its opinion to the President."),
    body("Article 143(2): Even matters excluded from Article 131 jurisdiction (e.g., pre-constitutional treaties) may be referred under Article 143(2), and in such cases the Court MUST report its opinion."),

    h3("Key Features of Advisory Jurisdiction"),
    bullet("The word 'may' in Article 143(1) means the Court has discretion to decline to answer — for strong, compelling and good reasons (e.g., hypothetical/speculative questions, purely political questions)."),
    bullet("The opinion is delivered in open court by a Bench of not less than 5 judges [Art. 145(3)], with majority concurrence [Art. 145(5)]. Dissenting opinions are permitted."),
    bullet("The President acts on Cabinet advice in making a reference."),
    bullet("The Court cannot be asked to reconsider its earlier decisions under Article 143 — this would convert advisory jurisdiction into appellate jurisdiction over itself."),
    bullet("The scope is wide — the President can refer any question of law or fact of public importance, not limited to Union Government's functions."),
    cn("Re: Special Reference No. 1 of 2012", "The Court has 'well within its jurisdiction to answer/advise the President... if the questions referred are likely to arise in future or such questions are of public importance or there is no decision of this Court which has already decided the question referred.'"),
    sub("Binding force of advisory opinion", "Controversial. H.M. Seervai argues only judgments have precedential value. But the Supreme Court since Re Delhi Laws, AIR 1951 SC 332 has consistently treated advisory opinions as binding law under Article 141."),
    src("M.P. Jain"),

    h2("6. Complete Justice — Article 142"),
    body("Article 142(1): The Supreme Court, in the exercise of its jurisdiction, may pass such decree or make such order as is NECESSARY FOR DOING COMPLETE JUSTICE in any cause or matter pending before it."),
    body("This is a plenary, supplementary, residual source of power. It is 'entirely of a different level and of a different quality' from ordinary law — it is not controlled by statutory provisions but cannot be used to 'supplant' substantive law or to build a new edifice where none existed. [Supreme Court Bar Association v UOI, (1998) 4 SCC 409]"),
    body("Conditions for exercise: (i) SC must have seisin of the matter; (ii) the order must be necessary for complete justice; (iii) cannot contravene a constitutional provision."),
    src("M.P. Jain"),
    sp(),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function sectionC2() {
  return [
    h1("C(i) contd. — Jurisdiction of High Courts"),
    rule(),

    h2("Overview"),
    body("High Courts are the highest courts of the States. Every High Court is a Court of Record [Article 215] and has all the powers of such a court, including the power to punish for contempt. The High Courts have inherent and plenary powers; unless expressly or impliedly barred, and subject to the SC's appellate/discretionary jurisdiction, 'the high courts have unlimited jurisdiction, including jurisdiction to determine their own powers.' [MV Elisabeth v Harwan, AIR 1993 SC 1014]"),
    src("D.D. Basu"),

    h2("1. Article 226 — Writ Jurisdiction (Broader than Article 32)"),
    body("Article 226(1): 'Notwithstanding anything in Article 32... every High Court shall have power, throughout the territories in relation to which it exercises jurisdiction, to issue to any person or authority, including in appropriate cases, any Government, within those territories directions, orders or writs, including writs in the nature of habeas corpus, mandamus, prohibition, quo warranto and certiorari, or any of them, for the ENFORCEMENT OF ANY OF THE RIGHTS CONFERRED BY PART III AND FOR ANY OTHER PURPOSE.'"),

    h3("Key Distinctions: Article 226 vs Article 32"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 3480, 3480],
      rows: [
        new TableRow({ children: [hCell("Feature", 2400), hCell("Article 32 (Supreme Court)", 3480), hCell("Article 226 (High Court)", 3480)] }),
        ...[
          ["Scope", "ONLY for enforcement of Fundamental Rights (Part III)", "For enforcement of Fundamental Rights AND ANY OTHER PURPOSE (broader)"],
          ["Right", "A Fundamental Right itself — cannot be abridged or suspended except during emergency", "Not a Fundamental Right — Parliament can restrict (but cannot abrogate) it"],
          ["Territorial", "Nationwide", "Within the HC's territorial jurisdiction (cause of action test under Art. 226(2))"],
          ["Discretion", "Wide — but the right to move SC under Art. 32 cannot be abrogated", "Highly discretionary — HC may decline to exercise on grounds of alternative remedy, delay, laches"],
        ].map((r, i) => new TableRow({ children: [tCell(r[0], i%2===0 ? LIGHT : "FFFFFF", 2400), tCell(r[1], i%2===0 ? LIGHT : "FFFFFF", 3480), tCell(r[2], i%2===0 ? LIGHT : "FFFFFF", 3480)] }))
      ]
    }),
    sp(),
    src("D.D. Basu | M.P. Jain | P.M. Bakshi"),

    h3("Territorial Jurisdiction under Article 226(2)"),
    body("A HC may issue writs even to a government/authority outside its territories if the CAUSE OF ACTION (wholly or in part) arose within its territorial jurisdiction. This resolves the issue of forum-shopping and ensures access to justice."),

    h3("Limitations on Article 226 Jurisdiction"),
    bullet("Alternative remedy available: HC generally declines writ if adequate alternative statutory remedy exists, unless (a) violation of Fundamental Rights; (b) HC jurisdiction ousted; (c) violation of natural justice; or (d) order patently without jurisdiction."),
    bullet("Delay and laches: Undue delay may defeat writ relief, especially where third-party rights have been acquired."),
    bullet("Disputed facts: HC reluctant to exercise Art. 226 where investigation of disputed facts is required."),
    bullet("Article 226(3): Mandatory obligation on HC to dispose of an application for vacation of ex parte interim order within 2 weeks; failing which the interim order stands vacated automatically."),
    src("D.D. Basu | M.P. Jain"),

    h2("2. Article 227 — Supervisory Jurisdiction"),
    body("Article 227 confers on the High Court superintendence over all courts and tribunals throughout its territories (including courts not within the appellate jurisdiction). This is an administrative and judicial superintendence to prevent grave injustice, or where there is no effective alternative remedy."),
    body("Scope: The HC can call for records, examine proceedings, issue orders to subordinate courts but cannot act as an appellate court on facts or legal errors in exercise of Article 227. Power is used to correct jurisdictional errors, denial of natural justice, or where the subordinate court acts de hors the law."),
    src("M.P. Jain"),

    h2("3. Original Jurisdiction of High Courts"),
    body("Some High Courts (e.g., Calcutta, Bombay, Madras — the old Presidency HCs) retain original civil jurisdiction under their respective Letters Patent. Under Article 225, the jurisdiction of every HC at commencement of the Constitution continues. The original jurisdiction of a HC is determined by the statute creating or governing it."),

    h2("4. Appellate Jurisdiction"),
    body("High Courts hear appeals from subordinate courts in civil and criminal matters. In criminal matters, the HC hears appeals from Sessions Courts and confirms death sentences. The HC is the final court of facts."),

    h2("5. Restriction on HC Jurisdiction — Article 212"),
    body("Under Article 212(1), courts cannot inquire into the validity of proceedings of the Legislature on the ground of mere irregularity of procedure. But courts CAN scrutinise proceedings for substantive or gross illegality or unconstitutionality. The scope of Article 226 is not affected by Article 212 when illegal or unconstitutional conduct is alleged."),
    sp(),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function sectionC3() {
  return [
    h1("C(ii). Appointment and Independence of the Judiciary"),
    rule(),

    h2("I. Appointment of Supreme Court Judges"),
    body("Article 124(2) provides that every Judge of the Supreme Court shall be appointed by the President by warrant under his hand and seal AFTER CONSULTATION with such of the Judges of the Supreme Court and of the High Courts in the States as the President may deem necessary. (The reference to the NJAC was removed following the Supreme Court's ruling in SCAORA v UOI (2015).)"),

    h3("A. Phase 1 — Pre-1993 (SP Gupta — Executive Primacy)"),
    body("In SP Gupta v UOI (First Judges Case, 1981 Supp SCC 516), a 7-judge bench by majority held that 'consultation' in Article 217(1) did not mean 'concurrence'. The opinion of the Chief Justice of India was not binding on the Executive, and the final power of appointment rested with the Union Government."),
    body("Effect: The Executive held primacy in judicial appointments. This coincided with the controversial period of judicial supersession (mid-1970s)."),
    src("M.P. Jain"),

    h3("B. Phase 2 — Post-1993 (Second Judges Case — Judicial Primacy / Collegium)"),
    body("In SC Advocates on Record Association v UOI (Second Judges Case, AIR 1994 SC 268), a 9-judge bench (7:2) OVERRULED the First Judges Case on the crucial question of primacy. The key holdings:"),
    bullet("'Consultation' in Articles 124(2) and 217(1) means effective and meaningful consultation — the CJI's opinion must be given primacy. The opinion is not merely of the CJI individually but as the head of the Judiciary, formed COLLECTIVELY by a body at the apex."),
    bullet("Collegium at SC level: CJI + 2 senior-most puisne Judges."),
    bullet("The opinion of the CJI formed by the collegium is PRIMAL — no appointment can be made contrary to it; if the CJI's recommendation is reiterated unanimously, the Executive is BOUND by it."),
    bullet("Independence of Judiciary is a basic feature of the Constitution [SP Gupta]; the Second Judges Case gave it teeth at the appointment stage."),
    bullet("Constitutional conventions (primacy of CJI's opinion established since 1950) are part of constitutional law and are enforceable."),
    body("Minority dissent (Ahmadi J): Giving primacy to CJI's view through 'consultation' amounted to rewriting the Constitution. The Constitution has not required 'concurrence'. The Executive power in appointment was clear."),
    src("M.P. Jain"),

    h3("C. Phase 3 — 1998 (Third Judges Case — Collegium Expanded)"),
    body("In Re: Special Reference No. 1 of 1998 (Third Judges Case), the President referred nine questions under Article 143. The Supreme Court (11 judges) unanimously answered and clarified the collegium process. Key changes/clarifications:"),
    bullet("Collegium for SC appointments EXPANDED to CJI + 4 senior-most puisne judges (from 2 in Second Judges Case)."),
    bullet("For HC appointments: Chief Justice of HC + 2 senior-most HC judges, with CJI consulting 2 senior-most SC judges."),
    bullet("If CJI is part of a case, the successor CJI (if in line of succession) must also be part of the collegium."),
    bullet("CJI must reduce entire consultative process IN WRITING in communication to the President."),
    bullet("Default: appointments must be made by forming consensus; friction must be averted."),
    bullet("Primacy of CJI's individual opinion is diluted — it is the COLLEGIUM'S opinion that counts."),
    src("M.P. Jain"),

    h3("D. Phase 4 — NJAC (99th Amendment, 2014) and SCAORA v UOI (2015)"),
    body("The National Judicial Appointments Commission (NJAC) was established by the 99th Constitutional Amendment and the NJAC Act, 2014 with the aim of making judicial appointments more transparent and accountable. NJAC composition: CJI (chair) + 2 senior-most SC judges + Union Law Minister + 2 eminent persons nominated by a committee of PM, CJI and Leader of Opposition."),
    body("Problem with NJAC: Section 5(2)(b) and 6(6) of the NJAC Act allowed any 2 members to veto a recommendation. This effectively allowed the Law Minister alone (with one other member) to stall any appointment, fundamentally compromising judicial independence."),
    body("In SCAORA v UOI (IV Judges Case) [Advocates-on-Record Association v UOI, (2016) 5 SCC 1], a 5-judge Constitution Bench by 4:1 STRUCK DOWN the 99th Amendment and the NJAC Act as unconstitutional, holding that they violated the basic feature of INDEPENDENCE OF JUDICIARY."),
    bullet("The collegium system was restored."),
    bullet("The Court recognised that the current collegium system has its own problems (lack of transparency, accountability) and called for its improvement through a Memorandum of Procedure."),
    bullet("Dissent (Chelameswar J): Favoured some executive/public participation in appointments; complete exclusion of the executive was itself a threat to independence."),
    src("M.P. Jain"),

    sp(),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 2450, 2450, 2460],
      rows: [
        new TableRow({ children: [hCell("Phase", 2000), hCell("Case / Law", 2450), hCell("Who Had Primacy?", 2450), hCell("Key Change", 2460)] }),
        ...[
          ["Pre-1993", "SP Gupta (1st Judges Case)", "Union Executive", "CJI's opinion not binding; consultation ≠ concurrence"],
          ["1993", "2nd Judges Case", "CJI's Collegium (CJI + 2 SC judges)", "Consultation = concurrence of CJI collegium; primacy to judiciary"],
          ["1998", "3rd Judges Case", "CJI's Collegium (CJI + 4 SC judges)", "Collegium expanded; written consultation process mandated"],
          ["2014-15", "99th Amdt + NJAC struck down", "Collegium restored", "NJAC unconstitutional; violated basic feature of judicial independence"],
        ].map((r, i) => new TableRow({ children: [tCell(r[0], i%2===0 ? LIGHT : "FFFFFF", 2000), tCell(r[1], i%2===0 ? LIGHT : "FFFFFF", 2450), tCell(r[2], i%2===0 ? LIGHT : "FFFFFF", 2450), tCell(r[3], i%2===0 ? LIGHT : "FFFFFF", 2460)] }))
      ]
    }),
    sp(),

    h2("II. Appointment of High Court Judges"),
    body("Article 217(1): Appointment by the President after consultation with (a) the Chief Justice of India, (b) the Governor of the State, and (c) the Chief Justice of the High Court concerned. The same 'consultation = concurrence' principle applies."),
    body("After Second and Third Judges Cases: The CJI (acting through the collegium) has primacy. No appointment to the HC can be made contrary to the CJI's opinion if it is reiterated unanimously."),
    body("Article 217(2): A person is not qualified unless they are a citizen of India AND has been a Judge of a High Court for at least 10 years, OR an advocate of a High Court for at least 10 years."),
    body("Article 217(3): The President (on reference from the HC) determines the age of a judge where there is a dispute."),
    src("M.P. Jain | D.D. Basu"),

    h2("III. Tenure and Removal of Judges"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2400, 3480, 3480],
      rows: [
        new TableRow({ children: [hCell("Aspect", 2400), hCell("Supreme Court", 3480), hCell("High Court", 3480)] }),
        ...[
          ["Retirement age", "65 years [Art. 124(2)]", "62 years [Art. 217(1)] — disparity exists (Jain notes this is unjustified)"],
          ["Removal", "By President on an address by BOTH Houses of Parliament in same session, by special majority (absolute + 2/3 present & voting), on ground of proved misbehaviour or incapacity [Art. 124(4)]", "Same procedure [Art. 217(1)(b) r/w Art. 124(4)]"],
          ["Inquiry", "Judges (Inquiry) Act, 1968 — Committee of (a) SC Judge, (b) HC Chief Justice, (c) eminent jurist", "Same inquiry process applies"],
          ["Transfer", "SC Judges may be transferred [Art. 222]; consent NOT required by Constitution; only CJI/collegium recommendation required", "HC Judges may be transferred in public interest; CJI must initiate and recommendation is binding on Executive"],
        ].map((r, i) => new TableRow({ children: [tCell(r[0], i%2===0 ? LIGHT : "FFFFFF", 2400), tCell(r[1], i%2===0 ? LIGHT : "FFFFFF", 3480), tCell(r[2], i%2===0 ? LIGHT : "FFFFFF", 3480)] }))
      ]
    }),
    sp(),

    h2("IV. Independence of Judiciary — Constitutional Safeguards"),

    h3("A. What Independence of Judiciary Means"),
    body("Independence of Judiciary is a BASIC FEATURE of the Constitution. [SP Gupta (1st Judges); confirmed in Second Judges Case; SCAORA (2015)]. It means:"),
    bullet("Freedom from Executive control and interference in appointments, transfers, and functioning."),
    bullet("Security of tenure — judges can only be removed by the elaborate impeachment process."),
    bullet("Security of remuneration — salaries, allowances and pensions are charged on the Consolidated Fund and cannot be reduced during service [Articles 125, 221]."),
    bullet("Prohibition on practice after retirement — SC judges cannot practice in any court in India except the SC [Art. 124(7)]. HC judges cannot practice in the HC from which they retired."),
    bullet("Freedom of decision-making — no punishment for decisions made in judicial capacity."),
    src("M.P. Jain | D.D. Basu"),

    h3("B. Institutional Safeguards"),
    bullet("Articles 121/211: No discussion in Parliament/Legislature about conduct of SC/HC judge in discharge of duties, except when a motion for impeachment is under consideration."),
    bullet("Article 129/215: Courts of Record — power to punish for contempt protects their authority."),
    bullet("Article 144: All authorities (civil and judicial) in India must ACT IN AID of the Supreme Court."),
    bullet("Article 361: President and Governors are immune from court proceedings; judges similarly protected by their institutional authority."),
    bullet("Collegium system: Primacy of Chief Justice/Collegium in appointments insulates selections from political influence."),
    src("D.D. Basu | M.P. Jain"),

    h3("C. Challenges to Independence — Critical Notes"),
    body("M.P. Jain critically notes:"),
    bullet("The collegium system itself lacks transparency and accountability — no public justification for most decisions."),
    bullet("Practice of appointing recently active politicians as Governors creates risks of politicisation of the institution — the same concern applies to judicial appointments if seniority norms are departed from without reasons."),
    bullet("Post-retirement employment of judges (tribunals, commissions) may create incentives that colour their judicial decision-making towards the end of tenure."),
    bullet("The 'cooling off period' convention is poorly observed in Governor appointments and needs formalisation."),
    src("M.P. Jain"),
    sp(),

    h2("V. Quick Reference — Key Cases"),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3400, 2200, 3760],
      rows: [
        new TableRow({ children: [hCell("Case"), hCell("Topic"), hCell("Principle")] }),
        ...[
          ["SP Gupta v UOI (1st Judges Case, 1981)", "Appointment", "Consultation ≠ concurrence; Executive had primacy. Now OVERRULED on this point."],
          ["SC Advocates on Record v UOI (2nd Judges Case, 1993)", "Appointment", "Consultation = concurrence; CJI's Collegium (CJI + 2) given primacy. Independence of Judiciary is basic feature."],
          ["Re: Special Reference No. 1 of 1998 (3rd Judges Case)", "Appointment", "Collegium expanded to CJI + 4; written consultation; consultation process must be reduced to writing."],
          ["SCAORA v UOI (IV Judges Case, 2015)", "NJAC", "NJAC and 99th Amendment struck down. Violated basic feature of judicial independence. Collegium restored."],
          ["Ram Jawaya Kapoor v State of Punjab, AIR 1955 SC 549", "Executive Power", "President is only formal/constitutional head; real power vests in Cabinet."],
          ["Samsher Singh v State of Punjab, AIR 1974 SC 2192", "President/Governor", "Satisfaction of President/Governor = satisfaction of Council of Ministers. Decision of Minister/officer under Rules of Business = decision of President/Governor."],
          ["Subhash Desai v Governor of Maharashtra, 2023", "Governor Discretion", "Governor must have objective material before calling floor test. Discretion not unfettered; subject to judicial review."],
          ["Rameshwar Prasad (VI) v UOI, (2006) 2 SCC 1", "Governor", "Governor's report for Art. 356 must be based on verified material, not ipse dixit. Dissolution declared unconstitutional."],
          ["Kaushal Kishore v State of UP, 2023 SCC Online SC 6", "Art. 32", "Horizontality of some Fundamental Rights (majority 3:2) — some rights enforceable against private parties."],
          ["Kunhayammed v State of Kerala, AIR 2000 SC 2587", "Art. 136 SLP", "Art. 136 is discretionary; to be exercised sparingly; substantial question of law or manifest injustice required."],
          ["Supreme Court Bar Assn v UOI, (1998) 4 SCC 409", "Art. 142", "Art. 142 power is supplementary and residual; cannot supplant substantive law or build new edifice against express statutory provisions."],
        ].map((r, i) => new TableRow({ children: [
          new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT : "FFFFFF", type: ShadingType.CLEAR }, width: { size: 3400, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r[0], font: "Palatino Linotype", size: 19, italics: true })] })] }),
          new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT : "FFFFFF", type: ShadingType.CLEAR }, width: { size: 2200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r[1], font: "Palatino Linotype", size: 19, bold: true, color: ACCENT })] })] }),
          new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT : "FFFFFF", type: ShadingType.CLEAR }, width: { size: 3760, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r[2], font: "Palatino Linotype", size: 19 })] })] }),
        ]})),
      ]
    }),
    sp(),
  ];
}

// ═══════════════════════════════════════════════════════════════════════════
// BUILD
// ═══════════════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "◆", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 360 } }, run: { color: GOLD } } },
      ]
    }]
  },
  styles: {
    default: { document: { run: { font: "Palatino Linotype", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Georgia" },
        paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Georgia" },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Georgia" },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1260, right: 1260, bottom: 1260, left: 1440 }
      }
    },
    children: [
      ...titlePage(),
      ...sectionB(),
      ...sectionC1(),
      ...sectionC2(),
      ...sectionC3(),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/ConLaw_Notes_Executive_Judiciary.docx', buf);
  console.log('Done!');
}).catch(e => console.error(e));
