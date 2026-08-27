const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  Header, convertInchesToTwip,
} = require("docx");

const FONT = "游ゴシック";

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: "000000" };

function cellBox(children, opts = {}) {
  return new TableCell({
    children,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shade ? { type: ShadingType.CLEAR, color: "auto", fill: "F2F2F2" } : undefined,
    borders: opts.borders,
    verticalAlign: "center",
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
  });
}

function labelRun(text) {
  return new TextRun({ text, font: FONT, size: 21, bold: true });
}

function replyRow(label, width1, width2) {
  return new TableRow({
    children: [
      cellBox([new Paragraph({ children: [labelRun(label)] })], {
        width: width1,
        shade: true,
        borders: { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder },
      }),
      cellBox([new Paragraph({ children: [new TextRun({ text: "", font: FONT })] })], {
        width: width2,
        borders: { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder },
      }),
    ],
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { after: 40 },
    children: [
      new TextRun({ text: "・", font: FONT, size: 20 }),
      new TextRun({ text, font: FONT, size: 20 }),
    ],
  });
}

const doc = new Document({
  sections: [
    {
      properties: {
        page: {
          margin: {
            top: convertInchesToTwip(0.5),
            bottom: convertInchesToTwip(0.45),
            left: convertInchesToTwip(0.7),
            right: convertInchesToTwip(0.7),
          },
        },
      },
      children: [
        // ---- 発信者レターヘッド ----
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 0 },
          children: [new TextRun({ text: "社会保険労務士法人　アンブレラ", font: FONT, size: 20, bold: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 0 },
          children: [new TextRun({ text: "代表社員　伊藤　泰人", font: FONT, size: 20, bold: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 0 },
          children: [new TextRun({ text: "立川市曙町２－３４－１３　オリンピック第３ビル６階", font: FONT, size: 18 })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 120 },
          children: [new TextRun({ text: "TEL：042-595-6103　　FAX：042-595-6132", font: FONT, size: 18 })],
        }),
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" } },
          spacing: { after: 160 },
          children: [new TextRun({ text: "", font: FONT })],
        }),

        // ---- 宛先 ----
        new Paragraph({
          spacing: { after: 240 },
          children: [new TextRun({ text: "税理士事務所　御中", font: FONT, size: 24, bold: true })],
        }),

        // ---- タイトル ----
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [new TextRun({ text: "顧問先様へのご案内に、ぜひご活用ください", font: FONT, size: 20 })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [
            new TextRun({ text: "職員を１人でも雇用していれば、", font: FONT, size: 30, bold: true }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "厚生労働省の助成金を申請できる可能性があります", font: FONT, size: 30, bold: true }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 220 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" } },
          children: [
            new TextRun({ text: "～ 顧問先企業様の雇用関係助成金、無料診断いたします ～", font: FONT, size: 20 }),
          ],
        }),

        // ---- 本文 ----
        new Paragraph({
          spacing: { after: 140 },
          children: [
            new TextRun({
              text: "税理士事務所の皆様におかれましては、日頃より格別のお引き立てを賜り、誠にありがとうございます。突然のFAXにて失礼いたします。",
              font: FONT, size: 20,
            }),
          ],
        }),
        new Paragraph({
          spacing: { after: 140 },
          children: [
            new TextRun({
              text: "厚生労働省が実施する雇用関係助成金は、従業員を１名でも雇用している事業所であれば、業種を問わず申請できる可能性がある制度です。しかし要件が複雑なため、多くの事業所様が「知らないまま」申請の機会を逃しているのが実情です。",
              font: FONT, size: 20,
            }),
          ],
        }),
        new Paragraph({
          spacing: { after: 160 },
          children: [
            new TextRun({
              text: "私ども社会保険労務士法人アンブレラでは、貴事務所の顧問先企業様に対する助成金診断・申請代行を通じて、税理士事務所様の付加価値向上のお手伝いをさせていただいております。",
              font: FONT, size: 20,
            }),
          ],
        }),

        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: "【こんな顧問先様はいらっしゃいませんか？】", font: FONT, size: 21, bold: true })],
        }),
        bullet("従業員を新たに採用した、または採用を予定している"),
        bullet("高齢者・障害者・母子家庭の方などを雇用している"),
        bullet("従業員のキャリアアップ・処遇改善に取り組んでいる"),
        bullet("働き方改革・両立支援（育児・介護）に取り組んでいる"),
        bullet("就業規則の整備・見直しを検討している"),
        new Paragraph({
          spacing: { before: 80, after: 160 },
          children: [
            new TextRun({ text: "上記に１つでも当てはまる場合、活用できる助成金がある可能性がございます。", font: FONT, size: 20 }),
          ],
        }),

        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: "【アンブレラにご依頼いただくメリット】", font: FONT, size: 21, bold: true })],
        }),
        bullet("助成金診断は無料で承ります"),
        bullet("完全成功報酬制のため、着手金は不要です"),
        bullet("貴事務所の顧問料収入に影響を与えません（業務範囲を明確に分離）"),
        bullet("煩雑な申請書類の作成・労働局対応はすべて代行いたします"),

        new Paragraph({
          spacing: { before: 160, after: 260 },
          children: [
            new TextRun({
              text: "まずは資料送付、またはお打ち合わせのお時間を頂戴できましたら幸いです。下記の返信欄にご記入のうえ、本FAX番号宛にご返信いただくか、お電話・メールにてお気軽にお問い合わせください。今後とも何卒よろしくお願い申し上げます。",
              font: FONT, size: 20,
            }),
          ],
        }),

        // ---- 返信欄見出し ----
        new Paragraph({
          shading: { type: ShadingType.CLEAR, color: "auto", fill: "1F1F1F" },
          spacing: { after: 100 },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "☎　資料請求・お問い合わせ　返信用紙　☎", font: FONT, size: 22, bold: true, color: "FFFFFF" }),
          ],
        }),
        new Paragraph({
          spacing: { after: 140 },
          children: [
            new TextRun({
              text: "下記にご記入のうえ、本用紙をそのままFAX：042-595-6132　までご返信ください。",
              font: FONT, size: 19, bold: true,
            }),
          ],
        }),

        new Table({
          width: { size: 10100, type: WidthType.DXA },
          rows: [
            replyRow("税理士事務所名", 2600, 7500),
            replyRow("ご担当者名", 2600, 7500),
            replyRow("電話番号", 2600, 7500),
            replyRow("FAX番号", 2600, 7500),
            replyRow("メールアドレス", 2600, 7500),
          ],
        }),

        new Paragraph({
          spacing: { before: 220 },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "お問い合わせ先：社会保険労務士法人アンブレラ　（担当：伊藤）",
              font: FONT, size: 18,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "立川市曙町２－３４－１３　オリンピック第３ビル６階　　TEL：042-595-6103　　FAX：042-595-6132",
              font: FONT, size: 18,
            }),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync("/home/user/umb/faxdm/税理士事務所向けFAXDM.docx", buf);
  console.log("done");
});
