import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "就業規則未整備リスク一覧（統合版）"

# ── データ定義 ────────────────────────────────────────────────
# (No, 項目名, 統合元, 説明, 想定リスク, カテゴリ, is_merged)
data = [
    (
        1,
        "懲戒処分の根拠が弱く・\nハラスメント対応も困難になる",
        "旧#1＋#10を統合",
        "懲戒事由・処分の種類・手続きが明確でないため、無断欠勤・業務命令違反・情報持ち出し・パワハラ等への対処根拠が弱くなります。フジ興産事件（最高裁2003年）でも就業規則への明記が有効要件とされています。また2022年4月からすべての事業主にパワハラ防止措置が義務付けられており（労働施策総合推進法）、対応が不十分な場合は会社が使用者責任を問われます。",
        "問題社員対応が後手に回る・懲戒無効・バックペイ・使用者責任・損害賠償請求・被害者離職リスク",
        "法的リスク", True
    ),
    (
        2,
        "解雇・本採用拒否・能力不足社員への\n対応力が弱くなる",
        "旧#2＋#11を統合",
        "能力不足・勤務態度不良・協調性欠如・試用期間中の適格性判断について、解雇理由・手続き・予告期間のルールが書面で存在しないため「不当解雇」を主張された際に会社側の根拠が著しく弱くなります。特に能力不足解雇は裁判所が厳格に判断しており、指導・教育・配置転換を行った証拠がなければ解雇無効とされるケースが多くあります。",
        "解雇無効・金銭解決・不当解雇主張・労働審判・問題社員の長期在籍・職場モラル低下リスク",
        "法的リスク", True
    ),
    (
        3,
        "助成金・制度導入で不利になる",
        "旧#3（単独）",
        "厚生労働省の多くの雇用関係助成金では就業規則の整備・届出が申請要件に含まれており、整備していないだけで申請資格を失います。規程整備・周知・制度運用が求められる助成金では申請準備も遅れます。",
        "助成金受給機会の逸失・申請準備遅延リスク",
        "経営リスク", False
    ),
    (
        4,
        "休職・長期欠勤の処理と\n出口戦略が曖昧になる",
        "旧#4＋#13を統合",
        "私傷病で働けなくなった場合の休職発令基準・期間上限・診断書提出義務・復職判断手順・期間満了時の退職扱いが定まらず、対応が「待ちの状態」のまま際限なく長期化します。解雇に踏み切っても就業規則に根拠規定がなければ無効とされるリスクがあります。",
        "長期欠勤・復職トラブル・給与負担の長期化・業務穴埋めコスト・解雇無効・自然退職扱い無効リスク",
        "労務管理リスク", True
    ),
    (
        5,
        "労働条件の変更が難しくなる",
        "旧#5（単独）",
        "賃金・手当・休日・勤務時間・賞与・退職金などの変更時に、基準や手続きが曖昧になり、不利益変更トラブルにつながりやすくなります。",
        "不利益変更トラブル",
        "法的リスク", False
    ),
    (
        6,
        "「言った・言わない」が増える",
        "旧#6（単独）",
        "欠勤連絡・有給申請・退職届・貸与物返還・服務規律などが口頭ルールになり、従業員から「聞いていない」と言われやすくなります。",
        "社長判断への不満・不公平感",
        "労務管理リスク", False
    ),
    (
        7,
        "残業代の設計不備・\n未払い遡及請求リスクが高まる",
        "旧#7＋#12を統合",
        "固定残業代を給与に含める場合、計算根拠・対象時間数・超過時の追加払いルールの明記が有効要件です。また労働時間の管理方法・認定基準・残業申請ルールが定められていないと実態と賃金台帳が乖離しやすくなります。2020年の労基法改正により未払い賃金の請求権消滅時効が3年に延長されており、退職者・在職者を問わず一括請求されるリスクが増大しています。",
        "固定残業代無効・3年分の未払い残業代・遅延損害金・複数社員による集団請求リスク",
        "法的リスク", True
    ),
    (
        8,
        "変形労働時間制・\nフレックスタイム制を導入できない",
        "旧#8（単独）",
        "シフト制や繁閑対応の柔軟な勤務体系を導入するには、就業規則への記載が労働基準法上の要件となっています（第32条の2・32条の4等）。",
        "制度無効・残業代発生・人件費増加リスク",
        "経営リスク", False
    ),
    (
        9,
        "採用力・人材定着率の低下",
        "旧#9（単独）",
        "転職経験のある求職者は就業規則の開示を求めることがあります。開示できない場合、職場環境が不透明と判断され、優秀な人材の採用に影響します。",
        "採用競争力低下・早期離職リスク",
        "経営リスク", False
    ),
]

# ── スタイル定義 ──────────────────────────────────────────────
def side(style="thin", color="BBBBBB"):
    return Side(style=style, color=color)

BORDER = Border(
    left=side(), right=side(), top=side(), bottom=side()
)
BORDER_MERGED = Border(
    left=side("medium", "2E75B6"),
    right=side("medium", "2E75B6"),
    top=side("medium", "2E75B6"),
    bottom=side("medium", "2E75B6"),
)

# 塗りつぶし
FILL = {
    "header":       PatternFill("solid", fgColor="1F3864"),
    "title_row":    PatternFill("solid", fgColor="2E4057"),
    "法的リスク":   PatternFill("solid", fgColor="FFE0E0"),
    "経営リスク":   PatternFill("solid", fgColor="FFF2CC"),
    "労務管理リスク": PatternFill("solid", fgColor="EAF0FF"),
    "法的リスク_cat":    PatternFill("solid", fgColor="C00000"),
    "経営リスク_cat":    PatternFill("solid", fgColor="ED7D31"),
    "労務管理リスク_cat": PatternFill("solid", fgColor="2E75B6"),
    "merged_mark":  PatternFill("solid", fgColor="D9E1F2"),
    "legend_法的":  PatternFill("solid", fgColor="FFE0E0"),
    "legend_経営":  PatternFill("solid", fgColor="FFF2CC"),
    "legend_労務":  PatternFill("solid", fgColor="EAF0FF"),
    "legend_統合":  PatternFill("solid", fgColor="D9E1F2"),
}

# フォント
FONT = {
    "header":   Font(name="Meiryo UI", bold=True, color="FFFFFF", size=10),
    "title":    Font(name="Meiryo UI", bold=True, color="FFFFFF", size=14),
    "no":       Font(name="Meiryo UI", bold=True, size=12),
    "item":     Font(name="Meiryo UI", bold=True, size=10, color="1F3864"),
    "item_m":   Font(name="Meiryo UI", bold=True, size=10, color="C00000"),
    "src":      Font(name="Meiryo UI", italic=True, size=9, color="666666"),
    "body":     Font(name="Meiryo UI", size=10, color="404040"),
    "risk":     Font(name="Meiryo UI", bold=True, size=10, color="C00000"),
    "cat":      Font(name="Meiryo UI", bold=True, size=9, color="FFFFFF"),
    "legend":   Font(name="Meiryo UI", bold=True, size=9, color="404040"),
    "note":     Font(name="Meiryo UI", size=9, color="666666"),
}

WRAP   = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")
VCENTER = Alignment(wrap_text=True, vertical="center", horizontal="left")

# ── 列幅設定 ──────────────────────────────────────────────────
col_widths = {
    1: 5,    # No
    2: 11,   # カテゴリ
    3: 28,   # 項目名
    4: 14,   # 統合元
    5: 46,   # 説明
    6: 36,   # 想定リスク
}
for ci, w in col_widths.items():
    ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

# ── タイトル行 ────────────────────────────────────────────────
ws.merge_cells("A1:F1")
c = ws["A1"]
c.value     = "就業規則未整備リスク一覧　（統合版：13項目 → 9項目）"
c.font      = FONT["title"]
c.fill      = FILL["header"]
c.alignment = CENTER
ws.row_dimensions[1].height = 36

# ── 凡例行 ────────────────────────────────────────────────────
ws.row_dimensions[2].height = 20
legends = [
    ("A2", "B2", "【カテゴリ】",   None,             "legend"),
    ("C2", "C2", "法的リスク",     "legend_法的",    "legend"),
    ("D2", "D2", "経営リスク",     "legend_経営",    "legend"),
    ("E2", "E2", "労務管理リスク", "legend_労務",    "legend"),
    ("F2", "F2", "■統合項目",     "legend_統合",    "legend"),
]
for start, end, label, fill_key, font_key in legends:
    if start != end:
        ws.merge_cells(f"{start}:{end}")
    c = ws[start]
    c.value     = label
    c.font      = FONT[font_key]
    c.fill      = FILL[fill_key] if fill_key else PatternFill("solid", fgColor="F2F2F2")
    c.alignment = CENTER
    c.border    = BORDER

# ── ヘッダー行 ────────────────────────────────────────────────
ws.row_dimensions[3].height = 24
headers = ["No", "カテゴリ", "項　目　名", "統合元", "説　　　明", "想定リスク"]
for ci, h in enumerate(headers, 1):
    c = ws.cell(row=3, column=ci, value=h)
    c.font      = FONT["header"]
    c.fill      = FILL["header"]
    c.alignment = CENTER
    c.border    = BORDER

# ── データ行 ──────────────────────────────────────────────────
CAT_FILL_KEY = {
    "法的リスク":    ("法的リスク",    "法的リスク_cat"),
    "経営リスク":    ("経営リスク",    "経営リスク_cat"),
    "労務管理リスク": ("労務管理リスク", "労務管理リスク_cat"),
}

for idx, (no, title, src, desc, risk, cat, is_merged) in enumerate(data):
    row = idx + 4
    base_fill_key, cat_fill_key = CAT_FILL_KEY[cat]
    base_fill = FILL[base_fill_key]
    border    = BORDER_MERGED if is_merged else BORDER

    # 行高さを説明文の長さに応じて設定
    lines = max(len(desc) // 30 + 2, 3)
    ws.row_dimensions[row].height = max(lines * 15, 60)

    # A: No
    c = ws.cell(row=row, column=1, value=no)
    c.font      = Font(name="Meiryo UI", bold=True, size=13,
                       color="C00000" if is_merged else "1F3864")
    c.fill      = FILL["merged_mark"] if is_merged else base_fill
    c.alignment = CENTER
    c.border    = border

    # B: カテゴリ（縦書き）
    c = ws.cell(row=row, column=2, value=cat)
    c.font      = FONT["cat"]
    c.fill      = FILL[cat_fill_key]
    c.alignment = Alignment(wrap_text=True, vertical="center",
                            horizontal="center", text_rotation=90)
    c.border    = border

    # C: 項目名
    c = ws.cell(row=row, column=3, value=title)
    c.font      = FONT["item_m"] if is_merged else FONT["item"]
    c.fill      = FILL["merged_mark"] if is_merged else base_fill
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
    c.border    = border

    # D: 統合元
    c = ws.cell(row=row, column=4, value=src)
    c.font      = Font(name="Meiryo UI", italic=True, bold=is_merged, size=9,
                       color="C00000" if is_merged else "888888")
    c.fill      = FILL["merged_mark"] if is_merged else base_fill
    c.alignment = CENTER
    c.border    = border

    # E: 説明
    c = ws.cell(row=row, column=5, value=desc)
    c.font      = FONT["body"]
    c.fill      = base_fill
    c.alignment = WRAP
    c.border    = border

    # F: 想定リスク
    c = ws.cell(row=row, column=6, value=risk)
    c.font      = FONT["risk"]
    c.fill      = base_fill
    c.alignment = WRAP
    c.border    = border

# ── 統合サマリー行 ────────────────────────────────────────────
summary_row = len(data) + 4
ws.row_dimensions[summary_row].height = 20
ws.merge_cells(f"A{summary_row}:F{summary_row}")
c = ws[f"A{summary_row}"]
c.value = "※ 赤枠・赤字の項目は複数項目を統合したもの（4組統合）。統合前13項目 → 統合後9項目。"
c.font  = Font(name="Meiryo UI", italic=True, size=9, color="C00000")
c.fill  = PatternFill("solid", fgColor="FFF0F0")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
c.border = BORDER

# ── ウィンドウ枠固定・フィルター ─────────────────────────────
ws.freeze_panes = "A4"
ws.auto_filter.ref = "A3:F3"

# ── 保存 ─────────────────────────────────────────────────────
output = "/home/user/umb/就業規則未整備リスク一覧_統合版.xlsx"
wb.save(output)
print(f"Saved: {output}")
