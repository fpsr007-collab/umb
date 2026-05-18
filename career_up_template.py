"""
キャリアアップ助成金（正社員化コース）要件チェック Excelテンプレート生成スクリプト
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation

# カラー定義
C_TITLE_BG    = "1F4E79"  # 濃紺（タイトル背景）
C_TITLE_FG    = "FFFFFF"  # 白（タイトル文字）
C_HEADER_BG   = "2E75B6"  # 中青（列ヘッダー背景）
C_SUBHEAD_BG  = "BDD7EE"  # 薄青（セクションヘッダー背景）
C_SUBHEAD_FG  = "1F4E79"  # 濃紺（セクションヘッダー文字）
C_INPUT_BG    = "FFF2CC"  # 薄黄（入力セル）
C_FORMULA_BG  = "F2F2F2"  # 薄灰（自動計算セル）
C_OK_BG       = "C6EFCE"  # 薄緑（OK）
C_OK_FG       = "276221"
C_NG_BG       = "FFC7CE"  # 薄赤（NG）
C_NG_FG       = "9C0006"
C_WARN_BG     = "FFEB9C"  # 薄黄（要確認）
C_WARN_FG     = "9C6500"


def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def cell_style(cell, value=None, bold=False, size=10, fg=None, bg=None,
               h_align="left", v_align="center", wrap=False, number_format=None, border=True):
    if value is not None:
        cell.value = value
    if fg:
        cell.font = Font(bold=bold, size=size, color=fg)
    else:
        cell.font = Font(bold=bold, size=size)
    if bg:
        cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    cell.alignment = Alignment(horizontal=h_align, vertical=v_align, wrap_text=wrap)
    if border:
        cell.border = thin_border()
    if number_format:
        cell.number_format = number_format


def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width


def merge_title(ws, row, col_start, col_end, value, size=12, bold=True, bg=C_TITLE_BG, fg=C_TITLE_FG):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start)
    cell_style(cell, value=value, bold=bold, size=size, fg=fg, bg=bg, h_align="center", border=False)


# ─────────────────────────────────────────────
# Sheet 1: 入力シート
# ─────────────────────────────────────────────
def build_input_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 18

    # タイトル
    merge_title(ws, 1, 1, 14, "キャリアアップ助成金（正社員化コース）要件チェックシート", size=14)
    merge_title(ws, 2, 1, 14, "※黄色セルに入力してください　　自動計算セルは変更しないでください", size=9,
                bg="DEEAF1", fg="1F4E79")

    # ── 企業情報セクション ──
    ws.row_dimensions[4].height = 20
    merge_title(ws, 4, 1, 14, "【企業情報】", size=11, bg=C_SUBHEAD_BG, fg=C_SUBHEAD_FG)

    labels_a = [
        (5, 1, "事業所名"),
        (5, 5, "企業規模"),
        (6, 1, "事業所番号（雇用保険）"),
        (6, 5, "キャリアアップ計画　提出日"),
        (7, 1, "所在地"),
        (7, 5, "キャリアアップ計画　受理日"),
        (8, 1, "担当社労士"),
        (8, 5, "就業規則　転換規定"),
    ]
    for r, c, txt in labels_a:
        ws.row_dimensions[r].height = 20
        cell_style(ws.cell(r, c), value=txt, bold=True, size=10, bg=C_SUBHEAD_BG, fg=C_SUBHEAD_FG,
                   h_align="right")
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 2)

    input_cells_a = [(5, 4), (6, 4), (7, 4), (8, 4)]
    for r, c in input_cells_a:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=4)
        cell_style(ws.cell(r, c), bg=C_INPUT_BG)

    right_input = [(5, 8), (6, 8), (7, 8), (8, 8)]
    for r, c in right_input:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=14)
        cell_style(ws.cell(r, c), bg=C_INPUT_BG)

    # 企業規模：データ検証
    dv_size = DataValidation(type="list", formula1='"中小企業,大企業"', allow_blank=True)
    ws.add_data_validation(dv_size)
    dv_size.add(ws.cell(5, 8))

    # 転換規定：データ検証
    dv_reg = DataValidation(type="list", formula1='"あり,なし,未確認"', allow_blank=True)
    ws.add_data_validation(dv_reg)
    dv_reg.add(ws.cell(8, 8))

    # 日付フォーマット
    ws.cell(6, 8).number_format = "YYYY/MM/DD"
    ws.cell(7, 8).number_format = "YYYY/MM/DD"

    # ── 対象者テーブル ──
    ws.row_dimensions[10].height = 20
    merge_title(ws, 10, 1, 14, "【対象者情報】", size=11, bg=C_SUBHEAD_BG, fg=C_SUBHEAD_FG)

    headers = [
        (1,  4,  "No."),
        (2,  16, "対象者氏名"),
        (3,  14, "雇用形態"),
        (4,  14, "雇用開始日"),
        (5,  14, "転換日"),
        (6,  10, "雇用期間\n(日数)"),
        (7,  14, "転換前賃金\n(月額)"),
        (8,  14, "転換後賃金\n(月額)"),
        (9,  10, "賃金UP率"),
        (10, 14, "会社都合\n離職者数\n(過去6ヶ月)"),
        (11, 14, "備考"),
    ]
    col = 1
    for idx, (_, w, h) in enumerate(headers, 1):
        ws.row_dimensions[11].height = 40
        cell_style(ws.cell(11, col), value=h, bold=True, size=9,
                   fg=C_TITLE_FG, bg=C_HEADER_BG, h_align="center", wrap=True)
        set_col_width(ws, col, w)
        col += 1

    # 対象者行（12〜31行、最大20人）
    dv_type = DataValidation(type="list", formula1='"有期契約,無期契約,派遣"', allow_blank=True)
    ws.add_data_validation(dv_type)

    for i in range(20):
        r = 12 + i
        ws.row_dimensions[r].height = 18
        # No.
        cell_style(ws.cell(r, 1), value=i + 1, h_align="center", bg="F8F8F8")
        # 氏名
        cell_style(ws.cell(r, 2), bg=C_INPUT_BG)
        # 雇用形態
        cell_style(ws.cell(r, 3), bg=C_INPUT_BG, h_align="center")
        dv_type.add(ws.cell(r, 3))
        # 雇用開始日
        c_start = ws.cell(r, 4)
        cell_style(c_start, bg=C_INPUT_BG, h_align="center", number_format="YYYY/MM/DD")
        # 転換日
        c_conv = ws.cell(r, 5)
        cell_style(c_conv, bg=C_INPUT_BG, h_align="center", number_format="YYYY/MM/DD")
        # 雇用期間（自動）
        c_days = ws.cell(r, 6)
        c_days.value = f'=IF(AND(D{r}<>"",E{r}<>""),E{r}-D{r},"")'
        cell_style(c_days, bg=C_FORMULA_BG, h_align="center", number_format="0")
        # 転換前賃金
        cell_style(ws.cell(r, 7), bg=C_INPUT_BG, h_align="right", number_format="#,##0")
        # 転換後賃金
        cell_style(ws.cell(r, 8), bg=C_INPUT_BG, h_align="right", number_format="#,##0")
        # 賃金UP率（自動）
        c_rate = ws.cell(r, 9)
        c_rate.value = f'=IF(AND(G{r}<>"",H{r}<>"",G{r}>0),(H{r}-G{r})/G{r},"")'
        cell_style(c_rate, bg=C_FORMULA_BG, h_align="center", number_format="0.0%")
        # 会社都合離職者数
        cell_style(ws.cell(r, 10), bg=C_INPUT_BG, h_align="center", number_format="0")
        # 備考
        cell_style(ws.cell(r, 11), bg=C_INPUT_BG)

    # 列幅調整（累計で設定済み、ここでは追加調整）
    set_col_width(ws, 1, 4)
    set_col_width(ws, 2, 16)
    set_col_width(ws, 3, 14)
    set_col_width(ws, 4, 14)
    set_col_width(ws, 5, 14)
    set_col_width(ws, 6, 10)
    set_col_width(ws, 7, 14)
    set_col_width(ws, 8, 14)
    set_col_width(ws, 9, 10)
    set_col_width(ws, 10, 12)
    set_col_width(ws, 11, 20)

    # 条件付き書式：賃金UP率
    ok_fill   = PatternFill(start_color=C_OK_BG, end_color=C_OK_BG, fill_type="solid")
    ng_fill   = PatternFill(start_color=C_NG_BG, end_color=C_NG_BG, fill_type="solid")
    ws.conditional_formatting.add(
        "I12:I31",
        CellIsRule(operator="greaterThanOrEqual", formula=["0.03"], fill=ok_fill)
    )
    ws.conditional_formatting.add(
        "I12:I31",
        CellIsRule(operator="lessThan", formula=["0.03"], fill=ng_fill)
    )

    # 条件付き書式：雇用期間
    ws.conditional_formatting.add(
        "F12:F31",
        CellIsRule(operator="greaterThanOrEqual", formula=["180"], fill=ok_fill)
    )
    ws.conditional_formatting.add(
        "F12:F31",
        CellIsRule(operator="lessThan", formula=["180"], fill=ng_fill)
    )


# ─────────────────────────────────────────────
# Sheet 2: 判定結果
# ─────────────────────────────────────────────
def build_result_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 18

    merge_title(ws, 1, 1, 12, "キャリアアップ助成金（正社員化コース）判定結果", size=14)
    merge_title(ws, 2, 1, 12, "※入力シートの値を自動で判定します　変更しないでください", size=9,
                bg="DEEAF1", fg="1F4E79")

    ws.row_dimensions[4].height = 45
    headers = [
        "No.", "対象者氏名", "雇用期間\n(6ヶ月以上)",
        "賃金UP\n(3%以上)", "就業規則\n転換規定",
        "会社都合\n離職者なし", "計画\n受理済み",
        "総合判定", "助成見込額\n(中小企業)", "申請期限\n(開始)", "申請期限\n(終了)", "備考"
    ]
    widths = [4, 16, 12, 10, 12, 12, 10, 10, 16, 14, 14, 20]
    for idx, (h, w) in enumerate(zip(headers, widths), 1):
        cell_style(ws.cell(4, idx), value=h, bold=True, size=9,
                   fg=C_TITLE_FG, bg=C_HEADER_BG, h_align="center", wrap=True)
        set_col_width(ws, idx, w)

    for i in range(20):
        r = 5 + i
        ws.row_dimensions[r].height = 18
        src = 12 + i  # 入力シートの行

        # No.
        cell_style(ws.cell(r, 1), value=f"=入力シート!A{src}", h_align="center", bg="F8F8F8")
        # 氏名
        cell_style(ws.cell(r, 2), value=f"=入力シート!B{src}", bg="F8F8F8")

        # 雇用期間チェック
        f_days = f'=IF(入力シート!D{src}="","－",IF(入力シート!F{src}>=180,"✅ OK","❌ NG"))'
        cell_style(ws.cell(r, 3), value=f_days, h_align="center", bg=C_FORMULA_BG)

        # 賃金UPチェック
        f_rate = f'=IF(入力シート!G{src}="","－",IF(入力シート!I{src}>=0.03,"✅ OK","❌ NG"))'
        cell_style(ws.cell(r, 4), value=f_rate, h_align="center", bg=C_FORMULA_BG)

        # 就業規則転換規定
        f_reg = f'=IF(入力シート!H8="","⚠️ 未確認",IF(入力シート!H8="あり","✅ OK","❌ NG"))'
        cell_style(ws.cell(r, 5), value=f_reg, h_align="center", bg=C_FORMULA_BG)

        # 会社都合離職者
        f_离 = f'=IF(入力シート!J{src}="","⚠️ 未確認",IF(入力シート!J{src}=0,"✅ OK","❌ NG"))'
        cell_style(ws.cell(r, 6), value=f_离, h_align="center", bg=C_FORMULA_BG)

        # キャリアアップ計画受理
        f_plan = '=IF(入力シート!H7="","⚠️ 未確認","✅ OK")'
        cell_style(ws.cell(r, 7), value=f_plan, h_align="center", bg=C_FORMULA_BG)

        # 総合判定
        f_total = (
            f'=IF(入力シート!B{src}="","",IF(AND('
            f'C{r}="✅ OK",'
            f'D{r}="✅ OK",'
            f'E{r}="✅ OK",'
            f'F{r}="✅ OK",'
            f'G{r}="✅ OK"'
            f'),"✅ 申請可","⚠️ 要確認"))'
        )
        cell_style(ws.cell(r, 8), value=f_total, h_align="center", bold=True, bg=C_FORMULA_BG)

        # 助成見込額
        f_amt = (
            f'=IF(入力シート!B{src}="","",IF(入力シート!H5="大企業",'
            f'IF(入力シート!C{src}="派遣",570000,IF(入力シート!C{src}="有期契約",425000,215000)),'
            f'IF(入力シート!C{src}="派遣",720000,IF(入力シート!C{src}="有期契約",570000,285000))))'
        )
        cell_style(ws.cell(r, 9), value=f_amt, h_align="right", bg=C_FORMULA_BG,
                   number_format="#,##0円")

        # 申請期限（転換日 + 6ヶ月）
        f_start = f'=IF(入力シート!E{src}="","",EDATE(入力シート!E{src},6))'
        cell_style(ws.cell(r, 10), value=f_start, h_align="center", bg=C_FORMULA_BG,
                   number_format="YYYY/MM/DD")

        # 申請期限（転換日 + 8ヶ月）
        f_end = f'=IF(入力シート!E{src}="","",EDATE(入力シート!E{src},8))'
        cell_style(ws.cell(r, 11), value=f_end, h_align="center", bg=C_FORMULA_BG,
                   number_format="YYYY/MM/DD")

        # 備考
        cell_style(ws.cell(r, 12), bg="FFFFFF")

    # 条件付き書式：総合判定
    ok_fill   = PatternFill(start_color=C_OK_BG,   end_color=C_OK_BG,   fill_type="solid")
    warn_fill = PatternFill(start_color=C_WARN_BG,  end_color=C_WARN_BG, fill_type="solid")
    ws.conditional_formatting.add("H5:H24", FormulaRule(
        formula=['H5="✅ 申請可"'], fill=ok_fill))
    ws.conditional_formatting.add("H5:H24", FormulaRule(
        formula=['H5="⚠️ 要確認"'], fill=warn_fill))


# ─────────────────────────────────────────────
# Sheet 3: 要件一覧（参照用）
# ─────────────────────────────────────────────
def build_ref_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    merge_title(ws, 1, 1, 4, "キャリアアップ助成金（正社員化コース）要件一覧【参照用】", size=13)

    sections = [
        ("【事前要件】転換前に確認", [
            ("キャリアアップ計画", "転換日の前日までに労働局へ提出・受理が必要"),
            ("就業規則の転換規定", "正社員転換を規定した条文が必要（転換日前に届出済みであること）"),
            ("雇用形態", "有期契約・無期契約・派遣社員が対象"),
            ("雇用期間", "転換前に同一事業所で6ヶ月以上継続雇用されていること"),
            ("会社都合離職者", "転換日の前日から過去6ヶ月以内に特定受給資格者となる離職者がいないこと"),
            ("雇用保険", "雇用保険適用事業所であること"),
        ]),
        ("【転換時の要件】", [
            ("転換後の雇用形態", "無期・フルタイムの正規雇用労働者であること"),
            ("賃金の増額", "転換前と比較して3%以上の賃金増額が必要"),
            ("転換根拠", "就業規則に基づく転換であること"),
        ]),
        ("【申請時の要件】", [
            ("支給申請期限", "転換日から6ヶ月後の翌日〜8ヶ月後の末日（2ヶ月間）"),
            ("賃金支払い", "転換後6ヶ月分の賃金を支払い済みであること"),
            ("添付書類", "雇用契約書、賃金台帳、転換前後の出勤簿、就業規則など"),
        ]),
        ("【助成額（2024年度）】", [
            ("有期→正規（中小企業）", "57万円／人（年間上限20人）"),
            ("有期→正規（大企業）", "42.5万円／人"),
            ("無期→正規（中小企業）", "28.5万円／人"),
            ("無期→正規（大企業）", "21.5万円／人"),
            ("派遣→正規（中小企業）", "72万円／人"),
            ("派遣→正規（大企業）", "57万円／人"),
        ]),
    ]

    row = 3
    set_col_width(ws, 1, 4)
    set_col_width(ws, 2, 30)
    set_col_width(ws, 3, 50)

    for section_title, items in sections:
        ws.row_dimensions[row].height = 22
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        cell_style(ws.cell(row, 1), value=section_title, bold=True, size=11,
                   fg=C_SUBHEAD_FG, bg=C_SUBHEAD_BG, border=False)
        row += 1

        for item, desc in items:
            ws.row_dimensions[row].height = 18
            cell_style(ws.cell(row, 1), value="", bg="F8F8F8")
            cell_style(ws.cell(row, 2), value=item, bold=True, size=10)
            cell_style(ws.cell(row, 3), value=desc, size=10, wrap=True)
            row += 1

        row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    cell_style(ws.cell(row, 1),
               value="※ 要件・助成額は法改正により変更となる場合があります。必ず最新の厚労省資料をご確認ください。",
               size=9, fg="FF0000", border=False)


# ─────────────────────────────────────────────
# メイン
# ─────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()

    ws_input  = wb.active
    ws_input.title = "入力シート"
    ws_result = wb.create_sheet("判定結果")
    ws_ref    = wb.create_sheet("要件一覧")

    build_input_sheet(ws_input)
    build_result_sheet(ws_result)
    build_ref_sheet(ws_ref)

    # 印刷設定
    for ws in [ws_input, ws_result]:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1

    path = "/home/user/umb/キャリアアップ助成金_要件チェックシート.xlsx"
    wb.save(path)
    print(f"✅ 作成完了: {path}")


if __name__ == "__main__":
    main()
