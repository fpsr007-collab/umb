"""
キャリアアップ助成金（正社員化コース）要件チェック Excelテンプレート生成スクリプト
令和8年度（2026年度）対応版
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

# カラー定義
C_TITLE_BG   = "1F4E79"
C_TITLE_FG   = "FFFFFF"
C_HEADER_BG  = "2E75B6"
C_SUBHEAD_BG = "BDD7EE"
C_SUBHEAD_FG = "1F4E79"
C_INPUT_BG   = "FFF2CC"
C_FORMULA_BG = "F2F2F2"
C_OK_BG      = "C6EFCE"
C_NG_BG      = "FFC7CE"
C_WARN_BG    = "FFEB9C"
C_NEW_BG     = "E2EFDA"  # 薄緑（新設項目）


def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def cell_style(cell, value=None, bold=False, size=10, fg=None, bg=None,
               h_align="left", v_align="center", wrap=False,
               number_format=None, border=True):
    if value is not None:
        cell.value = value
    cell.font = Font(bold=bold, size=size, color=fg if fg else "000000")
    if bg:
        cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    cell.alignment = Alignment(horizontal=h_align, vertical=v_align, wrap_text=wrap)
    if border:
        cell.border = thin_border()
    if number_format:
        cell.number_format = number_format


def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width


def merge_title(ws, row, col_start, col_end, value, size=12, bold=True,
                bg=C_TITLE_BG, fg=C_TITLE_FG):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start)
    cell_style(cell, value=value, bold=bold, size=size, fg=fg, bg=bg,
               h_align="center", border=False)


# ─────────────────────────────────────────────
# Sheet 1: 入力シート
# 企業情報: 行5〜9
# 対象者ヘッダー: 行11
# 対象者データ: 行12〜31（最大20人）
#
# 列構成（対象者テーブル）:
#  A=No. B=氏名 C=雇用形態 D=重点支援対象者 E=雇用開始日 F=転換日
#  G=雇用期間(自動) H=転換前賃金 I=転換後賃金 J=賃金UP率(自動)
#  K=会社都合離職者数 L=備考
# ─────────────────────────────────────────────
def build_input_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8

    merge_title(ws, 1, 1, 12,
                "キャリアアップ助成金（正社員化コース）要件チェックシート　令和8年度（2026年度）版",
                size=13)
    merge_title(ws, 2, 1, 12,
                "※黄色セルに入力してください　　灰色セルは自動計算のため変更不要",
                size=9, bg="DEEAF1", fg="1F4E79")

    # ── 企業情報セクション ──
    ws.row_dimensions[4].height = 20
    merge_title(ws, 4, 1, 12, "【企業情報】", size=11, bg=C_SUBHEAD_BG, fg=C_SUBHEAD_FG)

    # ラベル列 (col 1-3: 左側, col 5-7: 右側)
    left_labels = [
        (5, "事業所名"),
        (6, "事業所番号（雇用保険）"),
        (7, "所在地"),
        (8, "担当社労士"),
        (9, ""),
    ]
    right_labels = [
        (5, "企業規模"),
        (6, "キャリアアップ計画　提出日"),
        (7, "キャリアアップ計画　受理日"),
        (8, "就業規則　転換規定"),
        (9, "情報公表（令和8年度新設）"),  # 新設
    ]

    for r, txt in left_labels:
        ws.row_dimensions[r].height = 20
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        cell_style(ws.cell(r, 1), value=txt, bold=True, size=10,
                   bg=C_SUBHEAD_BG, fg=C_SUBHEAD_FG, h_align="right")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=4)
        cell_style(ws.cell(r, 4), bg=C_INPUT_BG)

    for r, txt in right_labels:
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        bg = C_NEW_BG if r == 9 else C_SUBHEAD_BG
        cell_style(ws.cell(r, 5), value=txt, bold=True, size=10,
                   bg=bg, fg=C_SUBHEAD_FG, h_align="right")
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=12)
        cell_style(ws.cell(r, 8), bg=C_INPUT_BG if r != 9 else C_NEW_BG)

    # データ検証：企業規模
    dv_size = DataValidation(type="list", formula1='"中小企業,大企業"', allow_blank=True)
    ws.add_data_validation(dv_size)
    dv_size.add(ws.cell(5, 8))

    # データ検証：就業規則転換規定
    dv_reg = DataValidation(type="list", formula1='"あり,なし,未確認"', allow_blank=True)
    ws.add_data_validation(dv_reg)
    dv_reg.add(ws.cell(8, 8))

    # データ検証：情報公表
    dv_pub = DataValidation(type="list", formula1='"あり,なし"', allow_blank=True)
    ws.add_data_validation(dv_pub)
    dv_pub.add(ws.cell(9, 8))

    # 日付フォーマット
    ws.cell(6, 8).number_format = "YYYY/MM/DD"
    ws.cell(7, 8).number_format = "YYYY/MM/DD"

    # ── 対象者テーブル ──
    ws.row_dimensions[10].height = 8
    ws.row_dimensions[11].height = 20
    merge_title(ws, 11, 1, 12, "【対象者情報】", size=11, bg=C_SUBHEAD_BG, fg=C_SUBHEAD_FG)

    ws.row_dimensions[12].height = 45
    col_headers = [
        (1,  4,  "No."),
        (2,  16, "対象者氏名"),
        (3,  14, "雇用形態"),
        (4,  14, "重点支援\n対象者\n※"),
        (5,  13, "雇用開始日"),
        (6,  13, "転換日"),
        (7,  10, "雇用期間\n(日数)\n自動"),
        (8,  13, "転換前賃金\n(月額)"),
        (9,  13, "転換後賃金\n(月額)"),
        (10, 10, "賃金UP率\n自動"),
        (11, 12, "会社都合\n離職者数\n(過去6ヶ月)"),
        (12, 18, "備考"),
    ]
    for col, width, header in col_headers:
        cell_style(ws.cell(12, col), value=header, bold=True, size=9,
                   fg=C_TITLE_FG, bg=C_HEADER_BG, h_align="center", wrap=True)
        set_col_width(ws, col, width)

    # 注釈行
    ws.row_dimensions[13].height = 14
    ws.merge_cells(start_row=13, start_column=1, end_row=13, end_column=12)
    cell_style(ws.cell(13, 1),
               value="※ 重点支援対象者: 派遣労働者・雇入れ3年以上の有期労働者・過去5年間で正規雇用1年以下かつ過去1年間正規未就労の有期労働者 など　派遣は自動的に「はい」",
               size=8, fg="9C0006", border=False)

    # データ検証：雇用形態
    dv_type = DataValidation(type="list", formula1='"有期契約,無期契約,派遣"', allow_blank=True)
    ws.add_data_validation(dv_type)

    # データ検証：重点支援対象者
    dv_key = DataValidation(type="list", formula1='"はい,いいえ"', allow_blank=True)
    ws.add_data_validation(dv_key)

    # 対象者データ行（14〜33）
    for i in range(20):
        r = 14 + i
        ws.row_dimensions[r].height = 18

        cell_style(ws.cell(r, 1), value=i + 1, h_align="center", bg="F8F8F8")
        cell_style(ws.cell(r, 2), bg=C_INPUT_BG)

        # 雇用形態
        cell_style(ws.cell(r, 3), bg=C_INPUT_BG, h_align="center")
        dv_type.add(ws.cell(r, 3))

        # 重点支援対象者（派遣なら自動「はい」）
        ws.cell(r, 4).value = f'=IF(C{r}="派遣","はい","")'
        cell_style(ws.cell(r, 4), bg=C_INPUT_BG, h_align="center")
        dv_key.add(ws.cell(r, 4))

        # 雇用開始日・転換日
        cell_style(ws.cell(r, 5), bg=C_INPUT_BG, h_align="center",
                   number_format="YYYY/MM/DD")
        cell_style(ws.cell(r, 6), bg=C_INPUT_BG, h_align="center",
                   number_format="YYYY/MM/DD")

        # 雇用期間（自動）
        ws.cell(r, 7).value = f'=IF(AND(E{r}<>"",F{r}<>""),F{r}-E{r},"")'
        cell_style(ws.cell(r, 7), bg=C_FORMULA_BG, h_align="center",
                   number_format="0")

        # 賃金
        cell_style(ws.cell(r, 8), bg=C_INPUT_BG, h_align="right",
                   number_format="#,##0")
        cell_style(ws.cell(r, 9), bg=C_INPUT_BG, h_align="right",
                   number_format="#,##0")

        # 賃金UP率（自動）
        ws.cell(r, 10).value = (
            f'=IF(AND(H{r}<>"",I{r}<>"",H{r}>0),(I{r}-H{r})/H{r},"")'
        )
        cell_style(ws.cell(r, 10), bg=C_FORMULA_BG, h_align="center",
                   number_format="0.0%")

        # 会社都合離職者数
        cell_style(ws.cell(r, 11), bg=C_INPUT_BG, h_align="center",
                   number_format="0")
        # 備考
        cell_style(ws.cell(r, 12), bg=C_INPUT_BG)

    # 条件付き書式：賃金UP率（緑=OK, 赤=NG）
    ok_fill = PatternFill(start_color=C_OK_BG, end_color=C_OK_BG, fill_type="solid")
    ng_fill = PatternFill(start_color=C_NG_BG, end_color=C_NG_BG, fill_type="solid")
    ws.conditional_formatting.add(
        "J14:J33",
        CellIsRule(operator="greaterThanOrEqual", formula=["0.03"], fill=ok_fill))
    ws.conditional_formatting.add(
        "J14:J33",
        CellIsRule(operator="lessThan", formula=["0.03"], fill=ng_fill))

    # 条件付き書式：雇用期間
    ws.conditional_formatting.add(
        "G14:G33",
        CellIsRule(operator="greaterThanOrEqual", formula=["180"], fill=ok_fill))
    ws.conditional_formatting.add(
        "G14:G33",
        CellIsRule(operator="lessThan", formula=["180"], fill=ng_fill))


# ─────────────────────────────────────────────
# Sheet 2: 判定結果
# 令和8年度: 2期制・重点支援対象者・情報公表加算に対応
#
# 企業情報参照:
#   入力シート!H5  = 企業規模
#   入力シート!H7  = キャリアアップ計画受理日
#   入力シート!H8  = 就業規則転換規定
#   入力シート!H9  = 情報公表
#
# 対象者参照（src = 14 + i）:
#   B=氏名 C=雇用形態 D=重点支援 F=転換日 G=雇用期間
#   H=転換前賃金 J=賃金UP率 K=会社都合離職者数
# ─────────────────────────────────────────────
def build_result_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8

    merge_title(ws, 1, 1, 16,
                "キャリアアップ助成金（正社員化コース）判定結果　令和8年度（2026年度）版",
                size=13)
    merge_title(ws, 2, 1, 16,
                "※入力シートの値を自動判定します　変更しないでください",
                size=9, bg="DEEAF1", fg="1F4E79")

    ws.row_dimensions[4].height = 50
    col_headers = [
        (1,  4,  "No."),
        (2,  16, "対象者氏名"),
        (3,  12, "雇用期間\n(6ヶ月以上)"),
        (4,  10, "賃金UP\n(3%以上)"),
        (5,  12, "就業規則\n転換規定"),
        (6,  12, "会社都合\n離職者なし"),
        (7,  10, "計画\n受理済み"),
        (8,  10, "総合\n判定"),
        (9,  14, "第1期\n助成額"),
        (10, 14, "第2期\n助成額\n(重点支援のみ)"),
        (11, 13, "情報公表\n加算\n(新設)"),
        (12, 14, "合計\n助成見込額"),
        (13, 14, "第1期\n申請期限\n(開始)"),
        (14, 14, "第1期\n申請期限\n(終了)"),
        (15, 14, "第2期\n申請期限\n(開始)"),
        (16, 14, "第2期\n申請期限\n(終了)"),
    ]
    for col, width, header in col_headers:
        bg = C_NEW_BG if col in (10, 11) else C_HEADER_BG
        fg = C_SUBHEAD_FG if col in (10, 11) else C_TITLE_FG
        cell_style(ws.cell(4, col), value=header, bold=True, size=9,
                   fg=fg, bg=bg, h_align="center", wrap=True)
        set_col_width(ws, col, width)

    for i in range(20):
        r = 5 + i
        ws.row_dimensions[r].height = 18
        src = 14 + i  # 入力シートのデータ行

        # No. / 氏名
        cell_style(ws.cell(r, 1), value=f"=入力シート!A{src}",
                   h_align="center", bg="F8F8F8")
        cell_style(ws.cell(r, 2), value=f"=入力シート!B{src}", bg="F8F8F8")

        # 雇用期間チェック（6ヶ月=180日以上）
        ws.cell(r, 3).value = (
            f'=IF(入力シート!B{src}="","",IF(入力シート!G{src}="","－",'
            f'IF(入力シート!G{src}>=180,"✅ OK","❌ NG")))'
        )
        cell_style(ws.cell(r, 3), h_align="center", bg=C_FORMULA_BG)

        # 賃金UPチェック（3%以上）
        ws.cell(r, 4).value = (
            f'=IF(入力シート!B{src}="","",IF(入力シート!H{src}="","－",'
            f'IF(入力シート!J{src}>=0.03,"✅ OK","❌ NG")))'
        )
        cell_style(ws.cell(r, 4), h_align="center", bg=C_FORMULA_BG)

        # 就業規則転換規定
        ws.cell(r, 5).value = (
            '=IF(入力シート!H8="","⚠️ 未確認",'
            'IF(入力シート!H8="あり","✅ OK","❌ NG"))'
        )
        cell_style(ws.cell(r, 5), h_align="center", bg=C_FORMULA_BG)

        # 会社都合離職者チェック
        ws.cell(r, 6).value = (
            f'=IF(入力シート!B{src}="","",IF(入力シート!K{src}="","⚠️ 未確認",'
            f'IF(入力シート!K{src}=0,"✅ OK","❌ NG")))'
        )
        cell_style(ws.cell(r, 6), h_align="center", bg=C_FORMULA_BG)

        # キャリアアップ計画受理チェック
        ws.cell(r, 7).value = (
            '=IF(入力シート!H7="","⚠️ 未確認","✅ OK")'
        )
        cell_style(ws.cell(r, 7), h_align="center", bg=C_FORMULA_BG)

        # 総合判定
        ws.cell(r, 8).value = (
            f'=IF(入力シート!B{src}="","",IF(AND('
            f'C{r}="✅ OK",'
            f'D{r}="✅ OK",'
            f'E{r}="✅ OK",'
            f'F{r}="✅ OK",'
            f'G{r}="✅ OK"'
            f'),"✅ 申請可","⚠️ 要確認"))'
        )
        cell_style(ws.cell(r, 8), h_align="center", bold=True, bg=C_FORMULA_BG)

        # 第1期助成額（中小:有期・派遣=40万, 無期=20万 / 大企業:有期・派遣=30万, 無期=15万）
        ws.cell(r, 9).value = (
            f'=IF(入力シート!B{src}="","",IF(入力シート!H5="大企業",'
            f'IF(入力シート!C{src}="無期契約",150000,300000),'
            f'IF(入力シート!C{src}="無期契約",200000,400000)))'
        )
        cell_style(ws.cell(r, 9), h_align="right", bg=C_FORMULA_BG,
                   number_format="#,##0円")

        # 第2期助成額（重点支援対象者のみ / 派遣は常に重点支援）
        ws.cell(r, 10).value = (
            f'=IF(入力シート!B{src}="","",IF(OR(入力シート!C{src}="派遣",'
            f'入力シート!D{src}="はい"),'
            f'IF(入力シート!H5="大企業",'
            f'IF(入力シート!C{src}="無期契約",150000,300000),'
            f'IF(入力シート!C{src}="無期契約",200000,400000)),0))'
        )
        cell_style(ws.cell(r, 10), h_align="right", bg=C_NEW_BG,
                   number_format="#,##0円")

        # 情報公表加算（中小=20万, 大企業=15万 / 情報公表=あり の場合のみ）
        ws.cell(r, 11).value = (
            f'=IF(入力シート!B{src}="","",IF(入力シート!H9="あり",'
            f'IF(入力シート!H5="大企業",150000,200000),0))'
        )
        cell_style(ws.cell(r, 11), h_align="right", bg=C_NEW_BG,
                   number_format="#,##0円")

        # 合計助成見込額
        ws.cell(r, 12).value = (
            f'=IF(入力シート!B{src}="","",I{r}+J{r}+K{r})'
        )
        cell_style(ws.cell(r, 12), h_align="right", bold=True, bg=C_FORMULA_BG,
                   number_format="#,##0円")

        # 第1期申請期限（転換日+6ヶ月〜+8ヶ月）
        ws.cell(r, 13).value = (
            f'=IF(入力シート!F{src}="","",EDATE(入力シート!F{src},6))'
        )
        cell_style(ws.cell(r, 13), h_align="center", bg=C_FORMULA_BG,
                   number_format="YYYY/MM/DD")
        ws.cell(r, 14).value = (
            f'=IF(入力シート!F{src}="","",EDATE(入力シート!F{src},8))'
        )
        cell_style(ws.cell(r, 14), h_align="center", bg=C_FORMULA_BG,
                   number_format="YYYY/MM/DD")

        # 第2期申請期限（転換日+12ヶ月〜+14ヶ月 / 重点支援のみ）
        ws.cell(r, 15).value = (
            f'=IF(OR(入力シート!C{src}="派遣",入力シート!D{src}="はい"),'
            f'IF(入力シート!F{src}="","",EDATE(入力シート!F{src},12)),"")'
        )
        cell_style(ws.cell(r, 15), h_align="center", bg=C_NEW_BG,
                   number_format="YYYY/MM/DD")
        ws.cell(r, 16).value = (
            f'=IF(OR(入力シート!C{src}="派遣",入力シート!D{src}="はい"),'
            f'IF(入力シート!F{src}="","",EDATE(入力シート!F{src},14)),"")'
        )
        cell_style(ws.cell(r, 16), h_align="center", bg=C_NEW_BG,
                   number_format="YYYY/MM/DD")

    # 条件付き書式：総合判定
    ok_fill   = PatternFill(start_color=C_OK_BG,   end_color=C_OK_BG,   fill_type="solid")
    warn_fill = PatternFill(start_color=C_WARN_BG,  end_color=C_WARN_BG, fill_type="solid")
    ws.conditional_formatting.add("H5:H24",
        FormulaRule(formula=['H5="✅ 申請可"'], fill=ok_fill))
    ws.conditional_formatting.add("H5:H24",
        FormulaRule(formula=['H5="⚠️ 要確認"'], fill=warn_fill))

    # 凡例
    ws.row_dimensions[26].height = 14
    ws.merge_cells(start_row=26, start_column=1, end_row=26, end_column=16)
    cell_style(ws.cell(26, 1),
               value="【令和8年度の変更点】①2期制導入（第2期は重点支援対象者のみ）　②情報公表加算（中小20万/大企業15万）が新設　③重点支援対象者は合計助成額が2倍",
               size=9, fg="9C0006", bg="FFF2CC", border=False)


# ─────────────────────────────────────────────
# Sheet 3: 要件一覧（参照用）
# ─────────────────────────────────────────────
def build_ref_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    merge_title(ws, 1, 1, 4,
                "キャリアアップ助成金（正社員化コース）要件一覧　令和8年度【参照用】",
                size=13)

    set_col_width(ws, 1, 4)
    set_col_width(ws, 2, 32)
    set_col_width(ws, 3, 55)

    sections = [
        ("【事前要件】転換前に確認", [
            ("キャリアアップ計画", "転換日の前日までに労働局へ提出・受理が必要"),
            ("就業規則の転換規定", "正社員転換を規定した条文が必要（転換日前に届出済みであること）"),
            ("雇用形態", "有期契約・無期契約・派遣社員が対象"),
            ("雇用期間", "転換前に同一事業所で6ヶ月以上継続雇用されていること"),
            ("会社都合離職者",
             "転換日の前日から過去6ヶ月以内に特定受給資格者となる離職者がいないこと"),
            ("雇用保険", "雇用保険適用事業所であること"),
        ]),
        ("【転換時の要件】", [
            ("転換後の雇用形態", "無期・フルタイムの正規雇用労働者であること"),
            ("賃金の増額", "転換前と比較して3%以上の賃金増額が必要"),
            ("転換根拠", "就業規則に基づく転換であること"),
        ]),
        ("【申請時の要件】令和8年度：2期制", [
            ("第1期申請期限",
             "転換後6ヶ月経過した日の翌日〜8ヶ月後末日（2ヶ月間）　※全対象者"),
            ("第2期申請期限",
             "転換後12ヶ月経過した日の翌日〜14ヶ月後末日（2ヶ月間）　※重点支援対象者のみ"),
            ("賃金支払い",
             "各申請期について、6ヶ月分の賃金を支払い済みであること"),
            ("添付書類",
             "雇用契約書、賃金台帳、転換前後の出勤簿、就業規則など"),
        ]),
        ("【重点支援対象者とは】令和8年度新設", [
            ("①派遣労働者", "すべての派遣労働者が対象"),
            ("②雇入れ3年以上の有期労働者",
             "同一事業所で3年以上継続して有期契約で雇用されている者"),
            ("③過去5年間で正規雇用が1年以下の有期労働者",
             "雇入れから3年未満で、過去5年間の正規雇用期間が通算1年以下かつ"
             "過去1年間に正規雇用されていない有期労働者"),
            ("④その他",
             "母子家庭の母等、特定の訓練修了者 など"),
        ]),
        ("【助成額（令和8年度）】", [
            ("有期→正規（通常）中小企業",
             "40万円／人　（第1期のみ）"),
            ("有期→正規（重点支援）中小企業",
             "80万円／人　（第1期40万 + 第2期40万）"),
            ("有期→正規（通常）大企業",
             "30万円／人　（第1期のみ）"),
            ("有期→正規（重点支援）大企業",
             "60万円／人　（第1期30万 + 第2期30万）"),
            ("無期→正規（通常）中小企業",
             "20万円／人　（第1期のみ）"),
            ("無期→正規（重点支援）中小企業",
             "40万円／人　（第1期20万 + 第2期20万）"),
            ("派遣→正規 中小企業",
             "80万円／人　（派遣は全員重点支援扱い）"),
            ("情報公表加算（新設）中小企業",
             "20万円／人　自社HPまたは職場情報総合サイトで転換情報を公表した場合"),
            ("情報公表加算（新設）大企業",
             "15万円／人"),
        ]),
    ]

    row = 3
    for section_title, items in sections:
        ws.row_dimensions[row].height = 22
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        bg = C_NEW_BG if "新設" in section_title or "令和8" in section_title else C_SUBHEAD_BG
        cell_style(ws.cell(row, 1), value=section_title, bold=True, size=11,
                   fg=C_SUBHEAD_FG, bg=bg, border=False)
        row += 1

        for item, desc in items:
            ws.row_dimensions[row].height = 20
            cell_style(ws.cell(row, 1), value="", bg="F8F8F8")
            cell_style(ws.cell(row, 2), value=item, bold=True, size=10)
            cell_style(ws.cell(row, 3), value=desc, size=10, wrap=True)
            row += 1

        row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    cell_style(ws.cell(row, 1),
               value="※ 要件・助成額は法改正により変更となる場合があります。"
                     "申請前に必ず厚生労働省の最新資料をご確認ください。",
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

    for ws in [ws_input, ws_result]:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1

    path = "/home/user/umb/キャリアアップ助成金_要件チェックシート_令和8年度.xlsx"
    wb.save(path)
    print(f"✅ 作成完了: {path}")


if __name__ == "__main__":
    main()
