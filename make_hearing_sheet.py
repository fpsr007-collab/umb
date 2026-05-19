"""就業規則コンサルヒアリングシート2026 Word生成スクリプト"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ページ余白設定
section = doc.sections[0]
section.page_width  = Cm(21)
section.page_height = Cm(29.7)
section.top_margin    = Cm(1.5)
section.bottom_margin = Cm(1.5)
section.left_margin   = Cm(1.8)
section.right_margin  = Cm(1.8)

# ========== ヘルパー関数 ==========

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), val.get('val', 'single'))
            el.set(qn('w:sz'), str(val.get('sz', 6)))
            el.set(qn('w:color'), val.get('color', '000000'))
            tcBorders.append(el)
    tcPr.append(tcBorders)

def set_table_border(table, color='AAAAAA'):
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:color'), color)
        tblBorders.append(el)
    tblPr.append(tblBorders)

def add_section_title(doc, number, title, bg='222222', fg='FFFFFF'):
    """濃い背景色のセクションタイトル行"""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run(f'{number}. {title}')
    run.bold = True
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor.from_string(fg)
    doc.add_paragraph()  # spacing
    return tbl

def add_subsection(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(1)
    p.paragraph_format.left_indent  = Cm(0)
    run = p.add_run(f'◆ {title}')
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    # 下線で区切り
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:color'), '888888')
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_memo(doc, text):
    """社労士メモ（★）黄色背景段落"""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, 'FFF9C4')
    set_cell_borders(cell,
        top={'val':'dashed','sz':4,'color':'E6A800'},
        bottom={'val':'dashed','sz':4,'color':'E6A800'},
        left={'val':'dashed','sz':4,'color':'E6A800'},
        right={'val':'dashed','sz':4,'color':'E6A800'},
    )
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    r1 = p.add_run('★ ')
    r1.bold = True
    r1.font.color.rgb = RGBColor(0x7A, 0x4F, 0x00)
    r1.font.size = Pt(9)
    r2 = p.add_run(text)
    r2.font.color.rgb = RGBColor(0x7A, 0x4F, 0x00)
    r2.font.size = Pt(9)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(1)

def add_row(doc, text, size=10):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(1)
    p.paragraph_format.left_indent  = Cm(0.3)
    for run in p.runs:
        run.font.size = Pt(size)

def add_note(doc, text):
    p = doc.add_paragraph(f'※ {text}')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(1)
    p.paragraph_format.left_indent  = Cm(0.5)
    for run in p.runs:
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

def add_blank_box(doc, height_cm=1.0, label=''):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    set_table_border(tbl, 'AAAAAA')
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, 'FAFAFA')
    p = cell.paragraphs[0]
    if label:
        run = p.add_run(label)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x55,0x55,0x55)
    # 高さ確保のため空行を追加
    lines = max(1, int(height_cm / 0.5))
    for _ in range(lines):
        cell.add_paragraph()
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(1)

def add_check_row(doc, items_per_line, items):
    """チェックボックス風の行（□ テキスト形式）"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(1)
    p.paragraph_format.left_indent  = Cm(0.3)
    count = 0
    for item in items:
        if count > 0 and count % items_per_line == 0:
            p.add_run('\n')
        run = p.add_run(f'□ {item}　')
        run.font.size = Pt(10)
        count += 1

def add_grid_table(doc, headers, rows_data, col_widths=None):
    """汎用テーブル"""
    n_cols = len(headers)
    tbl = doc.add_table(rows=1+len(rows_data), cols=n_cols)
    tbl.style = 'Table Grid'
    set_table_border(tbl, 'AAAAAA')
    # ヘッダー
    hdr = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        set_cell_bg(cell, 'DDDDDD')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(9)
    # データ行
    for ri, row_data in enumerate(rows_data):
        row = tbl.rows[ri+1]
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            p = cell.paragraphs[0]
            run = p.add_run(val)
            run.font.size = Pt(9.5)
    # 列幅
    if col_widths:
        for ri in range(len(tbl.rows)):
            for ci, w in enumerate(col_widths):
                tbl.rows[ri].cells[ci].width = Cm(w)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(1)
    return tbl

def add_badge(run_container_para, text, color='C0392B'):
    """法改正バッジ風テキスト"""
    run = run_container_para.add_run(f' [{text}]')
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor.from_string(color)
    run.bold = True

def spacer(doc, pt=3):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(pt)

# ========== ドキュメント本文 ==========

# タイトル
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_title.add_run('就業規則コンサルヒアリングシート 2026')
r.bold = True
r.font.size = Pt(16)
r.font.underline = True

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
rs = p_sub.add_run('2024年・2025年 法改正対応版　／　★印は社労士メモ（顧客へ渡す際は削除してください）')
rs.font.size = Pt(9)
rs.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
spacer(doc, 4)

# 顧客情報ヘッダー
tbl_hdr = doc.add_table(rows=2, cols=3)
tbl_hdr.style = 'Table Grid'
set_table_border(tbl_hdr, '333333')
labels = [['会社名', '', ''], ['担当社労士', 'ヒアリング日', '']]
for ri, row_labels in enumerate(labels):
    for ci, label in enumerate(row_labels):
        cell = tbl_hdr.rows[ri].cells[ci]
        set_cell_bg(cell, 'EEEEEE')
        if label:
            p = cell.paragraphs[0]
            run = p.add_run(f'{label}：')
            run.bold = True
            run.font.size = Pt(10)
            cell.add_paragraph()
col_widths_hdr = [7, 5, 5]
for ri in range(2):
    for ci, w in enumerate(col_widths_hdr):
        tbl_hdr.rows[ri].cells[ci].width = Cm(w)
# 会社名セルは2行にわたるよう結合
tbl_hdr.rows[0].cells[0].merge(tbl_hdr.rows[0].cells[1])
tbl_hdr.rows[0].cells[0].merge(tbl_hdr.rows[0].cells[2])
spacer(doc, 6)

# ==============================
# 1. 就業規則作成の背景・目的
# ==============================
add_section_title(doc, 1, '就業規則作成の背景・目的')
add_memo(doc, '作成したい「表向きの理由」と「裏の理由（実際のトラブル・課題）」を分けて把握する。裏の理由が就業規則のコアになる。')
add_row(doc, '表向きの理由：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿')
add_row(doc, '実際の課題・背景：')
add_blank_box(doc, 1.2)
add_row(doc, '以前の就業規則　□ 無　□ 有（最終改定：＿＿＿年頃）')
add_memo(doc, '既存規則がある場合、現行のどこが機能していないかを確認。古いまま放置されているケースが多い。')
spacer(doc)

# ==============================
# 2. 会社概要
# ==============================
add_section_title(doc, 2, '会社概要')
add_row(doc, '会社設立：＿＿年　＿＿月　／　業種・業態：＿＿＿＿＿＿　／　資本金：＿＿＿＿万円')
add_row(doc, '本社所在地：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿')

add_subsection(doc, '雇用状況（おおよその数字でOK）')
add_memo(doc, '無理に全員分を確認しなくてよい。10人以上で届出義務、300人超で育休取得率の公表義務あり。')
add_grid_table(doc,
    ['区分','役員','正社員','契約社員','パート','嘱託','アルバイト','派遣','合計'],
    [['人数','','','','','','','','']],
    col_widths=[1.8, 1.4, 1.4, 1.4, 1.4, 1.4, 1.6, 1.4, 1.4]
)
add_note(doc, '常時使用する労働者が10人以上で就業規則の届出義務あり（労基法89条）')

add_row(doc, '性別：男性＿＿人　女性＿＿人')
add_row(doc, '年齢層：20代＿＿　30代＿＿　40代＿＿　50代＿＿　60代以上＿＿')
add_row(doc, '職種：営業＿＿人　事務＿＿人　技術・SE＿＿人　製造・現場＿＿人　その他（＿＿＿）＿＿人')

add_subsection(doc, '勤務地・テレワーク')
add_row(doc, '在宅・テレワーク：□ 無　□ 有（□ 全員可　□ 限定：＿＿＿＿＿）')
add_row(doc, '関連会社・グループ会社出向：□ 無　□ 有（＿＿＿＿＿）')
add_row(doc, '外国人労働者：□ 無　□ 有（在留資格：＿＿＿＿＿）')
add_memo(doc, 'テレワーク規程・在宅勤務規程を別途作成するか確認。外国人がいる場合は就労可能な業務範囲も確認。')

add_subsection(doc, '障害者雇用')
p_badge = doc.add_paragraph()
p_badge.paragraph_format.left_indent = Cm(0.3)
r1 = p_badge.add_run('□ 無　□ 有（身体：＿人　知的：＿人　精神：＿人）')
r1.font.size = Pt(10)
add_badge(p_badge, '2024年改正', '1A6FA8')
add_note(doc, '法定雇用率：2024年4月〜2.5%、2026年7月〜2.7%（常用労働者40人以上で義務）')
spacer(doc)

# ==============================
# 3. 採用・入社手続き
# ==============================
add_section_title(doc, 3, '採用・入社手続き')

add_subsection(doc, '入社時の提出書類')
add_memo(doc, '身元保証書の有無と更新管理を確認。身元保証法改正（2020年）で期間上限5年。')
add_check_row(doc, 4, [
    'マイナンバー', '住民票記載事項証明書', '扶養控除等申告書', '誓約書（秘密保持・競業避止）',
    '身元保証書', '源泉徴収票', '資格証明書', '自動車免許証',
    '給与振込口座', 'その他：＿＿＿＿＿＿＿'
])

add_subsection(doc, '試用期間・有期契約')
p_badge2 = doc.add_paragraph()
p_badge2.paragraph_format.left_indent = Cm(0.3)
r2 = p_badge2.add_run('試用期間：□ 無　□ 有（＿＿か月）　有期契約期間：□ 無　□ 有（＿＿か月）')
r2.font.size = Pt(10)
add_badge(p_badge2, '2024年4月 明示ルール改正', 'C0392B')
add_memo(doc, '試用期間≠有期契約の違いを必ず説明。2024年4月から「就業場所・業務の変更範囲」「有期契約の更新上限・無期転換申込機会」の明示が義務化。')
add_row(doc, '有期契約の更新上限：□ 無　□ 有（上限回数：＿＿回 / ＿＿年）')
add_row(doc, '無期転換ルール（5年超）の対象者：□ 無　□ 有（＿＿人）')
add_note(doc, '無期転換申込機会・転換後の労働条件を就業規則に明記する必要あり')

add_subsection(doc, '同一労働同一賃金チェック')
add_memo(doc, '正社員とパート・有期社員の待遇差の説明義務（パートタイム・有期雇用労働法）。各手当の支給根拠が説明できるか確認。')
add_row(doc, '正規・非正規の待遇差の合理的説明：□ できる　□ 要検討　□ 未整備')
add_row(doc, '課題・コメント：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿')
spacer(doc)

# ==============================
# 4. 労働時間・休憩
# ==============================
add_section_title(doc, 4, '労働時間・休憩')

add_subsection(doc, '労働時間制の種類')
add_check_row(doc, 3, [
    '通常の労働時間制（固定）', '1か月単位変形労働時間制', '1年単位変形労働時間制',
    'フレックスタイム制', '事業場外みなし労働時間制', '裁量労働制（専門型・企画型）'
])
add_memo(doc, '変形労働・フレックスは労使協定が必要。裁量労働は要件が厳格。自社の実態に合った制度か確認する。')

add_subsection(doc, '通常（固定時間）')
add_row(doc, '始業（＿＿時＿＿分）　終業（＿＿時＿＿分）')

add_subsection(doc, 'シフト制')
add_memo(doc, 'シフト表の有無、いつシフトを確定するかを確認。直前変更ルールも規定に入れるか検討。')
add_grid_table(doc,
    ['パターン', '始業', '終業', '備考'],
    [['A','　　時　　分','　　時　　分',''],
     ['B','　　時　　分','　　時　　分',''],
     ['C','　　時　　分','　　時　　分',''],
     ['D','　　時　　分','　　時　　分','']],
    col_widths=[2, 4, 4, 7]
)

add_subsection(doc, '休憩時間')
add_memo(doc, '残業が多い会社では休憩時間で法定時間（6h超→45分、8h超→60分）を満たすよう調整することがある。')
add_row(doc, '（＿＿時＿＿分〜＿＿時＿＿分）　（＿＿時＿＿分〜＿＿時＿＿分）　合計（＿＿）分')

add_subsection(doc, '出退勤管理')
add_row(doc, '管理方法：□ 無　□ 出勤簿　□ タイムカード　□ ICカード　□ 勤怠システム（＿＿＿＿＿）')
add_row(doc, '欠勤・遅刻・早退の連絡：□ 電話　□ メール　□ 取得届　□ その他（＿＿＿＿）')

add_subsection(doc, '時間外・休日・深夜労働')
add_row(doc, '残業：□ 有　□ 無　　深夜業務（22時〜5時）：□ 有　□ 無　　休日業務：□ 有　□ 無')
add_row(doc, '36協定 締結：□ 済　□ 未締結　　特別条項：□ 有　□ 無　　月の上限残業時間：＿＿時間')
add_note(doc, '時間外上限規制：月45h・年360h、特別条項でも月100h未満・年720h以内（労基法36条）')
add_row(doc, '※残業代・残業時間に関する過去のトラブル')
add_blank_box(doc, 1.0)
spacer(doc)

# ==============================
# 5. 休日・休暇
# ==============================
add_section_title(doc, 5, '休日・休暇')

add_subsection(doc, '休日')
add_memo(doc, '年間カレンダーを作成・添付する。週1日以上の法定休日を確認。')
add_row(doc, '定休日：□月　□火　□水　□木　□金　□土　□日　□祝日')
add_row(doc, '夏季：＿＿日　GW：＿＿日　年末年始：＿＿日　年間休日合計：＿＿日')
add_row(doc, 'シフト制（週休＿＿日）※特定の曜日が決まっていない場合')

add_subsection(doc, '年次有給休暇')
p_badge3 = doc.add_paragraph()
p_badge3.paragraph_format.left_indent = Cm(0.3)
r3 = p_badge3.add_run('一斉付与制度：□ 有　□ 無　　半日単位：□ 有　□ 無　　時間単位：□ 有　□ 無（最大＿＿日分）')
r3.font.size = Pt(10)
add_badge(p_badge3, '年5日取得義務', 'C0392B')
add_memo(doc, '年5日の取得義務（労基法39条7項）を知っているか確認。管理簿の整備状況も確認。')
add_row(doc, '届出期限：取得予定日の＿＿日以上前　届出方法：□ メール　□ 取得届　□ システム　□ 口頭')
add_row(doc, '有給休暇管理簿：□ 整備済　□ 未整備')
add_row(doc, '※有給休暇に関する過去のトラブル')
add_blank_box(doc, 1.0)

add_subsection(doc, '特別休暇')
add_row(doc, '□ 無　□ 有')
add_grid_table(doc,
    ['種別', '日数', '有給/無給', '備考'],
    [['慶弔休暇（結婚・忌引）', '', '', ''],
     ['生理休暇', '必要日数', '無給', '（法定）'],
     ['裁判員休暇', '', '', ''],
     ['ボランティア休暇', '', '', ''],
     ['その他（＿＿＿＿＿）', '', '', '']],
    col_widths=[5, 2.5, 3, 6.5]
)

add_subsection(doc, '育児・介護休業')
p_badge4 = doc.add_paragraph()
p_badge4.paragraph_format.left_indent = Cm(0.3)
r4 = p_badge4.add_run('産後パパ育休の実績：□ 有　□ 無　　育児休業（通常）の実績：□ 有　□ 無')
r4.font.size = Pt(10)
add_badge(p_badge4, '2025年4月施行改正', 'C0392B')
add_memo(doc, '2025年4月施行：①子の看護休暇の対象拡大（小学校3年生修了まで）・事由拡大、②所定外労働免除の対象拡大（3歳→小学校就学前）、③育児休業の柔軟化。300人超は育休取得率の公表義務。')
add_row(doc, '介護休業の実績：□ 有　□ 無　　子の看護休暇の実績：□ 有　□ 無')
add_row(doc, '育休取得率の公表（300人超）：□ 対象外　□ 公表済　□ 未対応')
add_row(doc, '育休・介護休業規程の整備：□ 整備済　□ 未整備　□ 要更新')

add_subsection(doc, '休職制度')
add_memo(doc, 'メンタルヘルス不調による休職が増加。復職手続き・リハビリ勤務・自然退職の条件を明確にする。')
add_row(doc, '休職制度：□ 無　□ 有（傷病：＿＿か月、その他：＿＿＿＿＿）')
add_row(doc, '※メンタルヘルス・長期療養に関する過去のトラブル')
add_blank_box(doc, 1.0)
spacer(doc)

# ==============================
# 6. 賃金・退職金・定年
# ==============================
add_section_title(doc, 6, '賃金・退職金・定年')

add_subsection(doc, '正社員・契約社員の賃金')
add_row(doc, '基本給：月給（＿＿＿＿円〜＿＿＿＿円）　日給（＿＿＿＿円〜＿＿＿＿円）')
add_grid_table(doc,
    ['手当名', '金額（円）', '支給条件・備考'],
    [['　　　　　手当','',''],
     ['　　　　　手当','',''],
     ['　　　　　手当','',''],
     ['　　　　　手当','',''],
     ['通勤交通費','上限　　　　円','□ 実費　□ 定額']],
    col_widths=[4, 3.5, 9.5]
)
add_row(doc, '歩合給：□ 無　□ 有（計算書を従業員に公開：□ 有　□ 無）')

add_subsection(doc, 'パート・アルバイトの賃金')
add_row(doc, '時給：＿＿＿＿円〜＿＿＿＿円　通勤交通費：□ 無　□ 有（上限＿＿＿円）')
add_note(doc, '地域別最低賃金（2025年度見込み）を必ず確認。全国加重平均1,000円超が目安。')

add_subsection(doc, '共通事項')
add_row(doc, '賃金締切日・支払日：毎月＿＿日締切、＿＿日払（□ 当月　□ 翌月　□ 翌々月）')
add_row(doc, '銀行休日時：□ 前日払　□ 翌日払　　支払方法：□ 銀行振込　□ 現金払')
add_row(doc, '税・社保以外の控除：□ 無　□ 有（＿＿＿＿＿＿＿＿）')
add_row(doc, '昇給：□ 無　□ 有（年＿＿回、＿＿月）　賞与：□ 無　□ 有（＿月・＿月・＿月）')

add_subsection(doc, '退職金・定年')
add_row(doc, '退職金：□ 無　□ 有（□ 中退共　□ 生命保険　□ 確定拠出年金（iDeCo企業型）　□ その他：＿＿＿）')
add_row(doc, '定年：□ 有（＿＿歳）　□ 無　　継続雇用制度：□ 有（＿＿歳まで）　□ 無')
add_row(doc, '65歳超雇用推進（70歳就業確保）：□ 対応済　□ 検討中　□ 予定なし')
add_note(doc, '高年齢者雇用安定法：65歳までは雇用確保義務、70歳までは就業確保努力義務（2021年〜）')

add_subsection(doc, '退職手続き')
add_memo(doc, '民法では2週間前が原則。就業規則での延長は1か月程度が多い。長すぎると無効になるリスクあり。')
add_row(doc, '自己都合退職の届出期限：退職する＿＿日（または＿＿か月）以上前に届出')
add_row(doc, '※退職に関する過去のトラブル（入社時：＿＿＿＿＿＿＿＿＿＿＿＿）')
add_row(doc, '（退社時：＿＿＿＿＿＿＿＿＿＿＿＿）')
spacer(doc)

# ==============================
# 7. 社会保険・福利厚生
# ==============================
add_section_title(doc, 7, '社会保険・福利厚生')

add_row(doc, '社会保険（厚生年金・健康保険）：□ 全員加入済　□ 未加入者あり（＿＿人）')
add_row(doc, '雇用保険：□ 全員加入済　□ 未加入者あり（＿＿人）')
add_note(doc, '社会保険の適用拡大：2024年10月〜51人以上、2026年10月〜全企業（短時間労働者）')
add_row(doc, '健康診断：□ 無　□ 有（毎年＿月）　ストレスチェック（50人以上）：□ 実施済　□ 未実施')
add_row(doc, '研修制度：□ 有　□ 無　　資格取得奨励制度：□ 有　□ 無')
add_row(doc, '副業・兼業：□ 禁止　□ 全員可　□ 限定許可（＿＿＿＿＿＿）')
add_memo(doc, '副業容認の場合は「副業・兼業に関するガイドライン（厚労省）」に基づく規定と、労働時間通算の管理方法を検討。')
spacer(doc)

# ==============================
# 8. 服務規律・ハラスメント・懲戒
# ==============================
add_section_title(doc, 8, '服務規律・ハラスメント・懲戒')

add_subsection(doc, 'ハラスメント対策（全企業義務）')
add_memo(doc, 'パワハラ防止措置は全事業主に義務（2022年4月〜中小企業含む）。相談窓口の設置・周知が必要。')
add_row(doc, '相談窓口の設置：□ 有（担当：＿＿＿＿＿）　□ 無')
add_row(doc, 'ハラスメント研修：□ 定期実施　□ 実施済（単発）　□ 未実施')
add_row(doc, '※ハラスメントの過去のトラブル事例')
add_blank_box(doc, 1.0)

add_subsection(doc, '情報セキュリティ・SNS規定')
add_memo(doc, 'SNS投稿による風評被害・情報漏洩のトラブルが増加。就業規則にSNS利用ルールを明記することを推奨。')
add_row(doc, 'PC・スマホの貸与：□ 有（□ 全員　□ 限定）　□ 無')
add_row(doc, 'SNS・インターネット利用規定：□ 有　□ 無（要作成）')
add_row(doc, '秘密保持・競業避止（退職後）：□ 誓約書あり　□ 規定のみ　□ 未対応')
add_row(doc, '同業他社への転職制限：□ 無　□ 有（制限範囲：＿＿＿＿＿）')
add_memo(doc, '競業避止条項は「期間・地域・職種の範囲・代償措置」がなければ無効になりやすい。内容を精査する。')

add_subsection(doc, 'その他の服務事項')
add_row(doc, '社有車の業務使用：□ 有　□ 無')
add_row(doc, 'マイカー通勤：□ 有（駐車場：□有　□無　業務使用：□有　□無）　□ 無')
add_row(doc, '自転車通勤：□ 有（駐輪場：□有　□無　業務使用：□有　□無）　□ 無')
add_row(doc, 'アルコールチェック：□ 有（義務対象）　□ 有（任意）　□ 無')
add_row(doc, '出張：□ 無　□ 有（地域：＿＿＿＿＿　宿泊：□ 有　□ 無）')
add_row(doc, '損害賠償・弁償規定：□ 不要　□ 要整備（事例：＿＿＿＿＿＿）')
add_row(doc, '解雇した事例：□ 無　□ 有（時期：＿＿＿　概要：＿＿＿＿＿＿＿＿＿＿＿）')
spacer(doc)

# ==============================
# 9. 就業規則の整備・届出・周知
# ==============================
add_section_title(doc, 9, '就業規則の整備・届出・周知')
add_memo(doc, '就業規則は常時10人以上で届出義務あり。効力発生には「周知」が不可欠（労基法106条）。届け出ただけでは不十分。')
add_row(doc, '就業規則の届出：□ 済（最終届出：＿＿年）　□ 未届出　□ 届出義務なし（9人以下）')
add_row(doc, '周知方法：□ 備付け　□ 書面交付　□ 社内イントラ・共有フォルダ　□ 未周知')
add_row(doc, '別規程（パート・契約社員用）：□ 作成済　□ 不要　□ 要作成')
add_row(doc, '労働者代表の選出・意見書：□ 対応済　□ 未対応')
add_row(doc, 'その他付属規程（必要なもの）：')
add_check_row(doc, 4, [
    '賃金規程', '育児・介護休業規程', 'テレワーク規程', 'ハラスメント防止規程',
    '副業・兼業規程', '情報セキュリティ規程', '退職金規程', 'その他：＿＿＿＿＿'
])
spacer(doc)

# ==============================
# 10. 社労士メモ（最終確認）
# ==============================
add_section_title(doc, 10, '社労士メモ（最終確認・提案事項）　※顧客渡し版では削除', bg='7A4F00')

p_m1 = doc.add_paragraph()
p_m1.paragraph_format.left_indent = Cm(0.3)
r_m1 = p_m1.add_run('【優先度 高】すぐに対応が必要な事項')
r_m1.bold = True
r_m1.font.size = Pt(10)
r_m1.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
add_blank_box(doc, 1.5)

p_m2 = doc.add_paragraph()
p_m2.paragraph_format.left_indent = Cm(0.3)
r_m2 = p_m2.add_run('【提案・オプション】追加でご提案できる事項')
r_m2.bold = True
r_m2.font.size = Pt(10)
r_m2.font.color.rgb = RGBColor(0x15, 0x5B, 0x24)
add_blank_box(doc, 1.5)

add_row(doc, '見積もり目安：就業規則作成（＿＿＿＿円）　付属規程（＿＿＿円/1本）　次回日程：＿＿＿＿＿')

# ========== 保存 ==========
out_path = '/home/user/umb/就業規則コンサルヒアリングシート2026.docx'
doc.save(out_path)
print(f'saved: {out_path}')
