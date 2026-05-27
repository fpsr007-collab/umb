from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width = Cm(21.0)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)

# ---- ヘルパー関数 ----

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    elif level == 2:
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0x27, 0x6D, 0xC4)
    elif level == 3:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x37, 0x86, 0xC8)
    run.font.name = '游ゴシック'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '游ゴシック')
    return p

def add_body(doc, text, bold=False, indent=0):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(10.5)
    run.font.name = '游明朝'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')
    p.paragraph_format.space_after = Pt(4)
    return p

def add_bullet(doc, text, indent=0.5):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(indent)
    run = p.add_run(text)
    run.font.size = Pt(10.5)
    run.font.name = '游明朝'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')
    p.paragraph_format.space_after = Pt(3)
    return p

def add_info_table(doc, rows_data):
    table = doc.add_table(rows=len(rows_data), cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    col_widths = [Cm(3.8), Cm(12.4)]
    for i, row_data in enumerate(rows_data):
        row = table.rows[i]
        row.cells[0].width = col_widths[0]
        row.cells[1].width = col_widths[1]
        cell0 = row.cells[0]
        set_cell_bg(cell0, 'D6E4F7')
        p0 = cell0.paragraphs[0]
        run0 = p0.add_run(row_data[0])
        run0.bold = True
        run0.font.size = Pt(10)
        run0.font.name = '游ゴシック'
        run0._element.rPr.rFonts.set(qn('w:eastAsia'), '游ゴシック')
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell1 = row.cells[1]
        p1 = cell1.paragraphs[0]
        run1 = p1.add_run(row_data[1])
        run1.font.size = Pt(10)
        run1.font.name = '游明朝'
        run1._element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')
    doc.add_paragraph()
    return table

def add_course_divider(doc, course_name, color_hex='2E74B5'):
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.rows[0].cells[0]
    set_cell_bg(cell, color_hex)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(course_name)
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run.font.name = '游ゴシック'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '游ゴシック')
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def add_new_badge(doc):
    """令和8年度の変更ポイントを強調するバッジ行"""
    p = doc.add_paragraph()
    run = p.add_run('★ 令和8年度 変更ポイント')
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    run.font.name = '游ゴシック'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '游ゴシック')
    p.paragraph_format.space_after = Pt(2)
    return p

# ==============================
# 表紙
# ==============================
doc.add_paragraph()
doc.add_paragraph()

title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_run = title_p.add_run('働き方改革推進支援助成金')
title_run.bold = True
title_run.font.size = Pt(22)
title_run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
title_run.font.name = '游ゴシック'
title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '游ゴシック')

subtitle_p = doc.add_paragraph()
subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle_run = subtitle_p.add_run('各コース概要・申請ポイントまとめ（令和8年度版）')
subtitle_run.bold = True
subtitle_run.font.size = Pt(14)
subtitle_run.font.color.rgb = RGBColor(0x27, 0x6D, 0xC4)
subtitle_run.font.name = '游ゴシック'
subtitle_run._element.rPr.rFonts.set(qn('w:eastAsia'), '游ゴシック')

doc.add_paragraph()

note_p = doc.add_paragraph()
note_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
note_run = note_p.add_run('厚生労働省・各都道府県労働局資料をもとに作成')
note_run.font.size = Pt(10)
note_run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
note_run.font.name = '游明朝'
note_run._element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')

doc.add_page_break()

# ==============================
# 助成金の全体概要
# ==============================
add_heading(doc, '■ 助成金の全体概要（令和8年度）', level=1)
add_body(doc,
    '働き方改革推進支援助成金は、生産性を向上させながら労働時間の短縮・年次有給休暇の取得促進・'
    '勤務間インターバルの導入等に取り組む中小企業事業主等を支援する助成金制度です。'
    '令和8年度は予算が前年度比約9億円増の101億円に増額され、新たな加算措置も追加されました。'
)
add_body(doc,
    '交付申請受付開始：令和8年4月13日（月）　／　申請受付期限：令和8年11月30日（月）午後5時\n'
    '※予算上限に達した場合は期限前に受付終了',
    bold=True)
doc.add_paragraph()

add_info_table(doc, [
    ('対象', '労働者災害補償保険の適用を受ける中小企業事業主（業種・規模要件あり）'),
    ('助成率', '対象経費の3/4（労働者数30人以下で対象取り組みが特定の場合は4/5）'),
    ('申請窓口', '各都道府県労働局 雇用環境・均等部（室）\n（電子申請：Jグランツも利用可）'),
    ('コース数', '5コース（個社向け3コース＋団体・事業主団体向け2コース）'),
    ('令和8年度の\n主な変更点',
     '①勤務間インターバル導入コースの上限額を120万円→150万円に引き上げ\n'
     '②全コース共通で「割増賃金率引上げ加算」を新設\n'
     '　（割増賃金率5%以上引上げ→25万円加算、50%以上→100万円加算）\n'
     '③業種別課題対応コースに「所定外労働時間の削減」の成果目標を追加\n'
     '④予算額を92億円→101億円に増額'),
])

doc.add_page_break()

# ==============================
# コース1: 労働時間短縮・年休促進支援コース
# ==============================
add_course_divider(doc, 'コース① 労働時間短縮・年休促進支援コース', '1F497D')

add_heading(doc, '【概要】', level=3)
add_body(doc,
    '時間外労働の削減、年次有給休暇・特別休暇の取得促進に向けた環境整備に取り組む中小企業事業主を支援します。'
    '令和2年4月から中小企業にも時間外労働の上限規制が適用されており、本コースはその対応を後押しします。'
    '全業種が対象となる最も基本的なコースです。'
)

add_heading(doc, '【対象事業主の主な要件】', level=3)
add_info_table(doc, [
    ('規模要件', '中小企業事業主（業種別の資本金・労働者数の基準内）'),
    ('36協定', '時間外・休日労働に関する協定届を提出していること\n（月60時間超または月80時間超の36協定を締結している事業場が対象）'),
    ('年休整備', '全対象事業場で年5日の年次有給休暇取得に向けた就業規則等を整備済みであること'),
])

add_heading(doc, '【助成対象の主な取り組み（成果目標）】', level=3)
add_bullet(doc, '全ての対象事業場の36協定の時間外・休日労働時間数を縮減する（月60時間以下 等）')
add_bullet(doc, '年次有給休暇の計画的付与制度を新たに導入する')
add_bullet(doc, '時間単位年次有給休暇制度を新たに導入する')
add_bullet(doc, '特別休暇（病気・子の看護・介護・教育訓練・不妊治療等）を新たに導入する')
doc.add_paragraph()

add_heading(doc, '【助成額・助成率】', level=3)
add_info_table(doc, [
    ('助成率', '対象経費の3/4\n（労働者数30人以下かつ所要額が30万円超の特定取り組みは4/5）'),
    ('上限額（目安）',
     '・月60時間以下に縮減：最大150万円\n'
     '・賃金引上げ加算（3%・5%・7%以上の区分で加算額が異なる）'),
    ('割増賃金率\n引上げ加算\n【令和8年度新設】',
     '・割増賃金率を5%以上引き上げた場合：25万円加算\n'
     '・月60時間超の割増賃金率を50%以上に引き上げた場合：100万円加算'),
    ('対象経費例', '就業規則・労使協定等の作成・変更費用、コンサルタント費用、\n研修費、労務管理ソフト導入費 等'),
])

add_heading(doc, '【申請のポイント】', level=3)
add_bullet(doc, '【事前】交付申請前に「成果目標」を設定し、交付決定を受けてから取り組みを実施すること')
add_bullet(doc, '36協定の月あたり時間外労働の上限が高いほど対象になりやすいが、縮減目標が小さいと助成額も低くなる')
add_bullet(doc, '年次有給休暇の計画的付与制度・時間単位制度は未導入の事業場のみ申請可能')
add_bullet(doc, '特別休暇は「新たに導入」が条件のため、既存の休暇制度がある場合は対象外')
add_bullet(doc, '令和8年度新設の「割増賃金率引上げ加算」は就業規則改定が必要。賃金コスト増も踏まえて計画する')
add_bullet(doc, '賃金引上げ加算との併用も可能。加算の組み合わせで助成総額を最大化する設計を検討する')
add_bullet(doc, '申請期限（令和8年11月30日）までに交付申請を行い、翌年2月14日までに取り組みを完了する必要あり')
doc.add_paragraph()
doc.add_page_break()

# ==============================
# コース2: 勤務間インターバル導入コース
# ==============================
add_course_divider(doc, 'コース② 勤務間インターバル導入コース', '276DC4')

add_heading(doc, '【概要】', level=3)
add_body(doc,
    '勤務終了後から次の勤務開始まで一定の「休息時間（インターバル）」を設ける制度の導入を支援します。'
    '2019年4月から中小企業も制度導入が努力義務化されており、'
    '労働者の睡眠時間・生活時間の確保と過重労働防止が目的です。'
    '令和8年度は新規導入（11時間以上）の助成上限額が120万円から150万円に引き上げられました。'
)

add_new_badge(doc)
add_body(doc, '11時間以上の勤務間インターバルを新規導入する場合の助成上限額が 120万円 → 150万円 に引き上げ', bold=True)
doc.add_paragraph()

add_heading(doc, '【対象事業主の主な要件】', level=3)
add_info_table(doc, [
    ('規模要件', '中小企業事業主'),
    ('現状確認', '勤務間インターバルが未導入、または一部事業場に未導入の状態であること'),
    ('年休整備', '全対象事業場で年5日の年次有給休暇取得に向けた就業規則等を整備済みであること'),
])

add_heading(doc, '【助成対象の主な取り組み（成果目標）】', level=3)
add_bullet(doc, '勤務間インターバルを「新たに」導入する（就業規則等への規定）')
add_bullet(doc, 'インターバル時間の拡大（例：9時間→11時間以上）')
add_bullet(doc, '対象となる労働者の範囲の拡大')
doc.add_paragraph()

add_heading(doc, '【助成額・助成率】', level=3)
add_info_table(doc, [
    ('助成率', '対象経費の3/4（特定条件では4/5）'),
    ('上限額',
     '・新規導入（11時間以上）：最大150万円【令和8年度に引き上げ。令和7年度は120万円】\n'
     '・インターバル時間拡大・対象範囲拡大：取り組み内容により異なる\n'
     '・賃金引上げ加算（3%・5%・7%以上の区分）'),
    ('割増賃金率\n引上げ加算\n【令和8年度新設】',
     '・割増賃金率を5%以上引き上げた場合：25万円加算\n'
     '・月60時間超の割増賃金率を50%以上に引き上げた場合：100万円加算'),
    ('対象経費例', '就業規則・労使協定の変更費用、コンサルタント費用、\nシフト管理システム導入費、研修費 等'),
])

add_heading(doc, '【申請のポイント】', level=3)
add_bullet(doc, '「勤務間インターバル」を就業規則等に明記することが必須。口頭運用では不可')
add_bullet(doc, 'インターバルを長く設定するほど上限額が高くなる（9時間・10時間・11時間以上で区分）')
add_bullet(doc, '令和8年度から上限150万円に拡充。より積極的な導入を検討する好機')
add_bullet(doc, '導入後は実際の運用実績（出退勤記録等）を保存しておくこと（実施結果報告時に必要）')
add_bullet(doc, '賃金引上げ加算・割増賃金率引上げ加算との併用で助成総額の最大化が可能')
add_bullet(doc, '深夜・早朝業務の多い業種（飲食・宿泊・医療等）で特に効果的な取り組み')
doc.add_paragraph()
doc.add_page_break()

# ==============================
# コース3: 業種別課題対応コース
# ==============================
add_course_divider(doc, 'コース③ 業種別課題対応コース', '37558C')

add_heading(doc, '【概要】', level=3)
add_body(doc,
    '時間外労働の上限規制の適用が猶予・除外されていた業種（建設業・運送業・病院等・砂糖製造業）および'
    '令和7年度から追加された情報通信業・宿泊業を対象に、'
    '業種固有の課題に応じた時間外労働削減・休日確保の取り組みを支援します。'
    '令和8年度は成果目標に「所定外労働時間の削減」が新たに追加されました。'
)

add_new_badge(doc)
add_body(doc, '成果目標に「所定外労働時間の削減」が新規追加。割増賃金率引上げ加算も新設。', bold=True)
doc.add_paragraph()

add_heading(doc, '【対象業種・事業主の要件】', level=3)
add_info_table(doc, [
    ('対象業種',
     '①建設業（資本金3億円以下または常用労働者300人以下）\n'
     '②自動車運転業務（運送業：資本金3億円以下または300人以下）\n'
     '③病院等（医師・看護師等を含む医療機関：資本金5,000万円以下または100人以下）\n'
     '④砂糖製造業（さとうきびを原料とするもの）\n'
     '⑤情報通信業（令和7年度より追加・令和8年度継続）\n'
     '⑥宿泊業（令和7年度より追加・令和8年度継続）'),
    ('共通要件',
     '・労働者災害補償保険の適用事業主\n'
     '・全対象事業場で年5日の年次有給休暇取得に向けた就業規則等を整備済み\n'
     '・36協定を締結・届出済みであること'),
])

add_heading(doc, '【助成対象の主な取り組み（成果目標）】', level=3)
add_bullet(doc, '時間外・休日労働時間数の縮減（月60時間以下・45時間以下等）')
add_bullet(doc, '所定外労働時間の削減【令和8年度新規追加】')
add_bullet(doc, '年次有給休暇の計画的付与制度の新規導入')
add_bullet(doc, '時間単位年次有給休暇制度の新規導入')
add_bullet(doc, '勤務間インターバル制度の新規導入')
add_bullet(doc, '特別休暇（病気・育児・介護等）の新規導入')
add_bullet(doc, '建設業：4週8休以上の確保（いわゆる「週休2日」の実現）')
add_bullet(doc, '運送業：拘束時間・連続運転時間の削減に向けた取り組み')
doc.add_paragraph()

add_heading(doc, '【助成額・助成率】', level=3)
add_info_table(doc, [
    ('助成率', '対象経費の3/4（特定条件では4/5）'),
    ('上限額',
     '成果目標ごとの上限額の合計＋賃金引上げ加算額＋割増賃金率引上げ加算額\n'
     '（取り組みの組み合わせにより最大数百万円規模）'),
    ('割増賃金率\n引上げ加算\n【令和8年度新設】',
     '・割増賃金率を5%以上引き上げた場合：25万円加算\n'
     '・月60時間超の割増賃金率を50%以上に引き上げた場合：100万円加算'),
    ('対象経費例', '就業規則・労使協定の変更、コンサルタント、研修、\nシステム導入（勤怠管理・配車管理等）、人材確保費用 等'),
])

add_heading(doc, '【申請のポイント】', level=3)
add_bullet(doc, '自社が対象業種に該当するかを日本標準産業分類で確認すること')
add_bullet(doc, '令和8年度から「所定外労働時間の削減」が成果目標に追加。時間外だけでなく所定外も削減対象に')
add_bullet(doc, '建設業の「4週8休」は元請けとの連携が重要。受注側単独での達成が困難な場合は元請けを巻き込む')
add_bullet(doc, '運送業は荷主との協力がポイント。取引環境改善コース（コース⑤）との組み合わせも検討する')
add_bullet(doc, '医療機関は医師・看護師の時間外上限規制対応（年960時間・年1,860時間）と合わせて活用を検討する')
add_bullet(doc, '情報通信業・宿泊業は令和7年度からの対象。令和8年度も引き続き申請可能')
add_bullet(doc, '複数の成果目標＋割増賃金率引上げ加算の組み合わせで助成総額の最大化を狙う')
doc.add_paragraph()
doc.add_page_break()

# ==============================
# コース4: 団体推進コース
# ==============================
add_course_divider(doc, 'コース④ 団体推進コース', '1E5C99')

add_heading(doc, '【概要】', level=3)
add_body(doc,
    '事業主団体や共同事業主（10事業主以上で構成）が、傘下の中小企業事業主に対して'
    '労働条件の改善に向けた支援（セミナー開催、専門家派遣、ツール提供等）を実施する場合に'
    'その費用を助成します。個社ではなく「団体」が申請主体になるコースで、助成率100%が最大の特徴です。'
)

add_heading(doc, '【対象（申請主体の要件）】', level=3)
add_info_table(doc, [
    ('申請主体', '事業主団体（中小企業事業主で構成）または共同事業主'),
    ('構成要件',
     '・10事業主以上で構成されていること\n'
     '・1年以上の活動実績があること\n'
     '・傘下企業が中小企業事業主であること'),
    ('規模別上限',
     '・通常：助成上限額500万円\n'
     '・構成事業主が複数都道府県または市区町村にまたがる場合：上限1,000万円'),
])

add_heading(doc, '【助成対象の主な取り組み】', level=3)
add_bullet(doc, '傘下企業向けのセミナー・研修の開催')
add_bullet(doc, '傘下企業への専門家（社労士・コンサルタント等）の派遣')
add_bullet(doc, '傘下企業が活用できる就業規則モデル・ツール類の作成・提供')
add_bullet(doc, '傘下企業の労働条件改善に向けた個別相談・指導')
add_bullet(doc, '業種・地域の実態調査や課題把握のための調査・研究')
doc.add_paragraph()

add_heading(doc, '【助成額・助成率】', level=3)
add_info_table(doc, [
    ('助成率', '対象経費の10/10（100%）※他コースと異なり全額助成'),
    ('支給額の計算',
     '①対象経費の合計額\n'
     '②総事業費から収入額を控除した額\n'
     '③上限額（500万円または1,000万円）\n'
     '上記①〜③のうち最も低い金額が支給額'),
    ('上限額',
     '・500万円（通常）\n'
     '・1,000万円（複数都道府県・市区町村にまたがる場合）'),
    ('対象経費例', '専門家派遣費用、セミナー開催費（会場・講師料等）、\nモデル就業規則作成費、調査費、印刷・資料作成費 等'),
])

add_heading(doc, '【申請のポイント】', level=3)
add_bullet(doc, '助成率が100%という点が最大の特徴。団体として積極的な活用を検討する価値あり')
add_bullet(doc, '「1年以上の活動実績」が必要。設立間もない団体は要件を満たさない場合がある')
add_bullet(doc, '傘下企業が10社以上であることの証明書類（会員名簿等）を事前に整備しておく')
add_bullet(doc, '申請者は団体だが取り組みの恩恵は傘下の各事業主に及ぶ。傘下企業への周知・連携が重要')
add_bullet(doc, '事業計画書に実施内容・スケジュール・期待効果を具体的に記載することが採択のポイント')
add_bullet(doc, '収入（参加費・受講料等）が発生する場合は総事業費から控除されるため、収入の有無を整理しておく')
add_bullet(doc, '申請期限：令和8年11月30日（月）午後5時')
doc.add_paragraph()
doc.add_page_break()

# ==============================
# コース5: 取引環境改善コース
# ==============================
add_course_divider(doc, 'コース⑤ 取引環境改善コース', '153D6B')

add_heading(doc, '【概要】', level=3)
add_body(doc,
    '荷主集団（荷主・倉庫事業者・運送事業者等で構成）が連携して、'
    '荷待ち時間・荷役時間の短縮や取引慣行の見直しを行い、'
    'トラックドライバーをはじめとする自動車運転者の時間外労働削減に取り組む場合に支援します。'
    '2024年4月から運送業にも時間外労働の上限規制（年960時間）が適用されており、'
    '荷主側の協力も不可欠な状況に対応するコースです。令和8年度も引き続き実施されています。'
)

add_heading(doc, '【対象（申請主体・構成員の要件）】', level=3)
add_info_table(doc, [
    ('申請主体', '荷主集団等を代表する「代表事業主」が申請'),
    ('構成員',
     '・荷主（貨物輸送を依頼する事業者）\n'
     '・倉庫事業者（荷主との間で寄託契約を締結している者）\n'
     '・運送事業者（貨物自動車運送事業者または貨物利用運送事業者で、\n　 自動車運転者を雇用する中小企業事業主）'),
    ('構成要件',
     '・構成員の運送事業者の2分の1以上に荷待ち・荷役時間・労働時間の短縮効果が生じること\n'
     '・荷主と運送事業者の双方が構成員に含まれること'),
    ('実施期間', '交付決定日から当該年度の2月14日まで'),
])

add_heading(doc, '【助成対象の主な取り組み】', level=3)
add_bullet(doc, '荷待ち時間削減のための積卸しスケジュールの見直し・調整')
add_bullet(doc, '荷役作業の機械化・効率化（フォークリフト等の導入）')
add_bullet(doc, '附帯作業（手積み・手降ろし等）の削減に向けた取引慣行の見直し')
add_bullet(doc, '取引適正化に向けた荷主・運送事業者間の契約内容の整備（書面化等）')
add_bullet(doc, '運行管理システムや荷物追跡システムの導入による業務効率化')
add_bullet(doc, '関係者向けセミナー・研修の実施（取引適正化に関する理解促進）')
doc.add_paragraph()

add_heading(doc, '【助成額・助成率】', level=3)
add_info_table(doc, [
    ('助成率', '対象経費の3/4'),
    ('上限額', '取り組み内容・構成員数等により異なる\n（詳細は交付申請マニュアルを参照）'),
    ('対象経費例', '機械・設備の導入費、システム開発・導入費、\nコンサルタント費用、研修・セミナー開催費 等'),
])

add_heading(doc, '【申請のポイント】', level=3)
add_bullet(doc, '荷主・運送事業者の双方が参加する「集団」を組成することが前提。単独事業者は申請不可')
add_bullet(doc, '「構成員の運送事業者の2分の1以上」への効果が要件のため、参加事業者の範囲・規模を慎重に設定する')
add_bullet(doc, '業種別課題対応コース（コース③運送業）との違いを確認し、自社状況に適したコースを選ぶ')
add_bullet(doc, '荷主側の積極的な参加・協力が成否を左右する。事前に荷主へ制度の趣旨・メリットを説明し巻き込む')
add_bullet(doc, '荷待ち時間・荷役時間の現状把握（記録・ヒアリング）を申請前に実施しておくと計画の説得力が増す')
add_bullet(doc, '2024年4月の運送業の上限規制適用により、荷主にも「協力義務」がある。その観点からも荷主説得がしやすい')
doc.add_paragraph()

# ==============================
# まとめ・比較表
# ==============================
doc.add_page_break()
add_heading(doc, '■ コース比較一覧表（令和8年度）', level=1)
doc.add_paragraph()

compare_table = doc.add_table(rows=7, cols=5)
compare_table.style = 'Table Grid'
compare_table.alignment = WD_TABLE_ALIGNMENT.LEFT

headers = ['コース', '主な対象', '助成率', '上限額（目安）', '令和8年度のポイント']
col_widths_compare = [Cm(3.5), Cm(4.0), Cm(2.0), Cm(3.5), Cm(4.0)]

for j, (h, w) in enumerate(zip(headers, col_widths_compare)):
    cell = compare_table.rows[0].cells[j]
    cell.width = w
    set_cell_bg(cell, '1F497D')
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run.font.name = '游ゴシック'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '游ゴシック')

data_rows = [
    ['①労働時間短縮・年休促進支援', '全業種の中小企業', '3/4\n（条件付4/5）', '最大150万円\n＋各種加算', '割増賃金率引上げ加算が新設（最大+100万円）'],
    ['②勤務間インターバル導入', '全業種の中小企業', '3/4\n（条件付4/5）', '最大150万円\n＋各種加算', '新規導入の上限を120万円→150万円に引き上げ'],
    ['③業種別課題対応', '建設・運送・医療・\n砂糖・情報・宿泊業', '3/4\n（条件付4/5）', '取組内容により\n異なる', '所定外労働時間削減の成果目標を新設'],
    ['④団体推進', '事業主団体\n（10社以上）', '10/10\n（全額）', '500万円\n～1,000万円', '助成率100%。収入控除に注意'],
    ['⑤取引環境改善', '荷主・倉庫・\n運送の集団', '3/4', '取組内容により\n異なる', '令和8年度も継続実施。運送業の上限規制対応'],
]

for i, row_data in enumerate(data_rows, start=1):
    row = compare_table.rows[i]
    bg = 'EBF3FB' if i % 2 == 0 else 'FFFFFF'
    for j, (cell_text, w) in enumerate(zip(row_data, col_widths_compare)):
        cell = row.cells[j]
        cell.width = w
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        run = p.add_run(cell_text)
        run.font.size = Pt(9.5)
        run.font.name = '游明朝'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')

doc.add_paragraph()

note = doc.add_paragraph()
note.paragraph_format.left_indent = Cm(0.5)
run_note = note.add_run(
    '※ 本資料は厚生労働省公表情報（令和8年度）をもとに作成しています。\n'
    '　 助成額・要件等は変更される場合があります。申請前に必ず最新の公募要領・申請マニュアルをご確認ください。\n'
    '　 申請・相談窓口：各都道府県労働局 雇用環境・均等部（室）または働き方改革推進支援センター'
)
run_note.font.size = Pt(9)
run_note.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
run_note.font.name = '游明朝'
run_note._element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')

output_path = '/home/user/umb/働き方改革推進支援助成金_コース概要まとめ_令和8年度.docx'
doc.save(output_path)
print(f'保存完了: {output_path}')
