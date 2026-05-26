from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ページ余白設定
section = doc.sections[0]
section.top_margin = Cm(2)
section.bottom_margin = Cm(2)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def heading1(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # 背景色を段落シェーディングで表現（表の1セル行として作成）
    # 代わりに見出しスタイルを使用
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    # 段落に背景色を付ける（w:pPr/w:shd）
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '2E74B5')
    pPr.append(shd)
    return p

def heading2(text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    # 下線ボーダー
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '2E74B5')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p

def bullet(text, bold_part=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    if bold_part and text.startswith(bold_part):
        run = p.add_run(bold_part)
        run.bold = True
        run.font.size = Pt(10.5)
        p.add_run(text[len(bold_part):]).font.size = Pt(10.5)
    else:
        run = p.add_run(text)
        run.font.size = Pt(10.5)
    return p

def normal(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(2)
    for run in p.runs:
        run.font.size = Pt(10.5)
    return p

# ===== タイトル =====
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
tr = title.add_run('介護離職防止支援コース（両立支援等助成金）')
tr.bold = True
tr.font.size = Pt(16)
tr.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
title.paragraph_format.space_after = Pt(4)

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
sr = subtitle.add_run('概要・支給額・申請のポイント')
sr.font.size = Pt(11)
sr.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
subtitle.paragraph_format.space_after = Pt(12)

# ===== 1. 概要 =====
heading1('１．概要')

heading2('目的・対象')
bullet('仕事と介護の両立を支援し、介護を理由とした離職を防止するため、取組を実施した中小企業事業主に助成金を支給する')
bullet('対象：中小企業事業主（資本金額または常時雇用労働者数で判定）')
bullet('適用単位：事業所単位ではなく事業主（法人・個人）単位で支給')
bullet('根拠法令：雇用保険法第62条第1項第6号・雇保則第115条・116条')

heading2('助成金の種類（6種）')

table = doc.add_table(rows=7, cols=2)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

headers = ['種別', '概要']
for i, h in enumerate(headers):
    cell = table.cell(0, i)
    cell.text = h
    cell.paragraphs[0].runs[0].bold = True
    cell.paragraphs[0].runs[0].font.size = Pt(10.5)
    set_cell_bg(cell, '2E74B5')
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

rows_data = [
    ('① 介護休業', '両立支援プランに基づき介護休業を取得し、職場復帰した場合'),
    ('② 介護両立支援制度', '時差出勤・短時間勤務・在宅勤務・フレックス・介護サービス費用補助などを利用させた場合'),
    ('③ 業務代替支援', '休業・短時間勤務中に代替要員を新規雇用、または手当を支給した場合'),
    ('④ 介護休暇制度有給化支援', '有給の介護休暇制度を新たに導入し利用させた場合'),
    ('⑤ 有期雇用労働者加算', '①または②の対象者が有期雇用労働者だった場合（加算）'),
    ('⑥ 環境整備加算', '研修・相談体制・事例提供・方針周知の4つを全て実施した場合（加算）'),
]

for i, (k, v) in enumerate(rows_data):
    row = table.rows[i+1]
    row.cells[0].text = k
    row.cells[1].text = v
    for cell in row.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.size = Pt(10.5)
    if i % 2 == 0:
        set_cell_bg(row.cells[0], 'DEEAF1')
        set_cell_bg(row.cells[1], 'DEEAF1')

doc.add_paragraph()

# ===== 2. 支給額 =====
heading1('２．支給額')

heading2('① 介護休業')
bullet('連続5日以上14日以下の休業：40万円')
bullet('連続15日以上の休業：60万円')
bullet('上限：1事業主あたり5人まで（同一労働者・同一対象家族につき1回限り）')

heading2('② 介護両立支援制度')

tbl2 = doc.add_table(rows=3, cols=3)
tbl2.style = 'Table Grid'
h2_headers = ['導入制度数', '利用日数20日以上', '利用日数60日以上']
for i, h in enumerate(h2_headers):
    cell = tbl2.cell(0, i)
    cell.text = h
    cell.paragraphs[0].runs[0].bold = True
    cell.paragraphs[0].runs[0].font.size = Pt(10.5)
    set_cell_bg(cell, '2E74B5')
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

tbl2_data = [
    ('1つ導入・1つ利用', '20万円', '30万円'),
    ('2つ以上導入・1つ利用', '25万円', '40万円'),
]
for i, row_data in enumerate(tbl2_data):
    row = tbl2.rows[i+1]
    for j, val in enumerate(row_data):
        row.cells[j].text = val
        for run in row.cells[j].paragraphs[0].runs:
            run.font.size = Pt(10.5)
    if i % 2 == 0:
        for j in range(3):
            set_cell_bg(row.cells[j], 'DEEAF1')

doc.add_paragraph().paragraph_format.space_after = Pt(2)
bullet('上限：1事業主あたり5人まで（同一制度・同一対象家族につき1回限り、同一労働者は最大2回）')

heading2('③ 業務代替支援')

tbl3 = doc.add_table(rows=4, cols=3)
tbl3.style = 'Table Grid'
h3_headers = ['区分', '休業15日未満', '休業15日以上']
for i, h in enumerate(h3_headers):
    cell = tbl3.cell(0, i)
    cell.text = h
    cell.paragraphs[0].runs[0].bold = True
    cell.paragraphs[0].runs[0].font.size = Pt(10.5)
    set_cell_bg(cell, '2E74B5')
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

tbl3_data = [
    ('新規雇用', '20万円', '30万円'),
    ('手当支給等（介護休業）', '5万円', '10万円'),
    ('手当支給等（短時間勤務）', '3万円', '―'),
]
for i, row_data in enumerate(tbl3_data):
    row = tbl3.rows[i+1]
    for j, val in enumerate(row_data):
        row.cells[j].text = val
        for run in row.cells[j].paragraphs[0].runs:
            run.font.size = Pt(10.5)
    if i % 2 == 0:
        for j in range(3):
            set_cell_bg(row.cells[j], 'DEEAF1')

doc.add_paragraph().paragraph_format.space_after = Pt(2)
bullet('上限：1事業主あたり5人まで')

heading2('④ 介護休暇制度有給化支援')
bullet('時間単位取得可の有給介護休暇制度を導入：30万円（1事業主1回限り）')
bullet('10日以上の有給休暇とする場合：50万円（1事業主1回限り）')

heading2('⑤ 有期雇用労働者加算')
bullet('①または②に加算：10万円（1事業主いずれか1回限り）')

heading2('⑥ 環境整備加算')
bullet('①〜③に加算：10万円（1事業主1回限り）')

# ===== 3. 申請のポイント =====
heading1('３．申請のポイント')

heading2('申請先・提出期限')

tbl4 = doc.add_table(rows=5, cols=2)
tbl4.style = 'Table Grid'
h4_headers = ['種別', '申請期限']
for i, h in enumerate(h4_headers):
    cell = tbl4.cell(0, i)
    cell.text = h
    cell.paragraphs[0].runs[0].bold = True
    cell.paragraphs[0].runs[0].font.size = Pt(10.5)
    set_cell_bg(cell, '2E74B5')
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

tbl4_data = [
    ('申請先', '本社等の所在地を管轄する都道府県労働局長'),
    ('① 介護休業', '休業終了日の翌日から3か月経過後の翌日から2か月以内'),
    ('② 介護両立支援制度', '利用実績20日/60日到達後1か月経過の翌日から2か月以内'),
    ('③ 業務代替支援（新規雇用）', '休業終了日の翌日から2か月以内'),
]
for i, row_data in enumerate(tbl4_data):
    row = tbl4.rows[i+1]
    row.cells[0].text = row_data[0]
    row.cells[1].text = row_data[1]
    for cell in row.cells:
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(10.5)
    if i % 2 == 0:
        for j in range(2):
            set_cell_bg(row.cells[j], 'DEEAF1')

doc.add_paragraph().paragraph_format.space_after = Pt(2)

heading2('申請前に必要な準備')
bullet('「仕事と介護の両立支援プラン（面談シート兼用）」の作成（休業・制度利用開始前が原則）')
bullet('上司または人事担当者とのプラン策定面談を1回以上実施・記録')
bullet('就業規則・労働協約に介護休業関係制度および原職等復帰の規定を整備')
bullet('休業終了後のフォロー面談の実施・記録（①介護休業の場合）')
bullet('環境整備加算を申請する場合は、研修・相談体制・事例提供・方針周知の4つすべてを実施')

heading2('主な必要書類（①介護休業の場合）')
bullet('就業規則・労働協約（介護休業制度・原職復帰規定の記載箇所）')
bullet('両立支援プランの方針周知が確認できる書類（通達・社内報等）')
bullet('「仕事と介護の両立支援プラン（面談シート兼用）」')
bullet('対象家族の要介護状態を証明する書類（介護保険被保険者証等）')
bullet('労働条件通知書・雇用契約書')
bullet('介護休業申出書（期間変更がある場合は変更申出書も）')
bullet('休業期間中の出勤簿・タイムカード・賃金台帳')
bullet('職場復帰後3か月分の出勤簿・賃金台帳')

heading2('その他の注意点')
bullet('書類省略制度あり：過去に申請済みで内容変更がない書類は確認書（様式第6号）提出で省略可（電子申請除く）')
bullet('支給申請日前1年以内に育児・介護休業法違反があると不支給')
bullet('雇用形態・給与形態の不合理な変更を行った場合は不支給')
bullet('職場復帰後3か月間で就業予定日数の5割未満しか就業しなかった場合は不支給')
bullet('同一事業主に対する支給は各種別ごとに上限あり（原則5人まで、加算は1回限り）')

# 注釈
doc.add_paragraph()
note = doc.add_paragraph()
nr = note.add_run('※ 本資料は令和8年4月8日改正版の支給要領に基づき作成しています。最新情報は管轄労働局または厚生労働省ウェブサイトでご確認ください。')
nr.font.size = Pt(9)
nr.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
nr.italic = True

# 保存
output_path = '/home/user/umb/介護離職防止支援コース_概要まとめ.docx'
doc.save(output_path)
print(f'Saved: {output_path}')
