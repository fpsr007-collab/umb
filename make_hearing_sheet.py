"""就業規則リスク診断・コンサルヒアリングシート2026 改良版 Word生成スクリプト
修正点：①月60時間超割増率確認 ②賃金デジタル払い ③フリーランス保護法確認 ④個人情報保護規程追加
        ⑤有給比例付与（パート）確認 ⑥解雇予告手当確認
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ページ設定
sec = doc.sections[0]
sec.page_width        = Cm(21)
sec.page_height       = Cm(29.7)
sec.top_margin        = Cm(1.2)
sec.bottom_margin     = Cm(1.2)
sec.left_margin       = Cm(1.8)
sec.right_margin      = Cm(1.8)

# ============================================================
# ヘルパー関数
# ============================================================

def set_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def set_tbl_borders(tbl_obj, color='AAAAAA', sz=4):
    tblPr    = tbl_obj._tbl.tblPr
    tblBdrs  = OxmlElement('w:tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'),   'single')
        el.set(qn('w:sz'),    str(sz))
        el.set(qn('w:color'), color)
        tblBdrs.append(el)
    tblPr.append(tblBdrs)

def ns(p):
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)

def sp(doc, pt=3):
    p = doc.add_paragraph(); ns(p)
    p.paragraph_format.space_after = Pt(pt)

def footer_copyright(doc):
    p = doc.add_paragraph()
    ns(p); p.paragraph_format.space_after = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('© 社会保険労務士法人アンブレラ　　※顧客渡し版では社労士メモを削除')
    r.font.size = Pt(8); r.font.color.rgb = RGBColor(0x88,0x88,0x88)

# --- セクションタイトル ---
def sec_title(doc, title, bg='1F3864'):
    t = doc.add_table(rows=1, cols=1); t.style = 'Table Grid'
    set_tbl_borders(t, '333333', 6)
    c = t.rows[0].cells[0]; set_bg(c, bg)
    p = c.paragraphs[0]; ns(p)
    r = p.add_run(title); r.bold = True
    r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
    sp(doc, 2)

# --- サブセクション見出し ---
def sub(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run(f'◇ {title}'); r.bold = True
    r.font.size = Pt(10); r.font.color.rgb = RGBColor(0x1F,0x38,0x64)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'4'); b.set(qn('w:color'),'AAAAAA')
    pBdr.append(b); pPr.append(pBdr)

# --- 社労士メモ（★黄色） ---
def memo(doc, text):
    t = doc.add_table(rows=1, cols=1); t.style = 'Table Grid'
    c = t.rows[0].cells[0]; set_bg(c, 'FFF9C4')
    tcPr = c._tc.get_or_add_tcPr()
    tcB  = OxmlElement('w:tcBorders')
    for side in ['top','bottom','left','right']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'),'dashed'); el.set(qn('w:sz'),'4'); el.set(qn('w:color'),'E6A800')
        tcB.append(el)
    tcPr.append(tcB)
    p = c.paragraphs[0]; ns(p)
    r1 = p.add_run('★ '); r1.bold = True
    r1.font.size = Pt(9); r1.font.color.rgb = RGBColor(0x7A,0x4F,0x00)
    r2 = p.add_run(text)
    r2.font.size = Pt(9); r2.font.color.rgb = RGBColor(0x7A,0x4F,0x00)
    sp(doc, 2)

# --- 行テキスト ---
def row(doc, text, ind=0.3, size=10, bold=False):
    p = doc.add_paragraph(text); ns(p)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(ind)
    for r in p.runs:
        r.font.size = Pt(size); r.bold = bold

# --- 注釈 ---
def note(doc, text):
    p = doc.add_paragraph(f'※ {text}'); ns(p)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.5)
    for r in p.runs:
        r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x55,0x55,0x55)

# --- 罫線付き記入行 ---
def blank_lines(doc, n=2, label=''):
    for i in range(n):
        p = doc.add_paragraph(); ns(p)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.left_indent = Cm(0.3)
        if i == 0 and label:
            r = p.add_run(label); r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x55,0x55,0x55)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        b = OxmlElement('w:bottom')
        b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'4'); b.set(qn('w:color'),'AAAAAA')
        pBdr.append(b); pPr.append(pBdr)

# --- チェックボックス行 ---
def checks(doc, items, per_line=3, ind=0.3):
    for i in range(0, len(items), per_line):
        p = doc.add_paragraph(); ns(p)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.left_indent = Cm(ind)
        for item in items[i:i+per_line]:
            r = p.add_run(f'□ {item}　　'); r.font.size = Pt(10)

# --- グリッドテーブル ---
def gtable(doc, headers, data_rows, col_w=None, hbg='DDDDDD'):
    nc = len(headers)
    t  = doc.add_table(rows=1+len(data_rows), cols=nc)
    t.style = 'Table Grid'; set_tbl_borders(t, 'AAAAAA', 4)
    hr = t.rows[0]
    for i,h in enumerate(headers):
        c = hr.cells[i]; set_bg(c, hbg)
        p = c.paragraphs[0]; ns(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(9)
    for ri, rd in enumerate(data_rows):
        row_ = t.rows[ri+1]
        for ci, v in enumerate(rd):
            c = row_.cells[ci]; p = c.paragraphs[0]; ns(p)
            r = p.add_run(v); r.font.size = Pt(9.5)
    if col_w:
        for ri in range(len(t.rows)):
            for ci,w in enumerate(col_w):
                t.rows[ri].cells[ci].width = Cm(w)
    sp(doc, 2)
    return t

# --- 社労士診断テーブル（各セクション末尾） ---
def diag(doc):
    t = doc.add_table(rows=2, cols=4); t.style = 'Table Grid'
    set_tbl_borders(t, '555555', 4)
    hdrs = ['社労士診断','リスク度','規程整備','優先対応・提案メモ']
    for i,h in enumerate(hdrs):
        c = t.rows[0].cells[i]; set_bg(c,'444444')
        p = c.paragraphs[0]; ns(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold=True; r.font.size=Pt(9)
        r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
    data2 = ['','□ 高　□ 中　□ 低','□ 必須　□ 推奨　□ 不要','']
    for i,d in enumerate(data2):
        c = t.rows[1].cells[i]; p = c.paragraphs[0]; ns(p)
        r = p.add_run(d); r.font.size = Pt(9.5)
    ws = [2.5,3.5,4.0,7.0]
    for ri in range(2):
        for ci,w in enumerate(ws):
            t.rows[ri].cells[ci].width = Cm(w)
    sp(doc, 5)


# ============================================================
# ドキュメント本文
# ============================================================

# ヘッダー（コピーライト＋タイトル）
p_cr = doc.add_paragraph(); ns(p_cr)
p_cr.paragraph_format.space_after = Pt(1)
p_cr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r_cr = p_cr.add_run('© 社会保険労務士法人アンブレラ　※顧客渡し版では社労士メモを削除')
r_cr.font.size = Pt(8); r_cr.font.color.rgb = RGBColor(0x88,0x88,0x88)

p_ti = doc.add_paragraph(); ns(p_ti)
p_ti.paragraph_format.space_after = Pt(4)
p_ti.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_ti = p_ti.add_run('就業規則リスク診断・コンサルヒアリングシート 2026 改良版')
r_ti.bold = True; r_ti.font.size = Pt(16)
r_ti.font.underline = True

p_su = doc.add_paragraph(); ns(p_su)
p_su.paragraph_format.space_after = Pt(4)
p_su.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_su = p_su.add_run('★印は社労士メモです。顧客へ渡す際は削除してください。')
r_su.font.size = Pt(9); r_su.font.color.rgb = RGBColor(0x44,0x44,0x44)

# 基本情報テーブル
info_t = doc.add_table(rows=2, cols=6)
info_t.style = 'Table Grid'; set_tbl_borders(info_t, '444444', 4)
labels = [['会社名','','','担当社労士','','ヒアリング日'],
          ['面談者','','','役職','','連絡先']]
for ri,row_ in enumerate(labels):
    for ci,lbl in enumerate(row_):
        c = info_t.rows[ri].cells[ci]
        if lbl:
            set_bg(c,'EEEEEE')
            p = c.paragraphs[0]; ns(p)
            r = p.add_run(lbl+'　'); r.bold=True; r.font.size=Pt(9)
            c.add_paragraph()
# 列幅調整（会社名列は広く）
cw = [1.8,3.8,0.2,1.8,2.8,2.6]
for ri in range(2):
    for ci,w in enumerate(cw):
        info_t.rows[ri].cells[ci].width = Cm(w)
sp(doc, 3)

# 社労士メモ（冒頭）
memo(doc,'本シートは「情報収集」ではなく「リスク診断」と「次回提案」につなげるためのシート。'
         '表向きの理由だけで終わらせず、実際の不満・揉めごと・不安を必ず確認する。')

# ── 相談区分・ゴール設定 ──
sec_title(doc,'今回の相談区分・ゴール設定','2E4057')
checks(doc,[
    '新規作成','全面改定','法改正対応のみ','労務トラブル対応',
    '助成金・認定制度対応','採用強化・定着率改善','M&A・事業承継・組織再編対応',
    'その他：＿＿＿＿＿＿'
], per_line=4)
row(doc,'今回のゴール・完成イメージ：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿', bold=True)
row(doc,'社長・担当者が一番不安に感じていること：', bold=True)
blank_lines(doc, 2)
diag(doc)

# ── 1. 就業規則作成・改定の背景 ──
sec_title(doc,'1. 就業規則作成・改定の背景')
memo(doc,'「表向きの理由」と「裏の理由」を分ける。就業規則は、実際のトラブル・不満・経営課題から逆算して作る。')
row(doc,'表向きの理由：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿')
row(doc,'実際の課題・背景：')
blank_lines(doc, 2)
row(doc,'以前の就業規則：□ 無　□ 有（最終改定：＿＿＿年頃）　□ 不明')

sub(doc,'現行就業規則の問題点チェック')
checks(doc,[
    '最新法改正が反映されていない','実態と規程が合っていない','正社員用しかない','パート・契約社員用がない',
    '賃金規程が不明確','懲戒・解雇規定が弱い','休職・復職規定が弱い','育児介護休業規程が古い',
    'ハラスメント規程がない','テレワーク・副業・SNS規定がない','労使協定との整合性がない','周知されていない'
], per_line=4)
diag(doc)

# ── 2. 会社概要・雇用状況 ──
sec_title(doc,'2. 会社概要・雇用状況')
row(doc,'会社設立：＿＿年＿＿月　／　業種・業態：＿＿＿＿＿＿　／　資本金：＿＿＿＿万円　／　本社所在地：＿＿＿＿＿＿＿＿')

sub(doc,'雇用状況（おおよその数字で可）')
memo(doc,'常時使用する労働者10人以上で就業規則の届出義務。300人超は育休取得率公表等の確認が必要。')
gtable(doc,
    ['区分','役員','正社員','契約社員','パート','嘱託','アルバイト','派遣','合計'],
    [['人数','','','','','','','','']],
    col_w=[2.0,1.3,1.5,1.5,1.5,1.3,1.7,1.3,1.4]
)
row(doc,'性別：男性＿＿人　女性＿＿人　　／　年齢層：20代＿＿　30代＿＿　40代＿＿　50代＿＿　60代以上＿＿')
row(doc,'職種：営業＿＿人　事務＿＿人　技術・SE＿＿人　製造・現場＿＿人　その他（＿＿＿）＿＿人')

sub(doc,'勤務地・テレワーク・出向・外国人雇用')
row(doc,'在宅・テレワーク：□ 無　□ 有（□ 全員可　□ 限定：＿＿＿＿＿）　　関連会社・グループ会社出向：□ 無　□ 有（＿＿＿＿＿）')
row(doc,'外国人労働者：□ 無　□ 有（在留資格：＿＿＿＿＿）　就労可能な業務範囲の確認：□ 済　□ 未確認')

# ★修正③ フリーランス・業務委託の確認追加
sub(doc,'業務委託・フリーランス活用【フリーランス保護法 2024年11月施行】')
memo(doc,'フリーランス保護法（2024年11月施行）により、業務委託先への書面明示・報酬支払期限・育児介護等の配慮義務等が生じた。外注・業委を使っている会社は就業規則とセットで確認する。')
row(doc,'業務委託・フリーランスの活用：□ 無　□ 有（人数：＿＿人程度）')
row(doc,'フリーランス保護法への対応：□ 対応済（書面明示・報酬支払期限等）　□ 未対応　□ 要確認')
row(doc,'業務委託契約書・発注書の整備：□ 整備済　□ 未整備')

sub(doc,'障害者雇用')
row(doc,'□ 無　□ 有（身体：＿人　知的：＿人　精神：＿人）　障害者雇用納付金・報告義務：□ 対象　□ 対象外　□ 要確認')
note(doc,'法定雇用率：2024年4月から2.5%、2026年7月から2.7%。対象事業主は2024年4月から40.0人以上、2026年7月から37.5人以上を目安に確認。')
diag(doc)

# ── 3. 採用・入社手続き・労働条件明示 ──
sec_title(doc,'3. 採用・入社手続き・労働条件明示')

sub(doc,'入社時の提出書類')
memo(doc,'身元保証書は期間・極度額を確認。極度額がない身元保証は無効リスク（民法改正2020年4月）。')
checks(doc,[
    'マイナンバー','住民票記載事項証明書','扶養控除等申告書','誓約書（秘密保持・競業避止）',
    '身元保証書（極度額：＿＿＿円）','源泉徴収票','資格証明書','自動車免許証',
    '給与振込口座','その他：＿＿＿＿＿＿＿'
], per_line=3)

sub(doc,'労働条件通知書・雇用契約書')
row(doc,'使用状況：□ 使用している　□ 使用していない　□ 正社員のみ　□ パート・アルバイトにも使用')
row(doc,'2024年4月改正対応：□ 対応済　□ 未対応')
row(doc,'明示事項：□ 就業場所・業務の変更範囲　□ 有期契約の更新上限　□ 無期転換申込機会　□ 無期転換後の労働条件')

sub(doc,'試用期間・有期契約')
row(doc,'試用期間：□ 無　□ 有（＿＿か月）　　有期契約期間：□ 無　□ 有（＿＿か月）')
row(doc,'有期契約の更新上限：□ 無　□ 有（上限回数：＿＿回／＿＿年）　　無期転換ルール（5年超）の対象者：□ 無　□ 有（＿＿人）')

sub(doc,'同一労働同一賃金チェック')
row(doc,'正規・非正規の待遇差の合理的説明：□ できる　□ 要検討　□ 未整備')
row(doc,'課題・コメント：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿')
diag(doc)

# ── 4. 労働時間・休憩・勤怠管理 ──
sec_title(doc,'4. 労働時間・休憩・勤怠管理')

sub(doc,'労働時間制の種類')
memo(doc,'変形労働・フレックスは労使協定や規程の整備が必要。裁量労働は対象業務・手続・同意等の要件を確認。')
checks(doc,[
    '通常の労働時間制','1か月単位変形労働時間制','1年単位変形労働時間制',
    'フレックスタイム制','事業場外みなし労働時間制','裁量労働制（専門型・企画型）'
], per_line=3)

sub(doc,'通常勤務・シフト制')
row(doc,'通常：始業（＿＿時＿＿分）　終業（＿＿時＿＿分）　休憩（＿＿時＿＿分〜＿＿時＿＿分）　合計（＿＿）分')
gtable(doc,
    ['パターン','始業','終業','休憩','備考'],
    [['A','時　分','時　分','　　分',''],
     ['B','時　分','時　分','　　分',''],
     ['C','時　分','時　分','　　分',''],
     ['D','時　分','時　分','　　分','']],
    col_w=[2.0,3.5,3.5,2.5,5.5]
)
row(doc,'シフト表：□ 有　□ 無　　確定時期：＿＿日前　　直前変更ルール：□ 有　□ 無')

sub(doc,'出退勤管理')
row(doc,'管理方法：□ 無　□ 出勤簿　□ タイムカード　□ ICカード　□ 勤怠システム（＿＿＿＿＿）')
row(doc,'欠勤・遅刻・早退の連絡：□ 電話　□ メール　□ 届出書　□ システム　□ その他（＿＿＿＿）')

sub(doc,'時間外・休日・深夜労働')
row(doc,'残業：□ 有　□ 無　　深夜業務（22時〜5時）：□ 有　□ 無　　休日業務：□ 有　□ 無')
row(doc,'36協定：□ 済　□ 未締結　　特別条項：□ 有　□ 無　　月の上限残業時間：＿＿時間')
note(doc,'時間外上限規制：原則月45時間・年360時間。特別条項でも単月100時間未満、年720時間以内等を確認。')

# ★修正① 月60時間超の割増賃金率確認追加
row(doc,'月60時間超の割増賃金率（50%）の管理：□ 対応済（計算・支払い確認済）　□ 未対応　□ 要確認',
    bold=False)
note(doc,'月60時間超の割増率50%は2023年4月から中小企業にも適用。超過分の支払い漏れはトラブル直結。')

row(doc,'残業代・残業時間に関する過去のトラブル：')
blank_lines(doc, 2)
diag(doc)

# ── 5. 労使協定・届出書類チェック ──
sec_title(doc,'5. 労使協定・届出書類チェック')
memo(doc,'就業規則だけでなく、実態に必要な労使協定・届出書類が揃っているか確認する。ここが追加提案につながる。')
checks(doc,[
    '36協定','1か月単位変形労働時間制協定','1年単位変形労働時間制協定','フレックスタイム制協定',
    '時間単位年休協定','賃金控除協定','一斉休憩除外協定','育児介護休業等の適用除外協定',
    'その他：＿＿＿＿＿'
], per_line=4)
row(doc,'届出状況：□ 済　□ 未届出　□ 不明　　／　保管場所：＿＿＿＿＿＿＿＿＿＿＿＿')
diag(doc)

# ── 6. 休日・休暇 ──
sec_title(doc,'6. 休日・休暇')
memo(doc,'年間カレンダーを作成・添付する。週1日以上の法定休日、休日振替・代休ルールを確認。')

sub(doc,'休日')
row(doc,'定休日：□月　□火　□水　□木　□金　□土　□日　□祝日')
row(doc,'夏季：＿＿日　GW：＿＿日　年末年始：＿＿日　年間休日合計：＿＿日　シフト制：週休＿＿日')

sub(doc,'年次有給休暇')
row(doc,'一斉付与制度：□ 有　□ 無　　半日単位：□ 有　□ 無　　時間単位：□ 有　□ 無（最大＿＿日分）')
row(doc,'届出期限：取得予定日の＿＿日以上前　届出方法：□ メール　□ 取得届　□ システム　□ 口頭')
row(doc,'有給休暇管理簿：□ 整備済　□ 未整備')
# ★修正⑤ パート比例付与確認追加
row(doc,'パートタイマーへの比例付与：□ 適切に付与済　□ 付与漏れあり（週4日以下・年216日以下は比例付与）')
note(doc,'週4日以下または年216日以下のパートは比例付与の対象。付与漏れは未払い賃金リスク。')
row(doc,'有給休暇に関する過去のトラブル：')
blank_lines(doc, 2)

sub(doc,'特別休暇')
gtable(doc,
    ['種別','日数','有給／無給','備考'],
    [['慶弔休暇（結婚・忌引）','','',''],
     ['生理休暇','必要日数','無給','（法定）'],
     ['裁判員休暇','','',''],
     ['ボランティア休暇','','',''],
     ['その他（　　　　　）','','','']],
    col_w=[5.0,2.5,3.0,6.5]
)
diag(doc)

# ── 7. 育児・介護休業・柔軟な働き方 ──
sec_title(doc,'7. 育児・介護休業・柔軟な働き方')
memo(doc,'2025年4月施行分に加え、2025年10月施行分（柔軟な働き方措置・個別周知等）も確認。規程改定だけでなく運用資料も必要。')
row(doc,'産後パパ育休の実績：□ 有　□ 無　　育児休業の実績：□ 有　□ 無　　介護休業の実績：□ 有　□ 無　　子の看護等休暇の実績：□ 有　□ 無')
row(doc,'育休取得率の公表（300人超）：□ 対象外　□ 公表済　□ 未対応　　育児・介護休業規程：□ 整備済　□ 未整備　□ 要更新')

sub(doc,'3歳以上小学校就学前の子を養育する従業員への柔軟な働き方措置【2025年10月施行】')
checks(doc,[
    '始業時刻等の変更','テレワーク等','短時間勤務制度','新たな休暇制度',
    '保育施設等の設置運営その他これに準ずる便宜供与','未対応'
], per_line=3)
row(doc,'個別の意向確認：□ 実施済　□ 未対応　　意向聴取の記録：□ あり　□ なし　　説明資料：□ あり　□ なし')

sub(doc,'介護離職防止')
row(doc,'介護休業制度の個別周知：□ 実施済　□ 未対応　　相談体制：□ 有　□ 無　　研修・情報提供：□ 有　□ 無')
diag(doc)

# ── 8. 休職・復職・メンタルヘルス ──
sec_title(doc,'8. 休職・復職・メンタルヘルス')
memo(doc,'休職は就業規則の防御力が最も出る項目。休職命令、診断書、復職判定、試し出勤、自然退職を明確にする。')
row(doc,'休職制度：□ 無　□ 有（傷病：＿＿か月、私傷病以外：＿＿＿＿＿）　対象者：□ 正社員のみ　□ 有期・パート含む　□ 要検討')
row(doc,'復職手続：□ 主治医診断書　□ 会社指定医　□ 産業医面談　□ 試し出勤　□ リハビリ勤務　□ 自然退職規定')
row(doc,'メンタルヘルス・長期療養に関する過去のトラブル：')
blank_lines(doc, 2)
diag(doc)

# ── 9. 賃金・退職金・定年 ──
sec_title(doc,'9. 賃金・退職金・定年')

sub(doc,'正社員・契約社員の賃金')
row(doc,'基本給：月給（＿＿＿＿円〜＿＿＿＿円）　日給（＿＿＿＿円〜＿＿＿＿円）')
gtable(doc,
    ['手当名','金額（円）','支給条件・備考'],
    [['　　　　　手当','',''],
     ['　　　　　手当','',''],
     ['　　　　　手当','',''],
     ['　　　　　手当','',''],
     ['通勤交通費','上限　　　　円','□ 実費　□ 定額']],
    col_w=[4.5,3.5,9.0]
)
row(doc,'歩合給：□ 無　□ 有（計算書を従業員に公開：□ 有　□ 無）　　固定残業代：□ 無　□ 有（＿＿時間分／＿＿円）')

sub(doc,'パート・アルバイトの賃金')
row(doc,'時給：＿＿＿＿円〜＿＿＿＿円　通勤交通費：□ 無　□ 有（上限＿＿＿円）　最低賃金確認：□ 済　□ 未確認')

sub(doc,'共通事項')
row(doc,'賃金締切日・支払日：毎月＿＿日締切、＿＿日払（□ 当月　□ 翌月　□ 翌々月）　銀行休日時：□ 前日払　□ 翌日払')
# ★修正② 賃金デジタル払いを追加
row(doc,'支払方法：□ 銀行振込　□ 現金払　□ デジタル払い（資金移動業者：＿＿＿＿＿）')
note(doc,'賃金のデジタル払いは2023年4月解禁。対応する場合は厚労省指定の資金移動業者を確認。')
row(doc,'税・社保以外の控除：□ 無　□ 有（＿＿＿＿＿＿＿＿）')
row(doc,'昇給：□ 無　□ 有（年＿＿回、＿＿月）　賞与：□ 無　□ 有（＿月・＿月・＿月）')

sub(doc,'退職金・定年')
row(doc,'退職金：□ 無　□ 有（□ 中退共　□ 生命保険　□ 企業型DC　□ その他：＿＿＿）')
row(doc,'定年：□ 有（＿＿歳）　□ 無　継続雇用制度：□ 有（＿＿歳まで）　□ 無　70歳就業確保：□ 対応済　□ 検討中　□ 予定なし')

sub(doc,'退職手続き')
row(doc,'自己都合退職の届出期限：退職する＿＿日（または＿＿か月）以上前に届出')
# ★修正⑥ 解雇予告手当の確認追加
row(doc,'解雇予告・解雇予告手当：□ 手続き把握済（30日前予告 or 予告手当支払い）　□ 未把握')
note(doc,'解雇予告手当（平均賃金×30日以上）の支払い漏れは労基法違反。過去の解雇事例と合わせて確認。')
row(doc,'退職に関する過去のトラブル（入社時・退社時）：')
blank_lines(doc, 2)
diag(doc)

# ── 10. 社会保険・福利厚生 ──
sec_title(doc,'10. 社会保険・福利厚生')
row(doc,'社会保険（厚生年金・健康保険）：□ 全員加入済　□ 未加入者あり（＿＿人）　雇用保険：□ 全員加入済　□ 未加入者あり（＿＿人）')
note(doc,'短時間労働者の社会保険適用拡大は、企業規模・労働時間・賃金・学生除外等をヒアリング時点の最新情報で確認。51人以上は既に加入対象。2026年10月〜全企業に拡大。')
row(doc,'健康診断：□ 無　□ 有（毎年＿月）　ストレスチェック（50人以上）：□ 実施済　□ 未実施　□ 対象外')
row(doc,'研修制度：□ 有　□ 無　　資格取得奨励制度：□ 有　□ 無　　福利厚生規程：□ 有　□ 無')
row(doc,'副業・兼業：□ 禁止　□ 全員可　□ 限定許可（＿＿＿＿＿＿）　労働時間通算の管理方法：□ 整備済　□ 未整備')
diag(doc)

# ── 11. 服務規律・ハラスメント・情報管理・懲戒 ──
sec_title(doc,'11. 服務規律・ハラスメント・情報管理・懲戒')

sub(doc,'ハラスメント対策')
memo(doc,'パワハラ防止措置は全事業主に義務。相談窓口の設置、周知、相談対応フロー、再発防止まで確認。')
row(doc,'相談窓口：□ 有（担当：＿＿＿＿＿）　□ 無　　ハラスメント研修：□ 定期実施　□ 実施済（単発）　□ 未実施')
row(doc,'ハラスメントの過去のトラブル事例：')
blank_lines(doc, 2)

sub(doc,'情報セキュリティ・SNS・競業避止')
row(doc,'PC・スマホの貸与：□ 有（□ 全員　□ 限定）　□ 無　　SNS・インターネット利用規定：□ 有　□ 無（要作成）')
row(doc,'秘密保持・競業避止（退職後）：□ 誓約書あり　□ 規定のみ　□ 未対応　　同業他社への転職制限：□ 無　□ 有（制限範囲：＿＿＿＿＿）')

sub(doc,'その他の服務事項')
row(doc,'社有車の業務使用：□ 有　□ 無　　マイカー通勤：□ 有（駐車場：□有　□無　業務使用：□有　□無）　□ 無')
row(doc,'自転車通勤：□ 有（駐輪場：□有　□無　業務使用：□有　□無）　□ 無　アルコールチェック：□ 有（義務対象）　□ 有（任意）　□ 無')
row(doc,'出張：□ 無　□ 有（地域：＿＿＿＿＿　宿泊：□有　□無）　損害賠償・弁償規定：□ 不要　□ 要整備（事例：＿＿＿＿＿＿）')

sub(doc,'問題社員・懲戒対応')
row(doc,'過去にあった問題：')
checks(doc,[
    '無断欠勤','遅刻・早退','業務命令違反','能力不足','協調性不足',
    'ハラスメント加害','SNS投稿','情報漏えい','横領・不正請求','副業トラブル','メンタル不調'
], per_line=4)
row(doc,'対応履歴：□ 口頭注意　□ 書面注意　□ 始末書　□ 減給　□ 出勤停止　□ 退職勧奨　□ 解雇')
row(doc,'懲戒規定の整備：□ 十分　□ 不十分　□ 未整備　解雇した事例：□ 無　□ 有（時期：＿＿＿　概要：＿＿＿＿＿＿＿＿）')
diag(doc)

# ── 12. 就業規則の整備・届出・周知 ──
sec_title(doc,'12. 就業規則の整備・届出・周知')
memo(doc,'就業規則は届出だけでは不十分。効力発生には「周知」が重要。労働者代表の選出手続も確認する。')
row(doc,'就業規則の届出：□ 済（最終届出：＿＿年）　□ 未届出　□ 届出義務なし（9人以下）')
row(doc,'周知方法：□ 備付け　□ 書面交付　□ 社内イントラ・共有フォルダ　□ 未周知')
row(doc,'別規程（パート・契約社員用）：□ 作成済　□ 不要　□ 要作成　労働者代表の選出・意見書：□ 対応済　□ 未対応')

sub(doc,'必要な付属規程')
# ★修正④ 個人情報保護規程を追加
checks(doc,[
    '賃金規程','育児・介護休業規程','テレワーク規程','ハラスメント防止規程',
    '副業・兼業規程','情報セキュリティ規程','退職金規程','休職・復職規程',
    '個人情報保護規程',          # ← 追加
    '出張旅費規程','マイカー通勤規程','その他：＿＿＿＿＿'
], per_line=4)
note(doc,'個人情報保護規程はマイナンバー管理含め、ほぼ全社で必要。社内の個人情報取扱い手順と合わせて整備する。')
diag(doc)

# ── 13. 最終診断・提案事項（社労士メモ） ──
sec_title(doc,'13. 最終診断・提案事項（社労士メモ）','7A4F00')
memo(doc,'顧客渡し版では削除または提案書へ転記。必須対応・推奨対応・追加提案に分けるとクロージングしやすい。')

sub(doc,'A：必須対応（法令違反・行政対応リスク）')
gtable(doc,
    ['No.','対応事項','理由・リスク','期限'],
    [['1','','',''],['2','','',''],['3','','',''],['4','','','']],
    col_w=[1.2,5.5,7.5,2.8]
)

sub(doc,'B：推奨対応（トラブル予防・実態整備）')
gtable(doc,
    ['No.','対応事項','期待効果','優先度'],
    [['1','','','□高　□中　□低'],
     ['2','','','□高　□中　□低'],
     ['3','','','□高　□中　□低'],
     ['4','','','□高　□中　□低']],
    col_w=[1.2,6.0,7.0,2.8]
)

sub(doc,'C：追加提案（制度設計・顧問契約化）')
checks(doc,[
    '賃金制度整備','評価制度整備','育児介護休業規程改定','ハラスメント研修',
    '管理職研修','助成金診断','労務監査','顧問契約','その他：＿＿＿＿＿'
], per_line=4)

sub(doc,'見積もり・次回導線')
row(doc,'就業規則作成・改定：＿＿＿＿円　付属規程：＿＿＿＿円／1本　労使協定整備：＿＿＿＿円　顧問契約：月額＿＿＿＿円')
row(doc,'次回提案内容：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿', bold=True)
row(doc,'次回面談日：＿＿年＿＿月＿＿日　／　提出物：□ 現行就業規則　□ 賃金台帳　□ 労働条件通知書　□ 36協定　□ 年間カレンダー　□ その他：＿＿＿＿＿')
sp(doc, 5)

# ── 14. 顧客説明用まとめ ──
sec_title(doc,'14. 顧客説明用まとめ（必要に応じて使用）','2E4057')
body = (
    '今回のヒアリングでは、就業規則の条文そのものだけでなく、実際の働き方、労使協定、雇用契約書、'
    '休職・復職、ハラスメント、賃金、育児介護休業、周知状況まで確認します。\n\n'
    '就業規則は「作って終わり」ではなく、会社を守るための運用ルールです。法改正対応だけでなく、'
    'トラブルが起きた時に説明できる状態に整えることが重要です。\n\n'
)
p_b = doc.add_paragraph(body); ns(p_b)
p_b.paragraph_format.space_after = Pt(4)
p_b.paragraph_format.left_indent = Cm(0.3)
for r in p_b.runs: r.font.size = Pt(10)

p_em = doc.add_paragraph('本シートをもとに、優先順位を「必須対応」「推奨対応」「追加提案」に分けてご提案します。')
ns(p_em); p_em.paragraph_format.left_indent = Cm(0.3)
for r in p_em.runs:
    r.bold = True; r.font.size = Pt(10.5)

sp(doc, 6)
footer_copyright(doc)

# ── 保存 ──
out = '/home/user/umb/就業規則コンサルヒアリングシート2026.docx'
doc.save(out)
print(f'saved: {out}')
