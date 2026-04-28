"""
docx_exporter.py
"""
from pathlib import Path
from docxtpl import DocxTemplate, RichText  # 💡 新增 RichText 支援動態修改字體
from docx.shared import Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

MONTH_ZH = {1:'一',2:'二',3:'三',4:'四',5:'五',6:'六',
            7:'七',8:'八',9:'九',10:'十',11:'十一',12:'十二'}

def _find_template(base_dir, *names):
    for name in names:
        p = Path(base_dir) / name
        if p.exists():
            return str(p)
    return None

def _get_base():
    import sys
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent

def _get_day(d):
    if not d: return ''
    try: return str(int(str(d).split('/')[-1]))
    except: return str(d)

def _fmt(v):
    """格式化分數：去掉多餘小數（100.0 -> 100，91.5 -> 91.5）"""
    if v == '' or v is None: return ''
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else str(f)
    except:
        return str(v)

def _book_en(book_name):
    """套書名稱轉英文格式：文法3 → Grammar3，發音4 → Phonics4"""
    import re
    name = str(book_name)
    name = name.replace('文法', 'Grammar').replace('發音', 'Phonics')
    return name

def _auto_size(text):
    """💡 自動縮放文字大小：如果文字長度超過 3 個半形字元，將字體縮小為 10pt"""
    text = str(text)
    if not text: return ''
    
    # 計算字串長度（全形中文字算 2，半形英文數字算 1）
    length = sum(2 if ord(c) > 127 else 1 for c in text)
    
    if length > 3:
        # docxtpl 的 size 是以 half-points (半磅) 為單位，20 代表 10pt
        return RichText(text, size=20)
    return text

def export_monthly_batch(folder, template_path, cls_name, book_name,
                          teacher, year_month, students_data,
                          comments=None, goals=None, sig_b64=None):
    """
    comments: dict {zh: 評語文字}
    goals:    list [目標1, 目標2, ...]
    sig_b64:  base64 簽名圖字串
    """
    year, month = year_month.split('-')
    month_zh = MONTH_ZH.get(int(month), month)
    filepath = str(Path(folder) / f'兒美{cls_name}{int(month)}月成績單.docx')

    tmpl = template_path
    if not tmpl or not Path(tmpl).exists():
        base = _get_base()
        tmpl = _find_template(base, '成績單模板.docx', '兒美成績單模板.docx', '兒美A班成績單.docx')
    if not tmpl:
        return ''

    tpl = DocxTemplate(tmpl)
    context_students = []
    
    for item in students_data:
        stu = item['stu']
        records = item['records']
        d = {'zh': stu.get('zh',''), 'en': stu.get('en','')}

        exam  = [r for r in records if r.get('type')=='考試本']
        speak = [r for r in records if r.get('type')=='口說']
        vocab = [r for r in records if r.get('type')=='單字']

        for i in range(7):
            r = exam[i]  if i < len(exam)  else {}
            d[f'ex_d_{i}']    = _get_day(r.get('date',''))
            d[f'ex_rng_{i}']  = _auto_size(r.get('range','') or '')   # 💡 套用自動縮放字體
            d[f'ex_item_{i}'] = _auto_size(r.get('item','')  or '')   # 💡 套用自動縮放字體
            d[f'ex_sc_{i}']   = _fmt(r.get('score',''))
            d[f'ex_rt_{i}']   = _fmt(r.get('retake',''))

            r = speak[i] if i < len(speak) else {}
            d[f'sp_d_{i}']    = _get_day(r.get('date',''))
            d[f'sp_rng_{i}']  = _auto_size(r.get('range','') or '')   
            d[f'sp_item_{i}'] = _auto_size(r.get('item','')  or '')   # 💡 新增這行：口說項目
            d[f'sp_sc_{i}']   = _fmt(r.get('score',''))

            r = vocab[i] if i < len(vocab) else {}
            d[f'voc_d_{i}']    = _get_day(r.get('date',''))
            d[f'voc_rng_{i}']  = _auto_size(r.get('range','') or '')   
            d[f'voc_item_{i}'] = _auto_size(r.get('item','')  or '')   # 💡 新增這行：單字項目
            d[f'voc_sc_{i}']   = _fmt(r.get('score',''))
            d[f'voc_rt_{i}']   = _fmt(r.get('retake',''))

        # 帶入該學生的評語
        zh = item['stu'].get('zh', '')
        d['comment'] = (comments or {}).get(zh, '')

        context_students.append(d)

    # 教學目標
    goals_ctx = goals or []
    goal_vars = {f'goal_{i+1}': (goals_ctx[i] if i < len(goals_ctx) else '') for i in range(4)}

    # 簽名圖：base64 → InlineImage
    sig_image = None
    if sig_b64:
        try:
            import base64, io
            from docxtpl import InlineImage
            from docx.shared import Cm as _Cm
            img_data = base64.b64decode(sig_b64.split(',')[-1])
            sig_image = InlineImage(tpl, io.BytesIO(img_data), width=_Cm(2), height=_Cm(0.4))
        except Exception:
            pass

    context = {
        'month_zh': month_zh, 'cls_name': cls_name, 'book_name': book_name,
        'teacher': teacher, 'students': context_students,
        'goal_1': goal_vars['goal_1'], 'goal_2': goal_vars['goal_2'],
        'goal_3': goal_vars['goal_3'], 'goal_4': goal_vars['goal_4'],
        'sig_image': sig_image,
    }

    tpl.render(context)

    for section in tpl.sections:
        section.top_margin = Cm(1.27)
        section.bottom_margin = Cm(1.27)
        section.left_margin = Cm(1.27)
        section.right_margin = Cm(1.27)

    for table in tpl.tables:
        tbl_props = table._element.xpath('w:tblPr')[0]
        tbl_w = tbl_props.xpath('w:tblW')
        if not tbl_w:
            tbl_w = OxmlElement('w:tblW')
            tbl_props.append(tbl_w)
        else:
            tbl_w = tbl_w[0]
        tbl_w.set(qn('w:type'), 'auto')
        tbl_w.set(qn('w:w'), '0')

    # 檔案衝突檢查
    import os as _os
    if _os.path.exists(filepath):
        try:
            with open(filepath, 'a'): pass
        except PermissionError:
            raise PermissionError("同名檔案已開啟中，請先關閉後再重新輸出：\n" + filepath)
    tpl.save(filepath)
    return filepath


# ══════════════════════════════════════════════════════════════════════════════
# 升級考成績單（用 升級考模板.docx）
# ══════════════════════════════════════════════════════════════════════════════
def export_leveltest(filepath_or_folder, cls_name, book_name, date_str,
                     students, records, template_path=None,
                     lt_comments=None, sig_b64=None):
    """
    用 docxtpl 渲染升級考模板，自動產生檔名。
    檔名格式：A班升級考成績單(GM3).docx
    欄位邏輯：
      lvl_N_score  = 原始成績（過關成績欄）
      lvl_N_retake = 補考成績，且補考 >= 90（補考過關欄）
      lvl_N_fail   = 補考成績，且補考 < 90（不過關欄）
    """
    from pathlib import Path
    # 產生檔名（書名縮寫：文法3→GM3，發音4→PH4）
    book_abbr = book_name.replace("文法", "GM").replace("發音", "PH")
    filename  = f"{cls_name}升級考成績單({book_abbr}).docx"
    # 模板用英文書名
    book_name_en = _book_en(book_name)
    # 相容舊呼叫（傳完整路徑）和新呼叫（傳資料夾）
    p = Path(filepath_or_folder)
    filepath = str(p / filename) if p.is_dir() else filepath_or_folder

    # 固定優先用程式資料夾的模板（確保圖片完整）
    base = _get_base()
    tmpl_path = None
    for name in ['升級考模板.docx', '升級考成績單模板.docx']:
        p = base / name
        if p.exists():
            tmpl_path = str(p)
            break
    # fallback：用傳入的路徑
    if not tmpl_path and template_path and Path(template_path).exists():
        tmpl_path = template_path

    if not tmpl_path:
        print('找不到升級考模板，請確認資料夾內有「升級考模板.docx」')
        return

    tpl = DocxTemplate(tmpl_path)

    context_students = []
    for rec in records:
        zh  = rec['zh']
        stu = next((s for s in students if s['zh'] == zh), {'zh': zh, 'en': ''})
        d   = {'zh': zh, 'en': stu.get('en', '')}

        # 從 lt_comments 取評語、勾選、評鑑結果
        lt_rec = (lt_comments or {}).get(zh, {})
        if isinstance(lt_rec, dict):
            checks = lt_rec.get('checks', rec.get('checks', []))
            result = lt_rec.get('result', rec.get('result', ''))
            d['comment'] = lt_rec.get('comment', rec.get('comment', ''))
        else:
            checks = rec.get('checks', [])
            result = rec.get('result', '')
            d['comment'] = str(lt_rec) if lt_rec else rec.get('comment', '')

        d['chk_neat']      = '■' if 'neat'      in checks else '□'
        d['chk_pronounce'] = '■' if 'pronounce' in checks else '□'
        d['chk_vocab']     = '■' if 'vocab'     in checks else '□'
        d['chk_attitude']  = '■' if 'attitude'  in checks else '□'
        d['chk_express']   = '■' if 'express'   in checks else '□'
        d['chk_pass'] = '■' if result == 'pass' else '□'
        d['chk_fail'] = '■' if result == 'fail' else '□' 

        items = list(rec['items'].keys())
        for i in range(4):
            if i < len(items):
                it     = items[i]
                item_d = rec['items'][it]
                score  = item_d.get('score',  '')
                retake = item_d.get('retake', '')

                # 格式化（去掉 .0）
                score_fmt  = _fmt(score)
                retake_fmt = _fmt(retake)

                # 補考過關欄：補考成績 >= 90（數字）或 A系列
                # 不過關欄：補考成績 < 90
                retake_pass = ''
                retake_fail = ''
                if retake_fmt != '':
                    try:
                        if float(retake) >= 90:
                            retake_pass = retake_fmt
                        else:
                            retake_fail = retake_fmt
                    except:
                        # 非數字成績（不應出現，但保險起見）
                        retake_pass = retake_fmt

                d[f'lvl_{i}_score']  = score_fmt
                d[f'lvl_{i}_retake'] = retake_pass  # 補考過關欄
                d[f'lvl_{i}_fail']   = retake_fail  # 不過關欄
            else:
                d[f'lvl_{i}_score']  = ''
                d[f'lvl_{i}_retake'] = ''
                d[f'lvl_{i}_fail']   = ''

        # 💡 已移除會導致報錯的贅字程式碼 (name 'item' is not defined)

        context_students.append(d)

    # 簽名圖：base64 → InlineImage
    sig_image = None
    if sig_b64:
        try:
            import base64, io
            from docxtpl import InlineImage
            from docx.shared import Cm as _Cm
            img_data = base64.b64decode(sig_b64.split(',')[-1])
            sig_image = InlineImage(tpl, io.BytesIO(img_data), width=_Cm(2.5), height=_Cm(1))
        except Exception:
            pass

    context = {
        'book_name': book_name_en,
        'students':  context_students,
        'sig_image': sig_image,
    }

    tpl.render(context)

    # 檔案衝突檢查
    import os
    if os.path.exists(filepath):
        try:
            with open(filepath, 'a'):
                pass
        except PermissionError:
            raise PermissionError(f"檔案已開啟，請先關閉後再試：\n{filepath}")

    tpl.save(filepath)
    return filepath    # 檔案衝突檢查
    import os as _os
    if _os.path.exists(filepath):
        try:
            with open(filepath, 'a'): pass
        except PermissionError:
            raise PermissionError("同名檔案已開啟中，請先關閉後再重新輸出：\n" + filepath)
