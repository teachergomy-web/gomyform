"""
data_manager.py — 所有資料的讀寫與查詢
"""
import json
from pathlib import Path
import sys

BASE_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DATA_FILE = DATA_DIR / "grades.json"
DATA_DIR.mkdir(exist_ok=True)

DEFAULT_BOOKS = [
    {"name": "發音1", "items": ["單字", "口說", "寫作"]},
    {"name": "發音2", "items": ["單字", "口說", "寫作"]},
    {"name": "發音3", "items": ["單字", "口說", "寫作", "翻譯"]},
    {"name": "發音4", "items": ["單字", "口說", "寫作", "翻譯"]},
    {"name": "文法1", "items": ["單字", "口說", "寫作", "翻譯"]},
    {"name": "文法2", "items": ["單字", "口說", "寫作", "翻譯"]},
    {"name": "文法3", "items": ["單字", "口說", "寫作", "翻譯"]},
    {"name": "文法4", "items": ["單字", "口說", "寫作", "翻譯"]},
]

GRADES = ["幼稚園", "一年級", "二年級", "三年級", "四年級", "五年級", "六年級"]

def _default():
    return {
        "books": DEFAULT_BOOKS,
        # { "A班": {"book":"發音1", "students":[{"zh":"","en":"","grade":""},...]} }
        "classes": {},
        # { "A班": {"王小明": {"2025-03": [record,...] } } }
        # record: {"type":"考試本","date":"2025/03/05","range":"L1-L3",
        #          "item":"聽力","score":72,"retake":"","std":90}
        "monthly": {},
        # { "A班": [{"date":"","book":"","scores":{"王小明":{"單字":80}},
        #            "retakes":{"王小明":{"單字":90}},"stds":{"王小明":{"單字":90}}}, ...] }
        "leveltest": {},
        # { "A班": ["發音1", ...] }  已升級過的書
        "upgraded": {},
    }

def load():
    if DATA_FILE.exists():
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        for k, v in _default().items():
            if k not in d:
                d[k] = v
        return d
    return _default()

def save(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ── 套書 ──────────────────────────────────────────────────────────────────────
def book_names(data):
    return [b["name"] for b in data["books"]]

def book_items(data, book_name):
    for b in data["books"]:
        if b["name"] == book_name:
            return list(b["items"])
    return []

def next_book(data, book_name):
    names = book_names(data)
    if book_name in names:
        idx = names.index(book_name)
        if idx + 1 < len(names):
            return names[idx + 1]
    return None

# ── 班級 ──────────────────────────────────────────────────────────────────────
def class_list(data):
    return sorted(data["classes"].keys())

def class_students(data, cls):
    return data["classes"].get(cls, {}).get("students", [])

def student_zh_list(data, cls):
    return [s["zh"] for s in class_students(data, cls)]

def class_book(data, cls):
    return data["classes"].get(cls, {}).get("book", "")

def add_class(data, name, book="", teacher=""):
    if name not in data["classes"]:
        data["classes"][name] = {"book": book, "teacher": teacher, "students": []}
        data["monthly"].setdefault(name, {})
        data["leveltest"].setdefault(name, [])

def class_teacher(data, cls):
    return data["classes"].get(cls, {}).get("teacher", "")

def set_class_teacher(data, cls, teacher):
    if cls in data["classes"]:
        data["classes"][cls]["teacher"] = teacher

def teacher_list(data):
    """回傳老師名單"""
    return data.get("teachers", [])

def add_teacher(data, name):
    if "teachers" not in data:
        data["teachers"] = []
    if name and name not in data["teachers"]:
        data["teachers"].append(name)

def remove_teacher(data, name):
    if "teachers" in data:
        data["teachers"] = [t for t in data["teachers"] if t != name]
    # 清除使用這個老師的班級設定
    for cls_data in data.get("classes", {}).values():
        if cls_data.get("teacher") == name:
            cls_data["teacher"] = ""

def rename_class(data, old, new):
    if old in data["classes"] and new not in data["classes"]:
        data["classes"][new] = data["classes"].pop(old)
        data["monthly"][new] = data["monthly"].pop(old, {})
        data["leveltest"][new] = data["leveltest"].pop(old, [])
        data["upgraded"][new] = data["upgraded"].pop(old, [])

def delete_class(data, name):
    data["classes"].pop(name, None)
    # 保留月成績和升級考資料（歷史）

def add_student(data, cls, zh, en, grade):
    students = data["classes"][cls]["students"]
    if not any(s["zh"] == zh for s in students):
        students.append({"zh": zh, "en": en, "grade": grade})

def update_student(data, cls, old_zh, zh, en, grade):
    for s in data["classes"][cls]["students"]:
        if s["zh"] == old_zh:
            # 同步月成績 key
            if old_zh != zh:
                for ym_data in data["monthly"].get(cls, {}).values():
                    pass  # key 在外層
                if old_zh in data["monthly"].get(cls, {}):
                    data["monthly"][cls][zh] = data["monthly"][cls].pop(old_zh)
                # 同步升級考
                for rec in data["leveltest"].get(cls, []):
                    for d_map in [rec.get("scores",{}), rec.get("retakes",{}), rec.get("stds",{})]:
                        if old_zh in d_map:
                            d_map[zh] = d_map.pop(old_zh)
            s["zh"] = zh
            s["en"] = en
            s["grade"] = grade
            break

def remove_student(data, cls, zh):
    """離班：從名單移除，保留歷史成績"""
    data["classes"][cls]["students"] = [
        s for s in data["classes"][cls]["students"] if s["zh"] != zh
    ]

def transfer_student(data, src_cls, dst_cls, zh):
    """換班：移動學生資料與成績"""
    students = data["classes"][src_cls]["students"]
    stu = next((s for s in students if s["zh"] == zh), None)
    if not stu:
        return
    # 從原班移除
    data["classes"][src_cls]["students"] = [s for s in students if s["zh"] != zh]
    # 加入目標班
    if not any(s["zh"] == zh for s in data["classes"][dst_cls]["students"]):
        data["classes"][dst_cls]["students"].append(stu)
    # 移動月成績
    src_monthly = data["monthly"].get(src_cls, {})
    if zh in src_monthly:
        data["monthly"].setdefault(dst_cls, {})[zh] = src_monthly.pop(zh)

# ── 月成績 ────────────────────────────────────────────────────────────────────
def monthly_records(data, cls, zh, ym):
    return data["monthly"].get(cls, {}).get(zh, {}).get(ym, [])

def all_monthly_records(data, cls, ym):
    """回傳該班該月所有學生的成績 {zh: [records]}"""
    result = {}
    for zh in student_zh_list(data, cls):
        recs = monthly_records(data, cls, zh, ym)
        result[zh] = recs
    return result

def set_monthly_records(data, cls, zh, ym, records):
    data["monthly"].setdefault(cls, {}).setdefault(zh, {})[ym] = records

def missing_retake_count(data, cls, ym):
    """計算該班該月缺考（原始分空白）或未補考（有原始分未達標且補考空白）的筆數"""
    count = 0
    for zh in student_zh_list(data, cls):
        for rec in monthly_records(data, cls, zh, ym):
            score = rec.get("score", "")
            retake = rec.get("retake", "")
            std = rec.get("std", 90)
            if score == "":
                count += 1
            elif retake == "" and isinstance(score, (int, float)) and score < std:
                count += 1
    return count

# ── 升級考 ────────────────────────────────────────────────────────────────────
def leveltest_list(data, cls):
    # 優先從 classes[cls]["leveltests"] 讀（新結構）
    cls_lts = data.get("classes", {}).get(cls, {}).get("leveltests", [])
    if cls_lts:
        return cls_lts
    return data["leveltest"].get(cls, [])

def latest_leveltest(data, cls):
    recs = leveltest_list(data, cls)
    return recs[-1] if recs else None

def add_leveltest(data, cls, date, book, scores, retakes, stds):
    # 同時存到 classes[cls]["leveltests"]（供總覽讀取）
    if cls not in data.get("classes", {}):
        return
    lts = data["classes"][cls].setdefault("leveltests", [])
    # 如果同日期已存在就更新，否則新增
    existing = next((i for i, x in enumerate(lts) if x.get("date") == date), None)
    entry = {"date": date, "book": book, "scores": scores, "retakes": retakes, "stds": stds}
    if existing is not None:
        lts[existing] = entry
    else:
        lts.append(entry)
    # 也維持舊 leveltest 結構的相容性
    data["leveltest"].setdefault(cls, [])
    old_idx = next((i for i, x in enumerate(data["leveltest"][cls]) if x.get("date") == date), None)
    if old_idx is not None:
        data["leveltest"][cls][old_idx] = entry
    else:
        data["leveltest"][cls].append(entry)

def should_prompt_upgrade(data, cls):
    lt = latest_leveltest(data, cls)
    if not lt:
        return False
    book = lt["book"]
    if book in data["upgraded"].get(cls, []):
        return False
    return next_book(data, book) is not None

def do_upgrade(data, cls):
    lt = latest_leveltest(data, cls)
    if not lt:
        return
    nb = next_book(data, lt["book"])
    if nb:
        data["classes"][cls]["book"] = nb
        data["upgraded"].setdefault(cls, [])
        if lt["book"] not in data["upgraded"][cls]:
            data["upgraded"][cls].append(lt["book"])

def skip_upgrade(data, cls):
    lt = latest_leveltest(data, cls)
    if lt:
        data["upgraded"].setdefault(cls, [])
        if lt["book"] not in data["upgraded"][cls]:
            data["upgraded"][cls].append(lt["book"])

