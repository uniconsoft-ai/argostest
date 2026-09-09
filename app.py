"""
========================================================================================
🏛️ ARGOS.UZ - WEB VAKANSIYALAR TAHLILCHISI VA WORD KO'RUVCHI (WEB APP)
========================================================================================
Platforma: vacancy.argos.uz va hrm.argos.uz
Imkoniyatlar:
  1. Web brauzer orqali qulay boshqaruv va filtrlash
  2. Server-Sent Events (SSE) orqali jonli progress va topilgan vakansiyalar oqimi
  3. Word (.docx) hujjatini to'g'ridan-to'g'ri web sahifada ko'rish (In-Browser Word Viewer)
  4. docx-preview.js (sahifali Word ko'rinishi) va Mammoth HTML (o'qish ko'rinishi)
  5. Avval yaratilgan Word hujjatlari tarixi (History)
========================================================================================
"""
import os
import sys
import time
import re
import json
import glob
import urllib.parse
import threading
from datetime import datetime
from typing import Dict, Any, List, Optional, Tuple, Callable
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from requests.adapters import HTTPAdapter

import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import mammoth
from flask import Flask, render_template, request, Response, jsonify, send_from_directory, send_file

# Windows stdout kodirovkasi
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# --------------------------------------------------------------------------------------
# 1. KONFIGURATSIYA VA DOIMIY QIYMATLAR
# --------------------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Cloud (Google Cloud Run / Firebase / Serverless) muhitida /tmp/exports, lokalda exports papkasi
EXPORTS_DIR = os.environ.get("EXPORTS_DIR")
if not EXPORTS_DIR:
    if os.environ.get("K_SERVICE") or os.environ.get("GAE_INSTANCE") or os.environ.get("FUNCTION_TARGET"):
        EXPORTS_DIR = "/tmp/exports"
    else:
        EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
os.makedirs(EXPORTS_DIR, exist_ok=True)

DEFAULT_LIST_URL = "https://vacancy.argos.uz/hrm-vacancy-list"
BASE_VACANCY_URL = "https://vacancy.argos.uz"
BASE_HRM_URL = "https://hrm.argos.uz"

API_LIST_ENDPOINT = f"{BASE_VACANCY_URL}/api/Vacancy/HRMVacancy/GetHRMVacancyList"
API_DETAIL_ENDPOINT = f"{BASE_VACANCY_URL}/api/Vacancy/HRMVacancy/GetHRMVacancyDetail"
API_TEST_TYPES_ENDPOINT = f"{BASE_HRM_URL}/api/SelectItem/GetSelectItems/testTypes"
API_DISTRICTS_ENDPOINT = f"{BASE_HRM_URL}/api/SelectItem/GetSelectItems/districts"
API_INSTITUTIONS_ENDPOINT = f"{BASE_HRM_URL}/api/SelectItem/GetHeaderResult/institutions"

TEST_TYPES_MAP = {
    4: "Davlat fuqarolik xizmatchisi (Boshqaruv xodimi)",
    5: "Davlat fuqarolik xizmatchisi (Mutaxassis)",
    6: "Davlat fuqarolik xizmati",
    1225: "Hamshiralik ishi",
    1226: "Shifokorlar uchun",
    0: "Test talab etilmaydi",
}

TEST_TYPE_ALIASES = {
    "barchasi": "all",
    "all": "all",
    "4": 4,
    "boshqaruv": 4,
    "boshqaruv xodimi": 4,
    "davlat fuqarolik xizmatchisi (boshqaruv xodimi)": 4,
    "5": 5,
    "mutaxassis": 5,
    "mutaxassislik": 5,
    "davlat fuqarolik xizmatchisi (mutaxassis)": 5,
    "dfx": "dfx_all",
    "dfx_all": "dfx_all",
    "davlat fuqarolik xizmati": "dfx_all",
    "davlat fuqarolik xizmati (barchasi)": "dfx_all",
    "1225": 1225,
    "hamshira": 1225,
    "hamshiralik": 1225,
    "hamshiralik ishi": 1225,
    "1226": 1226,
    "shifokor": 1226,
    "shifokorlik": 1226,
    "shifokorlar uchun": 1226,
    "0": 0,
    "talab etilmaydi": 0,
    "test talab etilmaydi": 0,
    "mavjud emas": 0,
}

EDUCATION_LEVELS_MAP = {
    15: "O'rta",
    16: "O'rta maxsus",
    17: "Oliy",
    18: "Magistr",
    19: "Bakalavr"
}

UZBEKISTAN_REGIONS = [
    {"value": "10", "name": "Toshkent shahri", "code": "1726"},
    {"value": "11", "name": "Toshkent viloyati", "code": "1727"},
    {"value": "17", "name": "Andijon viloyati", "code": "1703"},
    {"value": "20", "name": "Buxoro viloyati", "code": "1706"},
    {"value": "15", "name": "Farg'ona viloyati", "code": "1730"},
    {"value": "13", "name": "Jizzax viloyati", "code": "1708"},
    {"value": "22", "name": "Xorazm viloyati", "code": "1733"},
    {"value": "16", "name": "Namangan viloyati", "code": "1714"},
    {"value": "21", "name": "Navoiy viloyati", "code": "1712"},
    {"value": "18", "name": "Qashqadaryo viloyati", "code": "1710"},
    {"value": "23", "name": "Qoraqalpog'iston Respublikasi", "code": "1735"},
    {"value": "14", "name": "Samarqand viloyati", "code": "1718"},
    {"value": "12", "name": "Sirdaryo viloyati", "code": "1724"},
    {"value": "19", "name": "Surxondaryo viloyati", "code": "1722"},
]

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "uz,ru;q=0.9,en;q=0.8",
    "Content-Type": "application/json"
}


# --------------------------------------------------------------------------------------
# 2. YORDAMCHI FUNKSIYALAR
# --------------------------------------------------------------------------------------
def extract_location(region_raw: Optional[str]) -> Tuple[str, str]:
    if not region_raw:
        return "Ko'rsatilmagan", "Ko'rsatilmagan"
    parts = [p.strip() for p in region_raw.split(",") if p.strip()]
    if len(parts) >= 2:
        return parts[0], parts[1]
    elif len(parts) == 1:
        return parts[0], "-"
    return "Ko'rsatilmagan", "Ko'rsatilmagan"


def format_date_str(date_val: Any) -> str:
    if not date_val:
        return "Belgilanmagan"
    if isinstance(date_val, str):
        try:
            cleaned = date_val.split(".")[0].replace("Z", "").replace("T", " ")
            dt = datetime.fromisoformat(cleaned)
            return dt.strftime("%d.%m.%Y %H:%M")
        except Exception:
            pass
    return str(date_val)[:16]


def format_salary(salary_val: Any) -> str:
    if not salary_val or salary_val == 0:
        return "Ko'rsatilmagan"
    try:
        val = int(salary_val)
        return f"{val:,} UZS".replace(",", " ")
    except (ValueError, TypeError):
        return f"{salary_val} UZS"


def format_salary_display(salary_val: Any) -> str:
    if not salary_val or salary_val == 0:
        return "Shtat jadvali bo'yicha"
    try:
        val = int(salary_val)
        return f"{val:,} UZS".replace(",", " ")
    except (ValueError, TypeError):
        return f"{salary_val} UZS"


UZ_MONTHS = {
    1: "sentabr" if False else "yanvar", 2: "fevral", 3: "mart", 4: "aprel",
    5: "may", 6: "iyun", 7: "iyul", 8: "avgust",
    9: "sentabr", 10: "oktabr", 11: "noyabr", 12: "dekabr"
}


def format_date_uzbek(date_val: Any) -> str:
    if not date_val:
        return "Belgilanmagan"
    if isinstance(date_val, str):
        try:
            cleaned = date_val.split(".")[0].replace("Z", "").replace("T", " ")
            dt = datetime.fromisoformat(cleaned)
            m_name = UZ_MONTHS.get(dt.month, "")
            return f"{dt.day:02d} {m_name} {dt.year}"
        except Exception:
            pass
    return str(date_val)[:10]


def format_vacancy_card_data(res: Dict[str, Any]) -> Dict[str, Any]:
    v_id = res.get("id")
    vil, tum = extract_location(res.get("region"))
    salary_raw = res.get("position_salary")
    salary_disp = format_salary_display(salary_raw)

    rate = res.get("position_rate")
    if rate is None or rate == 1.0 or rate == 1:
        work_disp = "To'liq"
    else:
        work_disp = f"{rate} stavka"

    exp = res.get("experience")
    exp_disp = f"Kamida {exp} yil" if (exp and exp > 0) else "Mehnat staji talab etilmaydi"

    raw_t = res.get("test_type_name")
    if not raw_t or str(raw_t).lower() in ("none", "null", "test turi #0", "ko'rsatilmagan", "noma'lum"):
        t_name = "Test talab etilmaydi"
    else:
        t_name = str(raw_t).strip()

    contract_disp = "Nomuayyan muddatli"
    d_start = format_date_uzbek(res.get("date_start"))
    d_stop = format_date_uzbek(res.get("date_stop"))

    struct = res.get("structure_name")
    if not struct or struct == "-" or str(struct).strip() == "":
        struct = "Boshqaruv apparati"

    return {
        "id": v_id,
        "position_name": res.get("position_name") or "Lavozim ko'rsatilmagan",
        "organization": res.get("organization") or "Tashkilot ko'rsatilmagan",
        "structure_name": struct,
        "region_display": res.get("region") or f"{vil}, {tum}",
        "region_only": vil,
        "district_only": tum,
        "salary_display": salary_disp,
        "work_type_display": work_disp,
        "contract_display": contract_disp,
        "experience_display": exp_disp,
        "test_type_name": t_name,
        "test_type_code": res.get("testTypeCode"),
        "date_start_uz": d_start,
        "date_stop_uz": d_stop,
        "views_count": res.get("views_count", 0),
        "likes_count": res.get("likesCount", 0),
        "candidates_count": res.get("get_candidates_count", 0),
        "is_internal": bool(res.get("is_internal")),
        "argos_url": f"https://vacancy.argos.uz/hrm-vacancy-detail/{v_id}"
    }


def clean_text_list(text: Optional[str]) -> List[str]:
    if not text:
        return []
    cleaned = re.sub(r'<[^>]+>', '\n', text)
    cleaned = cleaned.replace("&quot;", '"').replace("&nbsp;", " ")
    lines = []
    for line in cleaned.split("\n"):
        line = line.strip(" •-\t\r")
        if line:
            lines.append(line)
    return lines


def set_cell_background(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top: int = 100, bottom: int = 100, left: int = 150, right: int = 150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>\n'
        f'  <w:top w:w="{top}" w:type="dxa"/>\n'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>\n'
        f'  <w:left w:w="{left}" w:type="dxa"/>\n'
        f'  <w:right w:w="{right}" w:type="dxa"/>\n'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


# --------------------------------------------------------------------------------------
# 3. WORD (.DOCX) EKSPORTCHI
# --------------------------------------------------------------------------------------
class DocxExporter:
    def __init__(self, exports_dir: str = EXPORTS_DIR):
        self.exports_dir = exports_dir
        os.makedirs(self.exports_dir, exist_ok=True)

    def _setup_document(self) -> docx.Document:
        doc = docx.Document()
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)
        return doc

    def _write_vacancy_block(self, doc: docx.Document, item: Dict[str, Any], is_first: bool = True):
        if not is_first:
            doc.add_page_break()

        pos_name = item.get("position_name") or "Lavozim nomi ko'rsatilmagan"
        org_name = item.get("organization") or "Tashkilot nomi ko'rsatilmagan"

        raw_test = item.get("test_type_name")
        if not raw_test or str(raw_test).lower() in ("noma'lum", "none", "null", "test turi #0", "ko'rsatilmagan"):
            test_type = "Test talab etilmaydi"
        else:
            test_type = str(raw_test).strip()

        if "boshqaruv" in test_type.lower():
            badge_icon = "⭐"
            badge_color = "1E3A8A"  # To'q ko'k (#1E3A8A)
            val_bg_color = "EFF6FF"
            sub_title = "DAVLAT FUQAROLIK XIZMATI (BOSHQARUV XODIMI) VAKANT LAVOZIMI"
        elif "mutaxassis" in test_type.lower():
            badge_icon = "👔"
            badge_color = "0D9488"
            val_bg_color = "CCFBF1"
            sub_title = "DAVLAT FUQAROLIK XIZMATI (MUTAXASSIS) VAKANT LAVOZIMI"
        elif "hamshira" in test_type.lower() or "shifokor" in test_type.lower():
            badge_icon = "🩺"
            badge_color = "7C3AED"
            val_bg_color = "EDE9FE"
            sub_title = "TIBBIYOT VA SOG'LIQNI SAQLASH VAKANT LAVOZIMI"
        else:
            badge_icon = "ℹ️"
            badge_color = "64748B"
            val_bg_color = "F1F5F9"
            sub_title = "VAKANT ISH O'RNI"

        viloyat, tuman = extract_location(item.get("region"))
        org_full = item.get("organizationFull") or {}
        mail_addr = org_full.get("mail_address") or "-"

        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_sub = p_sub.add_run(sub_title)
        run_sub.font.size = Pt(9.5)
        run_sub.font.bold = True
        run_sub.font.color.rgb = RGBColor(100, 116, 139)

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_title = p_title.add_run(pos_name)
        run_title.font.size = Pt(16.5)
        run_title.font.bold = True
        run_title.font.color.rgb = RGBColor(14, 116, 144)

        p_org = doc.add_paragraph()
        run_org = p_org.add_run(f"🏛️ {org_name}")
        run_org.font.size = Pt(12)
        run_org.font.bold = True
        run_org.font.color.rgb = RGBColor(30, 41, 59)

        p_banner = doc.add_paragraph()
        p_banner.paragraph_format.space_before = Pt(4)
        p_banner.paragraph_format.space_after = Pt(12)

        run_lbl = p_banner.add_run(f"  {badge_icon} TEST TURI:  ")
        run_lbl.font.size = Pt(10.5)
        run_lbl.font.bold = True
        run_lbl.font.color.rgb = RGBColor(255, 255, 255)
        shd_lbl = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{badge_color}"/>')
        run_lbl._r.get_or_add_rPr().append(shd_lbl)

        run_val = p_banner.add_run(f"  {test_type}  ")
        run_val.font.size = Pt(11.5)
        run_val.font.bold = True
        if "boshqaruv" in test_type.lower():
            run_val.font.color.rgb = RGBColor(30, 58, 138)  # To'q ko'k
        else:
            run_val.font.color.rgb = RGBColor(15, 23, 42)
        shd_val = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{val_bg_color}"/>')
        run_val._r.get_or_add_rPr().append(shd_val)

        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        v_id = item.get("id")
        argos_link = f"https://vacancy.argos.uz/hrm-vacancy-detail/{v_id}" if v_id else "https://vacancy.argos.uz"

        params_data = [
            ("Lavozim nomi:", pos_name),
            ("Tashkilot:", org_name),
            ("Tarkibiy tuzilma / Bo'lim:", item.get("structure_name") or "Boshqaruv apparati"),
            ("📍 Hudud (Viloyat):", viloyat),
            ("🏙️ Tuman / Shahar:", tuman),
            ("📌 Tashkilot aniq manzili:", mail_addr),
            ("Oylik ish haqi:", format_salary_display(item.get("position_salary"))),
            ("⭐ Test turi:", test_type),
            ("Mehnat staji talabi:", f"Kamida {item.get('experience', 0)} yil" if item.get('experience') else "Mehnat staji talab etilmaydi"),
            ("Ta'lim darajasi:", item.get("education_level_name") or "Talab etilmaydi"),
            ("Bandlik shakli:", f"To'liq ({item.get('position_rate', 1.0)} stavka)"),
            ("Qabul muddati:", f"{format_date_uzbek(item.get('date_start'))} — {format_date_uzbek(item.get('date_stop'))}"),
            ("Nomzodlar soni:", f"{item.get('get_candidates_count', 0)} ta"),
            ("🔗 Rasmiy Argos havolasi:", argos_link),
            ("Vakansiya ID raqami:", str(v_id or "-"))
        ]

        for i, (k, v) in enumerate(params_data):
            row = table.add_row()
            c1, c2 = row.cells
            c1.width = Inches(2.4)
            c2.width = Inches(4.5)

            bg_color = "F8FAFC" if i % 2 == 0 else "FFFFFF"
            set_cell_background(c1, bg_color)
            set_cell_background(c2, bg_color)
            set_cell_margins(c1, top=70, bottom=70, left=120, right=120)
            set_cell_margins(c2, top=70, bottom=70, left=120, right=120)

            p1 = c1.paragraphs[0]
            p1.paragraph_format.space_before = Pt(2)
            p1.paragraph_format.space_after = Pt(2)
            r1 = p1.add_run(k)
            r1.font.size = Pt(10)
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(71, 85, 105)

            p2 = c2.paragraphs[0]
            p2.paragraph_format.space_before = Pt(2)
            p2.paragraph_format.space_after = Pt(2)
            r2 = p2.add_run(str(v))
            r2.font.size = Pt(10.5)

            if "Test turi" in k:
                r2.font.bold = True
                r2.font.color.rgb = RGBColor(30, 58, 138)  # To'q ko'k (#1E3A8A)
            elif "Rasmiy Argos" in k:
                r2.font.bold = True
                r2.font.color.rgb = RGBColor(37, 99, 235)  # Ko'k havola
            elif "Oylik ish haqi" in k:
                r2.font.bold = True
                r2.font.color.rgb = RGBColor(22, 101, 52)
            else:
                r2.font.color.rgb = RGBColor(30, 41, 59)

        req_lines = clean_text_list(item.get("position_requirements"))
        if req_lines:
            h1 = doc.add_heading(level=2)
            h1.paragraph_format.space_before = Pt(14)
            h1.paragraph_format.space_after = Pt(4)
            rh1 = h1.add_run("📋 Malakaviy talablar")
            rh1.font.size = Pt(12.5)
            rh1.font.bold = True
            rh1.font.color.rgb = RGBColor(14, 116, 144)

            for line in req_lines:
                p_item = doc.add_paragraph(style='List Bullet')
                p_item.paragraph_format.space_before = Pt(2)
                p_item.paragraph_format.space_after = Pt(2)
                r_item = p_item.add_run(line)
                r_item.font.size = Pt(10)
                r_item.font.color.rgb = RGBColor(51, 65, 85)

        duty_lines = clean_text_list(item.get("position_duties"))
        if duty_lines:
            h2 = doc.add_heading(level=2)
            h2.paragraph_format.space_before = Pt(12)
            h2.paragraph_format.space_after = Pt(4)
            rh2 = h2.add_run("⚖️ Lavozim majburiyatlari")
            rh2.font.size = Pt(12.5)
            rh2.font.bold = True
            rh2.font.color.rgb = RGBColor(14, 116, 144)

            for line in duty_lines:
                p_item = doc.add_paragraph(style='List Bullet')
                p_item.paragraph_format.space_before = Pt(2)
                p_item.paragraph_format.space_after = Pt(2)
                r_item = p_item.add_run(line)
                r_item.font.size = Pt(10)
                r_item.font.color.rgb = RGBColor(51, 65, 85)

        cond_text = item.get("position_conditions")
        if cond_text:
            h3 = doc.add_heading(level=2)
            h3.paragraph_format.space_before = Pt(12)
            h3.paragraph_format.space_after = Pt(4)
            rh3 = h3.add_run("💼 Ish sharoitlari")
            rh3.font.size = Pt(12.5)
            rh3.font.bold = True
            rh3.font.color.rgb = RGBColor(14, 116, 144)

            for line in clean_text_list(cond_text):
                p_cond = doc.add_paragraph(style='List Bullet')
                p_cond.paragraph_format.space_before = Pt(2)
                p_cond.paragraph_format.space_after = Pt(2)
                rc = p_cond.add_run(line)
                rc.font.size = Pt(10)
                rc.font.color.rgb = RGBColor(51, 65, 85)

        phone = org_full.get("phone")
        email = org_full.get("email")
        if phone or email:
            h4 = doc.add_heading(level=2)
            h4.paragraph_format.space_before = Pt(12)
            h4.paragraph_format.space_after = Pt(4)
            rh4 = h4.add_run("📞 Bog'lanish")
            rh4.font.size = Pt(12.5)
            rh4.font.bold = True
            rh4.font.color.rgb = RGBColor(14, 116, 144)

            if phone:
                p_ph = doc.add_paragraph()
                p_ph.add_run("Telefon: ").font.bold = True
                p_ph.add_run(str(phone))
            if email:
                p_em = doc.add_paragraph()
                p_em.add_run("Email: ").font.bold = True
                p_em.add_run(str(email))

        vac_id = item.get("id")
        if vac_id:
            p_link = doc.add_paragraph()
            p_link.paragraph_format.space_before = Pt(12)
            r_lk = p_link.add_run(f"Havola: https://vacancy.argos.uz/hrm-vacancy-detail/{vac_id}")
            r_lk.font.size = Pt(9)
            r_lk.font.italic = True
            r_lk.font.color.rgb = RGBColor(100, 116, 139)

    def export_single_vacancy(self, item: Dict[str, Any], filename: Optional[str] = None) -> str:
        doc = self._setup_document()
        self._write_vacancy_block(doc, item, is_first=True)

        if not filename:
            item_id = item.get("id", "vakansiya")
            pos = item.get("position_name", "lavozim")
            test = item.get("test_type_name", "test")
            safe_pos = re.sub(r'[\\/*?:"<>|]', '_', pos)[:35]
            safe_test = re.sub(r'[\\/*?:"<>|]', '_', test)[:25]
            now_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{item_id}_{safe_pos}_{safe_test}_{now_stamp}.docx"

        if not filename.endswith(".docx"):
            filename += ".docx"

        filepath = os.path.join(self.exports_dir, filename)
        doc.save(filepath)
        try:
            meta_path = filepath + ".json"
            card_items = [format_vacancy_card_data(item)]
            with open(meta_path, "w", encoding="utf-8") as mf:
                json.dump(card_items, mf, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return filepath

    def export_multiple_vacancies(self, items: List[Dict[str, Any]], filename: Optional[str] = None) -> str:
        if not items:
            raise ValueError("Vakansiyalar mavjud emas!")

        doc = self._setup_document()

        p_main = doc.add_paragraph()
        p_main.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_main.paragraph_format.space_before = Pt(10)
        p_main.paragraph_format.space_after = Pt(2)
        r_main = p_main.add_run("ARGOS.UZ VAKANSIYALAR TO'PLAMI")
        r_main.font.size = Pt(19)
        r_main.font.bold = True
        r_main.font.color.rgb = RGBColor(14, 116, 144)

        p_info = doc.add_paragraph()
        p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_info.paragraph_format.space_after = Pt(16)
        now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
        p_info.add_run(f"Yaratilgan sana: {now_str} | Jami vakansiyalar soni: {len(items)} ta").font.color.rgb = RGBColor(100, 116, 139)

        h_table = doc.add_heading(level=2)
        rh_t = h_table.add_run("📊 Vakansiyalar umumiy ro'yxati")
        rh_t.font.size = Pt(13)
        rh_t.font.bold = True
        rh_t.font.color.rgb = RGBColor(30, 41, 59)

        summary_table = doc.add_table(rows=1, cols=6)
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        summary_table.autofit = False

        headers = ["№", "Lavozim nomi", "Tashkilot", "Hudud (Viloyat)", "Tuman / Shahar", "Test turi"]
        hdr_row = summary_table.rows[0]
        widths = [Inches(0.4), Inches(2.0), Inches(1.8), Inches(1.1), Inches(1.1), Inches(1.5)]

        for idx, text in enumerate(headers):
            cell = hdr_row.cells[idx]
            cell.width = widths[idx]
            set_cell_background(cell, "0284C7")
            set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
            p = cell.paragraphs[0]
            run = p.add_run(text)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)

        for i, it in enumerate(items, 1):
            row = summary_table.add_row()
            bg = "F1F5F9" if i % 2 == 0 else "FFFFFF"
            viloyat, tuman = extract_location(it.get("region"))
            raw_t = it.get("test_type_name")
            if not raw_t or str(raw_t).lower() in ("noma'lum", "none", "null", "test turi #0", "ko'rsatilmagan"):
                display_t = "Test talab etilmaydi"
            else:
                display_t = str(raw_t).strip()

            vals = [
                str(i),
                it.get("position_name") or "-",
                it.get("organization") or "-",
                viloyat,
                tuman,
                display_t
            ]
            for c_idx, val in enumerate(vals):
                cell = row.cells[c_idx]
                cell.width = widths[c_idx]
                set_cell_background(cell, bg)
                set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
                p = cell.paragraphs[0]
                run = p.add_run(val)
                run.font.size = Pt(8.5)
                if c_idx == 5:
                    run.font.bold = True
                    if "boshqaruv" in display_t.lower():
                        run.font.color.rgb = RGBColor(3, 105, 161)
                    elif "mutaxassis" in display_t.lower():
                        run.font.color.rgb = RGBColor(13, 148, 136)
                    elif "hamshira" in display_t.lower() or "shifokor" in display_t.lower():
                        run.font.color.rgb = RGBColor(124, 58, 237)
                    else:
                        run.font.color.rgb = RGBColor(100, 116, 139)
                elif c_idx in (3, 4):
                    run.font.color.rgb = RGBColor(30, 41, 59)

        for item in items:
            self._write_vacancy_block(doc, item, is_first=False)

        if not filename:
            now_stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:19]
            filename = f"Argos_Vakansiyalar_Toplami_{now_stamp}.docx"

        if not filename.endswith(".docx"):
            filename += ".docx"

        filepath = os.path.join(self.exports_dir, filename)
        doc.save(filepath)
        try:
            meta_path = filepath + ".json"
            card_items = [format_vacancy_card_data(it) for it in items]
            with open(meta_path, "w", encoding="utf-8") as mf:
                json.dump(card_items, mf, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return filepath


# --------------------------------------------------------------------------------------
# 4. ARGOS API MIJOZI
# --------------------------------------------------------------------------------------
class ArgosApiClient:
    def __init__(self):
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=35, pool_maxsize=35, max_retries=3)
        self.session.mount('https://', adapter)
        self.session.mount('http://', adapter)
        self.session.headers.update(DEFAULT_HEADERS)
        self.test_types = dict(TEST_TYPES_MAP)
        self.education_levels = dict(EDUCATION_LEVELS_MAP)
        self._districts_cache: Optional[List[Dict[str, Any]]] = None
        self._institutions_cache: Optional[List[Dict[str, Any]]] = None
        self._load_test_types()
        self._load_districts()

    def _load_test_types(self) -> None:
        try:
            r = self.session.get(API_TEST_TYPES_ENDPOINT, timeout=5)
            if r.status_code == 200:
                items = r.json()
                for item in items:
                    code = item.get("code")
                    name = item.get("nameUz") or item.get("label")
                    if code and name and name != "Танланг":
                        try:
                            self.test_types[int(code)] = name
                        except ValueError:
                            pass
        except Exception:
            pass

    def _load_districts(self) -> None:
        try:
            r = self.session.get(API_DISTRICTS_ENDPOINT, timeout=6)
            if r.status_code == 200:
                self._districts_cache = r.json()
        except Exception:
            pass

    def _load_institutions(self) -> None:
        try:
            r = self.session.get(API_INSTITUTIONS_ENDPOINT, timeout=12)
            if r.status_code == 200:
                self._institutions_cache = r.json()
        except Exception:
            pass

    def get_organizations(
        self,
        region_id: Optional[str] = None,
        district_id: Optional[str] = None,
        search: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        if self._institutions_cache is None:
            self._load_institutions()
        if not self._institutions_cache:
            return []

        filtered = self._institutions_cache
        if region_id and str(region_id).strip() not in ("0", "Barchasi", ""):
            try:
                r_num = int(region_id)
                filtered = [it for it in filtered if it.get("regionId") == r_num]
            except ValueError:
                pass

        if district_id and str(district_id).strip() not in ("0", "Barchasi", ""):
            try:
                d_num = int(district_id)
                filtered = [it for it in filtered if it.get("districtId") == d_num]
            except ValueError:
                pass

        if search and search.strip():
            s_low = search.strip().lower()
            filtered = [
                it for it in filtered
                if s_low in (it.get("nameUz") or it.get("label") or "").lower()
            ]

        results = []
        for it in filtered[:300]:
            name = it.get("nameUz") or it.get("label")
            it_id = it.get("id")
            if name and it_id:
                results.append({
                    "id": it_id,
                    "name": name,
                    "tin": it.get("tin"),
                    "regionId": it.get("regionId"),
                    "districtId": it.get("districtId")
                })
        return results

    def get_test_type_name(self, code: Optional[Any]) -> str:
        if code is None or code == "" or str(code).lower() in ("none", "null"):
            return "Test talab etilmaydi"
        try:
            code_int = int(code)
            if code_int == 0:
                return "Test talab etilmaydi"
            return self.test_types.get(code_int, f"Test turi #{code_int}")
        except (ValueError, TypeError):
            return str(code)

    def resolve_test_type_target(self, filter_str: Optional[str]) -> Tuple[Optional[Any], str]:
        if not filter_str:
            return None, "Barchasi"
        clean = filter_str.strip().lower()
        if clean in ("barchasi", "all", "", "tanlang"):
            return None, "Barchasi"

        if clean in TEST_TYPE_ALIASES:
            alias_val = TEST_TYPE_ALIASES[clean]
            if alias_val == "all":
                return None, "Barchasi"
            elif alias_val == "dfx_all":
                return "dfx_all", "Davlat fuqarolik xizmati (Barchasi)"
            elif alias_val in self.test_types:
                return alias_val, self.test_types[alias_val]
            return alias_val, str(alias_val)

        if "boshqaruv" in clean:
            return 4, self.test_types.get(4, "Davlat fuqarolik xizmatchisi (Boshqaruv xodimi)")
        if "mutaxassis" in clean:
            return 5, self.test_types.get(5, "Davlat fuqarolik xizmatchisi (Mutaxassis)")
        if "dfx" in clean or "fuqarolik xizmati" in clean:
            return "dfx_all", "Davlat fuqarolik xizmati (Barchasi)"
        if "hamshira" in clean:
            return 1225, self.test_types.get(1225, "Hamshiralik ishi")
        if "shifokor" in clean:
            return 1226, self.test_types.get(1226, "Shifokorlar uchun")
        if "talab etilmaydi" in clean or "mavjud emas" in clean:
            return 0, "Test talab etilmaydi"

        return clean, filter_str

    def get_districts_by_region(self, region_value: str) -> List[Dict[str, str]]:
        if not region_value or region_value in ("0", "Barchasi"):
            return []
        if self._districts_cache is None:
            self._load_districts()
        if not self._districts_cache:
            return []
        results = []
        for d in self._districts_cache:
            if str(d.get("rootCode")) == str(region_value):
                name = d.get("nameUz") or d.get("label")
                val = d.get("value")
                if name and val and name != "Tanlang":
                    results.append({"name": name, "value": str(val)})
        return results


    def get_education_level_name(self, code: Optional[int]) -> str:
        if code is None:
            return "Talab etilmaydi"
        try:
            code_int = int(code)
            return self.education_levels.get(code_int, f"Daraja #{code}")
        except (ValueError, TypeError):
            return str(code)

    def parse_url(self, url: str) -> Tuple[str, Dict[str, Any]]:
        url = url.strip()
        detail_match = re.search(r'hrm-vacancy-detail/(\d+)', url)
        if detail_match:
            return 'detail', {'id': int(detail_match.group(1))}

        parsed = urllib.parse.urlparse(url)
        params = urllib.parse.parse_qs(parsed.query)

        if 'id' in params:
            try:
                return 'detail', {'id': int(params['id'][0])}
            except ValueError:
                pass

        search_payload: Dict[str, Any] = {
            "page": 0,
            "pageSize": 40,
            "search": None,
            "regionSoato": None,
            "districtSoato": None,
            "organizationId": None,
            "fatherOrganizationId": None,
            "civilServant": None,
            "vacancyType": None,
            "isDisablity": None,
            "isSmallEmployee": None,
            "isInternal": None,
            "minSalary": None,
            "workExperience": None,
            "organizationTin": None
        }

        if 'page' in params:
            try:
                search_payload['page'] = int(params['page'][0])
            except ValueError:
                pass
        if 'search' in params and params['search'][0]:
            search_payload['search'] = params['search'][0]
        reg_param = params.get('regionSoato', [None])[0] or params.get('region', [None])[0]
        if reg_param:
            try:
                search_payload['regionSoato'] = int(reg_param)
            except ValueError:
                search_payload['regionSoato'] = reg_param

        dist_param = (
            params.get('districtSoato', [None])[0] or 
            params.get('districtId', [None])[0] or 
            params.get('district', [None])[0]
        )
        if dist_param:
            try:
                search_payload['districtSoato'] = int(dist_param)
            except ValueError:
                search_payload['districtSoato'] = dist_param

        if 'civilServant' in params and params['civilServant'][0]:
            try:
                search_payload['civilServant'] = int(params['civilServant'][0])
            except ValueError:
                pass

        return 'list', search_payload


    def get_vacancy_list(self, payload: Optional[Dict[str, Any]] = None, page: int = 0, page_size: int = 40) -> Dict[str, Any]:
        data = dict(payload or {})
        data["page"] = page
        data["pageSize"] = min(page_size, 40)
        resp = self.session.post(API_LIST_ENDPOINT, json=data, timeout=10)
        resp.raise_for_status()
        return resp.json()

    def get_vacancy_detail(self, vacancy_id: int, retries: int = 2) -> Dict[str, Any]:
        url = f"{API_DETAIL_ENDPOINT}?id={vacancy_id}&isView=false"
        last_err = None
        for attempt in range(retries + 1):
            try:
                resp = self.session.post(url, timeout=8)
                resp.raise_for_status()
                data = resp.json()
                test_code = data.get("testTypeCode")
                data["test_type_name"] = self.get_test_type_name(test_code)
                data["education_level_name"] = self.get_education_level_name(data.get("education_Level_Id"))
                return data
            except Exception as e:
                last_err = e
                if attempt < retries:
                    time.sleep(0.3 * (attempt + 1))
        raise last_err or Exception(f"Vakansiya #{vacancy_id} tafsilotlarini olib bo'lmadi")

    def search_all_vacancies(
        self,
        base_payload: Optional[Dict[str, Any]] = None,
        max_items: Optional[int] = None,
        test_type_filter: Optional[str] = None,
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
        stop_check: Optional[Callable[[], bool]] = None
    ) -> List[Dict[str, Any]]:
        payload = dict(base_payload or {})
        page = payload.get("page", 0)
        page_size = 40
        collected: List[Dict[str, Any]] = []

        target_code, target_label = self.resolve_test_type_target(test_type_filter)
        has_filter = (target_code is not None)

        if target_code in (4, 5, "dfx_all") and payload.get("civilServant") is None:
            payload["civilServant"] = 1
        elif target_code in (1225, 1226) and payload.get("fatherOrganizationId") is None and payload.get("civilServant") is None:
            payload["fatherOrganizationId"] = 1052

        first_resp = self.get_vacancy_list(payload, page=page, page_size=page_size)
        total_count = first_resp.get("count", 0)
        results = first_resp.get("results", [])

        if progress_callback:
            filt_desc = f" [{target_label}]" if has_filter else ""
            kw = payload.get("search")
            if kw:
                filt_desc += f' (Kalit so\'z: "{kw}")'
            progress_callback({
                "type": "status",
                "message": f"Jami {total_count:,} ta vakansiya mavjud{filt_desc}. Tahlil boshlandi...",
                "cur": 0,
                "tot": total_count
            })

        def fetch_and_check(item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
            if stop_check and stop_check():
                return None
            vac_id = item.get("id")
            if not vac_id:
                return None

            try:
                detail = self.get_vacancy_detail(vac_id)
                t_code = detail.get("testTypeCode")
                t_name = detail.get("test_type_name", "")

                if has_filter:
                    try:
                        t_code_num = int(t_code) if t_code is not None and str(t_code).strip() not in ("", "none", "null") else None
                    except (ValueError, TypeError):
                        t_code_num = None

                    if target_code == "dfx_all":
                        if t_code_num not in (4, 5) and not ("boshqaruv" in t_name.lower() or "mutaxassis" in t_name.lower() or "fuqarolik" in t_name.lower()):
                            return None
                    elif target_code == 0:
                        if t_code_num not in (0, None) and "talab etilmaydi" not in t_name.lower():
                            return None
                    elif isinstance(target_code, int):
                        if t_code_num != target_code:
                            return None
                    else:
                        q = str(target_code).lower()
                        if q not in t_name.lower() and q not in str(t_code) and (t_code_num is None or str(t_code_num) != q):
                            return None

                return detail
            except Exception:
                return None

        with ThreadPoolExecutor(max_workers=10) as executor:
            while True:
                if stop_check and stop_check():
                    if progress_callback:
                        progress_callback({
                            "type": "stopped",
                            "message": f"To'xtatildi. Yig'ilgan: {len(collected)} ta",
                            "cur": len(collected),
                            "tot": total_count
                        })
                    break

                scanned_so_far = min((page + 1) * page_size, total_count)

                futures = [executor.submit(fetch_and_check, it) for it in results]
                for fut in as_completed(futures):
                    if stop_check and stop_check():
                        break
                    res = fut.result()
                    if res:
                        collected.append(res)
                        if progress_callback:
                            card_item = format_vacancy_card_data(res)
                            progress_callback({
                                "type": "found",
                                "count": len(collected),
                                "card": card_item,
                                "id": card_item["id"],
                                "position": card_item["position_name"],
                                "organization": card_item["organization"],
                                "region": card_item["region_only"],
                                "district": card_item["district_only"],
                                "test_type": card_item["test_type_name"],
                                "cur": scanned_so_far,
                                "tot": total_count
                            })
                        if max_items and len(collected) >= max_items:
                            break

                pct = int((scanned_so_far / max(1, total_count)) * 100)
                if progress_callback and (not max_items or len(collected) < max_items):
                    progress_callback({
                        "type": "progress",
                        "message": f"Skaner qilindi: {scanned_so_far:,} / {total_count:,} ({pct}%) | Tanlandi: {len(collected)} ta",
                        "cur": scanned_so_far,
                        "tot": total_count
                    })

                if max_items and len(collected) >= max_items:
                    break

                if scanned_so_far >= total_count or not results:
                    break

                page += 1
                try:
                    next_resp = self.get_vacancy_list(payload, page=page, page_size=page_size)
                    results = next_resp.get("results", [])
                except Exception as e:
                    if progress_callback:
                        progress_callback({
                            "type": "warning",
                            "message": f"Sahifa {page} yuklanishida ogohlantirish: {e}",
                            "cur": scanned_so_far,
                            "tot": total_count
                        })
                    break

        if max_items and len(collected) > max_items:
            collected = collected[:max_items]

        return collected



# --------------------------------------------------------------------------------------
# 5. FLASK WEB ILOVASI (Web Server)
# --------------------------------------------------------------------------------------
app = Flask(__name__, template_folder="templates", static_folder="static")
api_client = ArgosApiClient()
docx_exporter = DocxExporter(exports_dir=EXPORTS_DIR)

# Faol jarayonlarni boshqarish uchun flaglar
current_scan_lock = threading.Lock()
stop_flag = False


@app.route("/")
def index():
    """Asosiy Web sahifani yuklaydi"""
    return render_template("index.html")


@app.route("/favicon.ico")
def favicon():
    """Brauzer tab nishoni (O'zbekiston Davlat Gerbi)"""
    return send_from_directory(os.path.join(BASE_DIR, "static", "img"), "emblem.svg", mimetype="image/svg+xml")


@app.route("/assets/images/<path:filename>")
def assets_images(filename):
    """Argos statik rasmlari (masalan, independent.png)"""
    return send_from_directory(os.path.join(BASE_DIR, "static", "assets", "images"), filename)


@app.route("/api/regions")
def get_regions():
    """Viloyatlar ro'yxati"""
    return jsonify(UZBEKISTAN_REGIONS)


@app.route("/api/districts")
def get_districts():
    """Viloyat bo'yicha tumanlar ro'yxati"""
    region_val = request.args.get("region", "").strip()
    if not region_val:
        return jsonify([])
    districts = api_client.get_districts_by_region(region_val)
    return jsonify(districts)


@app.route("/api/organizations")
def get_organizations():
    """Tashkilotlar ro'yxati (viloyat va tuman bo'yicha filtrlangan)"""
    reg = request.args.get("region", "").strip()
    dist = request.args.get("district", "").strip()
    q = request.args.get("search", "").strip()
    items = api_client.get_organizations(region_id=reg, district_id=dist, search=q)
    return jsonify(items)


@app.route("/api/test-types")
def get_test_types():
    """Mavjud test turlari"""
    options = [
        {"value": "all", "label": "Barchasi"},
        {"value": "dfx_all", "label": "Davlat fuqarolik xizmati (Barchasi)"},
        {"value": "4", "label": "Davlat fuqarolik xizmatchisi (Boshqaruv xodimi)"},
        {"value": "5", "label": "Davlat fuqarolik xizmatchisi (Mutaxassis)"},
        {"value": "1225", "label": "Hamshiralik ishi"},
        {"value": "1226", "label": "Shifokorlar uchun"},
        {"value": "0", "label": "Test talab etilmaydi (Texnik / Boshqa)"}
    ]
    return jsonify(options)


# In-memory kesh (TTL: 60 soniya)
_vacancies_cache: Dict[str, Tuple[float, Dict[str, Any]]] = {}


@app.route("/api/vacancies", methods=["GET", "POST"])
def get_vacancies():
    """
    Vakansiyalar ro'yxatini sahifalangan holda tezkor qaytaradi (boshlang'ich default yoki filtrlangan).
    """
    if request.method == "POST":
        data = request.json or {}
    else:
        data = request.args.to_dict()

    try:
        page = int(data.get("page", 0))
    except (ValueError, TypeError):
        page = 0

    try:
        page_size = int(data.get("pageSize", data.get("page_size", 20)))
    except (ValueError, TypeError):
        page_size = 20
    page_size = max(1, min(page_size, 40))

    keyword = (data.get("search") or "").strip() or None
    region_val = (data.get("region") or "").strip() or None
    district_val = (data.get("district") or "").strip() or None
    organization_id = data.get("organizationId") or None
    direction = (data.get("direction") or "").strip() or None
    vacancy_type = data.get("vacancyType")
    vacancy_level = data.get("vacancyLevel")
    work_experience = data.get("workExperience")
    min_salary = data.get("minSalary")
    is_disability = data.get("isDisablity") if "isDisablity" in data else data.get("isDisability")
    test_type = (data.get("test_type") or "").strip() or None

    payload: Dict[str, Any] = {
        "page": page,
        "pageSize": page_size,
        "search": keyword,
        "regionSoato": None,
        "districtSoato": None,
        "organizationId": None,
        "fatherOrganizationId": None,
        "civilServant": None,
        "vacancyType": None,
        "isDisablity": None,
        "isSmallEmployee": None,
        "isInternal": None,
        "minSalary": None,
        "workExperience": None,
        "organizationTin": None
    }

    if region_val and region_val not in ("0", "Barchasi", ""):
        try:
            payload["regionSoato"] = int(region_val)
        except ValueError:
            payload["regionSoato"] = region_val

    if district_val and district_val not in ("0", "Barchasi", ""):
        try:
            payload["districtSoato"] = int(district_val)
        except ValueError:
            payload["districtSoato"] = district_val

    if organization_id and str(organization_id).strip() not in ("0", ""):
        try:
            payload["organizationId"] = int(organization_id)
        except ValueError:
            payload["organizationId"] = organization_id

    if direction:
        dir_str = str(direction).strip()
        if dir_str == "1":
            payload["fatherOrganizationId"] = 1052  # Sog'liqni saqlash
        elif dir_str == "2":
            payload["civilServant"] = 1  # Davlat fuqarolik xizmati
        elif dir_str == "3":
            payload["isOtherOrganization"] = True  # Boshqalar

    if vacancy_type is not None and str(vacancy_type).strip() not in ("", "null"):
        try:
            payload["isInternal"] = int(vacancy_type)
        except ValueError:
            pass

    if vacancy_level and str(vacancy_level).strip() not in ("", "0", "null"):
        payload["vacancyLevel"] = [str(vacancy_level).strip()]

    if work_experience is not None and str(work_experience).strip() not in ("", "null"):
        try:
            payload["workExperience"] = int(work_experience)
        except ValueError:
            payload["workExperience"] = str(work_experience).strip()

    if min_salary is not None and str(min_salary).strip() not in ("", "0", "null"):
        try:
            payload["minSalary"] = int(min_salary)
        except ValueError:
            pass

    if is_disability is True or str(is_disability).lower() in ("true", "1"):
        payload["isDisablity"] = True

    # Kesh tekshirish
    cache_key = json.dumps(payload, sort_keys=True) + f"_tt_{test_type}"
    now = time.time()
    if cache_key in _vacancies_cache:
        cached_time, cached_res = _vacancies_cache[cache_key]
        if now - cached_time < 60:
            return jsonify(cached_res)

    try:
        raw_list = api_client.get_vacancy_list(payload=payload, page=page, page_size=page_size)
        total_count = raw_list.get("count", 0)
        items = raw_list.get("results", [])

        # Har bir vakansiyaning to'liq ma'lumotlarini (test turi va h.k.) concurrent boyitish
        def enrich_item(item):
            v_id = item.get("id")
            try:
                detail = api_client.get_vacancy_detail(v_id)
                return format_vacancy_card_data(detail)
            except Exception:
                return format_vacancy_card_data(item)

        with ThreadPoolExecutor(max_workers=10) as executor:
            cards = list(executor.map(enrich_item, items))

        total_pages = (total_count + page_size - 1) // page_size if total_count > 0 else 0
        response_data = {
            "status": "ok",
            "count": total_count,
            "page": page,
            "pageSize": page_size,
            "totalPages": total_pages,
            "vacancies": cards
        }

        # Keshga saqlash (maksimal 100 ta kalit)
        _vacancies_cache[cache_key] = (now, response_data)
        if len(_vacancies_cache) > 100:
            oldest_keys = sorted(_vacancies_cache.keys(), key=lambda k: _vacancies_cache[k][0])[:25]
            for k in oldest_keys:
                _vacancies_cache.pop(k, None)

        return jsonify(response_data)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/scan", methods=["POST"])
def scan_vacancies():
    """
    Vakansiyalarni qidirish va Word hujjati yaratish (Server-Sent Events streaming).
    """
    global stop_flag
    data = request.json or {}

    url = (data.get("url") or "").strip() or DEFAULT_LIST_URL
    region_val = (data.get("region") or "").strip() or None
    district_val = (data.get("district") or "").strip() or None
    organization_id = data.get("organizationId") or None
    direction = (data.get("direction") or "").strip() or None
    vacancy_type = data.get("vacancyType")
    vacancy_level = data.get("vacancyLevel")
    work_experience = data.get("workExperience")
    min_salary = data.get("minSalary")
    is_disability = data.get("isDisablity") if "isDisablity" in data else data.get("isDisability")
    test_type = (data.get("test_type") or "").strip() or "Barchasi"
    keyword = (data.get("search") or "").strip() or None
    limit_raw = data.get("limit")

    limit = None
    if limit_raw and str(limit_raw).isdigit():
        limit = int(limit_raw)

    stop_flag = False

    def event_stream():
        global stop_flag
        yield f"data: {json.dumps({'type': 'status', 'message': f'URL tahlil qilinmoqda: {url}', 'cur': 0, 'tot': 0})}\n\n"

        page_type, parsed_params = api_client.parse_url(url)

        if page_type == "detail":
            vac_id = parsed_params["id"]
            yield f"data: {json.dumps({'type': 'status', 'message': f'Vakansiya #{vac_id} yuklanmoqda...', 'cur': 1, 'tot': 1})}\n\n"
            try:
                detail = api_client.get_vacancy_detail(vac_id)
                card_item = format_vacancy_card_data(detail)
                yield f"data: {json.dumps({'type': 'found', 'count': 1, 'card': card_item, 'id': vac_id, 'position': card_item['position_name'], 'organization': card_item['organization'], 'region': card_item['region_only'], 'district': card_item['district_only'], 'test_type': card_item['test_type_name'], 'cur': 1, 'tot': 1})}\n\n"
                
                yield f"data: {json.dumps({'type': 'status', 'message': 'Word (.docx) hujjati shakllantirilmoqda...', 'cur': 1, 'tot': 1})}\n\n"
                filepath = docx_exporter.export_single_vacancy(detail)
                fname = os.path.basename(filepath)

                yield f"data: {json.dumps({'type': 'done', 'filename': fname, 'count': 1, 'docx_url': f'/api/docx-raw/{fname}', 'preview_url': f'/api/preview/{fname}', 'vacancies': [card_item], 'message': 'Muvaffaqiyatli yakunlandi!'})}\n\n"
            except Exception as e:
                yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"
            return

        # Ro'yxat parametrlari
        if keyword:
            parsed_params["search"] = keyword
        if region_val and region_val not in ("0", "Barchasi"):
            try:
                parsed_params["regionSoato"] = int(region_val)
            except ValueError:
                parsed_params["regionSoato"] = region_val
        if district_val and district_val not in ("0", "Barchasi"):
            try:
                parsed_params["districtSoato"] = int(district_val)
            except ValueError:
                parsed_params["districtSoato"] = district_val

        if organization_id and str(organization_id).strip() not in ("0", ""):
            try:
                parsed_params["organizationId"] = int(organization_id)
            except ValueError:
                parsed_params["organizationId"] = organization_id

        if direction:
            dir_str = str(direction).strip()
            if dir_str == "1":
                parsed_params["fatherOrganizationId"] = 1052  # Sog'liqni saqlash
            elif dir_str == "2":
                parsed_params["civilServant"] = 1  # Davlat fuqarolik xizmati
            elif dir_str == "3":
                parsed_params["isOtherOrganization"] = True  # Boshqalar

        if vacancy_type is not None and str(vacancy_type).strip() not in ("", "null"):
            try:
                parsed_params["isInternal"] = int(vacancy_type)
            except ValueError:
                pass

        if vacancy_level and str(vacancy_level).strip() not in ("", "0", "null"):
            parsed_params["vacancyLevel"] = [str(vacancy_level).strip()]

        if work_experience is not None and str(work_experience).strip() not in ("", "null"):
            try:
                parsed_params["workExperience"] = int(work_experience)
            except ValueError:
                parsed_params["workExperience"] = str(work_experience).strip()

        if min_salary is not None and str(min_salary).strip() not in ("", "0", "null"):
            try:
                parsed_params["minSalary"] = int(min_salary)
            except ValueError:
                pass

        if is_disability is True or str(is_disability).lower() in ("true", "1"):
            parsed_params["isDisablity"] = True

        clean_filter = test_type if test_type.lower() not in ("barchasi", "all", "") else None

        collected_items = []
        event_queue = []

        def on_progress(evt_data):
            event_queue.append(evt_data)

        def should_stop():
            return stop_flag

        def _run_search():
            nonlocal collected_items
            try:
                collected_items = api_client.search_all_vacancies(
                    base_payload=parsed_params,
                    max_items=limit,
                    test_type_filter=clean_filter,
                    progress_callback=on_progress,
                    stop_check=should_stop
                )
            except Exception as e:
                event_queue.append({"type": "error", "message": str(e)})

        search_thread = threading.Thread(target=_run_search, daemon=True)
        search_thread.start()

        while search_thread.is_alive() or event_queue:
            while event_queue:
                evt = event_queue.pop(0)
                yield f"data: {json.dumps(evt)}\n\n"
            time.sleep(0.08)

        if not collected_items:
            yield f"data: {json.dumps({'type': 'done_empty', 'message': 'Belgilangan shartlarga mos vakansiyalar topilmadi.'})}\n\n"
            return

        yield f"data: {json.dumps({'type': 'status', 'message': f'Word (.docx) hujjati yaratilmoqda ({len(collected_items)} ta vakansiya)...', 'cur': len(collected_items), 'tot': len(collected_items)})}\n\n"

        try:
            filepath = docx_exporter.export_multiple_vacancies(collected_items)
            fname = os.path.basename(filepath)
            cards_list = [format_vacancy_card_data(it) for it in collected_items]
            yield f"data: {json.dumps({'type': 'done', 'filename': fname, 'count': len(collected_items), 'docx_url': f'/api/docx-raw/{fname}', 'preview_url': f'/api/preview/{fname}', 'vacancies': cards_list, 'message': f'Muvaffaqiyatli saqlandi: {len(collected_items)} ta vakansiya'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"

    return Response(event_stream(), mimetype="text/event-stream")


@app.route("/api/stop", methods=["POST"])
def stop_scan():
    """Jarayonni to'xtatish"""
    global stop_flag
    stop_flag = True
    return jsonify({"status": "ok", "message": "To'xtatish so'rovi qabul qilindi"})


@app.route("/api/preview/<filename>")
def preview_document(filename):
    """
    Word (.docx) hujjatini Mammoth orqali chiroyli HTML ko'rinishida qaytaradi.
    """
    safe_name = os.path.basename(filename)
    filepath = os.path.join(EXPORTS_DIR, safe_name)
    if not os.path.exists(filepath):
        return jsonify({"error": "Fayl topilmadi"}), 404

    try:
        with open(filepath, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
            return jsonify({
                "status": "ok",
                "filename": safe_name,
                "html": html,
                "messages": [str(m) for m in result.messages]
            })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/docx-raw/<filename>")
def get_docx_raw(filename):
    """
    Word (.docx) faylini client-side docx-preview.js rendereri uchun jo'natadi.
    """
    safe_name = os.path.basename(filename)
    filepath = os.path.join(EXPORTS_DIR, safe_name)
    if not os.path.exists(filepath):
        return "Fayl topilmadi", 404
    return send_file(
        filepath,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=False
    )


@app.route("/api/download/<filename>")
def download_document(filename):
    """Word (.docx) faylini yuklab olish"""
    safe_name = os.path.basename(filename)
    return send_from_directory(EXPORTS_DIR, safe_name, as_attachment=True)


@app.route("/api/history")
def get_history():
    """Avval yaratilgan Word hujjatlari tarixi"""
    files = glob.glob(os.path.join(EXPORTS_DIR, "*.docx"))
    files.sort(key=os.path.getmtime, reverse=True)

    items = []
    for f in files:
        fname = os.path.basename(f)
        stat = os.stat(f)
        size_kb = round(stat.st_size / 1024, 1)
        mtime = datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M:%S")

        # Nomidan sonini topish
        m_cnt = re.search(r'_(\d+)_ta', fname)
        count_str = f"{m_cnt.group(1)} ta" if m_cnt else "Vakansiyalar to'plami"

        items.append({
            "filename": fname,
            "size": f"{size_kb} KB",
            "date": mtime,
            "count": count_str,
            "preview_url": f"/api/preview/{fname}",
            "docx_url": f"/api/docx-raw/{fname}",
            "download_url": f"/api/download/{fname}"
        })
    return jsonify(items)


@app.route("/api/vacancies-by-doc/<path:filename>")
def get_vacancies_by_doc(filename):
    """Word hujjatiga tegishli vakansiya kartalarini qaytaradi"""
    safe_name = os.path.basename(filename)
    meta_path = os.path.join(EXPORTS_DIR, safe_name + ".json")
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                return jsonify(json.load(f))
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    return jsonify([])


@app.route("/api/export-single/<int:vac_id>", methods=["POST"])
def export_single_vacancy_api(vac_id):
    """
    Bitta tanlangan vakansiyani Word (.docx) hujjati qilib eksport qiladi.
    """
    try:
        detail = api_client.get_vacancy_detail(vac_id)
        if not detail:
            return jsonify({"error": f"Vakansiya #{vac_id} topilmadi"}), 404
        
        filepath = docx_exporter.export_single_vacancy(detail)
        fname = os.path.basename(filepath)
        card_item = format_vacancy_card_data(detail)
        
        return jsonify({
            "status": "ok",
            "filename": fname,
            "count": 1,
            "docx_url": f"/api/docx-raw/{fname}",
            "preview_url": f"/api/preview/{fname}",
            "download_url": f"/api/download/{fname}",
            "card": card_item,
            "message": f"Vakansiya #{vac_id} Word hujjati yaratildi!"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/history/<filename>", methods=["DELETE"])
def delete_history_document(filename):
    """
    Foydalanuvchi talabi bo'yicha hujjatlarni o'chirish taqiqlangan (arxiv saqlanadi).
    """
    return jsonify({"error": "Word tarixi hujjatlarini o'chirish taqiqlangan. Barcha hujjatlar arxivda saqlanadi."}), 403



# --------------------------------------------------------------------------------------
# 6. DASTUR KIRISH NUQTASI (main)
# --------------------------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = os.environ.get("HOST", "0.0.0.0")
    is_cloud = bool(os.environ.get("K_SERVICE") or os.environ.get("GAE_INSTANCE") or os.environ.get("FUNCTION_TARGET"))

    print("=" * 65)
    print("  ARGOS.UZ PROFESSIONAL WEB ILOVASI VA WORD VIEWER")
    print("=" * 65)
    print(f"[*] Server manzili: http://127.0.0.1:{port} (Tashqi: http://0.0.0.0:{port})")
    print(f"[*] Hujjatlar saqlanadi: {EXPORTS_DIR}")
    if not is_cloud:
        print("[*] Brauzeringiz avtomatik tarzda ochilmoqda...")
        import webbrowser
        threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    print("-" * 65)

    # Flask serverni ishga tushirish (0.0.0.0 lokal va bulutda ishonchli ishlaydi)
    app.run(host=host, port=port, debug=False)

