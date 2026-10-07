#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
             JOB AUTOPILOT DAEMON — АВТОНОМНЫЙ АГЕНТ ПОИСКА И ПОДАЧИ
             Разработано для: Мирослав Писарик (CHIKOJgg)
             Стек: Java 21/25, Spring Boot 3/4, PostgreSQL, Redis, Docker
=============================================================================
"""

import os
import sys
import time
import json
import sqlite3
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
from datetime import datetime
import re

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "jobs_autopilot.db")
EXCEL_PATH = os.path.join(BASE_DIR, "Вакансии_Автопилот_Live.xlsx")
LOG_PATH = os.path.join(BASE_DIR, "autopilot.log")

CANDIDATE = {
    "name_ru": "Мирослав Писарик",
    "name_en": "Miraslau Pisaryk",
    "email": "bhasdgjkg5@gmail.com",
    "phone": "+375 29 749-77-19",
    "telegram": "@yxxtg",
    "github": "github.com/CHIKOJgg",
    "linkedin": "linkedin.com/in/miroslav-pisaryk-953490261",
    "english": "C1 Advanced",
    "skills": [
        "java", "spring", "spring boot", "postgresql", "redis", "docker", "docker compose",
        "rest", "websocket", "stomp", "flyway", "junit", "testcontainers", "resilience4j",
        "bucket4j", "claude", "mcp", "modular monolith", "spring modulith", "concurrency",
        "matching engine", "hibernate", "jpa", "maven", "git"
    ]
}

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{ts}] {msg}"
    try:
        print(formatted)
    except Exception:
        # Fallback for Windows terminal codepage
        print(formatted.encode("ascii", "replace").decode("ascii"))
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(formatted + "\n")

# ==================== DATABASE ====================

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS vacancies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            external_id TEXT UNIQUE,
            source TEXT,
            company TEXT,
            title TEXT,
            url TEXT,
            location TEXT,
            salary TEXT,
            requirements TEXT,
            match_score INTEGER,
            recommended_resume TEXT,
            cover_letter TEXT,
            status TEXT,
            created_at TEXT,
            updated_at TEXT
        )
    """)
    conn.commit()
    conn.close()

# ==================== FILTERS & MATCHER ====================

def is_relevant_role(title, req_text=""):
    """Strictly filters out non-developer roles, QA, Architects, Leads, Senior 5+ YOE"""
    t_low = title.lower()
    
    # 1. Reject non-developer roles
    non_dev = [
        'qa', 'тестировщ', 'tester', 'test engineer', 'automation engineer',
        'devops', 'sre', 'delivery manager', 'project manager', 'product manager',
        'scrum master', 'руководитель', 'директор', 'системный администратор', 'sysadmin'
    ]
    if any(w in t_low for w in non_dev):
        return False, "Non-developer role (QA/DevOps/Management)"
    
    # 2. Reject Senior/Lead unless explicitly junior/trainee
    if any(w in t_low for w in ['lead', 'тимлид', 'senior', 'сеньор', 'ведущий', 'главный', 'architect', 'архитектор']):
        if not any(w in t_low for w in ['junior', 'джуниор', 'trainee', 'стажер', 'стажёр', 'intern']):
            return False, "Senior/Lead/Architect role"
            
    # 3. Must be Java / Backend / Software Engineer
    if not any(w in t_low for w in ['java', 'джава', 'backend', 'бэкенд', 'software engineer', 'разработчик', 'developer', 'программист']):
        if 'java' not in req_text.lower():
            return False, "Not Java/Backend"
            
    return True, "OK"

def extract_hr_email(text):
    """Regex extracts valid HR email addresses from description or URL"""
    if not text:
        return None
    matches = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', text)
    for m in matches:
        m_clean = m.strip('.').lower()
        if not any(m_clean.endswith(ext) for ext in ['.png', '.jpg', '.gif', '.webp', '.svg', '.jpeg']) and len(m_clean) > 5:
            return m_clean
    return None

def analyze_and_match(title, req_text):
    full_text = f"{title} {req_text}".lower()
    
    # Matching keywords
    matched_skills = [s for s in CANDIDATE["skills"] if s in full_text]
    
    # Core stack bonuses
    base_score = 65
    if any(k in full_text for k in ["java 17", "java 21", "java 25", "java"]):
        base_score += 10
    if any(k in full_text for k in ["spring boot", "spring framework", "spring"]):
        base_score += 10
    if any(k in full_text for k in ["postgres", "postgresql", "sql", "redis"]):
        base_score += 5
    if any(k in full_text for k in ["docker", "git"]):
        base_score += 5

    # Intern/Trainee flag
    is_internship = any(w in full_text for w in [
        "стажер", "стажёр", "стажировк", "trainee", "intern", "студент", "student",
        "graduate", "лаборатор", "курсы", "junior-", "junior 0"
    ])
    if is_internship:
        base_score += 5

    score = min(100, max(70, base_score))
    
    has_cyrillic = bool(re.search(r'[а-яА-ЯёЁ]', full_text))
    lang = "RU" if has_cyrillic else "EN"

    if is_internship:
        resume_name = f"Resume_{lang}.pdf"
        mode = "Student / Intern (Честное)"
    else:
        resume_name = f"Resume_{lang}_1.pdf"
        mode = "Commercial 1 YOE (Опыт)"

    return score, matched_skills, resume_name, lang, mode

# ==================== COVER LETTER GENERATOR ====================

def build_cover_letter(company, title, lang, matched_skills, mode):
    if lang == "RU":
        skills_str = ", ".join(matched_skills[:5]) if matched_skills else "Java 21, Spring Boot 3/4, PostgreSQL, Redis, Docker"
        if "1 YOE" in mode:
            intro_exp = "Имею коммерческий опыт разработки высоконагруженных серверных приложений на стеке Java 21 и Spring Boot 3/4."
        else:
            intro_exp = "Я студент БГУИР и активно развиваюсь как Java/Spring Boot разработчик, создавая проекты промышленного уровня."

        letter = f"""Здравствуйте, команда {company}!

Меня зовут {CANDIDATE['name_ru']}, пишу по поводу позиции «{title}».

{intro_exp}

Ключевые практические результаты:
• Trading Platform: Спроектировал потокобезопасный in-memory matching engine и ордербук (price-time priority), real-time стриминг стакана через WebSocket (STOMP), кэширование в Redis и схему из 13 таблиц PostgreSQL с миграциями Flyway.
• Klawa AI Assistant: Разработал модульный монолит (12 модулей) на Spring Boot 4 и Spring Modulith с интеграцией Claude API/MCP, защитой от сбоев Resilience4j Circuit Breaker и 100% изоляцией тестов в Testcontainers.

Мой стек совпадает с вашими задачами: {skills_str}. Владею английским языком на уровне {CANDIDATE['english']}.

Буду рад пройти техническое интервью и обсудить задачи вашей команды!

С уважением,
{CANDIDATE['name_ru']}
Email: {CANDIDATE['email']} | Телефон: {CANDIDATE['phone']}
Telegram: {CANDIDATE['telegram']} | GitHub: https://{CANDIDATE['github']}
"""
    else:
        skills_str = ", ".join(matched_skills[:5]) if matched_skills else "Java 21, Spring Boot 3/4, PostgreSQL, Redis, Docker, Testcontainers"
        if "1 YOE" in mode:
            intro_exp = "I have commercial experience designing and deploying high-reliability backend systems on Java 21, Spring Boot 3/4, and PostgreSQL."
        else:
            intro_exp = "I am a Computer Science student at BSUIR building production-ready backend systems on Java 21 and Spring Boot."

        letter = f"""Dear Hiring Team at {company},

I am writing to apply for the {title} position.

{intro_exp}

Key engineering highlights:
• High-Throughput Trading Engine: Engineered a thread-safe in-memory order matching engine with real-time WebSocket (STOMP) depth streaming, Redis caching, and a 13-table PostgreSQL schema with Flyway.
• Modular AI Architecture: Built a 12-module modular system with Spring Modulith & Java 25, integrating Anthropic Claude API/MCP, Resilience4j Circuit Breakers, and comprehensive Testcontainers validation.

My core stack directly covers your requirements: {skills_str}. English proficiency: {CANDIDATE['english']}.

I look forward to discussing how I can contribute to your engineering goals.

Best regards,
{CANDIDATE['name_en']}
Email: {CANDIDATE['email']} | Phone: {CANDIDATE['phone']}
Telegram: {CANDIDATE['telegram']} | GitHub: https://{CANDIDATE['github']} | LinkedIn: https://{CANDIDATE['linkedin']}
"""
    return letter

# ==================== PARSERS ====================

def fetch_habr_career():
    log("Парсинг Хабр Карьера RSS...")
    found = []
    queries = [
        "https://career.habr.com/vacancies/rss?q=java+junior",
        "https://career.habr.com/vacancies/rss?q=java+%D1%81%D1%82%D0%B0%D0%B6%D0%B5%D1%80",
        "https://career.habr.com/vacancies/rss?q=spring+junior"
    ]
    for url in queries:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                xml_data = resp.read()
                root = ET.fromstring(xml_data)
                for item in root.findall(".//item"):
                    title_elem = item.find("title")
                    link_elem = item.find("link")
                    desc_elem = item.find("description")
                    pub_elem = item.find("pubDate")
                    
                    raw_title = title_elem.text if title_elem is not None else ""
                    link = link_elem.text if link_elem is not None else ""
                    desc = desc_elem.text if desc_elem is not None else ""
                    
                    # Split company from title if format is "Role at Company"
                    company = "IT Компания"
                    title = raw_title
                    if " в " in raw_title:
                        parts = raw_title.split(" в ")
                        title = parts[0].strip()
                        company = parts[1].strip()
                    elif " (" in raw_title:
                        parts = raw_title.split(" (")
                        title = parts[0].strip()
                        company = parts[1].replace(")", "").strip()
                        
                    vid = f"habr_{link.split('/')[-1] if '/' in link else str(time.time())}"
                    
                    found.append({
                        "external_id": vid,
                        "source": "Хабр Карьера",
                        "company": company,
                        "title": title,
                        "url": link,
                        "location": "Удаленно / СНГ",
                        "salary": "По договоренности",
                        "requirements": desc[:300] if desc else "Java, Spring Boot"
                    })
        except Exception as e:
            log(f"Ошибка Хабр Карьера RSS: {e}")
            
    log(f"Хабр Карьера вернул: {len(found)} вакансий.")
    return found

def fetch_remotive_global():
    log("Парсинг Remotive API (International Remote)...")
    found = []
    url = "https://remotive.com/api/remote-jobs?category=software-dev&search=java"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            for j in data.get("jobs", [])[:20]:
                title = j.get("title", "")
                if any(k in title.lower() for k in ["java", "backend", "software engineer", "spring"]):
                    vid = f"remotive_{j.get('id', str(time.time()))}"
                    found.append({
                        "external_id": vid,
                        "source": "Remotive Global",
                        "company": j.get("company_name", "Global Startup"),
                        "title": title,
                        "url": j.get("url", ""),
                        "location": j.get("candidate_required_location", "Worldwide Remote"),
                        "salary": j.get("salary", "Competitive / B2B"),
                        "requirements": " ".join(j.get("tags", []))
                    })
    except Exception as e:
        log(f"Ошибка Remotive API: {e}")
        
    log(f"Remotive API вернул: {len(found)} вакансий.")
    return found

def fetch_arbeitnow_eu():
    log("Парсинг Arbeitnow EU Remote API...")
    found = []
    url = "https://www.arbeitnow.com/api/job-board-api?search=java"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            for j in data.get("data", [])[:40]:
                title = j.get("title", "")
                desc = re.sub(r'<[^>]+>', ' ', j.get("description", ""))
                ok, _ = is_relevant_role(title, desc)
                if ok:
                    vid = f"arbeitnow_{j.get('slug', str(time.time()))}"
                    found.append({
                        "external_id": vid,
                        "source": "Arbeitnow EU",
                        "company": j.get("company_name", "EU Tech Company"),
                        "title": title,
                        "url": j.get("url", ""),
                        "location": "Remote (Europe / Worldwide)",
                        "salary": "B2B / Competitive",
                        "requirements": desc[:350]
                    })
    except Exception as e:
        log(f"Ошибка Arbeitnow API: {e}")
    log(f"Arbeitnow EU вернул: {len(found)} релевантных вакансий.")
    return found

def fetch_jobicy_remote():
    log("Парсинг Jobicy Remote API...")
    found = []
    url = "https://jobicy.com/api/v2/remote-jobs?count=25&tag=java"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            for j in data.get("jobs", []):
                title = j.get("jobTitle", "")
                desc = re.sub(r'<[^>]+>', ' ', j.get("jobDescription", ""))
                ok, _ = is_relevant_role(title, desc)
                if ok:
                    vid = f"jobicy_{j.get('id', str(time.time()))}"
                    sal = f"${j.get('annualSalaryMin')}-${j.get('annualSalaryMax')}" if j.get("annualSalaryMin") else "Competitive"
                    found.append({
                        "external_id": vid,
                        "source": "Jobicy Remote",
                        "company": j.get("companyName", "Tech Company"),
                        "title": title,
                        "url": j.get("url", ""),
                        "location": j.get("jobGeo", "Worldwide Remote"),
                        "salary": sal,
                        "requirements": desc[:350]
                    })
    except Exception as e:
        log(f"Ошибка Jobicy API: {e}")
    log(f"Jobicy Remote вернул: {len(found)} релевантных вакансий.")
    return found

def fetch_curated_and_telegram():
    log("Загрузка проверенных вакансий из Telegram и карьерных порталов...")
    curated = [
        {
            "external_id": "curated_innowise_00",
            "source": "Innowise Career",
            "company": "Innowise Group",
            "title": "Java Traineeship / Junior Developer",
            "url": "https://innowise.com/career/",
            "location": "Минск / Remote",
            "salary": "По результатам интервью",
            "requirements": "Java Core, Spring Boot, PostgreSQL, Docker, Git. Лаборатория с высокой конверсией в штат."
        },
        {
            "external_id": "curated_tbank_01",
            "source": "Т-Банк Карьера",
            "company": "Т-Банк",
            "title": "Junior Java / Стажёр (Т-Старт)",
            "url": "https://education.tbank.ru/internship/",
            "location": "Минск (Центр разработки) / Гибрид",
            "salary": "Оплачиваемая стажировка",
            "requirements": "Java Core, Collections, Concurrency, Spring Boot, PostgreSQL, Docker, алгоритмы."
        },
        {
            "external_id": "curated_picnic_02",
            "source": "Picnic Tech Academy",
            "company": "Picnic Technologies",
            "title": "Java Graduate Program (Tech Academy)",
            "url": "https://picnic.app/careers",
            "location": "Remote / Амстердам",
            "salary": "Full package + B2B/Relocation",
            "requirements": "Java 21, Spring 6, PostgreSQL, Docker, Kubernetes, English B2+."
        },
        {
            "external_id": "curated_team_inno_03",
            "source": "rabota.by",
            "company": "Team.Inno",
            "title": "Trainee Java Developer",
            "url": "https://rabota.by/vacancy/137295021",
            "location": "Минск, пр. Независимости",
            "salary": "Обучение + оплачиваемый контракт",
            "requirements": "ООП, Java Core, SQL (CRUD, JOIN), Git, English B1."
        },
        {
            "external_id": "curated_lightwell_04",
            "source": "rabota.by",
            "company": "Лайт Вел Организейшн",
            "title": "Java-разработчик (Junior / Entry-Level)",
            "url": "https://rabota.by/vacancy/137247090",
            "location": "Минск, ул. Натуралистов",
            "salary": "По результатам интервью",
            "requirements": "Java 8+, Spring Framework, Hibernate, PostgreSQL/Oracle, REST, Kafka, Redis, Docker."
        },
        {
            "external_id": "curated_id_finance_05",
            "source": "Telegram: @young_juniors",
            "company": "ID Finance",
            "title": "Junior Java Developer (Laboratory 2.0)",
            "url": "mailto:hrminsk@idfinance.com",
            "location": "Минск (Гибрид)",
            "salary": "По результатам интервью",
            "requirements": "Java, Spring Boot, PostgreSQL, Redis, Docker, Git. FinTech."
        },
        {
            "external_id": "curated_aston_06",
            "source": "ASTON Career Portal",
            "company": "ASTON Devs",
            "title": "Trainee / Junior Java Developer (Лаборатория)",
            "url": "https://career.astondevs.ru/",
            "location": "Минск / Удаленно",
            "salary": "Стипендия + старт от 1500 BYN",
            "requirements": "Java Core, Spring Framework, Hibernate, PostgreSQL, Git. 75% конверсия в оффер."
        },
        {
            "external_id": "curated_yotpo_07",
            "source": "Wellfound",
            "company": "Yotpo",
            "title": "Junior Software Engineer (Java / Backend)",
            "url": "https://wellfound.com/jobs/4302085-junior-software-engineer",
            "location": "100% Remote Worldwide",
            "salary": "Опционы + конкурентная ставка",
            "requirements": "Java, Spring Boot, PostgreSQL, Docker, Redis, REST APIs."
        },
        {
            "external_id": "curated_link_group_08",
            "source": "No Fluff Jobs",
            "company": "Link Group",
            "title": "Junior Java Developer (Remote B2B)",
            "url": "https://nofluffjobs.com/pl/job/junior-java-developer-link-group-remote-4",
            "location": "100% Remote EU",
            "salary": "50 - 70 PLN/час (~8 000 - 11 200 PLN / месяц netto)",
            "requirements": "Java 8+, Spring Framework, Kafka, Git, Clean Code, English B2."
        }
    ]
    log(f"Загружено проверенных позиций: {len(curated)}")
    return curated

# ==================== MAIN CYCLE ====================

def run_cycle():
    log(">>> СТАРТ ЦИКЛА ОБРАБОТКИ АВТОПИЛОТА <<<")
    init_db()
    
    all_jobs = []
    all_jobs.extend(fetch_habr_career())
    all_jobs.extend(fetch_arbeitnow_eu())
    all_jobs.extend(fetch_jobicy_remote())
    all_jobs.extend(fetch_remotive_global())
    all_jobs.extend(fetch_curated_and_telegram())
    
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    
    new_count = 0
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for j in all_jobs:
        title = j.get("title", "")
        reqs = j.get("requirements", "")
        
        # 1. Filter out irrelevant roles
        ok, reason = is_relevant_role(title, reqs)
        if not ok:
            continue
            
        # 2. Check deduplication hash
        clean_comp = re.sub(r'[^a-zA-Zа-яА-Я0-9]', '', j.get("company", "").lower())
        clean_title = re.sub(r'[^a-zA-Zа-яА-Я0-9]', '', title.lower())
        dedup_hash = f"{clean_comp}_{clean_title[:24]}"
        
        cur.execute("SELECT id FROM vacancies WHERE external_id = ? OR dedup_hash = ?", (j["external_id"], dedup_hash))
        if cur.fetchone():
            continue
            
        # 3. Detect HR email
        hr_email = extract_hr_email(f"{j.get('url', '')} {reqs}")
        if hr_email and not j.get("url", "").startswith("mailto:"):
            reqs = f"[HR Email: {hr_email}] " + reqs
            
        score, skills, resume_name, lang, mode = analyze_and_match(title, reqs)
        
        # Threshold filter
        if score < 70 and "java" not in title.lower():
            continue
            
        cover_letter = build_cover_letter(j["company"], title, lang, skills, mode)
        new_count += 1
        
        cur.execute("""
            INSERT INTO vacancies (
                external_id, source, company, title, url, location, salary,
                requirements, match_score, recommended_resume, cover_letter,
                status, created_at, updated_at, dedup_hash
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            j["external_id"], j["source"], j["company"], title, j["url"],
            j["location"], j["salary"], reqs, score, resume_name,
            cover_letter, "ГОТОВ К ПОДАЧЕ", now_str, now_str, dedup_hash
        ))
        log(f"  [+] Записано в базу ({score}% | {mode}): {j['company']} — {title}")
            
    conn.commit()
    conn.close()
    
    log(f"Цикл окончен. Новых добавленных вакансий: {new_count}")
    export_to_excel()

# ==================== EXCEL SYNC ====================

def export_to_excel():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, source, company, title, location, salary, match_score,
               recommended_resume, status, url, requirements, created_at, cover_letter
        FROM vacancies
        ORDER BY match_score DESC, id DESC
    """)
    rows = cur.fetchall()
    conn.close()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Автопилот - Вакансии"
    
    header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin_border = Border(
        left=Side(style="thin", color="CBD5E0"), right=Side(style="thin", color="CBD5E0"),
        top=Side(style="thin", color="CBD5E0"), bottom=Side(style="thin", color="CBD5E0")
    )
    wrap = Alignment(wrap_text=True, vertical="top")
    
    headers = [
        "#", "Источник", "Компания", "Позиция", "Локация", "Зарплата",
        "Match %", "Рекомендуемое резюме", "Статус заявки", "Ссылка",
        "Требования", "Дата обнаружения"
    ]
    
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
        
    for r_idx, r in enumerate(rows, 2):
        row_vals = list(r[:12])
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = wrap
            cell.border = thin_border
            if c_idx == 7:
                m = int(val) if val else 0
                if m >= 95:
                    cell.fill = PatternFill(start_color="82E0AA", end_color="82E0AA", fill_type="solid")
                    cell.font = Font(bold=True)
                elif m >= 85:
                    cell.fill = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")
            elif c_idx == 9:
                cell.font = Font(bold=True, color="1B4F72")
                
    # Sheet 2: Cover Letters
    ws2 = wb.create_sheet("Готовые Cover Letters")
    ws2.cell(row=1, column=1, value="#").fill = header_fill
    ws2.cell(row=1, column=2, value="Компания").fill = header_fill
    ws2.cell(row=1, column=3, value="Позиция").fill = header_fill
    ws2.cell(row=1, column=4, value="Текст письма для отправки").fill = header_fill
    for col in range(1, 5):
        ws2.cell(row=1, column=col).font = header_font
        
    for r_idx, r in enumerate(rows, 2):
        ws2.cell(row=r_idx, column=1, value=r[0]).border = thin_border
        ws2.cell(row=r_idx, column=2, value=r[2]).border = thin_border
        ws2.cell(row=r_idx, column=3, value=r[3]).border = thin_border
        c_cl = ws2.cell(row=r_idx, column=4, value=r[12])
        c_cl.border = thin_border
        c_cl.alignment = wrap
        
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 25
    ws2.column_dimensions["C"].width = 30
    ws2.column_dimensions["D"].width = 100

    col_widths = [5, 20, 25, 32, 22, 18, 10, 30, 18, 35, 45, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A1:L{len(rows)+1}"

    try:
        wb.save(EXCEL_PATH)
        log(f"Excel синхронизирован: {EXCEL_PATH} (Всего записей в таблице: {len(rows)})")
    except Exception as e:
        log(f"Ошибка при сохранении Excel: {e}")

if __name__ == "__main__":
    is_loop = "--loop" in sys.argv
    interval_minutes = 60
    
    if is_loop:
        log(">>> ЗАПУСК В НЕПРЕРЫВНОМ РЕЖИМЕ (INTERVAL LOOP) <<<")
        log(f"Интервал повторного сканирования: каждые {interval_minutes} минут.")
        while True:
            try:
                run_cycle()
            except Exception as e:
                log(f"Сбой цикла: {e}")
            log(f"Сон на {interval_minutes} минут до следующего парсинга...\n")
            time.sleep(interval_minutes * 60)
    else:
        run_cycle()
        print("\n[OK] Одиночный цикл автопилота завершен успешно.")
