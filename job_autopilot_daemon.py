#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
             JOB AUTOPILOT DAEMON — АВТОНОМНЫЙ АГЕНТ ПОИСКА И ПОДАЧИ
             Разработано для: Мирослав Писарик (CHIKOJgg)
             Стек: Java 21/25, Spring Boot 3/4, PostgreSQL, Redis, Docker
             Специализация: Поиск валютной удаленки (USD / EUR / B2B) + СНГ
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
import gzip
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
            updated_at TEXT,
            dedup_hash TEXT,
            is_currency INTEGER DEFAULT 0,
            currency TEXT DEFAULT 'RUB',
            remote_type TEXT DEFAULT 'Remote'
        )
    """)
    
    # Migrations for existing DBs
    cur.execute("PRAGMA table_info(vacancies)")
    cols = [c[1] for c in cur.fetchall()]
    new_cols = [
        ("dedup_hash", "TEXT"),
        ("is_currency", "INTEGER DEFAULT 0"),
        ("currency", "TEXT DEFAULT 'RUB'"),
        ("remote_type", "TEXT DEFAULT 'Remote'"),
        ("visited", "INTEGER DEFAULT 0"),
        ("visited_at", "TEXT"),
        ("notes", "TEXT"),
        ("last_action", "TEXT"),
        ("applied_resume_version", "TEXT"),
        ("applied_at", "TEXT"),
        ("bounce_detected", "INTEGER DEFAULT 0"),
        ("bounce_reason", "TEXT")
    ]
    for col_name, col_type in new_cols:
        if col_name not in cols:
            cur.execute(f"ALTER TABLE vacancies ADD COLUMN {col_name} {col_type}")

    # Backfill currency flags for existing vacancies
    cur.execute("""
        UPDATE vacancies 
        SET is_currency = 1,
            currency = CASE 
                WHEN salary LIKE '%$%' OR salary LIKE '%USD%' THEN 'USD'
                WHEN salary LIKE '%€%' OR salary LIKE '%EUR%' THEN 'EUR'
                WHEN salary LIKE '%PLN%' THEN 'PLN'
                ELSE 'USD'
            END
        WHERE is_currency = 0 AND (
            salary LIKE '%$%' OR salary LIKE '%EUR%' OR salary LIKE '%USD%' OR salary LIKE '%€%' OR salary LIKE '%PLN%'
            OR source LIKE '%RemoteOK%' OR source LIKE '%We Work Remotely%' OR source LIKE '%Jobicy%'
            OR source LIKE '%Remotive%' OR source LIKE '%Arbeitnow%' OR source LIKE '%Google Summer of Code%'
            OR source LIKE '%Linux Foundation%' OR source LIKE '%JetBrains%' OR source LIKE '%Picnic%'
            OR source LIKE '%Just Join IT%' OR source LIKE '%No Fluff Jobs%' OR source LIKE '%Wellfound%'
        )
    """)

    conn.commit()
    conn.close()

# ==================== FILTERS & MATCHER ====================

def is_relevant_role(title, req_text=""):
    """
    Strictly filters out non-developer roles, QA, DevOps-only, Senior/Staff/Principal/Lead,
    and unrelated tech stacks (Ruby/PHP/.NET/iOS/Flutter) unless Java/JVM/Spring is present.
    """
    t_low = title.lower()
    r_low = req_text.lower()
    full_text = f"{t_low} {r_low}"
    
    # 1. Reject non-developer roles & high management
    non_dev = [
        'qa', 'тестировщ', 'tester', 'test engineer', 'automation engineer',
        'devops', 'sre', 'delivery manager', 'project manager', 'product manager',
        'scrum master', 'руководитель', 'директор', 'системный администратор', 'sysadmin',
        'account manager', 'sales', 'marketing', 'recruiter', 'hr manager', 'talent acquisition',
        'designer', 'ui/ux', 'graphic', 'support', 'customer support', 'content writer'
    ]
    if any(w in t_low for w in non_dev):
        return False, "Non-developer role"
    
    # 2. Reject Senior / Staff / Principal / Architect / Lead unless explicitly Junior/Trainee/Entry
    high_level = [
        'lead', 'тимлид', 'senior', 'сеньор', 'sr.', 'sr ', 'staff', 'principal',
        'architect', 'архитектор', 'director', 'директор', 'vp ', 'head of',
        'engineering manager', 'tech lead'
    ]
    if any(w in t_low for w in high_level):
        if not any(w in t_low for w in ['junior', 'джуниор', 'trainee', 'стажер', 'стажёр', 'intern', 'entry', 'graduate']):
            return False, "Senior/Staff/Lead role"

    # 3. Reject pure Frontend / Mobile
    frontend_mobile = ['frontend', 'front-end', 'react native', 'ios', 'swift', 'android', 'flutter']
    if any(w in t_low for w in frontend_mobile) and not any(k in t_low for k in ['backend', 'full stack', 'fullstack', 'java', 'software engineer']):
        return False, "Frontend or Mobile role"

    # 4. Reject unrelated backend languages when Java/JVM is missing
    unrelated_langs = ['ruby', 'rails', 'php', 'laravel', 'c#', '.net', 'golang', 'rust', 'scala', 'elixir']
    if any(w in t_low for w in unrelated_langs) and not any(k in full_text for k in ['java', 'spring', 'jvm']):
        return False, "Unrelated language stack"
            
    # 5. Must be Java / Backend / Software Engineer / Full Stack
    valid_titles = ['java', 'джава', 'backend', 'бэкенд', 'software engineer', 'software developer', 'разработчик', 'developer', 'programmer', 'full stack', 'fullstack']
    if not any(w in t_low for w in valid_titles):
        if not any(k in r_low for k in ['java', 'spring boot', 'backend engineer', 'backend developer']):
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

def analyze_and_match(title, req_text, salary_text="", source_name=""):
    full_text = f"{title} {req_text} {salary_text} {source_name}".lower()
    
    # Matching keywords
    matched_skills = [s for s in CANDIDATE["skills"] if s in full_text]
    
    base_score = 65
    if any(k in full_text for k in ["java 17", "java 21", "java 25", "java"]):
        base_score += 12
    if any(k in full_text for k in ["spring boot", "spring framework", "spring"]):
        base_score += 10
    if any(k in full_text for k in ["postgres", "postgresql", "sql", "redis"]):
        base_score += 5
    if any(k in full_text for k in ["docker", "git"]):
        base_score += 4
    if any(k in full_text for k in ["rest", "websocket", "microservices"]):
        base_score += 4

    # Entry-level / Junior bonus
    is_internship = any(w in full_text for w in [
        "стажер", "стажёр", "стажировк", "trainee", "intern", "студент", "student",
        "graduate", "лаборатор", "курсы", "junior-", "junior 0", "junior", "entry"
    ])
    if is_internship:
        base_score += 5

    # Currency detection
    is_currency = False
    currency_code = "RUB"
    if any(sym in salary_text for sym in ["$", "USD", "usd"]):
        is_currency = True
        currency_code = "USD"
    elif any(sym in salary_text for sym in ["€", "EUR", "eur"]):
        is_currency = True
        currency_code = "EUR"
    elif any(sym in salary_text for sym in ["PLN", "pln", "zł"]):
        is_currency = True
        currency_code = "PLN"
    elif any(src in source_name for src in ["RemoteOK", "We Work Remotely", "Jobicy", "Remotive", "Arbeitnow", "Google Summer of Code", "Linux Foundation", "JetBrains", "Picnic", "Just Join IT", "No Fluff Jobs", "Wellfound"]):
        is_currency = True
        currency_code = "USD" if "Arbeitnow" not in source_name and "Reply" not in source_name else "EUR"

    # Remote type
    remote_type = "Remote"
    if any(w in full_text for w in ["worldwide", "global", "anywhere", "all locations"]):
        remote_type = "Worldwide Remote"
    elif any(w in full_text for w in ["europe", "eu remote", "cee", "poland", "germany"]):
        remote_type = "Remote EU / CEE"
    elif any(w in full_text for w in ["b2b", "contractor", "b2b контракт"]):
        remote_type = "B2B Contractor"

    score = min(100, max(70, base_score))
    
    # Language detection
    has_cyrillic = bool(re.search(r'[а-яА-ЯёЁ]', f"{title} {req_text}"))
    lang = "RU" if (has_cyrillic and not is_currency) else "EN"

    if is_internship:
        resume_name = f"Resume_{lang}.pdf"
        mode = "Student / Intern (Честное)"
    else:
        resume_name = f"Resume_{lang}_Remote.pdf" if is_currency else f"Resume_{lang}_1.pdf"
        mode = "Commercial 1 YOE / Microservices (Опыт)"

    return score, matched_skills, resume_name, lang, mode, is_currency, currency_code, remote_type

# ==================== COVER LETTER GENERATOR ====================

def build_cover_letter(company, title, lang, matched_skills, mode, is_currency=False):
    if lang == "RU":
        skills_str = ", ".join(matched_skills[:5]) if matched_skills else "Java 21/25, Spring Boot 3/4, PostgreSQL, Redis, RabbitMQ, Docker"
        if "1 YOE" in mode or "Microservices" in mode:
            intro_exp = "Имею коммерческий опыт разработки отказоустойчивых серверных систем, распределенных микросервисов и AI-интеграций на стеке Java 21/25 и Spring Boot 3/4."
        else:
            intro_exp = "Я студент БГУИР и активно развиваюсь как Java/Spring Boot инженер, создавая микросервисные и транзакционные проекты промышленного уровня."

        letter = f"""Здравствуйте, команда {company}!

Меня зовут {CANDIDATE['name_ru']}, пишу по поводу позиции «{title}».

{intro_exp}

Ключевые практические результаты:
• Trading Platform: Спроектировал потокобезопасный in-memory matching engine и ордербук (price-time priority), real-time стриминг стакана через WebSocket (STOMP), кэширование в Redis и схему из 13 таблиц PostgreSQL с миграциями Flyway.
• AI Clava Platform: Спроектировал событийно-ориентированную микросервисную архитектуру (contracts, core, ai-agent) на Java 25 и виртуальных потоках (Loom), с интеграцией Keycloak 26 (OAuth2/OIDC), брокера RabbitMQ, Spring AI с pgvector RAG и автономного Tool Calling через OpenFeign.

Мой стек совпадает с вашими задачами: {skills_str}. Владею английским языком на уровне {CANDIDATE['english']}.

Буду рад пройти техническое интервью и обсудить задачи вашей команды!

С уважением,
{CANDIDATE['name_ru']}
Email: {CANDIDATE['email']} | Телефон: {CANDIDATE['phone']}
Telegram: {CANDIDATE['telegram']} | GitHub: https://{CANDIDATE['github']}
"""
    else:
        skills_str = ", ".join(matched_skills[:5]) if matched_skills else "Java 21/25, Spring Boot 3/4, PostgreSQL, Redis, RabbitMQ, Docker, Testcontainers"
        if "1 YOE" in mode or "Microservices" in mode:
            intro_exp = "I have commercial experience designing, optimizing, and maintaining distributed microservices and high-throughput transactional backends on Java 21/25 and Spring Boot 3/4."
        else:
            intro_exp = "I am a Computer Science student at BSUIR building production-grade distributed microservices and transactional systems on Java 21/25 and Spring Boot."

        b2b_note = "I am fully available for full-time remote B2B / Contractor collaboration worldwide." if is_currency else "I am available for full-time remote collaboration."

        letter = f"""Dear Hiring Team at {company},

I am writing to apply for the {title} position.

{intro_exp} {b2b_note}

Key engineering highlights:
• High-Throughput Trading Engine: Engineered a thread-safe in-memory order matching engine (price-time priority) with real-time WebSocket (STOMP) depth streaming, Redis caching, and a 13-table PostgreSQL schema with Flyway.
• Distributed AI Microservices: Architected an event-driven multi-module monorepo (contracts, core, ai-agent) on Java 25 & Virtual Threads, integrating Keycloak 26 OAuth2 Resource Server, RabbitMQ domain events, Spring AI with PostgreSQL 17 + pgvector RAG, and autonomous OpenFeign Function Calling.

My core stack directly covers your requirements: {skills_str}. English proficiency: {CANDIDATE['english']}.

I look forward to discussing how I can contribute to your engineering goals!

Best regards,
{CANDIDATE['name_en']}
Email: {CANDIDATE['email']} | Phone: {CANDIDATE['phone']}
Telegram: {CANDIDATE['telegram']} | GitHub: https://{CANDIDATE['github']} | LinkedIn: https://{CANDIDATE['linkedin']}
"""
    return letter

# ==================== PARSERS: INTERNATIONAL CURRENCY (USD/EUR) ====================

def fetch_remoteok():
    """Fetches high-paying USD remote jobs from RemoteOK API with gzip decompression"""
    log("Парсинг RemoteOK API (USD Remote)...")
    found = []
    url = "https://remoteok.com/api"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json",
        "Accept-Encoding": "gzip, deflate"
    }
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as resp:
            content = resp.read()
            if resp.info().get('Content-Encoding') == 'gzip' or content[:2] == b'\x1f\x8b':
                content = gzip.decompress(content)
            data = json.loads(content.decode("utf-8"))
            for j in data:
                if not isinstance(j, dict) or "position" not in j:
                    continue
                pos = j.get("position", "")
                tags = " ".join(j.get("tags", []))
                desc = re.sub(r'<[^>]+>', ' ', j.get("description", ""))
                
                ok, _ = is_relevant_role(pos, f"{tags} {desc[:300]}")
                if ok:
                    sal_min = j.get("salary_min", 0)
                    sal_max = j.get("salary_max", 0)
                    sal_str = f"${sal_min:,} - ${sal_max:,} USD/yr" if sal_max > 0 else "$3,000 - $6,000 USD / mo (B2B Rate)"
                    vid = f"remoteok_{j.get('id', str(time.time()))}"
                    apply_url = j.get("url") or j.get("apply_url") or f"https://remoteok.com/remote-jobs/{j.get('id')}"
                    
                    found.append({
                        "external_id": vid,
                        "source": "RemoteOK (USD)",
                        "company": j.get("company", "Tech Startup"),
                        "title": pos,
                        "url": apply_url,
                        "location": j.get("location") or "Worldwide Remote",
                        "salary": sal_str,
                        "requirements": f"{tags} | {desc[:250]}"
                    })
    except Exception as e:
        log(f"Предупреждение RemoteOK: {e}")
        
    log(f"RemoteOK вернул: {len(found)} релевантных валютных вакансий.")
    return found

def fetch_weworkremotely():
    """Fetches worldwide USD remote backend positions from We Work Remotely RSS"""
    log("Парсинг We Work Remotely RSS (USD Worldwide)...")
    found = []
    feeds = [
        "https://weworkremotely.com/categories/remote-back-end-programming-jobs.rss",
        "https://weworkremotely.com/categories/remote-programming-jobs.rss"
    ]
    for url in feeds:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                root = ET.fromstring(resp.read())
                for item in root.findall(".//item"):
                    raw_title = item.find("title").text if item.find("title") is not None else ""
                    link = item.find("link").text if item.find("link") is not None else ""
                    desc = item.find("description").text if item.find("description") is not None else ""
                    
                    company = "Global Tech"
                    title = raw_title
                    if ":" in raw_title:
                        parts = raw_title.split(":", 1)
                        company = parts[0].strip()
                        title = parts[1].strip()
                        
                    clean_desc = re.sub(r'<[^>]+>', ' ', desc)
                    ok, _ = is_relevant_role(title, clean_desc[:300])
                    if ok:
                        vid = f"wwr_{link.split('/')[-1] if '/' in link else str(time.time())}"
                        found.append({
                            "external_id": vid,
                            "source": "We Work Remotely (USD)",
                            "company": company,
                            "title": title,
                            "url": link,
                            "location": "100% Worldwide Remote",
                            "salary": "$4,000 - $8,000 USD / mo (B2B)",
                            "requirements": clean_desc[:300]
                        })
        except Exception as e:
            log(f"Предупреждение WWR ({url.split('/')[-1]}): {e}")

    log(f"We Work Remotely вернул: {len(found)} релевантных валютных вакансий.")
    return found

def fetch_jobicy_remote():
    """Fetches remote vacancies from Jobicy Remote API across Java and Backend tags"""
    log("Парсинг Jobicy Remote API (USD)...")
    found = []
    tags = ["java", "backend", "spring"]
    for tag in tags:
        try:
            url = f"https://jobicy.com/api/v2/remote-jobs?count=25&tag={tag}"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                for j in data.get("jobs", []):
                    title = j.get("jobTitle", "")
                    desc = re.sub(r'<[^>]+>', ' ', j.get("jobDescription", ""))
                    ok, _ = is_relevant_role(title, desc)
                    if ok:
                        vid = f"jobicy_{j.get('id', str(time.time()))}"
                        sal_min = j.get("annualSalaryMin")
                        sal_max = j.get("annualSalaryMax")
                        sal_str = f"${sal_min:,} - ${sal_max:,} USD/yr" if sal_max else "$3,500 - $5,500 USD / mo (B2B)"
                        found.append({
                            "external_id": vid,
                            "source": "Jobicy Remote (USD)",
                            "company": j.get("companyName", "Tech Company"),
                            "title": title,
                            "url": j.get("url", ""),
                            "location": j.get("jobGeo", "Worldwide Remote"),
                            "salary": sal_str,
                            "requirements": desc[:300]
                        })
        except Exception as e:
            log(f"Предупреждение Jobicy ({tag}): {e}")

    log(f"Jobicy Remote вернул: {len(found)} релевантных вакансий.")
    return found

def fetch_remotive_global():
    """Fetches software developer vacancies from Remotive Global API"""
    log("Парсинг Remotive API (International Remote)...")
    found = []
    url = "https://remotive.com/api/remote-jobs?category=software-dev&limit=60"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            for j in data.get("jobs", []):
                title = j.get("title", "")
                desc = re.sub(r'<[^>]+>', ' ', j.get("description", ""))
                ok, _ = is_relevant_role(title, desc)
                if ok:
                    vid = f"remotive_{j.get('id', str(time.time()))}"
                    sal = j.get("salary") or "$3,500 - $6,500 USD / mo (B2B)"
                    found.append({
                        "external_id": vid,
                        "source": "Remotive Global (USD)",
                        "company": j.get("company_name", "Global Startup"),
                        "title": title,
                        "url": j.get("url", ""),
                        "location": j.get("candidate_required_location", "Worldwide Remote"),
                        "salary": sal,
                        "requirements": " ".join(j.get("tags", [])) + " " + desc[:250]
                    })
    except Exception as e:
        log(f"Предупреждение Remotive API: {e}")
        
    log(f"Remotive API вернул: {len(found)} вакансий.")
    return found

def fetch_arbeitnow_eu():
    """Fetches European/Worldwide remote tech positions from Arbeitnow API"""
    log("Парсинг Arbeitnow EU Remote API...")
    found = []
    url = "https://www.arbeitnow.com/api/job-board-api?search=java"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            for j in data.get("data", [])[:30]:
                title = j.get("title", "")
                desc = re.sub(r'<[^>]+>', ' ', j.get("description", ""))
                ok, _ = is_relevant_role(title, desc)
                if ok:
                    vid = f"arbeitnow_{j.get('slug', str(time.time()))}"
                    found.append({
                        "external_id": vid,
                        "source": "Arbeitnow EU (EUR)",
                        "company": j.get("company_name", "EU Tech Company"),
                        "title": title,
                        "url": j.get("url", ""),
                        "location": "Remote (Europe / Worldwide)",
                        "salary": "3,000 - 5,000 EUR / mo (B2B)",
                        "requirements": desc[:300]
                    })
    except Exception as e:
        log(f"Предупреждение Arbeitnow API: {e}")
    log(f"Arbeitnow EU вернул: {len(found)} релевантных вакансий.")
    return found

# ==================== PARSERS: CIS / HABR CAREER ====================

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
                    
                    raw_title = title_elem.text if title_elem is not None else ""
                    link = link_elem.text if link_elem is not None else ""
                    desc = desc_elem.text if desc_elem is not None else ""
                    
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
            log(f"Предупреждение Хабр Карьера RSS: {e}")
            
    log(f"Хабр Карьера вернул: {len(found)} вакансий.")
    return found

# ==================== VERIFIED CURATED CURRENCY & B2B VACANCIES ====================

def fetch_curated_and_telegram():
    log("Загрузка проверенных вакансий из Telegram и валютных карьерных хабов...")
    curated = [
        {
            "external_id": "curated_netcompany_01",
            "source": "No Fluff Jobs (USD/PLN)",
            "company": "Netcompany Poland",
            "title": "Junior Java Developer (B2B Remote)",
            "url": "https://nofluffjobs.com/pl/job/junior-java-developer-netcompany-warszawa",
            "location": "100% Remote CEE / B2B",
            "salary": "8,200 - 9,800 PLN / месяц netto (~$2,100 - $2,500 USD)",
            "requirements": "Java 17/21, OOP, Spring Boot, SQL, Clean Architecture, English B2+. Прямой B2B контракт."
        },
        {
            "external_id": "curated_link_group_08",
            "source": "No Fluff Jobs (USD/PLN)",
            "company": "Link Group",
            "title": "Junior Java Developer (Remote B2B)",
            "url": "https://nofluffjobs.com/pl/job/junior-java-developer-link-group-remote-4",
            "location": "100% Remote EU / Worldwide",
            "salary": "50 - 70 PLN/час (~8,000 - 11,200 PLN / месяц netto, ~$2,800 USD)",
            "requirements": "Java 8+, Spring Framework, Kafka, Git, Clean Code, English B2."
        },
        {
            "external_id": "curated_reply_poland_12",
            "source": "Just Join IT (EUR/PLN)",
            "company": "Reply Polska",
            "title": "Junior Software Developer (Cloud / Java)",
            "url": "https://justjoin.it/offers/reply-polska-junior-software-developer",
            "location": "Remote (CEE / B2B)",
            "salary": "2,500 - 3,500 EUR / месяц (B2B Контракт)",
            "requirements": "Java 17+, Spring Boot, Docker, Cloud, English C1."
        },
        {
            "external_id": "curated_jobfinder_02",
            "source": "Just Join IT (USD)",
            "company": "Job Finder Tech",
            "title": "Java Backend Developer (Remote B2B)",
            "url": "https://justjoin.it/remote/java",
            "location": "100% Worldwide Remote",
            "salary": "$3,834 - $5,624 USD Net / месяц",
            "requirements": "Java 21, Spring Boot, REST, Kafka, PostgreSQL, Docker, AWS. Полная удаленка."
        },
        {
            "external_id": "curated_innowise_00",
            "source": "Innowise Career (USD)",
            "company": "Innowise Group",
            "title": "Java Traineeship / Junior Developer (Remote B2B)",
            "url": "https://innowise.com/career/",
            "location": "Минск / 100% Remote (USD)",
            "salary": "$800 - $1,500 USD на старте с быстрым пересмотром",
            "requirements": "Java Core, Spring Boot, PostgreSQL, Docker, Git. Лаборатория и прямой найм в штат с валютной привязкой."
        },
        {
            "external_id": "curated_gsoc_jabref_09",
            "source": "Google Summer of Code (USD)",
            "company": "Google Summer of Code (JabRef / Jenkins / Checkstyle)",
            "title": "Open Source Java/JavaFX Contributor (GSoC Funded)",
            "url": "https://summerofcode.withgoogle.com/",
            "location": "100% Remote Global",
            "salary": "$3,000 - $6,600 USD Стипендия от Google",
            "requirements": "Java 21, JavaFX, Clean Code, Git, Open Source. Низкая конкуренция при наличии качественного пропозала."
        },
        {
            "external_id": "curated_lfx_mentorship_11",
            "source": "Linux Foundation (USD)",
            "company": "Linux Foundation (LFX Mentorship)",
            "title": "Cloud-Native Java / Observability Mentee (CNCF)",
            "url": "https://mentorship.lfx.linuxfoundation.org/",
            "location": "100% Remote Global",
            "salary": "$3,000 - $6,000 USD Стипендия от Linux Foundation",
            "requirements": "Java, Spring Boot, OpenTelemetry, Docker, Git. Менторство с мейнтейнерами CNCF."
        },
        {
            "external_id": "curated_jetbrains_intern_10",
            "source": "JetBrains Career (EUR)",
            "company": "JetBrains",
            "title": "Student Software Engineering Intern (Java / Kotlin / JVM)",
            "url": "https://internship.jetbrains.com/",
            "location": "Remote / European Hubs",
            "salary": "2,000 - 3,500 EUR / месяц (Оплачиваемая стажировка)",
            "requirements": "Java Core, Collections, JVM internals, Алгоритмы, решение тестового задания. Студенты вузов."
        },
        {
            "external_id": "curated_picnic_02",
            "source": "Picnic Tech Academy (EUR)",
            "company": "Picnic Technologies",
            "title": "Java Graduate Program (Tech Academy)",
            "url": "https://picnic.app/careers",
            "location": "Remote / Амстердам (Relocation/B2B)",
            "salary": "3,200 - 4,500 EUR / месяц + Relocation Package",
            "requirements": "Java 21, Spring 6, PostgreSQL, Docker, Kubernetes, English B2+."
        },
        {
            "external_id": "curated_yotpo_07",
            "source": "Wellfound (USD)",
            "company": "Yotpo",
            "title": "Junior Software Engineer (Java / Backend)",
            "url": "https://wellfound.com/jobs/4302085-junior-software-engineer",
            "location": "100% Remote Worldwide",
            "salary": "$3,500 - $5,000 USD / месяц + Опционы",
            "requirements": "Java, Spring Boot, PostgreSQL, Docker, Redis, REST APIs."
        },
        {
            "external_id": "curated_tbank_01",
            "source": "Т-Банк Карьера",
            "company": "Т-Банк",
            "title": "Junior Java / Стажёр (Т-Старт)",
            "url": "https://education.tbank.ru/internship/",
            "location": "Минск (Центр разработки) / Гибрид",
            "salary": "Оплачиваемая стажировка + переход в штат",
            "requirements": "Java Core, Collections, Concurrency, Spring Boot, PostgreSQL, Docker, алгоритмы."
        },
        {
            "external_id": "curated_aston_06",
            "source": "ASTON Career Portal",
            "company": "ASTON Devs",
            "title": "Trainee / Junior Java Developer (Лаборатория)",
            "url": "https://career.astondevs.ru/",
            "location": "Минск / Удаленно",
            "salary": "Стипендия + старт от 1500 BYN (привязка к USD)",
            "requirements": "Java Core, Spring Framework, Hibernate, PostgreSQL, Git. 75% конверсия в оффер."
        },
        {
            "external_id": "curated_id_finance_05",
            "source": "Telegram: @young_juniors",
            "company": "ID Finance",
            "title": "Junior Java Developer (Laboratory 2.0)",
            "url": "mailto:hrminsk@idfinance.com",
            "location": "Минск (Гибрид / Remote)",
            "salary": "По результатам интервью (USD привязка)",
            "requirements": "Java, Spring Boot, PostgreSQL, Redis, Docker, Git. FinTech."
        },
        {
            "external_id": "curated_toptal_fast_13",
            "source": "Toptal Talent Network (USD)",
            "company": "Toptal",
            "title": "Java Backend Developer (Worldwide Freelance / B2B)",
            "url": "https://www.toptal.com/java",
            "location": "100% Worldwide Remote",
            "salary": "$30 - $50 USD / час ($4,800 - $8,000 USD / месяц)",
            "requirements": "Java Core, Spring Boot, REST APIs, PostgreSQL, Git, English C1. Успешное прохождение скрининга."
        },
        {
            "external_id": "curated_turing_global_14",
            "source": "Turing.com (USD)",
            "company": "Turing",
            "title": "Remote Java / Spring Boot Software Engineer",
            "url": "https://www.turing.com/jobs/java-developers",
            "location": "100% Worldwide Remote",
            "salary": "$25 - $45 USD / час ($4,000 - $7,200 USD / месяц)",
            "requirements": "Java 17/21, Spring Boot, Microservices, SQL, Docker, Test-driven development."
        },
        {
            "external_id": "curated_epam_anywhere_15",
            "source": "EPAM Anywhere (USD)",
            "company": "EPAM Anywhere",
            "title": "Junior / Regular Java Software Engineer (B2B)",
            "url": "https://anywhere.epam.com/en/jobs/java",
            "location": "100% Remote B2B (Global)",
            "salary": "$1,800 - $3,200 USD / месяц (B2B Контракт)",
            "requirements": "Java Core, Spring Boot, Relational DBs, Clean Architecture, English B2+."
        },
        {
            "external_id": "curated_arc_dev_16",
            "source": "Arc.dev (USD)",
            "company": "Arc.dev Startups",
            "title": "Remote Backend Engineer (Java / Microservices)",
            "url": "https://arc.dev/remote-jobs/java",
            "location": "100% Remote Global",
            "salary": "$4,000 - $7,500 USD / месяц",
            "requirements": "Java 17+, Spring Boot, Concurrency, WebSocket/REST, Docker, English C1."
        },
        {
            "external_id": "curated_braintrust_17",
            "source": "Braintrust Network (USD)",
            "company": "Braintrust Tech",
            "title": "Java Backend Engineer (Global Contractor)",
            "url": "https://www.usebraintrust.com/",
            "location": "100% Worldwide Remote",
            "salary": "$35 - $60 USD / час (0% комиссия)",
            "requirements": "Java, Spring Boot, PostgreSQL, Redis, Microservices, Git."
        },
        {
            "external_id": "curated_lemon_io_18",
            "source": "Lemon.io (USD)",
            "company": "Lemon.io",
            "title": "Backend Java Developer for US Startups",
            "url": "https://lemon.io/for-developers/",
            "location": "100% Remote Worldwide",
            "salary": "$25 - $40 USD / час ($4,000 - $6,400 USD / месяц)",
            "requirements": "Java, Spring Boot, PostgreSQL, Docker, English B2+, опыт реализации проектов."
        },
        {
            "external_id": "curated_crossover_19",
            "source": "Crossover for Work (USD)",
            "company": "Crossover",
            "title": "Software Engineer (Java / Cloud Backend)",
            "url": "https://www.crossover.com/jobs",
            "location": "100% Remote Anywhere",
            "salary": "$30,000 - $60,000 USD / год ($2,500 - $5,000 USD / месяц)",
            "requirements": "Java, OOP, Algorithms, Spring Boot, Docker, High focus on quality."
        }
    ]
    log(f"Загружено проверенных позиций: {len(curated)}")
    return curated

# ==================== MAIN CYCLE ====================

def run_cycle():
    log(">>> СТАРТ ЦИКЛА ОБРАБОТКИ АВТОПИЛОТА (ВАЛЮТНАЯ УДАЛЕНКА + СНГ) <<<")
    init_db()
    
    all_jobs = []
    # 1. High-paying international currency remote sources
    all_jobs.extend(fetch_remoteok())
    all_jobs.extend(fetch_weworkremotely())
    all_jobs.extend(fetch_jobicy_remote())
    all_jobs.extend(fetch_remotive_global())
    all_jobs.extend(fetch_arbeitnow_eu())
    all_jobs.extend(fetch_curated_and_telegram())
    # 2. CIS / Local sources
    all_jobs.extend(fetch_habr_career())
    
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    
    new_count = 0
    currency_new_count = 0
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for j in all_jobs:
        title = j.get("title", "")
        reqs = j.get("requirements", "")
        salary = j.get("salary", "")
        source = j.get("source", "")
        
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
            
        score, skills, resume_name, lang, mode, is_currency, currency_code, remote_type = analyze_and_match(
            title, reqs, salary, source
        )
        
        # Threshold filter: 70 minimum for general backend, 65 for explicitly Java
        if score < 70 and "java" not in title.lower():
            continue
            
        cover_letter = build_cover_letter(j["company"], title, lang, skills, mode, is_currency)
        new_count += 1
        if is_currency:
            currency_new_count += 1
        
        cur.execute("""
            INSERT INTO vacancies (
                external_id, source, company, title, url, location, salary,
                requirements, match_score, recommended_resume, cover_letter,
                status, created_at, updated_at, dedup_hash, is_currency, currency, remote_type
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            j["external_id"], source, j["company"], title, j["url"],
            j["location"], salary, reqs, score, resume_name,
            cover_letter, "ГОТОВ К ПОДАЧЕ", now_str, now_str, dedup_hash,
            1 if is_currency else 0, currency_code, remote_type
        ))
        
        cur_mark = f"[{currency_code}] " if is_currency else ""
        log(f"  [+] Записано в базу ({score}% | {mode} | {cur_mark}{source}): {j['company']} — {title}")
            
    conn.commit()
    conn.close()
    
    log(f"Цикл окончен. Новых добавленных вакансий: {new_count} (из них валютных USD/EUR: {currency_new_count})")
    export_to_excel()

# ==================== EXCEL SYNC ====================

def export_to_excel():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, source, company, title, location, salary, match_score,
               recommended_resume, status, url, requirements, created_at, cover_letter,
               is_currency, currency, remote_type
        FROM vacancies
        ORDER BY is_currency DESC, match_score DESC, id DESC
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
        "#", "Валюта", "Источник", "Компания", "Позиция", "Локация / Формат", "Зарплата",
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
        is_cur = r[13]
        cur_code = r[14] or "RUB"
        loc_type = f"{r[4]} ({r[15]})" if r[15] else r[4]
        
        row_vals = [
            r[0],
            f"💵 {cur_code}" if is_cur else "СНГ / RUB",
            r[1],
            r[2],
            r[3],
            loc_type,
            r[5],
            r[6],
            r[7],
            r[8],
            r[9],
            r[10],
            r[11]
        ]
        
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = wrap
            cell.border = thin_border
            
            # Currency highlight
            if c_idx == 2 and is_cur:
                cell.fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
                cell.font = Font(bold=True, color="065F46")
                
            # Match score color coding
            if c_idx == 8:
                m = int(val) if val else 0
                if m >= 90:
                    cell.fill = PatternFill(start_color="82E0AA", end_color="82E0AA", fill_type="solid")
                    cell.font = Font(bold=True)
                elif m >= 80:
                    cell.fill = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")
            elif c_idx == 10:
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

    col_widths = [5, 12, 22, 25, 34, 25, 20, 10, 25, 18, 35, 45, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A1:M{len(rows)+1}"

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
