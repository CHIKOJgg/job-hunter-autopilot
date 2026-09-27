#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
         DASHBOARD & VACANCY TRACKER v2.0 — ПАНЕЛЬ УПРАВЛЕНИЯ И ТРЕКЕР
=============================================================================
"""

import os
import sys
import sqlite3
import json
import subprocess
from datetime import datetime
from typing import Optional

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from fastapi import FastAPI, HTTPException, Body
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "jobs_autopilot.db")
EXCEL_PATH = os.path.join(BASE_DIR, "Вакансии_Автопилот_Live.xlsx")

app = FastAPI(title="Job Hunter & Application Tracker")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db_extra():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(vacancies)")
    cols = [c[1] for c in cur.fetchall()]
    new_cols = [
        ("visited", "INTEGER DEFAULT 0"),
        ("visited_at", "TEXT"),
        ("notes", "TEXT"),
        ("last_action", "TEXT")
    ]
    for col_name, col_type in new_cols:
        if col_name not in cols:
            cur.execute(f"ALTER TABLE vacancies ADD COLUMN {col_name} {col_type}")
    conn.commit()
    conn.close()

init_db_extra()

# ==================== API ENDPOINTS ====================

@app.get("/api/stats")
def get_stats():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM vacancies")
    total = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM vacancies WHERE status LIKE '%ОТПРАВЛЕНО%'")
    applied = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM vacancies WHERE status NOT LIKE '%ОТПРАВЛЕНО%' AND status NOT LIKE '%ОТКАЗ%' AND status NOT LIKE '%АРХИВ%'")
    to_apply = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM vacancies WHERE visited = 1")
    visited_count = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM vacancies WHERE status LIKE '%ИНТЕРВЬЮ%' OR status LIKE '%ПРИГЛАШЕНИЕ%'")
    interviews = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM vacancies WHERE status LIKE '%ТЕСТОВОЕ%'")
    test_tasks = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM vacancies WHERE status LIKE '%ОТКАЗ%'")
    rejections = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM vacancies WHERE match_score >= 85")
    top_matches = cur.fetchone()[0]

    conn.close()

    conversion_rate = round((interviews / applied * 100), 1) if applied > 0 else 0.0

    return {
        "total": total,
        "to_apply": to_apply,
        "applied": applied,
        "visited_count": visited_count,
        "interviews": interviews,
        "test_tasks": test_tasks,
        "rejections": rejections,
        "top_matches": top_matches,
        "conversion_rate": conversion_rate
    }

@app.get("/api/vacancies")
def get_vacancies(tab: str = "to_apply", search: str = "", stack_filter: str = ""):
    conn = get_db()
    cur = conn.cursor()

    query = "SELECT * FROM vacancies WHERE 1=1"
    params = []

    if tab == "to_apply":
        query += " AND status NOT LIKE '%ОТПРАВЛЕНО%' AND status NOT LIKE '%ОТКАЗ%' AND status NOT LIKE '%АРХИВ%'"
    elif tab == "applied":
        query += " AND status LIKE '%ОТПРАВЛЕНО%'"
    elif tab == "visited":
        query += " AND visited = 1"
    elif tab == "dialogs":
        query += " AND (status LIKE '%ИНТЕРВЬЮ%' OR status LIKE '%ПРИГЛАШЕНИЕ%' OR status LIKE '%ТЕСТОВОЕ%')"
    elif tab == "top":
        query += " AND match_score >= 85"
    elif tab == "all":
        pass

    if search:
        query += " AND (company LIKE ? OR title LIKE ? OR location LIKE ? OR requirements LIKE ?)"
        s_wild = f"%{search}%"
        params.extend([s_wild, s_wild, s_wild, s_wild])

    if stack_filter:
        query += " AND (requirements LIKE ? OR title LIKE ?)"
        st_wild = f"%{stack_filter}%"
        params.extend([st_wild, st_wild])

    query += " ORDER BY match_score DESC, id DESC"
    cur.execute(query, params)
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@app.post("/api/vacancies/{vid}/status")
def update_status(vid: int, payload: dict = Body(...)):
    new_status = payload.get("status")
    if not new_status:
        raise HTTPException(status_code=400, detail="Missing status")
    
    conn = get_db()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute("UPDATE vacancies SET status = ?, updated_at = ?, last_action = ? WHERE id = ?", 
                (new_status, now_str, f"Статус изменен на: {new_status}", vid))
    conn.commit()
    conn.close()
    return {"success": True, "status": new_status, "updated_at": now_str}

@app.post("/api/vacancies/{vid}/visited")
def mark_visited(vid: int):
    conn = get_db()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute("UPDATE vacancies SET visited = 1, visited_at = ? WHERE id = ?", (now_str, vid))
    conn.commit()
    conn.close()
    return {"success": True, "visited": 1, "visited_at": now_str}

@app.post("/api/vacancies/{vid}/notes")
def save_notes(vid: int, payload: dict = Body(...)):
    notes = payload.get("notes", "")
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE vacancies SET notes = ? WHERE id = ?", (notes, vid))
    conn.commit()
    conn.close()
    return {"success": True, "notes": notes}

@app.post("/api/action/apply_email/{vid}")
def apply_email(vid: int):
    from email_sender import send_email_application
    ok, msg = send_email_application(vid)
    return {"success": ok, "message": msg}

@app.post("/api/action/run_scan")
def run_scan():
    daemon_script = os.path.join(BASE_DIR, "job_autopilot_daemon.py")
    subprocess.Popen([sys.executable, daemon_script])
    return {"success": True, "message": "Автопилот сканирования запущен в фоне!"}

@app.post("/api/action/open_excel")
def open_excel():
    if os.path.exists(EXCEL_PATH):
        os.system(f'start "" "{EXCEL_PATH}"')
        return {"success": True, "message": "Excel таблица открыта."}
    return {"success": False, "message": "Файл Excel не найден."}

@app.get("/download_resume/{name}")
def download_resume(name: str):
    safe_name = os.path.basename(name)
    path = os.path.join(BASE_DIR, safe_name)
    if os.path.exists(path):
        return FileResponse(path, filename=safe_name)
    raise HTTPException(status_code=404, detail="Файл не найден")

# ==================== FRONTEND UI HTML ====================

@app.get("/", response_class=HTMLResponse)
def index():
    html_content = '''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Job Hunter & Tracker — Мирослав Писарик</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
<style>
    :root {
        --bg-main: #0B0F19;
        --bg-card: #131B2E;
        --bg-card-hover: #19233C;
        --bg-visited: #111827;
        --border: #1E293B;
        --border-active: #38BDF8;
        --text-primary: #F8FAFC;
        --text-secondary: #94A3B8;
        --accent-blue: #38BDF8;
        --accent-purple: #818CF8;
        --accent-green: #34D399;
        --accent-red: #F87171;
        --accent-yellow: #FBBF24;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', -apple-system, sans-serif; }
    body { background-color: var(--bg-main); color: var(--text-primary); padding: 20px 24px; min-height: 100vh; }
    .container { max-width: 1440px; margin: 0 auto; }

    /* Top Alert for Bounces */
    .bounce-banner {
        background: linear-gradient(90deg, rgba(248, 113, 113, 0.15) 0%, rgba(251, 191, 36, 0.15) 100%);
        border: 1px solid rgba(248, 113, 113, 0.4);
        border-radius: 12px;
        padding: 12px 18px;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
    }
    .bounce-banner-content { font-size: 13.5px; line-height: 1.4; color: #FCA5A5; }
    .bounce-banner-content strong { color: #FEF08A; }
    .bounce-btn { background: #DC2626; color: white; padding: 6px 14px; border-radius: 8px; text-decoration: none; font-size: 12px; font-weight: 600; white-space: nowrap; border: none; cursor: pointer; }

    /* Header */
    .header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: linear-gradient(135deg, #131B2E 0%, #1A233A 100%);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 20px 28px;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    }
    .profile-info h1 { font-size: 22px; font-weight: 800; display: flex; align-items: center; gap: 10px; }
    .profile-info p { color: var(--text-secondary); font-size: 13px; margin-top: 4px; }
    .badge-status {
        background: rgba(52, 211, 153, 0.15);
        color: var(--accent-green);
        border: 1px solid rgba(52, 211, 153, 0.3);
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
    }
    .resume-buttons { display: flex; gap: 8px; flex-wrap: wrap; }
    .btn {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid var(--border);
        color: var(--text-primary);
        padding: 8px 14px;
        border-radius: 8px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        text-decoration: none;
        transition: all 0.2s ease;
    }
    .btn:hover { background: rgba(255, 255, 255, 0.12); border-color: var(--text-secondary); transform: translateY(-1px); }
    .btn-primary { background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%); border-color: #3B82F6; color: white; }
    .btn-primary:hover { background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%); }
    .btn-success { background: rgba(52, 211, 153, 0.15); color: var(--accent-green); border-color: rgba(52, 211, 153, 0.3); }

    /* Stats Grid */
    .stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 12px;
        margin-bottom: 20px;
    }
    .stat-card {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 16px 20px;
        position: relative;
        overflow: hidden;
        cursor: pointer;
        transition: all 0.2s ease;
    }
    .stat-card:hover { transform: translateY(-2px); border-color: var(--accent-blue); }
    .stat-card.active { border-color: var(--accent-blue); background: #162038; box-shadow: 0 0 15px rgba(56, 189, 248, 0.2); }
    .stat-title { font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-secondary); font-weight: 600; }
    .stat-val { font-size: 26px; font-weight: 800; margin: 4px 0 2px; }
    .stat-sub { font-size: 11px; color: var(--text-secondary); }

    /* Main Section Controls */
    .controls-panel {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 16px 20px;
        margin-bottom: 20px;
        display: flex;
        flex-direction: column;
        gap: 12px;
    }
    .tabs-row {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        border-bottom: 1px solid var(--border);
        padding-bottom: 12px;
    }
    .tab-btn {
        background: transparent;
        border: 1px solid transparent;
        color: var(--text-secondary);
        padding: 8px 16px;
        border-radius: 8px;
        font-size: 13px;
        font-weight: 600;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        gap: 8px;
        transition: all 0.2s;
    }
    .tab-btn:hover { color: var(--text-primary); background: rgba(255,255,255,0.05); }
    .tab-btn.active {
        color: white;
        background: #2563EB;
        border-color: #3B82F6;
    }
    .tab-count {
        background: rgba(0,0,0,0.3);
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
    }

    .filter-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 14px;
        flex-wrap: wrap;
    }
    .search-box {
        flex: 1;
        min-width: 260px;
        position: relative;
    }
    .search-input {
        width: 100%;
        background: rgba(0, 0, 0, 0.25);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 9px 14px 9px 36px;
        color: white;
        font-size: 13px;
        outline: none;
    }
    .search-input:focus { border-color: var(--accent-blue); }
    .search-icon { position: absolute; left: 12px; top: 11px; font-size: 13px; color: var(--text-secondary); }

    .quick-chips { display: flex; gap: 6px; flex-wrap: wrap; align-items: center; }
    .chip {
        background: rgba(255,255,255,0.05);
        border: 1px solid var(--border);
        padding: 5px 10px;
        border-radius: 16px;
        font-size: 11px;
        cursor: pointer;
        color: var(--text-secondary);
        transition: all 0.15s;
    }
    .chip:hover, .chip.active { color: var(--accent-blue); border-color: var(--accent-blue); background: rgba(56, 189, 248, 0.1); }

    /* Vacancy Cards List */
    .vacancy-list { display: flex; flex-direction: column; gap: 12px; }
    .vac-card {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 18px 22px;
        transition: all 0.2s ease;
        position: relative;
    }
    .vac-card:hover { border-color: rgba(56, 189, 248, 0.5); background: var(--bg-card-hover); }
    .vac-card.is-visited {
        border-left: 4px solid var(--accent-purple);
    }
    .vac-card.is-applied {
        border-left: 4px solid var(--accent-green);
        background: rgba(52, 211, 153, 0.03);
    }
    .vac-card.is-failed {
        border-left: 4px solid var(--accent-red);
    }

    .vac-header {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 16px;
        margin-bottom: 8px;
    }
    .vac-title-group h3 {
        font-size: 16px;
        font-weight: 700;
        color: var(--text-primary);
        display: flex;
        align-items: center;
        gap: 10px;
        flex-wrap: wrap;
    }
    .company-name { color: var(--accent-blue); font-weight: 600; font-size: 14px; }
    .vac-meta {
        display: flex;
        align-items: center;
        gap: 14px;
        font-size: 12px;
        color: var(--text-secondary);
        margin-top: 4px;
        flex-wrap: wrap;
    }

    /* Badges */
    .match-badge {
        font-weight: 800;
        font-size: 12px;
        padding: 4px 10px;
        border-radius: 8px;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }
    .match-high { background: rgba(52, 211, 153, 0.2); color: var(--accent-green); border: 1px solid rgba(52, 211, 153, 0.4); }
    .match-med { background: rgba(56, 189, 248, 0.2); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.4); }
    .visited-badge {
        background: rgba(129, 140, 248, 0.15);
        color: var(--accent-purple);
        border: 1px solid rgba(129, 140, 248, 0.3);
        padding: 2px 8px;
        border-radius: 6px;
        font-size: 11px;
        font-weight: 600;
    }

    .vac-desc {
        font-size: 12.5px;
        color: #CBD5E1;
        line-height: 1.45;
        margin: 10px 0;
        background: rgba(0, 0, 0, 0.2);
        padding: 10px 14px;
        border-radius: 8px;
        border-left: 2px solid var(--border);
    }

    .vac-actions-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 12px;
        margin-top: 12px;
        padding-top: 12px;
        border-top: 1px solid rgba(255,255,255,0.06);
        flex-wrap: wrap;
    }
    .action-left { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .action-right { display: flex; align-items: center; gap: 8px; }

    /* Status Selector Dropdown */
    .status-select {
        background: #1E293B;
        color: white;
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 7px 12px;
        font-size: 12px;
        font-weight: 600;
        outline: none;
        cursor: pointer;
    }
    .status-select:focus { border-color: var(--accent-blue); }

    /* Modal */
    .modal-overlay {
        position: fixed;
        top: 0; left: 0; right: 0; bottom: 0;
        background: rgba(0, 0, 0, 0.75);
        display: none;
        align-items: center;
        justify-content: center;
        z-index: 1000;
        backdrop-filter: blur(4px);
    }
    .modal-box {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 16px;
        width: 90%;
        max-width: 650px;
        max-height: 85vh;
        overflow-y: auto;
        padding: 24px;
        box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
    }
    .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; }
    .modal-close { background: none; border: none; font-size: 20px; color: var(--text-secondary); cursor: pointer; }
    .letter-content {
        background: #0B0F19;
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 16px;
        font-family: 'JetBrains Mono', monospace;
        font-size: 12.5px;
        line-height: 1.5;
        white-space: pre-wrap;
        color: #E2E8F0;
        max-height: 380px;
        overflow-y: auto;
        margin-bottom: 16px;
    }

    /* Toast Notification */
    .toast {
        position: fixed;
        bottom: 24px;
        right: 24px;
        background: #1E293B;
        border: 1px solid var(--accent-green);
        color: white;
        padding: 12px 20px;
        border-radius: 10px;
        font-size: 13px;
        font-weight: 600;
        box-shadow: 0 10px 25px rgba(0,0,0,0.5);
        display: none;
        align-items: center;
        gap: 8px;
        z-index: 9999;
        animation: fadeIn 0.2s ease;
    }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
</style>
</head>
<body>

<div class="container">
    <!-- Notice for Bounced / Direct link vacancies -->
    <div class="bounce-banner">
        <div class="bounce-banner-content">
            ⚠️ <strong>Внимание по почте:</strong> Отклики в <strong>ASTON Devs</strong> не дошли по почте, так как адрес <code>cv@astondevs.ru</code> закрыт. В компании Aston подача переведена на сайт: <strong>career.astondevs.ru</strong>. По ID Finance отклики успешно доставлены!
        </div>
        <a class="bounce-btn" href="https://career.astondevs.ru/" target="_blank">Открыть форму Aston</a>
    </div>

    <!-- Header -->
    <header class="header">
        <div class="profile-info">
            <h1>Мирослав Писарик <span class="badge-status">Автопилот активен</span></h1>
            <p>Junior Java / Spring Boot Backend Engineer &bull; Минск &bull; C1 Advanced &bull; BSUIR (2024–2028)</p>
        </div>
        <div class="resume-buttons">
            <a class="btn" href="/download_resume/Resume_RU.pdf" target="_blank" title="Студенческое / Честное на русском">📄 Резюме RU (Честное)</a>
            <a class="btn" href="/download_resume/Resume_EN.pdf" target="_blank" title="Student / Honest in English">📄 Resume EN (Honest)</a>
            <a class="btn btn-primary" href="/download_resume/Resume_RU_1.pdf" target="_blank" title="1+ год коммерческого опыта на русском">⭐ Резюме RU 1 (Опыт)</a>
            <a class="btn btn-primary" href="/download_resume/Resume_EN_1.pdf" target="_blank" title="1+ YOE Commercial in English">⭐ Resume EN 1 (1 YOE)</a>
        </div>
    </header>

    <!-- Stats Grid -->
    <div class="stats-grid">
        <div class="stat-card active" onclick="switchTab('to_apply')">
            <div class="stat-title">🔥 Надо откликнуться</div>
            <div class="stat-val" id="stat-to-apply" style="color: var(--accent-yellow)">-</div>
            <div class="stat-sub">Высокий Match Score</div>
        </div>
        <div class="stat-card" onclick="switchTab('applied')">
            <div class="stat-title">📬 Уже отправлено</div>
            <div class="stat-val" id="stat-applied" style="color: var(--accent-green)">-</div>
            <div class="stat-sub">Успешные отклики</div>
        </div>
        <div class="stat-card" onclick="switchTab('visited')">
            <div class="stat-title">👁️ Где я кликал</div>
            <div class="stat-val" id="stat-visited" style="color: var(--accent-purple)">-</div>
            <div class="stat-sub">Просмотренные мной</div>
        </div>
        <div class="stat-card" onclick="switchTab('dialogs')">
            <div class="stat-title">🎯 Интервью / Тесты</div>
            <div class="stat-val" id="stat-interviews" style="color: var(--accent-blue)">-</div>
            <div class="stat-sub">Активные отклики</div>
        </div>
        <div class="stat-card" onclick="switchTab('all')">
            <div class="stat-title">📋 Всего в базе</div>
            <div class="stat-val" id="stat-total">-</div>
            <div class="stat-sub">Хабр, Telegram, Резерв</div>
        </div>
    </div>

    <!-- Controls Panel -->
    <div class="controls-panel">
        <div class="tabs-row">
            <button class="tab-btn active" id="tab-to_apply" onclick="switchTab('to_apply')">
                🔥 Нужно откликнуться <span class="tab-count" id="count-to_apply">0</span>
            </button>
            <button class="tab-btn" id="tab-applied" onclick="switchTab('applied')">
                📬 Уже отправлено <span class="tab-count" id="count-applied">0</span>
            </button>
            <button class="tab-btn" id="tab-visited" onclick="switchTab('visited')">
                👁️ Где я был (Кликнутые) <span class="tab-count" id="count-visited">0</span>
            </button>
            <button class="tab-btn" id="tab-dialogs" onclick="switchTab('dialogs')">
                🎯 Интервью и Тестовые <span class="tab-count" id="count-dialogs">0</span>
            </button>
            <button class="tab-btn" id="tab-all" onclick="switchTab('all')">
                📋 Все вакансии <span class="tab-count" id="count-all">0</span>
            </button>
        </div>

        <div class="filter-row">
            <div class="search-box">
                <span class="search-icon">🔍</span>
                <input type="text" class="search-input" id="search-input" placeholder="Поиск по компании, позиции, стеку..." oninput="handleSearch()">
            </div>
            <div class="quick-chips">
                <span style="font-size: 11px; color: var(--text-secondary); margin-right: 4px;">Фильтры:</span>
                <div class="chip" onclick="toggleChip(this, 'Java 21')">Java 21</div>
                <div class="chip" onclick="toggleChip(this, 'Spring Boot')">Spring Boot</div>
                <div class="chip" onclick="toggleChip(this, 'Remote')">Remote</div>
                <div class="chip" onclick="toggleChip(this, 'Минск')">Минск</div>
                <div class="chip" onclick="toggleChip(this, 'Docker')">Docker</div>
            </div>
            <div style="display: flex; gap: 8px;">
                <button class="btn btn-primary" onclick="triggerScan()">⚡ Сканировать новое</button>
                <button class="btn" onclick="openExcel()">📊 Открыть Excel</button>
            </div>
        </div>
    </div>

    <!-- Vacancy List -->
    <div class="vacancy-list" id="vacancies-container">
        <div style="text-align: center; padding: 40px; color: var(--text-secondary);">Загрузка вакансий...</div>
    </div>
</div>

<!-- Modal Cover Letter -->
<div class="modal-overlay" id="cover-modal" onclick="closeModal(event)">
    <div class="modal-box" onclick="event.stopPropagation()">
        <div class="modal-header">
            <h3 id="modal-title" style="font-size: 16px;">Сопроводительное письмо</h3>
            <button class="modal-close" onclick="closeModal()">&times;</button>
        </div>
        <div class="letter-content" id="modal-content"></div>
        <div style="display: flex; justify-content: flex-end; gap: 10px;">
            <button class="btn btn-primary" onclick="copyModalLetter()">📋 Скопировать текст письма</button>
            <button class="btn" onclick="closeModal()">Закрыть</button>
        </div>
    </div>
</div>

<!-- Toast -->
<div class="toast" id="toast-box">
    <span id="toast-icon">✓</span>
    <span id="toast-text">Готово!</span>
</div>

<script>
    let currentTab = 'to_apply';
    let currentSearch = '';
    let currentChip = '';
    let vacanciesData = [];

    async function loadStats() {
        try {
            const res = await fetch('/api/stats');
            const data = await res.json();
            document.getElementById('stat-to-apply').innerText = data.to_apply;
            document.getElementById('stat-applied').innerText = data.applied;
            document.getElementById('stat-visited').innerText = data.visited_count;
            document.getElementById('stat-interviews').innerText = data.interviews + data.test_tasks;
            document.getElementById('stat-total').innerText = data.total;

            document.getElementById('count-to_apply').innerText = data.to_apply;
            document.getElementById('count-applied').innerText = data.applied;
            document.getElementById('count-visited').innerText = data.visited_count;
            document.getElementById('count-dialogs').innerText = data.interviews + data.test_tasks;
            document.getElementById('count-all').innerText = data.total;
        } catch(e) {
            console.error('Stats error:', e);
        }
    }

    async function loadVacancies() {
        const container = document.getElementById('vacancies-container');
        try {
            let url = `/api/vacancies?tab=${currentTab}&search=${encodeURIComponent(currentSearch)}&stack_filter=${encodeURIComponent(currentChip)}`;
            const res = await fetch(url);
            vacanciesData = await res.json();

            if (vacanciesData.length === 0) {
                container.innerHTML = '<div style="text-align: center; padding: 60px; color: var(--text-secondary); background: var(--bg-card); border-radius: 12px;">В этом разделе пока нет вакансий по вашему фильтру.</div>';
                return;
            }

            container.innerHTML = vacanciesData.map(v => renderCard(v)).join('');
        } catch(e) {
            container.innerHTML = `<div style="text-align: center; color: var(--accent-red); padding: 40px;">Ошибка загрузки вакансий: ${e}</div>`;
        }
    }

    function renderCard(v) {
        const isVisited = v.visited === 1;
        const isApplied = (v.status || '').includes('ОТПРАВЛЕНО');
        const isFailed = (v.status || '').includes('не активен');
        const matchClass = v.match_score >= 85 ? 'match-high' : 'match-med';

        let cardClass = 'vac-card';
        if (isApplied) cardClass += ' is-applied';
        else if (isVisited) cardClass += ' is-visited';
        if (isFailed) cardClass += ' is-failed';

        const safeUrl = v.url || '#';
        const resFile = v.recommended_resume || 'Resume_RU_1.pdf';

        return `
        <div class="${cardClass}" id="card-${v.id}">
            <div class="vac-header">
                <div class="vac-title-group">
                    <span class="match-badge ${matchClass}">${v.match_score}% MATCH</span>
                    <h3>${v.title}</h3>
                    <span class="company-name">@ ${v.company}</span>
                    ${isVisited ? `<span class="visited-badge">👁️ Посещено (${v.visited_at ? v.visited_at.split(' ')[0] : 'ранее'})</span>` : ''}
                </div>
                <div style="display: flex; align-items: center; gap: 8px;">
                    <select class="status-select" onchange="changeStatus(${v.id}, this.value)">
                        <option value="ГОТОВ К ПОДАЧЕ" ${v.status === 'ГОТОВ К ПОДАЧЕ' ? 'selected' : ''}>🟡 Готов к подаче</option>
                        <option value="ОТПРАВЛЕНО (EMAIL)" ${v.status.includes('ОТПРАВЛЕНО') ? 'selected' : ''}>🟢 Отправлено</option>
                        <option value="ИНТЕРВЬЮ" ${v.status.includes('ИНТЕРВЬЮ') ? 'selected' : ''}>🟣 Приглашение / Интервью</option>
                        <option value="ТЕСТОВОЕ ЗАДАНИЕ" ${v.status.includes('ТЕСТОВОЕ') ? 'selected' : ''}>🟠 Тестовое задание</option>
                        <option value="ПОДАТЬ НА САЙТЕ (Email не активен)" ${v.status.includes('не активен') ? 'selected' : ''}>⚠️ Подать на сайте</option>
                        <option value="ОТКАЗ" ${v.status.includes('ОТКАЗ') ? 'selected' : ''}>🔴 Отказ</option>
                        <option value="АРХИВ" ${v.status.includes('АРХИВ') ? 'selected' : ''}>⚪ В архив</option>
                    </select>
                </div>
            </div>

            <div class="vac-meta">
                <span>📍 ${v.location || 'Удаленно / Remote'}</span>
                <span>💰 ${v.salary || 'Не указана'}</span>
                <span>🗓️ Источник: ${v.source || 'Хабр / Live'}</span>
                <span>📄 Резюме: <strong>${resFile}</strong></span>
            </div>

            <div class="vac-desc">
                ${v.requirements || 'Java, Spring Boot, PostgreSQL, Docker, Git'}
            </div>

            <div class="vac-actions-bar">
                <div class="action-left">
                    <a class="btn btn-primary" href="${safeUrl}" target="_blank" onclick="trackClick(${v.id})">
                        🔗 Открыть вакансию ↗
                    </a>
                    <button class="btn" onclick="openCoverModal(${v.id})">
                        📝 Сопроводительное
                    </button>
                    <button class="btn" onclick="copyLetterById(${v.id})">
                        📋 Скопировать письмо
                    </button>
                    <a class="btn" href="/download_resume/${resFile}" target="_blank" download>
                        💾 Скачать ${resFile}
                    </a>
                </div>
                <div class="action-right">
                    ${safeUrl.includes('@') || safeUrl.includes('mailto:') ? `
                        <button class="btn btn-success" onclick="sendDirectEmail(${v.id})">
                            ✉️ Отправить Email
                        </button>
                    ` : ''}
                </div>
            </div>
        </div>
        `;
    }

    async function trackClick(vid) {
        try {
            await fetch(`/api/vacancies/${vid}/visited`, { method: 'POST' });
            loadStats();
            // visually mark as visited
            const card = document.getElementById(`card-${vid}`);
            if (card && !card.classList.contains('is-applied')) {
                card.classList.add('is-visited');
            }
        } catch(e) {
            console.error('Track error:', e);
        }
    }

    async function changeStatus(vid, newStatus) {
        try {
            const res = await fetch(`/api/vacancies/${vid}/status`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ status: newStatus })
            });
            const data = await res.json();
            showToast(`Статус обновлен: ${newStatus}`);
            loadStats();
            loadVacancies();
        } catch(e) {
            showToast(`Ошибка: ${e}`);
        }
    }

    async function sendDirectEmail(vid) {
        showToast('Отправка email с резюме...');
        try {
            const res = await fetch(`/api/action/apply_email/${vid}`, { method: 'POST' });
            const data = await res.json();
            if (data.success) {
                showToast(`✓ Успешно отправлено!`);
                loadStats();
                loadVacancies();
            } else {
                showToast(`Ошибка: ${data.message}`);
            }
        } catch(e) {
            showToast(`Ошибка отправки: ${e}`);
        }
    }

    function switchTab(tabName) {
        currentTab = tabName;
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.stat-card').forEach(card => card.classList.remove('active'));

        const tabBtn = document.getElementById(`tab-${tabName}`);
        if (tabBtn) tabBtn.classList.add('active');

        loadVacancies();
    }

    let searchTimer = null;
    function handleSearch() {
        clearTimeout(searchTimer);
        searchTimer = setTimeout(() => {
            currentSearch = document.getElementById('search-input').value.trim();
            loadVacancies();
        }, 300);
    }

    function toggleChip(el, value) {
        if (currentChip === value) {
            currentChip = '';
            el.classList.remove('active');
        } else {
            document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
            currentChip = value;
            el.classList.add('active');
        }
        loadVacancies();
    }

    let currentLetterText = '';
    function openCoverModal(vid) {
        const v = vacanciesData.find(item => item.id === vid);
        if (!v) return;
        document.getElementById('modal-title').innerText = `Сопроводительное письмо — ${v.company}`;
        currentLetterText = v.cover_letter || 'Текст не сформирован.';
        document.getElementById('modal-content').innerText = currentLetterText;
        document.getElementById('cover-modal').style.display = 'flex';
    }

    function closeModal(e) {
        document.getElementById('cover-modal').style.display = 'none';
    }

    function copyModalLetter() {
        navigator.clipboard.writeText(currentLetterText);
        showToast('✓ Текст письма скопирован в буфер!');
    }

    function copyLetterById(vid) {
        const v = vacanciesData.find(item => item.id === vid);
        if (!v) return;
        navigator.clipboard.writeText(v.cover_letter || '');
        showToast(`✓ Сопроводительное для ${v.company} скопировано!`);
    }

    function showToast(text) {
        const toast = document.getElementById('toast-box');
        document.getElementById('toast-text').innerText = text;
        toast.style.display = 'flex';
        setTimeout(() => { toast.style.display = 'none'; }, 3000);
    }

    async function triggerScan() {
        showToast('Запуск автопилота сканирования...');
        await fetch('/api/action/run_scan', { method: 'POST' });
        setTimeout(() => { loadStats(); loadVacancies(); }, 4000);
    }

    async function openExcel() {
        await fetch('/api/action/open_excel', { method: 'POST' });
        showToast('Открываю Excel таблицу...');
    }

    // Init
    loadStats();
    loadVacancies();
    setInterval(loadStats, 15000);
</script>

</body>
</html>'''
    return html_content

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8050, log_level="warning")
