#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
         CLOUD SERVICE ENTRYPOINT — RAILWAY 24/7 BACKGROUND SERVICE
=============================================================================
Runs:
1. Web Dashboard (FastAPI / Uvicorn) on 0.0.0.0:$PORT
2. Background Vacancy Autopilot Daemon (scans Habr, Remotive, Telegram every 60 min)
3. Background Response Monitor (checks Gmail inbox via IMAP every 15 min)
"""

import os
import sys
import time
import threading
import uvicorn
from datetime import datetime

# Set up logging encoding
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from dashboard_server import app
import job_autopilot_daemon
import response_monitor

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [CLOUD RUNNER] {msg}", flush=True)

def background_autopilot_worker():
    """Background worker that continuously searches and updates vacancies"""
    log("Запуск фонового демона поиска вакансий (интервал: 60 мин)...")
    time.sleep(10)  # Wait for server to boot first
    while True:
        try:
            log("Старт планового цикла поиска вакансий...")
            job_autopilot_daemon.run_cycle()
            log("Плановый цикл поиска завершен успешно.")
        except Exception as e:
            log(f"Ошибка в цикле автопилота: {e}")
        
        # Sleep for 60 minutes
        time.sleep(60 * 60)

def background_inbox_worker():
    """Background worker that monitors incoming recruiter emails"""
    log("Запуск фонового монитора ответов HR (интервал: 15 мин)...")
    time.sleep(30)
    while True:
        try:
            log("Проверка новых входящих ответов от HR...")
            response_monitor.check_live_inbox()
        except Exception as e:
            log(f"Ошибка в мониторе почты: {e}")
        
        # Sleep for 15 minutes
        time.sleep(15 * 60)

@app.on_event("startup")
def start_background_tasks():
    log("Инициализация облачных фоновых задач...")
    t_daemon = threading.Thread(target=background_autopilot_worker, daemon=True, name="AutopilotWorker")
    t_daemon.start()

    t_inbox = threading.Thread(target=background_inbox_worker, daemon=True, name="InboxMonitorWorker")
    t_inbox.start()
    log("Все фоновые потоки успешно активированы.")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    log(f"Запуск веб-сервера на 0.0.0.0:{port}...")
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
