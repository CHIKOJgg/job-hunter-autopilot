#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
             RESPONSE MONITOR — МОДУЛЬ МОНИТОРИНГА И АНАЛИЗА ОТВЕТОВ
=============================================================================
"""

import os
import sys
import json
import sqlite3
import imaplib
import email
from email.header import decode_header
from datetime import datetime
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
DB_PATH = os.path.join(BASE_DIR, "jobs_autopilot.db")

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def classify_response_text(subject, body):
    full = f"{subject} {body}".lower()

    # Priority 1: Interview
    interview_patterns = [
        "приглаш", "собеседован", "интервью", "инвайт", "созвон", "онлайн-встреч",
        "хотим пообщаться", "interview", "invitation", "screening", "schedule a call",
        "congratulations", "offer", "оффер"
    ]
    if any(p in full for p in interview_patterns):
        return "ПРИГЛАШЕНИЕ (ИНТЕРВЬЮ)", "INTERVIEW"

    # Priority 2: Test task
    test_patterns = [
        "тестовое задани", "тестовое", "техническое задани", "ссылк на тест",
        "assessment", "coding challenge", "hackerrank", "codility", "leetcode"
    ]
    if any(p in full for p in test_patterns):
        return "ТЕСТОВОЕ ЗАДАНИЕ", "TEST_TASK"

    # Priority 3: Rejection
    reject_patterns = [
        "к сожалению", "другой кандидат", "выбрали другого", "отказ", "не готовы пригласить",
        "regret to inform", "unfortunate", "moved forward with another", "not selected"
    ]
    if any(p in full for p in reject_patterns):
        return "ОТКАЗ", "REJECTION"

    # Priority 4: Question / Inquiry
    inquiry_patterns = [
        "уточните", "зарплатн", "вилка", "когда готовы", "локаци", "вопрос",
        "clarify", "salary expectations", "notice period", "availability"
    ]
    if any(p in full for p in inquiry_patterns):
        return "УТОЧНЕНИЕ ДАННЫХ", "INQUIRY"

    return "ОТВЕТ ПОЛУЧЕН", "GENERAL"

def process_single_response(company_hint, sender, subject, body, received_at=None):
    if not received_at:
        received_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    status_label, category = classify_response_text(subject, body)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # Try to find matching vacancy by company or domain
    cur.execute("SELECT id, company, title, status FROM vacancies")
    all_vacs = cur.fetchall()

    matched_id = None
    for vid, comp, title, st in all_vacs:
        c_clean = comp.lower().replace(" ", "")
        if c_clean in company_hint.lower() or c_clean in sender.lower() or c_clean in subject.lower():
            matched_id = vid
            break

    snippet = (body[:250] + "...") if len(body) > 250 else body

    if matched_id:
        cur.execute("""
            UPDATE vacancies
            SET status = ?, updated_at = ?
            WHERE id = ?
        """, (status_label, received_at, matched_id))
        print(f"[+] Распознан ответ для вакансии ID {matched_id} ({status_label}): {subject}")
    else:
        # Create a new logged response entry if not matched
        print(f"[i] Получен ответ от '{sender}', тема: '{subject}' -> Категория: {status_label}")

    conn.commit()
    conn.close()
    return status_label, category, matched_id

def check_live_inbox():
    cfg = load_config()
    email_cfg = cfg.get("email", {})
    user = email_cfg.get("username")
    pwd = email_cfg.get("app_password")

    if not pwd or email_cfg.get("dry_run", True):
        print("[RESPONSE MONITOR] Учетные данные IMAP не заданы или активен dry_run. Запуск эмуляции проверок...")
        return 0

    try:
        mail = imaplib.IMAP4_SSL(email_cfg.get("imap_server", "imap.gmail.com"), email_cfg.get("imap_port", 993))
        mail.login(user, pwd.replace(" ", "").strip())
        mail.select("inbox")

        # Search unread messages
        status, messages = mail.search(None, '(UNSEEN)')
        mail_ids = messages[0].split()
        print(f"[RESPONSE MONITOR] Найдено непрочитанных писем: {len(mail_ids)}")

        processed = 0
        for mid in mail_ids[-10:]: # Process last 10
            res, msg_data = mail.fetch(mid, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8", errors="replace")
                    sender = msg.get("From", "")

                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                body = part.get_payload(decode=True).decode(errors="replace")
                                break
                    else:
                        body = msg.get_payload(decode=True).decode(errors="replace")

                    process_single_response(sender, sender, subject, body)
                    processed += 1

        mail.close()
        mail.logout()
        return processed
    except Exception as e:
        print(f"[RESPONSE MONITOR Ошибка IMAP]: {e}")
        return 0

def simulate_sample_responses():
    """Демонстрационная эмуляция ответов для проверки работы воронки аналитики"""
    print("\n[ЭМУЛЯТОР ОТВЕТОВ] Генерация тестовых ответов от работодателей для демонстрации аналитики...")
    samples = [
        ("Т-Банк", "hr@tbank.ru", "Приглашение на технический скрининг — Java разработчик", "Мирослав, добрый день! Ваше резюме произвело отличное впечатление. Предлагаем созвониться в Google Meet на 30 минут для знакомства с командой."),
        ("Picnic", "careers@picnic.app", "Picnic Tech Academy — Invitation to Online Assessment", "Hi Miroslav, thank you for applying to Picnic! We would like to invite you to our HackerRank engineering challenge."),
        ("ID Finance", "hrminsk@idfinance.com", "Отклик на позицию Junior Java Developer", "Здравствуйте, Мирослав! Изучили ваш проект с matching engine. Предлагаем созвониться на этой неделе."),
        ("Крупный холдинг", "auto-reply@recruiting.com", "Результаты рассмотрения резюме", "К сожалению, в настоящий момент мы приняли решение продолжить общение с другими кандидатами.")
    ]
    for comp, sender, subj, body in samples:
        process_single_response(comp, sender, subj, body)

if __name__ == "__main__":
    if "--simulate" in sys.argv:
        simulate_sample_responses()
    else:
        print("Проверка входящих сообщений...")
        cnt = check_live_inbox()
        print(f"Обработано ответов: {cnt}")
