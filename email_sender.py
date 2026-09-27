#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
                  EMAIL AUTOPILOT SENDER — МОДУЛЬ АВТОПОДАЧИ ПО EMAIL
=============================================================================
"""

import os
import sys
import json
import sqlite3
import smtplib

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
DB_PATH = os.path.join(BASE_DIR, "jobs_autopilot.db")

def load_config():
    cfg = {}
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    if "APP_PASSWORD" in os.environ:
        cfg.setdefault("email", {})["app_password"] = os.environ["APP_PASSWORD"]
    if "EMAIL_USERNAME" in os.environ:
        cfg.setdefault("email", {})["username"] = os.environ["EMAIL_USERNAME"]
    return cfg

def send_email_application(vacancy_id, target_email=None):
    cfg = load_config()
    email_cfg = cfg.get("email", {})
    cand_cfg = cfg.get("candidate", {})

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, company, title, url, recommended_resume, cover_letter, status
        FROM vacancies WHERE id = ?
    """, (vacancy_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, f"Вакансия ID {vacancy_id} не найдена."

    v_id, company, title, url, resume_name, cover_letter, status = row

    # Determine destination email
    dest_email = target_email
    if not dest_email:
        if "mailto:" in url:
            dest_email = url.replace("mailto:", "").strip()
        elif "@" in url:
            dest_email = url.strip()
        elif "id finance" in company.lower():
            dest_email = "hrminsk@idfinance.com"
        elif "innowise" in company.lower():
            dest_email = "lab@innowise.com"
        elif "leverx" in company.lower():
            dest_email = "training@leverx.com"
        elif "aston" in company.lower():
            dest_email = "cv@astondevs.ru"

    if not dest_email or "@" not in dest_email:
        conn.close()
        return False, f"Не удалось определить Email для компании '{company}'."

    resume_path = os.path.join(BASE_DIR, resume_name)
    if not os.path.exists(resume_path):
        # Fallback to RU 1YOE
        resume_path = os.path.join(BASE_DIR, "Resume_RU_1.pdf")

    subject = f"Отклик на позицию {title} — {cand_cfg.get('name_ru', 'Мирослав Писарик')}"
    dry_run = email_cfg.get("dry_run", True)

    print(f"\n[EMAIL SENDER] Подготовка отклика для {company} ({dest_email})")
    print(f"  Тема: {subject}")
    print(f"  Прикрепляемый файл: {os.path.basename(resume_path)}")
    print(f"  Режим: {'[ТЕСТОВЫЙ / DRY RUN - письмо не отправляется в сеть]' if dry_run else '[БОЕВОЙ / LIVE]'}")

    if not dry_run:
        user = email_cfg.get("username")
        pwd = email_cfg.get("app_password")
        if not user or not pwd:
            conn.close()
            return False, "Не указан username или app_password в config.json"

        try:
            msg = MIMEMultipart()
            msg['From'] = f"{cand_cfg.get('name_ru', 'Мирослав Писарик')} <{user}>"
            msg['To'] = dest_email
            msg['Subject'] = subject

            msg.attach(MIMEText(cover_letter, 'plain', 'utf-8'))

            if os.path.exists(resume_path):
                with open(resume_path, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(resume_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(resume_path)}"'
                msg.attach(part)

            import socket
            server = None
            host = email_cfg.get("smtp_server", "smtp.gmail.com")
            port = email_cfg.get("smtp_port", 587)

            # Dynamically detect non-tun physical IPv4 addresses
            detected_ips = []
            try:
                hostname = socket.gethostname()
                for ip in socket.gethostbyname_ex(hostname)[2]:
                    if not ip.startswith('127.') and not ip.startswith('169.254.') and not ip.startswith('172.19.'):
                        if ip not in detected_ips:
                            detected_ips.append(ip)
            except Exception:
                pass
            detected_ips.append(None) # fallback to default routing

            for local_ip in detected_ips:
                try:
                    kwargs = {'timeout': 6}
                    if local_ip:
                        kwargs['source_address'] = (local_ip, 0)
                    s = smtplib.SMTP(host, port, **kwargs)
                    s.starttls()
                    s.login(user, pwd.replace(" ", "").strip())
                    server = s
                    break
                except Exception:
                    continue

            if not server:
                conn.close()
                return False, "Не удалось подключиться к SMTP серверу (таймаут соединения)."

            server.send_message(msg)
            server.quit()
            print("  [OK] Письмо успешно отправлено через SMTP!")
        except Exception as e:
            conn.close()
            return False, f"Ошибка SMTP: {e}"
    else:
        print("  [DRY RUN] Имитация успешной отправки отклика.")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute("""
        UPDATE vacancies
        SET status = 'ОТПРАВЛЕНО (EMAIL)', updated_at = ?
        WHERE id = ?
    """, (now_str, vacancy_id))
    conn.commit()
    conn.close()

    return True, f"Отклик успешно отправлен на {dest_email} (Статус обновлен в БД)"

def batch_send_pending():
    """Отправляет отклики на все вакансии с известными email адресами"""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, company, title, url
        FROM vacancies
        WHERE (url LIKE '%@%' OR url LIKE '%mailto:%' OR company LIKE '%ID Finance%' OR company LIKE '%Innowise%' OR company LIKE '%ASTON%')
          AND status != 'ОТПРАВЛЕНО (EMAIL)'
    """)
    rows = cur.fetchall()
    conn.close()

    print(f"Найдено подходящих email-вакансий для отправки: {len(rows)}")
    success_count = 0
    for r in rows:
        ok, msg = send_email_application(r[0])
        print(f"ID {r[0]}: {msg}")
        if ok:
            success_count += 1
    return success_count

if __name__ == "__main__":
    if len(sys.argv) > 1:
        target_id = int(sys.argv[1])
        ok, res = send_email_application(target_id)
        print(res)
    else:
        print("Запуск пакетной отправки по известным email-адресам...")
        cnt = batch_send_pending()
        print(f"\nВсего обработано: {cnt}")
