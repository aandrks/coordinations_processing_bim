import re
import json
import pandas as pd
from datetime import datetime, timedelta, date
from collections import defaultdict
from pathlib import Path
import unicodedata
import io

import streamlit as st
from rapidfuzz import fuzz
import openpyxl
import xlsxwriter

# -------------------- CONFIG --------------------
public_domains = {'mail', 'yandex', 'gmail', 'yahoo', 'hotmail', 'outlook'}
holidays = ['01-01', '02-01', '03-01', '04-01', '05-01', '06-01', '07-01',
            '23-02', '08-03', '01-05', '09-05', '12-06', '03-11', '04-11']
working_holidays = ['01-11']
SPEC_CONFIG_FILE = 'spec_config.json'

# -------------------- DATABASE HELPERS --------------------
def load_employee_db():
    if 'employee_db' not in st.session_state:
        st.session_state.employee_db = {'employees': [], 'companies': set()}
    return st.session_state.employee_db

def save_employee_db(db):
    st.session_state.employee_db = db

# -------------------- TEXT PROCESSING --------------------
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFKD', text)
    text = ''.join([c for c in text if not unicodedata.combining(c)])
    text = re.sub(r'[^\w\s.]', '', text)
    return text.lower().strip()

def is_initial(part):
    return len(part) <= 2 or (len(part) == 2 and part.endswith('.'))

def extract_name_components(name):
    if not isinstance(name, str):
        return "", ""
    clean_name = re.sub(r'[^а-яА-ЯёЁa-zA-Z\s.]', '', name).strip()
    if ',' in clean_name:
        parts = [p.strip() for p in clean_name.split(',')]
        if len(parts) >= 2:
            return parts[0], parts[1]
    parts = [p for p in re.split(r'\s+', clean_name) if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    if is_initial(parts[-1]):
        return parts[0], parts[-1]
    if is_initial(parts[0]):
        return parts[-1], parts[0]
    return parts[-1], " ".join(parts[:-1])

# -------------------- PARSING --------------------
def parse_company_person_data(file_content, db, public_assignments=None):
    company_person_map = defaultdict(list)
    new_employees = []
    seen_emails = {e['email'] for e in db['employees']}
    team_id_counter = 1
    lines = file_content.split('\n')

    for i, line in enumerate(lines):
        if not line.strip():
            continue
        if line.endswith('/') and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if next_line:
                line = line.rstrip('/') + ' ' + next_line
                lines[i + 1] = ""

        for block in re.findall(r'(?:\(| - )([^()]+?\s+[^\s@]+@[^\s/@]+(?:\s*/\s*[^()]+?\s+[^\s@]+@[^\s/@]+)*)', line):
            if '@' not in block:
                continue
            team_members = [p.strip() for p in block.split('/')]
            team_id = f"team_{team_id_counter}"
            team_id_counter += 1
            team_company = None
            team_emails = []

            for person in team_members:
                match = re.search(r'([^@]+)\s+([^\s@]+@[^\s@]+)', person)
                if not match:
                    continue
                name, email = match.group(1).strip(), match.group(2).strip()
                email = re.sub(r'[),.;]+$', '', email).strip()
                if email in seen_emails:
                    continue
                seen_emails.add(email)
                team_emails.append(email)

                domain = email.split('@')[-1].split('.')[0]
                surname, given_names = extract_name_components(name)
                normalized_name = normalize_text(name)

                if domain in public_domains:
                    company = public_assignments.get(email) if public_assignments else None
                    if not company:
                        continue
                    if team_company is None:
                        team_company = company
                        db['companies'].add(company)
                    new_employees.append({
                        'name': name,
                        'email': email,
                        'normalized_name': normalized_name,
                        'surname': surname,
                        'given_names': given_names,
                        'company': company,
                        'source': 'manual',
                        'team_id': team_id,
                        'team_emails': team_emails
                    })
                    company_person_map[company].append({
                        'name': name,
                        'email': email,
                        'normalized_name': normalized_name,
                        'surname': surname,
                        'given_names': given_names,
                        'team_id': team_id,
                        'team_emails': team_emails
                    })
                else:
                    if team_company is None:
                        team_company = domain
                        db['companies'].add(domain)
                    new_employees.append({
                        'name': name,
                        'email': email,
                        'normalized_name': normalized_name,
                        'surname': surname,
                        'given_names': given_names,
                        'company': domain,
                        'source': 'auto',
                        'team_id': team_id,
                        'team_emails': team_emails
                    })
                    company_person_map[domain].append({
                        'name': name,
                        'email': email,
                        'normalized_name': normalized_name,
                        'surname': surname,
                        'given_names': given_names,
                        'team_id': team_id,
                        'team_emails': team_emails
                    })

    db['employees'].extend(new_employees)
    save_employee_db(db)
    return db, company_person_map

# -------------------- MATCHING --------------------
def find_best_match(target_name, candidates):
    target_surname, target_given = extract_name_components(target_name)
    for candidate in candidates:
        if normalize_text(target_surname) != normalize_text(candidate['surname']):
            continue
        cand_given = normalize_text(candidate['given_names'])
        target_given_norm = normalize_text(target_given)
        if cand_given.startswith(target_given_norm) or target_given_norm.startswith(cand_given):
            return candidate
    best_score = 0
    best_match = None
    for candidate in candidates:
        score = fuzz.token_set_ratio(normalize_text(target_name), candidate['normalized_name'])
        if score > 65 and score > best_score:
            best_score = score
            best_match = candidate
    return best_match

def add_working_days(start_date, working_days):
    if working_days <= 0:
        return start_date
    current_date = start_date
    days_added = 0
    while days_added < working_days:
        current_date += timedelta(days=1)
        monthday = current_date.strftime("%d-%m")
        is_weekday = current_date.weekday() < 5
        is_holiday = monthday in holidays
        is_working_holiday = monthday in working_holidays
        if (is_weekday and not is_holiday) or is_working_holiday:
            days_added += 1
    return current_date

def load_spec_config():
    if Path(SPEC_CONFIG_FILE).exists():
        with open(SPEC_CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"2": {"раздела КР": 2}, "3": {}, "4": {}}

def get_working_days(step_text, workflow_text):
    if "Утверждение" in step_text:
        return (2, 4, "Stage 4")
    stage_match = re.search(r'Шаг (\d+)', step_text)
    if not stage_match:
        return (0, 0, "No stage number found")
    step_number = int(stage_match.group(1))
    stage_number = step_number + 1
    spec_config = load_spec_config()
    default_days = {2: 3, 3: 5, 4: 2}
    if stage_number not in spec_config:
        days = default_days.get(stage_number, 0)
        return (days, stage_number, f"Stage {stage_number}: default {days} days")
    stage_keywords = spec_config[stage_number]
    workflow_lower = workflow_text.lower()
    for keyword, days in stage_keywords.items():
        if keyword.lower() in workflow_lower:
            return (days, stage_number, f"Stage {stage_number}: keyword '{keyword}' → {days} days")
    days = default_days.get(stage_number, 0)
    return (days, stage_number, f"Stage {stage_number}: no keyword, default {days} days")

def extract_start_date_from_lifecycle(lifecycle_text, current_step_number):
    if not lifecycle_text or pd.isna(lifecycle_text):
        return None
    current_stage = current_step_number + 1
    target_step = current_stage - 2
    if target_step < 0:
        return None
    step_pattern = rf'Шаг {target_step}.*?(\d{{2}}\.\d{{2}}\.\d{{2}} \d{{2}}:\d{{2}})'
    matches = re.findall(step_pattern, lifecycle_text, re.IGNORECASE | re.DOTALL)
    if matches:
        try:
            return datetime.strptime(matches[-1], '%d.%m.%y %H:%M')
        except ValueError:
            return None
    return None

def is_team_checked(approver_name, all_people, checked_approvers):
    best_match = find_best_match(approver_name, all_people)
    if not best_match:
        return False
    team_id = best_match.get('team_id')
    team_emails = best_match.get('team_emails', [])
    if not team_id or len(team_emails) <= 1:
        return False
    team_members = [p for p in all_people if p.get('team_id') == team_id]
    for tm in team_members:
        for checked_name in checked_approvers:
            if find_best_match(checked_name, [tm]):
                return True
    return False

def process_coordinations(df, company_person_map, today_date):
    overdue_counts = defaultdict(int)
    overdue_emails = []
    overdue_ids = []
    coordination_details = []
    all_people = []
    for company, persons in company_person_map.items():
        for person in persons:
            person['company'] = company
            all_people.append(person)

    id_column = df.columns[0] if len(df.columns) > 0 else 'id'

    for idx, row in df.iterrows():
        if not all(col in row for col in ['Не проверили на текущем шаге', 'Шаг', 'Рабочий процесс']):
            continue
        coord_id = row.get(id_column, 'N/A')
        step_text = str(row['Шаг'])
        workflow_text = str(row['Рабочий процесс'])
        working_days, stage_number, days_explanation = get_working_days(step_text, workflow_text)

        try:
            start_date = None
            if 'Жизненный цикл' in row and row['Жизненный цикл']:
                lifecycle_text = str(row['Жизненный цикл'])
                step_match = re.search(r'Шаг (\d+)', step_text)
                if "Утверждение" in step_text:
                    current_step_number = 3
                    start_date = extract_start_date_from_lifecycle(lifecycle_text, current_step_number)
                elif step_match:
                    current_step_number = int(step_match.group(1))
                    start_date = extract_start_date_from_lifecycle(lifecycle_text, current_step_number)
            if start_date is None:
                start_date_str = str(row['Дата и время создания согласования'])
                start_date = datetime.strptime(start_date_str, '%Y-%m-%d %H:%M:%S')
            deadline = add_working_days(start_date, working_days)
            if deadline.date() >= today_date:
                continue
        except Exception:
            continue

        not_checked_text = str(row['Не проверили на текущем шаге'])
        not_checked_approvers = [name.strip() for name in not_checked_text.split(',') if name.strip()]
        checked_text = str(row['Проверили на текущем шаге'])
        checked_approvers = [name.strip() for name in checked_text.split(',') if name.strip()]

        coord_emails = []
        coord_companies = set()
        for approver_name in not_checked_approvers:
            if is_team_checked(approver_name, all_people, checked_approvers):
                continue
            best_match = find_best_match(approver_name, all_people)
            if best_match:
                coord_emails.append(best_match['email'])
                coord_companies.add(best_match['company'])

        for company in coord_companies:
            overdue_counts[company] += 1
        overdue_emails.extend(coord_emails)
        overdue_ids.append(coord_id)
        coordination_details.append({
            'id': coord_id,
            'company': ', '.join(coord_companies) if coord_companies else '',
            'start_date': start_date.date() if start_date else '',
            'deadline': deadline.date(),
            'working_days': working_days,
            'not_checked_count': len(not_checked_approvers),
            'explanation': days_explanation,
            'emails': coord_emails
        })

    return overdue_counts, overdue_emails, overdue_ids, coordination_details

# -------------------- STREAMLIT UI --------------------
st.set_page_config(page_title="Обработка согласований", layout="wide")
st.title("📋 Система контроля просроченных согласований")

if 'employee_db' not in st.session_state:
    st.session_state.employee_db = {'employees': [], 'companies': set()}

menu = st.sidebar.radio(
    "Режим работы",
    ["🏢 Загрузка данных сотрудников", "📊 Обработка согласований", "📂 Загрузить готовую базу JSON"]
)

# ---------- ЗАГРУЗКА ДАННЫХ ----------
if menu == "🏢 Загрузка данных сотрудников":
    st.header("Загрузка данных сотрудников")
    db = st.session_state.employee_db
    st.info(f"В базе {len(db['employees'])} сотрудников из {len(db['companies'])} организаций")

    uploaded_file = st.file_uploader("Файл со списком сотрудников (CSV/Excel)", type=["csv", "xlsx", "txt"])
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv') or uploaded_file.name.endswith('.txt'):
                df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            file_content = '\n'.join(df.astype(str).values.flatten().tolist())

            public_emails = []
            seen_emails = {e['email'] for e in db['employees']}
            lines = file_content.split('\n')
            for line in lines:
                if not line.strip():
                    continue
                for block in re.findall(r'(?:\(| - )([^()]+?\s+[^\s@]+@[^\s/@]+(?:\s*/\s*[^()]+?\s+[^\s@]+@[^\s/@]+)*)', line):
                    for person in block.split('/'):
                        match = re.search(r'([^@]+)\s+([^\s@]+@[^\s@]+)', person.strip())
                        if match:
                            name, email = match.group(1).strip(), match.group(2).strip()
                            email = re.sub(r'[),.;]+$', '', email)
                            domain = email.split('@')[-1].split('.')[0]
                            if domain in public_domains and email not in seen_emails:
                                public_emails.append((name, email))
            public_emails = list(set(public_emails))

            if public_emails:
                st.warning(f"Найдено {len(public_emails)} email на публичных доменах. Укажите организации.")
                assignments = {}
                for name, email in public_emails:
                    col1, col2 = st.columns([3, 2])
                    with col1:
                        st.write(f"{name} <{email}>")
                    with col2:
                        org = st.text_input(f"Организация", key=email)
                        if org:
                            assignments[email] = org
                if st.button("Сохранить и обработать"):
                    if len(assignments) != len(public_emails):
                        st.error("Назначьте организации для всех публичных адресов.")
                    else:
                        db, _ = parse_company_person_data(file_content, db, assignments)
                        st.success(f"Готово! В базе {len(db['employees'])} сотрудников.")
            else:
                if st.button("Обработать файл"):
                    db, _ = parse_company_person_data(file_content, db, {})
                    st.success(f"Готово! В базе {len(db['employees'])} сотрудников.")
        except Exception as e:
            st.error(f"Ошибка: {e}")

# ---------- ОБРАБОТКА СОГЛАСОВАНИЙ ----------
elif menu == "📊 Обработка согласований":
    st.header("Анализ просроченных согласований")
    db = st.session_state.employee_db
    if len(db['employees']) == 0:
        st.error("База сотрудников пуста. Сначала загрузите данные.")
    else:
        company_person_map = defaultdict(list)
        for emp in db['employees']:
            company_person_map[emp['company']].append({
                'name': emp['name'],
                'email': emp['email'],
                'normalized_name': emp['normalized_name'],
                'surname': emp['surname'],
                'given_names': emp['given_names'],
                'team_id': emp.get('team_id', ''),
                'team_emails': emp.get('team_emails', [])
            })

        uploaded_file = st.file_uploader("Файл с согласованиями (CSV/Excel)", type=["csv", "xlsx"])
        if uploaded_file:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
                else:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                check_date = st.date_input("Дата проверки", value=datetime.today().date())
                if st.button("🔍 Найти просрочки"):
                    with st.spinner("Обрабатываем..."):
                        overdue_counts, _, _, details = process_coordinations(df, company_person_map, check_date)

                    # Отчёт по сотрудникам
                    person_overdue = defaultdict(lambda: {'company': '', 'count': 0, 'overdue_days': []})
                    for d in details:
                        dd = d['deadline']
                        days_late = (check_date - dd).days if isinstance(dd, date) else (check_date - dd.date()).days
                        for email in d['emails']:
                            person_overdue[email]['count'] += 1
                            person_overdue[email]['overdue_days'].append(days_late)
                    for emp in db['employees']:
                        if emp['email'] in person_overdue:
                            person_overdue[emp['email']]['company'] = emp['company']
                            person_overdue[emp['email']]['name'] = emp['name']
                    report = []
                    for email, data in person_overdue.items():
                        report.append({
                            'Сотрудник': data.get('name', email.split('@')[0]),
                            'Email': email,
                            'Организация': data['company'] or '—',
                            'Просрочек': data['count'],
                            'Макс. дней': max(data['overdue_days'])
                        })
                    report.sort(key=lambda x: x['Просрочек'], reverse=True)
                    df_report = pd.DataFrame(report)

                    st.subheader("📊 По организациям")
                    for comp, cnt in sorted(overdue_counts.items(), key=lambda x: x[1], reverse=True):
                        st.write(f"- **{comp}**: {cnt} просрочек")

                    st.subheader("👤 По сотрудникам")
                    st.dataframe(df_report, use_container_width=True)

                    # Скачивание отчётов
                    col1, col2 = st.columns(2)
                    with col1:
                        csv_data = df_report.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("📥 Скачать CSV", csv_data, "person_overdue_report.csv", "text/csv")
                    with col2:
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df_report.to_excel(writer, index=False, sheet_name='Сотрудники')
                        st.download_button("📥 Скачать Excel", output.getvalue(), "person_overdue_report.xlsx")

                    st.subheader("📄 Детали просроченных согласований")
                    st.dataframe(pd.DataFrame(details), use_container_width=True)
            except Exception as e:
                st.error(f"Ошибка: {e}")

# ---------- ЗАГРУЗКА JSON ----------
elif menu == "📂 Загрузить готовую базу JSON":
    st.header("Загрузка базы сотрудников из JSON")
    uploaded_json = st.file_uploader("employee_database.json", type="json")
    if uploaded_json:
        try:
            data = json.load(uploaded_json)
            if 'employees' in data and 'companies' in data:
                data['companies'] = set(data['companies'])
                st.session_state.employee_db = data
                st.success(f"База загружена: {len(data['employees'])} сотрудников.")
            else:
                st.error("Неверный формат JSON.")
        except Exception as e:
            st.error(f"Ошибка: {e}")

    if st.button("💾 Скачать текущую базу"):
        db = st.session_state.employee_db
        db_json = {'employees': db['employees'], 'companies': list(db['companies'])}
        st.download_button(
            "📥 Сохранить JSON",
            json.dumps(db_json, ensure_ascii=False, indent=2),
            "employee_database.json",
            "application/json"
        )