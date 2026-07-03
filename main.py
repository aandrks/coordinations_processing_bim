import re
import json
import pandas as pd
from datetime import datetime, timedelta, date
from collections import defaultdict
from pathlib import Path
import traceback
import unicodedata
import io

import streamlit as st
from rapidfuzz import fuzz
import openpyxl
import xlsxwriter

EMPLOYEE_DB_FILE = 'employee_database.json'
public_domains = {'mail', 'yandex', 'gmail', 'yahoo', 'hotmail', 'outlook'}
holidays = ['01-01', '02-01', '03-01', '04-01', '05-01', '06-01', '07-01',
            '23-02', '08-03', '01-05', '09-05', '12-06', '03-11', '04-11']
working_holidays = ['01-11']
SPEC_CONFIG_FILE = 'spec_config.json'
no_match_array = []


def load_employee_db():
    try:
        if Path(EMPLOYEE_DB_FILE).exists():
            with open(EMPLOYEE_DB_FILE, 'r', encoding='utf-8') as f:
                db = json.load(f)
                db['companies'] = set(db['companies'])
                return db
    except Exception as e:
        st.error(f"Error loading employee database: {e}")
    return {'employees': [], 'companies': set()}


def save_employee_db(db):
    try:
        db_to_save = {'employees': db['employees'], 'companies': list(db['companies'])}
        with open(EMPLOYEE_DB_FILE, 'w', encoding='utf-8') as f:
            json.dump(db_to_save, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"Error saving employee database: {e}")


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


def parse_company_person_data(file_content, db, public_assignments=None):
    """
    public_assignments: dict {email: company} для публичных доменов.
    Если None – обрабатывается как пустой (публичные домены игнорируются).
    """
    company_person_map = defaultdict(list)
    new_employees = []
    seen_emails = {e['email'] for e in db['employees']}
    team_id_counter = 1

    lines = file_content.split('\n')

    for line_num in range(len(lines)):
        line = lines[line_num].strip()
        if not line:
            continue

        if line.endswith('/') and line_num + 1 < len(lines):
            next_line = lines[line_num + 1].strip()
            if next_line:
                line = line.rstrip('/') + ' ' + next_line
                lines[line_num + 1] = ""

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
                        'name': name, 'email': email, 'normalized_name': normalized_name,
                        'surname': surname, 'given_names': given_names,
                        'company': company, 'source': 'manual',
                        'team_id': team_id, 'team_emails': team_emails
                    })
                    company_person_map[company].append({
                        'name': name, 'email': email, 'normalized_name': normalized_name,
                        'surname': surname, 'given_names': given_names,
                        'team_id': team_id, 'team_emails': team_emails
                    })
                else:
                    if team_company is None:
                        team_company = domain
                        db['companies'].add(domain)

                    new_employees.append({
                        'name': name, 'email': email, 'normalized_name': normalized_name,
                        'surname': surname, 'given_names': given_names,
                        'company': domain, 'source': 'auto',
                        'team_id': team_id, 'team_emails': team_emails
                    })
                    company_person_map[domain].append({
                        'name': name, 'email': email, 'normalized_name': normalized_name,
                        'surname': surname, 'given_names': given_names,
                        'team_id': team_id, 'team_emails': team_emails
                    })

    db['employees'].extend(new_employees)
    save_employee_db(db)
    return db, company_person_map


def find_best_match(target_name, candidates, debug_info=None):
    if debug_info is None:
        debug_info = []

    debug_info.append(f"    Finding best match for: {target_name}")
    debug_info.append(f"    Candidates: {[c['name'] for c in candidates]}")
    target_surname, target_given = extract_name_components(target_name)
    debug_info.append(f"    Target surname: {target_surname}, given: {target_given}")
    target_possible_givens = [target_given]

    for candidate in candidates:
        if normalize_text(target_surname) != normalize_text(candidate['surname']):
            continue
        candidate_given_norm = normalize_text(candidate['given_names'])
        for possible_given in target_possible_givens:
            possible_given_norm = normalize_text(possible_given)
            if candidate_given_norm.startswith(possible_given_norm):
                debug_info.append(f"    Initial-based match: {candidate['name']} ...")
                return candidate
            if possible_given_norm.startswith(candidate_given_norm):
                debug_info.append(f"    Initial-based match: {candidate['name']} ...")
                return candidate

    best_score = 0
    best_match = None
    for candidate in candidates:
        score = fuzz.token_set_ratio(normalize_text(target_name), candidate['normalized_name'])
        debug_info.append(f"    Fuzzy score for {candidate['name']}: {score}")
        if score > 65 and score > best_score:
            best_score = score
            best_match = candidate
    if best_match:
        debug_info.append(f"    Fuzzy match selected: {best_match['name']} (score: {best_score})")
    else:
        debug_info.append("    No suitable match found")
    return best_match


def add_working_days(start_date, working_days):
    if working_days <= 0:
        return start_date
    current_date = start_date
    days_added = 0
    iterations = 0
    while days_added < working_days and iterations < 1000:
        current_date += timedelta(days=1)
        iterations += 1
        monthday = current_date.strftime("%d-%m")
        is_weekday = current_date.weekday() < 5
        is_holiday = monthday in holidays
        is_working_holiday = monthday in working_holidays
        if (is_weekday and not is_holiday) or is_working_holiday:
            days_added += 1
    return current_date


def load_spec_config():
    try:
        if Path(SPEC_CONFIG_FILE).exists():
            with open(SPEC_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
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
        return (days, stage_number, f"Stage {stage_number}: not configured, using default {days} days")
    stage_keywords = spec_config[stage_number]
    workflow_lower = workflow_text.lower()
    for keyword, days in stage_keywords.items():
        if keyword.lower() in workflow_lower:
            return (days, stage_number, f"Stage {stage_number}: keyword '{keyword}' → {days} days")
    days = default_days.get(stage_number, 0)
    return (days, stage_number, f"Stage {stage_number}: no keywords found, using default {days} days")


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
        except:
            return None
    return None


def is_team_checked(approver_name, all_people, checked_approvers, matching_log):
    best_match = find_best_match(approver_name, all_people, matching_log)
    if not best_match:
        matching_log.append(f"    No match found for team check: {approver_name}")
        return False
    team_id = best_match.get('team_id')
    team_emails = best_match.get('team_emails', [])
    if not team_id or len(team_emails) <= 1:
        matching_log.append(f"    No team found for: {approver_name}")
        return False
    matching_log.append(f"    Checking team {team_id} with {len(team_emails)} members")
    team_members = [p for p in all_people if p.get('team_id') == team_id]
    for team_member in team_members:
        for checked_name in checked_approvers:
            if find_best_match(checked_name, [team_member], matching_log):
                matching_log.append(f"    ✓ Team member {team_member['name']} is already checked")
                return True
    matching_log.append(f"    ✗ No team members found in checked list")
    return False


def generate_company_report(overdue_counts, person_report, overdue_coordination_ids):
    from collections import defaultdict

    companies = defaultdict(lambda: {'total': 0, 'max_days': 0, 'employees': []})
    for p in person_report:
        comp = p['Организация']

        companies[comp]['max_days'] = max(companies[comp]['max_days'], p['Макс. дней'])
        companies[comp]['employees'].append(p)

    for comp in companies:
        companies[comp]['total'] = overdue_counts.get(comp, 0)

    sorted_companies = sorted(companies.items(), key=lambda x: x[1]['total'], reverse=True)

    lines = []
    total_overdue = len(overdue_coordination_ids)
    lines.append(f"Общее количество просроченных согласований - {total_overdue}")
    lines.append("")

    for comp, data in sorted_companies:
        if data['total'] == 0:
            continue

        sorted_emps = sorted(data['employees'], key=lambda x: x['Просрочек'], reverse=True)

        # пункт 4: вставляем email в заголовок
        emp_names = ', '.join([f'{e["Сотрудник"]} - "{e["Email"]}"' for e in sorted_emps])
        lines.append(
            f"Количество просроченных согласований {comp} ({emp_names}) - {data['total']}, "
            f"макс. срок задержки - {data['max_days']} дня:"
        )
        for emp in sorted_emps:
            lines.append(f"- {emp['Сотрудник']} - {emp['Просрочек']} шт. {emp['Макс. дней']} дня")
        lines.append("")

    return "\n".join(lines)


def process_coordinations(df, company_person_map, today_date, day_period='вечер'):
    overdue_counts = defaultdict(int)
    overdue_emails = []
    overdue_coordination_ids = []
    coordination_details = []
    debug_info = []
    ambiguous_matches = []
    matching_log = []

    all_people = []
    for company, persons in company_person_map.items():
        for person in persons:
            person['company'] = company
            all_people.append(person)

    id_column = df.columns[0] if len(df.columns) > 0 else 'id'

    for idx, row in df.iterrows():
        if not all(col in row for col in ['Не проверили на текущем шаге', 'Шаг', 'Рабочий процесс']):
            matching_log.append(f"Row {idx}: Missing required columns, skipping")
            continue

        coord_id = row.get(id_column, 'N/A')
        step_text = str(row['Шаг'])
        workflow_text = str(row['Рабочий процесс'])
        matching_log.append(f"\nProcessing coordination ID: {coord_id}")
        matching_log.append(f"Step: {step_text}, Workflow: {workflow_text}")

        working_days, stage_number, days_explanation = get_working_days(step_text, workflow_text)
        matching_log.append(f"Working days calculation: {days_explanation}")

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
                    if start_date:
                        matching_log.append(f"  Found start date from lifecycle: {start_date}")
                    else:
                        matching_log.append("  No valid start date found in lifecycle")
            if start_date is None:
                start_date_str = str(row['Дата и время создания согласования'])
                start_date = datetime.strptime(start_date_str, '%Y-%m-%d %H:%M:%S')
                matching_log.append(f"  Using creation date as start date: {start_date}")

            deadline = add_working_days(start_date, working_days)
            matching_log.append(f"Start date: {start_date.date()}, Deadline: {deadline.date()}")

            # пункт 5: переключатель утро/вечер
            if day_period == 'утро':
                if deadline.date() >= today_date:
                    debug_info.append({
                        'id': coord_id,
                        'status': 'Not overdue',
                        'start_date': start_date.date(),
                        'deadline': deadline.date(),
                        'working_days': working_days,
                        'explanation': days_explanation,
                        'today': today_date
                    })
                    matching_log.append("Coordination is not overdue, skipping")
                    continue
            else:  # вечер
                if deadline.date() > today_date:
                    debug_info.append({
                        'id': coord_id,
                        'status': 'Not overdue',
                        'start_date': start_date.date(),
                        'deadline': deadline.date(),
                        'working_days': working_days,
                        'explanation': days_explanation,
                        'today': today_date
                    })
                    matching_log.append("Coordination is not overdue, skipping")
                    continue
        except Exception as e:
            debug_info.append({
                'id': coord_id,
                'status': f'Date error: {str(e)}',
                'start_date_str': start_date_str if 'start_date_str' in locals() else 'N/A',
                'working_days': working_days,
                'explanation': days_explanation,
                'error': e
            })
            matching_log.append(f"Date parsing error: {str(e)}")
            continue

        not_checked_text = str(row['Не проверили на текущем шаге'])
        not_checked_approvers = [name.strip() for name in not_checked_text.split(',') if name.strip()]
        checked_text = str(row['Проверили на текущем шаге'])
        checked_approvers = [name.strip() for name in checked_text.split(',') if name.strip()]

        matching_log.append(f"Not checked approvers: {not_checked_approvers}")
        matching_log.append(f"Checked approvers: {checked_approvers}")

        coord_emails = []
        coord_companies = set()
        checked_members = []
        for approver_name in not_checked_approvers:
            if is_team_checked(approver_name, all_people, checked_approvers, matching_log):
                matching_log.append(f"    Skipping {approver_name} - team member already checked")
                checked_members.append(f"{approver_name} checked!")
                continue

            best_match = find_best_match(approver_name, all_people, matching_log)
            if best_match:
                coord_emails.append(best_match['email'])
                coord_companies.add(best_match['company'])
                matching_log.append(f"    Matched: {approver_name} → {best_match['name']} <{best_match['email']}>")
            else:
                matching_log.append(f"    No match found for: {approver_name}")
                no_match_array.append(approver_name)

        for company in coord_companies:
            overdue_counts[company] += 1
        overdue_emails.extend(coord_emails)
        overdue_coordination_ids.append(coord_id)

        coordination_details.append({
            'id': coord_id,
            'company': ', '.join(coord_companies),
            'start_date': start_date.date(),
            'deadline': deadline.date(),
            'working_days': working_days,
            'not_checked_count': len(not_checked_approvers),
            'explanation': days_explanation,
            'emails': coord_emails
        })

        debug_info.append({
            'id': coord_id,
            'status': 'Overdue',
            'start_date': start_date.date(),
            'deadline': deadline.date(),
            'working_days': working_days,
            'explanation': days_explanation,
            'companies': list(coord_companies),
            'checked_members': checked_members
        })

    return overdue_counts, overdue_emails, overdue_coordination_ids, coordination_details, debug_info, ambiguous_matches, matching_log


st.set_page_config(page_title="Координации", layout="wide")

# st.title("📋 Система контроля просроченных согласований 2")


if 'employee_db' not in st.session_state:
    st.session_state.employee_db = {'employees': [], 'companies': set()}

menu = st.sidebar.radio("2.0.5/nРежим", ["🏢 Загрузка данных", "📊 Обработка согласований", "📂 Загрузить JSON"])

if menu == "🏢 Загрузка данных":
    st.header("Загрузка сотрудников")
    db = st.session_state.employee_db
    st.write(f"В базе {len(db['employees'])} сотрудников, {len(db['companies'])} компаний")

    uploaded_file = st.file_uploader("Файл с сотрудниками (CSV/Excel/TXT)", type=["csv", "xlsx", "txt"])
    if uploaded_file:

        if uploaded_file.name.endswith('.csv') or uploaded_file.name.endswith('.txt'):
            df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')

        file_content = '\n'.join([str(x) for x in df.values.flatten().tolist()])

        public_emails = []
        seen = {e['email'] for e in db['employees']}
        for line in file_content.split('\n'):
            if not line.strip(): continue
            for block in re.findall(r'(?:\(| - )([^()]+?\s+[^\s@]+@[^\s/@]+(?:\s*/\s*[^()]+?\s+[^\s@]+@[^\s/@]+)*)',
                                    line):
                for person in block.split('/'):
                    m = re.search(r'([^@]+)\s+([^\s@]+@[^\s@]+)', person.strip())
                    if m:
                        email = re.sub(r'[),.;]+$', '', m.group(2).strip())
                        domain = email.split('@')[-1].split('.')[0]
                        if domain in public_domains and email not in seen:
                            public_emails.append((m.group(1).strip(), email))
        public_emails = list(set(public_emails))

        assignments = {}
        if public_emails:
            st.warning(f"Обнаружено {len(public_emails)} публичных email – укажите компании")
            for name, email in public_emails:
                col1, col2 = st.columns([3, 2])
                with col1:
                    st.write(f"{name} <{email}>")
                with col2:
                    org = st.text_input(f"Компания", key=email)
                    if org:
                        assignments[email] = org
            if st.button("Сохранить"):
                if len(assignments) != len(public_emails):
                    st.error("Назначьте компании для всех публичных адресов")
                else:
                    db, _ = parse_company_person_data(file_content, db, assignments)
                    st.session_state.employee_db = db
                    st.success(f"Готово! Сотрудников: {len(db['employees'])}")
                    # пункт 9: показать загруженных сотрудников
                    if db['employees']:
                        st.subheader("Загруженные сотрудники")
                        st.dataframe(pd.DataFrame(db['employees']))
        else:
            if st.button("Загрузить"):
                db, _ = parse_company_person_data(file_content, db, {})
                st.session_state.employee_db = db
                st.success(f"Готово! Сотрудников: {len(db['employees'])}")
                # пункт 9: показать загруженных сотрудников
                if db['employees']:
                    st.subheader("Загруженные сотрудники")
                    st.dataframe(pd.DataFrame(db['employees']))


elif menu == "📊 Обработка согласований":
    st.header("Просроченные согласования")
    db = st.session_state.employee_db
    if not db['employees']:
        st.error("Сначала загрузите сотрудников")
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

        uploaded_file = st.file_uploader("Файл согласований (CSV/Excel)", type=["csv", "xlsx"])
        if uploaded_file:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            check_date = st.date_input("Дата проверки", value=datetime.today().date())
            # пункт 5: переключатель утро/вечер
            day_period = st.radio("Время отсечки", ["утро", "вечер"], index=1,
                                  help="утро – дедлайн сегодня ещё НЕ просрочен; вечер – дедлайн сегодня УЖЕ просрочен")
            if st.button("Найти просрочки"):
                with st.spinner("Анализируем..."):
                    (overdue_counts, overdue_emails, overdue_ids,
                     coordination_details, debug_info, ambiguous_matches,
                     matching_log) = process_coordinations(df, company_person_map, check_date, day_period)

                person_overdue = defaultdict(lambda: {'company': '', 'count': 0, 'overdue_days': []})
                for d in coordination_details:
                    dd = d['deadline']
                    days_late = (check_date - dd).days if isinstance(dd, date) else (check_date - dd.date()).days
                    if day_period == 'утро':
                        days_late = max(0, days_late - 1)
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

                st.subheader("По организациям")
                for comp, cnt in sorted(overdue_counts.items(), key=lambda x: x[1], reverse=True):
                    st.write(f"- **{comp}**: {cnt}")
                st.subheader("По сотрудникам")
                st.dataframe(df_report, use_container_width=True)

                csv = df_report.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 CSV", csv, "person_overdue_report.csv")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_report.to_excel(writer, index=False, sheet_name='Отчёт')
                st.download_button("📥 Excel", output.getvalue(), "person_overdue_report.xlsx")

                report_text = generate_company_report(overdue_counts, report, overdue_ids)
                st.subheader("📊 Сводный отчёт по компаниям и сотрудникам")
                st.code(report_text, language='text')  # или st.text(report_text)

                if overdue_emails:
                    emails_txt = '\n'.join(sorted(set(overdue_emails)))
                    st.download_button("📥 Скачать список Email (overdue_emails.txt)",
                                       emails_txt, "overdue_emails.txt", "text/plain")
                    st.subheader("Email адреса просрочивших")
                    st.code(emails_txt, language='text')

                st.download_button(
                    "📥 Скачать отчёт (TXT)",
                    report_text,
                    "overdue_report.txt",
                    "text/plain"
                )

                st.subheader("Детали согласований")
                st.dataframe(pd.DataFrame(coordination_details), use_container_width=True)


elif menu == "📂 Загрузить JSON":
    st.header("Импорт/Экспорт базы")
    uploaded_json = st.file_uploader("Загрузить employee_database.json", type="json")
    if uploaded_json:
        try:
            data = json.load(uploaded_json)
            if 'employees' in data and 'companies' in data:
                data['companies'] = set(data['companies'])
                st.session_state.employee_db = data
                save_employee_db(data)
                st.success(f"Загружено {len(data['employees'])} сотрудников")
                # пункт 9: показать загруженных сотрудников (при загрузке JSON тоже)
                if data['employees']:
                    st.subheader("Загруженные сотрудники")
                    st.dataframe(pd.DataFrame(data['employees']))
            else:
                st.error("Неверный формат")
        except Exception as e:
            st.error(f"Ошибка: {e}")
    if st.button("Скачать текущую базу"):
        db = st.session_state.employee_db
        db_json = {'employees': db['employees'], 'companies': list(db['companies'])}
        st.download_button("Сохранить JSON", json.dumps(db_json, ensure_ascii=False, indent=2),
                           "employee_database.json", "application/json")