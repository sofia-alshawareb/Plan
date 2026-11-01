"""Генерация плана работ по поддержанию работоспособности системы (plan.xlsx, plan.md).

Длительность — в рабочих днях. Даты начала считаются от окончания предшествующих
задач (формулы WORKDAY в Excel), поэтому при сдвиге одной задачи сдвигаются и зависимые.
"""
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJECT_START = dt.date(2026, 11, 9)

# Нерабочие праздничные дни по производственному календарю РФ на 2026–2027 гг.
HOLIDAYS = [dt.date(2026, 11, 4), dt.date(2026, 12, 31)] + [
    dt.date(2027, 1, d) for d in range(1, 9)
]

# (номер, задача, длительность, предшественники, ответственные, результат)
# Строки с длительностью None — заголовки этапов, их сроки и % считаются по подзадачам.
TASKS = [
    ("0", "Организационный этап", None, [], "ЦПС, УИСО, УКЗ", ""),
    ("0.1", "Согласование плана, назначение контактных лиц от ЦПС, УИСО, УКЗ", 3, [],
     "ЦПС (отв.), УИСО, УКЗ", "Утверждённый план, список контактов"),
    ("0.2", "Согласование окон регламентных работ и порядка оповещения пользователей", 2, ["0.1"],
     "УИСО (отв.), ЦПС", "График окон работ"),
    ("0.3", "Актуализация доступов исполнителей к ландшафту (УЗ, VPN, права)", 2, ["0.1"],
     "УИСО (отв.), УКЗ", "Доступы предоставлены"),

    ("1", "Выполнение рекомендаций подрядчика по отладке работоспособности", None, [], "ЦПС", ""),
    ("1.1", "Анализ рекомендаций, план работ и план отката", 3, ["0.1"],
     "ЦПС", "План работ и отката"),
    ("1.2", "Резервное копирование: снапшоты ВМ, бэкап БД и конфигураций", 1, ["1.1", "0.2", "0.3"],
     "ЦПС (отв.), УИСО (снапшоты ВМ)", "Резервные копии"),
    ("1.3", "Обновление ПО на тестовом контуре, проверка", 3, ["1.2"],
     "ЦПС", "Протокол проверки"),
    ("1.4", "Обновление ПО на продуктивном контуре (в окно работ)", 2, ["1.3"],
     "ЦПС (отв.), УИСО (окно, контроль)", "ПО обновлено"),
    ("1.5", "Замена TLS-сертификата кластера СУБД", None, [], "ЦПС, УКЗ", ""),
    ("1.5.1", "Формирование запроса на сертификат (CSR)", 1, ["1.1"],
     "ЦПС", "CSR передан в УКЗ"),
    ("1.5.2", "Выпуск сертификата в корпоративном УЦ", 3, ["1.5.1"],
     "УКЗ", "Сертификат выпущен"),
    ("1.5.3", "Установка сертификата, перезапуск узлов кластера, проверка репликации и подключений", 1,
     ["1.5.2", "1.4"], "ЦПС", "Сертификат заменён"),
    ("1.6", "Продление лицензии Angie PRO", None, [], "ЦПС", ""),
    ("1.6.1", "Запрос и получение лицензии у вендора", 10, ["0.1"],
     "ЦПС", "Лицензия получена"),
    ("1.6.2", "Установка лицензии, проверка работы", 1, ["1.6.1", "1.4"],
     "ЦПС", "Лицензия активна"),
    ("1.7", "Продление etcd (лицензия / перевыпуск сертификатов), поочерёдный перезапуск членов кластера", 1,
     ["1.4"], "ЦПС", "etcd работает, срок продлён"),
    ("1.8", "Контроль работоспособности после работ, отчёт о выполнении рекомендаций", 2,
     ["1.5.3", "1.6.2", "1.7"], "ЦПС (отв.), УИСО (приёмка)", "Отчёт"),

    ("2", "Настройка KES (Kaspersky Endpoint Security)", None, [], "УКЗ, ЦПС, УИСО", ""),
    ("2.1", "Обследование: перечень серверов, ОС, совместимость KES for Linux с ПО системы", 3, ["0.1"],
     "УКЗ (отв.), ЦПС", "Перечень узлов"),
    ("2.2", "Перечень исключений (доверенная зона): каталоги СУБД, etcd, Angie, процессы приложений", 3,
     ["2.1"], "ЦПС (отв.), УКЗ (согласование)", "Согласованные исключения"),
    ("2.3", "Правила МСЭ: узлы ↔ сервер KSC (порты Агента администрирования)", 3, ["2.1"],
     "УИСО (исполн.), УКЗ (заявка)", "Доступ открыт"),
    ("2.4", "Настройка KSC: группа администрирования, политики, задачи обновления баз и проверки вне пиковых часов",
     3, ["2.2"], "УКЗ", "Политики настроены"),
    ("2.5", "Пилот: установка Агента администрирования и KES на тестовый контур", 2,
     ["2.3", "2.4", "1.3"], "УКЗ (отв.), ЦПС (контроль)", "KES на тестовом контуре"),
    ("2.6", "Пилотная эксплуатация: нагрузка, ложные срабатывания, корректировка исключений", 5, ["2.5"],
     "ЦПС, УКЗ", "Протокол пилота"),
    ("2.7", "Установка на продуктивный контур (поэтапно, в окна работ)", 3, ["2.6", "1.8"],
     "УКЗ (отв.), ЦПС, УИСО", "KES на всех узлах"),
    ("2.8", "Проверка статуса защиты в KSC, актуализация документации", 2, ["2.7"],
     "УКЗ (отв.), ЦПС", "Акт / отчёт"),

    ("3", "Настройка Zabbix", None, [], "УИСО, ЦПС, УКЗ", ""),
    ("3.1", "Требования к мониторингу: узлы, сервисы, метрики, пороги, получатели оповещений", 3, ["0.1"],
     "ЦПС (отв.), УИСО", "Перечень объектов мониторинга"),
    ("3.2", "Матрица сетевого взаимодействия (10050/10051, доступ к СУБД, API), согласование", 2, ["3.1"],
     "ЦПС (отв.), УИСО, УКЗ (согласование)", "Согласованная матрица"),
    ("3.3", "Настройка правил МСЭ", 3, ["3.2"],
     "УИСО (отв.), УКЗ", "Доступы открыты"),
    ("3.4", "Серверная часть: ресурсы ВМ, Zabbix server / proxy, БД, веб-интерфейс, резервное копирование", 5,
     ["3.1"], "УИСО", "Сервер Zabbix готов"),
    ("3.5", "Сервисные УЗ: агент в ОС, мониторинговые пользователи СУБД, etcd, Angie API; согласование прав", 3,
     ["3.1"], "ЦПС (создание), УИСО (доменные УЗ), УКЗ (согласование)", "УЗ созданы"),
    ("3.6", "Расстановка Zabbix Agent 2 на узлы ландшафта, шифрование PSK/TLS", 3, ["3.3", "3.4", "3.5"],
     "ЦПС (отв.), УИСО", "Агенты на всех узлах"),
    ("3.7", "Шаблоны и триггеры: ОС, СУБД, etcd, Angie, приложения, сроки сертификатов и лицензий", 5, ["3.6"],
     "ЦПС, УИСО", "Мониторинг настроен"),
    ("3.8", "Оповещения, эскалации, дашборды", 2, ["3.7"],
     "УИСО (отв.), ЦПС", "Оповещения работают"),
    ("3.9", "Опытная эксплуатация, корректировка порогов", 10, ["3.8"],
     "УИСО (отв.), ЦПС", "Протокол ОЭ"),
    ("3.10", "Приёмка, регламент реагирования на события мониторинга", 2, ["3.9"],
     "УИСО (отв.), ЦПС, УКЗ", "Регламент, акт"),

    ("4", "Итоговая приёмка, закрытие плана", 1, ["1.8", "2.8", "3.10"],
     "ЦПС, УИСО, УКЗ", "Протокол закрытия"),
]

HEADERS = ["№", "Задача", "Длительность, раб. дн.", "Начало", "Окончание",
           "Ответственные", "% готовности", "Предшествующие", "Результат"]


def next_workday(d, n=1):
    step = 1 if n >= 0 else -1
    while n:
        d += dt.timedelta(days=step)
        if d.weekday() < 5 and d not in HOLIDAYS:
            n -= step
    return d


def workdays_between(a, b):
    count, d = 0, a
    while d <= b:
        if d.weekday() < 5 and d not in HOLIDAYS:
            count += 1
        d += dt.timedelta(days=1)
    return count


def compute_schedule():
    dates = {}
    for num, _, dur, preds, *_ in TASKS:
        if dur is None:
            continue
        start = next_workday(max(dates[p][1] for p in preds)) if preds else PROJECT_START
        dates[num] = (start, next_workday(start, dur - 1))
    for num, _, dur, *_ in TASKS:
        if dur is None:
            kids = [v for k, v in dates.items() if k.startswith(num + ".")]
            dates[num] = (min(s for s, _ in kids), max(e for _, e in kids))
    return dates


def children(num):
    """Прямые подзадачи этапа (без вложенных уровней)."""
    depth = num.count(".") + 1
    return [t[0] for t in TASKS if t[0].startswith(num + ".") and t[0].count(".") == depth]


def build_xlsx(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "План работ"

    hol = wb.create_sheet("Праздники")
    hol["A1"] = "Нерабочие праздничные дни"
    hol["A1"].font = Font(bold=True)
    for i, d in enumerate(HOLIDAYS, start=2):
        hol.cell(row=i, column=1, value=d).number_format = "DD.MM.YYYY"
    hol.column_dimensions["A"].width = 28
    hol_ref = f"Праздники!$A$2:$A${len(HOLIDAYS) + 1}"

    ws["A1"] = "План работ по поддержанию работоспособности системы"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Дата начала работ:"
    ws["C2"] = PROJECT_START
    ws["C2"].number_format = "DD.MM.YYYY"
    ws["C2"].font = Font(bold=True, color="C00000")
    ws["D2"] = "← при изменении все даты пересчитаются"
    ws["D2"].font = Font(italic=True, color="808080")

    header_row = 4
    thin = Side(style="thin", color="A6A6A6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    row_of = {t[0]: header_row + 1 + i for i, t in enumerate(TASKS)}
    stage_fill = PatternFill("solid", fgColor="D9E1F2")
    group_fill = PatternFill("solid", fgColor="EDEDED")

    for num, name, dur, preds, resp, result in TASKS:
        r = row_of[num]
        is_group = dur is None
        is_stage = "." not in num
        ws.cell(row=r, column=1, value=num)
        ws.cell(row=r, column=2, value=("    " * num.count(".")) + name)
        ws.cell(row=r, column=6, value=resp)
        ws.cell(row=r, column=8, value=", ".join(preds))
        ws.cell(row=r, column=9, value=result)

        if is_group:
            kids = [row_of[k] for k in children(num)]
            starts = ",".join(f"D{k}" for k in kids)
            ends = ",".join(f"E{k}" for k in kids)
            durs = ",".join(f"C{k}" for k in kids)
            ws.cell(row=r, column=4, value=f"=MIN({starts})")
            ws.cell(row=r, column=5, value=f"=MAX({ends})")
            ws.cell(row=r, column=3, value=f"=NETWORKDAYS(D{r},E{r},{hol_ref})")
            weighted = "+".join(f"C{k}*G{k}" for k in kids)
            ws.cell(row=r, column=7, value=f"=({weighted})/SUM({durs})")
        else:
            ws.cell(row=r, column=3, value=dur)
            if preds:
                prev = ",".join(f"E{row_of[p]}" for p in preds)
                ws.cell(row=r, column=4, value=f"=WORKDAY(MAX({prev}),1,{hol_ref})")
            else:
                ws.cell(row=r, column=4, value="=$C$2")
            ws.cell(row=r, column=5, value=f"=WORKDAY(D{r},C{r}-1,{hol_ref})")
            ws.cell(row=r, column=7, value=0)

        for c in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal="center" if c in (1, 3, 4, 5, 7) else "left")
            if is_group:
                cell.font = Font(bold=True)
                cell.fill = stage_fill if is_stage else group_fill
        ws.cell(row=r, column=4).number_format = "DD.MM.YYYY"
        ws.cell(row=r, column=5).number_format = "DD.MM.YYYY"
        ws.cell(row=r, column=7).number_format = "0%"

    last = header_row + len(TASKS)
    total = last + 1
    ws.cell(row=total, column=2, value="ИТОГО по плану").font = Font(bold=True)
    ws.cell(row=total, column=4, value=f"=MIN(D{header_row + 1}:D{last})")
    ws.cell(row=total, column=5, value=f"=MAX(E{header_row + 1}:E{last})")
    ws.cell(row=total, column=3, value=f"=NETWORKDAYS(D{total},E{total},{hol_ref})")
    top = [row_of[t[0]] for t in TASKS if "." not in t[0]]
    ws.cell(row=total, column=7,
            value="=(" + "+".join(f"C{k}*G{k}" for k in top) + ")/(" + "+".join(f"C{k}" for k in top) + ")")
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=total, column=c)
        cell.font = Font(bold=True)
        cell.border = border
        cell.fill = PatternFill("solid", fgColor="FFF2CC")
        cell.alignment = Alignment(horizontal="center" if c in (1, 3, 4, 5, 7) else "left", vertical="center")
    ws.cell(row=total, column=4).number_format = "DD.MM.YYYY"
    ws.cell(row=total, column=5).number_format = "DD.MM.YYYY"
    ws.cell(row=total, column=7).number_format = "0%"

    ws.conditional_formatting.add(
        f"G{header_row + 1}:G{total}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"),
    )

    notes = [
        "Примечания:",
        "1. Длительность — в рабочих днях; выходные и праздники (лист «Праздники») исключаются.",
        "2. % готовности заполняется только по подзадачам; по этапам и итогу считается автоматически (взвешенно по длительности).",
        "3. «(отв.)» — ответственный за результат; остальные — соисполнители / согласующие.",
        "4. Исполнитель правил МСЭ указан как УИСО — уточнить, если МСЭ администрирует УКЗ.",
        "5. etcd — свободное ПО: уточнить у подрядчика, речь о лицензии коммерческой сборки или о перевыпуске сертификатов etcd.",
        "6. Срок получения лицензии Angie PRO (10 дн.) зависит от закупочных процедур — уточнить у ЦПС.",
    ]
    for i, text in enumerate(notes):
        cell = ws.cell(row=total + 2 + i, column=1, value=text)
        cell.font = Font(bold=(i == 0), italic=(i > 0), color="404040")

    widths = [7, 70, 13, 12, 12, 34, 12, 16, 28]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[header_row].height = 32
    ws.freeze_panes = ws.cell(row=header_row + 1, column=3)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(HEADERS))}{last}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(path)


def build_md(path, dates):
    fmt = "%d.%m.%Y"
    lines = [
        "# План работ по поддержанию работоспособности системы",
        "",
        f"Дата начала: {PROJECT_START.strftime(fmt)}. Длительность — в рабочих днях. "
        "«(отв.)» — ответственный за результат.",
        "",
        "| № | Задача | Длит., раб. дн. | Начало | Окончание | Ответственные | % готовности |",
        "|---|---|---|---|---|---|---|",
    ]
    for num, name, dur, _, resp, _ in TASKS:
        s, e = dates[num]
        d = dur if dur is not None else workdays_between(s, e)
        row = [num, name, str(d), s.strftime(fmt), e.strftime(fmt), resp, "0%"]
        if dur is None:
            row = [f"**{x}**" for x in row]
        lines.append("| " + " | ".join(row) + " |")
    s = min(v[0] for v in dates.values())
    e = max(v[1] for v in dates.values())
    lines.append(f"| | **ИТОГО** | **{workdays_between(s, e)}** | **{s.strftime(fmt)}** | "
                 f"**{e.strftime(fmt)}** | | **0%** |")
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    here = Path(__file__).parent
    schedule = compute_schedule()
    build_xlsx(here / "plan.xlsx")
    build_md(here / "plan.md", schedule)
