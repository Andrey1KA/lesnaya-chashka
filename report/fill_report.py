"""
Заполнение еженедельного отчёта «Лесная Чашка» за недели 1–4 (01.12–27.12.2026).

Правила (из задания):
- время московское;
- неделя чека — по дате загрузки на сайт (promo_receipt.created_at);
- регистрации — по date_joined;
- сотрудники и служебные аккаунты в отчёт для клиента не входят.

Запуск из корня проекта:
  backend\\.venv\\Scripts\\python.exe report\\fill_report.py
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import psycopg
from openpyxl import load_workbook
from zoneinfo import ZoneInfo

MSK = ZoneInfo("Europe/Moscow")
PERIOD_START = date(2026, 12, 1)
PERIOD_END = date(2026, 12, 27)

WEEKS = [
    (1, date(2026, 12, 1), date(2026, 12, 6)),
    (2, date(2026, 12, 7), date(2026, 12, 13)),
    (3, date(2026, 12, 14), date(2026, 12, 20)),
    (4, date(2026, 12, 21), date(2026, 12, 27)),
]

REJECT_COLUMNS = {
    "Чек уже зарегистрирован": "D",
    "Нет фото чека": "E",
    "Магазин не участвует в акции": "F",
    "В чеке нет акционного продукта": "G",
    "Недостаточная итоговая сумма по акционным товарам в чеке": "H",
    "Чек не прошёл проверку фискальных данных": "I",
    "Чек сфотографирован неполностью": "J",
    "Чек плохого качества": "K",
    "На фото более одного чека": "L",
    "Дата покупки в чеке не соответствует правилам": "M",
}

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "report" / "report.xlsx"

# Служебные аккаунты: сотрудник и тестовые ящики агентства / прошлого задания.
EXCLUDE_SQL = """
    u.is_staff
    OR lower(split_part(u.email, '@', 2)) <> 'mail.example'
"""


def msk_date(dt: datetime) -> date:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo("UTC"))
    return dt.astimezone(MSK).date()


def week_of(d: date) -> int | None:
    for num, start, end in WEEKS:
        if start <= d <= end:
            return num
    return None


def fetch():
    conn = psycopg.connect(
        host="localhost",
        port=5432,
        dbname="lesnaya_chashka_v2",
        user="lesnaya",
        password="lesnaya-local-2026",
    )
    with conn:
        with conn.cursor() as cur:
            cur.execute(
                f"""
                SELECT id, email, is_staff, participant_code, prev_promo_participant, date_joined
                FROM promo_user u
                WHERE NOT ({EXCLUDE_SQL})
                """
            )
            users = [
                {
                    "id": row[0],
                    "email": row[1],
                    "is_staff": row[2],
                    "code": row[3],
                    "prev": row[4],
                    "joined": msk_date(row[5]),
                }
                for row in cur.fetchall()
            ]

            cur.execute(
                f"""
                SELECT r.id, r.user_id, r.status, r.reject_reason, r.created_at, r.points, r.prize_amount
                FROM promo_receipt r
                JOIN promo_user u ON u.id = r.user_id
                WHERE NOT ({EXCLUDE_SQL})
                """
            )
            receipts = [
                {
                    "id": row[0],
                    "user_id": row[1],
                    "status": row[2],
                    "reason": row[3] or "",
                    "day": msk_date(row[4]),
                    "points": row[5] or 0,
                    "prize": row[6] or 0,
                }
                for row in cur.fetchall()
            ]

            cur.execute(
                f"""
                SELECT r.id, p.pack_format, i.quantity
                FROM promo_receiptitem i
                JOIN promo_receipt r ON r.id = i.receipt_id
                JOIN promo_product p ON p.id = i.product_id
                JOIN promo_user u ON u.id = r.user_id
                WHERE r.status = 'accepted'
                  AND NOT ({EXCLUDE_SQL})
                """
            )
            items = cur.fetchall()

            cur.execute(
                f"""
                SELECT u.id, u.email, u.is_staff, split_part(lower(u.email), '@', 2) AS domain,
                       COUNT(r.id) AS receipts
                FROM promo_user u
                LEFT JOIN promo_receipt r ON r.user_id = u.id
                WHERE {EXCLUDE_SQL}
                GROUP BY u.id
                ORDER BY u.id
                """
            )
            excluded = cur.fetchall()

    conn.close()
    return users, receipts, items, excluded


def main():
    users, receipts, items, excluded = fetch()
    users_by_id = {u["id"]: u for u in users}

    first_receipt_day = {}
    for r in receipts:
        prev = first_receipt_day.get(r["user_id"])
        if prev is None or r["day"] < prev:
            first_receipt_day[r["user_id"]] = r["day"]

    weekly = {
        n: {"D": 0, "E": set(), "F": 0, "H": 0, "J": 0, "K": 0, "L": 0, "reject": defaultdict(int)}
        for n, _, _ in WEEKS
    }
    daily_reg = defaultdict(int)
    daily_accepted = defaultdict(int)

    for u in users:
        if PERIOD_START <= u["joined"] <= PERIOD_END:
            daily_reg[u["joined"]] += 1
            w = week_of(u["joined"])
            if w:
                weekly[w]["D"] += 1

    packs_by_receipt = defaultdict(int)
    packs_by_format = defaultdict(int)
    receipt_by_id = {r["id"]: r for r in receipts}
    for receipt_id, pack_format, qty in items:
        r = receipt_by_id.get(receipt_id)
        if r is None or not (PERIOD_START <= r["day"] <= PERIOD_END):
            continue
        packs_by_receipt[receipt_id] += qty
        packs_by_format[pack_format] += qty

    for r in receipts:
        if not (PERIOD_START <= r["day"] <= PERIOD_END):
            continue
        w = week_of(r["day"])
        if not w:
            continue
        weekly[w]["J"] += 1
        weekly[w]["E"].add(r["user_id"])
        if r["status"] == "accepted":
            weekly[w]["K"] += 1
            weekly[w]["L"] += packs_by_receipt.get(r["id"], 0)
            daily_accepted[r["day"]] += 1
        elif r["status"] == "rejected":
            weekly[w]["reject"][r["reason"]] += 1

    for user_id, first_day in first_receipt_day.items():
        w = week_of(first_day)
        if not w:
            continue
        weekly[w]["F"] += 1
        if users_by_id[user_id]["prev"]:
            weekly[w]["H"] += 1

    unique_with_receipt = {r["user_id"] for r in receipts if PERIOD_START <= r["day"] <= PERIOD_END}
    accepted_total = sum(1 for r in receipts if r["status"] == "accepted" and PERIOD_START <= r["day"] <= PERIOD_END)
    packs_total = sum(packs_by_format.values())

    # Participants with at least one receipt in the period.
    stats = {}
    for r in receipts:
        if not (PERIOD_START <= r["day"] <= PERIOD_END):
            continue
        s = stats.setdefault(
            r["user_id"],
            {"uploaded": 0, "accepted": 0, "points": 0, "prize": 0},
        )
        s["uploaded"] += 1
        if r["status"] == "accepted":
            s["accepted"] += 1
            s["points"] += r["points"]
            s["prize"] += r["prize"]

    participants = []
    for user_id, s in stats.items():
        u = users_by_id[user_id]
        participants.append(
            {
                "code": u["code"] or "",
                "joined": u["joined"],
                **s,
                "prev": "да" if u["prev"] else "нет",
            }
        )
    participants.sort(key=lambda p: (p["code"], p["joined"]))

    wb = load_workbook(XLSX)
    total = wb["Тотал"]
    for num, start, end in WEEKS:
        row = 8 + num  # 9..12
        w = weekly[num]
        total[f"D{row}"] = w["D"]
        total[f"E{row}"] = len(w["E"])
        total[f"F{row}"] = w["F"]
        total[f"H{row}"] = w["H"]
        total[f"J{row}"] = w["J"]
        total[f"K{row}"] = w["K"]
        total[f"L{row}"] = w["L"]

    dynamics = wb["Динамика ежедневная"]
    for row in range(4, 67):
        cell_date = dynamics[f"A{row}"].value
        if not isinstance(cell_date, datetime):
            continue
        d = cell_date.date()
        if PERIOD_START <= d <= PERIOD_END:
            dynamics[f"C{row}"] = daily_reg.get(d, 0)
            dynamics[f"D{row}"] = daily_accepted.get(d, 0)

    reasons = wb["Причины отклонения чеков"]
    for num, start, end in WEEKS:
        row = 3 + num  # 4..7
        for reason, col in REJECT_COLUMNS.items():
            reasons[f"{col}{row}"] = weekly[num]["reject"].get(reason, 0)

    sku = wb["SKU"]
    sku["B3"] = packs_by_format.get("bags20", 0)
    sku["B4"] = packs_by_format.get("bags100", 0)
    sku["B5"] = packs_by_format.get("loose75", 0)

    people = wb["Участники"]
    for offset, p in enumerate(participants):
        row = 4 + offset
        people[f"A{row}"] = p["code"]
        people[f"B{row}"] = p["joined"].strftime("%d.%m.%Y")
        people[f"C{row}"] = p["uploaded"]
        people[f"D{row}"] = p["accepted"]
        people[f"E{row}"] = p["points"]
        people[f"F{row}"] = p["prize"]
        people[f"G{row}"] = p["prev"]

    wb.save(XLSX)

    print("saved", XLSX)
    print("excluded")
    for row in excluded:
        print(" ", row)
    print()
    print("weeks")
    for num, start, end in WEEKS:
        w = weekly[num]
        print(
            f"  W{num} {start}..{end}  D={w['D']} E={len(w['E'])} F={w['F']} H={w['H']} "
            f"J={w['J']} K={w['K']} L={w['L']}"
        )
    print()
    print("unique participants with receipts (4 weeks):", len(unique_with_receipt))
    print("accepted receipts:", accepted_total)
    print("packs:", packs_total)
    print("participants rows:", len(participants))
    print("regs 01-27:", sum(daily_reg.values()))
    print("sku", dict(packs_by_format))


if __name__ == "__main__":
    main()
