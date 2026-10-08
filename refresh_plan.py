from datetime import date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from config import SIZ_PRICES, DELIVERY_SCHEDULE, OUTPUT_DIR


def build_refresh_plan(siz_cards: list[dict]) -> list[dict]:
    today = date.today()
    current_year = today.year
    today_iso = today.isoformat()

    plan = []
    for name, price in SIZ_PRICES.items():
        rows = [r for r in siz_cards
                if r["name"] == name and r["expiry_date"] >= today_iso]
        if not rows:
            continue
        years = sorted(int(r["expiry_date"][:4]) for r in rows)
        year = max(years[0], current_year)
        qty = sum(r["quantity"] for r in rows
                  if int(r["expiry_date"][:4]) == year)
        if qty == 0:
            continue
        plan.append({
            "year":       year,
            "name":       name,
            "quantity":   qty,
            "unit_price": price,
            "total":      qty * price,
            "delivery":   DELIVERY_SCHEDULE.get(name, "—"),
        })
    plan.sort(key=lambda x: (x["year"], x["name"]))
    return plan


def export_excel(plan: list[dict]) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "План освежения СИЗ"

    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    hdr_fill = PatternFill("solid", fgColor="D9E1F2")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = "ПРИЛОЖЕНИЕ №11 к Приказу №___"
    ws["A1"].font = Font(bold=True, italic=True)
    ws["A2"] = "ПЛАН ОСВЕЖЕНИЯ (ЗАКУПОК) СИЗ НА 2026–2030 ГГ."
    ws["A2"].font = Font(bold=True, size=14)
    ws.merge_cells("A2:F2"); ws["A2"].alignment = center

    headers = ["Год закупки", "Наименование СИЗ",
               "Кол-во к списанию (шт.)", "Прогнозная стоимость (ед., руб.)",
               "Итог по бюджету (руб.)", "Срок поставки"]
    for c, txt in enumerate(headers, 1):
        cell = ws.cell(row=4, column=c, value=txt)
        cell.font = bold; cell.alignment = center
        cell.fill = hdr_fill; cell.border = border

    total = 0
    for i, p in enumerate(plan, 5):
        vals = [p["year"], p["name"], p["quantity"],
                p["unit_price"], p["total"], p["delivery"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = border; cell.alignment = center
        total += p["total"]

    last = 5 + len(plan)
    ws.cell(row=last, column=5, value=f"ИТОГО: {total} руб.").font = bold
    ws.cell(row=last + 2, column=1,
            value="Источники финансирования:").font = bold
    ws.cell(row=last + 3, column=1,
            value="• Бюджет организации (статья «Мобилизационная подготовка»)")
    ws.cell(row=last + 4, column=1, value="• Резервный фонд")
    ws.cell(row=last + 5, column=1,
            value="• Средства от утилизации списанных СИЗ")
    ws.cell(row=last + 7, column=1,
            value="Подписано УКЭП руководителя ГО.").font = bold

    for col, w in zip("ABCDEF", [12, 30, 22, 28, 22, 14]):
        ws.column_dimensions[col].width = w

    path = OUTPUT_DIR / "Пункт8_План_освежения_СИЗ.xlsx"
    wb.save(path)
    return path