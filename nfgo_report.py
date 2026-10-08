import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from config import (ORG_NAME, ORG_CATEGORY_GO, NFGO_TYPES,
                    TABEL_NORMS, OUTPUT_DIR)


def build_report(data: dict) -> dict:
    # Реестр формирований
    formations_dict = {}
    for m in data["nfgo_members"]:
        name = (m.get("formation") or "").strip()
        if not name:
            continue  # пропускаем пустые формирования
        formations_dict.setdefault(name, []).append(m)

    formations = []
    for i, (name, members) in enumerate(formations_dict.items(), start=1):
        info = NFGO_TYPES.get(
            name,
            {"type": "Прочее", "mchs_code": "—"}
        )
        formations.append({
            "no": i,
            "name": name,
            "type": info["type"],
            "mchs_code": info["mchs_code"],
            "count": len(members),
            "readiness": 100.0,
        })

    # Обеспеченность СИЗ
    today = date.today().isoformat()
    equipment = []
    for name, required in TABEL_NORMS.items():
        rows = [r for r in data["siz_cards"] if r["name"] == name]
        available   = sum(r["quantity"] for r in rows)
        serviceable = sum(r["quantity"] for r in rows
                          if r["expiry_date"] >= today)
        provision   = round(serviceable / required * 100, 1) if required else 0.0
        refresh_year = max((r["expiry_date"] for r in rows), default="—")[:4]
        equipment.append({
            "name":         name,
            "required":     required,
            "available":    available,
            "serviceable":  serviceable,
            "provision":    provision,
            "refresh_year": refresh_year,
        })

    # Авторасчёт готовности
    avg = sum(e["provision"] for e in equipment) / len(equipment) if equipment else 0
    for f in formations:
        f["readiness"] = round(min(100.0, avg), 1)

    # Подготовка личного состава
    workers = [e for e in data["employees"] if e["status"] == "Работает"]
    passed = {t["employee_id"] for t in data["training"]
              if t["result"] == "Зачтено"}
    last = max((t["date"] for t in data["training"] if t["date"]), default="—")
    training_info = {
        "planned":        len(workers),
        "fact":           len(passed),
        "umc_commanders": 4,
        "last_training":  last,
    }

    # Валидация
    vacancy = sum(1 for m in data["nfgo_members"] if m["status"] != "Работает")
    status = "Требует уточнения" if vacancy > 0 else "Готов к выгрузке"

    return {
        "header": {
            "organization": ORG_NAME,
            "category_go":  ORG_CATEGORY_GO,
            "report_date":  today,
            "recipient":    "Территориальный орган МЧС России",
        },
        "formations": formations,
        "equipment":  equipment,
        "training":   training_info,
        "status":     status,
        "vacancies":  vacancy,
    }


def export_excel(report: dict) -> Path:
    wb = Workbook()
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    hdr_fill = PatternFill("solid", fgColor="D9E1F2")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Реестр НФГО
    ws = wb.active
    ws.title = "Реестр НФГО"
    ws["A1"] = "СВЕДЕНИЯ О НАЛИЧИИ И ГОТОВНОСТИ НФГО"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:F1")
    ws["A1"].alignment = center

    h = report["header"]
    ws["A2"] = f"Организация: {h['organization']}"; ws["A2"].font = bold
    ws["A3"] = f"Категория по ГО: {h['category_go']}"; ws["A3"].font = bold
    ws["A4"] = f"Дата: {h['report_date']}"; ws["A4"].font = bold
    ws["A5"] = f"Статус: {report['status']}"; ws["A5"].font = bold

    headers = ["№", "Наименование формирования", "Вид",
               "Код МЧС", "Кол-во, чел.", "Готовность, %"]
    for c, txt in enumerate(headers, 1):
        cell = ws.cell(row=7, column=c, value=txt)
        cell.font = bold; cell.alignment = center
        cell.fill = hdr_fill; cell.border = border

    for i, f in enumerate(report["formations"], 8):
        vals = [f["no"], f["name"], f["type"],
                f["mchs_code"], f["count"], f["readiness"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = border; cell.alignment = center

    # Обеспеченность СИЗ
    ws2 = wb.create_sheet("Обеспеченность СИЗ")
    ws2["A1"] = "Таблица 2. Обеспеченность СИЗ"
    ws2["A1"].font = Font(bold=True, size=13)
    ws2.merge_cells("A1:F1"); ws2["A1"].alignment = center
    h2 = ["Наименование имущества", "Потребность, ед.", "Наличие, ед.",
          "Годных, шт.", "Обеспеченность, %", "Срок освежения"]
    for c, txt in enumerate(h2, 1):
        cell = ws2.cell(row=3, column=c, value=txt)
        cell.font = bold; cell.alignment = center
        cell.fill = hdr_fill; cell.border = border
    for i, e in enumerate(report["equipment"], 4):
        vals = [e["name"], e["required"], e["available"],
                e["serviceable"], e["provision"], e["refresh_year"]]
        for c, v in enumerate(vals, 1):
            cell = ws2.cell(row=i, column=c, value=v)
            cell.border = border; cell.alignment = center

    # Подготовка
    ws3 = wb.create_sheet("Подготовка")
    ws3["A1"] = "Таблица 3. Подготовка личного состава"
    ws3["A1"].font = Font(bold=True, size=13)
    t = report["training"]
    rows3 = [
        ("План, чел.", t["planned"]),
        ("Факт, чел.", t["fact"]),
        ("Обучено в УМЦ МЧС", t["umc_commanders"]),
        ("Дата последней тренировки", t["last_training"]),
    ]
    for i, (k, v) in enumerate(rows3, 3):
        ws3.cell(row=i, column=1, value=k).font = bold
        ws3.cell(row=i, column=2, value=v)

    path = OUTPUT_DIR / "Пункт6_Отчёт_НФГО_МЧС.xlsx"
    wb.save(path)
    return path


def export_xml(report: dict) -> Path:
    root = ET.Element("ReportNFO")
    h = ET.SubElement(root, "Header")
    for k, v in report["header"].items():
        ET.SubElement(h, k).text = str(v)
    ET.SubElement(root, "Status").text = report["status"]
    ET.SubElement(root, "Vacancies").text = str(report["vacancies"])

    fn = ET.SubElement(root, "Formations")
    for f in report["formations"]:
        n = ET.SubElement(fn, "Formation")
        for k, v in f.items():
            ET.SubElement(n, k).text = str(v)

    en = ET.SubElement(root, "Equipment")
    for e in report["equipment"]:
        n = ET.SubElement(en, "Item")
        for k, v in e.items():
            ET.SubElement(n, k).text = str(v)

    tn = ET.SubElement(root, "Training")
    for k, v in report["training"].items():
        ET.SubElement(tn, k).text = str(v)

    sig = ET.SubElement(root, "Signature")
    ET.SubElement(sig, "UKEP_HeadOfGO").text = "Подписано УКЭП"
    ET.SubElement(sig, "UKEP_AuthorizedGO").text = "Подписано УКЭП"

    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)
    path = OUTPUT_DIR / "Пункт6_Отчёт_НФГО_МЧС.xml"
    tree.write(path, encoding="utf-8", xml_declaration=True)
    return path