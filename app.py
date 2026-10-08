from flask import (Flask, render_template, request, redirect, url_for,
                   send_file, flash)
from werkzeug.utils import secure_filename
from datetime import datetime

from integration_files import integrate_uploaded_auto
from nfgo_report import build_report, export_excel as nfgo_xlsx, export_xml
from refresh_plan import build_refresh_plan, export_excel as plan_xlsx
import storage

import traceback

app = Flask(__name__)
app.secret_key = "go-museum-secret-key"
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024

# Перехватчик всех ошибок
@app.errorhandler(Exception)
def handle_exception(e):

    tb = traceback.format_exc()
    ts = datetime.now().strftime("%H:%M:%S")

    STATE.log.append(f"[{ts}] Ошибка: {type(e).__name__}: {e}")
    STATE.log.append("Проверьте: правильность формата файлов, "
                     "наличие колонок, отсутствие повреждённых файлов.")
    STATE.log.append("Техническая информация:")
    for line in tb.split("\n")[-6:]:
        if line.strip():
            STATE.log.append(f"  {line}")

    flash(f"Произошла ошибка: {e}. Подробности в журнале обработки.",
          "danger")
    return redirect(url_for("index"))

# Состояние (в памяти только для UI-логов и готовых отчётов)
class State:
    def __init__(self):
        self.log = []

STATE = State()


def log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    STATE.log.append(f"[{ts}] {msg}")


# Утилита: получить данные из БД + построить отчёты
def _get_data_and_reports():
    """Читает данные из SQLite и строит отчёты «на лету».
       Возвращает (data, report, plan, done)."""
    data = storage.load_data()
    if data is None:
        return None, None, None, False
    report = build_report(data)
    plan = build_refresh_plan(data["siz_cards"])
    return data, report, plan, True

# Загрузка инетграция

@app.route("/")
def index():
    data = storage.load_data()
    done = data is not None
    return render_template(
        "index.html",
        integration_done=done,
        log=STATE.log[-30:],
        counts={
            "employees": len(data["employees"]) if data else 0,
            "siz":       len(data["siz_cards"]) if data else 0,
            "nfgo":      len(data["nfgo_members"]) if data else 0,
            "training":  len(data["training"]) if data else 0,
        },
    )


@app.route("/integration/run", methods=["POST"])
def run_integration():
    STATE.log.clear()
    log("Начата обработка загруженных файлов")

    files_in = request.files.getlist("files")
    files = []
    for f in files_in:
        if f and f.filename:
            files.append((secure_filename(f.filename), f.read()))

    if not files:
        log("  Ни одного файла не получено.")
        flash("Не загружено ни одного файла. Выберите файлы в поле загрузки.",
              "warning")
        return redirect(url_for("index"))

    log(f"  Получено файлов: {len(files)}")
    for fname, fbytes in files:
        log(f"    • {fname} ({len(fbytes)} байт)")

    # Интеграция
    new_data = integrate_uploaded_auto(files, STATE.log)

    # Сохранение в БД (upsert)
    log("Сохранение результатов в базу данных…")
    storage.save_integrated_data(new_data)

    total_data = storage.load_data()
    log(f"В базе сейчас: "
        f"сотрудников={len(total_data['employees'])}, "
        f"СИЗ={len(total_data['siz_cards'])}, "
        f"НФГО={len(total_data['nfgo_members'])}, "
        f"обучения={len(total_data['training'])}")

    log("Интеграция завершена.")
    flash(f"Обработано файлов: {len(files)}. Данные сохранены в базу.",
          "success")
    return redirect(url_for("index"))


@app.route("/integration/reset", methods=["POST"])
def reset_integration():
    storage.clear_all()
    STATE.log.clear()
    flash("Все данные удалены из базы", "info")
    return redirect(url_for("index"))

# просмотр данных

@app.route("/data")
def data_view():
    data, _, _, done = _get_data_and_reports()
    return render_template(
        "data_view.html",
        data=data,
        integration_done=done,
        log=STATE.log[-10:],
    )

# отчет по НФГО

@app.route("/nfgo")
def nfgo_page():
    _, report, _, done = _get_data_and_reports()
    return render_template(
        "nfgo.html",
        report=report,
        integration_done=done,
        log=STATE.log[-10:],
    )


@app.route("/nfgo/export/xlsx")
def nfgo_export_xlsx():
    _, report, _, done = _get_data_and_reports()
    if not done:
        flash("Сначала загрузите и обработайте файлы", "warning")
        return redirect(url_for("index"))
    return send_file(nfgo_xlsx(report), as_attachment=True)

# план освежения

@app.route("/plan")
def plan_page():
    _, _, plan, done = _get_data_and_reports()
    total = sum(p["total"] for p in plan) if plan else 0
    return render_template(
        "plan.html",
        plan=plan,
        total=total,
        integration_done=done,
        log=STATE.log[-10:],
    )


@app.route("/plan/export/xlsx")
def plan_export_xlsx():
    _, _, plan, done = _get_data_and_reports()
    if not done:
        flash("Сначала загрузите и обработайте файлы", "warning")
        return redirect(url_for("index"))
    return send_file(plan_xlsx(plan), as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)