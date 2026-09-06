
# ============================================================
# TRAINING SYSTEM v4.7 — WEB APP
# ============================================================
# Мобільний веб-інтерфейс поверх існуючого Training System.
# Ядро v2.9 не переписується: цей файл підключає його функції.
# ============================================================

from pathlib import Path
from flask import Flask, request, redirect, url_for, render_template_string, jsonify
from datetime import date

# Завантажуємо існуюче ядро v2.9 як модуль.
import importlib.util

BASE_DIR = Path(__file__).resolve().parent

# Шукаємо робоче ядро автоматично.
# Це прибирає проблему, коли файл ядра має іншу назву.
CORE_CANDIDATES = [
    "training_system_v2_9_fixed.py",
    "training_system_v2_9.py",
    "training_system_v2_8.py",
]

CORE_PATH = next(
    (
        BASE_DIR / name
        for name in CORE_CANDIDATES
        if (BASE_DIR / name).exists()
    ),
    None
)

if CORE_PATH is None:
    raise FileNotFoundError(
        "Не знайдено ядро Training System. "
        "Поклади поруч із цим файлом "
        "training_system_v2_9_fixed.py "
        "або training_system_v2_9.py."
    )

spec = importlib.util.spec_from_file_location(
    "training_core",
    CORE_PATH
)

if spec is None or spec.loader is None:
    raise RuntimeError(
        f"Не вдалося підключити ядро: {CORE_PATH.name}"
    )

core = importlib.util.module_from_spec(spec)
spec.loader.exec_module(core)

# ============================================================
# DATA-SCHEMA COMPATIBILITY PATCH
# ============================================================
# v2.9 writes result rows in the same order as the result sheets
# created by ensure_result_sheets(). Some older reader functions
# expected a different column order. v3.6 uses these normalized
# readers so history/progress can see the records that were
# actually saved.
# ============================================================

def _v36_get_strength_rows(workbook):
    if "STRENGTH_RESULTS" not in workbook.sheetnames:
        return []

    sheet = workbook["STRENGTH_RESULTS"]
    rows = []

    for row in range(2, sheet.max_row + 1):
        rows.append({
            "id": sheet.cell(row, 1).value,
            "date": sheet.cell(row, 2).value,
            "workout_id": sheet.cell(row, 3).value,
            "week": sheet.cell(row, 4).value,
            "session": sheet.cell(row, 5).value,
            "exercise": sheet.cell(row, 6).value,
            "set": sheet.cell(row, 7).value,
            "weight": core.safe_float(sheet.cell(row, 8).value),
            "reps": core.safe_float(sheet.cell(row, 9).value),
            "rir": core.safe_float(sheet.cell(row, 10).value),
            "notes": sheet.cell(row, 11).value,
        })

    return rows


def _v36_get_wod_result_rows(workbook):
    if "WOD_RESULTS" not in workbook.sheetnames:
        return []

    sheet = workbook["WOD_RESULTS"]
    rows = []

    for row in range(2, sheet.max_row + 1):
        rows.append({
            "id": sheet.cell(row, 1).value,
            "date": sheet.cell(row, 2).value,
            "workout_id": sheet.cell(row, 3).value,
            "week": sheet.cell(row, 4).value,
            "session": sheet.cell(row, 5).value,
            "wod_id": sheet.cell(row, 6).value,
            "result": sheet.cell(row, 7).value,
            "rpe": core.safe_float(sheet.cell(row, 8).value),
            "notes": sheet.cell(row, 9).value,
            "format": None,
            "scaling": None,
            "time": None,
            "rounds": None,
        })

    return rows


def _v36_get_workout_rows(workbook):
    if "WORKOUTS" not in workbook.sheetnames:
        return []

    sheet = workbook["WORKOUTS"]
    rows = []

    for row in range(2, sheet.max_row + 1):
        rows.append({
            "id": sheet.cell(row, 1).value,
            "date": sheet.cell(row, 2).value,
            "week": sheet.cell(row, 3).value,
            "session": sheet.cell(row, 4).value,
            "wod_id": sheet.cell(row, 5).value,
            "wod": sheet.cell(row, 6).value,
            "format": sheet.cell(row, 7).value,
            "duration": core.safe_float(sheet.cell(row, 8).value),
            "result": sheet.cell(row, 9).value,
            "rpe": core.safe_float(sheet.cell(row, 10).value),
            "pr": sheet.cell(row, 11).value,
            "notes": sheet.cell(row, 12).value,
        })

    return rows


# Replace the readers used by the web app and core analytics.
core.get_strength_rows = _v36_get_strength_rows
core.get_wod_result_rows = _v36_get_wod_result_rows
core.get_workout_rows = _v36_get_workout_rows


app = Flask(__name__)

FILE_NAME = core.FILE_NAME

SESSION_TYPES = [
    "LOWER",
    "PUSH",
    "PULL",
    "FULL BODY"
]

CURRENT_WEEK = 1


def get_workbook():
    file_path = core.get_file_path()

    if not file_path.exists():
        raise RuntimeError(
            f"Excel-файл не знайдено: {file_path}"
        )

    try:
        workbook = core.open_training_database()
    except Exception as exc:
        raise RuntimeError(
            f"Не вдалося відкрити Excel: {file_path}\\n"
            f"Причина: {exc}"
        ) from exc

    if workbook is None:
        raise RuntimeError(
            f"Не вдалося відкрити Excel: {file_path}"
        )

    return workbook


def page(title, body):
    return render_template_string(
        """
<!doctype html>
<html lang="uk">
<head>
<meta charset="utf-8">
<meta name="viewport"
      content="width=device-width, initial-scale=1,
               maximum-scale=1">
<title>{{ title }}</title>
<style>
* { box-sizing: border-box; }
body {
    margin: 0;
    font-family: Arial, sans-serif;
    background: #f4f5f7;
    color: #171717;
}
.header {
    background: #111827;
    color: white;
    padding: 18px 16px;
    position: sticky;
    top: 0;
    z-index: 10;
}
.header h1 {
    margin: 0;
    font-size: 21px;
}
.container {
    max-width: 760px;
    margin: auto;
    padding: 14px;
}
.card {
    background: white;
    border-radius: 16px;
    padding: 16px;
    margin-bottom: 12px;
    box-shadow: 0 2px 10px rgba(0,0,0,.07);
}
.grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 10px;
}
.button {
    display: block;
    text-decoration: none;
    text-align: center;
    background: #111827;
    color: white;
    padding: 14px 10px;
    border-radius: 12px;
    font-weight: 700;
    margin-top: 8px;
}
.button.secondary {
    background: #e5e7eb;
    color: #111827;
}
.stat {
    font-size: 25px;
    font-weight: 800;
}
.muted {
    color: #6b7280;
    font-size: 13px;
}
.exercise {
    border-top: 1px solid #e5e7eb;
    padding: 12px 0;
}
.exercise:first-child {
    border-top: 0;
}
input, select, textarea {
    width: 100%;
    padding: 12px;
    border: 1px solid #d1d5db;
    border-radius: 10px;
    font-size: 16px;
    margin: 6px 0 12px;
}
label {
    font-weight: 700;
    font-size: 14px;
}
h2, h3 {
    margin-top: 0;
}
@media (max-width: 480px) {
    .grid { grid-template-columns: 1fr; }
}

.mobile-nav {
    position: fixed;
    left: 0;
    right: 0;
    bottom: 0;
    z-index: 1000;
    display: flex;
    justify-content: space-around;
    align-items: center;
    padding: 8px 6px calc(8px + env(safe-area-inset-bottom));
    background: rgba(255,255,255,.96);
    border-top: 1px solid #ddd;
    backdrop-filter: blur(8px);
}

.mobile-nav .nav-item {
    flex: 1;
    text-align: center;
    text-decoration: none;
    color: inherit;
    font-size: 12px;
    padding: 5px 2px;
}

.mobile-nav .nav-item span {
    display: block;
    font-size: 20px;
    line-height: 22px;
}

.mobile-nav .nav-item small {
    display: block;
    margin-top: 2px;
}

body {
    padding-bottom: 82px;
}

@media (min-width: 800px) {
    .mobile-nav {
        max-width: 720px;
        left: 50%;
        right: auto;
        width: 100%;
        transform: translateX(-50%);
        border-left: 1px solid #ddd;
        border-right: 1px solid #ddd;
    }
}
</style>
</head>
<body>
<div class="header">
    <h1>🔥 TRAINING SYSTEM</h1>
</div>
<div class="container">
{{ body|safe }}
</div>

        <nav class="mobile-nav">
            <a href="/" class="nav-item">
                <span>🏠</span>
                <small>Головна</small>
            </a>
            <a href="/next" class="nav-item">
                <span>🏋️</span>
                <small>Тренування</small>
            </a>
            <a href="/exercise" class="nav-item">
                <span>📈</span>
                <small>Прогрес</small>
            </a>
            <a href="/history" class="nav-item">
                <span>🏆</span>
                <small>Історія</small>
            </a>
        </nav>
</body>
</html>
        """,
        title=title,
        body=body
    )



# ============================================================
# FLEXIBLE TRAINING SEQUENCE — v3.2
# ============================================================

TRAINING_SEQUENCE = [
    "LOWER",
    "PUSH",
    "PULL",
    "FULL BODY"
]


def get_last_completed_session(workbook):
    """
    Визначає останній фактично виконаний тип тренування.
    Не використовує день тижня.
    """
    rows = core.get_workout_rows(workbook)

    if not rows:
        return None

    latest = None
    latest_key = None

    for row in rows:
        session = str(row.get("session") or "").strip().upper()

        if session not in TRAINING_SEQUENCE:
            continue

        date_value = str(row.get("date") or "")
        workout_id = core.safe_float(row.get("id"))

        key = (
            date_value,
            workout_id if workout_id is not None else -1
        )

        if latest_key is None or key > latest_key:
            latest_key = key
            latest = session

    return latest


def get_next_session(workbook):
    """
    Повертає наступне тренування за гнучкою чергою.
    Якщо історії немає — починаємо з LOWER.
    """
    last_session = get_last_completed_session(workbook)

    if last_session is None:
        return TRAINING_SEQUENCE[0]

    index = TRAINING_SEQUENCE.index(last_session)

    return TRAINING_SEQUENCE[
        (index + 1) % len(TRAINING_SEQUENCE)
    ]


def get_last_workout_summary(workbook):
    rows = core.get_workout_rows(workbook)

    if not rows:
        return None

    candidates = []

    for row in rows:
        session = str(row.get("session") or "").strip().upper()

        if session not in TRAINING_SEQUENCE:
            continue

        workout_id = core.safe_float(row.get("id"))

        candidates.append(
            (
                str(row.get("date") or ""),
                workout_id if workout_id is not None else -1,
                row
            )
        )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: (x[0], x[1])
    )

    return candidates[-1][2]


@app.route("/")
def home():
    workbook = get_workbook()

    info = core.get_week_info(workbook, CURRENT_WEEK)

    workouts = core.get_workout_rows(workbook)
    strength = core.get_strength_rows(workbook)
    wods = core.get_wod_result_rows(workbook)

    phase = core.value_or_dash(
        info["phase"] if info else None
    )
    target_rir = core.value_or_dash(
        info["target_rir"] if info else None
    )

    next_session = get_next_session(workbook)
    last_workout = get_last_workout_summary(workbook)

    if last_workout:
        last_session = core.value_or_dash(
            last_workout.get("session")
        )
        last_date = core.value_or_dash(
            last_workout.get("date")
        )
        last_workout_id = core.safe_float(
            last_workout.get("id")
        )
        last_link = (
            f'/history/{int(last_workout_id)}'
            if last_workout_id is not None
            else '/history'
        )
    else:
        last_session = "Ще немає"
        last_date = "—"
        last_link = "/history"

    avg_rpe = core.get_average([
        row["rpe"]
        for row in workouts
        if row.get("rpe") is not None
    ])

    avg_rir = core.get_average([
        row["rir"]
        for row in strength
        if row.get("rir") is not None
    ])

    recent = sorted(
        [
            row for row in workouts
            if core.safe_float(row.get("id")) is not None
        ],
        key=lambda row: (
            str(row.get("date") or ""),
            core.safe_float(row.get("id")) or -1
        ),
        reverse=True
    )[:3]

    recent_html = ""

    for row in recent:
        workout_id = int(
            core.safe_float(row["id"])
        )

        recent_html += f"""
<a class="card"
   style="display:block;text-decoration:none;color:inherit;"
   href="/history/{workout_id}">
    <div class="muted">
        {core.value_or_dash(row.get("date"))}
    </div>
    <h3>{core.value_or_dash(row.get("session"))}</h3>
    <p>
        RPE {core.value_or_dash(row.get("rpe"))}
        · {core.value_or_dash(row.get("duration"))} хв
    </p>
</a>
"""

    body = f"""
<div class="card">
    <div class="muted">🔥 TRAINING SYSTEM v4.7</div>
    <h1>Твій тренувальний центр</h1>
    <p>
        Week {CURRENT_WEEK} / 8 · {phase} ·
        Target RIR {target_rir}
    </p>
</div>

<div class="card">
    <div class="muted">🎯 НАСТУПНЕ ТРЕНУВАННЯ</div>
    <h1>{next_session}</h1>
    <p>
        Система визначила наступну сесію
        за фактичною історією.
    </p>

    <a class="button" href="/workout/{next_session}">
        🏋️ ПОЧАТИ {next_session}
    </a>
    <a class="button secondary" href="/workout/{next_session}/guided">
        🧭 Guided Workout
    </a>

    <a class="button secondary" href="/coach">
        🧠 Smart Coach
    </a>
</div>

<div class="grid">
    <div class="card">
        <div class="muted">ТРЕНУВАНЬ</div>
        <div class="stat">{len(workouts)}</div>
    </div>

    <div class="card">
        <div class="muted">PR</div>
        <div class="stat">{core.get_pr_count(workbook)}</div>
    </div>

    <div class="card">
        <div class="muted">AVG RPE</div>
        <div class="stat">
            {f"{avg_rpe:.1f}" if avg_rpe is not None else "—"}
        </div>
    </div>

    <div class="card">
        <div class="muted">AVG RIR</div>
        <div class="stat">
            {f"{avg_rir:.1f}" if avg_rir is not None else "—"}
        </div>
    </div>
</div>

<div class="card">
    <div class="muted">ОСТАННЯ СЕСІЯ</div>
    <h2>{last_session}</h2>
    <p>📅 {last_date}</p>
    <a class="button secondary" href="{last_link}">
        ВІДКРИТИ
    </a>
</div>

<div class="card">
    <h2>📊 ШВИДКИЙ ДОСТУП</h2>

    <a class="button secondary" href="/history">
        🏆 Історія тренувань
    </a>

    <a class="button secondary" href="/exercise">
        📈 Прогрес вправи
    </a>

    <a class="button secondary" href="/exercise/chart">
        📈 Графік прогресу
    </a>

    <a class="button secondary" href="/achievements">
        🏆 PR & Досягнення
    </a>

    <a class="button secondary" href="/performance">
        🧠 Performance
    </a>

    <a class="button secondary" href="/cycle">
        🗓️ 8-тижневий цикл
    </a>
</div>

<div class="card">
    <h2>🕘 ОСТАННІ ТРЕНУВАННЯ</h2>
    {recent_html or '<p>Тренувань поки немає.</p>'}
</div>
"""

    return page(
        "Training System",
        body
    )


@app.route("/next")
def next_workout():
    workbook = get_workbook()
    session = get_next_session(workbook)

    return redirect(
        url_for(
            "workout",
            session=session
        )
    )


@app.route("/dashboard")
def dashboard():
    return redirect(url_for("home"))


@app.route("/cycle")
def cycle():
    workbook = get_workbook()

    info = core.get_week_info(
        workbook,
        CURRENT_WEEK
    )

    stats = core.get_all_cycle_stats(workbook)

    cards = ""

    for week in core.get_available_weeks(workbook):
        week_info = core.get_week_info(workbook, week)
        week_stats = stats.get(week, {"total": 0})

        active = "border:2px solid #111827;" if week == CURRENT_WEEK else ""

        cards += f"""
<div class="card" style="{active}">
    <h3>Week {week}</h3>
    <div>
        {core.value_or_dash(week_info["phase"] if week_info else None)}
        · RIR {core.value_or_dash(week_info["target_rir"] if week_info else None)}
    </div>
    <p class="muted">
        Виконано тренувань: {week_stats["total"]}
    </p>
</div>
"""

    body = f"""
<div class="card">
    <h2>🧠 8-тижневий цикл</h2>
    <p>Поточний: Week {CURRENT_WEEK}</p>
</div>
{cards}
<a class="button secondary" href="/">← Головна</a>
"""

    return page("Cycle", body)



def build_exercise_history(workbook, exercise_name):
    rows = core.get_strength_rows(workbook)
    matching = [
        row for row in rows
        if core.normalize(row["exercise"]) == core.normalize(exercise_name)
        and row["weight"] is not None
        and row["reps"] is not None
    ]

    sessions = {}
    for row in matching:
        workout_id = core.safe_float(row["workout_id"])
        if workout_id is None:
            continue
        sessions.setdefault(workout_id, {
            "date": row.get("date"),
            "session": row.get("session"),
            "sets": []
        })
        sessions[workout_id]["sets"].append(row)

    history = []
    for workout_id, data in sessions.items():
        sets = data["sets"]
        history.append({
            "workout_id": workout_id,
            "date": data["date"],
            "session": data["session"],
            "max_weight": max(row["weight"] for row in sets),
            "total_reps": sum(row["reps"] for row in sets),
            "volume": sum(row["weight"] * row["reps"] for row in sets),
            "avg_rir": core.get_average([
                row["rir"] for row in sets if row["rir"] is not None
            ])
        })

    history.sort(key=lambda x: (str(x["date"] or ""), x["workout_id"]))
    return history



def build_chart_data(history):
    labels = []
    weights = []
    volumes = []

    for item in history:
        labels.append(str(item["date"] or ""))
        weights.append(item["max_weight"])
        volumes.append(item["volume"])

    return labels, weights, volumes


@app.route("/exercise/chart")
def exercise_chart():
    workbook = get_workbook()

    rows = core.get_strength_rows(workbook)

    exercises = core.unique_preserve_order(
        row["exercise"]
        for row in rows
        if row["exercise"]
    )

    selected = request.args.get("exercise", "").strip()

    if not selected and exercises:
        selected = exercises[0]

    history = (
        build_exercise_history(workbook, selected)
        if selected else []
    )

    labels, weights, volumes = build_chart_data(history)

    # Простий SVG без сторонніх бібліотек.
    width = 720
    height = 300
    pad_left = 55
    pad_right = 20
    pad_top = 25
    pad_bottom = 45

    plot_w = width - pad_left - pad_right
    plot_h = height - pad_top - pad_bottom

    if weights:
        min_value = min(weights)
        max_value = max(weights)

        if min_value == max_value:
            min_value -= 1
            max_value += 1

        points = []

        for i, value in enumerate(weights):
            x = (
                pad_left
                if len(weights) == 1
                else pad_left + (
                    i / (len(weights) - 1)
                ) * plot_w
            )

            y = pad_top + (
                (max_value - value)
                / (max_value - min_value)
            ) * plot_h

            points.append(f"{x:.1f},{y:.1f}")

        polyline = " ".join(points)

        circles = "".join(
            f'<circle cx="{p.split(",")[0]}" '
            f'cy="{p.split(",")[1]}" r="5" />'
            for p in points
        )

        svg = f"""
<svg viewBox="0 0 {width} {height}"
     width="100%"
     role="img"
     aria-label="Графік ваги">
    <line x1="{pad_left}" y1="{pad_top}"
          x2="{pad_left}" y2="{height-pad_bottom}"
          stroke="currentColor" />
    <line x1="{pad_left}" y1="{height-pad_bottom}"
          x2="{width-pad_right}" y2="{height-pad_bottom}"
          stroke="currentColor" />

    <polyline
        points="{polyline}"
        fill="none"
        stroke="currentColor"
        stroke-width="3"
    />

    {circles}

    <text x="8" y="{pad_top + 5}"
          font-size="12">{max_value:g} kg</text>

    <text x="8" y="{height-pad_bottom}"
          font-size="12">{min_value:g} kg</text>

    <text x="{pad_left}"
          y="{height-10}"
          font-size="12">
        {labels[0]}
    </text>

    <text x="{width-pad_right-80}"
          y="{height-10}"
          font-size="12">
        {labels[-1]}
    </text>
</svg>
"""
    else:
        svg = """
<div style="padding:30px;text-align:center;">
    Поки недостатньо даних для графіка.
</div>
"""

    options = ""

    for exercise in exercises:
        selected_attr = " selected" if exercise == selected else ""
        options += (
            f'<option value="{exercise}"{selected_attr}>'
            f'{exercise}</option>'
        )

    volume_text = (
        f"{history[-1]['volume']:g}"
        if history else "—"
    )

    body = f"""
<div class="card">
    <h2>📈 ГРАФІК ПРОГРЕСУ</h2>

    <form method="get" action="/exercise/chart">
        <label>Вправа</label>
        <select name="exercise" onchange="this.form.submit()">
            {options}
        </select>
    </form>
</div>

<div class="card">
    <h3>{selected or "Вправа"}</h3>
    <p class="muted">Максимальна вага по кожній сесії</p>
    {svg}
</div>

<div class="grid">
    <div class="card">
        <div class="muted">СЕСІЙ</div>
        <div class="stat">{len(history)}</div>
    </div>

    <div class="card">
        <div class="muted">ОСТАННІЙ ОБСЯГ</div>
        <div class="stat">{volume_text}</div>
    </div>
</div>

<a class="button secondary"
   href="/exercise?exercise={selected}">
    📊 Таблична історія
</a>

<a class="button secondary" href="/">
    ← Головна
</a>
"""

    return page("Графік прогресу", body)



def get_workout_by_id(workbook, workout_id):
    rows = core.get_workout_rows(workbook)
    for row in rows:
        if core.safe_float(row.get("id")) == workout_id:
            return row
    return None


@app.route("/history")
def workout_history():
    workbook = get_workbook()
    rows = core.get_workout_rows(workbook)

    valid = [
        row for row in rows
        if core.safe_float(row.get("id")) is not None
    ]

    valid.sort(
        key=lambda row: (
            str(row.get("date") or ""),
            core.safe_float(row.get("id")) or -1
        ),
        reverse=True
    )

    cards = ""
    for row in valid:
        workout_id = int(core.safe_float(row["id"]))
        cards += f"""
<div class="card">
    <div class="muted">{core.value_or_dash(row.get("date"))} · Workout #{workout_id}</div>
    <h2>{core.value_or_dash(row.get("session"))}</h2>
    <p>RPE: <b>{core.value_or_dash(row.get("rpe"))}</b>
       · {core.value_or_dash(row.get("duration"))} хв</p>
    <p>🏆 PR: {core.value_or_dash(row.get("pr"))}</p>
    <a class="button secondary" href="/history/{workout_id}">
        ВІДКРИТИ ТРЕНУВАННЯ
    </a>
</div>
"""

    if not cards:
        cards = '<div class="card"><h3>🏆 Історія порожня</h3><p>Завершені тренування з’являться тут.</p></div>'

    body = f"""
<div class="card">
    <h1>🏆 ІСТОРІЯ ТРЕНУВАНЬ</h1>
    <p class="muted">Усі завершені сесії.</p>
</div>
{cards}
<a class="button secondary" href="/">← Головна</a>
"""
    return page("Історія тренувань", body)


@app.route("/history/<int:workout_id>")
def workout_history_detail(workout_id):
    workbook = get_workbook()
    workout = get_workout_by_id(workbook, workout_id)

    if workout is None:
        return "Тренування не знайдено", 404

    strength_rows = [
        row for row in core.get_strength_rows(workbook)
        if core.safe_float(row.get("workout_id")) == workout_id
    ]

    exercises = core.unique_preserve_order(
        row["exercise"] for row in strength_rows if row.get("exercise")
    )

    strength_html = ""
    for exercise in exercises:
        sets = [
            row for row in strength_rows
            if core.normalize(row["exercise"]) == core.normalize(exercise)
        ]
        sets.sort(key=lambda row: core.safe_float(row.get("set")) or 0)

        set_html = ""
        for row in sets:
            set_html += f"""
<div class="exercise">
    <b>Set {core.value_or_dash(row.get("set"))}</b>
    <p>{core.value_or_dash(row.get("weight"))} kg ×
       {core.value_or_dash(row.get("reps"))} reps ·
       RIR {core.value_or_dash(row.get("rir"))}</p>
    <p class="muted">{core.value_or_dash(row.get("notes"))}</p>
</div>
"""
        strength_html += f'<div class="card"><h2>{exercise}</h2>{set_html}</div>'

    wod_rows = [
        row for row in core.get_wod_result_rows(workbook)
        if core.safe_float(row.get("workout_id")) == workout_id
    ]

    wod_html = ""
    for row in wod_rows:
        wod_html += f"""
<div class="card">
    <h2>🔥 WOD</h2>
    <p><b>Результат:</b> {core.value_or_dash(row.get("result"))}</p>
    <p><b>RPE:</b> {core.value_or_dash(row.get("rpe"))}</p>
    <p>{core.value_or_dash(row.get("notes"))}</p>
</div>
"""

    body = f"""
<div class="card">
    <div class="muted">Workout #{workout_id}</div>
    <h1>🏋️ {core.value_or_dash(workout.get("session"))}</h1>
    <p>📅 {core.value_or_dash(workout.get("date"))}</p>
    <p>Week {core.value_or_dash(workout.get("week"))}</p>
    <p>⏱️ {core.value_or_dash(workout.get("duration"))} хв
       · RPE {core.value_or_dash(workout.get("rpe"))}</p>
    <p>🏆 PR: {core.value_or_dash(workout.get("pr"))}</p>
    <p>{core.value_or_dash(workout.get("notes"))}</p>
</div>

<h2>💪 СИЛОВА ЧАСТИНА</h2>
{strength_html or '<div class="card"><p>Силові результати не знайдені.</p></div>'}

{wod_html or '<div class="card"><h2>🔥 WOD</h2><p>WOD результат не знайдений.</p></div>'}

<div class="card">
    <h2>📊 Загальний результат</h2>
    <p>{core.value_or_dash(workout.get("result"))}</p>
</div>

<a class="button secondary" href="/history">← Історія</a>
<a class="button secondary" href="/">← Головна</a>
"""
    return page(f"Workout #{workout_id}", body)



# ============================================================
# v3.7 — ATHLETE PERFORMANCE ANALYTICS
# ============================================================

def _v37_session_summaries(workbook):
    rows = core.get_strength_rows(workbook)

    grouped = {}

    for row in rows:
        workout_id = core.safe_float(row.get("workout_id"))
        if workout_id is None:
            continue

        key = (
            workout_id,
            str(row.get("exercise") or "").strip()
        )

        grouped.setdefault(key, []).append(row)

    summaries = []

    for (workout_id, exercise), items in grouped.items():
        valid = [
            row for row in items
            if row.get("weight") is not None
            and row.get("reps") is not None
        ]

        if not valid:
            continue

        summaries.append({
            "workout_id": workout_id,
            "date": items[0].get("date"),
            "session": items[0].get("session"),
            "exercise": exercise,
            "max_weight": max(row["weight"] for row in valid),
            "volume": sum(
                row["weight"] * row["reps"]
                for row in valid
            ),
            "reps": sum(row["reps"] for row in valid),
            "avg_rir": core.get_average([
                row["rir"]
                for row in valid
                if row.get("rir") is not None
            ])
        })

    return summaries


def _v37_compare_exercises(workbook):
    summaries = _v37_session_summaries(workbook)

    by_exercise = {}

    for item in summaries:
        by_exercise.setdefault(
            core.normalize(item["exercise"]),
            []
        ).append(item)

    results = []

    for _, items in by_exercise.items():
        items.sort(
            key=lambda x: (
                str(x["date"] or ""),
                x["workout_id"]
            )
        )

        if not items:
            continue

        exercise = items[-1]["exercise"]

        if len(items) < 2:
            results.append({
                "exercise": exercise,
                "status": "⚪ НОВА",
                "message": "Поки недостатньо історії для порівняння."
            })
            continue

        previous = items[-2]
        current = items[-1]

        if (
            current["volume"] > previous["volume"]
            and current["max_weight"] >= previous["max_weight"]
        ):
            status = "🟢 ПРОГРЕС"
            message = "Остання сесія сильніша за попередню."
        elif (
            current["volume"] < previous["volume"]
            and current["max_weight"] < previous["max_weight"]
        ):
            status = "🔴 РЕГРЕС"
            message = "Остання сесія нижча за попередню."
        else:
            status = "🟡 СТАБІЛЬНО"
            message = "Результат близький до попередньої сесії."

        results.append({
            "exercise": exercise,
            "status": status,
            "message": message
        })

    return results



# ============================================================
# v3.8 — PR & ACHIEVEMENTS
# ============================================================

def _v38_pr_data(workbook):
    rows = core.get_strength_rows(workbook)

    valid = [
        row for row in rows
        if row.get("exercise")
        and row.get("weight") is not None
        and row.get("reps") is not None
    ]

    by_exercise = {}

    for row in valid:
        exercise = str(row["exercise"]).strip()
        key = core.normalize(exercise)
        by_exercise.setdefault(key, {
            "exercise": exercise,
            "rows": []
        })
        by_exercise[key]["rows"].append(row)

    prs = []

    for data in by_exercise.values():
        items = data["rows"]

        max_weight = max(row["weight"] for row in items)
        max_volume_set = max(
            row["weight"] * row["reps"]
            for row in items
        )
        max_reps = max(row["reps"] for row in items)

        weight_row = next(
            row for row in items
            if row["weight"] == max_weight
        )
        volume_row = next(
            row for row in items
            if row["weight"] * row["reps"] == max_volume_set
        )
        reps_row = next(
            row for row in items
            if row["reps"] == max_reps
        )

        prs.append({
            "exercise": data["exercise"],
            "max_weight": max_weight,
            "weight_date": weight_row.get("date"),
            "max_volume": max_volume_set,
            "volume_date": volume_row.get("date"),
            "max_reps": max_reps,
            "reps_date": reps_row.get("date")
        })

    prs.sort(key=lambda x: x["exercise"].lower())
    return prs


def _v38_recent_prs(workbook):
    rows = core.get_strength_rows(workbook)

    valid = [
        row for row in rows
        if row.get("exercise")
        and row.get("weight") is not None
        and row.get("reps") is not None
    ]

    valid.sort(
        key=lambda row: (
            str(row.get("date") or ""),
            core.safe_float(row.get("workout_id")) or -1,
            core.safe_float(row.get("set")) or -1
        )
    )

    best_weight = {}
    best_volume = {}
    achievements = []

    for row in valid:
        key = core.normalize(row["exercise"])
        weight = row["weight"]
        volume = row["weight"] * row["reps"]

        if key not in best_weight:
            best_weight[key] = weight
            best_volume[key] = volume
            continue

        if weight > best_weight[key]:
            achievements.append({
                "date": row.get("date"),
                "exercise": row["exercise"],
                "type": "🏆 НОВА МАКСИМАЛЬНА ВАГА",
                "value": f"{weight:g} kg"
            })
            best_weight[key] = weight

        if volume > best_volume[key]:
            achievements.append({
                "date": row.get("date"),
                "exercise": row["exercise"],
                "type": "📈 НОВИЙ VOLUME PR",
                "value": f"{volume:g}"
            })
            best_volume[key] = volume

    return achievements[-12:][::-1]


@app.route("/achievements")
def achievements():
    workbook = get_workbook()

    prs = _v38_pr_data(workbook)
    recent = _v38_recent_prs(workbook)

    pr_cards = ""

    for item in prs:
        pr_cards += f"""
<div class="card">
    <h3>{item["exercise"]}</h3>
    <p>🏋️ Максимальна вага:
       <b>{item["max_weight"]:g} kg</b>
       <span class="muted">({core.value_or_dash(item["weight_date"])})</span>
    </p>
    <p>📈 Найбільший set-volume:
       <b>{item["max_volume"]:g}</b>
       <span class="muted">({core.value_or_dash(item["volume_date"])})</span>
    </p>
    <p>🔁 Найбільше повторень:
       <b>{item["max_reps"]:g}</b>
       <span class="muted">({core.value_or_dash(item["reps_date"])})</span>
    </p>
</div>
"""

    recent_cards = ""

    for item in recent:
        recent_cards += f"""
<div class="card">
    <div class="muted">{core.value_or_dash(item["date"])}</div>
    <h3>{item["type"]}</h3>
    <p><b>{item["exercise"]}</b> · {item["value"]}</p>
</div>
"""

    body = f"""
<div class="card">
    <h1>🏆 PR & ДОСЯГНЕННЯ</h1>
    <p class="muted">
        Особисті максимуми, визначені з фактичних результатів.
    </p>
</div>

<div class="card">
    <h2>🏆 ОСОБИСТІ РЕКОРДИ</h2>
    <p>Вага · set-volume · повторення</p>
</div>

{pr_cards or """
<div class="card">
    <p>Поки немає достатньо фактичних силових результатів.</p>
</div>
"""}

<div class="card">
    <h2>🔥 ОСТАННІ ДОСЯГНЕННЯ</h2>
</div>

{recent_cards or """
<div class="card">
    <p>Нові досягнення зявляться після перевищення попередніх максимумів.</p>
</div>
"""}

<a class="button secondary" href="/performance">
    🧠 Performance
</a>

<a class="button secondary" href="/">
    ← Головна
</a>
"""

    return page("PR & Досягнення", body)



# ============================================================
# v3.9 — SMART COACH
# ============================================================

@app.route("/coach")
def smart_coach():
    workbook = get_workbook()

    next_session = get_next_session(workbook)

    exercises = core.get_strength_plan(
        workbook,
        CURRENT_WEEK,
        next_session
    )

    recommendation_cards = ""

    for exercise in exercises:
        try:
            label, message = core.calculate_next_session_recommendation(
                workbook,
                exercise,
                next_session,
                CURRENT_WEEK
            )
        except Exception:
            label = "ПОЧАТИ ЗА ПЛАНОМ"
            message = "Персональна рекомендація поки недоступна."

        recommendation_cards += f"""
<div class="card">
    <h2>{exercise["exercise"]}</h2>
    <p>
        <b>План:</b> {exercise["sets"]} × {exercise["reps"]}
    </p>
    <p>
        🎯 <b>{label}</b>
    </p>
    <p class="muted">{message}</p>
</div>
"""

    if not recommendation_cards:
        recommendation_cards = """
<div class="card">
    <h3>Силову частину не знайдено</h3>
    <p>Перевір план поточного циклу.</p>
</div>
"""

    body = f"""
<div class="card">
    <div class="muted">🧠 SMART COACH</div>
    <h1>🎯 {next_session}</h1>
    <p>
        Рекомендації сформовані з фактичної історії,
        поточної фази циклу та записаних результатів.
    </p>
</div>

{recommendation_cards}

<div class="card">
    <h2>⚠️ Важливий принцип</h2>
    <p>
        Система не змінює програму автоматично.
        Вона дає рекомендацію, а рішення про виконання
        залишається за тобою.
    </p>
</div>

<a class="button" href="/workout/{next_session}">
    🏋️ ПОЧАТИ {next_session}
</a>

<a class="button secondary" href="/performance">
    🧠 Performance
</a>

<a class="button secondary" href="/">
    ← Головна
</a>
"""

    return page("Smart Coach", body)



# ============================================================
# v4.7 — POST WORKOUT ANALYSIS
# ============================================================

def _v47_num(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _v47_workout_analysis(workbook, workout_id):
    strength_rows = [
        row for row in core.get_strength_rows(workbook)
        if core.safe_float(row.get("workout_id")) == workout_id
    ]

    grouped = {}

    for row in strength_rows:
        exercise = str(row.get("exercise") or "").strip()
        if not exercise:
            continue

        key = core.normalize(exercise)
        item = grouped.setdefault(key, {
            "exercise": exercise,
            "sets": 0,
            "volume": 0,
            "best_weight": None,
            "best_reps": None,
            "rir_values": []
        })

        item["sets"] += 1

        weight = _v47_num(row.get("weight"))
        reps = _v47_num(row.get("reps"))
        rir = _v47_num(row.get("rir"))

        if weight is not None and reps is not None:
            item["volume"] += weight * reps
            item["best_weight"] = (
                weight if item["best_weight"] is None
                else max(item["best_weight"], weight)
            )
            item["best_reps"] = (
                reps if item["best_reps"] is None
                else max(item["best_reps"], reps)
            )

        if rir is not None:
            item["rir_values"].append(rir)

    exercise_rows = list(grouped.values())
    all_rir = [
        rir for item in exercise_rows
        for rir in item["rir_values"]
    ]

    avg_rir = sum(all_rir) / len(all_rir) if all_rir else None

    best_set = None
    for row in strength_rows:
        weight = _v47_num(row.get("weight"))
        reps = _v47_num(row.get("reps"))
        if weight is None or reps is None:
            continue

        volume = weight * reps
        if best_set is None or volume > best_set["volume"]:
            best_set = {
                "exercise": row.get("exercise"),
                "weight": weight,
                "reps": reps,
                "volume": volume
            }

    return {
        "sets": len(strength_rows),
        "exercises": len(exercise_rows),
        "total_volume": sum(x["volume"] for x in exercise_rows),
        "avg_rir": avg_rir,
        "best_set": best_set,
        "exercise_rows": exercise_rows
    }


@app.route("/workout/complete/<int:workout_id>")
def workout_complete(workout_id):
    workbook = get_workbook()
    analysis = _v47_workout_analysis(workbook, workout_id)

    if analysis["sets"] == 0:
        status = "⚪ НЕМАЄ ДАНИХ"
        summary = "Для цього тренування не знайдено збережених SET."
    elif analysis["avg_rir"] is not None and analysis["avg_rir"] <= 1:
        status = "🔴 ВИСОКА ІНТЕНСИВНІСТЬ"
        summary = "Середній RIR був 1 або нижче."
    elif analysis["avg_rir"] is not None and analysis["avg_rir"] <= 2:
        status = "🟡 СИЛЬНА РОБОТА"
        summary = "Середній RIR був близько робочої межі."
    else:
        status = "🟢 КОНТРОЛЬОВАНА РОБОТА"
        summary = "Робота виконана із запасом."

    best = analysis["best_set"]
    best_html = ""

    if best:
        best_html = f"""
<div class="card">
    <div class="muted">🏆 НАЙКРАЩИЙ SET ЗА VOLUME</div>
    <h2>{best["exercise"]}</h2>
    <p>{best["weight"]:g} kg × {best["reps"]:g}</p>
    <p>Set-volume: <b>{best["volume"]:g}</b></p>
</div>
"""

    exercise_html = ""
    for item in analysis["exercise_rows"]:
        weight_text = (
            f'{item["best_weight"]:g} kg'
            if item["best_weight"] is not None else "—"
        )
        reps_text = (
            f'{item["best_reps"]:g}'
            if item["best_reps"] is not None else "—"
        )

        exercise_html += f"""
<div class="card">
    <h3>{item["exercise"]}</h3>
    <p>{item["sets"]} сетів · Volume {item["volume"]:g}</p>
    <p class="muted">
        Best weight: {weight_text}
        · Best reps: {reps_text}
    </p>
</div>
"""

    avg_rir_text = (
        f'{analysis["avg_rir"]:.1f}'
        if analysis["avg_rir"] is not None else "—"
    )

    body = f"""
<div class="card">
    <div class="muted">✅ ТРЕНУВАННЯ ЗАВЕРШЕНО</div>
    <h1>Workout #{workout_id}</h1>
    <h2>{status}</h2>
    <p>{summary}</p>
</div>

<div class="grid">
    <div class="card">
        <div class="muted">SETS</div>
        <div class="stat">{analysis["sets"]}</div>
    </div>
    <div class="card">
        <div class="muted">ВПРАВИ</div>
        <div class="stat">{analysis["exercises"]}</div>
    </div>
    <div class="card">
        <div class="muted">VOLUME</div>
        <div class="stat">{analysis["total_volume"]:g}</div>
    </div>
    <div class="card">
        <div class="muted">AVG RIR</div>
        <div class="stat">{avg_rir_text}</div>
    </div>
</div>

{best_html}

<div class="card">
    <h2>📊 ПО ВПРАВАХ</h2>
</div>

{exercise_html or "<div class='card'><p>Деталей немає.</p></div>"}

<a class="button secondary" href="/history">🏆 Історія</a>
<a class="button secondary" href="/performance">🧠 Performance</a>
<a class="button secondary" href="/">← Головна</a>
"""

    return page("Тренування завершено", body)

@app.route("/performance")
def athlete_performance():
    workbook = get_workbook()

    workouts = core.get_workout_rows(workbook)
    strength = core.get_strength_rows(workbook)

    valid_workouts = [
        row for row in workouts
        if core.safe_float(row.get("id")) is not None
    ]

    avg_rpe = core.get_average([
        row["rpe"]
        for row in valid_workouts
        if row.get("rpe") is not None
    ])

    total_volume = sum(
        row["weight"] * row["reps"]
        for row in strength
        if row.get("weight") is not None
        and row.get("reps") is not None
    )

    comparisons = _v37_compare_exercises(workbook)

    progress_count = sum(
        1 for item in comparisons
        if item["status"] == "🟢 ПРОГРЕС"
    )

    stable_count = sum(
        1 for item in comparisons
        if item["status"] == "🟡 СТАБІЛЬНО"
    )

    regression_count = sum(
        1 for item in comparisons
        if item["status"] == "🔴 РЕГРЕС"
    )

    cards = ""

    for item in comparisons:
        cards += f"""
<div class="card">
    <h3>{item["exercise"]}</h3>
    <p><b>{item["status"]}</b></p>
    <p>{item["message"]}</p>
</div>
"""

    if not cards:
        cards = """
<div class="card">
    <h3>Поки недостатньо даних</h3>
    <p>Потрібно завершити хоча б одне тренування.</p>
</div>
"""

    body = f"""
<div class="card">
    <h1>🧠 PERFORMANCE</h1>
    <p class="muted">
        Загальна оцінка фактичних тренувальних результатів.
    </p>
</div>

<div class="grid">
    <div class="card">
        <div class="muted">ТРЕНУВАНЬ</div>
        <div class="stat">{len(valid_workouts)}</div>
    </div>

    <div class="card">
        <div class="muted">СЕРЕДНІЙ RPE</div>
        <div class="stat">
            {f"{avg_rpe:.1f}" if avg_rpe is not None else "—"}
        </div>
    </div>

    <div class="card">
        <div class="muted">ЗАГАЛЬНИЙ VOLUME</div>
        <div class="stat">{total_volume:g}</div>
    </div>

    <div class="card">
        <div class="muted">PR</div>
        <div class="stat">{core.get_pr_count(workbook)}</div>
    </div>
</div>

<div class="card">
    <h2>📈 СТАН ВПРАВ</h2>
    <p>🟢 Прогрес: <b>{progress_count}</b></p>
    <p>🟡 Стабільно: <b>{stable_count}</b></p>
    <p>🔴 Регрес: <b>{regression_count}</b></p>
</div>

{cards}

<a class="button secondary" href="/dashboard">
    📊 Dashboard
</a>

<a class="button secondary" href="/">
    ← Головна
</a>
"""

    return page("Performance", body)


@app.route("/exercise")
def exercise_history():
    workbook = get_workbook()
    rows = core.get_strength_rows(workbook)

    exercises = core.unique_preserve_order(
        row["exercise"] for row in rows if row["exercise"]
    )

    selected = request.args.get("exercise", "").strip()
    if not selected and exercises:
        selected = exercises[0]

    history = build_exercise_history(workbook, selected) if selected else []

    latest = history[-1] if history else None
    best_weight = max((x["max_weight"] for x in history), default=None)
    best_volume = max((x["volume"] for x in history), default=None)

    options = ""
    for exercise in exercises:
        selected_attr = " selected" if exercise == selected else ""
        options += (
            f'<option value="{exercise}"{selected_attr}>{exercise}</option>'
        )

    rows_html = ""
    for item in reversed(history):
        rir = f'{item["avg_rir"]:.1f}' if item["avg_rir"] is not None else "—"
        rows_html += (
            "<tr>"
            f'<td style="padding:8px">{core.value_or_dash(item["date"])}</td>'
            f'<td style="padding:8px">{core.value_or_dash(item["session"])}</td>'
            f'<td style="padding:8px">{item["max_weight"]:g}</td>'
            f'<td style="padding:8px">{item["total_reps"]:g}</td>'
            f'<td style="padding:8px">{item["volume"]:g}</td>'
            f'<td style="padding:8px">{rir}</td>'
            "</tr>"
        )

    latest_weight = f'{latest["max_weight"]:g} kg' if latest else "—"
    best_weight_text = f'{best_weight:g} kg' if best_weight is not None else "—"
    best_volume_text = f'{best_volume:g}' if best_volume is not None else "—"

    body = f"""
<div class="card">
    <h2>📈 ПРОГРЕС ВПРАВИ</h2>
    <form method="get" action="/exercise">
        <label>Вправа</label>
        <select name="exercise" onchange="this.form.submit()">
            {options}
        </select>
    </form>
</div>

<div class="grid">
    <div class="card">
        <div class="muted">ОСТАННЯ ВАГА</div>
        <div class="stat">{latest_weight}</div>
    </div>
    <div class="card">
        <div class="muted">НАЙКРАЩА ВАГА</div>
        <div class="stat">{best_weight_text}</div>
    </div>
    <div class="card">
        <div class="muted">НАЙБІЛЬШИЙ ОБСЯГ</div>
        <div class="stat">{best_volume_text}</div>
    </div>
    <div class="card">
        <div class="muted">СЕСІЙ</div>
        <div class="stat">{len(history)}</div>
    </div>
</div>

<div class="card">
    <h3>📊 Історія</h3>
    <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
            <tr>
                <th style="padding:8px;text-align:left;">Дата</th>
                <th style="padding:8px;text-align:left;">Тип</th>
                <th style="padding:8px;text-align:left;">kg</th>
                <th style="padding:8px;text-align:left;">Reps</th>
                <th style="padding:8px;text-align:left;">Volume</th>
                <th style="padding:8px;text-align:left;">RIR</th>
            </tr>
            {rows_html if rows_html else '<tr><td colspan="6" style="padding:12px;">Даних поки немає.</td></tr>'}
        </table>
    </div>
</div>

<a class="button secondary" href="/progress">📈 Оцінка прогресу</a>
<a class="button secondary" href="/">← Головна</a>
"""

    return page("Прогрес вправи", body)


@app.route("/progress")
def progress():
    workbook = get_workbook()

    rows = core.get_strength_rows(workbook)

    exercises = core.unique_preserve_order(
        row["exercise"] for row in rows
    )

    result_html = ""

    for exercise in exercises:
        sessions = core.unique_preserve_order(
            row["session"]
            for row in rows
            if core.normalize(row["exercise"]) ==
               core.normalize(exercise)
        )

        for session in sessions:
            result = core.evaluate_exercise_progress(
                workbook,
                exercise,
                session
            )

            result_html += f"""
<div class="card">
    <h3>{exercise}</h3>
    <div><b>{result["label"]}</b></div>
    <p>{result["message"]}</p>
</div>
"""

    if not result_html:
        result_html = """
<div class="card">
    <h3>Поки недостатньо даних</h3>
    <p>Потрібно більше фактичних результатів тренувань.</p>
</div>
"""

    body = f"""
<div class="card">
    <h2>📈 Оцінка прогресу</h2>
    <p>Порівняння останніх фактичних сесій.</p>
</div>
{result_html}
<a class="button secondary" href="/">← Головна</a>
"""

    return page("Progress", body)



def web_save_workout(
    workbook,
    current_week,
    session,
    strength_entries,
    wod_result,
    wod_rpe,
    wod_notes,
    duration,
    overall_result,
    overall_rpe,
    pr,
    notes
):
    """Зберігає одне повне тренування з веб-форми."""
    core.ensure_result_sheets(workbook)

    workout_id = core.next_workout_id(workbook)
    date_value = date.today().isoformat()

    for entry in strength_entries:
        core.append_strength_result(
            workbook,
            date_value,
            workout_id,
            current_week,
            session,
            entry["exercise"],
            entry["set_number"],
            entry["weight"],
            entry["reps"],
            entry["rir"],
            entry["notes"]
        )

    wods = core.get_wods(
        workbook,
        current_week,
        session
    )

    wod_id = wods[0].get("id") if wods else None
    wod_format = wods[0].get("format") if wods else None

    if wods:
        core.append_wod_result(
            workbook,
            date_value,
            workout_id,
            current_week,
            session,
            wod_id,
            wod_result,
            wod_rpe,
            wod_notes
        )

    core.append_workout_record(
        workbook,
        workout_id,
        date_value,
        current_week,
        session,
        wod_id,
        wod_format,
        duration,
        overall_result,
        overall_rpe if overall_rpe is not None else wod_rpe,
        pr,
        notes
    )

    workbook.save(core.get_file_path())

    return workout_id


@app.route("/workout/<session>/save", methods=["POST"])
def save_web_workout(session):
    session = session.upper()

    if session not in SESSION_TYPES:
        return "Невірний тип тренування", 404

    workbook = get_workbook()

    exercises = core.get_strength_plan(
        workbook,
        CURRENT_WEEK,
        session
    )

    strength_entries = []

    for index, exercise in enumerate(exercises):
        try:
            sets_count = int(float(exercise["sets"]))
        except (TypeError, ValueError):
            sets_count = 1

        for set_number in range(1, sets_count + 1):
            prefix = f"ex_{index}_set_{set_number}"

            weight_raw = request.form.get(
                prefix + "_weight",
                ""
            ).strip().replace(",", ".")

            reps_raw = request.form.get(
                prefix + "_reps",
                ""
            ).strip()

            rir_raw = request.form.get(
                prefix + "_rir",
                ""
            ).strip().replace(",", ".")

            set_notes = request.form.get(
                prefix + "_notes",
                ""
            ).strip()

            # Порожній сет не записуємо.
            if weight_raw == "" and reps_raw == "":
                continue

            try:
                weight = float(weight_raw)
                reps = int(reps_raw)
            except ValueError:
                return (
                    f"Помилка у вправі {exercise['exercise']}, "
                    f"set {set_number}: перевір вагу та повторення.",
                    400
                )

            rir = None

            if rir_raw != "":
                try:
                    rir = float(rir_raw)
                except ValueError:
                    return (
                        f"Помилка RIR у вправі "
                        f"{exercise['exercise']}, set {set_number}.",
                        400
                    )

            strength_entries.append({
                "exercise": exercise["exercise"],
                "set_number": set_number,
                "weight": weight,
                "reps": reps,
                "rir": rir,
                "notes": set_notes
            })

    wod_result = request.form.get("wod_result", "").strip()
    wod_notes = request.form.get("wod_notes", "").strip()

    wod_rpe_raw = request.form.get(
        "wod_rpe",
        ""
    ).strip().replace(",", ".")

    duration_raw = request.form.get(
        "duration",
        ""
    ).strip().replace(",", ".")

    overall_rpe_raw = request.form.get(
        "overall_rpe",
        ""
    ).strip().replace(",", ".")

    overall_result = request.form.get(
        "overall_result",
        ""
    ).strip()

    pr = request.form.get(
        "pr",
        "ні"
    ).strip()

    notes = request.form.get(
        "notes",
        ""
    ).strip()

    try:
        wod_rpe = float(wod_rpe_raw) if wod_rpe_raw else None
        duration = float(duration_raw) if duration_raw else None
        overall_rpe = (
            float(overall_rpe_raw)
            if overall_rpe_raw
            else None
        )
    except ValueError:
        return "Перевір числові поля RPE та тривалості.", 400

    workout_id = web_save_workout(
        workbook,
        CURRENT_WEEK,
        session,
        strength_entries,
        wod_result,
        wod_rpe,
        wod_notes,
        duration,
        overall_result,
        overall_rpe,
        pr,
        notes
    )

    return redirect(
        url_for(
            "workout_complete",
            workout_id=workout_id
        )
    )


def get_previous_exercise_summary(
    workbook,
    exercise_name,
    session,
    current_workout_id
):
    rows = core.get_strength_rows(workbook)

    matching = [
        row for row in rows
        if core.normalize(row["exercise"]) ==
           core.normalize(exercise_name)
        and core.normalize(row["session"]) ==
           core.normalize(session)
        and core.safe_float(row["workout_id"]) is not None
        and core.safe_float(row["workout_id"]) <
            current_workout_id
        and row["weight"] is not None
        and row["reps"] is not None
    ]

    if not matching:
        return None

    previous_ids = sorted({
        core.safe_float(row["workout_id"])
        for row in matching
    })

    if not previous_ids:
        return None

    previous_id = previous_ids[-1]

    previous = [
        row for row in matching
        if core.safe_float(row["workout_id"]) == previous_id
    ]

    return {
        "workout_id": previous_id,
        "max_weight": max(row["weight"] for row in previous),
        "total_reps": sum(row["reps"] for row in previous),
        "volume": sum(
            row["weight"] * row["reps"]
            for row in previous
        )
    }


def get_current_exercise_summary(
    workbook,
    exercise_name,
    session,
    current_workout_id
):
    rows = core.get_strength_rows(workbook)

    current = [
        row for row in rows
        if core.normalize(row["exercise"]) ==
           core.normalize(exercise_name)
        and core.normalize(row["session"]) ==
           core.normalize(session)
        and core.safe_float(row["workout_id"]) ==
            current_workout_id
        and row["weight"] is not None
        and row["reps"] is not None
    ]

    if not current:
        return None

    return {
        "max_weight": max(row["weight"] for row in current),
        "total_reps": sum(row["reps"] for row in current),
        "volume": sum(
            row["weight"] * row["reps"]
            for row in current
        )
    }


@app.route("/workout/complete/<int:workout_id>", endpoint="workout_complete_v2")
def workout_complete_v2(workout_id):
    workbook = get_workbook()

    workout_rows = core.get_workout_rows(workbook)

    current_workout = next(
        (
            row for row in workout_rows
            if core.safe_float(row.get("id")) == workout_id
        ),
        None
    )

    if current_workout is None:
        return "Тренування не знайдено", 404

    session = str(
        current_workout.get("session") or ""
    ).strip().upper()

    strength_rows = [
        row for row in core.get_strength_rows(workbook)
        if core.safe_float(row["workout_id"]) == workout_id
    ]

    exercises = core.unique_preserve_order(
        row["exercise"]
        for row in strength_rows
    )

    progress_cards = ""
    progress_count = 0
    stable_count = 0
    regression_count = 0

    for exercise in exercises:
        current = get_current_exercise_summary(
            workbook,
            exercise,
            session,
            workout_id
        )

        previous = get_previous_exercise_summary(
            workbook,
            exercise,
            session,
            workout_id
        )

        if not current:
            continue

        if not previous:
            label = "⚪ НОВА / НЕДОСТАТНЬО ІСТОРІЇ"
            message = "Для порівняння потрібна попередня сесія."
        elif (
            current["volume"] > previous["volume"]
            and current["max_weight"] >= previous["max_weight"]
        ):
            label = "🟢 ПРОГРЕС"
            message = "Тренувальний обсяг або навантаження зросли."
            progress_count += 1
        elif (
            current["volume"] < previous["volume"]
            and current["max_weight"] < previous["max_weight"]
        ):
            label = "🔴 РЕГРЕС"
            message = "Останній результат нижчий за попередній."
            regression_count += 1
        else:
            label = "🟡 СТАБІЛЬНО"
            message = "Ключові показники без явного прогресу."
            stable_count += 1

        if previous:
            comparison = f"""
<p class="muted">
Попередня: {previous["max_weight"]:g} kg ·
{previous["total_reps"]:g} reps ·
Volume {previous["volume"]:g}
</p>
<p>
Сьогодні: <b>{current["max_weight"]:g} kg</b> ·
<b>{current["total_reps"]:g} reps</b> ·
Volume <b>{current["volume"]:g}</b>
</p>
"""
        else:
            comparison = f"""
<p>
Сьогодні: <b>{current["max_weight"]:g} kg</b> ·
<b>{current["total_reps"]:g} reps</b> ·
Volume <b>{current["volume"]:g}</b>
</p>
"""

        progress_cards += f"""
<div class="card">
    <h3>{exercise}</h3>
    <p><b>{label}</b></p>
    <p>{message}</p>
    {comparison}
</div>
"""

    duration = current_workout.get("duration")
    overall_rpe = current_workout.get("rpe")
    wod_result = current_workout.get("result")

    next_session = get_next_session(workbook)

    body = f"""
<div class="card">
    <div class="muted">Тренування #{workout_id}</div>
    <h1>✅ {session}</h1>
    <p>Week {current_workout.get("week", "—")}</p>
</div>

<div class="grid">
    <div class="card">
        <div class="muted">ТРИВАЛІСТЬ</div>
        <div class="stat">
            {core.value_or_dash(duration)} хв
        </div>
    </div>
    <div class="card">
        <div class="muted">RPE</div>
        <div class="stat">
            {core.value_or_dash(overall_rpe)}
        </div>
    </div>
</div>

<div class="card">
    <h2>📈 ЩО ЗМІНИЛОСЯ</h2>
    <p>🟢 Прогрес: <b>{progress_count}</b></p>
    <p>🟡 Стабільно: <b>{stable_count}</b></p>
    <p>🔴 Регрес: <b>{regression_count}</b></p>
</div>

{progress_cards or """
<div class="card">
    <h3>💪 Силові результати</h3>
    <p>Дані цього тренування не знайдені.</p>
</div>
"""}

<div class="card">
    <h2>🔥 WOD</h2>
    <p>{core.value_or_dash(wod_result)}</p>
</div>

<div class="card">
    <div class="muted">🎯 НАСТУПНА СЕСІЯ</div>
    <h2>{next_session}</h2>
    <a class="button" href="/workout/{next_session}">
        🏋️ ПОЧАТИ {next_session}
    </a>
</div>

<a class="button secondary" href="/dashboard">
    📊 Dashboard
</a>

<a class="button secondary" href="/">
    ← Головна
</a>
"""

    return page(
        "Тренування завершено",
        body
    )



# ============================================================
# v4.2 — GUIDED WORKOUT MODE
# ============================================================



# ============================================================
# v4.4 — GUIDED SESSION SAVE
# ============================================================

@app.route("/api/guided/save", methods=["POST"])
def guided_save():
    workbook = get_workbook()

    payload = request.get_json(silent=True) or {}

    session = str(payload.get("session") or "").strip().upper()
    week = payload.get("week")
    workout_id = payload.get("workout_id")
    exercise = str(payload.get("exercise") or "").strip()
    set_number = payload.get("set")
    weight = payload.get("weight")
    reps = payload.get("reps")
    rir = payload.get("rir")
    notes = str(payload.get("notes") or "").strip()

    if not session or not exercise:
        return jsonify({
            "ok": False,
            "error": "Не вистачає session або exercise."
        }), 400

    if weight in ("", None) or reps in ("", None):
        return jsonify({
            "ok": False,
            "error": "Введи вагу та повторення."
        }), 400

    try:
        weight = float(weight)
        reps = float(reps)
    except (TypeError, ValueError):
        return jsonify({
            "ok": False,
            "error": "Вага та повторення мають бути числами."
        }), 400

    try:
        rir = float(rir) if rir not in ("", None) else None
    except (TypeError, ValueError):
        rir = None

    try:
        set_number = int(set_number)
    except (TypeError, ValueError):
        set_number = 1

    try:
        week = int(week)
    except (TypeError, ValueError):
        week = CURRENT_WEEK

    try:
        workout_id = int(float(workout_id))
    except (TypeError, ValueError):
        workout_id = None

    if workout_id is None:
        # У поточному core v2.9 немає get_or_create_workout_id().
        # Для Guided Mode беремо наступний ID з WORKOUTS.
        workout_ids = []

        if "WORKOUTS" in workbook.sheetnames:
            ws_workouts = workbook["WORKOUTS"]

            for row in range(2, ws_workouts.max_row + 1):
                value = ws_workouts.cell(row, 1).value
                try:
                    workout_ids.append(int(float(value)))
                except (TypeError, ValueError):
                    pass

        workout_id = (
            max(workout_ids) + 1
            if workout_ids
            else 1
        )

    if "STRENGTH_RESULTS" not in workbook.sheetnames:
        core.ensure_result_sheets(workbook)

    sheet = workbook["STRENGTH_RESULTS"]

    # Уникаємо дублювання того самого сету при повторному натисканні SAVE.
    existing_row = None

    for row in range(2, sheet.max_row + 1):
        same_workout = (
            sheet.cell(row, 3).value == workout_id
        )
        same_exercise = (
            str(sheet.cell(row, 6).value or "").strip().lower()
            == exercise.lower()
        )
        same_set = (
            sheet.cell(row, 7).value == set_number
        )

        if same_workout and same_exercise and same_set:
            existing_row = row
            break

    if existing_row is None:
        existing_row = sheet.max_row + 1

        # ID
        sheet.cell(existing_row, 1).value = existing_row - 1

        # Date
        sheet.cell(existing_row, 2).value = date.today().strftime(
            "%Y-%m-%d"
        )

    sheet.cell(existing_row, 3).value = workout_id
    sheet.cell(existing_row, 4).value = week
    sheet.cell(existing_row, 5).value = session
    sheet.cell(existing_row, 6).value = exercise
    sheet.cell(existing_row, 7).value = set_number
    sheet.cell(existing_row, 8).value = weight
    sheet.cell(existing_row, 9).value = reps
    sheet.cell(existing_row, 10).value = rir
    sheet.cell(existing_row, 11).value = notes

    workbook.save(core.get_file_path())

    return jsonify({
        "ok": True,
        "message": "SET записано в Excel.",
        "row": existing_row,
        "workout_id": workout_id
    })



# ============================================================
# v4.5 — GUIDED WORKOUT FINALIZATION
# ============================================================

@app.route("/api/guided/finish", methods=["POST"])
def guided_finish():
    workbook = get_workbook()
    payload = request.get_json(silent=True) or {}

    session = str(payload.get("session") or "").strip().upper()

    if session not in SESSION_TYPES:
        return jsonify({
            "ok": False,
            "error": "Невірний тип тренування."
        }), 400

    try:
        workout_id = int(float(payload.get("workout_id")))
    except (TypeError, ValueError):
        return jsonify({
            "ok": False,
            "error": "Не знайдено Workout ID."
        }), 400

    strength_rows = [
        row for row in core.get_strength_rows(workbook)
        if core.safe_float(row.get("workout_id")) == workout_id
    ]

    if not strength_rows:
        return jsonify({
            "ok": False,
            "error": "У цьому тренуванні ще немає збережених SET."
        }), 400

    try:
        week = int(payload.get("week", CURRENT_WEEK))
    except (TypeError, ValueError):
        week = CURRENT_WEEK

    wod_result = str(payload.get("wod_result") or "").strip()
    wod_notes = str(payload.get("wod_notes") or "").strip()
    overall_result = str(
        payload.get("overall_result") or ""
    ).strip()
    pr = str(payload.get("pr") or "ні").strip()
    notes = str(payload.get("notes") or "").strip()

    def number_or_none(value):
        if value in ("", None):
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    wod_rpe = number_or_none(payload.get("wod_rpe"))
    duration = number_or_none(payload.get("duration"))
    overall_rpe = number_or_none(payload.get("overall_rpe"))

    core.ensure_result_sheets(workbook)

    date_value = date.today().isoformat()

    # Створюємо заголовок тренування, якщо його ще немає.
    workout_rows = core.get_workout_rows(workbook)

    existing_workout = next(
        (
            row for row in workout_rows
            if core.safe_float(row.get("id")) == workout_id
        ),
        None
    )

    wods = core.get_wods(
        workbook,
        week,
        session
    )

    wod_id = wods[0].get("id") if wods else None
    wod_format = wods[0].get("format") if wods else None

    if existing_workout is None:
        core.append_workout_record(
            workbook,
            workout_id,
            date_value,
            week,
            session,
            wod_id,
            wod_format,
            duration,
            overall_result,
            overall_rpe if overall_rpe is not None else wod_rpe,
            pr,
            notes
        )
    else:
        # Якщо заголовок уже існує — оновлюємо підсумкові поля.
        sheet = workbook["WORKOUTS"]
        target_row = None

        for row in range(2, sheet.max_row + 1):
            if core.safe_float(sheet.cell(row, 1).value) == workout_id:
                target_row = row
                break

        if target_row is not None:
            sheet.cell(target_row, 2).value = date_value
            sheet.cell(target_row, 3).value = week
            sheet.cell(target_row, 4).value = session
            sheet.cell(target_row, 5).value = wod_id
            sheet.cell(target_row, 6).value = wod_format
            sheet.cell(target_row, 7).value = (
                wods[0].get("format") if wods else None
            )
            sheet.cell(target_row, 8).value = duration
            sheet.cell(target_row, 9).value = overall_result
            sheet.cell(target_row, 10).value = (
                overall_rpe if overall_rpe is not None else wod_rpe
            )
            sheet.cell(target_row, 11).value = pr
            sheet.cell(target_row, 12).value = notes

    # Додаємо WOD result лише один раз.
    if wods:
        wod_result_rows = core.get_wod_result_rows(workbook)

        already_saved = any(
            core.safe_float(row.get("workout_id")) == workout_id
            for row in wod_result_rows
        )

        if not already_saved:
            core.append_wod_result(
                workbook,
                date_value,
                workout_id,
                week,
                session,
                wod_id,
                wod_result,
                wod_rpe,
                wod_notes
            )

    # Один фінальний SAVE після всіх змін.
    workbook.save(core.get_file_path())

    return jsonify({
        "ok": True,
        "workout_id": workout_id,
        "message": "Тренування завершено та записано."
    })


@app.route("/workout/<session>/guided")
def guided_workout(session):
    workbook = get_workbook()
    session = str(session).strip().upper()

    if session not in SESSION_TYPES:
        return "Невірний тип тренування", 404

    strength_plan = core.get_strength_plan(
        workbook,
        CURRENT_WEEK,
        session
    )
    wods = core.get_wods(
        workbook,
        CURRENT_WEEK,
        session
    )

    if not strength_plan:
        return page("Тренування", f"""
<div class="card">
    <h1>🏋️ {session}</h1>
    <p>Силову частину для цього тренування не знайдено.</p>
</div>
<a class="button secondary" href="/">← Головна</a>
""")

    try:
        index = int(request.args.get("exercise", "0"))
    except ValueError:
        index = 0

    index = max(0, min(index, len(strength_plan) - 1))
    exercise = strength_plan[index]

    try:
        sets_count = int(float(exercise.get("sets", 1)))
    except (TypeError, ValueError):
        sets_count = 1

    try:
        recommendation = core.calculate_next_session_recommendation(
            workbook, exercise, session, CURRENT_WEEK
        )
    except Exception:
        recommendation = (
            "ЗА ПЛАНОМ",
            "Виконай заплановані сети та зафіксуй фактичний RIR."
        )

    progress = int(((index + 1) / len(strength_plan)) * 100)
    sets_html = ""

    for set_number in range(1, sets_count + 1):
        key = f"guided_{session}_{index}_{set_number}"
        sets_html += f"""
<div class="card" style="margin-top:10px;">
    <h3>SET {set_number}</h3>
    <label>Вага (kg)</label>
    <input id="{key}_weight" inputmode="decimal"
           placeholder="Наприклад 100" autocomplete="off">
    <label>Повторення</label>
    <input id="{key}_reps" inputmode="numeric"
           placeholder="Наприклад 10" autocomplete="off">
    <label>RIR</label>
    <input id="{key}_rir" inputmode="decimal"
           placeholder="Наприклад 2" autocomplete="off">
    <label>Notes</label>
    <input id="{key}_notes"
           placeholder="Необов'язково" autocomplete="off">
    <button type="button" class="button secondary"
            style="border:0;width:100%;"
            onclick="saveGuidedSet('{key}')">
        💾 ЗБЕРЕГТИ SET
    </button>
    <div id="{key}_status" class="muted"></div>
</div>
"""

    if index > 0:
        previous_url = f"/workout/{session}/guided?exercise={index-1}"
    else:
        previous_url = f"/workout/{session}"

    if index < len(strength_plan) - 1:
        next_url = f"/workout/{session}/guided?exercise={index+1}"
        next_label = "НАСТУПНА ВПРАВА →"
    else:
        next_url = f"/workout/{session}/guided/wod"
        next_label = "🔥 ПЕРЕЙТИ ДО WOD"

    wod_summary = ""
    if wods:
        wod = wods[0]
        wod_summary = f"""
<div class="card">
    <div class="muted">ПІСЛЯ СИЛОВОЇ</div>
    <h3>🔥 {core.value_or_dash(wod.get("format"))}</h3>
    <p>{core.value_or_dash(wod.get("description"))}</p>
</div>
"""

    load_calls = "\n".join(
        f"loadGuidedSet('guided_{session}_{index}_{n}');"
        for n in range(1, sets_count + 1)
    )

    body = f"""
<div class="card">
    <div class="muted">{session} · Week {CURRENT_WEEK}</div>
    <h1>🏋️ ВПРАВА {index + 1}/{len(strength_plan)}</h1>

    <div style="height:8px;background:#e5e7eb;border-radius:8px;
                overflow:hidden;margin:12px 0 20px;">
        <div style="width:{progress}%;height:100%;
                    background:#111827;"></div>
    </div>

    <h2>{exercise["exercise"]}</h2>

    <div class="grid">
        <div>
            <div class="muted">SETS</div>
            <div class="stat">{exercise["sets"]}</div>
        </div>
        <div>
            <div class="muted">REPS</div>
            <div class="stat">{exercise["reps"]}</div>
        </div>
    </div>

    <p>
        <b>Tempo:</b> {core.value_or_dash(exercise.get("tempo"))}
        · <b>Rest:</b> {core.value_or_dash(exercise.get("rest"))} sec
    </p>
    <p>
        🎯 Target RIR:
        <b>{core.value_or_dash(exercise.get("target_rir"))}</b>
    </p>
</div>

<div class="card">
    <h3>🧠 РЕКОМЕНДАЦІЯ</h3>
    <p><b>{recommendation[0]}</b></p>
    <p class="muted">{recommendation[1]}</p>
</div>

{sets_html}

<div class="card">
    <h2>⏱️ ТАЙМЕР ВІДПОЧИНКУ</h2>

    <div id="restTimerDisplay"
         style="font-size:42px;font-weight:800;text-align:center;
                margin:14px 0;">
        00:00
    </div>

    <div class="grid">
        <button type="button"
                class="button secondary"
                style="border:0;"
                onclick="startRestTimer(30)">
            30 сек
        </button>

        <button type="button"
                class="button secondary"
                style="border:0;"
                onclick="startRestTimer(60)">
            60 сек
        </button>

        <button type="button"
                class="button secondary"
                style="border:0;"
                onclick="startRestTimer(90)">
            90 сек
        </button>

        <button type="button"
                class="button secondary"
                style="border:0;"
                onclick="startRestTimer(120)">
            120 сек
        </button>
    </div>

    <button type="button"
            class="button"
            style="border:0;width:100%;"
            onclick="startRestTimer({int(float(exercise.get("rest", 60) or 60)) if str(exercise.get("rest", "")).replace(".", "", 1).isdigit() else 60})">
        ▶️ ПОЧАТИ ВІДПОЧИНОК ЗА ПЛАНОМ
    </button>

    <button type="button"
            class="button secondary"
            style="border:0;width:100%;"
            onclick="stopRestTimer()">
        ⏹️ ЗУПИНИТИ
    </button>
</div>

<div class="card">
    <p class="muted">
        Збережені сети залишаються в браузері до фінального запису.
    </p>
    <a class="button" href="{next_url}">{next_label}</a>
    <a class="button secondary" href="{previous_url}">← Назад</a>
</div>

{wod_summary}

<script>
let restTimerInterval = null;
let restSeconds = 0;

function renderRestTimer() {{
    const display = document.getElementById("restTimerDisplay");
    if (!display) return;

    const minutes = Math.floor(restSeconds / 60);
    const seconds = restSeconds % 60;

    display.textContent =
        String(minutes).padStart(2, "0") + ":" +
        String(seconds).padStart(2, "0");
}}

function startRestTimer(seconds) {{
    stopRestTimer();

    restSeconds = Number(seconds) || 0;
    renderRestTimer();

    restTimerInterval = setInterval(function() {{
        restSeconds -= 1;
        renderRestTimer();

        if (restSeconds <= 0) {{
            stopRestTimer();
            if (navigator.vibrate) {{
                navigator.vibrate([200, 100, 200]);
            }}
        }}
    }}, 1000);
}}

function stopRestTimer() {{
    if (restTimerInterval !== null) {{
        clearInterval(restTimerInterval);
        restTimerInterval = null;
    }}
}}

function guidedKey(prefix) {{
    return "TRAINING_SYSTEM_GUIDED_" + prefix;
}}

async function saveGuidedSet(key) {{
    const data = {{
        weight: document.getElementById(key + "_weight").value,
        reps: document.getElementById(key + "_reps").value,
        rir: document.getElementById(key + "_rir").value,
        notes: document.getElementById(key + "_notes").value
    }};

    localStorage.setItem(guidedKey(key), JSON.stringify(data));

    const parts = key.split("_");
    const setNumber = parts[parts.length - 1];
    const exerciseIndex = parts[parts.length - 2];

    const workoutKey =
        "TRAINING_SYSTEM_WORKOUT_ID_{session}_{CURRENT_WEEK}_" +
        new Date().toISOString().slice(0, 10);

    const savedWorkoutId =
        localStorage.getItem(workoutKey) || null;

    const response = await fetch("/api/guided/save", {{
        method: "POST",
        headers: {{
            "Content-Type": "application/json"
        }},
        body: JSON.stringify({{
            session: "{session}",
            week: {CURRENT_WEEK},
            workout_id: savedWorkoutId,
            exercise: {exercise["exercise"]!r},
            set: setNumber,
            weight: data.weight,
            reps: data.reps,
            rir: data.rir,
            notes: data.notes
        }})
    }});

    const status = document.getElementById(key + "_status");

    try {{
        const result = await response.json();

        if (result.ok && result.workout_id) {{
            localStorage.setItem(
                workoutKey,
                String(result.workout_id)
            );
        }}

        status.textContent = result.ok
            ? "✅ SET записано в Excel"
            : "❌ " + (result.error || "Помилка запису");

    }} catch (error) {{
        console.error("Guided save error:", error);
        status.textContent =
            "❌ Сервер не повернув коректну відповідь";

    }}
}}

function loadGuidedSet(key) {{
    const raw = localStorage.getItem(guidedKey(key));
    if (!raw) return;

    try {{
        const data = JSON.parse(raw);
        document.getElementById(key + "_weight").value = data.weight || "";
        document.getElementById(key + "_reps").value = data.reps || "";
        document.getElementById(key + "_rir").value = data.rir || "";
        document.getElementById(key + "_notes").value = data.notes || "";
        document.getElementById(key + "_status").textContent =
            "📱 Дані відновлено";
    }} catch (e) {{}}
}}

window.addEventListener("DOMContentLoaded", function() {{
    {load_calls}
}});
</script>
"""

    return page(f"{session} — Guided", body)



@app.route("/workout/<session>/guided/wod")
def guided_wod(session):
    workbook = get_workbook()
    session = str(session).strip().upper()

    if session not in SESSION_TYPES:
        return "Невірний тип тренування", 404

    wods = core.get_wods(
        workbook,
        CURRENT_WEEK,
        session
    )

    wod_html = ""

    for index, wod in enumerate(wods):
        wod_html += f"""
<div class="card">
    <div class="muted">🔥 WOD {index + 1}</div>
    <h2>{core.value_or_dash(wod.get("format"))}</h2>
    <p>{core.value_or_dash(wod.get("description"))}</p>
</div>
"""

    if not wod_html:
        wod_html = """
<div class="card">
    <h3>🔥 WOD не знайдено</h3>
    <p>Для цієї сесії WOD у плані відсутній.</p>
</div>
"""

    body = f"""
<div class="card">
    <div class="muted">{session} · Week {CURRENT_WEEK}</div>
    <h1>🔥 WOD</h1>
    <p>
        Силова частина завершена.
        Виконай WOD і внеси фінальні дані.
    </p>
</div>

{wod_html}

<div class="card">
    <h2>📝 РЕЗУЛЬТАТ WOD</h2>

    <label>Результат</label>
    <textarea id="finish_wod_result"
              rows="3"
              placeholder="Час / reps / rounds / результат"></textarea>

    <label>WOD RPE</label>
    <input id="finish_wod_rpe"
           inputmode="decimal"
           placeholder="Наприклад 8">

    <label>WOD Notes</label>
    <textarea id="finish_wod_notes"
              rows="2"
              placeholder="Необов'язково"></textarea>
</div>

<div class="card">
    <h2>📊 ПІДСУМОК ТРЕНУВАННЯ</h2>

    <label>Тривалість (хв)</label>
    <input id="finish_duration"
           inputmode="decimal"
           placeholder="Наприклад 65">

    <label>Загальний результат</label>
    <textarea id="finish_overall_result"
              rows="2"
              placeholder="Необов'язково"></textarea>

    <label>Загальний RPE</label>
    <input id="finish_overall_rpe"
           inputmode="decimal"
           placeholder="Наприклад 8">

    <label>PR</label>
    <select id="finish_pr">
        <option value="ні">Ні</option>
        <option value="так">Так</option>
    </select>

    <label>Загальні Notes</label>
    <textarea id="finish_notes"
              rows="3"
              placeholder="Необов'язково"></textarea>
</div>

<div id="finish_status" class="card muted"></div>

<button type="button"
        class="button"
        style="border:0;width:100%;font-size:18px;"
        onclick="finishGuidedWorkout()">
    ✅ ЗАВЕРШИТИ ТРЕНУВАННЯ
</button>

<a class="button secondary"
   href="/workout/{session}/guided">
    ← До силової частини
</a>

<a class="button secondary" href="/">
    ← Головна
</a>

<script>
async function finishGuidedWorkout() {{
    const status = document.getElementById("finish_status");

    const workoutKey =
        "TRAINING_SYSTEM_WORKOUT_ID_{session}_{CURRENT_WEEK}_" +
        new Date().toISOString().slice(0, 10);

    const workoutId = localStorage.getItem(workoutKey);

    if (!workoutId) {{
        status.textContent =
            "❌ Workout ID не знайдено. Спочатку збережи хоча б один SET.";
        return;
    }}

    const payload = {{
        session: "{session}",
        week: {CURRENT_WEEK},
        workout_id: workoutId,

        wod_result:
            document.getElementById("finish_wod_result").value,

        wod_rpe:
            document.getElementById("finish_wod_rpe").value,

        wod_notes:
            document.getElementById("finish_wod_notes").value,

        duration:
            document.getElementById("finish_duration").value,

        overall_result:
            document.getElementById("finish_overall_result").value,

        overall_rpe:
            document.getElementById("finish_overall_rpe").value,

        pr:
            document.getElementById("finish_pr").value,

        notes:
            document.getElementById("finish_notes").value
    }};

    status.textContent = "⏳ Зберігаємо тренування...";

    try {{
        const response = await fetch(
            "/api/guided/finish",
            {{
                method: "POST",
                headers: {{
                    "Content-Type": "application/json"
                }},
                body: JSON.stringify(payload)
            }}
        );

        const result = await response.json();

        if (!result.ok) {{
            status.textContent =
                "❌ " + (result.error || "Помилка збереження");
            return;
        }}

        localStorage.removeItem(workoutKey);

        status.textContent =
            "✅ Тренування завершено!";

        window.location.href =
            "/workout/complete/" + result.workout_id;

    }} catch (error) {{
        console.error(error);
        status.textContent =
            "❌ Не вдалося зв'язатися із сервером.";
    }}
}}
</script>
"""

    return page(
        f"{session} — WOD",
        body
    )



@app.route("/workout/<session>")
def workout(session):
    session = session.upper()

    if session not in SESSION_TYPES:
        return "Невірний тип тренування", 404

    workbook = get_workbook()

    exercises = core.get_strength_plan(
        workbook,
        CURRENT_WEEK,
        session
    )

    wods = core.get_wods(
        workbook,
        CURRENT_WEEK,
        session
    )

    exercise_html = ""

    for index, exercise in enumerate(exercises):
        try:
            sets_count = int(float(exercise["sets"]))
        except (TypeError, ValueError):
            sets_count = 1

        recommendation = core.calculate_next_session_recommendation(
            workbook,
            exercise,
            session,
            CURRENT_WEEK
        )

        sets_html = ""

        for set_number in range(1, sets_count + 1):
            prefix = f"ex_{index}_set_{set_number}"

            sets_html += f"""
<div class="card">
    <b>SET {set_number}</b>

    <label>Вага (kg)</label>
    <input
        inputmode="decimal"
        name="{prefix}_weight"
        placeholder="Наприклад 100"
    >

    <label>Повторення</label>
    <input
        inputmode="numeric"
        name="{prefix}_reps"
        placeholder="Наприклад 10"
    >

    <label>RIR</label>
    <input
        inputmode="decimal"
        name="{prefix}_rir"
        placeholder="Наприклад 3"
    >

    <label>Notes</label>
    <input
        name="{prefix}_notes"
        placeholder="Необов'язково"
    >
</div>
"""

        exercise_html += f"""
<div class="card">
    <h2>{index + 1}. {exercise["exercise"]}</h2>

    <p>
        <b>План: {exercise["sets"]} × {exercise["reps"]}</b>
    </p>

    <p class="muted">
        Tempo: {core.value_or_dash(exercise["tempo"])}
        · Rest: {core.value_or_dash(exercise["rest"])} sec
    </p>

    <p>
        🎯 <b>{recommendation[0]}</b><br>
        <span class="muted">{recommendation[1]}</span>
    </p>

    {sets_html}
</div>
"""

    wod_html = ""

    for index, wod in enumerate(wods):
        wod_html += f"""
<div class="card">
    <h2>🔥 WOD {index + 1}</h2>
    <p><b>{core.value_or_dash(wod["format"])}</b></p>
    <p>{core.value_or_dash(wod["description"])}</p>
</div>
"""

    body = f"""
<form method="post" action="/workout/{session}/save">

<div class="card">
    <h1>🏋️ {session}</h1>
    <p>Week {CURRENT_WEEK} · введення результатів</p>
</div>

{exercise_html}

{wod_html}

<div class="card">
    <h2>🔥 WOD результат</h2>

    <label>Результат WOD</label>
    <textarea
        name="wod_result"
        rows="3"
        placeholder="Час / reps / rounds / результат"
    ></textarea>

    <label>WOD RPE</label>
    <input
        inputmode="decimal"
        name="wod_rpe"
        placeholder="Наприклад 8"
    >

    <label>WOD Notes</label>
    <textarea
        name="wod_notes"
        rows="2"
    ></textarea>
</div>

<div class="card">
    <h2>📊 Підсумок</h2>

    <label>Тривалість (хв)</label>
    <input
        inputmode="decimal"
        name="duration"
        placeholder="Наприклад 65"
    >

    <label>Загальний результат</label>
    <textarea
        name="overall_result"
        rows="2"
    ></textarea>

    <label>Загальний RPE</label>
    <input
        inputmode="decimal"
        name="overall_rpe"
        placeholder="Наприклад 8"
    >

    <label>PR</label>
    <select name="pr">
        <option value="ні">Ні</option>
        <option value="так">Так</option>
    </select>

    <label>Загальні Notes</label>
    <textarea
        name="notes"
        rows="3"
    ></textarea>
</div>

<button
    class="button"
    type="submit"
    style="border:0;width:100%;font-size:17px;"
>
    ✅ ЗАВЕРШИТИ ТРЕНУВАННЯ
</button>

</form>

<a class="button secondary" href="/">← Скасувати</a>
"""

    return page(
        f"{session} — Training System",
        body
    )


if __name__ == "__main__":
    print("=" * 60)
    print("🔥 TRAINING SYSTEM v4.7 — WEB APP")
    print("=" * 60)
    print(f"Ядро: {CORE_PATH.name}")
    print(f"Excel: {FILE_NAME}")
    print("Відкрий у браузері: http://127.0.0.1:5000")
    print("Для телефону в тій самій Wi-Fi мережі:")
    print("http://IP-АДРЕСА-КОМП'ЮТЕРА:5000")
    print("=" * 60)

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=False
    )
