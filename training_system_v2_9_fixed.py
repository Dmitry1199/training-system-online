from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook


# ============================================================
# TRAINING SYSTEM v2.1
# Працює з training_data_v2.xlsx
# ============================================================

FILE_NAME = "training_data_v2.xlsx"


# ============================================================
# НАЛАШТУВАННЯ
# ============================================================

def get_base_folder():
    """Папка, де знаходиться цей Python-файл."""
    return Path(__file__).resolve().parent


def get_file_path():
    """Повний шлях до Excel-бази."""
    return get_base_folder() / FILE_NAME


# ============================================================
# EXCEL
# ============================================================

def open_training_database():
    file_path = get_file_path()

    if not file_path.exists():
        print("\n❌ Excel-файл не знайдено.")
        print(f"Очікуваний файл:")
        print(file_path)
        print("\nПоклади training_data_v2.xlsx в ту саму папку, що й training_system_v2.py.")
        return None

    try:
        workbook = load_workbook(file_path)
        print("\n✅ Excel-база успішно підключена!")
        print(f"📁 {file_path}")
        return workbook

    except PermissionError:
        print("\n❌ Не вдалося відкрити Excel.")
        print("Закрий training_data_v2.xlsx в Excel і запусти програму ще раз.")
        return None

    except Exception as error:
        print("\n❌ Помилка відкриття Excel:")
        print(error)
        return None


def save_database(workbook):
    file_path = get_file_path()

    try:
        workbook.save(file_path)
        return True

    except PermissionError:
        print("\n❌ Excel-файл відкритий в Excel.")
        print("Закрий training_data_v2.xlsx і повтори збереження.")
        return False

    except Exception as error:
        print("\n❌ Помилка збереження:")
        print(error)
        return False


# ============================================================
# ДОПОМІЖНІ ФУНКЦІЇ
# ============================================================

def normalize(value):
    if value is None:
        return ""
    return str(value).strip().casefold()


def value_or_dash(value):
    if value is None or str(value).strip() == "":
        return "—"
    return value


def ask_required(prompt):
    while True:
        value = input(prompt).strip()

        if value:
            return value

        print("❌ Поле не може бути порожнім.")


def ask_optional(prompt):
    return input(prompt).strip()


def ask_number(prompt, allow_empty=False):
    while True:
        value = input(prompt).strip().replace(",", ".")

        if allow_empty and value == "":
            return ""

        try:
            return float(value)
        except ValueError:
            print("❌ Введи число.")


def ask_integer(prompt, allow_empty=False):
    while True:
        value = input(prompt).strip()

        if allow_empty and value == "":
            return ""

        try:
            return int(value)
        except ValueError:
            print("❌ Введи ціле число.")


# ============================================================
# WEEKLY PLAN
# ============================================================

def get_week_info(workbook, week):
    if "WEEKLY_PLAN" not in workbook.sheetnames:
        return None

    sheet = workbook["WEEKLY_PLAN"]

    for row in range(2, sheet.max_row + 1):
        if str(sheet.cell(row, 1).value).strip() == str(week):
            return {
                "target_rir": sheet.cell(row, 2).value,
                "phase": sheet.cell(row, 3).value,
                "instructions": sheet.cell(row, 4).value,
            }

    return None


def get_available_weeks(workbook):
    if "WEEKLY_PLAN" not in workbook.sheetnames:
        return []

    sheet = workbook["WEEKLY_PLAN"]
    weeks = []

    for row in range(2, sheet.max_row + 1):
        value = sheet.cell(row, 1).value

        if value is not None:
            try:
                weeks.append(int(value))
            except (TypeError, ValueError):
                pass

    return sorted(set(weeks))


def choose_week(workbook, current_week):
    weeks = get_available_weeks(workbook)

    if not weeks:
        print("\n❌ У WEEKLY_PLAN немає тижнів.")
        return current_week

    print("\nДоступні тижні:")
    print("  " + "  ".join(str(w) for w in weeks))
    print(f"\nПоточний тиждень: {current_week}")

    value = input(
        "Введи номер тижня або Enter, щоб залишити поточний: "
    ).strip()

    if value == "":
        return current_week

    try:
        week = int(value)

        if week in weeks:
            return week

    except ValueError:
        pass

    print("❌ Невірний тиждень. Залишаю поточний.")
    return current_week


# ============================================================
# STRENGTH PLAN
# ============================================================

def get_strength_plan(workbook, week, session):
    if "STRENGTH_PLAN" not in workbook.sheetnames:
        return []

    sheet = workbook["STRENGTH_PLAN"]
    result = []

    for row in range(2, sheet.max_row + 1):
        row_week = sheet.cell(row, 1).value
        row_session = sheet.cell(row, 2).value

        if str(row_week).strip() != str(week):
            continue

        if normalize(row_session) != normalize(session):
            continue

        result.append({
            "exercise": sheet.cell(row, 3).value,
            "sets": sheet.cell(row, 4).value,
            "reps": sheet.cell(row, 5).value,
            "tempo": sheet.cell(row, 6).value,
            "rest": sheet.cell(row, 7).value,
            "rir": sheet.cell(row, 8).value,
            "notes": sheet.cell(row, 9).value,
            "alternatives": sheet.cell(row, 10).value,
        })

    return result


# ============================================================
# WOD PLAN
# ============================================================

def get_wods(workbook, week, session):
    if "WOD_PLAN" not in workbook.sheetnames:
        return []

    sheet = workbook["WOD_PLAN"]
    result = []

    for row in range(2, sheet.max_row + 1):
        row_week = sheet.cell(row, 2).value
        row_session = sheet.cell(row, 3).value

        if str(row_week).strip() != str(week):
            continue

        if normalize(row_session) != normalize(session):
            continue

        result.append({
            "id": sheet.cell(row, 1).value,
            "week": sheet.cell(row, 2).value,
            "session": sheet.cell(row, 3).value,
            "format": sheet.cell(row, 4).value,
            "description": sheet.cell(row, 5).value,
            "time_cap": sheet.cell(row, 6).value,
            "intensity": sheet.cell(row, 7).value,
            "benchmark": sheet.cell(row, 8).value,
        })

    return result


# ============================================================
# WOD COMPONENTS
# ============================================================

def get_wod_components(workbook, wod_id):
    if "WOD_COMPONENTS" not in workbook.sheetnames:
        return []

    sheet = workbook["WOD_COMPONENTS"]
    result = []

    for row in range(2, sheet.max_row + 1):
        if str(sheet.cell(row, 1).value).strip() != str(wod_id):
            continue

        result.append({
            "order": sheet.cell(row, 2).value,
            "component": sheet.cell(row, 3).value,
            "target": sheet.cell(row, 4).value,
            "detail": sheet.cell(row, 5).value,
            "notes": sheet.cell(row, 6).value,
        })

    return sorted(result, key=lambda x: int(x["order"]) if str(x["order"]).isdigit() else 999)


# ============================================================
# DISPLAY CURRENT PLAN
# ============================================================

def show_week_info(workbook, week):
    info = get_week_info(workbook, week)

    print("\n" + "=" * 70)
    print(f"📅 WEEK {week}")
    print("=" * 70)

    if not info:
        print("❌ Інформацію про тиждень не знайдено.")
        return

    print(f"Фаза:        {value_or_dash(info['phase'])}")
    print(f"Target RIR:  {value_or_dash(info['target_rir'])}")
    print(f"Правила:     {value_or_dash(info['instructions'])}")


def show_strength_plan(workbook, week, session):
    exercises = get_strength_plan(workbook, week, session)

    print("\n" + "-" * 70)
    print("💪 СИЛОВА ЧАСТИНА")
    print("-" * 70)

    if not exercises:
        print("Силову частину не знайдено.")
        return

    for index, exercise in enumerate(exercises, 1):
        print(f"\n{index}. {value_or_dash(exercise['exercise'])}")
        print(f"   Sets:       {value_or_dash(exercise['sets'])}")
        print(f"   Reps:       {value_or_dash(exercise['reps'])}")
        print(f"   Tempo:      {value_or_dash(exercise['tempo'])}")
        print(f"   Rest:       {value_or_dash(exercise['rest'])} sec")
        print(f"   Target RIR: {value_or_dash(exercise['rir'])}")

        if exercise["notes"]:
            print(f"   Notes:      {exercise['notes']}")

        if exercise["alternatives"]:
            print(f"   Альтернативи: {exercise['alternatives']}")


def show_wod_plan(workbook, week, session):
    wods = get_wods(workbook, week, session)

    print("\n" + "-" * 70)
    print("🔥 WOD")
    print("-" * 70)

    if not wods:
        print("WOD не знайдено.")
        return

    for index, wod in enumerate(wods, 1):
        print(f"\nWOD {index} | ID: {wod['id']}")
        print(f"   Format:       {value_or_dash(wod['format'])}")
        print(f"   Description:  {value_or_dash(wod['description'])}")
        print(f"   Time / Rest:  {value_or_dash(wod['time_cap'])}")
        print(f"   Intensity:    {value_or_dash(wod['intensity'])}")

        if normalize(wod["benchmark"]) in ("yes", "так", "true"):
            print("   ⭐ BENCHMARK")

        components = get_wod_components(workbook, wod["id"])

        if components:
            print("   Components:")

            for component in components:
                print(
                    f"      {component['order']}. "
                    f"{value_or_dash(component['component'])} — "
                    f"{value_or_dash(component['target'])} "
                    f"{value_or_dash(component['detail'])}"
                )


# ============================================================
# CHOOSE WOD
# ============================================================

def choose_wod(workbook, week, session):
    wods = get_wods(workbook, week, session)

    if not wods:
        return None

    if len(wods) == 1:
        return wods[0]

    print("\nОберіть WOD:")

    for index, wod in enumerate(wods, 1):
        print(
            f"{index} — "
            f"{wod['format']} | "
            f"{wod['description']}"
        )

    while True:
        choice = input("\nНомер WOD: ").strip()

        try:
            index = int(choice)

            if 1 <= index <= len(wods):
                return wods[index - 1]

        except ValueError:
            pass

        print("❌ Невірний вибір.")


# ============================================================
# WORKOUT ID
# ============================================================

def get_next_id(sheet):
    ids = []

    for row in range(2, sheet.max_row + 1):
        value = sheet.cell(row, 1).value

        try:
            ids.append(int(value))
        except (TypeError, ValueError):
            pass

    return max(ids, default=0) + 1


# ============================================================
# START SESSION
# ============================================================

def create_workout_record(workbook, week, session, wod):
    if "WORKOUTS" not in workbook.sheetnames:
        print("\n❌ Аркуш WORKOUTS не знайдено.")
        return None

    sheet = workbook["WORKOUTS"]
    workout_id = get_next_id(sheet)
    today = datetime.now().strftime("%Y-%m-%d")

    row = sheet.max_row + 1

    sheet.cell(row, 1).value = workout_id
    sheet.cell(row, 2).value = today
    sheet.cell(row, 3).value = week
    sheet.cell(row, 4).value = session

    if wod:
        sheet.cell(row, 5).value = wod["id"]
        sheet.cell(row, 6).value = wod["description"]
        sheet.cell(row, 7).value = wod["format"]

    return workout_id


# ============================================================
# EXERCISE BRIEFING — v2.3
# ============================================================

def show_exercise_briefing(workbook, week, session, exercise):
    """
    Показує план + попередній результат + рекомендацію
    без зміни даних у плані.
    """
    print("\n" + "=" * 70)
    print(f"🏋️ {exercise['exercise']}")
    print("=" * 70)

    print(
        f"План: {exercise['sets']} × {exercise['reps']}"
    )
    print(
        f"Tempo: {value_or_dash(exercise['tempo'])} | "
        f"Rest: {value_or_dash(exercise['rest'])} sec"
    )
    print(
        f"Target RIR: {value_or_dash(exercise['rir'])}"
    )

    if exercise["notes"]:
        print(f"Notes: {exercise['notes']}")

    if exercise["alternatives"]:
        print(
            f"Альтернативи: {exercise['alternatives']}"
        )

    previous = get_previous_exercise_data(
        workbook,
        exercise["exercise"],
        session
    )

    recommendation = recommend_progression(
        exercise,
        previous,
        week
    )

    print("\n📊 ПОПЕРЕДНІЙ РЕЗУЛЬТАТ")

    if not previous:
        print("Немає попередньої історії.")
        print(
            "➡️ Рекомендація: виконати поточний план."
        )
        return

    print(
        f"Workout #{previous['workout_id']} | "
        f"{previous['date']}"
    )

    for item in previous["sets"]:
        print(
            f"  Set {item['set']}: "
            f"{value_or_dash(item['weight'])} kg × "
            f"{value_or_dash(item['reps'])} reps | "
            f"RIR {value_or_dash(item['rir'])}"
        )

    print("\n🧠 РЕКОМЕНДАЦІЯ")

    if recommendation["status"] == "PROGRESS":
        print("🟢 МОЖНА ПРОГРЕСУВАТИ")
    elif recommendation["status"] == "HOLD":
        print("🟡 ЗАЛИШИТИ ВАГУ")
    elif recommendation["status"] == "DELOAD":
        print("🔵 DELOAD")
    else:
        print("⚠️ ПОТРІБНА ДОДАТКОВА ОЦІНКА")

    print(recommendation["message"])
    print(f"➡️ {recommendation['recommendation']}")


# ============================================================
# ENTER STRENGTH RESULTS
# ============================================================

def choose_exercise_version(exercise):
    """
    Дозволяє залишити основну вправу або вибрати одну з альтернатив.
    Повертає назву вправи, яку реально виконуватимемо.
    """
    main_exercise = str(exercise["exercise"]).strip()
    alternatives_raw = exercise["alternatives"]

    if not alternatives_raw:
        return main_exercise

    alternatives = [
        item.strip()
        for item in str(alternatives_raw).split("→")
        if item.strip()
    ]

    print("\nВибір вправи:")
    print(f"1 — {main_exercise}")

    for index, alternative in enumerate(alternatives, start=2):
        print(f"{index} — {alternative}")

    print("Enter — залишити основну вправу")

    while True:
        choice = input("Обери: ").strip()

        if choice == "":
            return main_exercise

        try:
            number = int(choice)

            if number == 1:
                return main_exercise

            alternative_index = number - 2

            if 0 <= alternative_index < len(alternatives):
                return alternatives[alternative_index]

        except ValueError:
            pass

        print("❌ Невірний вибір.")


def enter_strength_results(workbook, week, session, workout_id):
    exercises = get_strength_plan(workbook, week, session)

    if not exercises:
        print("\nℹ️ Силових вправ для цього тренування немає.")
        return

    sheet = workbook["STRENGTH_RESULTS"]
    next_result_id = get_next_id(sheet)
    today = datetime.now().strftime("%Y-%m-%d")

    print("\n" + "=" * 70)
    print("💪 ВНЕСЕННЯ ФАКТИЧНИХ РЕЗУЛЬТАТІВ")
    print("=" * 70)
    print(f"Workout ID: {workout_id}")
    print(f"Дата:       {today}")
    print(f"Week:       {week}")
    print(f"Session:    {session}")

    for exercise in exercises:
        show_exercise_briefing(
            workbook,
            week,
            session,
            exercise
        )

        actual_exercise = choose_exercise_version(exercise)

        if actual_exercise != str(exercise["exercise"]).strip():
            print(
                f"🔄 Виконується альтернатива: {actual_exercise}"
            )

        try:
            number_of_sets = int(exercise["sets"])
        except (TypeError, ValueError):
            print("⚠️ Кількість підходів не визначена — пропускаю.")
            continue

        for set_number in range(1, number_of_sets + 1):
            print(f"\n{actual_exercise} — Set {set_number}")

            weight = ask_number(
                "  Вага (kg): ",
                allow_empty=True
            )

            reps = ask_integer(
                "  Повтори: ",
                allow_empty=True
            )

            rir = ask_number(
                "  Фактичний RIR: ",
                allow_empty=True
            )

            notes = ask_optional(
                "  Notes: "
            )

            row = sheet.max_row + 1

            sheet.cell(row, 1).value = next_result_id
            sheet.cell(row, 2).value = workout_id
            sheet.cell(row, 3).value = today
            sheet.cell(row, 4).value = week
            sheet.cell(row, 5).value = session
            sheet.cell(row, 6).value = actual_exercise
            sheet.cell(row, 7).value = set_number
            sheet.cell(row, 8).value = weight
            sheet.cell(row, 9).value = reps
            sheet.cell(row, 10).value = rir
            sheet.cell(row, 11).value = notes

            next_result_id += 1

    print("\n✅ Силові результати підготовлені.")


# ============================================================
# ENTER WOD RESULT
# ============================================================

def enter_wod_result(workbook, week, session, workout_id, wod):
    if "WOD_RESULTS" not in workbook.sheetnames:
        print("\n❌ Аркуш WOD_RESULTS не знайдено.")
        return

    if not wod:
        print("\nℹ️ WOD для цього тренування немає.")
        return

    sheet = workbook["WOD_RESULTS"]
    result_id = get_next_id(sheet)
    today = datetime.now().strftime("%Y-%m-%d")

    print("\n" + "=" * 70)
    print("🔥 РЕЗУЛЬТАТ WOD")
    print("=" * 70)

    print(f"WOD ID:       {wod['id']}")
    print(f"Format:       {wod['format']}")
    print(f"Description:  {wod['description']}")

    scaling = ask_optional(
        "\nВага / Scaling: "
    )

    time_result = ask_optional(
        "Час: "
    )

    rounds_result = ask_optional(
        "Раунди / повтори / калорії: "
    )

    rpe = ask_number(
        "RPE (1-10): ",
        allow_empty=True
    )

    notes = ask_optional(
        "Notes: "
    )

    row = sheet.max_row + 1

    sheet.cell(row, 1).value = result_id
    sheet.cell(row, 2).value = workout_id
    sheet.cell(row, 3).value = today
    sheet.cell(row, 4).value = week
    sheet.cell(row, 5).value = session
    sheet.cell(row, 6).value = wod["id"]
    sheet.cell(row, 7).value = wod["format"]
    sheet.cell(row, 8).value = scaling
    sheet.cell(row, 9).value = time_result
    sheet.cell(row, 10).value = rounds_result
    sheet.cell(row, 11).value = rpe
    sheet.cell(row, 12).value = notes

    print("\n✅ Результат WOD підготовлений.")


# ============================================================
# FINISH WORKOUT
# ============================================================

def finish_workout(workbook, workout_id):
    if "WORKOUTS" not in workbook.sheetnames:
        return

    sheet = workbook["WORKOUTS"]

    target_row = None

    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row, 1).value == workout_id:
            target_row = row
            break

    if target_row is None:
        print("\n❌ Workout ID не знайдено.")
        return

    print("\n" + "=" * 70)
    print("📝 ЗАВЕРШЕННЯ ТРЕНУВАННЯ")
    print("=" * 70)

    duration = ask_optional(
        "Тривалість тренування (хв): "
    )

    result = ask_optional(
        "Загальний результат WOD: "
    )

    rpe = ask_number(
        "Загальний RPE (1-10): ",
        allow_empty=True
    )

    pr = ask_optional(
        "PR? (так/ні): "
    )

    notes = ask_optional(
        "Загальні нотатки: "
    )

    sheet.cell(target_row, 8).value = duration
    sheet.cell(target_row, 9).value = result
    sheet.cell(target_row, 10).value = rpe
    sheet.cell(target_row, 11).value = pr
    sheet.cell(target_row, 12).value = notes

    print("\n✅ Загальна інформація про тренування заповнена.")


# ============================================================
# START WORKOUT
# ============================================================

def start_workout(workbook, week, session):
    show_week_info(workbook, week)

    print("\n" + "=" * 70)
    print(f"🔥 {session.upper()} — WEEK {week}")
    print("=" * 70)

    show_strength_plan(workbook, week, session)
    show_wod_plan(workbook, week, session)

    print("\n" + "=" * 70)
    print("1 — Почати це тренування")
    print("2 — Назад")

    choice = input("\nОбери дію: ").strip()

    if choice != "1":
        return

    wod = choose_wod(workbook, week, session)

    workout_id = create_workout_record(
        workbook,
        week,
        session,
        wod
    )

    if workout_id is None:
        return

    # Переконуємося, що аркуші існують.
    if "STRENGTH_RESULTS" not in workbook.sheetnames:
        print("\n❌ STRENGTH_RESULTS відсутній.")
        print("Перевір структуру training_data_v2.xlsx.")
        return

    # Силова частина
    enter_strength_results(
        workbook,
        week,
        session,
        workout_id
    )

    # WOD
    enter_wod_result(
        workbook,
        week,
        session,
        workout_id,
        wod
    )

    # Загальна інформація
    finish_workout(
        workbook,
        workout_id
    )

    # Зберігаємо все одним save()
    if save_database(workbook):
        print("\n" + "=" * 70)
        print("🎉 ТРЕНУВАННЯ ПОВНІСТЮ ЗБЕРЕЖЕНО!")
        print("=" * 70)
        print(f"Workout ID: {workout_id}")
        print(f"Excel: {get_file_path()}")
    else:
        print("\n⚠️ Дані залишилися в пам'яті програми, але файл не збережено.")

    input("\nНатисни Enter...")


# ============================================================
# HISTORY
# ============================================================

def show_history(workbook):
    if "WORKOUTS" not in workbook.sheetnames:
        print("\n❌ WORKOUTS не знайдено.")
        return

    sheet = workbook["WORKOUTS"]

    print("\n" + "=" * 90)
    print("📚 ІСТОРІЯ ТРЕНУВАНЬ")
    print("=" * 90)

    if sheet.max_row <= 1:
        print("\nІсторія поки порожня.")
        return

    start_row = max(2, sheet.max_row - 14)

    for row in range(start_row, sheet.max_row + 1):
        workout_id = sheet.cell(row, 1).value
        date = sheet.cell(row, 2).value
        week = sheet.cell(row, 3).value
        session = sheet.cell(row, 4).value
        wod_id = sheet.cell(row, 5).value
        duration = sheet.cell(row, 8).value
        result = sheet.cell(row, 9).value
        rpe = sheet.cell(row, 10).value
        pr = sheet.cell(row, 11).value

        print(
            f"\n#{workout_id} | {date} | "
            f"Week {week} | {session}"
        )

        print(
            f"   WOD ID: {wod_id} | "
            f"Duration: {value_or_dash(duration)} | "
            f"Result: {value_or_dash(result)} | "
            f"RPE: {value_or_dash(rpe)} | "
            f"PR: {value_or_dash(pr)}"
        )


# ============================================================
# ПРОГРЕСІЯ — v2.2
# ============================================================

def parse_rep_range(value):
    """
    Перетворює Target Reps на (min_reps, max_reps).
    Підтримує:
    8
    8–10
    8-10
    8–10/нога
    10/нога
    """
    if value is None:
        return None, None

    text = str(value).strip()
    text = text.replace(",", ".")
    text = text.replace("–", "-")
    text = text.split("/")[0].strip()

    import re

    numbers = re.findall(r"\d+(?:\.\d+)?", text)

    if not numbers:
        return None, None

    if len(numbers) == 1:
        number = float(numbers[0])
        return number, number

    return float(numbers[0]), float(numbers[1])


def get_last_exercise_session(workbook, exercise_name, session):
    """
    Знаходить останнє тренування цієї вправи
    саме в цьому типі сесії.
    """
    if "STRENGTH_RESULTS" not in workbook.sheetnames:
        return []

    sheet = workbook["STRENGTH_RESULTS"]

    matches = []

    for row in range(2, sheet.max_row + 1):
        stored_exercise = sheet.cell(row, 6).value
        stored_session = sheet.cell(row, 5).value

        if normalize(stored_exercise) != normalize(exercise_name):
            continue

        if normalize(stored_session) != normalize(session):
            continue

        matches.append(row)

    if not matches:
        return []

    # Останній Workout ID = остання виконана сесія.
    workout_ids = []

    for row in matches:
        value = sheet.cell(row, 2).value
        try:
            workout_ids.append(int(value))
        except (TypeError, ValueError):
            pass

    if not workout_ids:
        return []

    last_workout_id = max(workout_ids)

    return [
        row for row in matches
        if sheet.cell(row, 2).value == last_workout_id
    ]


def get_previous_exercise_data(workbook, exercise_name, session):
    """
    Повертає фактичні підходи з останньої сесії.
    """
    rows = get_last_exercise_session(
        workbook,
        exercise_name,
        session
    )

    if not rows:
        return None

    sheet = workbook["STRENGTH_RESULTS"]

    sets = []

    for row in sorted(
        rows,
        key=lambda r: (
            int(sheet.cell(r, 7).value)
            if str(sheet.cell(r, 7).value).isdigit()
            else 999
        )
    ):
        weight = sheet.cell(row, 8).value
        reps = sheet.cell(row, 9).value
        rir = sheet.cell(row, 10).value

        try:
            weight = float(weight)
        except (TypeError, ValueError):
            weight = None

        try:
            reps = int(reps)
        except (TypeError, ValueError):
            reps = None

        try:
            rir = float(rir)
        except (TypeError, ValueError):
            rir = None

        sets.append({
            "set": sheet.cell(row, 7).value,
            "weight": weight,
            "reps": reps,
            "rir": rir
        })

    return {
        "workout_id": sheet.cell(rows[0], 2).value,
        "date": sheet.cell(rows[0], 3).value,
        "sets": sets
    }


def recommend_progression(
    exercise_plan,
    previous_data,
    week
):
    """
    Консервативна рекомендація.
    На v2.2 система НІКОЛИ сама не змінює план.
    """

    if not previous_data:
        return {
            "status": "NO_HISTORY",
            "message": "Попереднього результату немає.",
            "recommendation": "Виконай вправу за поточним планом."
        }

    min_reps, max_reps = parse_rep_range(
        exercise_plan["reps"]
    )

    if max_reps is None:
        return {
            "status": "NO_REP_RANGE",
            "message": "Не вдалося визначити діапазон повторень.",
            "recommendation": "Оціни прогрес вручну."
        }

    valid_sets = [
        item for item in previous_data["sets"]
        if item["weight"] is not None
        and item["reps"] is not None
    ]

    if not valid_sets:
        return {
            "status": "NO_VALID_SETS",
            "message": "Немає достатньо фактичних даних.",
            "recommendation": "Повтори вправу за планом."
        }

    # Якщо дані були записані порожніми — не робимо висновок.
    if len(valid_sets) < len(previous_data["sets"]):
        return {
            "status": "INCOMPLETE",
            "message": "Попереднє тренування записане не повністю.",
            "recommendation": "Спочатку отримай повний набір даних."
        }

    weights = [item["weight"] for item in valid_sets]
    reps = [item["reps"] for item in valid_sets]

    if len(set(weights)) != 1:
        return {
            "status": "VARIABLE_LOAD",
            "message": "У підходах використовувалась різна вага.",
            "recommendation": "Повтори структуру або оціни прогрес вручну."
        }

    previous_weight = weights[0]

    reached_top = all(
        reps_value >= max_reps
        for reps_value in reps
    )

    rir_values = [
        item["rir"]
        for item in previous_data["sets"]
        if item["rir"] is not None
    ]

    target_rir_text = str(
        exercise_plan["rir"]
    ).strip()

    # --------------------------------------------------------
    # Deload: ніякого підвищення.
    # --------------------------------------------------------

    if int(week) == 8:
        return {
            "status": "DELOAD",
            "message": "8-й тиждень — Deload.",
            "recommendation": (
                "Не підвищувати вагу. "
                "Виконувати план у зниженій інтенсивності."
            ),
            "previous_weight": previous_weight
        }

    # --------------------------------------------------------
    # Якщо всі підходи досягли верхньої межі,
    # пропонуємо невелике підвищення.
    # --------------------------------------------------------

    if reached_top:
        if previous_weight <= 30:
            increase = 2
        elif previous_weight <= 60:
            increase = 2.5
        elif previous_weight <= 100:
            increase = 5
        else:
            increase = 5

        recommended_weight = previous_weight + increase

        return {
            "status": "PROGRESS",
            "message": (
                f"Усі підходи досягли верхньої межі "
                f"{int(max_reps) if max_reps.is_integer() else max_reps} повторень."
            ),
            "recommendation": (
                f"Спробувати {recommended_weight:g} kg "
                f"на наступній сесії."
            ),
            "previous_weight": previous_weight,
            "recommended_weight": recommended_weight,
            "previous_reps": reps,
            "previous_rir": rir_values,
            "target_rir": target_rir_text
        }

    # --------------------------------------------------------
    # Якщо верхню межу не досягнуто — залишаємо вагу.
    # --------------------------------------------------------

    return {
        "status": "HOLD",
        "message": (
            f"Не всі підходи досягли верхньої межі "
            f"{int(max_reps) if max_reps.is_integer() else max_reps}."
        ),
        "recommendation": (
            f"Залишити {previous_weight:g} kg "
            f"і спробувати збільшити повторення."
        ),
        "previous_weight": previous_weight,
        "previous_reps": reps,
        "previous_rir": rir_values,
        "target_rir": target_rir_text
    }


def show_progression_recommendation(
    workbook,
    week,
    session,
    exercise
):
    previous = get_previous_exercise_data(
        workbook,
        exercise["exercise"],
        session
    )

    recommendation = recommend_progression(
        exercise,
        previous,
        week
    )

    print("\n" + "-" * 70)
    print(f"🧠 ПРОГРЕСІЯ: {exercise['exercise']}")
    print("-" * 70)

    if not previous:
        print("\n📌 Історії ще немає.")
        print(
            "Рекомендація: виконати вправу за поточним планом."
        )
        return

    print(
        f"\nОстанній Workout ID: "
        f"{previous['workout_id']}"
    )

    print(
        f"Дата: {previous['date']}"
    )

    print("\nОстанні підходи:")

    for item in previous["sets"]:
        print(
            f"  Set {item['set']}: "
            f"{value_or_dash(item['weight'])} kg × "
            f"{value_or_dash(item['reps'])} reps | "
            f"RIR {value_or_dash(item['rir'])}"
        )

    print(
        f"\nПлан: {exercise['sets']} × {exercise['reps']}"
    )

    print(
        f"Target RIR: {exercise['rir']}"
    )

    if recommendation["status"] == "PROGRESS":
        print("\n🟢 ПРОГРЕС")
    elif recommendation["status"] == "HOLD":
        print("\n🟡 ЗАЛИШИТИ ВАГУ")
    elif recommendation["status"] == "DELOAD":
        print("\n🔵 DELOAD")
    else:
        print("\n⚠️ ПОТРІБНА ДОДАТКОВА ОЦІНКА")

    print(
        f"\n{recommendation['message']}"
    )

    print(
        f"➡️ {recommendation['recommendation']}"
    )


def show_all_progression(workbook, week, session):
    exercises = get_strength_plan(
        workbook,
        week,
        session
    )

    if not exercises:
        print("\n❌ Силову частину не знайдено.")
        return

    print("\n" + "=" * 70)
    print(
        f"🧠 ПРОГРЕСІЯ — {session} — WEEK {week}"
    )
    print("=" * 70)

    for exercise in exercises:
        show_progression_recommendation(
            workbook,
            week,
            session,
            exercise
        )


# ============================================================
# EXERCISE HISTORY
# ============================================================

def show_exercise_history(workbook):
    if "STRENGTH_RESULTS" not in workbook.sheetnames:
        print("\n❌ STRENGTH_RESULTS не знайдено.")
        return

    sheet = workbook["STRENGTH_RESULTS"]

    exercise_name = ask_required(
        "\nВведи назву вправи: "
    )

    matches = []

    for row in range(2, sheet.max_row + 1):
        exercise = sheet.cell(row, 6).value

        if normalize(exercise) == normalize(exercise_name):
            matches.append(row)

    print("\n" + "=" * 80)
    print(f"📈 ІСТОРІЯ: {exercise_name}")
    print("=" * 80)

    if not matches:
        print("\nРезультатів не знайдено.")
        return

    for row in matches:
        print(
            f"\nWorkout #{sheet.cell(row, 2).value} | "
            f"{sheet.cell(row, 3).value} | "
            f"Week {sheet.cell(row, 4).value} | "
            f"{sheet.cell(row, 5).value}"
        )

        print(
            f"   Set {sheet.cell(row, 7).value}: "
            f"{value_or_dash(sheet.cell(row, 8).value)} kg × "
            f"{value_or_dash(sheet.cell(row, 9).value)} reps | "
            f"RIR {value_or_dash(sheet.cell(row, 10).value)}"
        )

        if sheet.cell(row, 11).value:
            print(
                f"   Notes: {sheet.cell(row, 11).value}"
            )


# ============================================================
# ANALYTICS — v2.4
# ============================================================

def safe_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def get_strength_rows(workbook):
    if "STRENGTH_RESULTS" not in workbook.sheetnames:
        return []

    sheet = workbook["STRENGTH_RESULTS"]
    rows = []

    for row in range(2, sheet.max_row + 1):
        rows.append({
            "id": sheet.cell(row, 1).value,
            "workout_id": sheet.cell(row, 2).value,
            "date": sheet.cell(row, 3).value,
            "week": sheet.cell(row, 4).value,
            "session": sheet.cell(row, 5).value,
            "exercise": sheet.cell(row, 6).value,
            "set": sheet.cell(row, 7).value,
            "weight": safe_float(sheet.cell(row, 8).value),
            "reps": safe_float(sheet.cell(row, 9).value),
            "rir": safe_float(sheet.cell(row, 10).value),
            "notes": sheet.cell(row, 11).value,
        })

    return rows


def get_workout_rows(workbook):
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
            "duration": safe_float(sheet.cell(row, 8).value),
            "result": sheet.cell(row, 9).value,
            "rpe": safe_float(sheet.cell(row, 10).value),
            "pr": sheet.cell(row, 11).value,
            "notes": sheet.cell(row, 12).value,
        })

    return rows


def get_wod_result_rows(workbook):
    if "WOD_RESULTS" not in workbook.sheetnames:
        return []

    sheet = workbook["WOD_RESULTS"]
    rows = []

    for row in range(2, sheet.max_row + 1):
        rows.append({
            "id": sheet.cell(row, 1).value,
            "workout_id": sheet.cell(row, 2).value,
            "date": sheet.cell(row, 3).value,
            "week": sheet.cell(row, 4).value,
            "session": sheet.cell(row, 5).value,
            "wod_id": sheet.cell(row, 6).value,
            "format": sheet.cell(row, 7).value,
            "scaling": sheet.cell(row, 8).value,
            "time": sheet.cell(row, 9).value,
            "rounds": sheet.cell(row, 10).value,
            "rpe": safe_float(sheet.cell(row, 11).value),
            "notes": sheet.cell(row, 12).value,
        })

    return rows


def unique_preserve_order(values):
    result = []

    for value in values:
        if value is None:
            continue

        text = str(value).strip()

        if text and text not in result:
            result.append(text)

    return result


def show_strength_analytics(workbook):
    rows = get_strength_rows(workbook)

    print("\n" + "=" * 80)
    print("📈 СИЛОВА АНАЛІТИКА")
    print("=" * 80)

    if not rows:
        print("\nДаних ще немає.")
        return

    exercises = unique_preserve_order(
        row["exercise"] for row in rows
    )

    print(f"\nВправ у фактичній історії: {len(exercises)}")
    print(f"Всього записаних підходів: {len(rows)}")

    for exercise in exercises:
        exercise_rows = [
            row for row in rows
            if normalize(row["exercise"]) == normalize(exercise)
            and row["weight"] is not None
            and row["reps"] is not None
        ]

        if not exercise_rows:
            continue

        max_weight = max(
            row["weight"] for row in exercise_rows
        )

        best_volume = max(
            row["weight"] * row["reps"]
            for row in exercise_rows
        )

        latest_workout_id = max(
            (
                safe_float(row["workout_id"])
                for row in exercise_rows
                if safe_float(row["workout_id"]) is not None
            ),
            default=None
        )

        latest_rows = [
            row for row in exercise_rows
            if safe_float(row["workout_id"]) == latest_workout_id
        ]

        latest_weight = max(
            (row["weight"] for row in latest_rows),
            default=None
        )

        latest_reps = sum(
            row["reps"]
            for row in latest_rows
            if row["reps"] is not None
        )

        print("\n" + "-" * 80)
        print(f"🏋️ {exercise}")
        print(f"   Максимальна вага: {max_weight:g} kg")
        print(f"   Найкращий set-volume: {best_volume:g}")
        print(f"   Остання вага:      {latest_weight:g} kg")
        print(f"   Остання сума reps: {latest_reps:g}")

        if latest_weight == max_weight:
            print("   📌 Поточна сесія містить максимальну вагу.")
        else:
            print(
                f"   📊 До максимальної ваги: "
                f"{max_weight:g} kg"
            )


def show_session_analytics(workbook):
    workout_rows = get_workout_rows(workbook)

    print("\n" + "=" * 80)
    print("📊 АНАЛІТИКА ТРЕНУВАНЬ")
    print("=" * 80)

    if not workout_rows:
        print("\nВиконаних тренувань ще немає.")
        return

    valid_rpe = [
        row["rpe"]
        for row in workout_rows
        if row["rpe"] is not None
    ]

    valid_duration = [
        row["duration"]
        for row in workout_rows
        if row["duration"] is not None
    ]

    sessions = {}

    for row in workout_rows:
        session = str(row["session"]).strip()

        if session not in sessions:
            sessions[session] = 0

        sessions[session] += 1

    print(f"\nВсього тренувань: {len(workout_rows)}")

    if valid_rpe:
        average_rpe = sum(valid_rpe) / len(valid_rpe)
        print(f"Середній RPE: {average_rpe:.1f}")

    if valid_duration:
        average_duration = (
            sum(valid_duration) / len(valid_duration)
        )
        print(
            f"Середня тривалість: "
            f"{average_duration:.1f} хв"
        )

    print("\nЗа типами тренувань:")

    for session, count in sessions.items():
        print(f"   {session}: {count}")

    prs = [
        row for row in workout_rows
        if normalize(row["pr"]) in ("так", "yes", "true", "1")
    ]

    print(f"\n🏆 PR-сесій: {len(prs)}")


def show_wod_analytics(workbook):
    rows = get_wod_result_rows(workbook)

    print("\n" + "=" * 80)
    print("🔥 WOD АНАЛІТИКА")
    print("=" * 80)

    if not rows:
        print("\nРезультатів WOD ще немає.")
        return

    wod_ids = unique_preserve_order(
        row["wod_id"] for row in rows
    )

    print(f"\nWOD із результатами: {len(wod_ids)}")
    print(f"Всього результатів: {len(rows)}")

    rpe_values = [
        row["rpe"]
        for row in rows
        if row["rpe"] is not None
    ]

    if rpe_values:
        print(
            f"Середній WOD RPE: "
            f"{sum(rpe_values) / len(rpe_values):.1f}"
        )

    print("\nРезультати:")

    for wod_id in wod_ids:
        wod_rows = [
            row for row in rows
            if str(row["wod_id"]).strip() == str(wod_id).strip()
        ]

        latest = wod_rows[-1]

        print(
            f"\n   WOD {wod_id} | "
            f"Session: {latest['session']} | "
            f"Week: {latest['week']}"
        )
        print(
            f"      Time: {value_or_dash(latest['time'])}"
        )
        print(
            f"      Rounds/Reps/Calories: "
            f"{value_or_dash(latest['rounds'])}"
        )
        print(
            f"      Scaling: "
            f"{value_or_dash(latest['scaling'])}"
        )
        print(
            f"      RPE: {value_or_dash(latest['rpe'])}"
        )


def show_cycle_analytics(workbook, week):
    strength_rows = get_strength_rows(workbook)
    workout_rows = get_workout_rows(workbook)
    wod_rows = get_wod_result_rows(workbook)

    print("\n" + "=" * 80)
    print(f"🧠 АНАЛІТИКА WEEK {week}")
    print("=" * 80)

    week_strength = [
        row for row in strength_rows
        if str(row["week"]).strip() == str(week)
    ]

    week_workouts = [
        row for row in workout_rows
        if str(row["week"]).strip() == str(week)
    ]

    week_wods = [
        row for row in wod_rows
        if str(row["week"]).strip() == str(week)
    ]

    info = get_week_info(workbook, week)

    if info:
        print(
            f"\nФаза: {value_or_dash(info['phase'])}"
        )
        print(
            f"Target RIR: {value_or_dash(info['target_rir'])}"
        )
        print(
            f"Правила: {value_or_dash(info['instructions'])}"
        )

    print(f"\nТренувань за тиждень: {len(week_workouts)}")
    print(f"Силових підходів: {len(week_strength)}")
    print(f"WOD результатів: {len(week_wods)}")

    if week_workouts:
        rpes = [
            row["rpe"]
            for row in week_workouts
            if row["rpe"] is not None
        ]

        if rpes:
            print(
                f"Середній RPE тижня: "
                f"{sum(rpes) / len(rpes):.1f}"
            )

    if week_strength:
        rir_values = [
            row["rir"]
            for row in week_strength
            if row["rir"] is not None
        ]

        if rir_values:
            print(
                f"Середній фактичний RIR: "
                f"{sum(rir_values) / len(rir_values):.1f}"
            )


def show_analytics_menu(workbook, current_week):
    while True:
        print("\n" + "=" * 70)
        print("📊 АНАЛІТИКА")
        print("=" * 70)

        print("\n1 — Силова аналітика")
        print("2 — Аналітика тренувань")
        print("3 — WOD аналітика")
        print("4 — Аналітика поточного тижня")
        print("5 — Назад")

        choice = input("\nОбери: ").strip()

        if choice == "1":
            show_strength_analytics(workbook)
            input("\nНатисни Enter...")

        elif choice == "2":
            show_session_analytics(workbook)
            input("\nНатисни Enter...")

        elif choice == "3":
            show_wod_analytics(workbook)
            input("\nНатисни Enter...")

        elif choice == "5":
            show_cycle_analytics(
                workbook,
                current_week
            )
            input("\nНатисни Enter...")

        elif choice == "6":
            break

        else:
            print("\n❌ Невірний вибір.")


# ============================================================
# WOD HISTORY
# ============================================================

def show_wod_history(workbook):
    if "WOD_RESULTS" not in workbook.sheetnames:
        print("\n❌ WOD_RESULTS не знайдено.")
        return

    sheet = workbook["WOD_RESULTS"]

    print("\n" + "=" * 90)
    print("🔥 ІСТОРІЯ WOD")
    print("=" * 90)

    if sheet.max_row <= 1:
        print("\nРезультатів WOD поки немає.")
        return

    for row in range(2, sheet.max_row + 1):
        print(
            f"\n#{sheet.cell(row, 1).value} | "
            f"{sheet.cell(row, 3).value} | "
            f"Week {sheet.cell(row, 4).value} | "
            f"{sheet.cell(row, 5).value}"
        )

        print(
            f"   WOD ID: {sheet.cell(row, 6).value} | "
            f"Format: {sheet.cell(row, 7).value}"
        )

        print(
            f"   Weight/Scaling: {value_or_dash(sheet.cell(row, 8).value)}"
        )

        print(
            f"   Time: {value_or_dash(sheet.cell(row, 9).value)} | "
            f"Rounds/Reps/Calories: {value_or_dash(sheet.cell(row, 10).value)} | "
            f"RPE: {value_or_dash(sheet.cell(row, 11).value)}"
        )


# ============================================================
# PROGRAM OVERVIEW
# ============================================================

def show_program_overview(workbook):
    print("\n" + "=" * 70)
    print("📋 ПРОГРАМА")
    print("=" * 70)

    weeks = get_available_weeks(workbook)

    print(f"\nТижнів: {len(weeks)}")
    print("Тижні:", ", ".join(map(str, weeks)))

    if "WOD_PLAN" in workbook.sheetnames:
        wod_sheet = workbook["WOD_PLAN"]
        print(f"WOD у базі: {max(0, wod_sheet.max_row - 1)}")

    if "STRENGTH_PLAN" in workbook.sheetnames:
        strength_sheet = workbook["STRENGTH_PLAN"]
        print(
            f"Планових силових рядків: "
            f"{max(0, strength_sheet.max_row - 1)}"
        )

    if "EXERCISES" in workbook.sheetnames:
        exercise_sheet = workbook["EXERCISES"]
        print(
            f"Вправ у базі: "
            f"{max(0, exercise_sheet.max_row - 1)}"
        )


# ============================================================
# CYCLE MANAGEMENT — v2.5
# ============================================================

def get_cycle_stats(workbook, week):
    """Статистика фактичних тренувань за конкретний тиждень."""
    workout_rows = get_workout_rows(workbook)

    week_rows = [
        row for row in workout_rows
        if str(row["week"]).strip() == str(week)
    ]

    sessions = {
        "LOWER": 0,
        "PUSH": 0,
        "PULL": 0,
        "FULL BODY": 0
    }

    for row in week_rows:
        session = str(row["session"]).strip().upper()
        if session in sessions:
            sessions[session] += 1

    return {"total": len(week_rows), "sessions": sessions}


def get_all_cycle_stats(workbook):
    return {
        week: get_cycle_stats(workbook, week)
        for week in get_available_weeks(workbook)
    }


def get_cycle_progress_percent(workbook, current_week):
    weeks = get_available_weeks(workbook)
    if not weeks:
        return 0

    minimum = min(weeks)
    maximum = max(weeks)

    if maximum == minimum:
        return 100

    return max(
        0,
        min(
            100,
            ((current_week - minimum) / (maximum - minimum)) * 100
        )
    )


def get_next_week(workbook, current_week):
    for week in get_available_weeks(workbook):
        if week > current_week:
            return week
    return None


def show_cycle_dashboard(workbook, current_week):
    info = get_week_info(workbook, current_week)
    stats = get_cycle_stats(workbook, current_week)
    progress = get_cycle_progress_percent(workbook, current_week)
    next_week = get_next_week(workbook, current_week)

    print("\n" + "=" * 80)
    print("🧠 TRAINING CYCLE DASHBOARD")
    print("=" * 80)

    print(f"\nПоточний тиждень: {current_week} / 8")
    print(f"Фаза:        {value_or_dash(info['phase'] if info else None)}")
    print(f"Target RIR:  {value_or_dash(info['target_rir'] if info else None)}")
    print(f"Правила:     {value_or_dash(info['instructions'] if info else None)}")

    print(f"\nПрогрес циклу: {progress:.0f}%")

    bar_length = 30
    filled = int(bar_length * progress / 100)
    print("[" + "█" * filled + "░" * (bar_length - filled) + "]")

    print("\n📊 ПОТОЧНИЙ ТИЖДЕНЬ")
    print(f"Тренувань: {stats['total']}")

    for session, count in stats["sessions"].items():
        print(f"   {session}: {count}")

    if next_week is not None:
        next_info = get_week_info(workbook, next_week)
        print(f"\n➡️ Наступний тиждень: {next_week}")

        if next_info:
            print(f"   Фаза: {value_or_dash(next_info['phase'])}")
            print(f"   Target RIR: {value_or_dash(next_info['target_rir'])}")
    else:
        print("\n🏁 Це останній тиждень циклу.")
        print("Після Deload можна готувати новий цикл.")


def show_full_cycle_overview(workbook, current_week):
    stats = get_all_cycle_stats(workbook)

    print("\n" + "=" * 90)
    print("📅 ПОВНИЙ 8-ТИЖНЕВИЙ ЦИКЛ")
    print("=" * 90)

    for week in get_available_weeks(workbook):
        info = get_week_info(workbook, week)
        week_stats = stats.get(week, {"total": 0})

        marker = "👉" if week == current_week else "  "

        print(f"\n{marker} WEEK {week}")
        print(f"   Phase: {value_or_dash(info['phase'] if info else None)}")
        print(f"   Target RIR: {value_or_dash(info['target_rir'] if info else None)}")
        print(f"   Виконано тренувань: {week_stats['total']}")
        print(f"   Rule: {value_or_dash(info['instructions'] if info else None)}")


def change_cycle_week(workbook, current_week):
    weeks = get_available_weeks(workbook)

    print("\n" + "=" * 70)
    print("🔄 ЗМІНА ТИЖНЯ")
    print("=" * 70)
    print(f"\nПоточний тиждень: {current_week}")

    for week in weeks:
        info = get_week_info(workbook, week)
        print(
            f"   {week} — "
            f"{value_or_dash(info['phase'] if info else None)}"
        )

    value = input(
        "\nВведи новий тиждень "
        "(Enter — залишити поточний): "
    ).strip()

    if value == "":
        return current_week

    try:
        selected = int(value)
    except ValueError:
        print("\n❌ Введи номер тижня.")
        return current_week

    if selected not in weeks:
        print("\n❌ Такого тижня немає.")
        return current_week

    if selected > current_week:
        print(
            f"\n⚠️ Перехід з Week {current_week} "
            f"на Week {selected}."
        )

        confirm = input("Підтвердити? (так/ні): ").strip().casefold()

        if confirm not in ("так", "yes", "y"):
            print("\nПерехід скасовано.")
            return current_week

    print(f"\n✅ Поточний тиждень: {selected}")
    return selected


# ============================================================
# ATHLETE DASHBOARD — v2.6
# ============================================================

def get_pr_count(workbook):
    workouts = get_workout_rows(workbook)

    return sum(
        1
        for row in workouts
        if normalize(row["pr"]) in ("так", "yes", "true", "1")
    )


def get_average(values):
    values = [value for value in values if value is not None]

    if not values:
        return None

    return sum(values) / len(values)


def get_latest_workout(workbook):
    workouts = get_workout_rows(workbook)

    if not workouts:
        return None

    valid = []

    for row in workouts:
        workout_id = safe_float(row["id"])

        if workout_id is not None:
            valid.append((workout_id, row))

    if not valid:
        return workouts[-1]

    return max(valid, key=lambda item: item[0])[1]


def get_latest_week_with_data(workbook):
    workouts = get_workout_rows(workbook)

    weeks = []

    for row in workouts:
        try:
            weeks.append(int(row["week"]))
        except (TypeError, ValueError):
            pass

    return max(weeks) if weeks else None


def get_stalled_exercises(workbook, minimum_sessions=2):
    """
    Просте виявлення вправ, де вага не зростала
    в останніх сесіях.
    Це індикатор для уваги, а не автоматичний висновок.
    """
    rows = get_strength_rows(workbook)

    exercises = unique_preserve_order(
        row["exercise"] for row in rows
    )

    stalled = []

    for exercise in exercises:
        exercise_rows = [
            row for row in rows
            if normalize(row["exercise"]) == normalize(exercise)
            and row["weight"] is not None
        ]

        workout_ids = sorted(
            set(
                safe_float(row["workout_id"])
                for row in exercise_rows
                if safe_float(row["workout_id"]) is not None
            )
        )

        if len(workout_ids) < minimum_sessions:
            continue

        last_ids = workout_ids[-minimum_sessions:]

        last_weights = []

        for workout_id in last_ids:
            session_weights = [
                row["weight"]
                for row in exercise_rows
                if safe_float(row["workout_id"]) == workout_id
            ]

            if session_weights:
                last_weights.append(max(session_weights))

        if len(last_weights) < minimum_sessions:
            continue

        if len(set(last_weights)) == 1:
            stalled.append({
                "exercise": exercise,
                "weight": last_weights[-1],
                "sessions": minimum_sessions
            })

    return stalled


def show_athlete_dashboard(workbook, current_week):
    workouts = get_workout_rows(workbook)
    strength = get_strength_rows(workbook)
    wods = get_wod_result_rows(workbook)

    info = get_week_info(
        workbook,
        current_week
    )

    latest = get_latest_workout(workbook)

    workout_rpes = [
        row["rpe"]
        for row in workouts
        if row["rpe"] is not None
    ]

    strength_rirs = [
        row["rir"]
        for row in strength
        if row["rir"] is not None
    ]

    wod_rpes = [
        row["rpe"]
        for row in wods
        if row["rpe"] is not None
    ]

    print("\n" + "=" * 90)
    print("🏆 ATHLETE DASHBOARD")
    print("=" * 90)

    print("\n🧠 ЦИКЛ")
    print(
        f"   Week: {current_week} / 8"
    )
    print(
        f"   Phase: "
        f"{value_or_dash(info['phase'] if info else None)}"
    )
    print(
        f"   Target RIR: "
        f"{value_or_dash(info['target_rir'] if info else None)}"
    )

    print("\n📅 АКТИВНІСТЬ")
    print(
        f"   Всього тренувань: {len(workouts)}"
    )
    print(
        f"   Силових підходів: {len(strength)}"
    )
    print(
        f"   WOD результатів: {len(wods)}"
    )
    print(
        f"   🏆 PR: {get_pr_count(workbook)}"
    )

    if latest:
        print("\n🕒 ОСТАННЄ ТРЕНУВАННЯ")
        print(
            f"   #{latest['id']} | "
            f"{latest['date']} | "
            f"{latest['session']}"
        )
        print(
            f"   Week: {latest['week']} | "
            f"WOD: {value_or_dash(latest['wod_id'])}"
        )

    print("\n📊 СЕРЕДНІ ПОКАЗНИКИ")

    average_rpe = get_average(workout_rpes)
    average_rir = get_average(strength_rirs)
    average_wod_rpe = get_average(wod_rpes)

    print(
        f"   Середній RPE тренувань: "
        f"{average_rpe:.1f}" if average_rpe is not None
        else "   Середній RPE тренувань: —"
    )

    print(
        f"   Середній фактичний RIR: "
        f"{average_rir:.1f}" if average_rir is not None
        else "   Середній фактичний RIR: —"
    )

    print(
        f"   Середній WOD RPE: "
        f"{average_wod_rpe:.1f}" if average_wod_rpe is not None
        else "   Середній WOD RPE: —"
    )

    print("\n⚠️ ПОТРЕБУЄ УВАГИ")

    stalled = get_stalled_exercises(workbook)

    if stalled:
        for item in stalled:
            print(
                f"   🟡 {item['exercise']}: "
                f"{item['weight']:g} kg без росту "
                f"останні {item['sessions']} сесії"
            )
    else:
        print(
            "   Поки що недостатньо даних "
            "для виявлення зупинок."
        )

    print("\n🎯 НАСТУПНИЙ КРОК")

    if info:
        phase = normalize(info["phase"])

        if phase == "deload":
            print(
                "   🔵 Deload: знизити обсяг та інтенсивність "
                "відповідно до плану."
            )
        elif phase == "peak":
            print(
                "   🔴 Peak: пріоритет — якість виконання "
                "та контроль інтенсивності."
            )
        elif phase in ("progression", "hard progression"):
            print(
                "   🟢 Progression: використовувати "
                "рекомендації прогресії."
            )
        else:
            print(
                "   🟡 Base: накопичувати якісні "
                "тренувальні дані."
            )

    print("\n" + "=" * 90)



# ============================================================
# PROGRESS EVALUATION — v2.7
# ============================================================

def evaluate_exercise_progress(workbook, exercise_name, session):
    """Порівнює дві останні фактичні сесії вправи."""
    rows = get_strength_rows(workbook)

    matching = [
        row for row in rows
        if normalize(row["exercise"]) == normalize(exercise_name)
        and normalize(row["session"]) == normalize(session)
    ]

    workout_ids = sorted({
        safe_float(row["workout_id"])
        for row in matching
        if safe_float(row["workout_id"]) is not None
    })

    if len(workout_ids) < 2:
        return {
            "status": "INSUFFICIENT_DATA",
            "label": "НЕДОСТАТНЬО ДАНИХ",
            "message": "Потрібно щонайменше 2 виконані сесії."
        }

    previous_id = workout_ids[-2]
    latest_id = workout_ids[-1]

    def summarize(workout_id):
        data = [
            row for row in matching
            if safe_float(row["workout_id"]) == workout_id
            and row["weight"] is not None
            and row["reps"] is not None
        ]

        if not data:
            return None

        rir_values = [
            row["rir"] for row in data
            if row["rir"] is not None
        ]

        return {
            "max_weight": max(row["weight"] for row in data),
            "total_reps": sum(row["reps"] for row in data),
            "volume": sum(
                row["weight"] * row["reps"]
                for row in data
            ),
            "avg_rir": (
                sum(rir_values) / len(rir_values)
                if rir_values else None
            )
        }

    previous = summarize(previous_id)
    current = summarize(latest_id)

    if previous is None or current is None:
        return {
            "status": "INCOMPLETE_DATA",
            "label": "НЕПОВНІ ДАНІ",
            "message": "Не вистачає ваги або повторень."
        }

    if (
        current["volume"] > previous["volume"]
        and current["max_weight"] >= previous["max_weight"]
    ):
        status = "PROGRESS"
        label = "ПРОГРЕС"
        message = "Тренувальний обсяг або навантаження зросли."
    elif (
        current["volume"] < previous["volume"]
        and current["max_weight"] < previous["max_weight"]
    ):
        status = "REGRESSION"
        label = "РЕГРЕС"
        message = "Останній результат нижчий за попередній."
    elif (
        current["volume"] == previous["volume"]
        and current["max_weight"] == previous["max_weight"]
    ):
        status = "PLATEAU"
        label = "ПЛАТО"
        message = "Ключові показники залишилися на тому самому рівні."
    else:
        status = "MIXED"
        label = "ЗМІШАНА ДИНАМІКА"
        message = "Частина показників змінилась у різні боки."

    if current["avg_rir"] is not None and current["avg_rir"] <= 0:
        message += " Середній RIR дуже низький — інтенсивність була високою."

    return {
        "status": status,
        "label": label,
        "message": message,
        "previous": previous,
        "current": current,
        "previous_id": previous_id,
        "latest_id": latest_id
    }


def show_progress_evaluation(workbook, current_week, session):
    exercises = get_strength_plan(
        workbook,
        current_week,
        session
    )

    print("\n" + "=" * 80)
    print(f"📈 ОЦІНКА ПРОГРЕСУ — {session} — WEEK {current_week}")
    print("=" * 80)

    if not exercises:
        print("\nСилову частину не знайдено.")
        return

    labels = {
        "PROGRESS": "🟢 ПРОГРЕС",
        "PLATEAU": "🟡 ПЛАТО",
        "REGRESSION": "🔴 РЕГРЕС",
        "MIXED": "🔵 ЗМІШАНА ДИНАМІКА",
        "INSUFFICIENT_DATA": "⚪ НЕДОСТАТНЬО ДАНИХ",
        "INCOMPLETE_DATA": "⚠️ НЕПОВНІ ДАНІ"
    }

    for exercise in exercises:
        name = str(exercise["exercise"]).strip()

        result = evaluate_exercise_progress(
            workbook,
            name,
            session
        )

        print("\n" + "-" * 80)
        print(f"🏋️ {name}")
        print(f"   {labels.get(result['status'], result['label'])}")
        print(f"   {result['message']}")

        previous = result.get("previous")
        current = result.get("current")

        if previous and current:
            print(
                f"\n   Попередня сесія #{result['previous_id']}: "
                f"{previous['max_weight']:g} kg | "
                f"{previous['total_reps']:g} reps | "
                f"Volume {previous['volume']:g}"
            )
            print(
                f"   Остання сесія #{result['latest_id']}: "
                f"{current['max_weight']:g} kg | "
                f"{current['total_reps']:g} reps | "
                f"Volume {current['volume']:g}"
            )

            if current["avg_rir"] is not None:
                print(
                    f"   Останній середній RIR: "
                    f"{current['avg_rir']:.1f}"
                )


def show_overall_progress(workbook):
    rows = get_strength_rows(workbook)

    exercises = unique_preserve_order(
        row["exercise"] for row in rows
    )

    print("\n" + "=" * 80)
    print("🏆 ЗАГАЛЬНА ОЦІНКА ПРОГРЕСУ")
    print("=" * 80)

    if not exercises:
        print("\nНедостатньо даних.")
        return

    counters = {
        "PROGRESS": 0,
        "PLATEAU": 0,
        "REGRESSION": 0,
        "MIXED": 0,
        "INSUFFICIENT_DATA": 0,
        "INCOMPLETE_DATA": 0
    }

    for exercise in exercises:
        sessions = unique_preserve_order(
            row["session"]
            for row in rows
            if normalize(row["exercise"]) == normalize(exercise)
        )

        for session in sessions:
            result = evaluate_exercise_progress(
                workbook,
                exercise,
                session
            )

            counters[result["status"]] += 1

            if result["status"] in (
                "PROGRESS",
                "PLATEAU",
                "REGRESSION",
                "MIXED"
            ):
                print(
                    f"\n   {exercise} [{session}] — "
                    f"{result['label']}"
                )

    print("\n📊 ПІДСУМОК")
    print(f"   🟢 Прогрес: {counters['PROGRESS']}")
    print(f"   🟡 Плато: {counters['PLATEAU']}")
    print(f"   🔴 Регрес: {counters['REGRESSION']}")
    print(f"   🔵 Змішана динаміка: {counters['MIXED']}")
    print(f"   ⚪ Недостатньо даних: {counters['INSUFFICIENT_DATA']}")
    print("\n⚠️ Це аналітичний індикатор. План автоматично не змінюється.")


def show_progress_evaluation_menu(workbook, current_week):
    while True:
        print("\n" + "=" * 70)
        print("📈 ОЦІНКА ПРОГРЕСУ")
        print("=" * 70)
        print("\n1 — Загальна оцінка")
        print("2 — LOWER")
        print("3 — PUSH")
        print("4 — PULL")
        print("5 — FULL BODY")
        print("6 — Назад")

        choice = input("\nОбери: ").strip()

        if choice == "1":
            show_overall_progress(workbook)
            input("\nНатисни Enter...")
        elif choice in {"2", "3", "4", "5"}:
            sessions = {
                "2": "LOWER",
                "3": "PUSH",
                "4": "PULL",
                "5": "FULL BODY"
            }
            show_progress_evaluation(
                workbook,
                current_week,
                sessions[choice]
            )
            input("\nНатисни Enter...")
        elif choice == "6":
            return
        else:
            print("\n❌ Невірний вибір.")


# ============================================================
# MAIN MENU
# ============================================================


# ============================================================
# NEXT SESSION RECOMMENDATIONS — v2.8
# ============================================================

def parse_rep_range(rep_text):
    import re

    if rep_text is None:
        return None, None

    text = str(rep_text).replace("–", "-")
    numbers = re.findall(r"\d+(?:\.\d+)?", text)

    if not numbers:
        return None, None

    values = [float(x) for x in numbers]

    if len(values) == 1:
        return values[0], values[0]

    return values[0], values[1]


def get_target_rir_number(value):
    import re

    if value is None:
        return None

    numbers = re.findall(r"\d+(?:\.\d+)?", str(value))

    if not numbers:
        return None

    if len(numbers) > 1:
        return sum(float(x) for x in numbers) / len(numbers)

    return float(numbers[0])


def calculate_next_session_recommendation(
    workbook,
    exercise,
    session,
    current_week
):
    info = get_week_info(workbook, current_week)

    phase = normalize(info["phase"] if info else "")
    target_rir = get_target_rir_number(
        info["target_rir"] if info else None
    )

    result = evaluate_exercise_progress(
        workbook,
        exercise["exercise"],
        session
    )

    if result["status"] in ("INSUFFICIENT_DATA", "INCOMPLETE_DATA"):
        return (
            "ПОЧАТИ ЗА ПЛАНОМ",
            "Недостатньо історії для персональної рекомендації."
        )

    current = result["current"]
    rep_min, rep_max = parse_rep_range(exercise["reps"])
    avg_rir = current["avg_rir"]

    if phase == "deload":
        return (
            "DELOAD",
            "Не підвищувати навантаження. Виконувати Deload за планом."
        )

    if avg_rir is not None and avg_rir < 1:
        return (
            "ЗАЛИШИТИ / КОНТРОЛЮВАТИ",
            "Фактичний RIR дуже низький. Не підвищувати вагу."
        )

    if result["status"] == "REGRESSION":
        return (
            "СТАБІЛІЗУВАТИ",
            "Повторити попереднє навантаження та сфокусуватися на якості."
        )

    if phase == "peak":
        return (
            "ВИКОНАТИ ЗА ПЛАНОМ",
            "Peak: пріоритет — якість виконання та контроль інтенсивності."
        )

    # Перевіряємо, чи всі фактичні сети останньої сесії
    # досягли верхньої межі діапазону.
    strength_rows = get_strength_rows(workbook)

    latest_sets = [
        row for row in strength_rows
        if normalize(row["exercise"]) == normalize(exercise["exercise"])
        and normalize(row["session"]) == normalize(session)
        and safe_float(row["workout_id"]) == result["latest_id"]
        and row["reps"] is not None
    ]

    all_top = (
        bool(latest_sets)
        and rep_max is not None
        and all(row["reps"] >= rep_max for row in latest_sets)
    )

    rir_allows = (
        target_rir is None
        or avg_rir is None
        or avg_rir >= target_rir - 0.5
    )

    if (
        result["status"] == "PROGRESS"
        and all_top
        and rir_allows
    ):
        return (
            "МОЖНА РОЗГЛЯНУТИ ПІДВИЩЕННЯ",
            "Усі сети досягли верхньої межі при прийнятній інтенсивності."
        )

    if result["status"] == "PLATEAU":
        return (
            "ЗАЛИШИТИ ВАГУ",
            "Спробувати покращити повторення без зміни ваги."
        )

    return (
        "ЗАЛИШИТИ ВАГУ",
        "Продовжити роботу з поточним навантаженням і покращувати виконання."
    )


def show_next_session_recommendations(
    workbook,
    current_week,
    session
):
    exercises = get_strength_plan(
        workbook,
        current_week,
        session
    )

    print("\n" + "=" * 85)
    print(f"🎯 РЕКОМЕНДАЦІЇ — {session} — WEEK {current_week}")
    print("=" * 85)

    if not exercises:
        print("\nСилову частину не знайдено.")
        return

    for exercise in exercises:
        label, message = calculate_next_session_recommendation(
            workbook,
            exercise,
            session,
            current_week
        )

        print("\n" + "-" * 85)
        print(f"🏋️ {exercise['exercise']}")
        print(f"   План: {exercise['sets']} × {exercise['reps']}")
        print(f"   🧠 {label}")
        print(f"   {message}")


def show_next_session_menu(workbook, current_week):
    sessions = {
        "1": "LOWER",
        "2": "PUSH",
        "3": "PULL",
        "4": "FULL BODY"
    }

    print("\n" + "=" * 70)
    print("🎯 РЕКОМЕНДАЦІЇ НА НАСТУПНЕ ТРЕНУВАННЯ")
    print("=" * 70)

    print("\n1 — LOWER")
    print("2 — PUSH")
    print("3 — PULL")
    print("4 — FULL BODY")
    print("5 — Назад")

    choice = input("\nОбери: ").strip()

    if choice in sessions:
        show_next_session_recommendations(
            workbook,
            current_week,
            sessions[choice]
        )
        input("\nНатисни Enter...")


# ============================================================
# INTEGRATED WORKOUT SESSION — v2.9
# ============================================================

def ensure_result_sheets(workbook):
    """Створює аркуші результатів, якщо їх ще немає."""
    created = False

    if "STRENGTH_RESULTS" not in workbook.sheetnames:
        ws = workbook.create_sheet("STRENGTH_RESULTS")
        ws.append([
            "ID",
            "Date",
            "Workout ID",
            "Week",
            "Session",
            "Exercise",
            "Set",
            "Weight (kg)",
            "Reps",
            "RIR",
            "Notes"
        ])
        created = True

    if "WOD_RESULTS" not in workbook.sheetnames:
        ws = workbook.create_sheet("WOD_RESULTS")
        ws.append([
            "ID",
            "Date",
            "Workout ID",
            "Week",
            "Session",
            "WOD ID",
            "Result",
            "RPE",
            "Notes"
        ])
        created = True

    return created


def next_numeric_id(rows, key):
    values = []

    for row in rows:
        value = safe_float(row.get(key))

        if value is not None:
            values.append(int(value))

    return max(values, default=0) + 1


def append_strength_result(
    workbook,
    date_value,
    workout_id,
    week,
    session,
    exercise,
    set_number,
    weight,
    reps,
    rir,
    notes
):
    ws = workbook["STRENGTH_RESULTS"]

    rows = get_strength_rows(workbook)

    result_id = next_numeric_id(
        rows,
        "id"
    )

    ws.append([
        result_id,
        date_value,
        workout_id,
        week,
        session,
        exercise,
        set_number,
        weight,
        reps,
        rir,
        notes
    ])


def append_wod_result(
    workbook,
    date_value,
    workout_id,
    week,
    session,
    wod_id,
    result,
    rpe,
    notes
):
    ws = workbook["WOD_RESULTS"]

    rows = get_wod_result_rows(workbook)

    result_id = next_numeric_id(
        rows,
        "id"
    )

    ws.append([
        result_id,
        date_value,
        workout_id,
        week,
        session,
        wod_id,
        result,
        rpe,
        notes
    ])


def get_workout_result_sheet(workbook):
    if "WORKOUTS" not in workbook.sheetnames:
        return None

    return workbook["WORKOUTS"]


def next_workout_id(workbook):
    rows = get_workout_rows(workbook)

    return next_numeric_id(
        rows,
        "id"
    )


def append_workout_record(
    workbook,
    workout_id,
    date_value,
    week,
    session,
    wod_id,
    wod_format,
    duration,
    result,
    rpe,
    pr,
    notes
):
    ws = get_workout_result_sheet(workbook)

    if ws is None:
        raise RuntimeError(
            "Аркуш WORKOUTS не знайдено."
        )

    ws.append([
        workout_id,
        date_value,
        week,
        session,
        "STRENGTH + WOD",
        wod_id,
        wod_format,
        duration,
        result,
        rpe,
        pr,
        notes
    ])


def ask_float(prompt, allow_blank=True):
    while True:
        value = input(prompt).strip().replace(",", ".")

        if value == "" and allow_blank:
            return None

        try:
            return float(value)
        except ValueError:
            print("❌ Введи число.")


def ask_int(prompt, allow_blank=True):
    while True:
        value = input(prompt).strip()

        if value == "" and allow_blank:
            return None

        try:
            return int(value)
        except ValueError:
            print("❌ Введи ціле число.")


def start_integrated_workout(workbook, current_week, session):
    """
    v2.9 integrated entry point.

    Reuses the existing, tested workout engine:
    plan -> exercise briefing/recommendation -> strength results
    -> WOD result -> workout summary -> one save.
    """
    return start_workout(
        workbook,
        current_week,
        session
    )


def main():
    workbook = open_training_database()

    if workbook is None:
        input("\nНатисни Enter для виходу...")
        return

    required_sheets = [
        "WORKOUTS",
        "STRENGTH_PLAN",
        "WOD_PLAN",
        "EXERCISES",
        "WOD_COMPONENTS",
        "WEEKLY_PLAN",
        "STRENGTH_RESULTS",
        "WOD_RESULTS",
    ]

    missing = [
        sheet
        for sheet in required_sheets
        if sheet not in workbook.sheetnames
    ]

    if missing:
        print("\n❌ У файлі відсутні аркуші:")
        for sheet in missing:
            print(f"   - {sheet}")

        input("\nНатисни Enter для виходу...")
        return

    available_weeks = get_available_weeks(workbook)
    current_week = available_weeks[0] if available_weeks else 1

    while True:
        print("\n" + "=" * 70)
        print("🔥 TRAINING SYSTEM v2.9 FIXED")
        print("=" * 70)

        info = get_week_info(workbook, current_week)

        if info:
            print(
                f"\nWEEK {current_week} | "
                f"{value_or_dash(info['phase'])} | "
                f"Target RIR: {value_or_dash(info['target_rir'])}"
            )

        print("\n1 — Почати тренування")
        print("2 — 🏆 Dashboard спортсмена")
        print("3 — 🧠 Dashboard циклу")
        print("4 — Змінити поточний тиждень")
        print("5 — Переглянути план")
        print("6 — Історія тренувань")
        print("7 — Історія силової вправи")
        print("8 — Історія WOD")
        print("9 — Аналіз прогресії")
        print("10 — 📈 Оцінка прогресу")
        print("11 — 🎯 Рекомендації на наступне тренування")
        print("12 — 📊 Аналітика")
        print("13 — Огляд 8-тижневого циклу")
        print("14 — Вихід")

        choice = input("\nОбери дію: ").strip()

        if choice == "1":
            sessions = {
                "1": "LOWER",
                "2": "PUSH",
                "3": "PULL",
                "4": "FULL BODY"
            }

            print("\nОберіть тип тренування:")
            print("1 — LOWER")
            print("2 — PUSH")
            print("3 — PULL")
            print("4 — FULL BODY")

            session = sessions.get(
                input("\nОбери: ").strip()
            )

            if session:
                start_integrated_workout(
                    workbook,
                    current_week,
                    session
                )
            else:
                print("\n❌ Невірний вибір.")

        elif choice == "2":
            show_athlete_dashboard(
                workbook,
                current_week
            )
            input("\nНатисни Enter...")

        elif choice == "3":
            show_cycle_dashboard(
                workbook,
                current_week
            )
            input("\nНатисни Enter...")

        elif choice == "4":
            current_week = change_cycle_week(
                workbook,
                current_week
            )

        elif choice == "5":
            sessions = {
                "1": "LOWER",
                "2": "PUSH",
                "3": "PULL",
                "4": "FULL BODY"
            }

            print("\nОберіть тип тренування:")
            print("1 — LOWER")
            print("2 — PUSH")
            print("3 — PULL")
            print("4 — FULL BODY")

            session = sessions.get(input("\nОбери: ").strip())

            if session:
                show_week_info(
                    workbook,
                    current_week
                )
                show_strength_plan(
                    workbook,
                    current_week,
                    session
                )
                show_wod_plan(
                    workbook,
                    current_week,
                    session
                )
                input("\nНатисни Enter...")
            else:
                print("\n❌ Невірний вибір.")

        elif choice == "6":
            show_history(workbook)
            input("\nНатисни Enter...")

        elif choice == "7":
            show_exercise_history(workbook)
            input("\nНатисни Enter...")

        elif choice == "8":
            show_wod_history(workbook)
            input("\nНатисни Enter...")

        elif choice == "9":
            sessions = {
                "1": "LOWER",
                "2": "PUSH",
                "3": "PULL",
                "4": "FULL BODY"
            }

            print("\nОберіть тип тренування:")
            print("1 — LOWER")
            print("2 — PUSH")
            print("3 — PULL")
            print("4 — FULL BODY")

            session = sessions.get(input("\nОбери: ").strip())

            if session:
                show_all_progression(
                    workbook,
                    current_week,
                    session
                )
                input("\nНатисни Enter...")
            else:
                print("\n❌ Невірний вибір.")

        elif choice == "10":
            show_progress_evaluation_menu(
                workbook,
                current_week
            )

        elif choice == "11":
            show_next_session_menu(
                workbook,
                current_week
            )

        elif choice == "12":
            show_analytics_menu(
                workbook,
                current_week
            )

        elif choice == "13":
            show_full_cycle_overview(
                workbook,
                current_week
            )
            input("\nНатисни Enter...")

        elif choice == "14":
            print("\n👋 Training System завершено.")
            break

        else:
            print("\n❌ Невірний вибір.")



# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()
