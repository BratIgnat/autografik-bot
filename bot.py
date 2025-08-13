import os
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from supabase import create_client, Client
from datetime import datetime, timedelta
from aiogram.utils.keyboard import InlineKeyboardBuilder
from dotenv import load_dotenv
from aiogram.types import FSInputFile
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery
from uuid import uuid4
import matplotlib.pyplot as plt

load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
bot = Bot(token=TELEGRAM_TOKEN)
dp = Dispatcher()

class SlotState(StatesGroup):
    waiting_for_date = State()
    waiting_for_slot = State()

# Новые состояния FSM для команд
class JoinTeamState(StatesGroup):
    waiting_for_invite = State()

class CreateTeamState(StatesGroup):
    waiting_for_team_name = State()

def menu_keyboard():
    kb = [
        [KeyboardButton(text="📅 Расписание"), KeyboardButton(text="📝 Моя смена")],
        [KeyboardButton(text="🖼 Поменять фон"), KeyboardButton(text="👥 Пригласить сотрудника")],
        [KeyboardButton(text="👤 Выдать роль"), KeyboardButton(text="❓ Помощь")],
    ]
    return ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True)

def start_keyboard():
    kb = [
        [KeyboardButton(text="➕ Создать команду")],
        [KeyboardButton(text="🔑 Вступить по коду")]
    ]
    return ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True)

def get_active_week(team_id):
    week_resp = supabase.table("weeks").select("*").eq("team_id", team_id).eq("is_active", True).execute()
    if not week_resp.data:
        return None
    return week_resp.data[0]

def get_week_dates(start_date, end_date):
    wdays = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    dates = []
    d0 = datetime.strptime(start_date, "%Y-%m-%d")
    d1 = datetime.strptime(end_date, "%Y-%m-%d")
    for i in range((d1 - d0).days + 1):
        cur = d0 + timedelta(days=i)
        dates.append({
            "weekday": wdays[cur.weekday()],
            "date": cur.strftime("%d.%m"),
            "date_iso": cur.strftime("%Y-%m-%d"),
        })
    return dates

def make_schedule_image(users, week_days, shifts):
    columns = ["ФИО"] + [f"{day['weekday']} {day['date']}" for day in week_days]
    role_map = {
        "employee": "Официанты",
        "barman": "Бармен",
        "host": "Хостес",
        "runner": "Ранеры",
        "admin": "Админы",
        "trainee": "Стажёры"
    }
    roles_order = ["employee", "barman", "host", "runner", "admin", "trainee", "other"]
    data_rows = []
    header_rows = []  # индексы строк-заголовков групп для стилизации

    for role in roles_order:
        if role == "other":
            role_users = [u for u in users if not u.get('role') or u.get('role') not in roles_order[:-1]]
            if not role_users:
                continue
            data_rows.append(["Другие"] + [""] * len(week_days))
            header_rows.append(len(data_rows)-1)
        else:
            role_users = [u for u in users if u.get('role') == role]
            if not role_users:
                continue
            data_rows.append([role_map[role]] + [""] * len(week_days))
            header_rows.append(len(data_rows)-1)
        for u in role_users:
            row = [u["name"]]
            for day in week_days:
                slot = next((s["slot"] for s in shifts if s["user_id"] == u["id"] and s["date"] == day["date_iso"]), "-")
                row.append(slot)
            data_rows.append(row)

    n_cols = len(columns)
    n_rows = len(data_rows)
    fig_w = min(max(2 + n_cols * 1.35, 8), 24)
    fig_h = min(max(1.8 + n_rows * 0.7, 3), 28)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis('off')
    table = ax.table(cellText=data_rows, colLabels=columns, cellLoc='center', loc='center', bbox=[0,0,1,1])
    table.auto_set_font_size(False)
    table.set_fontsize(13)
    table.auto_set_column_width(col=list(range(n_cols)))

     # --- ПАТЧ НАЧИНАЕТСЯ ЗДЕСЬ ---
    # Список ролей для подсветки
    ROLE_HEADERS = {
        'Официанты',
        'Бармен',
        'Хостес',
        'Ранеры',
        'Админы',
        'Стажёры',
        'Другие'
    }

    for (row, col), cell in table.get_celld().items():
        # 1) Шапка таблицы
        if row == 0:
            cell.set_fontsize(14)
            cell.set_text_props(weight="bold")
            cell.set_facecolor("#e3ebfa")

        # 2) Если это первая колонка и значение — роль из списка
        elif col == 0 and row > 0 and data_rows[row - 1][0] in ROLE_HEADERS:
            cell.set_facecolor("#FFD580")
            cell.set_text_props(weight="bold", color="black")

        # 3) Всё остальное — обычный текст и белый фон
        else:
            cell.set_facecolor("white")
            cell.set_text_props(weight="normal", color="black")
    # --- ПАТЧ КОНЧАЕТСЯ ЗДЕСЬ ---

    plt.tight_layout()
    plt.savefig("schedule.png", bbox_inches='tight', transparent=True, dpi=170)
    plt.close(fig)
    return "schedule.png"


@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    user = supabase.table("users").select("team_id").eq("telegram_id", message.from_user.id).execute().data
    # Если пользователя вообще нет в базе — создать пустого пользователя (без команды)
    if not user:
        supabase.table("users").insert({
            "telegram_id": message.from_user.id,
            "name": message.from_user.full_name or message.from_user.username or f"user_{message.from_user.id}",
            "team_id": None,
            "is_owner": False,
            "is_admin": False,
            "role": None
        }).execute()
        user = [{"team_id": None}]
    if user and user[0].get("team_id"):
        await message.answer(
            "Добро пожаловать! Используй меню или команды для работы с ботом.",
            reply_markup=menu_keyboard()
        )
    else:
        await message.answer(
            "Ты не в команде! Создай команду или вступи по коду:",
            reply_markup=start_keyboard()
        )


# Реализация новых пользователей

@dp.message(F.text == "➕ Создать команду")
async def btn_create_team(message: types.Message, state: FSMContext):
    await message.answer("Введи название для своей команды (одной строкой):", reply_markup=ReplyKeyboardRemove())
    await state.set_state(CreateTeamState.waiting_for_team_name)

@dp.message(CreateTeamState.waiting_for_team_name)
async def create_team_name(message: types.Message, state: FSMContext):
    name = message.text.strip()
    invite_code = str(uuid4()).split('-')[0].upper()
    team_id = str(uuid4())
    supabase.table('teams').insert({
        "id": team_id,
        "name": name,
        "invite_code": invite_code,
    }).execute()
    # Создатель — owner/admin
    supabase.table('users').update({
        "team_id": team_id,
        "is_owner": True,
        "is_admin": True
    }).eq('telegram_id', message.from_user.id).execute()
    await message.answer(
        f"Команда <b>{name}</b> создана!\n"
        f"Твой код для приглашения: <code>{invite_code}</code>\n"
        "Ты назначен владельцем и администратором.",
        parse_mode="HTML",
        reply_markup=menu_keyboard()
    )
    await state.clear()

@dp.message(F.text == "🔑 Вступить по коду")
async def btn_join_team(message: types.Message, state: FSMContext):
    await message.answer("Введи код приглашения (invite_code) команды:", reply_markup=ReplyKeyboardRemove())
    await state.set_state(JoinTeamState.waiting_for_invite)

@dp.message(JoinTeamState.waiting_for_invite)
async def join_team_code(message: types.Message, state: FSMContext):
    code = message.text.strip().upper()
    team = supabase.table('teams').select('id', 'name').eq('invite_code', code).execute().data
    if not team:
        await message.answer("Команда с таким кодом не найдена. Проверь правильность кода и попробуй снова.")
        return
    team_id = team[0]['id']
    update_result = supabase.table('users').update({
        "team_id": team_id,
        "is_owner": False,
        "is_admin": False
    }).eq('telegram_id', message.from_user.id).execute()
    if update_result.count == 0:
        supabase.table('users').insert({
            "telegram_id": message.from_user.id,
            "name": message.from_user.full_name or message.from_user.username or f"user_{message.from_user.id}",
            "team_id": team_id,
            "is_owner": False,
            "is_admin": False,
            "role": None
        }).execute()
    await message.answer(
        f"Ты успешно вступил в команду <b>{team[0]['name']}</b>!",
        parse_mode="HTML",
        reply_markup=menu_keyboard()
    )
    await state.clear()


@dp.message(F.text == "📅 Расписание")
async def btn_schedule(message: types.Message, state: FSMContext):
    user_resp = supabase.table("users").select("team_id").eq("telegram_id", message.from_user.id).execute()
    if not user_resp.data or not user_resp.data[0].get("team_id"):
        await message.answer("Ты не состоишь ни в одной команде.")
        return
    team_id = user_resp.data[0]["team_id"]
    week = get_active_week(team_id)
    if not week:
        await message.answer("Нет активной недели. Пусть владелец команды её создаст.")
        return
    week_days = get_week_dates(week["start_date"], week["end_date"])
    users = supabase.table("users").select("id,name,role").eq("team_id", team_id).execute().data
    shifts = supabase.table("shifts").select("*").eq("team_id", team_id).execute().data
    img_path = make_schedule_image(users, week_days, shifts)
    photo = FSInputFile(img_path)
    await message.answer_photo(photo, caption="Текущее расписание:", reply_markup=menu_keyboard())

@dp.message(F.text == "👥 Пригласить сотрудника")
async def btn_invite(message: types.Message, state: FSMContext):
    user = supabase.table("users").select("team_id").eq("telegram_id", message.from_user.id).execute().data
    if not user or not user[0].get("team_id"):
        await message.answer("Ты не состоишь ни в одной команде.")
        return
    team_id = user[0]["team_id"]
    team = supabase.table("teams").select("invite_code").eq("id", team_id).execute().data
    invite_code = team[0]["invite_code"] if team and team[0].get("invite_code") else "Нет кода"
    await message.answer(f"Код приглашения для вашей команды: <code>{invite_code}</code>", parse_mode="HTML")

@dp.message(F.text == "📝 Моя смена")
async def myslot_start(message: types.Message, state: FSMContext):
    user = supabase.table("users").select("team_id").eq("telegram_id", message.from_user.id).execute().data
    if not user or not user[0].get("team_id"):
        await message.answer("Ты не состоишь ни в одной команде.")
        return
    team_id = user[0]["team_id"]
    week = get_active_week(team_id)
    if not week:
        await message.answer("Нет активной недели для выбора смены.")
        return
    week_days = get_week_dates(week["start_date"], week["end_date"])
    kb = [[KeyboardButton(text=f"{day['weekday']} {day['date']}")] for day in week_days]
    await message.answer("Выбери день для смены:", reply_markup=ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True))
    await state.set_state(SlotState.waiting_for_date)
    await state.update_data(week_days=week_days, team_id=team_id)

@dp.message(F.text == "👤 Выдать роль")
async def btn_give_role(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    user = supabase.table('users').select('id, team_id, is_admin, is_owner').eq('telegram_id', user_id).execute().data
    if not user or not (user[0].get('is_admin') or user[0].get('is_owner')):
        await message.answer("Только админ или владелец команды может выдавать роли.")
        return

    team_id = user[0]['team_id']
    members = supabase.table('users').select('id, name, role').eq('team_id', team_id).execute().data
    if not members:
        await message.answer("В команде нет сотрудников.")
        return

    keyboard = InlineKeyboardBuilder()
    for member in members:
        keyboard.button(
            text=f"{member['name']} ({member.get('role', '')})",
            callback_data=f"setrole_{member['id']}"
        )
    await message.answer("Выберите сотрудника для назначения роли:", reply_markup=keyboard.as_markup())

@dp.callback_query(F.data.startswith("setrole_"))
async def callback_choose_role(call: CallbackQuery, state: FSMContext):
    member_id = call.data.replace("setrole_", "")
    await state.update_data(member_id=member_id)
    roles = [
        ("ОФИЦИАНТ", "employee"),
        ("ХОСТ", "host"),
        ("БАРМЕН", "barman"),
        ("РАНЕР", "runner"),
        ("АДМИН", "admin"),
        ("СТАЖЁР", "trainee"),
    ]
    keyboard = InlineKeyboardBuilder()
    for title, code in roles:
        keyboard.button(
            text=title,
            callback_data=f"setroleto_{code}"
        )
    await call.message.edit_text("Выберите новую роль для сотрудника:", reply_markup=keyboard.as_markup())
    await call.answer()

@dp.callback_query(F.data.startswith("setroleto_"))
async def callback_set_role(call: CallbackQuery, state: FSMContext):
    role_code = call.data.replace("setroleto_", "")
    data = await state.get_data()
    member_id = data.get("member_id")
    if not member_id:
        await call.answer("Ошибка: не выбран сотрудник.", show_alert=True)
        return

    supabase.table('users').update({'role': role_code}).eq('id', member_id).execute()
    await call.message.edit_text(f"Роль успешно обновлена!")
    await call.answer("Роль назначена.", show_alert=True)

@dp.message(SlotState.waiting_for_date)
async def slot_choose_day(message: types.Message, state: FSMContext):
    data = await state.get_data()
    week_days = data.get("week_days")
    team_id = data.get("team_id")
    selected = message.text
    day = next((d for d in week_days if f"{d['weekday']} {d['date']}" == selected), None)
    if not day:
        await message.answer("Неверная дата. Попробуй ещё раз.")
        return
    slots = [
        KeyboardButton(text="09:30-23:00"),
        KeyboardButton(text="10:00-23:00"),
        KeyboardButton(text="11:00-23:00"),
        KeyboardButton(text="12:00-23:00"),
        KeyboardButton(text="13:00-23:00"),
        KeyboardButton(text="вых"),
        KeyboardButton(text="17:00-23:00"),
    ]
    await state.update_data(selected_date=day["date_iso"])
    await message.answer("Выбери смену:", reply_markup=ReplyKeyboardMarkup(keyboard=[[s] for s in slots], resize_keyboard=True))
    await state.set_state(SlotState.waiting_for_slot)

@dp.message(SlotState.waiting_for_slot)
async def slot_choose_slot(message: types.Message, state: FSMContext):
    data = await state.get_data()
    slot = message.text
    user = supabase.table("users").select("id,team_id,role").eq("telegram_id", message.from_user.id).execute().data[0]
    date = data["selected_date"]
    team_id = user["team_id"]
    user_id = user["id"]
    role = user["role"]
    existing = supabase.table("shifts").select("id").eq("user_id", user_id).eq("date", date).eq("team_id", team_id).execute().data
    if existing:
        supabase.table("shifts").update({"slot": slot}).eq("id", existing[0]["id"]).execute()
    else:
        supabase.table("shifts").insert({
            "user_id": user_id,
            "team_id": team_id,
            "date": date,
            "slot": slot
        }).execute()
    await message.answer(f"Готово! Ты выбрал смену {slot} на {date}.", reply_markup=menu_keyboard())
    await btn_schedule(message, state)
    await state.clear()

if __name__ == "__main__":
    import asyncio
    import logging
    logging.basicConfig(level=logging.INFO)
    print("Бот стартовал, polling запущен!")
    import sys
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    dp.run_polling(bot)
