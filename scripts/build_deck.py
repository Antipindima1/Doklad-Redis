"""Сборка презентации «Redis: кэширование и паттерны кэширования» (12 слайдов)."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn
from lxml import etree

# ---------- палитра ----------
NAVY = RGBColor(0x0E, 0x1D, 0x30)
INK = RGBColor(0x16, 0x28, 0x3D)
RED = RGBColor(0xC6, 0x30, 0x2A)
RED_SOFT = RGBColor(0xF7, 0xE7, 0xE5)
LIGHT = RGBColor(0xF2, 0xF5, 0xF9)
LINE = RGBColor(0xD7, 0xDF, 0xE9)
GRAY = RGBColor(0x5F, 0x6E, 0x80)
MUTED = RGBColor(0x8A, 0x99, 0xAB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x1F, 0x7A, 0x4D)
GREEN_SOFT = RGBColor(0xE6, 0xF2, 0xEB)

FONT = "Arial"

W, H = 13.333, 7.5
M = 0.7  # левое поле
CW = W - 2 * M  # 11.933


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    return prs


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _set_bullet(p, char="•", indent=0.24, color=None):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Emu(int(Inches(indent)))))
    pPr.set("indent", str(Emu(int(-Inches(indent)))))
    for tag in ("a:buNone", "a:buChar", "a:buFont", "a:buClr"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    if color is not None:
        buClr = etree.SubElement(pPr, qn("a:buClr"))
        srgb = etree.SubElement(buClr, qn("a:srgbClr"))
        srgb.set("val", str(color))
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", FONT)
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", char)


def _no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    for tag in ("a:buChar", "a:buFont", "a:buClr"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    if pPr.find(qn("a:buNone")) is None:
        etree.SubElement(pPr, qn("a:buNone"))


def tb(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, margins=(0, 0, 0, 0)):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(margins[0])
    tf.margin_right = Inches(margins[1])
    tf.margin_top = Inches(margins[2])
    tf.margin_bottom = Inches(margins[3])
    tf.vertical_anchor = anchor
    return tf


def para(tf, parts, size=16, color=INK, bold=False, first=False, align=PP_ALIGN.LEFT,
         space_before=0, space_after=6, line=1.0, bullet=False, bullet_color=None,
         italic=False, indent=0.24):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line
    if bullet:
        _set_bullet(p, color=bullet_color or RED, indent=indent)
    else:
        _no_bullet(p)
    if isinstance(parts, str):
        parts = [(parts, bold, color)]
    for text, b, c in parts:
        r = p.add_run()
        r.text = text
        f = r.font
        f.name = FONT
        f.size = Pt(size)
        f.bold = b
        f.italic = italic
        f.color.rgb = c
    return p


def card(slide, x, y, w, h, fill=LIGHT, line=None, radius=0.06, shadow=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1.0)
    shp.shadow.inherit = shadow
    shp.text_frame.word_wrap = True
    return shp


def rect(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1.0)
    shp.shadow.inherit = False
    return shp


def line_h(slide, x, y, w, color=LINE, weight=1.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln


def arrow(slide, x1, y1, x2, y2, color=GRAY, weight=1.5):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    lnEl = ln.line._get_or_add_ln()
    tail = etree.SubElement(lnEl, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")
    return ln


def page(prs, title, kicker=None):
    slide = blank(prs)
    bg(slide, WHITE)
    if kicker:
        tf = tb(slide, M, 0.38, CW, 0.28)
        para(tf, kicker.upper(), size=10.5, color=MUTED, bold=True, first=True, space_after=0)
    tf = tb(slide, M, 0.62, CW, 0.7)
    para(tf, title, size=30, color=NAVY, bold=True, first=True, space_after=0)
    rect(slide, M, 1.36, 1.35, 0.055, RED)
    # футер
    ftf = tb(slide, M, 7.03, 8.0, 0.28)
    para(ftf, "Redis · проект «Система каршеринга»", size=9.5, color=MUTED, first=True, space_after=0)
    ntf = tb(slide, W - M - 1.0, 7.03, 1.0, 0.28)
    para(ntf, f"{len(prs.slides):02d}", size=9.5, color=MUTED, first=True, space_after=0, align=PP_ALIGN.RIGHT)
    return slide


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# =====================================================================
prs = new_deck()

# ---------------------------------------------------------------- 1
s = blank(prs)
bg(s, NAVY)
rect(s, 0, 0, W, 0.09, RED)
tf = tb(s, 0.95, 0.95, 8.0, 0.3)
para(tf, "ВНУТРЕННИЙ ТЕХНИЧЕСКИЙ МИТАП", size=11, color=RGBColor(0x9A, 0xAE, 0xC6), bold=True, first=True, space_after=0)

tf = tb(s, 0.95, 1.55, 9.2, 1.5)
para(tf, "Redis", size=64, color=WHITE, bold=True, first=True, space_after=2)
para(tf, "кэширование и паттерны кэширования", size=27, color=WHITE, bold=True, space_after=0, line=1.05)

rect(s, 0.95, 3.55, 1.5, 0.06, RED)

tf = tb(s, 0.95, 3.85, 9.4, 0.9)
para(tf, "Практика применения на проекте «Система каршеринга»", size=19, color=RGBColor(0xC9, 0xD6, 0xE6), first=True, space_after=0, line=1.2)

tf = tb(s, 0.95, 6.35, 11.0, 0.5)
para(tf, [("Антипин Д. А.", True, WHITE), ("   ·   ноябрь 2026   ·   Jira SA-2377", False, RGBColor(0x8C, 0xA0, 0xB8))], size=13, first=True, space_after=0)

# декоративная точка
dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11.55), Inches(5.1), Inches(1.0), Inches(1.0))
dot.fill.solid()
dot.fill.fore_color.rgb = RED
dot.line.fill.background()
dot.shadow.inherit = False
dot.rotation = 0

notes(s, "Приветствие. Тема: Redis и паттерны кэширования. Контекст — учебный проект «Система каршеринга», "
         "в архитектуре которого Redis зафиксирован как слой текущего состояния и геопоиска. "
         "План: задача, почему Redis, суть технологии, практика, результаты, трудности, рекомендации.")

# ---------------------------------------------------------------- 2
s = page(prs, "С какой задачей столкнулась команда", "контекст проекта")

kpis = [("10 000", "автомобилей в парке"),
        ("3 000", "активных поездок одновременно"),
        ("5 000", "активных броней")]
for i, (num, label) in enumerate(kpis):
    x = M + i * (3.85 + 0.2)
    card(s, x, 1.62, 3.85, 1.22, fill=LIGHT, line=LINE)
    tf = tb(s, x + 0.35, 1.78, 3.2, 0.6)
    para(tf, num, size=32, color=RED, bold=True, first=True, space_after=0)
    tf = tb(s, x + 0.35, 2.38, 3.25, 0.35)
    para(tf, label, size=12.5, color=GRAY, first=True, space_after=0)

tf = tb(s, M, 2.95, 10.5, 0.3)
para(tf, "Проектные допущения из BRD (SA-2273). Требование: решение по брони — не более 3 секунд в 95% штатных случаев.",
     size=11.5, color=MUTED, italic=True, first=True, space_after=0)

tf = tb(s, M, 3.45, CW, 2.5)
para(tf, [("Позиции меняются непрерывно. ", True, NAVY),
          ("Поток телеметрии обновляет координаты и статус каждой машины — данные устаревают за секунды.", False, INK)],
     size=15.5, first=True, bullet=True, bullet_color=RED, space_after=10, line=1.12)
para(tf, [("Поиск — геопоиск на каждый запрос. ", True, NAVY),
          ("Клиенту нужны ближайшие свободные автомобили в радиусе, и это чтение, а не запись.", False, INK)],
     size=15.5, bullet=True, bullet_color=RED, space_after=10, line=1.12)
para(tf, [("Прямой геопоиск в PostgreSQL/PostGIS дорог. ", True, NAVY),
          ("Геозапросы конкурируют за ресурсы с транзакциями брони и оплаты, а позиции в базе отстают от потока.", False, INK)],
     size=15.5, bullet=True, bullet_color=RED, space_after=0, line=1.12)

card(s, M, 6.05, CW, 0.72, fill=NAVY)
tf = tb(s, M + 0.4, 6.2, CW - 0.8, 0.45)
para(tf, [("Нужен быстрый слой чтения, который не становится вторым источником истины.", True, WHITE)],
     size=15, first=True, space_after=0)

notes(s, "Профиль нагрузки из BRD: 10 000 авто, 3 000 активных поездок, 5 000 броней. Это проектные допущения "
         "учебного сценария, не статистика компании. Три проблемы: частота обновления телеметрии, геопоиск на каждое "
         "чтение и стоимость запросов к PostGIS. Ключевая формулировка задачи — быстрый слой, но не источник истины.")

# ---------------------------------------------------------------- 3
s = page(prs, "Почему выбрали Redis", "обоснование решения")

tf = tb(s, M, 1.6, 4.6, 0.3)
para(tf, "ТРЕБОВАНИЕ", size=10.5, color=MUTED, bold=True, first=True, space_after=0)
tf = tb(s, 6.55, 1.6, 5.9, 0.3)
para(tf, "ЧТО ДАЁТ REDIS", size=10.5, color=MUTED, bold=True, first=True, space_after=0)

rows = [("Задержка чтения", "данные в RAM — миллисекунды без похода на диск"),
        ("Поиск по радиусу", "гео-команды GEOADD / GEOSEARCH «из коробки»"),
        ("Общий слой для инстансов", "внешнее хранилище: все экземпляры Backend видят одно состояние"),
        ("Автоматическое устаревание", "TTL встроен, не нужен отдельный фоновый процесс"),
        ("Восстановимость", "производный индекс — пересобирается из потока телеметрии")]
for i, (req, prop) in enumerate(rows):
    y = 1.95 + i * 0.78
    card(s, M, y, 4.7, 0.62, fill=LIGHT, line=LINE)
    t = tb(s, M + 0.28, y + 0.18, 4.2, 0.3)
    para(t, req, size=13, color=NAVY, bold=True, first=True, space_after=0)
    arrow(s, M + 4.78, y + 0.31, 6.45, y + 0.31, color=RED, weight=1.5)
    card(s, 6.55, y, 5.95, 0.62, fill=WHITE, line=LINE)
    t = tb(s, 6.8, y + 0.19, 5.5, 0.3)
    para(t, prop, size=12.5, color=INK, first=True, space_after=0)

tf = tb(s, M, 5.98, CW, 0.9)
para(tf, [("Почему не альтернативы:  ", True, NAVY),
          ("кэш в памяти приложения — рассинхронизация между инстансами и потеря при рестарте;  ", False, GRAY),
          ("PostGIS на каждый запрос — нагрузка на OLTP и латентность;  ", False, GRAY),
          ("Memcached — нет гео-структур.", False, GRAY)],
     size=12, first=True, space_after=0, line=1.25)

notes(s, "В архитектуре SA-2274 зафиксировано решение: Redis + GEO — быстрый слой состояния и геопоиска. "
         "Пять критериев: задержка, гео-команды, общий слой для инстансов, TTL, восстановимость. "
         "Альтернативы отпали по понятным причинам: локальный кэш не масштабируется, PostGIS не держит профиль чтения.")

# ---------------------------------------------------------------- 4
s = page(prs, "Что такое кэширование", "суть технологии")

card(s, M, 1.6, 6.3, 1.45, fill=LIGHT, line=LINE)
tf = tb(s, M + 0.35, 1.78, 5.6, 1.15)
para(tf, [("Кэширование", True, NAVY), (" — хранение данных максимально близко к месту использования.", False, INK)],
     size=14.5, first=True, space_after=6, line=1.15)
para(tf, "Цель — меньше задержки и ниже нагрузка на серверы, БД и другие компоненты системы.",
     size=13, color=GRAY, space_after=0, line=1.15)

tf = tb(s, M, 3.3, 6.3, 0.3)
para(tf, "КАК ЭТО РАБОТАЕТ", size=10.5, color=MUTED, bold=True, first=True, space_after=0)

steps = [("Первый запрос", "Данные из БД попадают в кэш"),
         ("Повторный запрос", "Данные приходят из кэша — быстро и дёшево"),
         ("Промах (miss)", "Запрос в БД и запись в кэш с заданным TTL"),
         ("Попадание (hit)", "Ответ отдаётся без обращения к БД")]
for i, (head, body) in enumerate(steps):
    y = 3.62 + i * 0.75
    rect(s, M, y + 0.03, 0.08, 0.34, RED)
    tf = tb(s, M + 0.28, y, 6.0, 0.62)
    para(tf, [(head + ". ", True, NAVY), (body, False, INK)], size=13.5, first=True, space_after=0, line=1.1)

tf = tb(s, 7.45, 1.62, 5.2, 0.3)
para(tf, "КЛЮЧЕВЫЕ ПОНЯТИЯ", size=10.5, color=MUTED, bold=True, first=True, space_after=0)

terms = [("Cache hit / miss", "данные нашлись в кэше или нет"),
         ("Hit ratio", "доля попаданий — главная метрика кэша"),
         ("TTL", "время жизни записи до автоустаревания"),
         ("Инвалидация", "принудительное удаление или обновление"),
         ("Вытеснение", "удаление старых записей при нехватке памяти")]
for i, (term, desc) in enumerate(terms):
    y = 1.95 + i * 0.98
    card(s, 7.45, y, 5.18, 0.82, fill=WHITE, line=LINE)
    tf = tb(s, 7.7, y + 0.14, 4.7, 0.6)
    para(tf, term, size=13, color=NAVY, bold=True, first=True, space_after=2)
    para(tf, desc, size=11.5, color=GRAY, space_after=0)

notes(s, "Определение и цель кэширования. Классический сценарий: первый запрос заполняет кэш, повторный читает из кэша. "
         "Ключевые понятия: hit/miss, hit ratio, TTL, инвалидация, вытеснение — они понадобятся дальше в кейсе.")

# ---------------------------------------------------------------- 5
s = page(prs, "Где живёт кэш: уровни", "кэш работает на всех уровнях стека")

levels = [("1", "Процессор", "L1, L2, L3 — быстрая память внутри CPU"),
          ("2", "ОС и диск", "страничный кэш: часто запрашиваемые блоки файлов в RAM"),
          ("3", "Браузер", "статические ресурсы через заголовки Cache-Control, ETag"),
          ("4", "CDN и reverse proxy", "Nginx, Cloudflare — готовые копии без похода в бэкенд"),
          ("5", "Приложение · внешний кэш", "Redis, Memcached — решение нашего проекта"),
          ("6", "База данных", "Buffer Pool и кэш результатов SQL-запросов")]
for i, (num, name, desc) in enumerate(levels):
    col, row = i % 2, i // 2
    x = M + col * 6.05
    y = 1.65 + row * 1.62
    highlighted = num == "5"
    card(s, x, y, 5.85, 1.32, fill=RED_SOFT if highlighted else LIGHT,
         line=RED if highlighted else LINE, radius=0.05)
    tf = tb(s, x + 0.32, y + 0.36, 0.5, 0.5)
    para(tf, num, size=22, color=RED if highlighted else MUTED, bold=True, first=True, space_after=0)
    tf = tb(s, x + 0.85, y + 0.2, 4.75, 0.92, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, name, size=14, color=NAVY, bold=True, first=True, space_after=3)
    para(tf, desc, size=11.5, color=GRAY if not highlighted else INK, space_after=0, line=1.1)

tf = tb(s, M, 6.55, CW, 0.35)
para(tf, [("Наше решение — пятый уровень: ", True, NAVY), ("внешний кэш рядом с приложением, а не память процесса.", False, INK)],
     size=13.5, first=True, space_after=0)

notes(s, "Шесть уровней кэширования. Важно понимать, где узкое место. Наш слой — пятый: внешний кэш уровня приложения "
         "(Redis), а не внутренний кэш процесса — именно потому, что инстансов несколько и данные меняются постоянно.")

# ---------------------------------------------------------------- 6
s = page(prs, "Паттерны кэширования", "что именно применили")

blocks = [
    ("Cache-Aside", "Основной паттерн чтения: приложение спрашивает кэш, при промахе идёт в БД и кладёт результат с TTL.",
     "спрашиваем Redis → при miss идём в БД → пишем в Redis с TTL"),
    ("TTL и инвалидация", "TTL страхует от бесконечного устаревания. Дополнительно — ручная и событийная инвалидация, версионирование.",
     "TTL + удаление ключа по событию изменения"),
    ("Кэш как проекция", "Наш случай: кэш обновляет не сервис чтения, а асинхронный обработчик из потока событий.",
     "пишет обработчик телеметрии, читает Backend"),
    ("Вытеснение", "Когда память заканчивается, Redis удаляет записи по политике — это нужно планировать заранее.",
     "политика вытеснения + расчёт объёма RAM"),
]
for i, (title, body, note_) in enumerate(blocks):
    col, row = i % 2, i // 2
    x = M + col * 6.05
    y = 1.62 + row * 2.12
    highlighted = i == 2
    card(s, x, y, 5.85, 1.95, fill=WHITE if not highlighted else RED_SOFT, line=RED if highlighted else LINE)
    tf = tb(s, x + 0.35, y + 0.2, 5.15, 1.6, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, title, size=16, color=NAVY, bold=True, first=True, space_after=6)
    para(tf, body, size=12, color=INK, space_after=8, line=1.15)
    para(tf, note_, size=11, color=RED if highlighted else GRAY, italic=True, space_after=0, line=1.1)

tf = tb(s, M, 6.3, CW, 0.5)
para(tf, [("Почему это важно: ", True, NAVY),
          ("кэш как проекция снимает с сервиса чтения ответственность за наполнение — но добавляет отставание от источника истины.", False, INK)],
     size=13.5, first=True, space_after=0, line=1.15)

notes(s, "Четыре паттерна. Ключевой для проекта — не классический Cache-Aside на стороне сервиса чтения, "
         "а кэш как проекция: наполняет обработчик телеметрии, читает Backend. Плата — отставание кэша от базы, "
         "и именно поэтому перед бронью нужна проверка в PostgreSQL.")

# ---------------------------------------------------------------- 7
s = page(prs, "Как применили на практике", "схема решения")

# полоса 1
tf = tb(s, M, 1.55, 1.6, 0.3)
para(tf, "1 · ИСТОЧНИК", size=9.5, color=MUTED, bold=True, first=True, space_after=0)
card(s, 2.35, 1.5, 10.28, 0.62, fill=LIGHT, line=LINE)
tf = tb(s, 2.6, 1.65, 9.9, 0.32)
para(tf, "Телематическая платформа — поток координат, статусов и датчиков", size=12.5, color=INK, first=True, space_after=0)

arrow(s, 7.45, 2.15, 7.45, 2.45, color=RED, weight=1.75)

# полоса 2
tf = tb(s, M, 2.52, 1.6, 0.3)
para(tf, "2 · ОБРАБОТКА", size=9.5, color=MUTED, bold=True, first=True, space_after=0)
card(s, 2.35, 2.47, 3.2, 0.72, fill=LIGHT, line=LINE)
tf = tb(s, 2.58, 2.62, 2.8, 0.45)
para(tf, "Приёмник телеметрии", size=12, color=NAVY, bold=True, first=True, space_after=0)
card(s, 5.95, 2.47, 3.0, 0.72, fill=LIGHT, line=LINE)
tf = tb(s, 6.17, 2.62, 2.6, 0.45)
para(tf, "Kafka · telemetry.raw", size=12, color=NAVY, bold=True, first=True, space_after=0)
card(s, 9.33, 2.47, 3.3, 0.72, fill=LIGHT, line=LINE)
tf = tb(s, 9.56, 2.56, 2.9, 0.55)
para(tf, "Обработчик телеметрии", size=12, color=NAVY, bold=True, first=True, space_after=1)
para(tf, "валидация · дедупликация", size=9.5, color=GRAY, space_after=0)

arrow(s, 7.45, 3.25, 7.45, 3.55, color=RED, weight=1.75)

# полоса 3
tf = tb(s, M, 3.62, 1.6, 0.3)
para(tf, "3 · ХРАНИЛИЩА", size=9.5, color=MUTED, bold=True, first=True, space_after=0)
card(s, 2.35, 3.57, 4.55, 0.95, fill=RED_SOFT, line=RED)
tf = tb(s, 2.6, 3.7, 4.1, 0.75)
para(tf, "Redis · состояние + GEO", size=13, color=RED, bold=True, first=True, space_after=1)
para(tf, "пишет обработчик, читает Backend", size=10.5, color=INK, space_after=0)
card(s, 7.15, 3.57, 2.75, 0.95, fill=LIGHT, line=LINE)
tf = tb(s, 7.38, 3.7, 2.4, 0.75)
para(tf, "TimescaleDB", size=13, color=NAVY, bold=True, first=True, space_after=1)
para(tf, "история телеметрии", size=10.5, color=GRAY, space_after=0)
card(s, 10.13, 3.57, 2.5, 0.95, fill=LIGHT, line=LINE)
tf = tb(s, 10.36, 3.7, 2.15, 0.75)
para(tf, "PostgreSQL", size=13, color=NAVY, bold=True, first=True, space_after=1)
para(tf, "+ PostGIS · истина", size=10.5, color=GRAY, space_after=0)

arrow(s, 7.45, 4.6, 7.45, 4.9, color=RED, weight=1.75)

# полоса 4
tf = tb(s, M, 4.97, 1.6, 0.3)
para(tf, "4 · ЧТЕНИЕ", size=9.5, color=MUTED, bold=True, first=True, space_after=0)
card(s, 2.35, 4.92, 10.28, 0.95, fill=NAVY)
tf = tb(s, 2.65, 5.06, 9.7, 0.7)
para(tf, [("Backend каршеринга:  ", True, WHITE),
          ("1) кандидаты рядом — из Redis GEO   →   2) доступность — транзакционно в PostgreSQL   →   3) бронь и холд", False, RGBColor(0xD5, 0xDF, 0xEC))],
     size=12.5, first=True, space_after=0, line=1.2)

tf = tb(s, M, 6.15, CW, 0.8)
para(tf, [("Ключевое: ", True, NAVY),
          ("Redis отвечает на вопрос «кто рядом», а не «кому можно продать». Решение о брони принимает транзакционная база.", False, INK)],
     size=13.5, first=True, space_after=0, line=1.2)

notes(s, "Поток телеметрии: платформа → приёмник → Kafka telemetry.raw → обработчик → Redis (состояние и GEO) "
         "и TimescaleDB (история). Геопоиск: Backend берёт кандидатов из Redis GEO, затем транзакционно проверяет "
         "доступность в PostgreSQL и только потом создаёт бронь. Redis — производный индекс, не источник истины.")

# ---------------------------------------------------------------- 8
s = page(prs, "Почему это безопасно", "граница ответственности кэша")

items = [("Производный индекс", "Redis GEO можно пересобрать из потока или хранилища — потеря кэша не равна потере данных."),
         ("Проверка перед бронью", "Доступность автомобиля подтверждается транзакционно в PostgreSQL, а не по данным кэша."),
         ("Геозоны — в PostGIS", "При завершении поездки проверяется авторитетная геозона, а не её копия в кэше."),
         ("Идемпотентность", "Обработчик телеметрии дедуплицирует повторы и не портит состояние запоздавшими событиями.")]
for i, (head, body) in enumerate(items):
    col, row = i % 2, i // 2
    x = M + col * 6.05
    y = 1.62 + row * 1.5
    card(s, x, y, 5.85, 1.32, fill=WHITE, line=LINE)
    tf = tb(s, x + 0.35, y + 0.18, 0.4, 0.4)
    para(tf, "✓", size=17, color=GREEN, bold=True, first=True, space_after=0)
    tf = tb(s, x + 0.8, y + 0.17, 4.8, 1.05)
    para(tf, head, size=14, color=NAVY, bold=True, first=True, space_after=3)
    para(tf, body, size=11.5, color=GRAY, space_after=0, line=1.12)

card(s, M, 4.85, CW, 0.85, fill=NAVY)
tf = tb(s, M + 0.4, 5.02, CW - 0.8, 0.55)
para(tf, [("Правило: ", True, RGBColor(0xFF, 0x8B, 0x84)),
          ("кэш ускоряет чтение, но не принимает решения о доступности автомобиля и деньгах.", True, WHITE)],
     size=15, first=True, space_after=0)

tf = tb(s, M, 5.95, CW, 0.6)
para(tf, "Именно эта граница — главный проектный результат: технология выбрана не «для скорости», а под конкретную ответственность.",
     size=12.5, color=GRAY, first=True, space_after=0, line=1.2)

notes(s, "Четыре механизма, которые не дают кэшу стать источником истины: восстановимый индекс, проверка перед бронью, "
         "геозоны в PostGIS, идемпотентный обработчик. Главный вывод — граница ответственности важнее выбора СУБД.")

# ---------------------------------------------------------------- 9
s = page(prs, "Результаты", "что даёт решение")

res = [("Геопоиск разгрузил PostGIS", "поиск кандидатов ушёл в in-memory слой, база занята бронями и оплатами"),
       ("Единое состояние для всех инстансов", "горизонтальное масштабирование Backend без рассинхронизации кэшей"),
       ("Запас по латентности", "под требование «бронь не более 3 секунд в 95% случаев»"),
       ("Восстановимость", "потеря Redis не означает потерю бизнес-данных — индекс пересобирается из потока"),
       ("Чтение масштабируется отдельно от БД", "нагрузку поиска можно наращивать, не трогая PostgreSQL")]
for i, (head, body) in enumerate(res):
    y = 1.62 + i * 0.92
    rect(s, M, y + 0.1, 0.07, 0.55, RED)
    tf = tb(s, M + 0.3, y + 0.02, 8.2, 0.8)
    para(tf, head, size=15, color=NAVY, bold=True, first=True, space_after=2)
    para(tf, body, size=12, color=GRAY, space_after=0)
    line_h(s, M, y + 0.84, 8.4)

card(s, 9.0, 1.62, 3.63, 2.5, fill=LIGHT, line=LINE)
tf = tb(s, 9.3, 1.8, 3.05, 2.15)
para(tf, "Чего нет", size=14, color=NAVY, bold=True, first=True, space_after=8)
para(tf, "Измеримых цифр «до/после»: проект учебный, production-замеров не было.", size=11.5, color=INK, space_after=8, line=1.15)
para(tf, "Hit ratio, p95 латентности поиска и экономия ресурсов БД — следующий шаг, после нагрузочного тестирования.",
     size=11.5, color=GRAY, space_after=0, line=1.15)

card(s, 9.0, 4.35, 3.63, 1.4, fill=WHITE, line=RED)
tf = tb(s, 9.3, 4.52, 3.05, 1.1)
para(tf, "Что честно фиксируем", size=12.5, color=RED, bold=True, first=True, space_after=5)
para(tf, "Эффекты описаны как проектные ожидания, а не как измеренный результат.", size=11, color=GRAY, space_after=0, line=1.15)

notes(s, "Результаты: разгрузка PostGIS, единое состояние для инстансов, запас под SLA брони, восстановимость, "
         "независимое масштабирование чтения. Важно честно сказать: измеримых цифр нет — проект учебный, "
         "замеры hit ratio и p95 — следующий шаг.")

# ---------------------------------------------------------------- 10
s = page(prs, "Трудности и ограничения", "чем пришлось платить")

diffs = [("Консистентность", "кэш отстаёт от источника истины на время обработки события — лечится архитектурно, а не «более быстрым Redis»"),
         ("TTL требует расчёта", "в архитектуре TTL зафиксирован как допущение: слишком большой — устаревшие данные, слишком малый — промахи и нагрузка на БД"),
         ("Потеря кэша и холодный старт", "после сбоя индекс пуст, пока не пересоберётся из потока: нужны прогрев и деградация поиска"),
         ("Память и стоимость", "10 000 авто с непрерывной телеметрией — постоянный объём RAM, а память дороже диска"),
         ("Сложность отладки", "данные живут в двух системах — расхождения между Redis и PostgreSQL нужно уметь диагностировать"),
         ("Нельзя кэшировать всё", "финансовые состояния и статусы платежей как источник истины в кэше держать нельзя")]
for i, (head, body) in enumerate(diffs):
    col, row = i % 2, i // 2
    x = M + col * 6.05
    y = 1.62 + row * 1.65
    card(s, x, y, 5.85, 1.45, fill=WHITE, line=LINE)
    rect(s, x, y + 0.16, 0.07, 1.13, RED)
    tf = tb(s, x + 0.32, y + 0.16, 5.25, 1.15, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, head, size=13.5, color=NAVY, bold=True, first=True, space_after=4)
    para(tf, body, size=11, color=GRAY, space_after=0, line=1.12)

notes(s, "Трудности: консистентность, расчёт TTL, холодный старт после потери кэша, память и стоимость, "
         "сложность отладки распределённого состояния и запрет на кэширование транзакционных данных как истины.")

# ---------------------------------------------------------------- 11
s = page(prs, "Когда Redis стоит брать, а когда — нет", "рекомендации коллегам")

card(s, M, 1.6, 5.85, 3.55, fill=GREEN_SOFT, line=GREEN)
tf = tb(s, M + 0.35, 1.78, 5.15, 3.2)
para(tf, "СТОИТ", size=11, color=GREEN, bold=True, first=True, space_after=10)
for t in ["данные читаются часто, а меняются редко — справочники, тарифы, конфигурации",
          "нужен общий слой состояния для нескольких инстансов — сессии, права",
          "нужен быстрый геопоиск или специализированный индекс",
          "данные можно пересобрать из источника истины",
          "есть требование к отзывчивости, которое не выполняется на диске"]:
    para(tf, t, size=11.5, color=INK, bullet=True, bullet_color=GREEN, space_after=6, line=1.1)

card(s, 6.78, 1.6, 5.85, 3.55, fill=RED_SOFT, line=RED)
tf = tb(s, 7.13, 1.78, 5.15, 3.2)
para(tf, "НЕ СТОИТ", size=11, color=RED, bold=True, first=True, space_after=10)
for t in ["данные — источник истины для денег: балансы, проводки, статусы платежей",
          "нужна гарантия «ровно один раз» и строгая транзакционность",
          "данные регуляторные, и хранение вне основной БД не согласовано",
          "нет ресурсов на расчёт TTL, прогрев и мониторинг — рисков больше, чем пользы"]:
    para(tf, t, size=11.5, color=INK, bullet=True, bullet_color=RED, space_after=6, line=1.1)

card(s, M, 5.35, CW, 1.2, fill=LIGHT, line=LINE)
tf = tb(s, M + 0.35, 5.5, CW - 0.7, 0.95)
para(tf, "ЧЕК-ЛИСТ ПЕРЕД ВНЕДРЕНИЕМ", size=10, color=MUTED, bold=True, first=True, space_after=6)
para(tf, "источник истины и правило проверки  ·  TTL и инвалидация  ·  поведение при потере кэша  ·  метрики (hit ratio, латентность, память)  ·  тесты на расхождение",
     size=11.5, color=INK, space_after=0, line=1.15)

notes(s, "Рекомендации: когда Redis уместен (частые чтения, общий слой, гео, восстановимые данные) "
         "и когда нет (деньги, строгая транзакционность, регуляторика, отсутствие ресурсов на эксплуатацию). "
         "Чек-лист из пяти пунктов — минимальный набор перед внедрением.")

# ---------------------------------------------------------------- 12
s = page(prs, "Демо, выводы и ссылки", "завершение")

card(s, M, 1.6, 7.3, 2.45, fill=NAVY)
tf = tb(s, M + 0.35, 1.78, 6.6, 2.1)
para(tf, "ДЕМО · 2 минуты", size=10, color=RGBColor(0xFF, 0x8B, 0x84), bold=True, first=True, space_after=6)
para(tf, "docker run --rm -p 6379:6379 redis", size=10.5, color=RGBColor(0xD5, 0xDF, 0xEC), space_after=3)
para(tf, "redis-cli GEOADD cars:geo 37.6173 55.7558 car-001", size=10.5, color=RGBColor(0xD5, 0xDF, 0xEC), space_after=3)
para(tf, "redis-cli GEOSEARCH cars:geo FROMMEMBER car-001 BYRADIUS 1 km ASC", size=10.5, color=RGBColor(0xD5, 0xDF, 0xEC), space_after=3)
para(tf, "redis-cli SET vehicle:state:car-001 available EX 30", size=10.5, color=RGBColor(0xD5, 0xDF, 0xEC), space_after=0)

card(s, 8.28, 1.6, 4.35, 2.45, fill=RED_SOFT, line=RED)
tf = tb(s, 8.58, 1.78, 3.75, 2.1)
para(tf, "ЗАГЛУШКА ПОД ВИДЕО", size=10, color=RED, bold=True, first=True, space_after=6)
para(tf, "Сюда подставляется ссылка на запись экрана с демонстрацией GEOADD / GEOSEARCH и TTL.",
     size=11, color=INK, space_after=6, line=1.15)
para(tf, "Видео записывается отдельно, слайд остаётся как инструкция.", size=10, color=GRAY, italic=True, space_after=0, line=1.15)

tf = tb(s, M, 4.3, 5.9, 0.3)
para(tf, "ВЫВОДЫ", size=10.5, color=MUTED, bold=True, first=True, space_after=0)
concl = ["Redis выбран под конкретную задачу — real-time состояние и геопоиск, а не «кэш ради кэша».",
         "Главное решение — граница ответственности: скорость в Redis, истина в PostgreSQL.",
         "Риски лежат в консистентности: TTL, инвалидация и перепроверка важнее выбора СУБД.",
         "Внедрение требует подготовки: прогрев, восстановление, метрики, деградация."]
tf = tb(s, M, 4.6, 6.0, 2.2)
for i, c in enumerate(concl):
    para(tf, c, size=12, color=INK, bullet=True, bullet_color=RED, first=(i == 0), space_after=7, line=1.1)

tf = tb(s, 7.15, 4.3, 5.4, 0.3)
para(tf, "ССЫЛКИ", size=10.5, color=MUTED, bold=True, first=True, space_after=0)
links = [("GitHub", "github.com/Antipindima1/Doklad-Redis"),
         ("Jira", "javacode.atlassian.net/browse/SA-2377"),
         ("Confluence", "страница с макросом MacroPack · Markdown"),
         ("Документация", "redis.io · GEO и persistence")]
for i, (name, val) in enumerate(links):
    y = 4.62 + i * 0.55
    tf = tb(s, 7.15, y, 5.4, 0.5)
    para(tf, name, size=11, color=NAVY, bold=True, first=True, space_after=1)
    para(tf, val, size=10, color=GRAY, space_after=0)

notes(s, "Демо: docker run redis, GEOADD/GEOSEARCH, TTL, INFO stats — показываем гео-поиск и попадания. "
         "Выводы: технология под задачу, граница ответственности, риски в консистентности, внедрение требует подготовки. "
         "Ссылки: репозиторий, задача в Jira, страница в Confluence.")

from pathlib import Path
OUT = Path(__file__).resolve().parent.parent / "presentation" / "redis-caching-carsharing.pptx"
prs.save(OUT)
print("saved", len(prs.slides), "slides ->", OUT)
