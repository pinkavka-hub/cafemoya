#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cafe Moya – generátor webu (směr B: černá a zlatá)

Použití:   python3 build.py            sestaví web do složky dist/ (to dělá i Cloudflare)
           python3 build.py --nahled   navíc vytvoří preview.html (celý web v jednom souboru)

Polední menu se načítá z týdenní tabulky xlsx – stačí ji nahrát na GitHub do hlavní složky.
"""
import datetime
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
DIST = ROOT / "dist"

# ---------------------------------------------------------------- údaje podniku
SITE = "https://www.cafemoya.cz"
ORDER = "https://cafe-moya.order.app.hd.digital/menus"
TEL, TEL_H = "+420605226781", "605 226 781"
TEL2, TEL2_H = "+420702360519", "702 360 519"
EMAIL = "info@cafemoya.cz"
FACEBOOK = "https://www.facebook.com/cafemoyarousinov/"
FIRMY = "https://www.firmy.cz/detail/13414793-cafe-moya-rousinov.html"
INSTAGRAM = "https://www.instagram.com/cafe.moya/"
MAPY = "https://mapy.com/cs/zakladni?source=firm&id=13414793&x=16.882403&y=49.203344&z=17"
GMAPS = ("https://www.google.com/maps/search/?api=1&query="
         + quote("Cafe Moya, Sušilovo náměstí 61, Rousínov"))
# Klíč z web3forms.com (poptávky pak chodí na info@cafemoya.cz). Lze ho zadat i v Cloudflare
# jako proměnnou WEB3FORMS_KEY. Prázdný klíč = formulář jen otevře e-mail s vyplněnými údaji.
WEB3FORMS_KEY = os.environ.get("WEB3FORMS_KEY", "")
LUNCH_END = "14:00"         # po této hodině ukáže úvodní stránka menu na další den, kdy vaříme
TODAY = datetime.date.today()
UPDATED = TODAY.isoformat()
NB = "\u00a0"

VILLAGES = ["Holubice", "Pozořice", "Viničné Šumice", "Vítovice", "Královopolské Vážany",
            "Olšany", "Habrovany", "Nemojany", "Luleč", "Drnovice", "Vyškov"]

# ---------------------------------------------------------------- polední menu z tabulky xlsx
# Týdenní jídelní lístek stačí nahrát na GitHub (nejlépe do hlavní složky, ale najde se kdekoli).
# Čte se stejná tabulka, ze které vznikají příspěvky na Facebook: sloupec B = den, C = pořadí
# jídla, D = název, E = alergeny, F = cena; v hlavičce „47. týden 16.-20.11.2026“.
# Na webu jsou týdny, které ještě neskončily (při nahrání menu na další týden v pátek tedy oba);
# pokud žádný takový není, zůstane poslední známý týden a web upozorní, že nové menu připravujeme.
DAY_NAMES = ["Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek"]
DAY_IDS = {"Pondělí": ("pondeli", "Po"), "Úterý": ("utery", "Út"), "Středa": ("streda", "St"),
           "Čtvrtek": ("ctvrtek", "Čt"), "Pátek": ("patek", "Pá")}
MONTHS_GEN = ["ledna", "února", "března", "dubna", "května", "června", "července", "srpna", "září",
              "října", "listopadu", "prosince"]
LUNCH_PRICE_DEFAULT = (155, 229)    # jen pro texty, kdyby na GitHubu nebyla žádná tabulka s menu
XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
XLSX_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
WEEK_RANGE = re.compile(r"(\d{1,2})\.\s*(?:(\d{1,2})\.\s*)?(?:(\d{4})\s*)?[-–—]\s*"
                        r"(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})")
CLOSED_WORDS = ("nevar", "zavren", "svatek", "dovolen", "nevydav")


def ascii_key(text):
    """Text bez diakritiky malými písmeny: „Středa:“ → „streda:“."""
    return unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode().lower()


def xlsx_rows(path):
    """První list sešitu jako {řádek: {sloupec: hodnota}} – jen se standardní knihovnou Pythonu."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(XLSX_NS + "si"):
                t = si.find(XLSX_NS + "t")
                runs = [t] if t is not None else si.findall(f"{XLSX_NS}r/{XLSX_NS}t")
                shared.append("".join(x.text or "" for x in runs))
        sheet = ET.fromstring(z.read("xl/workbook.xml")).find(f"{XLSX_NS}sheets/{XLSX_NS}sheet")
        rid = sheet.get(XLSX_REL)
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in rels if r.get("Id") == rid)
        target = target.lstrip("/") if target.startswith("/") else "xl/" + target
        data = ET.fromstring(z.read(target))
    rows = {}
    for c in data.iter(XLSX_NS + "c"):
        m = re.fullmatch(r"([A-Z]+)(\d+)", c.get("r") or "")
        if not m:
            continue
        kind, v = c.get("t"), c.find(XLSX_NS + "v")
        if kind == "s" and v is not None:
            value = shared[int(v.text)]
        elif kind == "inlineStr":
            value = "".join(x.text or "" for x in c.iter(XLSX_NS + "t"))
        elif v is not None and v.text is not None:
            value = v.text
            if kind not in ("str", "b", "e"):
                try:
                    value = float(value)
                except ValueError:
                    pass
        else:
            continue
        col = 0
        for ch in m.group(1):
            col = col * 26 + ord(ch) - 64
        rows.setdefault(int(m.group(2)), {})[col] = value
    return rows


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, float):
        v = f"{v:g}"
    text = re.sub(r"\s+", " ", str(v)).strip()
    return re.sub(r" - ", " – ", text)          # pomlčka mezi slovy podle české typografie


def allergens(v):
    """Alergeny jako „1, 3, 7“. Excel z „1,3“ někdy udělá číslo 1.3 a ze „7,10“ číslo 7.1."""
    if isinstance(v, float):
        whole, _, frac = f"{v:.6f}".rstrip("0").rstrip(".").partition(".")
        if frac and int(frac) <= int(whole):
            frac += "0"
        return ", ".join(x for x in (whole, frac) if x)
    return ", ".join(re.findall(r"\d+", cell_text(v)))


def price(v):
    if isinstance(v, float):
        return int(round(v))
    m = re.search(r"\d+", cell_text(v).replace(" ", ""))
    return int(m.group()) if m else None


def dish_number(v):
    m = re.fullmatch(r"(\d{1,2})\s*\.?", cell_text(v))
    return int(m.group(1)) if m else None


def day_index(v):
    if not isinstance(v, str):
        return None
    key = re.sub(r"[^a-z]", "", ascii_key(v))
    keys = [k for k, _ in DAY_IDS.values()]
    return keys.index(key) if key in keys else None


def split_dish(text):
    """„Kuřecí steak, restované špecle, omáčka“ → („Kuřecí steak“, „restované špecle, omáčka“)."""
    depth = 0
    for i, ch in enumerate(text):
        depth += (ch == "(") - (ch == ")")
        if ch == "," and depth == 0:
            return text[:i].strip(), text[i + 1:].strip()
    return text, ""


def week_dates(texts, filename):
    """Pondělí, první a poslední den a číslo týdne z hlavičky, případně z názvu souboru."""
    numbers = [int(m.group(1)) for t in texts + [filename]
               for m in [re.search(r"\b(\d{1,2}) ?t ?y?den\b", re.sub(r"[^a-z0-9]+", " ", ascii_key(t)))] if m]
    for t in texts:
        m = WEEK_RANGE.search(t)
        if not m:
            continue
        d1, m1, y1, d2, m2, y2 = m.groups()
        end = datetime.date(int(y2), int(m2), int(d2))
        m1 = int(m1) if m1 else end.month
        start = datetime.date(int(y1) if y1 else end.year - (m1 > end.month), m1, int(d1))
        monday = start - datetime.timedelta(days=start.weekday())
        if numbers and numbers[0] != monday.isocalendar()[1]:
            print(f"VAROVÁNÍ: {filename}: v hlavičce je {numbers[0]}. týden, ale datum {start:%d. %m. %Y} "
                  f"patří do {monday.isocalendar()[1]}. týdne. Použito datum.")
        return monday, start, end
    if numbers:     # bez data: číslo týdne v roce, který je nejblíž dnešku
        options = []
        for year in (TODAY.year - 1, TODAY.year, TODAY.year + 1):
            try:
                options.append(datetime.date.fromisocalendar(year, numbers[0], 1))
            except ValueError:
                pass
        if options:
            monday = min(options, key=lambda d: abs((d - TODAY).days))
            return monday, monday, monday + datetime.timedelta(days=4)
    raise ValueError("v tabulce chybí datum týdne (např. „47. týden 16.-20.11.2026“)")


def week_label(start, end):
    if (start.year, start.month) == (end.year, end.month):
        return f"{start.day}.–{end.day}.{NB}{MONTHS_GEN[end.month - 1]} {end.year}"
    first = f"{start.day}.{NB}{MONTHS_GEN[start.month - 1]}" + (f" {start.year}" if start.year != end.year else "")
    return f"{first} – {end.day}.{NB}{MONTHS_GEN[end.month - 1]} {end.year}"


def read_week(path):
    rows = xlsx_rows(path)
    texts = [cell_text(v) for r in sorted(rows) for v in rows[r].values() if isinstance(v, str)]
    monday, start, end = week_dates(texts, path.name)
    found, day, base = {}, None, 2
    for r in sorted(rows):
        row = rows[r]
        day_col = next((c for c in sorted(row) if day_index(row[c]) is not None), None)
        if day_col is not None:
            base = day_col
            idx = day_index(row[day_col])
            if idx in found:
                print(f"VAROVÁNÍ: {path.name}: den {DAY_NAMES[idx]} je v tabulce dvakrát, beru první výskyt.")
                day = {"texts": [], "dishes": []}
            else:
                day = found[idx] = {"texts": [], "dishes": []}
        elif day is None:
            continue
        text = cell_text(row.get(base + 2))
        if not text:
            continue
        num = None if day_col is not None else dish_number(row.get(base + 1))
        if num is None:
            if not day["dishes"]:       # polévka nebo „nevaříme“; poznámky pod jídly se nezobrazují
                day["texts"].append((text, allergens(row.get(base + 3))))
        else:
            name, garnish = split_dish(text)
            day["dishes"].append((name, garnish, allergens(row.get(base + 3)), price(row.get(base + 4))))
    days = []
    for idx in sorted(found):
        f = found[idx]
        closed = None
        if not f["dishes"] and f["texts"] and any(w in ascii_key(f["texts"][0][0]) for w in CLOSED_WORDS):
            closed = f["texts"][0][0]
        days.append({"date": (monday + datetime.timedelta(days=idx)).isoformat(), "name": DAY_NAMES[idx],
                     "week": monday.isocalendar()[1], "closed": closed,
                     "soup": None if closed or not f["texts"] else f["texts"][0],
                     "dishes": f["dishes"], "complete": bool(closed or f["dishes"])})
    if not days:
        raise ValueError("v tabulce nejsou žádné dny (sloupec B: Pondělí, Úterý…)")
    return {"week": monday.isocalendar()[1], "monday": monday.isoformat(),
            "end": (monday + datetime.timedelta(days=4)).isoformat(),
            "label": week_label(start, end), "days": days, "source": path.name}


def git_time(path):
    """Kdy byl soubor naposledy nahrán na GitHub (0, když to nejde zjistit)."""
    try:
        out = subprocess.run(["git", "log", "-1", "--format=%ct", "--", str(path)], cwd=ROOT,
                             capture_output=True, text=True, timeout=30).stdout.strip()
        return int(out or 0)
    except (OSError, ValueError, subprocess.SubprocessError):
        return 0


def load_weeks():
    found = {}
    for path in sorted(ROOT.rglob("*.xlsx")):
        rel = path.relative_to(ROOT)
        if rel.parts[0] in ("dist", ".git", "node_modules") or path.name.startswith("~$"):
            continue
        try:
            week = read_week(path)
        except Exception as e:      # vadný soubor nesmí shodit celý web, jen se přeskočí
            print(f"VAROVÁNÍ: menu ze souboru {rel} se nepodařilo načíst: {e}")
            continue
        found.setdefault(week["monday"], []).append((path, week))
    weeks = []
    for monday in sorted(found):
        files = found[monday]
        if len(files) > 1:      # stejný týden víckrát: platí naposledy nahraný, jinak „opravený“
            files.sort(key=lambda x: (git_time(x[0]), "oprav" in ascii_key(x[0].name), x[0].name))
        weeks.append(files[-1][1])
    current = [w for w in weeks if w["end"] >= TODAY.isoformat()]
    return current or weeks[-1:]


WEEKS = load_weeks()

PLATTERS = [("Salámová mísa", [("1 kg", 739), ("1,5 kg", 1109), ("2 kg", 1478)]),
            ("Sýrová mísa", [("1 kg", 769), ("1,5 kg", 1154), ("2 kg", 1538)]),
            ("Mix mísa", [("1 kg", 749), ("1,5 kg", 1124), ("2 kg", 1498)])]

K_MENU = [("Domácí limonáda 0,5 l", "35 Kč"), ("Točená limonáda 0,3 / 0,5 l", "19 / 25 Kč"),
          ("Čepovaná Plzeň 0,4 l", "49 Kč")]

ALLERGENS = ["Obiloviny obsahující lepek", "Korýši", "Vejce", "Ryby", "Arašídy", "Sójové boby", "Mléko",
             "Skořápkové plody", "Celer", "Hořčice", "Sezamová semena", "Oxid siřičitý a siřičitany",
             "Vlčí bob (lupina)", "Měkkýši"]

FAQ_ROZVOZ = [
    ("Kolik stojí rozvoz?",
     "Rozvoz poledního menu je zdarma. Platíte jen za jídlo a polévka je v ceně menu."),
    ("Do kdy musím oběd objednat?",
     "Online objednávky na daný den přijímáme do 9:00, pro Vyškov do 10:30. Později nám prosím zavolejte na 702 360 519."),
    ("Kdy oběd přivezete?",
     "Rozvážíme ve dvou trasách. Na trasu směr Pozořice vyjíždíme v 10:00, na trasu Rousínov – Vyškov v 11:00. "
     "Do Vyškova obvykle dorazíme mezi 12:00 a 12:30. Před příjezdem vám kurýr pošle SMS."),
    ("Můžu si oběd vyzvednout osobně?",
     "Ano. Objednávky s osobním odběrem vydáváme v kavárně na Sušilově náměstí od 11:00 do 14:00."),
    ("Objednáváte pro celou kancelář nebo na celý týden?",
     "Zavolejte nám na 605 226 781 nebo napište na info@cafemoya.cz a domluvíme hromadnou objednávku."),
]

# ---------------------------------------------------------------- ikony
CAMERA = ('<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" '
          'stroke-linecap="round" stroke-linejoin="round"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/>'
          '<circle cx="12" cy="13" r="3.5"/></svg>')
PIN = ('<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" '
       'stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s-7-6.2-7-11.5a7 7 0 0 1 14 0C19 14.8 12 21 12 21z"/>'
       '<circle cx="12" cy="9.5" r="2.5"/></svg>')
PLUS = ('<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
        'stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>')
BURGER = ('<svg aria-hidden="true" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
          'stroke-width="2" stroke-linecap="round"><path d="M4 7h16M4 12h16M4 17h16"/></svg>')

# ---------------------------------------------------------------- pomocníci
def esc(s):
    return html.escape(s, quote=True)


def kc(n):
    return f"{n:,}".replace(",", NB) + NB + "Kč"


def dm(iso):
    _, m, d = iso.split("-")
    return f"{int(d)}.{NB}{int(m)}."


def ph(label):
    return f'<figure class="ph">{CAMERA}<figcaption>Foto: {label}</figcaption></figure>'


def tel_links(text):
    text = text.replace(TEL2_H, f'<a href="tel:{TEL2}">{TEL2_H}</a>')
    text = text.replace(TEL_H, f'<a href="tel:{TEL}">{TEL_H}</a>')
    return text.replace(EMAIL, f'<a href="mailto:{EMAIL}">{EMAIL}</a>')


# Česká typografie: jednopísmenné předložky a spojky nenechat na konci řádku
VLNKA = re.compile(r"(?<!\w)([AaIiKkOoSsUuVvZz])\s+(?=\S|$)")


def vlnka(markup):
    out, skip = [], False
    for part in re.split(r"(<[^>]+>)", markup):
        if part.startswith("<"):
            low = part.lower()
            if low.startswith(("<script", "<style")):
                skip = True
            elif low.startswith(("</script", "</style")):
                skip = False
            out.append(part)
        else:
            out.append(part if skip else VLNKA.sub(lambda m: m.group(1) + NB, part))
    return "".join(out)


LOGO = '<img src="/assets/logo-cafe-moya.svg" alt="" width="72" height="72">'
FOOTER_LOGO = ('<img class="footer-logo" src="/assets/logo-cafe-moya.svg" alt="Cafe Moya" '
               'width="84" height="84" loading="lazy">')


def shot(name, caption, alt, tall=False):
    """Fotka do galerie s popiskem: na šířku 4:3, nebo s tall=True na výšku 3:4."""
    files = ", ".join(f"/assets/foto/{name}-{w}.webp {w}w" for w in (400, 800))
    height, cls = (1067, "shot shot--tall") if tall else (600, "shot")
    return (f'<figure class="{cls}"><img src="/assets/foto/{name}-800.webp" srcset="{files}" '
            f'sizes="(max-width: 860px) calc(100vw - 40px), 390px" width="800" height="{height}" '
            f'alt="{alt}" loading="lazy" decoding="async"><figcaption>{caption}</figcaption></figure>')


def photo(name, alt):
    """Fotka na výšku ve třech velikostech (480, 720 a 960 px široká)."""
    files = ", ".join(f"/assets/foto/{name}-{w}.webp {w}w" for w in (480, 720, 960))
    return (f'<img class="photo" src="/assets/foto/{name}-960.webp" srcset="{files}" '
            'sizes="(max-width: 960px) calc(100vw - 40px), 480px" width="960" height="1286" '
            f'alt="{alt}" loading="lazy" decoding="async">')


def wide(name, alt):
    """Fotka na šířku 4:3 přes celý sloupec dvousloupcového bloku (600, 900 a 1200 px)."""
    files = ", ".join(f"/assets/foto/{name}-{w}.webp {w}w" for w in (600, 900, 1200))
    return (f'<img class="photo-wide" src="/assets/foto/{name}-1200.webp" srcset="{files}" '
            'sizes="(max-width: 960px) calc(100vw - 40px), 564px" width="1200" height="900" '
            f'alt="{alt}" loading="lazy" decoding="async">')


ADDRESS ='<address class="addr">Cafe Moya<br>Sušilovo náměstí 61/39<br>683 01 Rousínov</address>'
HOURS_ROWS = [("Pondělí–pátek", "8:00–19:00"), ("Sobota a neděle", "9:00–19:00"),
              ("Polední menu", "Po–Pá 11:00–14:00")]


def hours(extra=()):
    rows = "".join(f"<div><dt>{a}</dt><dd>{b}</dd></div>" for a, b in list(HOURS_ROWS) + list(extra))
    return f'<dl class="hours">{rows}</dl>'


# Mapa OpenStreetMap: nesleduje návštěvníky, a proto nepotřebuje souhlas s cookies
OSM_EMBED = ("https://www.openstreetmap.org/export/embed.html"
             "?bbox=16.8789%2C49.2017%2C16.8859%2C49.2050&amp;layer=mapnik&amp;marker=49.203344%2C16.882403")
MAP = ('<figure class="map">'
       f'<iframe src="{OSM_EMBED}" title="Mapa: Cafe Moya, Sušilovo náměstí 61/39, Rousínov" '
       'loading="lazy" referrerpolicy="no-referrer"></iframe>'
       f'<figcaption class="map-links"><a href="{MAPY}">Otevřít v Mapy.com</a>'
       f'<a href="{GMAPS}">Otevřít v Google Maps</a></figcaption></figure>')

COLS = ('<div class="cols">'
        '<div class="col"><h3>Káva</h3><p>Espresso, doppio, cappuccino a kávové speciality</p></div>'
        '<div class="col"><h3>Ledové kávy a nealko</h3><p>Ledové kávy, domácí limonády, čaj a horká čokoláda</p></div>'
        '<div class="col"><h3>Víno a koktejly</h3><p>Víno, koktejly, gin &amp; tonic, čepovaná Plzeň a destiláty</p></div>'
        '<div class="col"><h3>Snídaně a sladké</h3><p>Snídaně, panini, palačinky a zmrzlinové poháry</p></div>'
        '</div>')

# ---------------------------------------------------------------- polední menu
INCOMPLETE = '<p class="menu-note">Další jídla doplníme z týdenního jídelního lístku.</p>'


def text_html(s):
    return html.escape(s, quote=False)


def soup_li(soup):
    name, alg = soup
    alg = f' <span class="alg">({alg})</span>' if alg else ""
    return (f'<li><span class="dish"><b>Polévka: {text_html(name[0].lower() + name[1:])}</b>{alg}</span>'
            f'<span class="price price--incl">v ceně menu</span></li>')


def dish_li(i, dish):
    name, garnish, alg, price = dish
    g = f' <span class="garnish">{text_html(garnish)}</span>' if garnish else ""
    alg = f' <span class="alg">({alg})</span>' if alg else ""
    cena = kc(price) if price else ""
    return (f'<li><span class="dish"><span class="num">{i}.</span><b>{text_html(name)}</b>{g}{alg}</span>'
            f'<span class="price">{cena}</span></li>')


def day_items(day):
    soup = soup_li(day["soup"]) if day["soup"] else ""
    return soup + "".join(dish_li(i, x) for i, x in enumerate(day["dishes"], 1))


def day_body(day):
    if day["closed"]:
        return f'<p class="menu-note">{text_html(day["closed"])}</p>'
    return f'<ul class="menu-list">{day_items(day)}</ul>' + ("" if day["complete"] else INCOMPLETE)


def cooking_days():
    return [d for w in WEEKS for d in w["days"] if not d["closed"] and (d["soup"] or d["dishes"])]


def today_panel():
    """Úvodní stránka: bloky všech dnů, kdy vaříme; site.js podle českého času ukáže ten pravý."""
    days = cooking_days()
    ready = [d for d in days if d["complete"]] or days
    first = next((d for d in ready if d["date"] >= UPDATED), ready[0] if ready else None)
    out = []
    for d in days:
        attrs = f' data-day-block data-date="{d["date"]}"' + ("" if d is first else " hidden")
        out.append(f'<div class="menu-day"{attrs}><div class="menu-day-head"><h3>{d["name"]} {dm(d["date"])}</h3>'
                   f'<p data-when>Polední menu, {d["week"]}. týden</p></div>{day_body(d)}</div>')
    hidden = " hidden" if first else ""
    out.append(f'<div class="menu-day" data-day-fallback{hidden}><div class="menu-day-head"><h3>Polední menu</h3></div>'
               '<p class="menu-note">Menu na další dny právě připravujeme. Rádi vám ho řekneme na '
               f'<a href="tel:{TEL}">{TEL_H}</a>.</p></div>')
    return "\n".join(out)


def week_menu(week, suffix=""):
    out = []
    for d in week["days"]:
        sid = DAY_IDS[d["name"]][0] + suffix
        out.append(f'<section class="menu-day" id="{sid}" data-menu-date="{d["date"]}" aria-labelledby="h-{sid}">'
                   f'<div class="menu-day-head"><h3 id="h-{sid}">{d["name"]} {dm(d["date"])}</h3>'
                   f'<span class="today-badge" hidden>Dnes</span></div>{day_body(d)}</section>')
    return '<div class="week">\n' + "\n".join(out) + '\n</div>'


def day_links(week, suffix=""):
    return "".join(f'<a href="#{DAY_IDS[d["name"]][0]}{suffix}" aria-label="{d["name"]}">{DAY_IDS[d["name"]][1]}</a>'
                   for d in week["days"])


def weeks_html():
    """Stránka Polední menu: každý týden ve vlastní sekci; site.js skryje týdny, které už skončily."""
    if not WEEKS:
        return ('<section class="sec sec--paper" aria-label="Polední menu">\n<div class="wrap">\n'
                '<p class="menu-note">Menu na tento týden právě připravujeme. Rádi vám ho řekneme na '
                f'<a href="tel:{TEL}">{TEL_H}</a>.</p>\n</div>\n</section>')
    out = []
    for i, w in enumerate(WEEKS):
        suffix = "" if i == 0 else f"-{w['week']}"
        old = ""
        if i == len(WEEKS) - 1:
            old = ('\n<p class="menu-note" data-menu-old hidden>Menu na další týden právě připravujeme. '
                   f'Níže je menu, které platilo do {dm(w["end"])} Rádi vám nové menu řekneme na '
                   f'<a href="tel:{TEL}">{TEL_H}</a>.</p>')
        out.append(f'''<section class="sec sec--paper" aria-labelledby="h-tyden{suffix}" data-week-end="{w["end"]}">
<div class="wrap">
<div class="week-head">
<h2 id="h-tyden{suffix}">{w["week"]}. týden, {w["label"]}</h2>
<nav class="day-jump" aria-label="Dny {w["week"]}. týdne">{day_links(w, suffix)}</nav>{old}
</div>
{week_menu(w, suffix)}
</div>
</section>''')
    return "\n".join(out)


def k_menu():
    return "".join(f"<li><span>{a}</span><span>{b.replace(' Kč', NB + 'Kč')}</span></li>" for a, b in K_MENU)


def platters():
    cards = []
    for name, sizes in PLATTERS:
        rows = "".join(f"<div><dt>{w}</dt><dd>{kc(p)}</dd></div>" for w, p in sizes)
        cards.append(f'<div class="platter"><h3>{name}</h3><dl>{rows}</dl></div>')
    return '<div class="platters">' + "".join(cards) + "</div>"


def villages():
    return '<ul class="villages">' + "".join(f"<li>{v}</li>" for v in VILLAGES) + "</ul>"


def faq_html(items):
    out = []
    for i, (q, a) in enumerate(items):
        opened = " open" if i == 0 else ""
        out.append(f"<details{opened}><summary>{q}{PLUS}</summary><p>{tel_links(links_html(a))}</p></details>")
    return '<div class="faq">' + "".join(out) + "</div>"


def faq_section(sid, items, title="Časté dotazy"):
    return (f'<section class="sec sec--paper sec--tight" aria-labelledby="{sid}">\n<div class="wrap">\n'
            f'<h2 id="{sid}">{title}</h2>\n{faq_html(items)}\n</div>\n</section>')


# ---------------------------------------------------------------- texty pro vyhledávače a AI
# Jedna věta, která přesně říká, kdo jsme, co děláme a kde. Stejné údaje jsou na stránkách,
# ve strukturovaných datech i v llms.txt, aby vyhledávače i AI asistenti dostávali všude totéž.
BUSINESS_DESC = ("Cafe Moya je rodinná kavárna a cukrárna na Sušilově náměstí 61/39 v Rousínově "
                 "(okres Vyškov). Pečeme dorty na objednávku a zákusky, připravujeme obložené mísy "
                 "a catering a ve všední dny vaříme polední menu s rozvozem zdarma do Rousínova, "
                 "Vyškova a okolních obcí.")

LUNCH_PRICES = [x[3] for d in cooking_days() for x in d["dishes"] if x[3]] or LUNCH_PRICE_DEFAULT
LUNCH_MIN, LUNCH_MAX = min(LUNCH_PRICES), max(LUNCH_PRICES)
PLATTER_MIN = min(p for _, sizes in PLATTERS for w, p in sizes if w == "1 kg")
MD_LINK = re.compile(r"\[([^\]]+)\]\(([^)\s]+)\)")


def links_html(text):
    """Odkazy zapsané jako [text](adresa) převede na HTML odkazy."""
    return MD_LINK.sub(lambda m: f'<a href="{m.group(2)}">{m.group(1)}</a>', text)


def plain(text):
    """Text bez odkazů a pevných mezer, pro strukturovaná data a llms.txt."""
    return MD_LINK.sub(lambda m: m.group(1), text).replace(NB, " ")


def villages_text():
    names = [v for v in VILLAGES if v != "Vyškov"]
    return ", ".join(names[:-1]) + " a " + names[-1]


def lunch_price_text():
    if LUNCH_MIN == LUNCH_MAX:
        return f"stojí hlavní jídla {kc(LUNCH_MIN)}"
    return f"stojí hlavní jídla od {kc(LUNCH_MIN)} do {kc(LUNCH_MAX)}"


def platter_sentence():
    out = []
    for name, sizes in PLATTERS:
        parts = [f"{kc(p)} za {w}" for w, p in sizes]
        out.append(f"{name} stojí " + ", ".join(parts[:-1]) + " a " + parts[-1] + ".")
    return " ".join(out)


# Vybavení a služby kavárny (potvrzené majitelem). Používá se ve výpisu „Dobré vědět“,
# v častých dotazech, ve strukturovaných datech i v llms.txt.
PAYMENT_TEXT = "Platit můžete kartou i stravenkami Edenred, Up (dříve Chèque Déjeuner) a Pluxee (dříve Sodexo)."
PARKING_TEXT = "Parkovat můžete přímo před kavárnou na Sušilově náměstí, hodina parkování je zdarma."
GLUTEN_TEXT = "Vaříme i bezlepková jídla a některá další jídla můžeme připravit bez lepku změnou přílohy."
AMENITIES = [("Platba kartou", " i stravenkami Edenred, Up a Pluxee"), ("Wi-Fi", " pro hosty"),
             ("Bezbariérový přístup", ""), ("Pejsci vítáni", ""),
             ("Parkování přímo před kavárnou", ", hodina zdarma"),
             ("Bezlepková jídla", " a úprava dalších jídel bez lepku změnou přílohy")]
AMENITY_FEATURES = ["Platba kartou", "Wi-Fi", "Bezbariérový přístup", "Psi vítáni",
                    "Parkování před kavárnou, hodina zdarma", "Venkovní zahrádka", "Bezlepková jídla"]


def perks():
    items = "".join(f"<li><b>{name}</b>{rest}</li>" for name, rest in AMENITIES)
    return f'<ul class="perks">{items}</ul>'


FAQ_HOME = [
    ("Kde Cafe Moya najdu?",
     "Na Sušilově náměstí 61/39 v centru Rousínova (okres Vyškov), mezi Brnem a Vyškovem. "
     "Mapu a navigaci najdete na stránce [Kontakt](/kontakt/)."),
    ("Kde u vás zaparkuji?",
     "Přímo před kavárnou na Sušilově náměstí, hodina parkování je zdarma. Do kavárny vede bezbariérový přístup."),
    ("Jakou máte otevírací dobu?",
     "Kavárna je otevřená v pondělí až pátek od 8 do 19 hodin a v sobotu a neděli od 9 do 19 hodin. "
     "Polední menu vydáváme ve všední dny od 11 do 14 hodin."),
    ("Vaříte polední menu s rozvozem?",
     f"Ano. Ve všední dny vaříme polévku a pět hlavních jídel, v aktuálním menu {lunch_price_text()} "
     "a polévka je v ceně. Rozvoz je zdarma po Rousínově, do Vyškova a okolních obcí. "
     "Podrobnosti najdete na stránce [Rozvoz obědů](/rozvoz-obedu/)."),
    ("Upečete dort na objednávku?",
     "Ano. Dorty pečeme na narozeniny, výročí, křtiny i firemní akce podle vašeho přání. Poptávku pošlete "
     "přes stránku [Dorty a zákusky](/dorty-na-objednavku/) nebo zavolejte na 605 226 781."),
    ("Připravujete obložené mísy?",
     f"Ano. Kilogramová mísa stojí od {kc(PLATTER_MIN)}, vybrat si můžete salámovou, sýrovou nebo mix mísu "
     "a kilogram obvykle stačí pro 5–6 osob. Mísy objednávejte nejméně jeden pracovní den předem. "
     "Celý ceník je na stránce [Catering](/catering/)."),
    ("Můžu platit kartou nebo stravenkami?",
     "Ano. " + PAYMENT_TEXT),
]

FAQ_MENU = [
    ("Kolik stojí polední menu?",
     f"V aktuálním menu {lunch_price_text()}. Polévka je v ceně menu."),
    ("Kdy polední menu vydáváte?",
     "Ve všední dny od 11:00 do 14:00 v kavárně na Sušilově náměstí v Rousínově. O víkendu polední menu nevaříme."),
    ("Jak si oběd objednám s rozvozem?",
     f"Online přes [objednávkový systém]({ORDER}), na daný den do 9:00, pro Vyškov do 10:30. "
     "Rozvoz je zdarma. Později nám prosím zavolejte na 702 360 519."),
    ("Máte bezlepková jídla?",
     "Ano. " + GLUTEN_TEXT + " Zeptejte se nás při objednávce."),
    ("Kde najdu alergeny?",
     "Čísla v závorkách u každého jídla označují alergeny. Jejich seznam najdete pod menu."),
]

FAQ_ROZVOZ = [("Kam obědy rozvážíte?",
               f"Po Rousínově, do Vyškova a do obcí {villages_text()}. "
               "Nevíte, jestli vozíme i k vám? Zavolejte na 605 226 781.")] + FAQ_ROZVOZ

FAQ_DORTY = [
    ("Jak si u vás objednám dort?",
     "Vyplňte poptávku níže na této stránce nebo zavolejte na 605 226 781. Napište, kdy dort potřebujete, "
     "pro kolik osob má být a jakou příchuť a zdobení si přejete. Ozveme se vám a potvrdíme cenu i termín."),
    ("Na jaké příležitosti dorty pečete?",
     "Na narozeniny, výročí, křtiny, firemní akce i další oslavy. Každý dort pečeme podle domluvy s vámi."),
    ("Prodáváte zákusky i bez objednávky?",
     "Ano. Ve vitríně v kavárně najdete zákusky a dezerty z naší výroby. Dáte si je u nás ke kávě, "
     "nebo si je odnesete domů."),
]

FAQ_CATERING = [
    ("Kolik stojí obložené mísy?", platter_sentence()),
    ("Pro kolik lidí stačí jedna mísa?", "Kilogram mísy obvykle stačí pro 5–6 osob."),
    ("Jak dlouho předem musím mísu objednat?",
     "Nejméně jeden pracovní den předem, a to i na víkend. Mísy připravujeme z čerstvých surovin."),
    ("Jak objednávka mísy probíhá?",
     "Vyberte mísu a termín ve formuláři níže nebo zavolejte na 605 226 781. Po objednání vám zavoláme, "
     "objednávku zkontrolujeme a doptáme se na podrobnosti."),
    ("Připravíte i chlebíčky a pohoštění na akci?",
     "Ano. Chlebíčky, zákusky z naší cukrárny a další pohoštění připravíme na rodinné oslavy, porady i firemní "
     "akce. Napište nám na info@cafemoya.cz, na kdy a pro kolik hostů pohoštění potřebujete, a připravíme vám nabídku."),
]

FAQ_KAVARNA = [
    ("Co si u vás můžu dát?",
     "Kávu a kávové speciality, ledové kávy, domácí limonády, čaj, horkou čokoládu, víno, koktejly a čepovanou "
     "Plzeň. K jídlu snídaně, panini, palačinky, zmrzlinové poháry a zákusky z naší cukrárny."),
    ("Máte venkovní zahrádku?",
     "Ano. Za hezkého počasí můžete sedět venku na zahrádce pod slunečníky."),
    ("Můžu k vám přijít se psem?",
     "Ano, pejsci jsou u nás vítáni."),
    ("Je kavárna bezbariérová?",
     "Ano, do kavárny vede bezbariérový přístup."),
    ("Máte Wi-Fi?",
     "Ano, pro hosty máme Wi-Fi."),
    ("Můžu si rezervovat stůl?",
     "Ano, přes formulář níže nebo telefonicky na 605 226 781. Rezervaci vám potvrdíme telefonicky."),
]

# ---------------------------------------------------------------- formuláře
def field(fid, label, kind="text", required=False, attrs="", options=None):
    name = fid.split("-", 1)[1]
    star = " *" if required else ""
    req = " required" if required else ""
    if kind == "select":
        opts = "".join(f"<option>{o}</option>" for o in options)
        control = f'<select id="{fid}" name="{name}"{req}{attrs}>{opts}</select>'
    elif kind == "textarea":
        control = f'<textarea id="{fid}" name="{name}"{req}{attrs}></textarea>'
    else:
        control = f'<input id="{fid}" name="{name}" type="{kind}"{req}{attrs}>'
    return f'<div class="field"><label for="{fid}">{label}{star}</label>{control}</div>'


def row(*fields):
    return '<div class="form-row">' + "".join(fields) + "</div>"


def form(subject, parts, submit):
    if WEB3FORMS_KEY:
        endpoint = "https://api.web3forms.com/submit"
        head = (f'<form class="form" data-form data-subject="{subject}" data-endpoint="{endpoint}" '
                f'action="{endpoint}" method="post">'
                f'<input type="hidden" name="access_key" value="{esc(WEB3FORMS_KEY)}">'
                f'<input type="hidden" name="subject" value="{subject}">'
                '<input type="hidden" name="from_name" value="Web Cafe Moya">'
                '<input type="checkbox" name="botcheck" tabindex="-1" autocomplete="off" hidden>')
    else:
        head = (f'<form class="form" data-form data-subject="{subject}" '
                f'action="mailto:{EMAIL}" method="post" enctype="text/plain">')
    return (head + "".join(parts) +
            '<p class="form-note">Položky označené * jsou povinné. Údaje použijeme jen k vyřízení vaší žádosti, '
            'více v <a href="/ochrana-osobnich-udaju/">zásadách ochrany osobních údajů</a>.</p>'
            f'<div><button class="btn btn-ink" type="submit">{submit}</button></div>'
            '<p class="form-msg" role="status"></p></form>')


NAME = ' autocomplete="name"'
PHONE = ' autocomplete="tel"'
MAIL = ' autocomplete="email"'

CAKE_FORM = form("Poptávka dortu z webu", [
    row(field("d-jmeno", "Jméno a příjmení", required=True, attrs=NAME),
        field("d-telefon", "Telefon", "tel", True, PHONE)),
    row(field("d-email", "E-mail", "email", attrs=MAIL),
        field("d-datum", "Kdy dort potřebujete", "date", True)),
    row(field("d-porce", "Počet porcí", "number", attrs=' min="1" inputmode="numeric"'),
        field("d-prilezitost", "Příležitost", "select",
              options=["Narozeniny", "Výročí", "Křtiny", "Firemní akce", "Jiná příležitost"])),
    field("d-predstava", "Vaše představa", "textarea", attrs=' placeholder="Příchuť, zdobení, nápis, alergie…"'),
], "Odeslat poptávku")

PLATTER_FORM = form("Objednávka obložené mísy z webu", [
    row(field("m-jmeno", "Jméno a příjmení", required=True, attrs=NAME),
        field("m-telefon", "Telefon", "tel", True, PHONE)),
    row(field("m-misa", "Mísa", "select", True, options=["Salámová", "Sýrová", "Mix (salámy i sýry)"]),
        field("m-velikost", "Velikost", "select", True, options=["1 kg", "1,5 kg", "2 kg"]),
        field("m-pocet", "Počet kusů", "number", True, ' min="1" value="1" inputmode="numeric"')),
    row(field("m-datum", "Na kdy mísu potřebujete", "date", True),
        field("m-cas", "V kolik hodin", "time")),
    field("m-email", "E-mail", "email", attrs=MAIL),
    field("m-poznamka", "Poznámka", "textarea"),
], "Odeslat objednávku")

TABLE_FORM = form("Rezervace stolu z webu", [
    row(field("r-jmeno", "Jméno a příjmení", required=True, attrs=NAME),
        field("r-telefon", "Telefon", "tel", True, PHONE)),
    row(field("r-datum", "Datum", "date", True), field("r-cas", "Čas", "time", True),
        field("r-osob", "Počet osob", "number", True, ' min="1" inputmode="numeric"')),
    field("r-poznamka", "Poznámka", "textarea", attrs=' placeholder="Například oslava nebo místo na zahrádce"'),
], "Odeslat rezervaci")

# ---------------------------------------------------------------- hlavička a patička
NAV = [("/kavarna/", "Kavárna"), ("/dorty-na-objednavku/", "Dorty a zákusky"), ("/catering/", "Catering"),
       ("/obedove-menu/", "Polední menu"), ("/rozvoz-obedu/", "Rozvoz"), ("/kontakt/", "Kontakt")]


def header(active):
    items = []
    for path, label in NAV:
        cur = ' aria-current="page"' if path == active else ""
        items.append(f'<li><a href="{path}"{cur}>{label}</a></li>')
    return ('<a class="skip" href="#obsah">Přeskočit na obsah</a>\n'
            '<header class="site-header"><div class="wrap header-in">'
            f'<a class="logo" href="/" aria-label="Cafe Moya, úvodní stránka">{LOGO}</a>'
            '<button class="nav-toggle" type="button" aria-expanded="false" aria-controls="hlavni-menu">'
            f'{BURGER}<span>Menu</span></button>'
            '<nav class="nav" id="hlavni-menu" aria-label="Hlavní menu"><ul>' + "".join(items) + '</ul>'
            f'<a class="btn btn-gold-line" href="{ORDER}">Objednat oběd</a></nav>'
            '</div></header>')


def footer():
    services = [("/kavarna/", "Kavárna Rousínov"), ("/dorty-na-objednavku/", "Dorty na objednávku"),
                ("/catering/", "Obložené mísy a catering"), ("/obedove-menu/", "Obědové menu Rousínov"),
                ("/rozvoz-obedu/", "Rozvoz obědů Vyškov a okolí"), ("/kontakt/", "Kontakt a otevírací doba")]
    links = "".join(f'<li><a href="{p}">{t}</a></li>' for p, t in services)
    return ('<footer class="site-footer"><div class="wrap"><div class="footer-grid">'
            f'<div class="footer-brand">{FOOTER_LOGO}<p>Rodinná kavárna a cukrárna na Sušilově náměstí v Rousínově. '
            'Catering na oslavy a obědy s rozvozem zdarma.</p></div>'
            f'<nav aria-label="Služby"><h2>Služby</h2><ul>{links}</ul></nav>'
            f'<div><h2>Kontakt</h2><p>Sušilovo náměstí 61/39<br>683 01 Rousínov<br>'
            f'<a href="tel:{TEL}">{TEL_H}</a><br><a href="mailto:{EMAIL}">{EMAIL}</a></p></div>'
            '<div><h2>Otevírací doba</h2><p>Po–Pá 8:00–19:00<br>So–Ne 9:00–19:00</p>'
            f'<h2>Sledujte nás</h2><ul><li><a href="{FACEBOOK}">Facebook</a></li>'
            f'<li><a href="{INSTAGRAM}">Instagram</a></li></ul></div>'
            f'</div><div class="footer-bottom"><p>© {TODAY.year} Cafe Moya</p>'
            '<ul><li><a href="/kontakt/#prace">Práce u nás</a></li>'
            '<li><a href="/ochrana-osobnich-udaju/">Ochrana osobních údajů</a></li></ul>'
            '</div></div></footer>')


def page_hero(crumb, title, lead, buttons=""):
    btns = f'<div class="btns">{buttons}</div>' if buttons else ""
    return ('<section class="page-hero"><div class="wrap">'
            '<nav class="crumbs" aria-label="Drobečková navigace"><ol><li><a href="/">Úvod</a></li>'
            f'<li><span aria-current="page">{crumb}</span></li></ol></nav>'
            f'<h1>{title}</h1><p class="lead">{lead}</p>{btns}</div></section>')

# ---------------------------------------------------------------- stránky
HOME = f'''<section class="hero sec--night">
<div class="wrap hero-in">
<div class="hero-text">
<h1>Kavárna a cukrárna na Sušilově náměstí v Rousínově</h1>
<p class="lead">Rodinná kavárna a cukrárna v centru Rousínova u Vyškova. Pečeme dorty a zákusky, připravujeme obložené mísy a ve všední dny vaříme polední menu s rozvozem zdarma.</p>
<div class="btns"><a class="btn btn-gold" href="/dorty-na-objednavku/">Dort na objednávku</a><a class="btn btn-line" href="/obedove-menu/">Polední menu</a></div>
<p class="hero-note">Otevřeno Po–Pá 8–19 h, So–Ne 9–19 h</p>
</div>
<div class="hero-media"><div class="ring"><img src="/assets/foto/uvod-kava-1000.webp" srcset="/assets/foto/uvod-kava-600.webp 600w, /assets/foto/uvod-kava-1000.webp 1000w" sizes="(max-width: 720px) 270px, 470px" width="1000" height="1000" alt="Cappuccino s latte art na dřevěném baru kavárny" fetchpriority="high"></div></div>
</div>
</section>
<section class="sec sec--paper" aria-labelledby="home-poledne" data-lunch-section="weekday" data-lunch-end="{LUNCH_END}">
<div class="wrap lunch">
<div class="lunch-text">
<h2 id="home-poledne">Polední menu s rozvozem zdarma</h2>
<p class="muted">Ve všední dny od 11 do 14 h vaříme polévku a pět jídel. Polévka je v ceně menu. Rozvážíme po Rousínově, do Vyškova a <a href="/rozvoz-obedu/">dalších obcí v okolí</a>.</p>
<div class="btns"><a class="btn btn-ink" href="{ORDER}">Objednat oběd</a><a class="btn btn-line" href="/obedove-menu/">Menu na celý týden</a></div>
<p class="small muted">Platit můžete kartou i stravenkami Edenred, Up a Pluxee.</p>
</div>
<div class="lunch-menu" data-menu-slot>
{today_panel()}
</div>
</div>
</section>
<section class="sec sec--paper" aria-labelledby="home-vikend" data-lunch-section="weekend" hidden>
<div class="wrap lunch">
<div class="lunch-text">
<h2 id="home-vikend">O víkendu vás čekáme u kávy a dortu</h2>
<p class="muted">V sobotu a v neděli máme otevřeno od 9 do 19 h. Polední menu vaříme zase od pondělí.</p>
<div class="btns"><a class="btn btn-ink" href="/dorty-na-objednavku/">Dort na objednávku</a><a class="btn btn-line" href="/catering/">Nabídka mís</a></div>
<p class="small muted">Obložené mísy na oslavu objednávejte nejméně jeden pracovní den předem, i na víkend.</p>
</div>
<div class="lunch-menu" data-menu-slot>
<img class="weekend-photo" data-weekend-photo src="/assets/foto/dorty-960.webp" srcset="/assets/foto/dorty-480.webp 480w, /assets/foto/dorty-720.webp 720w, /assets/foto/dorty-960.webp 960w" sizes="(max-width: 960px) calc(100vw - 40px), 700px" width="960" height="1286" alt="Dort a zákusky z naší cukrárny" loading="lazy" decoding="async">
</div>
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="home-dorty">
<div class="wrap split">
{photo("dorty", "Čokoládový dort se zdobením a další zákusky na stole")}
<div class="prose">
<h2 id="home-dorty">Dorty a zákusky z naší cukrárny</h2>
<p>Zákusky a dezerty vyrábíme sami. Na objednávku upečeme dort na narozeniny, výročí nebo firemní oslavu podle vašeho přání a počtu hostů.</p>
<p class="muted">Objednávky přijímáme telefonicky, e-mailem nebo přímo v kavárně.</p>
<div class="btns"><a class="btn btn-ink" href="/dorty-na-objednavku/">Poptat dort</a></div>
</div>
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="home-catering">
<div class="wrap split split--rev">
<div class="prose">
<h2 id="home-catering">Obložené mísy a catering</h2>
<p>Salámové, sýrové i kombinované mísy z čerstvých surovin, chlebíčky a pohoštění na oslavy i firemní akce. Kilogram mísy obvykle stačí pro 5–6 osob.</p>
<p class="muted">Mísy od {kc(PLATTER_MIN)} za kilogram. Objednávejte nejméně jeden pracovní den předem.</p>
<div class="btns"><a class="btn btn-ink" href="/catering/">Nabídka mís</a></div>
</div>
{photo("misy", "Tři obložené mísy se salámy, sýry a zeleninou")}
</div>
</section>
<section class="sec sec--night" aria-labelledby="home-kavarna">
<div class="wrap">
<div class="sec-head">
<h2 id="home-kavarna">Kavárna otevřená každý den</h2>
<p class="muted">Ve všední dny od 8 do 19 h, o víkendu od 9. Za hezkého počasí můžete sedět venku na zahrádce.</p>
</div>
{COLS}
<div class="btns sec-foot"><a class="btn btn-gold-line" href="/kavarna/">O kavárně</a></div>
</div>
</section>
{faq_section("home-faq", FAQ_HOME)}
<section class="sec sec--paper" aria-labelledby="home-kontakt">
<div class="wrap contact">
<div class="contact-info prose">
<h2 id="home-kontakt">Přijďte k nám</h2>
{ADDRESS}
{hours()}
<p class="muted">{PARKING_TEXT} Do kavárny vede bezbariérový přístup.</p>
<p><a href="tel:{TEL}"><b>+420 {TEL_H}</b></a><br><a href="mailto:{EMAIL}"><b>{EMAIL}</b></a></p>
<div class="btns"><a class="btn btn-ink" href="{MAPY}">Navigovat k nám</a></div>
</div>
{MAP}
</div>
</section>'''

MENU_PAGE = page_hero(
    "Polední menu", "Obědové menu v Rousínově",
    "Každý všední den vaříme polévku a pět jídel. Výdej Po–Pá 11–14 h, polévka je v ceně menu. "
    "Oběd vám přivezeme zdarma, nebo si ho vyzvedněte u nás.",
    f'<a class="btn btn-gold" href="{ORDER}">Objednat online</a><a class="btn btn-line" href="/rozvoz-obedu/">Jak funguje rozvoz</a>'
) + f'''
{weeks_html()}
<section class="sec sec--paper sec--tight" aria-label="Nápoje, stravenky a objednávky">
<div class="wrap info-row">
<div><h3>K menu</h3><ul class="price-list">{k_menu()}</ul></div>
<div><h3>Platba</h3><p class="muted">{PAYMENT_TEXT}</p></div>
<div><h3>Bez lepku</h3><p class="muted">{GLUTEN_TEXT}</p></div>
<div><h3>Objednávky na rozvoz</h3><p class="muted">Online na daný den do 9:00, pro Vyškov do 10:30. Později volejte na <a href="tel:{TEL2}">{TEL2_H}</a>.</p></div>
</div>
</section>
<section class="sec sec--paper sec--tight" id="alergeny" aria-labelledby="h-alergeny">
<div class="wrap">
<h2 class="h3" id="h-alergeny">Seznam alergenů</h2>
<p class="muted">Čísla v závorkách u jídel označují alergeny.</p>
<ol class="allergens">{"".join(f"<li>{a}</li>" for a in ALLERGENS)}</ol>
</div>
</section>
{faq_section("h-faq-menu", FAQ_MENU)}'''

ROZVOZ_PAGE = page_hero(
    "Rozvoz obědů", "Rozvoz obědů zdarma v Rousínově, Vyškově a okolí",
    "Polední menu vám přivezeme domů i do kanceláře. Rozvoz je zdarma a polévka je v ceně menu.",
    f'<a class="btn btn-gold" href="{ORDER}">Objednat oběd</a><a class="btn btn-line" href="/obedove-menu/">Menu na tento týden</a>'
) + f'''
<section class="sec sec--paper" aria-labelledby="h-jak">
<div class="wrap">
<h2 id="h-jak">Jak to funguje</h2>
<ol class="steps">
<li><span class="step-num" aria-hidden="true">1</span><h3>Vyberte si z menu</h3><p class="muted">Polévka a pět jídel. Menu zveřejňujeme vždy na celý týden dopředu.</p></li>
<li><span class="step-num" aria-hidden="true">2</span><h3>Objednejte online</h3><p class="muted">Na daný den do 9:00, pro Vyškov do 10:30. Později nám zavolejte na <a href="tel:{TEL2}">{TEL2_H}</a>.</p></li>
<li><span class="step-num" aria-hidden="true">3</span><h3>Přivezeme zdarma</h3><p class="muted">Rozvážíme hromadně po trasách. Před příjezdem vám kurýr pošle SMS.</p></li>
</ol>
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-kdy">
<div class="wrap">
<h2 id="h-kdy">Kdy vozíme</h2>
<div class="routes">
<div class="route"><p class="route-time">10:00</p><h3>Trasa směr Pozořice</h3><p class="muted">Vyjíždíme směr Pozořice a zpět.</p></div>
<div class="route"><p class="route-time">11:00</p><h3>Trasa Rousínov – Vyškov</h3><p class="muted">Do Vyškova obvykle dorazíme mezi 12:00 a 12:30.</p></div>
</div>
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-kam">
<div class="wrap">
<h2 id="h-kam">Kam vozíme</h2>
<p class="muted">Rozvážíme po Rousínově a do okolních obcí:</p>
{villages()}
<p class="muted">Vozíme i do dalších obcí v okolí. Nevíte, jestli i k vám? Zavolejte na <a href="tel:{TEL}">{TEL_H}</a>.</p>
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-faq">
<div class="wrap">
<h2 id="h-faq">Časté dotazy</h2>
{faq_html(FAQ_ROZVOZ)}
</div>
</section>
<section class="sec sec--night">
<div class="wrap band-in">
<h2>Objednejte si oběd na dnešek</h2>
<a class="btn btn-gold" href="{ORDER}">Objednat online</a>
</div>
</section>'''

DORTY_PAGE = page_hero(
    "Dorty a zákusky", "Dorty na objednávku a zákusky z naší cukrárny",
    "Zákusky a dezerty vyrábíme sami v naší cukrárně v Rousínově. Na objednávku upečeme dort na narozeniny, "
    "výročí i firemní oslavu.",
    f'<a class="btn btn-gold" href="#poptavka-dortu">Poptat dort</a><a class="btn btn-line" href="tel:{TEL}">Zavolat {TEL_H}</a>'
) + f'''
<section class="sec sec--paper" aria-labelledby="h-dort">
<div class="wrap split">
{wide("dort-na-objednavku", "Dětský dort na objednávku v mořském stylu s mušlemi, korálem a narvalem")}
<div class="prose">
<h2 id="h-dort">Dort podle vašeho přání</h2>
<p>Řekněte nám, pro koho dort bude, kolik hostů čekáte a jaké chutě máte rádi. Domluvíme velikost, příchuť i zdobení a termín, kdy si dort vyzvednete.</p>
<p class="muted">Dort prosím objednávejte alespoň [doplnit: kolik dní] předem.</p>
</div>
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-zakusky">
<div class="wrap split split--rev">
<div class="prose">
<h2 id="h-zakusky">Zákusky z vitríny</h2>
<p>Ve vitríně v kavárně najdete zákusky a dezerty z naší výroby. Dejte si je ke kávě, nebo si je odneste domů.</p>
<p>Na oslavu vám zákusky rádi připravíme i na objednávku.</p>
</div>
{wide("zakusky-vitrina", "Domácí cheesecaky zdobené malinami, borůvkami, čokoládou a preclíky")}
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-ukazky">
<div class="wrap">
<h2 class="h3" id="h-ukazky">Ukázky naší práce</h2>
<div class="gallery">{shot("ukazka-zakusky", "Větrník a další zákusky", "Čokoládový zákusek s pipetou, choux s čokoládovým krémem a malinou a karamelový větrník")}{shot("ukazka-pavlovy", "Pavlovy s ovocem", "Tři pavlovy se šlehačkou, jahodami a borůvkami a šálek espressa")}{shot("ukazka-panna-cotta", "Panna cotta s ovocem", "Panna cotta s červenou a žlutou ovocnou vrstvou ve třech kelímcích")}</div>
</div>
</section>
{faq_section("h-faq-dorty", FAQ_DORTY)}
<section class="sec sec--paper sec--tight" id="poptavka-dortu" aria-labelledby="h-poptavka-dortu">
<div class="wrap form-wrap">
<div class="form-intro prose">
<h2 id="h-poptavka-dortu">Poptávka dortu</h2>
<p class="muted">Vyplňte, co o dortu zatím víte. Ozveme se vám, doladíme podrobnosti a potvrdíme cenu i termín.</p>
<p class="muted">Raději voláte? <a href="tel:{TEL}">{TEL_H}</a></p>
</div>
{CAKE_FORM}
</div>
</section>'''

CATERING_PAGE = page_hero(
    "Catering", "Obložené mísy a catering na oslavy",
    "Plánujete oslavu, posezení s přáteli nebo firemní akci? V Rousínově vám připravíme obložené mísy "
    "z čerstvých surovin, chlebíčky i další pohoštění.",
    f'<a class="btn btn-gold" href="#objednavka-misy">Objednat mísu</a><a class="btn btn-line" href="tel:{TEL}">Zavolat {TEL_H}</a>'
) + f'''
<section class="sec sec--paper" aria-labelledby="h-misy">
<div class="wrap">
<div class="sec-head">
<h2 id="h-misy">Obložené mísy</h2>
<p class="muted">Vybrat si můžete salámovou, sýrovou nebo kombinovanou mísu. Kilogram mísy obvykle stačí pro 5–6 osob. Mísa je i originální dárek.</p>
</div>
{platters()}
<p class="note-box">Mísy připravujeme z čerstvých surovin, proto je prosím objednávejte nejméně 1 pracovní den předem, i na víkend. Po objednání vám zavoláme, objednávku zkontrolujeme a doptáme se na podrobnosti.</p>
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-pohosteni">
<div class="wrap split">
{wide("chlebicky-pohosteni", "Pohoštění na akci: paštika s brusinkami, tatarák na topinkách a chlebíčky se salámem v krabici")}
<div class="prose">
<h2 id="h-pohosteni">Chlebíčky a pohoštění na akce</h2>
<p>Chlebíčky, zákusky z naší cukrárny a další pohoštění připravíme na rodinné oslavy, porady i firemní akce.</p>
<p>Napište nám, na kdy a pro kolik hostů pohoštění potřebujete, a připravíme vám nabídku.</p>
<div class="btns"><a class="btn btn-ink" href="mailto:{EMAIL}?subject={quote("Pohoštění na akci")}">Napsat poptávku</a></div>
</div>
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-misy-foto">
<div class="wrap">
<h2 class="h3" id="h-misy-foto">Jak mísy vypadají</h2>
<div class="gallery">{shot("misa-salamova", "Salámová mísa", "Salámová mísa se šunkou, salámy a klobáskami")}{shot("misa-syrova", "Sýrová mísa", "Sýrová mísa s plísňovými sýry, mozzarellou a zeleninou")}{shot("misa-mix", "Mix mísa", "Mix mísa se šunkou, salámem, sýry a mozzarellou")}</div>
</div>
</section>
{faq_section("h-faq-catering", FAQ_CATERING)}
<section class="sec sec--paper sec--tight" id="objednavka-misy" aria-labelledby="h-objednavka-misy">
<div class="wrap form-wrap">
<div class="form-intro prose">
<h2 id="h-objednavka-misy">Objednávka mísy</h2>
<p class="muted">Vyberte mísu a termín. Objednávku vám potvrdíme telefonicky.</p>
<p class="muted">Raději voláte? <a href="tel:{TEL}">{TEL_H}</a></p>
</div>
{PLATTER_FORM}
</div>
</section>'''

KAVARNA_PAGE = page_hero(
    "Kavárna", "Rodinná kavárna na Sušilově náměstí v Rousínově",
    "Dobrá káva, zákusky z naší cukrárny, snídaně i sklenka vína. Otevřeno každý den, za hezkého počasí i na zahrádce.",
    '<a class="btn btn-gold" href="#rezervace">Rezervovat stůl</a><a class="btn btn-line" href="/kontakt/">Otevírací doba</a>'
) + f'''
<section class="sec sec--paper" aria-labelledby="h-onas">
<div class="wrap split">
{wide("kavarna-interier", "Interiér kavárny Cafe Moya s dřevěnými stoly, zelenými obrazy z mechu a barem")}
<div class="prose">
<h2 id="h-onas">Místo pro setkání</h2>
<p>Cafe Moya je rodinná kavárna v centru Rousínova. Kávu připravujeme s péčí, dezerty vyrábíme sami a dbáme na kvalitní suroviny.</p>
<p>Přijďte se potkat s přáteli, s rodinou nebo na pracovní schůzku. Rádi vám vyjdeme vstříc.</p>
<h3>Dobré vědět</h3>
{perks()}
</div>
</div>
</section>
<section class="sec sec--night" id="napojovy-listek" aria-labelledby="h-nabidka">
<div class="wrap">
<div class="sec-head">
<h2 id="h-nabidka">Co u nás ochutnáte</h2>
<p class="muted">Ve všední dny od 8 do 19 h, o víkendu od 9 do 19 h.</p>
</div>
{COLS}
<p class="sec-foot muted">[Sem vložíme celý nápojový lístek s cenami.]</p>
</div>
</section>
<section class="sec sec--paper" aria-labelledby="h-zahradka">
<div class="wrap split split--rev">
<div class="prose">
<h2 id="h-zahradka">Klidná zahrádka</h2>
<p>Za hezkého počasí můžete sedět venku na zahrádce pod slunečníky. Hodí se na kávu, zmrzlinový pohár i sklenku vína.</p>
</div>
{wide("kavarna-zahradka", "Zahrádka kavárny Cafe Moya se slunečníky, stolky a květinami")}
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-kavarna-foto">
<div class="wrap">
<h2 class="h3" id="h-kavarna-foto">Z naší kavárny</h2>
<div class="gallery">{shot("kavarna-kava", "Káva a zákusky", "Cappuccino, ovocný dortík a letní drink s jahodou na stolku na zahrádce", tall=True)}{shot("kavarna-napoje", "Nápoje a dezerty", "Matcha latte, ledový nápoj a dezert z pusinek s jahodami", tall=True)}{shot("kavarna-snidane", "Snídaně", "Snídaně: míchaná vajíčka na toastu, palačinky s jahodami, cappuccino a čaj", tall=True)}</div>
</div>
</section>
{faq_section("h-faq-kavarna", FAQ_KAVARNA)}
<section class="sec sec--paper sec--tight" id="rezervace" aria-labelledby="h-rezervace">
<div class="wrap form-wrap">
<div class="form-intro prose">
<h2 id="h-rezervace">Rezervace stolu</h2>
<p class="muted">Napište nám, kdy a kolik vás přijde. Rezervaci vám potvrdíme telefonicky.</p>
<p class="muted">Raději voláte? <a href="tel:{TEL}">{TEL_H}</a></p>
</div>
{TABLE_FORM}
</div>
</section>'''

KONTAKT_PAGE = page_hero(
    "Kontakt", "Kontakt a otevírací doba",
    "Najdete nás v centru Rousínova na Sušilově náměstí. Rádi vám poradíme s objednávkou i rezervací.",
    f'<a class="btn btn-gold" href="tel:{TEL}">Zavolat {TEL_H}</a><a class="btn btn-line" href="mailto:{EMAIL}">Napsat e-mail</a>'
) + f'''
<section class="sec sec--paper" aria-labelledby="h-adresa">
<div class="wrap contact">
<div class="contact-info prose">
<h2 id="h-adresa">Adresa a otevírací doba</h2>
{ADDRESS}
{hours([("Osobní odběr obědů", "11:00–14:00")])}
<h3>Telefon a e-mail</h3>
<dl class="hours">
<div><dt>Kavárna, dorty, catering</dt><dd><a href="tel:{TEL}">{TEL_H}</a></dd></div>
<div><dt>Rozvoz obědů po 9:00</dt><dd><a href="tel:{TEL2}">{TEL2_H}</a></dd></div>
<div><dt>E-mail</dt><dd><a href="mailto:{EMAIL}">{EMAIL}</a></dd></div>
</dl>
<h3>Parkování a přístup</h3>
<p>{PARKING_TEXT} Do kavárny vede bezbariérový přístup a pejsci jsou u nás vítáni.</p>
<div class="btns"><a class="btn btn-ink" href="{MAPY}">Navigovat přes Mapy.com</a><a class="btn btn-line" href="{GMAPS}">Google Maps</a></div>
</div>
{MAP}
</div>
</section>
<section class="sec sec--paper sec--tight" id="prace" aria-labelledby="h-prace">
<div class="wrap prose">
<h2 id="h-prace">Práce u nás</h2>
<p>Hledáme číšníka nebo servírku do kavárny, na hlavní pracovní poměr i na dohodu. Vyučení v oboru je výhodou, ale není podmínkou.</p>
<p>Ozvěte se na <a href="tel:{TEL}">{TEL_H}</a> nebo na <a href="mailto:{EMAIL}">{EMAIL}</a>.</p>
</div>
</section>'''

PRIVACY_PAGE = page_hero(
    "Ochrana osobních údajů", "Ochrana osobních údajů",
    "Jak nakládáme s údaji, které nám pošlete přes formuláře na webu."
) + f'''
<section class="sec sec--paper">
<div class="wrap prose">
<p>Správcem osobních údajů je [doplnit: obchodní firma, IČO a sídlo provozovatele] (dále jen „Cafe Moya“).</p>
<h2 class="h3">Jaké údaje zpracováváme a proč</h2>
<p>Když nám přes web pošlete poptávku dortu, objednávku mísy nebo rezervaci stolu, zpracujeme vaše jméno, telefon, e-mail a údaje, které do formuláře napíšete. Použijeme je jen k vyřízení vaší žádosti a uchováme je jen po dobu nezbytně nutnou.</p>
<h2 class="h3">Cookies</h2>
<p>Web nepoužívá analytické ani reklamní cookies. Mapa na stránkách pochází z OpenStreetMap; při jejím zobrazení se váš prohlížeč připojí k jejich serverům, které návštěvníky nesledují.</p>
<h2 class="h3">Vaše práva</h2>
<p>Máte právo na přístup ke svým údajům, na jejich opravu nebo výmaz a právo vznést námitku. Stačí napsat na <a href="mailto:{EMAIL}">{EMAIL}</a>. Stížnost můžete podat u Úřadu pro ochranu osobních údajů (<a href="https://www.uoou.cz">www.uoou.cz</a>).</p>
</div>
</section>'''

NOTFOUND_PAGE = '''<section class="sec sec--night">
<div class="wrap prose">
<h1>Tuhle stránku jsme nenašli</h1>
<p class="lead muted">Možná se změnila adresa. Zkuste polední menu nebo úvodní stránku.</p>
<div class="btns"><a class="btn btn-gold" href="/obedove-menu/">Polední menu</a><a class="btn btn-line" href="/">Úvodní stránka</a></div>
</div>
</section>'''

# ---------------------------------------------------------------- strukturovaná data
# Na každé stránce je jeden propojený graf schema.org: web, podnik, stránka, drobečková navigace
# a podle stránky menu nebo služba. Všechny údaje odpovídají viditelnému textu na stránce.
BIZ_ID = SITE + "/#kavarna"
WEB_ID = SITE + "/#web"
CITIES = ["Rousínov"] + VILLAGES
SERVICES = [("/rozvoz-obedu/", "Rozvoz obědů zdarma"), ("/dorty-na-objednavku/", "Dorty na objednávku a zákusky"),
            ("/catering/", "Obložené mísy a catering")]


def cities():
    return [{"@type": "City", "name": c} for c in CITIES]


def business_node():
    return {
        "@type": ["CafeOrCoffeeShop", "Bakery"],
        "@id": BIZ_ID,
        "name": "Cafe Moya",
        "description": BUSINESS_DESC,
        "url": SITE + "/",
        "logo": SITE + "/assets/apple-touch-icon.png",
        "image": [SITE + "/assets/foto/kavarna-interier-1200.webp", SITE + "/assets/foto/kavarna-zahradka-1200.webp",
                  SITE + "/assets/foto/uvod-kava-1000.webp"],
        "telephone": TEL,
        "email": EMAIL,
        "address": {"@type": "PostalAddress", "streetAddress": "Sušilovo náměstí 61/39",
                    "addressLocality": "Rousínov", "postalCode": "683 01",
                    "addressRegion": "Jihomoravský kraj", "addressCountry": "CZ"},
        "geo": {"@type": "GeoCoordinates", "latitude": 49.203344, "longitude": 16.882403},
        "hasMap": MAPY,
        "openingHoursSpecification": [
            {"@type": "OpeningHoursSpecification",
             "dayOfWeek": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"], "opens": "08:00", "closes": "19:00"},
            {"@type": "OpeningHoursSpecification",
             "dayOfWeek": ["Saturday", "Sunday"], "opens": "09:00", "closes": "19:00"}],
        "servesCuisine": ["Česká kuchyně", "Káva", "Dezerty"],
        "hasMenu": SITE + "/obedove-menu/",
        "acceptsReservations": True,
        "currenciesAccepted": "CZK",
        "paymentAccepted": "Platební karta, stravenky Edenred, Up, Pluxee",
        "amenityFeature": [{"@type": "LocationFeatureSpecification", "name": name, "value": True}
                           for name in AMENITY_FEATURES],
        "areaServed": cities(),
        "sameAs": [FACEBOOK, INSTAGRAM, FIRMY],
        "hasOfferCatalog": {"@type": "OfferCatalog", "name": "Nabídka Cafe Moya", "itemListElement": [
            {"@type": "Offer", "itemOffered": {"@type": "Service", "@id": SITE + path + "#sluzba",
                                               "name": name, "url": SITE + path}}
            for path, name in SERVICES]},
        "potentialAction": [
            {"@type": "OrderAction", "name": "Objednat polední menu s rozvozem nebo osobním odběrem",
             "target": {"@type": "EntryPoint", "urlTemplate": ORDER, "inLanguage": "cs",
                        "actionPlatform": ["https://schema.org/DesktopWebPlatform",
                                           "https://schema.org/MobileWebPlatform"]},
             "deliveryMethod": ["http://purl.org/goodrelations/v1#DeliveryModeOwnFleet",
                                "http://purl.org/goodrelations/v1#DeliveryModePickUp"]},
            {"@type": "ReserveAction", "name": "Rezervovat stůl v kavárně",
             "target": {"@type": "EntryPoint", "urlTemplate": SITE + "/kavarna/#rezervace", "inLanguage": "cs"},
             "result": {"@type": "FoodEstablishmentReservation", "name": "Rezervace stolu"}}],
    }


def service_node(path, name, service_type, description, **extra):
    node = {"@type": "Service", "@id": SITE + path + "#sluzba", "name": name, "serviceType": service_type,
            "description": description, "url": SITE + path, "provider": {"@id": BIZ_ID}}
    node.update(extra)
    return node


ROZVOZ_SERVICE = service_node(
    "/rozvoz-obedu/", "Rozvoz obědů zdarma", "Rozvoz poledního menu",
    f"Rozvoz poledního menu zdarma po Rousínově, do Vyškova a do obcí {villages_text()}. "
    "Objednávky na daný den online do 9:00, pro Vyškov do 10:30.",
    areaServed=cities(),
    offers={"@type": "Offer", "price": "0", "priceCurrency": "CZK",
            "description": "Rozvoz je zdarma, platí se jen jídlo."},
    availableChannel={"@type": "ServiceChannel", "serviceUrl": ORDER,
                      "servicePhone": {"@type": "ContactPoint", "telephone": TEL2}})

DORTY_SERVICE = service_node(
    "/dorty-na-objednavku/", "Dorty na objednávku a zákusky", "Cukrářská výroba",
    "Dorty na objednávku na narozeniny, výročí, křtiny i firemní akce a zákusky z vlastní výroby. "
    "Poptávka přes web nebo telefon.")

CATERING_SERVICE = service_node(
    "/catering/", "Obložené mísy a catering", "Catering",
    "Obložené mísy salámové, sýrové i mix, chlebíčky a pohoštění na oslavy a firemní akce. "
    "Objednávky nejméně 1 pracovní den předem.",
    hasOfferCatalog={"@type": "OfferCatalog", "name": "Obložené mísy", "itemListElement": [
        {"@type": "Offer", "name": f"{name} {w}", "price": str(price), "priceCurrency": "CZK",
         "itemOffered": {"@type": "MenuItem", "name": f"{name} {w}"}}
        for name, sizes in PLATTERS for w, price in sizes]})


def menu_node():
    sections = []
    for d in cooking_days():
        items = []
        if d["soup"]:
            soup, alg = d["soup"]
            items.append({"@type": "MenuItem", "name": soup,
                          "description": "Polévka v ceně menu." + (f" Alergeny: {alg}" if alg else "")})
        for name, garnish, alg, price in d["dishes"]:
            desc = ". ".join(x for x in (garnish, f"Alergeny: {alg}" if alg else "") if x)
            item = {"@type": "MenuItem", "name": name}
            if desc:
                item["description"] = desc
            if price:
                item["offers"] = {"@type": "Offer", "price": str(price), "priceCurrency": "CZK",
                                  "validFrom": d["date"], "validThrough": d["date"]}
            items.append(item)
        sections.append({"@type": "MenuSection",
                         "name": f"{d['name']} {plain(dm(d['date']))} {d['date'][:4]}",
                         "hasMenuItem": items})
    weeks = " a ".join(f"{w['week']}. týden ({plain(w['label'])})" for w in WEEKS)
    node = {"@type": "Menu", "@id": SITE + "/obedove-menu/#menu",
            "name": "Polední menu Cafe Moya" + (f", {weeks}" if weeks else ""),
            "description": "Polední menu ve všední dny od 11 do 14 h: polévka a pět hlavních jídel, polévka je "
                           "v ceně. " + GLUTEN_TEXT,
            "inLanguage": "cs", "mainEntityOfPage": {"@id": SITE + "/obedove-menu/#stranka"}}
    if sections:
        node["hasMenuSection"] = sections
    return node


def website_node():
    return {"@type": "WebSite", "@id": WEB_ID, "url": SITE + "/", "name": "Cafe Moya", "inLanguage": "cs",
            "publisher": {"@id": BIZ_ID}}


def page_node(p):
    url = SITE + p["path"]
    base = p.get("page_type", "WebPage")
    node = {"@type": [base, "FAQPage"] if p.get("faq") else base, "@id": url + "#stranka", "url": url,
            "name": p["title"], "description": p["desc"], "inLanguage": "cs",
            "isPartOf": {"@id": WEB_ID}, "about": {"@id": BIZ_ID}}
    if p.get("image"):
        node["primaryImageOfPage"] = {"@type": "ImageObject", "url": SITE + p["image"]}
    if p.get("crumb"):
        node["breadcrumb"] = {"@id": url + "#drobecky"}
    if p.get("faq"):
        node["mainEntity"] = [{"@type": "Question", "name": q,
                               "acceptedAnswer": {"@type": "Answer", "text": plain(a)}} for q, a in p["faq"]]
    return node


def crumbs_node(p):
    url = SITE + p["path"]
    return {"@type": "BreadcrumbList", "@id": url + "#drobecky", "itemListElement": [
        {"@type": "ListItem", "position": 1, "name": "Úvod", "item": SITE + "/"},
        {"@type": "ListItem", "position": 2, "name": p["crumb"], "item": url}]}


def graph(p):
    nodes = [website_node(), business_node(), page_node(p)]
    if p.get("crumb"):
        nodes.append(crumbs_node(p))
    nodes += p.get("nodes", [])
    return {"@context": "https://schema.org", "@graph": nodes}

# ---------------------------------------------------------------- seznam stránek
PAGES = [
    dict(path="/", title="Cafe Moya Rousínov – kavárna, cukrárna a obědy s rozvozem",
         desc="Rodinná kavárna a cukrárna na Sušilově náměstí v Rousínově u Vyškova. Dorty na objednávku, "
              "zákusky, obložené mísy a polední menu s rozvozem zdarma.",
         body=HOME, faq=FAQ_HOME, image="/assets/foto/uvod-kava-1000.webp"),
    dict(path="/obedove-menu/", title="Obědové menu Rousínov – denní menu na celý týden | Cafe Moya",
         desc=plain(f"Polední menu Cafe Moya v Rousínově: polévka a pět jídel od {kc(LUNCH_MIN)}, výdej Po–Pá "
                    "11–14 h. Rozvoz zdarma do Rousínova, Vyškova a okolních obcí."),
         body=MENU_PAGE, crumb="Polední menu", faq=FAQ_MENU, nodes=[menu_node()]),
    dict(path="/rozvoz-obedu/", title="Rozvoz obědů zdarma – Rousínov, Vyškov a okolí | Cafe Moya",
         desc="Rozvážíme polední menu zdarma po Rousínově, do Vyškova, Pozořic, Holubic a dalších obcí. "
              "Objednávky online do 9:00, pro Vyškov do 10:30.",
         body=ROZVOZ_PAGE, crumb="Rozvoz obědů", faq=FAQ_ROZVOZ, nodes=[ROZVOZ_SERVICE]),
    dict(path="/dorty-na-objednavku/", title="Dorty na objednávku Rousínov u Vyškova – cukrárna Cafe Moya",
         desc="Dorty na objednávku na narozeniny, výročí, křtiny i firemní akce a zákusky z vlastní výroby. "
              "Cukrárna Cafe Moya na Sušilově náměstí v Rousínově.",
         body=DORTY_PAGE, crumb="Dorty a zákusky", faq=FAQ_DORTY, nodes=[DORTY_SERVICE],
         image="/assets/foto/dort-na-objednavku-1200.webp"),
    dict(path="/catering/", title="Obložené mísy a catering – Rousínov u Vyškova | Cafe Moya",
         desc=plain(f"Obložené mísy salámové, sýrové i mix od {kc(PLATTER_MIN)} za kilogram, chlebíčky a pohoštění "
                    "na oslavy i firemní akce. Objednávky nejméně 1 pracovní den předem."),
         body=CATERING_PAGE, crumb="Catering", faq=FAQ_CATERING, nodes=[CATERING_SERVICE],
         image="/assets/foto/chlebicky-pohosteni-1200.webp"),
    dict(path="/kavarna/", title="Kavárna Rousínov – káva, zákusky a snídaně | Cafe Moya",
         desc="Rodinná kavárna na Sušilově náměstí v Rousínově. Káva, domácí zákusky, snídaně, víno a koktejly. "
              "Zahrádka, Wi-Fi, pejsci vítáni, parkování před kavárnou.",
         body=KAVARNA_PAGE, crumb="Kavárna", faq=FAQ_KAVARNA, image="/assets/foto/kavarna-interier-1200.webp"),
    dict(path="/kontakt/", title="Kontakt a otevírací doba – Cafe Moya Rousínov",
         desc="Cafe Moya, Sušilovo náměstí 61/39, Rousínov. Otevřeno Po–Pá 8–19 h, So–Ne 9–19 h. "
              "Parkování přímo před kavárnou, hodina zdarma. Telefon 605 226 781.",
         body=KONTAKT_PAGE, crumb="Kontakt", page_type="ContactPage"),
    dict(path="/ochrana-osobnich-udaju/", title="Ochrana osobních údajů | Cafe Moya",
         desc="Jak Cafe Moya nakládá s osobními údaji, které nám pošlete přes formuláře na webu.",
         body=PRIVACY_PAGE, crumb="Ochrana osobních údajů", sitemap=False),
]
NOT_FOUND = dict(path="/404.html", title="Stránka nenalezena | Cafe Moya",
                 desc="Tuto stránku jsme nenašli.", body=NOTFOUND_PAGE, noindex=True)

# ---------------------------------------------------------------- sestavení dokumentů
FONT_FACES = """/* Písma Comfortaa a Nunito Sans (SIL Open Font License), uložená přímo na webu */
@font-face { font-family: 'Comfortaa'; font-style: normal; font-display: swap; font-weight: 300 700; src: url(/assets/fonts/comfortaa-latin-ext-wght-normal.woff2) format('woff2'); unicode-range: U+0100-02BA, U+02BD-02C5, U+02C7-02CC, U+02CE-02D7, U+02DD-02FF, U+0304, U+0308, U+0329, U+1D00-1DBF, U+1E00-1E9F, U+1EF2-1EFF, U+2020, U+20A0-20AB, U+20AD-20C0, U+2113, U+2C60-2C7F, U+A720-A7FF; }
@font-face { font-family: 'Comfortaa'; font-style: normal; font-display: swap; font-weight: 300 700; src: url(/assets/fonts/comfortaa-latin-wght-normal.woff2) format('woff2'); unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+0304, U+0308, U+0329, U+2000-206F, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD; }
@font-face { font-family: 'Nunito Sans'; font-style: normal; font-display: swap; font-weight: 200 1000; src: url(/assets/fonts/nunito-sans-latin-ext-wght-normal.woff2) format('woff2'); unicode-range: U+0100-02BA, U+02BD-02C5, U+02C7-02CC, U+02CE-02D7, U+02DD-02FF, U+0304, U+0308, U+0329, U+1D00-1DBF, U+1E00-1E9F, U+1EF2-1EFF, U+2020, U+20A0-20AB, U+20AD-20C0, U+2113, U+2C60-2C7F, U+A720-A7FF; }
@font-face { font-family: 'Nunito Sans'; font-style: normal; font-display: swap; font-weight: 200 1000; src: url(/assets/fonts/nunito-sans-latin-wght-normal.woff2) format('woff2'); unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+0304, U+0308, U+0329, U+2000-206F, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD; }

"""


def ld_scripts(objs):
    out = []
    for o in objs:
        s = json.dumps(o, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        out.append(f'<script type="application/ld+json">{s}</script>')
    return "\n".join(out)


def document(p, ver):
    url = SITE + p["path"]
    meta_index = ('<meta name="robots" content="noindex">' if p.get("noindex") else
                  f'<link rel="canonical" href="{url}">\n<meta name="robots" content="max-image-preview:large">')
    return f'''<!doctype html>
<html lang="cs" class="no-js">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{esc(p["title"])}</title>
<meta name="description" content="{esc(p["desc"])}">
{meta_index}
<meta property="og:type" content="website">
<meta property="og:locale" content="cs_CZ">
<meta property="og:site_name" content="Cafe Moya">
<meta property="og:title" content="{esc(p["title"])}">
<meta property="og:description" content="{esc(p["desc"])}">
<meta property="og:url" content="{url}">
<meta property="og:image" content="{SITE}/assets/og-image.jpg">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta name="twitter:card" content="summary_large_image">
<meta name="theme-color" content="#000000">
<link rel="icon" href="/assets/favicon-48.png" sizes="48x48" type="image/png">
<link rel="icon" href="/assets/favicon.svg" type="image/svg+xml">
<link rel="apple-touch-icon" href="/assets/apple-touch-icon.png">
<link rel="preload" href="/assets/fonts/comfortaa-latin-wght-normal.woff2" as="font" type="font/woff2" crossorigin>
<link rel="preload" href="/assets/fonts/nunito-sans-latin-wght-normal.woff2" as="font" type="font/woff2" crossorigin>
<link rel="stylesheet" href="/assets/styles.css?v={ver}">
<script>document.documentElement.className = "js";</script>
{"" if p.get("noindex") else ld_scripts([graph(p)])}
</head>
<body>
{vlnka(header(p["path"]))}
<main id="obsah">
{vlnka(p["body"])}
</main>
{vlnka(footer())}
<script src="/assets/site.js?v={ver}" defer></script>
</body>
</html>
'''


def rewrite_for_preview(text):
    def repl(m):
        href = m.group(1)
        if "#" in href:
            return 'href="#' + href.split("#", 1)[1] + '"'
        return 'href="#' + href + '"'
    return re.sub(r'href="(/[^"]*)"', repl, text)


def inline_images(text):
    """Pro jednosouborový náhled: fotky z /assets vložit přímo do HTML."""
    import base64
    text = re.sub(r'\s(?:srcset|sizes)="[^"]*"', "", text)
    def repl(m):
        data = (SRC / "assets" / m.group(1)).read_bytes()
        ext = m.group(1).rsplit(".", 1)[1]
        mime = {"webp": "image/webp", "jpg": "image/jpeg", "png": "image/png", "svg": "image/svg+xml"}[ext]
        return 'src="data:' + mime + ';base64,' + base64.b64encode(data).decode("ascii") + '"'
    return re.sub(r'src="/assets/([^"]+\.(?:webp|jpg|png|svg))"', repl, text)


def preview_map(text):
    note = ('<div class="map-preview">' + PIN + '<p><b>Sušilovo náměstí 61/39, Rousínov</b></p>'
            '<p>Tady bude mapa OpenStreetMap. V náhledu se nenačte, na webu ano.</p></div>')
    return re.sub(r"<iframe[^>]*></iframe>", note, text)


def preview_document(css, js, router):
    fonts = ("https://fonts.googleapis.com/css2?family=Comfortaa:wght@400;700"
             "&family=Nunito+Sans:wght@400;600;700;800&display=swap").replace("&", "&amp;")
    pages = []
    for p in PAGES:
        hidden = "" if p["path"] == "/" else " hidden"
        pages.append(f'<div class="page" data-route="{p["path"]}" data-title="{esc(p["title"])}"{hidden}>\n'
                     f'{vlnka(rewrite_for_preview(preview_map(inline_images(p["body"]))))}\n</div>')
    return f'''<!doctype html>
<html lang="cs" class="js">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{esc(PAGES[0]["title"])}</title>
<link rel="stylesheet" href="{fonts}">
<style>
{css}
.page[hidden] {{ display: none; }}
.sim {{ position: fixed; left: 16px; bottom: calc(16px + env(safe-area-inset-bottom, 0px)); z-index: 30; display: flex; flex-wrap: wrap; align-items: center; gap: 6px; max-width: calc(100% - 32px); padding: 10px 12px; border: 2px solid var(--gold); border-radius: 16px; background: #000; color: #fff; font-size: 14px; box-shadow: 0 6px 24px rgba(0, 0, 0, .35); }}
.sim span {{ margin-right: 4px; font-weight: 800; }}
.sim button {{ min-height: 40px; padding: 0 14px; border: 2px solid #fff; border-radius: 999px; background: transparent; color: #fff; font: inherit; font-weight: 700; cursor: pointer; }}
.sim button[aria-pressed="true"] {{ background: var(--gold); border-color: var(--gold); color: #000; }}
</style>
</head>
<body>
{vlnka(rewrite_for_preview(inline_images(header(None))))}
<main id="obsah">
{"".join(pages)}
</main>
{vlnka(rewrite_for_preview(inline_images(footer())))}
<script>window.MOYA_PREVIEW = true;</script>
<script>
{js}
</script>
<script>
{router}
</script>
</body>
</html>
'''

# ---------------------------------------------------------------- doplňkové soubory
OLD_URLS = [("inpage/poledni-menu", "/obedove-menu/"), ("inpage/poledni-menu-old", "/obedove-menu/"),
            ("inpage/jak-u-nas-rozvoz-funguje", "/rozvoz-obedu/"),
            ("inpage/potvrzeni-objednavky-rozvozu", "/rozvoz-obedu/"),
            ("inpage/oblozene-misy", "/catering/"), ("inpage/silvestrovske-oslavy", "/catering/"),
            ("inpage/kontakt", "/kontakt/")]


def htaccess():
    rules = "\n".join(f"RewriteRule ^{old}/?$ {new} [R=301,L]" for old, new in OLD_URLS)
    return f"""# Cafe Moya – nastavení pro hostingy s Apache (Wedos, Forpsi, Active24 a další)
RewriteEngine On

# Vždy https://www.cafemoya.cz
RewriteCond %{{HTTPS}} !=on [OR]
RewriteCond %{{HTTP_HOST}} !^www\\.cafemoya\\.cz$ [NC]
RewriteRule ^ https://www.cafemoya.cz%{{REQUEST_URI}} [R=301,L]

# Staré adresy z inPage přesměrovat na nové stránky
{rules}
RewriteRule ^inpage/ / [R=301,L]

ErrorDocument 404 /404.html

<IfModule mod_expires.c>
  ExpiresActive On
  ExpiresByType text/html "access plus 0 seconds"
  ExpiresByType text/css "access plus 1 year"
  ExpiresByType application/javascript "access plus 1 year"
  ExpiresByType font/woff2 "access plus 1 year"
  ExpiresByType image/svg+xml "access plus 1 year"
  ExpiresByType image/jpeg "access plus 1 month"
  ExpiresByType image/png "access plus 1 month"
  ExpiresByType image/webp "access plus 1 month"
</IfModule>

<IfModule mod_deflate.c>
  AddOutputFilterByType DEFLATE text/html text/css application/javascript image/svg+xml application/xml text/plain
</IfModule>
"""


def redirects():
    lines = [f"/{old}/  {new}  301" for old, new in OLD_URLS]
    lines.append("/inpage/*  /  301")
    return "# Netlify / Cloudflare Pages: staré adresy z inPage\n" + "\n".join(lines) + "\n"


def sitemap():
    urls = "".join(f"  <url><loc>{SITE}{p['path']}</loc><lastmod>{UPDATED}</lastmod></url>\n"
                   for p in PAGES if p.get("sitemap", True))
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' + urls + '</urlset>\n')


ROBOTS = f"User-agent: *\nAllow: /\n\nSitemap: {SITE}/sitemap.xml\n"


def llms_menu():
    """Aktuální polední menu v textu, aby AI asistenti uměli odpovědět, co se kdy vaří."""
    out = []
    for w in WEEKS:
        out += [f"Polední menu, {w['week']}. týden ({plain(w['label'])}); polévka je v ceně menu, "
                "čísla v závorkách jsou alergeny:", ""]
        for d in w["days"]:
            head = f"- {d['name']} {plain(dm(d['date']))}: "
            if d["closed"]:
                out.append(head + d["closed"])
                continue
            parts = []
            if d["soup"]:
                soup, alg = d["soup"]
                parts.append(f"polévka {soup[0].lower() + soup[1:]}" + (f" ({alg})" if alg else ""))
            for i, (name, garnish, alg, price) in enumerate(d["dishes"], 1):
                parts.append(f"{i}. {name}" + (f", {garnish}" if garnish else "")
                             + (f" ({alg})" if alg else "") + (f" {plain(kc(price))}" if price else ""))
            out.append(head + ("; ".join(parts) if parts else "menu doplníme"))
        out.append("")
    return out


def llms_txt():
    """Souhrn webu v Markdownu podle návrhu llmstxt.org pro AI asistenty a agenty."""
    pages = [("/obedove-menu/", "Polední menu", "menu na celý týden s cenami a alergeny"),
             ("/rozvoz-obedu/", "Rozvoz obědů", "rozvoz zdarma, obce, časy rozvozu a uzávěrky objednávek"),
             ("/dorty-na-objednavku/", "Dorty a zákusky", "dorty na objednávku, zákusky z vitríny a poptávkový formulář"),
             ("/catering/", "Catering a obložené mísy", "ceník mís, pohoštění na akce a objednávkový formulář"),
             ("/kavarna/", "Kavárna", "nabídka kavárny, zahrádka a rezervace stolu"),
             ("/kontakt/", "Kontakt", "adresa, otevírací doba, telefony a mapa")]
    platters = "; ".join(f"{name} " + ", ".join(f"{w} {plain(kc(p))}" for w, p in sizes) for name, sizes in PLATTERS)
    lines = [
        "# Cafe Moya", "",
        f"> {BUSINESS_DESC}", "",
        f"Údaje platí k {TODAY.day}. {TODAY.month}. {TODAY.year}. Aktuální stav je vždy na stránkách níže.", "",
        "- Adresa: Sušilovo náměstí 61/39, 683 01 Rousínov (okres Vyškov, Jihomoravský kraj)",
        "- Otevírací doba kavárny: pondělí až pátek 8:00–19:00, sobota a neděle 9:00–19:00",
        f"- Polední menu: pondělí až pátek 11:00–14:00, polévka a pět hlavních jídel, polévka je v ceně; "
        f"v aktuálním menu {plain(lunch_price_text())}",
        f"- Rozvoz obědů: zdarma po Rousínově, do Vyškova a do obcí {villages_text()}; objednávky na daný den "
        f"online do 9:00, pro Vyškov do 10:30, později telefonicky na +420 {TEL2_H}",
        f"- Online objednávka obědů: {ORDER}",
        f"- Obložené mísy: {platters}; kilogram obvykle stačí pro 5–6 osob; objednávky nejméně 1 pracovní den předem",
        "- Dorty na objednávku: na narozeniny, výročí, křtiny i firemní akce; poptávka přes web nebo telefon",
        "- Platba: kartou i stravenkami Edenred, Up, Pluxee",
        "- Parkování: přímo před kavárnou na Sušilově náměstí, hodina parkování zdarma",
        "- Vybavení: bezbariérový přístup, Wi-Fi pro hosty, pejsci vítáni, venkovní zahrádka",
        f"- Bezlepková jídla: {GLUTEN_TEXT[0].lower() + GLUTEN_TEXT[1:]}",
        f"- Rezervace stolu: formulář na {SITE}/kavarna/#rezervace nebo telefon",
        f"- Telefon: +420 {TEL_H} (kavárna, dorty, catering), +420 {TEL2_H} (rozvoz obědů po 9:00)",
        f"- E-mail: {EMAIL}", ""]
    lines += llms_menu()
    lines += ["## Stránky", ""]
    lines += [f"- [{title}]({SITE}{path}): {desc}" for path, title, desc in pages]
    lines += ["", "## Optional", "",
              f"- [Ochrana osobních údajů]({SITE}/ochrana-osobnich-udaju/): jak zpracováváme údaje z formulářů",
              f"- [Facebook]({FACEBOOK}): novinky z kavárny",
              f"- [Instagram]({INSTAGRAM}): fotky z kavárny a cukrárny",
              f"- [Firmy.cz]({FIRMY}): profil kavárny", ""]
    return "\n".join(lines)

def out_path(path):
    if path == "/":
        return DIST / "index.html"
    if path.endswith(".html"):
        return DIST / path.lstrip("/")
    return DIST / path.strip("/") / "index.html"


def build(preview=False):
    css = (SRC / "styles.css").read_text(encoding="utf-8")
    js = (SRC / "site.js").read_text(encoding="utf-8")
    ver = hashlib.md5((css + js).encode("utf-8")).hexdigest()[:8]

    shutil.rmtree(DIST, ignore_errors=True)
    shutil.copytree(SRC / "assets", DIST / "assets")
    (DIST / "assets" / "styles.css").write_text(FONT_FACES + css, encoding="utf-8")
    (DIST / "assets" / "site.js").write_text(js, encoding="utf-8")

    for p in PAGES + [NOT_FOUND]:
        target = out_path(p["path"])
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(document(p, ver), encoding="utf-8")

    (DIST / "sitemap.xml").write_text(sitemap(), encoding="utf-8")
    (DIST / "robots.txt").write_text(ROBOTS, encoding="utf-8")
    (DIST / "llms.txt").write_text(llms_txt(), encoding="utf-8")
    (DIST / ".htaccess").write_text(htaccess(), encoding="utf-8")
    (DIST / "_redirects").write_text(redirects(), encoding="utf-8")
    (DIST / ".assetsignore").write_text(".htaccess\n*.xlsx\n", encoding="utf-8")   # na web nepatří

    if preview:
        router = (SRC / "preview.js").read_text(encoding="utf-8")
        (ROOT / "preview.html").write_text(preview_document(css, js, router), encoding="utf-8")

    forms = "Web3Forms" if WEB3FORMS_KEY else "e-mail (bez klíče Web3Forms)"
    menu = ", ".join(f"{w['week']}. týden ze souboru {w['source']}" for w in WEEKS) or "žádná tabulka s menu"
    print(f"Hotovo: {len(PAGES) + 1} stránek ve složce dist, verze {ver}, formuláře: {forms}")
    print(f"Polední menu: {menu}")


if __name__ == "__main__":
    build(preview="--nahled" in sys.argv)
