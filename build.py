#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cafe Moya – generátor webu (směr B: černá a zlatá)

Použití:   python3 build.py            sestaví web do složky dist/ (to dělá i Cloudflare)
           python3 build.py --nahled   navíc vytvoří preview.html (celý web v jednom souboru)

Polední menu se mění v proměnné MENU níže (později ji může plnit týdenní xlsx).
"""
import datetime
import hashlib
import html
import json
import os
import re
import shutil
import sys
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

# Polední menu: jídla jsou (název, příloha / popis, alergeny, cena)
MENU = {
    "week": 42,
    "label": "12.–16. října 2026",
    "days": [
        {"date": "2026-10-12", "name": "Pondělí", "soup": ("Kulajda s houbami", "1, 3, 7"),
         "dishes": [("Špagety caprese", "rajčata, mozzarella, bazalka", "1, 7", 155),
                    ("Kuřecí plátek se slaninou", "zapečený hermelínem, šťouchané brambory", "7", 159),
                    ("Tomatové risotto", "kuřecí supreme, Gran Biraghi", "7, 12", 199),
                    ("Zeleninový salát", "kuřecí nudličky, medovo-hořčičný dresink, pečivo", "1, 3, 7, 10", 189),
                    ("Sýrová dvojka", "gouda a hermelín, hranolky, tatarská omáčka", "1, 3, 7", 165)],
         "complete": True},
        {"date": "2026-10-13", "name": "Úterý", "soup": ("Krupicová s vejcem", "1, 3, 9"),
         "dishes": [("Brněnský řízek", "bramborová kaše", "1, 3, 7", 155)], "complete": False},
        {"date": "2026-10-14", "name": "Středa", "soup": ("Boršč", "1, 7, 9"),
         "dishes": [("Vepřové kostky na paprice", "těstoviny", "1, 7", 155)], "complete": False},
        {"date": "2026-10-15", "name": "Čtvrtek", "soup": ("Drůbeží vývar s nudlemi a zeleninou", "1, 3, 9"),
         "dishes": [("Pečená sekaná", "vařené brambory, kompot", "1, 3, 7", 155)], "complete": False},
        {"date": "2026-10-16", "name": "Pátek", "soup": ("Gulášovka s chlebem", "1"),
         "dishes": [("Přírodní rybí filé pečené v troubě", "vařené brambory, citronová omáčka", "4, 7, 10", 155)],
         "complete": False},
    ],
}
DAY_IDS = {"Pondělí": ("pondeli", "Po"), "Úterý": ("utery", "Út"), "Středa": ("streda", "St"),
           "Čtvrtek": ("ctvrtek", "Čt"), "Pátek": ("patek", "Pá")}

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


def shot(name, caption, alt):
    """Fotka do galerie na šířku 4:3 s popiskem pod ní."""
    files = ", ".join(f"/assets/foto/{name}-{w}.webp {w}w" for w in (400, 800))
    return (f'<figure class="shot"><img src="/assets/foto/{name}-800.webp" srcset="{files}" '
            'sizes="(max-width: 860px) calc(100vw - 40px), 390px" width="800" height="600" '
            f'alt="{alt}" loading="lazy" decoding="async"><figcaption>{caption}</figcaption></figure>')


def photo(name, alt):
    """Fotka na výšku ve třech velikostech (480, 720 a 960 px široká)."""
    files = ", ".join(f"/assets/foto/{name}-{w}.webp {w}w" for w in (480, 720, 960))
    return (f'<img class="photo" src="/assets/foto/{name}-960.webp" srcset="{files}" '
            'sizes="(max-width: 960px) calc(100vw - 40px), 480px" width="960" height="1286" '
            f'alt="{alt}" loading="lazy" decoding="async">')
ADDRESS = '<address class="addr">Cafe Moya<br>Sušilovo náměstí 61/39<br>683 01 Rousínov</address>'
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


def soup_li(soup):
    name, alg = soup
    return (f'<li><span class="dish"><b>Polévka: {name[0].lower() + name[1:]}</b> '
            f'<span class="alg">({alg})</span></span><span class="price price--incl">v ceně menu</span></li>')


def dish_li(i, dish):
    name, garnish, alg, price = dish
    g = f' <span class="garnish">{garnish}</span>' if garnish else ""
    return (f'<li><span class="dish"><span class="num">{i}.</span><b>{name}</b>{g} '
            f'<span class="alg">({alg})</span></span><span class="price">{kc(price)}</span></li>')


def day_items(day):
    return soup_li(day["soup"]) + "".join(dish_li(i, x) for i, x in enumerate(day["dishes"], 1))


def today_panel():
    days = MENU["days"]
    first = next((d for d in days if d["complete"]), days[0])
    out = []
    for d in days:
        attrs = f' data-day-block data-date="{d["date"]}"'
        if d is not first:
            attrs += " hidden"
        note = "" if d["complete"] else INCOMPLETE
        out.append(f'<div class="menu-day"{attrs}><div class="menu-day-head"><h3>{d["name"]} {dm(d["date"])}</h3>'
                   f'<p data-when>Polední menu, {MENU["week"]}. týden</p></div>'
                   f'<ul class="menu-list">{day_items(d)}</ul>{note}</div>')
    out.append('<div class="menu-day" data-day-fallback hidden><div class="menu-day-head"><h3>Polední menu</h3></div>'
               '<p class="menu-note">Menu na další dny právě připravujeme. Rádi vám ho řekneme na '
               f'<a href="tel:{TEL}">{TEL_H}</a>.</p></div>')
    return "\n".join(out)


def week_menu():
    out = []
    for d in MENU["days"]:
        sid = DAY_IDS[d["name"]][0]
        note = "" if d["complete"] else INCOMPLETE
        out.append(f'<section class="menu-day" id="{sid}" data-menu-date="{d["date"]}" aria-labelledby="h-{sid}">'
                   f'<div class="menu-day-head"><h3 id="h-{sid}">{d["name"]} {dm(d["date"])}</h3>'
                   f'<span class="today-badge" hidden>Dnes</span></div>'
                   f'<ul class="menu-list">{day_items(d)}</ul>{note}</section>')
    return '<div class="week">\n' + "\n".join(out) + '\n</div>'


def day_links():
    return "".join(f'<a href="#{DAY_IDS[d["name"]][0]}" aria-label="{d["name"]}">{DAY_IDS[d["name"]][1]}</a>'
                   for d in MENU["days"])


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
        out.append(f"<details{opened}><summary>{q}{PLUS}</summary><p>{tel_links(a)}</p></details>")
    return '<div class="faq">' + "".join(out) + "</div>"

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
            f'<h2>Sledujte nás</h2><ul><li><a href="{FACEBOOK}">Facebook</a></li></ul></div>'
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
<p class="lead">Zákusky a dorty z vlastní výroby, obložené mísy na oslavy a polední menu s rozvozem zdarma.</p>
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
<p class="small muted">Bereme stravenky Edenred, Up a Pluxee.</p>
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
<p class="muted">Mísy od {kc(739)} za kilogram. Objednávejte nejméně jeden pracovní den předem.</p>
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
<section class="sec sec--paper" aria-labelledby="home-kontakt">
<div class="wrap contact">
<div class="contact-info prose">
<h2 id="home-kontakt">Přijďte k nám</h2>
{ADDRESS}
{hours()}
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
<section class="sec sec--paper" aria-labelledby="h-tyden">
<div class="wrap">
<div class="week-head">
<h2 id="h-tyden">{MENU["week"]}. týden, {MENU["label"]}</h2>
<nav class="day-jump" aria-label="Dny v týdnu">{day_links()}</nav>
</div>
{week_menu()}
</div>
</section>
<section class="sec sec--paper sec--tight" aria-label="Nápoje, stravenky a objednávky">
<div class="wrap info-row">
<div><h3>K menu</h3><ul class="price-list">{k_menu()}</ul></div>
<div><h3>Stravenky</h3><p class="muted">Bereme stravenky Edenred, Up (dříve Chèque Déjeuner) a Pluxee (dříve Sodexo).</p></div>
<div><h3>Objednávky na rozvoz</h3><p class="muted">Online na daný den do 9:00, pro Vyškov do 10:30. Později volejte na <a href="tel:{TEL2}">{TEL2_H}</a>.</p></div>
</div>
</section>
<section class="sec sec--paper sec--tight" id="alergeny" aria-labelledby="h-alergeny">
<div class="wrap">
<h2 class="h3" id="h-alergeny">Seznam alergenů</h2>
<p class="muted">Čísla v závorkách u jídel označují alergeny.</p>
<ol class="allergens">{"".join(f"<li>{a}</li>" for a in ALLERGENS)}</ol>
</div>
</section>'''

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
    "Zákusky a dezerty vyrábíme sami. Na objednávku upečeme dort na narozeniny, výročí i firemní oslavu.",
    f'<a class="btn btn-gold" href="#poptavka-dortu">Poptat dort</a><a class="btn btn-line" href="tel:{TEL}">Zavolat {TEL_H}</a>'
) + f'''
<section class="sec sec--paper" aria-labelledby="h-dort">
<div class="wrap split">
{ph("dort na objednávku, detail zdobení")}
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
{ph("zákusky ve vitríně")}
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-ukazky">
<div class="wrap">
<h2 class="h3" id="h-ukazky">Ukázky naší práce</h2>
<div class="gallery">{ph("narozeninový dort")}{ph("dort na firemní akci")}{ph("zákusky na tácu")}</div>
</div>
</section>
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
    "Plánujete oslavu, posezení s přáteli nebo firemní akci? Připravíme vám obložené mísy z čerstvých surovin, "
    "chlebíčky i další pohoštění.",
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
{ph("chlebíčky a pohoštění na oslavu")}
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
    "Kavárna", "Rodinná kavárna na Sušilově náměstí",
    "Dobrá káva, zákusky z naší cukrárny, snídaně i sklenka vína. Otevřeno každý den, za hezkého počasí i na zahrádce.",
    '<a class="btn btn-gold" href="#rezervace">Rezervovat stůl</a><a class="btn btn-line" href="/kontakt/">Otevírací doba</a>'
) + f'''
<section class="sec sec--paper" aria-labelledby="h-onas">
<div class="wrap split">
{ph("interiér kavárny")}
<div class="prose">
<h2 id="h-onas">Místo pro setkání</h2>
<p>Cafe Moya je rodinná kavárna v centru Rousínova. Kávu připravujeme s péčí, dezerty vyrábíme sami a dbáme na kvalitní suroviny.</p>
<p>Přijďte se potkat s přáteli, s rodinou nebo na pracovní schůzku. Rádi vám vyjdeme vstříc.</p>
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
<h2 id="h-zahradka">Zahrádka na náměstí</h2>
<p>Za hezkého počasí můžete sedět venku na naší klidné zahrádce. Hodí se na kávu, zmrzlinový pohár i sklenku vína.</p>
</div>
{ph("zahrádka kavárny")}
</div>
</section>
<section class="sec sec--paper sec--tight" aria-labelledby="h-kavarna-foto">
<div class="wrap">
<h2 class="h3" id="h-kavarna-foto">Z naší kavárny</h2>
<div class="gallery">{ph("káva a zákusek")}{ph("zmrzlinový pohár")}{ph("snídaně")}</div>
</div>
</section>
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
def business_ld():
    return {
        "@context": "https://schema.org",
        "@type": "CafeOrCoffeeShop",
        "@id": SITE + "/#kavarna",
        "name": "Cafe Moya",
        "url": SITE + "/",
        "image": SITE + "/assets/og-image.jpg",
        "telephone": TEL,
        "email": EMAIL,
        "address": {"@type": "PostalAddress", "streetAddress": "Sušilovo náměstí 61/39",
                    "addressLocality": "Rousínov", "postalCode": "683 01", "addressCountry": "CZ"},
        "geo": {"@type": "GeoCoordinates", "latitude": 49.203344, "longitude": 16.882403},
        "openingHoursSpecification": [
            {"@type": "OpeningHoursSpecification",
             "dayOfWeek": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"], "opens": "08:00", "closes": "19:00"},
            {"@type": "OpeningHoursSpecification",
             "dayOfWeek": ["Saturday", "Sunday"], "opens": "09:00", "closes": "19:00"}],
        "servesCuisine": ["Česká kuchyně", "Káva", "Dezerty"],
        "hasMenu": SITE + "/obedove-menu/",
        "acceptsReservations": True,
        "areaServed": ["Rousínov"] + VILLAGES,
        "sameAs": [FACEBOOK],
    }


def menu_ld():
    sections = []
    for d in MENU["days"]:
        items = [{"@type": "MenuItem", "name": d["soup"][0],
                  "description": f"Polévka v ceně menu. Alergeny: {d['soup'][1]}"}]
        for name, garnish, alg, price in d["dishes"]:
            items.append({"@type": "MenuItem", "name": name,
                          "description": (garnish + ". " if garnish else "") + f"Alergeny: {alg}",
                          "offers": {"@type": "Offer", "price": str(price), "priceCurrency": "CZK"}})
        sections.append({"@type": "MenuSection", "name": f"{d['name']} {dm(d['date']).replace(NB, ' ')}",
                         "hasMenuItem": items})
    return {"@context": "https://schema.org", "@type": "Menu", "@id": SITE + "/obedove-menu/#menu",
            "name": f"Obědové menu Cafe Moya, {MENU['week']}. týden", "inLanguage": "cs",
            "hasMenuSection": sections}


def faq_ld(items):
    return {"@context": "https://schema.org", "@type": "FAQPage",
            "mainEntity": [{"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}}
                           for q, a in items]}


def crumbs_ld(name, path):
    return {"@context": "https://schema.org", "@type": "BreadcrumbList", "itemListElement": [
        {"@type": "ListItem", "position": 1, "name": "Úvod", "item": SITE + "/"},
        {"@type": "ListItem", "position": 2, "name": name, "item": SITE + path}]}

# ---------------------------------------------------------------- seznam stránek
PAGES = [
    dict(path="/", title="Cafe Moya Rousínov – kavárna, cukrárna a obědy s rozvozem",
         desc="Rodinná kavárna a cukrárna na Sušilově náměstí v Rousínově. Dorty a zákusky z vlastní výroby, "
              "obložené mísy na oslavy a polední menu s rozvozem zdarma.",
         body=HOME, ld=[business_ld()]),
    dict(path="/obedove-menu/", title="Obědové menu Rousínov – denní menu na celý týden | Cafe Moya",
         desc="Polední menu Cafe Moya v Rousínově: polévka a pět jídel od 155 Kč, výdej Po–Pá 11–14 h. "
              "Rozvoz zdarma do Rousínova, Vyškova a okolních obcí.",
         body=MENU_PAGE, ld=[menu_ld(), crumbs_ld("Polední menu", "/obedove-menu/")]),
    dict(path="/rozvoz-obedu/", title="Rozvoz obědů zdarma – Rousínov, Vyškov a okolí | Cafe Moya",
         desc="Rozvážíme polední menu zdarma po Rousínově, do Vyškova, Pozořic, Holubic a dalších obcí. "
              "Objednávky online do 9:00, pro Vyškov do 10:30.",
         body=ROZVOZ_PAGE, ld=[faq_ld(FAQ_ROZVOZ), crumbs_ld("Rozvoz obědů", "/rozvoz-obedu/")]),
    dict(path="/dorty-na-objednavku/", title="Dorty na objednávku a zákusky v Rousínově | Cafe Moya",
         desc="Zákusky a dorty z vlastní výroby. Na objednávku upečeme dort na narozeniny, výročí nebo firemní "
              "oslavu. Cukrárna Cafe Moya na Sušilově náměstí v Rousínově.",
         body=DORTY_PAGE, ld=[crumbs_ld("Dorty a zákusky", "/dorty-na-objednavku/")]),
    dict(path="/catering/", title="Obložené mísy a catering – Rousínov, Vyškov | Cafe Moya",
         desc="Obložené mísy salámové, sýrové i mix od 739 Kč za kilogram, chlebíčky a pohoštění na oslavy "
              "i firemní akce. Objednávky nejméně 1 pracovní den předem.",
         body=CATERING_PAGE, ld=[crumbs_ld("Catering", "/catering/")]),
    dict(path="/kavarna/", title="Kavárna Rousínov – káva, zákusky a snídaně | Cafe Moya",
         desc="Rodinná kavárna na Sušilově náměstí v Rousínově. Káva, domácí zákusky, snídaně, panini, "
              "zmrzlinové poháry, víno a koktejly. Otevřeno denně, v létě zahrádka.",
         body=KAVARNA_PAGE, ld=[crumbs_ld("Kavárna", "/kavarna/")]),
    dict(path="/kontakt/", title="Kontakt a otevírací doba – Cafe Moya Rousínov",
         desc="Cafe Moya, Sušilovo náměstí 61/39, 683 01 Rousínov. Otevřeno Po–Pá 8–19 h, So–Ne 9–19 h. "
              "Telefon 605 226 781, e-mail info@cafemoya.cz.",
         body=KONTAKT_PAGE, ld=[business_ld(), crumbs_ld("Kontakt", "/kontakt/")]),
    dict(path="/ochrana-osobnich-udaju/", title="Ochrana osobních údajů | Cafe Moya",
         desc="Jak Cafe Moya nakládá s osobními údaji, které nám pošlete přes formuláře na webu.",
         body=PRIVACY_PAGE, ld=[crumbs_ld("Ochrana osobních údajů", "/ochrana-osobnich-udaju/")], sitemap=False),
]
NOT_FOUND = dict(path="/404.html", title="Stránka nenalezena | Cafe Moya",
                 desc="Tuto stránku jsme nenašli.", body=NOTFOUND_PAGE, ld=[], noindex=True)

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
    meta_index = '<meta name="robots" content="noindex">' if p.get("noindex") else f'<link rel="canonical" href="{url}">'
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
{ld_scripts(p.get("ld", []))}
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
    (DIST / ".htaccess").write_text(htaccess(), encoding="utf-8")
    (DIST / "_redirects").write_text(redirects(), encoding="utf-8")
    (DIST / ".assetsignore").write_text(".htaccess\n", encoding="utf-8")   # Cloudflare ho nepotřebuje

    if preview:
        router = (SRC / "preview.js").read_text(encoding="utf-8")
        (ROOT / "preview.html").write_text(preview_document(css, js, router), encoding="utf-8")

    forms = "Web3Forms" if WEB3FORMS_KEY else "e-mail (bez klíče Web3Forms)"
    print(f"Hotovo: {len(PAGES) + 1} stránek ve složce dist, verze {ver}, formuláře: {forms}")


if __name__ == "__main__":
    build(preview="--nahled" in sys.argv)
