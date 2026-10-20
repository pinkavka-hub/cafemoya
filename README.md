# Web Cafe Moya

Zdroj webu www.cafemoya.cz. Stránky se generují skriptem `build.py` a po každé změně v tomto
repozitáři je Cloudflare Pages sám sestaví a zveřejní.

## Nastavení projektu v Cloudflare Pages

| Položka | Hodnota |
|---|---|
| Framework preset | None |
| Build command | `python3 build.py` |
| Build output directory | `dist` |
| Proměnná prostředí (volitelně) | `WEB3FORMS_KEY` = klíč z web3forms.com |

## Co kde je

- `build.py` – texty všech stránek, polední menu (`MENU`), ceny mís, kontakty, titulky a popisy pro vyhledávače
- `src/styles.css` – vzhled webu
- `src/site.js` – mobilní menu, zobrazení dnešního menu, odesílání formulářů
- `src/assets/` – písma, ikony a obrázek pro sdílení na sociálních sítích
- `src/preview.js` – jen pro náhled celého webu v jednom souboru

## Formuláře

Poptávka dortu, objednávka mísy a rezervace stolu se odesílají přes Web3Forms na info@cafemoya.cz,
jakmile je vyplněný klíč `WEB3FORMS_KEY` (v Cloudflare v nastavení proměnných, nebo přímo v `build.py`).
Bez klíče formulář jen otevře e-mail s vyplněnými údaji.

## Náhled na vlastním počítači

```
python3 build.py --nahled
```

Vznikne složka `dist` s hotovým webem a soubor `preview.html`, který stačí otevřít v prohlížeči.
