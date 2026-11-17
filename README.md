# Web Cafe Moya

Zdroj webu www.cafemoya.cz. Stránky se generují skriptem `build.py` a po každé změně v tomto
repozitáři je Cloudflare sám sestaví a zveřejní.

## Nastavení projektu v Cloudflare (Workers)

V projektu otevřete Settings → Build a nastavte:

| Položka | Hodnota |
|---|---|
| Build command | `python3 build.py` |
| Deploy command | `npx wrangler deploy` |
| Proměnná prostředí (volitelně) | `WEB3FORMS_KEY` = klíč z web3forms.com |

Soubor `wrangler.jsonc` říká Cloudflaru, že hotový web je ve složce `dist`. Jméno v něm (`cafemoya`)
musí odpovídat názvu projektu v Cloudflare. Soubor index.html a ostatní stránky v repozitáři nehledejte,
vznikají až při sestavení.

## Co kde je

- `build.py` – texty všech stránek, ceny mís, kontakty, titulky a popisy pro vyhledávače
- `*.xlsx` – týdenní jídelní lístky, ze kterých se skládá polední menu (viz níže)
- `src/styles.css` – vzhled webu
- `src/site.js` – mobilní menu, zobrazení dnešního menu, odesílání formulářů
- `src/assets/` – písma, ikony a obrázek pro sdílení na sociálních sítích
- `src/preview.js` – jen pro náhled celého webu v jednom souboru

## Polední menu

Týdenní tabulku s menu (xlsx) stačí nahrát na hlavní stránku repozitáře: Add file → Upload files.
Web si z ní při sestavení přečte týden, dny, jídla, alergeny i ceny. Na názvu souboru nezáleží.

- Tabulka má stejný tvar jako pro Facebook: v hlavičce „47. týden 16.-20.11.2026“, ve sloupci B den,
  v C pořadí jídla, v D název, v E alergeny a v F cena. Polévka je na řádku s názvem dne.
- Den bez vaření stačí označit textem u dne, například „Státní svátek - nevaříme“.
- Menu na další týden můžete nahrát dopředu. Web ho ukáže, až skončí to současné.
- Opravenou tabulku nahrajte se stejným názvem, GitHub ji přepíše. Staré tabulky mazat nemusíte.
- Když nové menu chybí, úvodní stránka i stránka s menu napíšou, že ho připravujeme.

## Formuláře

Poptávka dortu, objednávka mísy a rezervace stolu se odesílají přes Web3Forms na info@cafemoya.cz,
jakmile je vyplněný klíč `WEB3FORMS_KEY` (v Cloudflare v nastavení proměnných, nebo přímo v `build.py`).
Bez klíče formulář jen otevře e-mail s vyplněnými údaji.

## Náhled na vlastním počítači

```
python3 build.py --nahled
```

Vznikne složka `dist` s hotovým webem a soubor `preview.html`, který stačí otevřít v prohlížeči.
