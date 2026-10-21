/* Cafe Moya – drobné skripty webu */
(function () {
  var d = document;
  d.documentElement.classList.add('js');

  /* Mobilní menu */
  var toggle = d.querySelector('.nav-toggle');
  var nav = d.getElementById('hlavni-menu');
  if (toggle && nav) {
    toggle.addEventListener('click', function () {
      var open = toggle.getAttribute('aria-expanded') === 'true';
      toggle.setAttribute('aria-expanded', String(!open));
      nav.classList.toggle('is-open', !open);
    });
  }

  /* Český čas bez ohledu na to, odkud se návštěvník dívá */
  function pragueNow() {
    if (window.MOYA_NOW) return window.MOYA_NOW;
    var o = {};
    new Intl.DateTimeFormat('en-GB', {
      timeZone: 'Europe/Prague', year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
    }).formatToParts(new Date()).forEach(function (p) { o[p.type] = p.value; });
    return { date: o.year + '-' + o.month + '-' + o.day, minutes: Number(o.hour) * 60 + Number(o.minute) };
  }
  function utc(iso) { var a = iso.split('-'); return Date.UTC(+a[0], +a[1] - 1, +a[2]); }
  function weekday(iso) { return new Date(utc(iso)).getUTCDay(); }          /* 0 = neděle */
  function addDays(iso, n) { return new Date(utc(iso) + n * 86400000).toISOString().slice(0, 10); }
  var DAY_IN = ['V neděli', 'V pondělí', 'V úterý', 'Ve středu', 'Ve čtvrtek', 'V pátek', 'V sobotu'];

  /* Úvodní stránka: polední menu podle dne a hodiny
     Po–Pá do konce výdeje  → dnešní menu
     po skončení výdeje      → menu na další den, kdy vaříme
     Pá odpoledne a víkend   → víkendový blok (a pondělní menu, pokud už je zveřejněné) */
  var weekdaySec = d.querySelector('[data-lunch-section="weekday"]');
  var weekendSec = d.querySelector('[data-lunch-section="weekend"]');
  if (weekdaySec && weekendSec) {
    var blocks = [].slice.call(d.querySelectorAll('[data-day-block]'));
    var slotWeekday = weekdaySec.querySelector('[data-menu-slot]');
    var slotWeekend = weekendSec.querySelector('[data-menu-slot]');
    var photo = weekendSec.querySelector('[data-weekend-photo]');
    var fallback = d.querySelector('[data-day-fallback]');
    var end = (weekdaySec.getAttribute('data-lunch-end') || '14:00').split(':');
    var endMin = Number(end[0]) * 60 + Number(end[1]);

    var update = function () {
      var now = pragueNow();
      var today = now.date;
      var wd = weekday(today);
      var servingToday = blocks.filter(function (b) {
        return b.getAttribute('data-date') === today;
      })[0];
      var pick = null;
      var label = '';

      if (servingToday && now.minutes < endMin) {
        pick = servingToday;
        label = 'Dnes v poledne';
      } else {
        pick = blocks.filter(function (b) { return b.getAttribute('data-date') > today; })[0] || null;
        if (pick) {
          var next = pick.getAttribute('data-date');
          label = next === addDays(today, 1) ? 'Zítra vaříme' : DAY_IN[weekday(next)] + ' vaříme';
        }
      }

      var weekend = wd === 6 || wd === 0 || (wd === 5 && now.minutes >= endMin);
      var showWeekend = weekend && pick !== servingToday;
      weekdaySec.hidden = showWeekend;
      weekendSec.hidden = !showWeekend;

      blocks.forEach(function (b) { b.hidden = b !== pick; });
      if (pick) {
        (showWeekend ? slotWeekend : slotWeekday).appendChild(pick);
        var when = pick.querySelector('[data-when]');
        if (when) { when.textContent = label; when.className = 'today-badge'; }
      }
      if (photo) photo.hidden = !!pick;
      if (fallback) fallback.hidden = showWeekend || !!pick;
    };

    update();
    setInterval(update, 5 * 60 * 1000);
    window.moyaUpdateLunch = update;
  }

  /* Stránka s menu: označit dnešní den */
  var day = d.querySelector('[data-menu-date="' + pragueNow().date + '"]');
  if (day) {
    var badge = day.querySelector('.today-badge');
    if (badge) badge.hidden = false;
  }

  /* Formuláře: poptávka dortu, objednávka mísy, rezervace stolu */
  [].slice.call(d.querySelectorAll('form[data-form]')).forEach(function (form) {
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      if (!form.reportValidity()) return;
      var msg = form.querySelector('.form-msg');

      if (window.MOYA_PREVIEW) {
        show(msg, 'Ukázka: na hotovém webu se formulář odešle na info@cafemoya.cz.');
        return;
      }

      var endpoint = form.getAttribute('data-endpoint');
      if (endpoint) {
        fetch(endpoint, { method: 'POST', body: new FormData(form), headers: { Accept: 'application/json' } })
          .then(function (r) {
            if (!r.ok) throw new Error('send');
            show(msg, 'Děkujeme, máme to. Brzy se vám ozveme.');
            form.reset();
          })
          .catch(function () {
            show(msg, 'Odeslání se nepovedlo. Zavolejte nám prosím na 605 226 781.');
          });
        return;
      }

      /* Bez odesílací služby: otevřít e-mail s vyplněnými údaji */
      var lines = [];
      [].slice.call(form.elements).forEach(function (el) {
        if (!el.name || !el.value) return;
        var label = form.querySelector('label[for="' + el.id + '"]');
        lines.push((label ? label.textContent.replace('*', '').trim() : el.name) + ': ' + el.value);
      });
      window.location.href = 'mailto:info@cafemoya.cz?subject=' +
        encodeURIComponent(form.getAttribute('data-subject') || 'Zpráva z webu') +
        '&body=' + encodeURIComponent(lines.join('\n'));
      show(msg, 'Otevřeli jsme vám e-mail s vyplněnými údaji. Stačí ho odeslat.');
    });
  });

  function show(el, text) {
    if (!el) return;
    el.textContent = text;
    el.hidden = false;
  }
})();
