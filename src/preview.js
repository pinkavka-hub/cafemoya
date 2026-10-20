/* Jen pro jednosouborový náhled: přepínání stránek podle #/adresy/ */
(function () {
  var pages = [].slice.call(document.querySelectorAll('.page'));
  var links = [].slice.call(document.querySelectorAll('.nav ul a'));
  var nav = document.getElementById('hlavni-menu');
  var toggle = document.querySelector('.nav-toggle');

  function show(route) {
    var target = pages.filter(function (p) { return p.getAttribute('data-route') === route; })[0] || pages[0];
    pages.forEach(function (p) { p.hidden = p !== target; });
    document.title = target.getAttribute('data-title');
    links.forEach(function (a) {
      if (a.getAttribute('href') === '#' + target.getAttribute('data-route')) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    });
    if (nav && toggle) { nav.classList.remove('is-open'); toggle.setAttribute('aria-expanded', 'false'); }
    return target;
  }

  function route() {
    var h = location.hash;
    if (!h || h === '#' || h.indexOf('#/') === 0) {
      show(h && h.length > 1 ? h.slice(1) : '/');
      window.scrollTo(0, 0);
      return;
    }
    var el = document.getElementById(decodeURIComponent(h.slice(1)));
    if (!el) return;
    var page = el.closest('.page');
    if (page && page.hidden) show(page.getAttribute('data-route'));
    el.scrollIntoView();
  }

  window.addEventListener('hashchange', route);
  route();

  /* Vyzkoušení poledního bloku v různých časech (jen v náhledu, na webu není) */
  var dates = [].slice.call(document.querySelectorAll('[data-day-block]')).map(function (b) {
    return b.getAttribute('data-date');
  });
  if (!dates.length || !window.moyaUpdateLunch) return;
  function plus(iso, n) {
    var a = iso.split('-');
    return new Date(Date.UTC(+a[0], +a[1] - 1, +a[2]) + n * 86400000).toISOString().slice(0, 10);
  }
  var mon = dates[0];
  var sims = [['Teď', null], ['Po 10:00', [mon, 600]], ['Po 15:00', [mon, 900]],
              ['Pá 15:00', [plus(mon, 4), 900]], ['So 10:00', [plus(mon, 5), 600]]];
  var panel = document.createElement('div');
  panel.className = 'sim';
  panel.setAttribute('role', 'group');
  panel.setAttribute('aria-label', 'Vyzkoušet polední blok v jiném čase');
  panel.innerHTML = '<span>Vyzkoušet čas:</span>';
  sims.forEach(function (s) {
    var b = document.createElement('button');
    b.type = 'button';
    b.textContent = s[0];
    b.setAttribute('aria-pressed', s[1] ? 'false' : 'true');
    b.addEventListener('click', function () {
      window.MOYA_NOW = s[1] ? { date: s[1][0], minutes: s[1][1] } : null;
      [].forEach.call(panel.querySelectorAll('button'), function (x) { x.setAttribute('aria-pressed', String(x === b)); });
      window.moyaUpdateLunch();
      if (location.hash && location.hash !== '#/') location.hash = '#/';
      setTimeout(function () {
        var sec = document.querySelector('[data-lunch-section]:not([hidden])');
        if (sec) sec.scrollIntoView();
      }, 60);
    });
    panel.appendChild(b);
  });
  document.body.appendChild(panel);
})();
