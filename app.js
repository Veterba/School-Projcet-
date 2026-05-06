// HMS — small enhancements only. Page works fully without JS.
(function () {
  'use strict';

  var doc = document.documentElement;
  var prefersReduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  // -------------------------------------------------------------------------
  // Mobile drawer toggle + dropdowns
  // -------------------------------------------------------------------------
  var navToggle = document.querySelector('.nav-toggle');
  var nav = document.querySelector('.nav');

  if (navToggle && nav) {
    navToggle.addEventListener('click', function () {
      var open = nav.getAttribute('data-open') === 'true';
      nav.setAttribute('data-open', String(!open));
      navToggle.setAttribute('aria-expanded', String(!open));
      navToggle.setAttribute('aria-label', !open ? 'Lukk meny' : 'Åpne meny');
    });
  }

  var groups = document.querySelectorAll('.nav__group');
  var isMobile = function () { return window.matchMedia('(max-width: 767.98px)').matches; };

  groups.forEach(function (group) {
    var toggle = group.querySelector('.nav__group-toggle');
    if (!toggle) return;
    toggle.addEventListener('click', function () {
      // On desktop the dropdown opens on hover/focus; on mobile we toggle.
      if (!isMobile()) return;
      var open = group.getAttribute('data-open') === 'true';
      group.setAttribute('data-open', String(!open));
      toggle.setAttribute('aria-expanded', String(!open));
    });
  });

  // -------------------------------------------------------------------------
  // Theme toggle (manual override on top of prefers-color-scheme)
  // -------------------------------------------------------------------------
  var themeBtn = document.querySelector('.theme-toggle');
  if (themeBtn) {
    themeBtn.addEventListener('click', function () {
      var current = doc.getAttribute('data-theme');
      var systemDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
      var resolved = current || (systemDark ? 'dark' : 'light');
      var next = resolved === 'dark' ? 'light' : 'dark';
      doc.setAttribute('data-theme', next);
      try { localStorage.setItem('hms-theme', next); } catch (e) {}
    });
  }

  // -------------------------------------------------------------------------
  // Auto-build TOC from the current article's <h2 id="...">
  // -------------------------------------------------------------------------
  var tocList = document.getElementById('toc-list');
  var article = document.querySelector('.prose');
  if (tocList && article) {
    var headings = article.querySelectorAll('h2[id], section[id] > h2');
    var items = [];
    headings.forEach(function (h2) {
      var id = h2.id || (h2.parentElement && h2.parentElement.id);
      if (!id) return;
      var li = document.createElement('li');
      var a = document.createElement('a');
      a.href = '#' + id;
      a.textContent = h2.textContent.trim();
      a.dataset.tocLink = id;
      li.appendChild(a);
      tocList.appendChild(li);
      items.push({ id: id, link: a });
    });

    // Scroll-spy: highlight the closest section in view
    if (items.length && 'IntersectionObserver' in window) {
      var observer = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
          if (!entry.isIntersecting) return;
          var id = entry.target.id;
          items.forEach(function (it) {
            it.link.dataset.active = String(it.id === id);
          });
        });
      }, { rootMargin: '-20% 0px -70% 0px', threshold: 0 });

      items.forEach(function (it) {
        var section = document.getElementById(it.id);
        if (section) observer.observe(section);
      });
    }
  }

  // -------------------------------------------------------------------------
  // Fade-in on scroll (gated on reduced-motion)
  // -------------------------------------------------------------------------
  if (!prefersReduced && 'IntersectionObserver' in window) {
    var fades = document.querySelectorAll('.fade-in');
    if (fades.length) {
      var fadeObs = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            entry.target.classList.add('is-visible');
            fadeObs.unobserve(entry.target);
          }
        });
      }, { rootMargin: '0px 0px -10% 0px', threshold: 0.1 });
      fades.forEach(function (el) { fadeObs.observe(el); });
    }
  } else {
    // Reduced motion → reveal everything immediately
    document.querySelectorAll('.fade-in').forEach(function (el) {
      el.classList.add('is-visible');
    });
  }

  // -------------------------------------------------------------------------
  // Year stamp in footer
  // -------------------------------------------------------------------------
  var year = document.getElementById('year');
  if (year) year.textContent = String(new Date().getFullYear());
})();
