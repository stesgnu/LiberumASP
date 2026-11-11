/* Convierte todos los <select> de la aplicacion en listas con buscador.
 * - El <select> original se conserva (oculto): el formulario, los onchange y
 *   los scripts existentes siguen funcionando igual.
 * - Se omiten: select multiple, listas con size>1, selects con solo numeros
 *   (fechas) y los que tengan el atributo data-nosearch.
 */
(function () {
  if (window.__lhdSearchSelect) return;
  window.__lhdSearchSelect = true;

  function norm(s) {
    s = (s || '').toLowerCase();
    return s.normalize ? s.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : s;
  }
  function esc(s) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  function eligible(sel) {
    if (sel.getAttribute('data-nosearch') !== null) return false;
    if (sel.getAttribute('data-lhd-ss')) return false;
    if (sel.multiple || sel.size > 1) return false;
    if (sel.options.length < 2) return false;
    var allNum = true;
    for (var i = 0; i < sel.options.length; i++) {
      if (!/^\s*-?\d+\s*$/.test(sel.options[i].text)) { allNum = false; break; }
    }
    return !allNum;
  }

  function enhance(sel) {
    sel.setAttribute('data-lhd-ss', '1');

    var wrap = document.createElement('span');
    wrap.className = 'lhd-ss';
    var w = sel.offsetWidth;
    wrap.style.minWidth = (w > 0 ? Math.max(w, 140) : 180) + 'px';
    if (sel.style.width) wrap.style.width = sel.style.width;
    if (sel.className.indexOf('form-control') > -1) wrap.className += ' lhd-ss-block';

    wrap.style.position = 'relative';
    wrap.style.display = (sel.className.indexOf('form-control') > -1) ? 'block' : 'inline-block';
    wrap.style.verticalAlign = 'middle';
    wrap.style.boxSizing = 'border-box';

    var input = document.createElement('input');
    input.type = 'text';
    input.className = 'lhd-ss-input';
    input.style.cssText = 'width:100%;box-sizing:border-box;padding:7px 28px 7px 10px;cursor:text;background:#fff;color:#1a202c;border:1px solid #cbd5e0;border-radius:6px;font-size:14px;';
    input.setAttribute('autocomplete', 'off');
    if (sel.style.fontSize) input.style.fontSize = sel.style.fontSize;

    var arrow = document.createElement('span');
    arrow.className = 'lhd-ss-arrow';
    arrow.style.cssText = 'position:absolute;right:10px;top:50%;margin-top:-9px;color:#718096;cursor:pointer;line-height:18px;';
    arrow.innerHTML = '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg>';

    var list = document.createElement('ul');
    list.className = 'lhd-ss-list';
    list.style.cssText = 'display:none;position:absolute;z-index:2147483000;left:0;top:100%;min-width:100%;max-height:240px;overflow-y:auto;margin:4px 0 0;padding:4px;list-style:none;background:#fff;border:1px solid #cbd5e0;border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.18);box-sizing:border-box;text-align:left;';

    sel.parentNode.insertBefore(wrap, sel);
    wrap.appendChild(sel);
    wrap.appendChild(input);
    wrap.appendChild(arrow);
    // La lista vive en <body>: asi ningun contenedor (overflow:hidden, z-index, etc.) puede recortarla ni taparla
    document.body.appendChild(list);
    sel.style.display = 'none';

    var active = -1, lis = [];

    function currentText() {
      var o = sel.options[sel.selectedIndex];
      return o ? o.text : '';
    }
    function syncInput() {
      if (document.activeElement !== input) input.value = currentText();
      input.disabled = sel.disabled;
      wrap.className = wrap.className.replace(' lhd-ss-off', '') + (sel.disabled ? ' lhd-ss-off' : '');
    }
    function setActive(n) {
      var v = [], i;
      for (i = 0; i < lis.length; i++) { lis[i].className = lis[i].className.replace(' active', ''); lis[i].style.background = '#fff'; if (lis[i].style.display !== 'none') v.push(lis[i]); }
      if (!v.length) { active = -1; return; }
      active = (n + v.length) % v.length;
      v[active].className += ' active';
      v[active].style.background = '#e8f0fb';
      if (v[active].scrollIntoView) v[active].scrollIntoView({ block: 'nearest' });
    }
    function build() {
      list.innerHTML = '';
      lis = [];
      for (var i = 0; i < sel.options.length; i++) {
        var o = sel.options[i];
        var li = document.createElement('li');
        li.setAttribute('data-i', i);
        li.setAttribute('data-txt', o.text);
        li.className = (i === sel.selectedIndex ? 'sel' : '') + (o.disabled ? ' dis' : '');
        li.style.cssText = 'display:block;margin:0;padding:7px 10px;border-radius:4px;cursor:pointer;font-size:14px;color:#1a202c;background:#fff;white-space:nowrap;list-style:none;' + (i === sel.selectedIndex ? 'font-weight:600;' : '');
        li.innerHTML = esc(o.text) || '&nbsp;';
        li.onmouseover = (function (el) { return function () {
          var v = [], k;
          for (k = 0; k < lis.length; k++) if (lis[k].style.display !== 'none') v.push(lis[k]);
          for (k = 0; k < v.length; k++) if (v[k] === el) { setActive(k); break; }
        }; })(li);
        list.appendChild(li);
        lis.push(li);
      }
      var empty = document.createElement('li');
      empty.className = 'lhd-ss-empty';
      empty.textContent = 'Sin coincidencias';
      empty.style.cssText = 'display:none;padding:7px 10px;color:#718096;list-style:none;';
      list.appendChild(empty);
      list._empty = empty;
    }
    function filter(showAll) {
      var q = showAll ? '' : norm(input.value), shown = 0;
      for (var i = 0; i < lis.length; i++) {
        var txt = lis[i].getAttribute('data-txt');
        var pos = q ? norm(txt).indexOf(q) : 0;
        if (pos > -1) {
          shown++;
          lis[i].style.display = '';
          lis[i].innerHTML = (q ? esc(txt.substr(0, pos)) + '<mark>' + esc(txt.substr(pos, q.length)) + '</mark>' + esc(txt.substr(pos + q.length)) : esc(txt)) || '&nbsp;';
        } else {
          lis[i].style.display = 'none';
        }
      }
      list._empty.style.display = shown ? 'none' : 'block';
      setActive(0);
    }
    // La lista usa position:fixed para no quedar recortada por contenedores con
    // overflow:hidden (tablas, tarjetas). Se recalcula al abrir, filtrar y hacer scroll.
    function placeList() {
      var rc = input.getBoundingClientRect();
      var vh = window.innerHeight || document.documentElement.clientHeight;
      var below = vh - rc.bottom - 12;
      var above = rc.top - 12;
      var maxH = 240;
      list.style.position = 'fixed';
      list.style.left = rc.left + 'px';
      list.style.minWidth = rc.width + 'px';
      list.style.right = 'auto';
      list.style.width = 'auto';
      list.style.margin = '0';
      if (below < 140 && above > below) {          // poco espacio abajo: abrir hacia arriba
        list.style.top = 'auto';
        list.style.bottom = (vh - rc.top + 4) + 'px';
        list.style.maxHeight = Math.min(maxH, above) + 'px';
      } else {
        list.style.bottom = 'auto';
        list.style.top = (rc.bottom + 4) + 'px';
        list.style.maxHeight = Math.min(maxH, Math.max(120, below)) + 'px';
      }
    }
    function onReposition() { if (wrap.className.indexOf(' open') > -1) placeList(); }
    window.addEventListener('scroll', onReposition, true);
    window.addEventListener('resize', onReposition);

    function open() {
      if (sel.disabled) return;
      build();
      wrap.className = wrap.className.replace(' open', '') + ' open';
      wrap.style.zIndex = '100000';
      list.style.display = 'block';
      placeList();
      filter(true);
      // posicionar sobre la opcion elegida
      for (var i = 0; i < lis.length; i++) if (lis[i].className.indexOf('sel') > -1) { if (lis[i].scrollIntoView) lis[i].scrollIntoView({ block: 'nearest' }); break; }
    }
    function close() {
      wrap.className = wrap.className.replace(' open', '');
      wrap.style.zIndex = '';
      list.style.display = 'none';
      input.value = currentText();
    }
    function choose(li) {
      var i = parseInt(li.getAttribute('data-i'), 10);
      if (sel.options[i] && sel.options[i].disabled) return;
      var changed = sel.selectedIndex !== i;
      sel.selectedIndex = i;
      close();
      if (changed) {
        var ev;
        if (typeof Event === 'function') { ev = new Event('change', { bubbles: true }); }
        else { ev = document.createEvent('Event'); ev.initEvent('change', true, true); }
        sel.dispatchEvent(ev);
      }
    }

    input.onfocus = function () { input.select(); open(); };
    input.onclick = function () { if (wrap.className.indexOf(' open') < 0) open(); };
    arrow.onmousedown = function (e) { e.preventDefault(); if (wrap.className.indexOf(' open') < 0) { input.focus(); } else { input.blur(); } };
    input.oninput = function () {
      if (wrap.className.indexOf(' open') < 0) open();
      filter(false);
      placeList();
    };
    input.onkeydown = function (e) {
      var k = e.keyCode;
      if (k === 40) { if (wrap.className.indexOf(' open') < 0) open(); else setActive(active + 1); e.preventDefault(); }
      else if (k === 38) { setActive(active - 1); e.preventDefault(); }
      else if (k === 13) {
        if (wrap.className.indexOf(' open') > -1) {
          var v = [];
          for (var i = 0; i < lis.length; i++) if (lis[i].style.display !== 'none') v.push(lis[i]);
          if (v[active]) choose(v[active]);
          input.blur();
          e.preventDefault();   // Enter no envia el formulario mientras se elige
        }
      }
      else if (k === 27) { input.blur(); }
    };
    input.onblur = function () { setTimeout(close, 120); };
    list.onmousedown = function (e) {
      e = e || window.event;
      var t = e.target || e.srcElement;
      while (t && t !== list && t.tagName !== 'LI') t = t.parentNode;
      if (t && t.tagName === 'LI' && t.getAttribute('data-i') !== null) { choose(t); e.preventDefault(); input.blur(); }
      else e.preventDefault();
    };

    // Mantener sincronizado si otro script cambia el valor o el estado
    sel.addEventListener('change', syncInput);
    if (sel.form) sel.form.addEventListener('reset', function () { setTimeout(syncInput, 0); });
    setInterval(syncInput, 400);
    syncInput();
  }

  function init() {
    var all = document.getElementsByTagName('select'), arr = [], i;
    for (i = 0; i < all.length; i++) arr.push(all[i]);
    for (i = 0; i < arr.length; i++) if (eligible(arr[i])) enhance(arr[i]);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
