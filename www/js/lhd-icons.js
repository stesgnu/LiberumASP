/* Iconos de la aplicacion.
 * Las vistas usan emojis / simbolos Unicode como iconos (p. ej. &#128203;). Si el equipo
 * no tiene una fuente con esos glifos, el navegador dibuja solo un cuadro. Este script
 * los sustituye por iconos SVG en linea (no dependen de ninguna fuente, heredan el color
 * del texto y se ajustan al tamano de letra), tanto en el HTML inicial como en el
 * contenido que se agregue despues con JavaScript.
 */
(function (window, document) {
  'use strict';
  if (window.__lhdIcons) { return; }
  window.__lhdIcons = true;

  var P = {
    clip:   '<rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/>',
    plus:   '<path d="M12 5v14M5 12h14"/>',
    check:  '<path d="M20 6 9 17l-5-5"/>',
    search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
    user:   '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
    users:  '<circle cx="9" cy="8" r="4"/><path d="M2 21a7 7 0 0 1 14 0"/><path d="M16 4a4 4 0 0 1 0 8M22 21a7 7 0 0 0-5-6.7"/>',
    gear:   '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M4.9 19.1 7 17M17 7l2.1-2.1"/>',
    warn:   '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
    book:   '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/>',
    pencil: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/>',
    lock:   '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
    note:   '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h8M8 9h2"/>',
    refresh:'<path d="M21 12a9 9 0 1 1-3-6.7L21 8"/><path d="M21 3v5h-5"/>',
    chart:  '<path d="M3 3v18h18"/><path d="M7 15V9M12 15V5M17 15v-3"/>',
    save:   '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><path d="M17 21v-8H7v8M7 3v5h8"/>',
    eye:    '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z"/><circle cx="12" cy="12" r="3"/>',
    x:      '<path d="M18 6 6 18M6 6l12 12"/>',
    timer:  '<circle cx="12" cy="13" r="8"/><path d="M12 9v4l2 2M9 2h6"/>',
    info:   '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
    print:  '<path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/>',
    file:   '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/>',
    folder: '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
    okc:    '<circle cx="12" cy="12" r="10"/><path d="m8 12 3 3 5-6"/>',
    logout: '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="m16 17 5-5-5-5M21 12H9"/>',
    chat:   '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    upload: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5M12 3v12"/>',
    home:   '<path d="m3 11 9-8 9 8"/><path d="M5 10v10h14V10"/>',
    trophy: '<path d="M8 21h8M12 17v4M7 4h10v5a5 5 0 0 1-10 0z"/><path d="M17 5h3v2a3 3 0 0 1-3 3M7 5H4v2a3 3 0 0 0 3 3"/>',
    zap:    '<path d="M13 2 3 14h9l-1 8 10-12h-9z"/>',
    menu:   '<path d="M3 6h18M3 12h18M3 18h18"/>',
    trash:  '<path d="M3 6h18"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>',
    clock:  '<circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/>',
    wrench: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18v3h3l6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.5-.5-.5-2.5z"/>',
    flame:  '<path d="M12 2s5 4 5 9a5 5 0 0 1-10 0c0-2 1-3 2-4 0 2 1 3 2 3 0-3-1-5 1-8z"/>',
    clipx:  '<path d="m21 11-9.5 9.5a5 5 0 0 1-7-7L14 4a3.5 3.5 0 0 1 5 5L9.5 18.5a2 2 0 0 1-3-3L15 7"/>',
    palette:'<circle cx="13.5" cy="6.5" r="1"/><circle cx="17.5" cy="10.5" r="1"/><circle cx="8.5" cy="7.5" r="1"/><circle cx="6.5" cy="12.5" r="1"/><path d="M12 2a10 10 0 0 0 0 20c1.7 0 2-1 2-2s-1-1.5-1-3 1-2 3-2h2a4 4 0 0 0 4-4 10 10 0 0 0-10-9z"/>',
    chevron:'<path d="m6 9 6 6 6-6"/>',
    bell:   '<path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.7 21a2 2 0 0 1-3.4 0"/>',
    star:   '<path d="m12 2 3.1 6.3 6.9 1-5 4.9 1.2 6.8L12 17.8 5.8 21l1.2-6.8-5-4.9 6.9-1z"/>',
    mail:   '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>',
    phone:  '<path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 1.9.7 2.8a2 2 0 0 1-.5 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7a2 2 0 0 1 1.7 2z"/>',
    calendar:'<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
    tag:    '<path d="M20.6 13.4 13.4 20.6a2 2 0 0 1-2.8 0L2 12V2h10l8.6 8.6a2 2 0 0 1 0 2.8z"/><circle cx="7" cy="7" r="1"/>',
    link:   '<path d="M10 13a5 5 0 0 0 7.5.5l3-3a5 5 0 0 0-7-7l-1.7 1.7"/><path d="M14 11a5 5 0 0 0-7.5-.5l-3 3a5 5 0 0 0 7 7l1.7-1.7"/>',
    lightbulb:'<path d="M9 18h6M10 22h4M12 2a7 7 0 0 0-4 12.7c.6.5 1 1.3 1 2.3h6c0-1 .4-1.8 1-2.3A7 7 0 0 0 12 2z"/>',
    target: '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
    shield: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>',
    globe:  '<circle cx="12" cy="12" r="10"/><path d="M2 12h20M12 2a15 15 0 0 1 0 20M12 2a15 15 0 0 0 0 20"/>',
    arrowr: '<path d="M5 12h14M12 5l7 7-7 7"/>',
    arrowl: '<path d="M19 12H5M12 19l-7-7 7-7"/>',
    updown: '<path d="M7 15l5 5 5-5M7 9l5-5 5 5"/>',
    up:     '<path d="m18 15-6-6-6 6"/>',
    down:   '<path d="m6 9 6 6 6-6"/>',
    inbox:  '<path d="M22 12h-6l-2 3h-4l-2-3H2"/><path d="M5.5 5.1 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.5-6.9A2 2 0 0 0 16.7 4H7.3a2 2 0 0 0-1.8 1.1z"/>'
  };

  // codigo Unicode -> icono
  var MAP = {
    0x1F4CB: 'clip', 0x2795: 'plus', 0x2714: 'check', 0x2705: 'okc', 0x1F50D: 'search', 0x1F50E: 'search',
    0x1F464: 'user', 0x1F465: 'users', 0x2699: 'gear', 0x26A0: 'warn', 0x1F4DA: 'book', 0x1F4D6: 'book',
    0x270E: 'pencil', 0x270F: 'pencil', 0x1F512: 'lock', 0x1F510: 'lock', 0x1F4DD: 'note', 0x1F504: 'refresh', 0x21BB: 'refresh',
    0x1F4CA: 'chart', 0x1F4C8: 'chart', 0x1F4BE: 'save', 0x1F441: 'eye', 0x2715: 'x', 0x2716: 'x', 0x274C: 'x',
    0x23F1: 'timer', 0x2139: 'info', 0x1F5A8: 'print', 0x1F4C4: 'file', 0x1F4C3: 'file',
    0x1F4C2: 'folder', 0x1F4C1: 'folder', 0x1F6AA: 'logout', 0x1F5E3: 'chat', 0x1F4AC: 'chat',
    0x1F4E4: 'upload', 0x1F3E0: 'home', 0x1F3C6: 'trophy', 0x26A1: 'zap', 0x2630: 'menu',
    0x1F5D1: 'trash', 0x1F550: 'clock', 0x1F557: 'clock', 0x1F527: 'wrench', 0x1F525: 'flame', 0x1F4CE: 'clipx',
    0x1F3A8: 'palette', 0x1F514: 'bell', 0x2B50: 'star', 0x1F4E7: 'mail', 0x2709: 'mail', 0x1F4DE: 'phone',
    0x1F4C5: 'calendar', 0x1F3F7: 'tag', 0x1F517: 'link', 0x1F4A1: 'lightbulb', 0x1F3AF: 'target',
    0x1F6E1: 'shield', 0x1F310: 'globe', 0x27A1: 'arrowr', 0x2192: 'arrowr', 0x2190: 'arrowl', 0x2195: 'updown',
    0x25B2: 'up', 0x25BC: 'down', 0x1F4E5: 'inbox'
  };

  var chars = [], k;
  for (k in MAP) { if (MAP.hasOwnProperty(k)) { chars.push(String.fromCodePoint(parseInt(k, 10))); } }
  var RE = new RegExp('(' + chars.join('|') + ')[\\uFE0E\\uFE0F]?', 'g');
  var RE_TEST = new RegExp('(' + chars.join('|') + ')');
  var SKIP = { SCRIPT: 1, STYLE: 1, TEXTAREA: 1, NOSCRIPT: 1, SVG: 1, CODE: 1, PRE: 1 };

  function svg(name) {
    return '<svg viewBox="0 0 24 24" width="1.1em" height="1.1em" fill="none" stroke="currentColor" ' +
           'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">' +
           P[name] + '</svg>';
  }
  window.LhdIcon = svg;

  function skipEl(el) {
    while (el && el.nodeType === 1) {
      if (SKIP[el.tagName.toUpperCase()]) { return true; }
      var c = el.className;
      if (typeof c === 'string' && (c.indexOf('lhd-rich') > -1 || c.indexOf('ql-editor') > -1 || c.indexOf('lhd-ico') > -1)) { return true; }
      el = el.parentNode;
    }
    return false;
  }

  function convertText(node) {
    var txt = node.nodeValue;
    if (!txt || !RE_TEST.test(txt)) { return; }
    var parent = node.parentNode;
    if (!parent || skipEl(parent)) { return; }

    // <option> y similares no admiten HTML: se quita el simbolo
    if (parent.tagName === 'OPTION' || parent.tagName === 'TITLE') {
      node.nodeValue = txt.replace(RE, '').replace(/^\s+/, '');
      return;
    }
    var frag = document.createDocumentFragment(), last = 0, m;
    RE.lastIndex = 0;
    while ((m = RE.exec(txt)) !== null) {
      if (m.index > last) { frag.appendChild(document.createTextNode(txt.substring(last, m.index))); }
      var sp = document.createElement('span');
      sp.className = 'lhd-ico';
      sp.innerHTML = svg(MAP[m[1].codePointAt(0)]);
      frag.appendChild(sp);
      last = m.index + m[0].length;
    }
    if (last < txt.length) { frag.appendChild(document.createTextNode(txt.substring(last))); }
    parent.replaceChild(frag, node);
  }

  function convertInputs(root) {
    var els = root.querySelectorAll ? root.querySelectorAll('input[type=submit],input[type=button],input[type=reset]') : [];
    for (var i = 0; i < els.length; i++) {
      var v = els[i].value;
      if (v && RE_TEST.test(v)) { els[i].value = v.replace(RE, '').replace(/^\s+/, ''); }
    }
  }

  function scan(root) {
    if (!root) { return; }
    if (root.nodeType === 3) { convertText(root); return; }
    if (root.nodeType !== 1 || skipEl(root)) { return; }
    var w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false), list = [], n;
    while ((n = w.nextNode())) { if (RE_TEST.test(n.nodeValue)) { list.push(n); } }
    for (var i = 0; i < list.length; i++) { convertText(list[i]); }
    convertInputs(root);
  }

  function start() {
    scan(document.body);
    if (typeof MutationObserver === 'function') {
      new MutationObserver(function (muts) {
        for (var i = 0; i < muts.length; i++) {
          var a = muts[i].addedNodes;
          for (var j = 0; j < a.length; j++) { scan(a[j]); }
          if (muts[i].type === 'characterData') { convertText(muts[i].target); }
        }
      }).observe(document.body, { childList: true, subtree: true });
    }
  }

  var css = document.createElement('style');
  css.appendChild(document.createTextNode(
    '.lhd-ico{display:inline-flex;align-items:center;justify-content:center;vertical-align:-0.18em;line-height:1}' +
    '.lhd-ico svg{display:block}'));
  document.getElementsByTagName('head')[0].appendChild(css);

  if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', start); }
  else { start(); }
})(window, document);
