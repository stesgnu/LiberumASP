/* Notificaciones de la aplicacion, basadas en Toastify JS (js/toastify/).
 * Uso: LhdToast.success('Ticket guardado'); LhdToast.error('...'); .warning(); .info();
 * Desde ASP: Call LhdToast("success", "mensaje")  (ver public.asp)
 */
(function (window, document) {
  'use strict';

  var COLORS = {
    success: ['#1A7F45', '#2FA463'],
    error:   ['#B02A37', '#D6453A'],
    warning: ['#B7791F', '#D69E2E'],
    info:    ['#0057B8', '#0082C8']
  };
  var ICONS = { success: '✔', error: '✖', warning: '⚠', info: 'ℹ' };

  function show(type, message, opts) {
    type = COLORS[type] ? type : 'info';
    opts = opts || {};
    if (typeof window.Toastify !== 'function') {
      // Toastify no cargo: no perder el mensaje
      if (type === 'error' || type === 'warning') { window.alert(message); }
      return;
    }
    var node = document.createElement('div');
    node.className = 'lhd-toast-body';
    var ic = document.createElement('span');
    ic.className = 'lhd-toast-icon';
    ic.textContent = ICONS[type];
    var tx = document.createElement('span');
    tx.textContent = message;   // textContent: el mensaje nunca se interpreta como HTML
    node.appendChild(ic);
    node.appendChild(tx);

    window.Toastify({
      node: node,
      duration: opts.duration || (type === 'error' || type === 'warning' ? 7000 : 4000),
      gravity: 'top',
      position: 'right',
      close: true,
      stopOnFocus: true,
      style: { background: 'linear-gradient(135deg,' + COLORS[type][0] + ',' + COLORS[type][1] + ')' },
      className: 'lhd-toast'
    }).showToast();
  }

  window.LhdToast = {
    show: show,
    success: function (m, o) { show('success', m, o); },
    error:   function (m, o) { show('error', m, o); },
    warning: function (m, o) { show('warning', m, o); },
    info:    function (m, o) { show('info', m, o); }
  };

  // Mensajes dejados por el servidor (cola impresa por DisplayFooter)
  var q = window.__lhdToastQueue || [];
  for (var i = 0; i < q.length; i++) { show(q[i][0], q[i][1]); }
  window.__lhdToastQueue = [];
})(window, document);
