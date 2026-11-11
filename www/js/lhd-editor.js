/*
 * Liberum Help Desk - editor WYSIWYG y adjuntos para tickets
 * Requiere Quill 1.3.7 (cargado antes de este archivo).
 * Las imagenes y los archivos se suben a attach/upload.asp y quedan en la base de datos.
 */
(function (window, document) {
  'use strict';

  var MAX_BYTES_DEFAULT = 2000000;

  // 2000000 -> "1.9 MB"; 500000 -> "488 KB"
  function sizeLabel(bytes) {
    if (bytes >= 1048576) { return (Math.round(bytes / 1048576 * 10) / 10) + ' MB'; }
    return Math.max(1, Math.round(bytes / 1024)) + ' KB';
  }

  // Notificaciones con Toastify (js/lhd-toast.js); alert() solo si no cargo.
  function notify(type, msg) {
    if (window.LhdToast) { window.LhdToast.show(type, msg); } else { window.alert(msg); }
  }

  var SPIN_CSS =
    '.lhd-spin{display:inline-block;width:14px;height:14px;border:2px solid #cbd5e1;' +
    'border-top-color:#2563eb;border-radius:50%;animation:lhd-spin .8s linear infinite;' +
    'vertical-align:middle;margin-right:6px}' +
    '@keyframes lhd-spin{to{transform:rotate(360deg)}}' +
    '.lhd-attach-row.is-error{color:#b91c1c}' +
    '.lhd-attach-status{font-size:13px;color:#1d4ed8;margin-bottom:6px;font-weight:600}' +
    '.lhd-dropzone{border:2px dashed #94a3b8;border-radius:10px;background:#f8fafc;padding:22px 16px;text-align:center;cursor:pointer;transition:all .15s ease;color:#475569}' +
    '.lhd-dropzone:hover{border-color:#0057b8;background:#eef4fc}' +
    '.lhd-dropzone.dz-drag-hover{border-color:#0057b8;background:#e0ecfb;box-shadow:0 0 0 4px rgba(0,87,184,.12)}' +
    '.lhd-dropzone .lhd-dz-icon{font-size:30px;line-height:1;margin-bottom:6px}' +
    '.lhd-dropzone .lhd-dz-title{font-weight:600;color:#1e293b}' +
    '.lhd-dropzone .lhd-dz-sub{font-size:12px;color:#64748b;margin-top:2px}' +
    '.lhd-dropzone .dz-message{margin:0}';
  var cssInjected = false;

  function injectCss() {
    if (cssInjected) { return; }
    cssInjected = true;
    var st = document.createElement('style');
    st.type = 'text/css';
    st.appendChild(document.createTextNode(SPIN_CSS));
    document.getElementsByTagName('head')[0].appendChild(st);
  }

  // Llama al servidor y devuelve el JSON. Si la respuesta no es JSON (error de
  // servidor), informa el codigo HTTP y un extracto del texto recibido.
  function req(url, method, body, cb) {
    fetch(url, { method: method, body: body, credentials: 'same-origin' })
      .then(function (r) {
        return r.text().then(function (txt) {
          var j = null;
          try { j = JSON.parse(txt); } catch (ex) { j = null; }
          if (j === null) {
            var extracto = txt.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().substring(0, 200);
            cb(null, { ok: false, error: 'Error del servidor (HTTP ' + r.status + '): ' + (extracto || 'respuesta vacia') });
            return;
          }
          cb(null, j);
        });
      })
      .catch(function (e) { cb(e, null); });
  }

  // Crea un editor WYSIWYG sobre un contenedor. El HTML se copia al campo oculto (textarea)
  // cada vez que cambia el contenido, para que el formulario lo envie.
  function createEditor(opts) {
    var container = document.getElementById(opts.editorId);
    var field = document.getElementById(opts.fieldId);
    if (!container || !field || typeof window.Quill === 'undefined') {
      return null;
    }
    var maxBytes = opts.maxBytes || MAX_BYTES_DEFAULT;

    function pickImage(quill) {
      var input = document.createElement('input');
      input.type = 'file';
      input.accept = 'image/png,image/jpeg,image/gif,image/webp';
      input.onchange = function () {
        var file = input.files && input.files[0];
        if (!file) { return; }
        if (file.size > maxBytes) {
          notify('error', 'La imagen supera el tamaño máximo permitido (' + sizeLabel(maxBytes) + ').');
          return;
        }
        var fd = new FormData();
        fd.append('file', file);
        req(opts.uploadUrl + '?token=' + encodeURIComponent(opts.token), 'POST', fd, function (err, j) {
          if (err || !j || !j.ok) {
            notify('error', (j && j.error) || 'No se pudo subir la imagen.');
            return;
          }
          var range = quill.getSelection(true);
          var index = range ? range.index : quill.getLength();
          quill.insertEmbed(index, 'image', j.url, 'user');
          quill.setSelection(index + 1, 0);
          notify('success', 'Imagen subida correctamente.');
        });
      };
      input.click();
    }

    var toolbar = [
      [{ header: [1, 2, 3, false] }],
      ['bold', 'italic', 'underline', 'strike'],
      [{ list: 'ordered' }, { list: 'bullet' }],
      [{ color: [] }, { background: [] }],
      ['blockquote', 'code-block', 'link', 'image'],
      ['clean']
    ];

    var quill = new window.Quill(container, {
      theme: 'snow',
      placeholder: opts.placeholder || '',
      modules: {
        toolbar: {
          container: toolbar,
          handlers: { image: function () { pickImage(quill); } }
        }
      }
    });

    if (field.value) {
      quill.clipboard.dangerouslyPasteHTML(field.value);
    }

    quill.on('text-change', function () {
      var html = quill.root.innerHTML;
      field.value = (html === '<p><br></p>') ? '' : html;
    });
    field.value = (quill.root.innerHTML === '<p><br></p>') ? '' : quill.root.innerHTML;
    return quill;
  }

  // Lista de adjuntos. Varios archivos se encolan y se suben uno a la vez, en el orden
  // seleccionado. Cada archivo muestra su estado (subiendo / listo / error) y se muestra
  // el progreso general. Mientras haya subidas pendientes no se permite enviar el formulario.
  function attachments(opts) {
    var input = document.getElementById(opts.fileInputId);
    var list = document.getElementById(opts.listId);
    var idsField = document.getElementById(opts.idsId);
    if (!input || !list || !idsField) { return; }
    var maxBytes = opts.maxBytes || MAX_BYTES_DEFAULT;
    var ids = [];
    var queue = [];
    var busy = false;
    var pending = 0;      // archivos encolados o subiendo
    var batchTotal = 0;   // archivos seleccionados en la tanda actual
    var batchDone = 0;    // archivos terminados (ok o error) en la tanda actual

    injectCss();
    var status = document.createElement('div');
    status.className = 'lhd-attach-status';
    status.style.display = 'none';
    status.setAttribute('role', 'status');
    list.parentNode.insertBefore(status, list);

    function refresh() {
      idsField.value = ids.join(',');
    }

    // Zona "arrastrar y soltar" (Dropzone). Dropzone solo captura los archivos:
    // la subida la hace la cola de esta funcion (enqueue/next), asi se conservan
    // el estado por archivo, las notificaciones y el bloqueo al enviar.
    function setupDropzone() {
      if (typeof window.Dropzone !== 'function') { return; }   // sin Dropzone: queda el input normal
      window.Dropzone.autoDiscover = false;
      var zone = document.createElement('div');
      zone.className = 'lhd-dropzone dropzone';
      zone.innerHTML =
        '<div class="dz-message">' +
        '<div class="lhd-dz-icon">&#128206;</div>' +
        '<div class="lhd-dz-title">Arrastra los archivos aquí o haz clic para seleccionarlos</div>' +
        '<div class="lhd-dz-sub">Puedes adjuntar varios archivos a la vez' +
        (opts.maxBytes ? ' (máx. ' + sizeLabel(opts.maxBytes || MAX_BYTES_DEFAULT) + ' c/u)' : '') +
        '</div></div>';
      input.parentNode.insertBefore(zone, input);
      input.style.display = 'none';

      var dz = new window.Dropzone(zone, {
        url: 'about:blank',
        autoProcessQueue: false,
        previewsContainer: false,
        clickable: true,
        uploadMultiple: false,
        maxFilesize: 100000,
        addRemoveLinks: false
      });
      dz.on('addedfile', function (file) {
        enqueue(file);
        // el archivo ya paso a nuestra cola; se libera de Dropzone
        setTimeout(function () { try { dz.removeFile(file); } catch (e) {} }, 0);
      });
    }

    function updateStatus() {
      if (pending > 0) {
        status.style.display = 'block';
        status.textContent = 'Adjuntando archivo ' + (batchDone + 1) + ' de ' + batchTotal + '…';
      } else {
        status.style.display = 'none';
        status.textContent = '';
      }
    }

    function sizeKb(bytes) {
      return Math.max(1, Math.round((bytes || 0) / 1024));
    }

    function addRemoveButton(row, onClick) {
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'btn btn-secondary btn-sm';
      btn.textContent = 'Quitar';
      btn.onclick = onClick;
      row.appendChild(btn);
    }

    function makeRow(label) {
      var row = document.createElement('div');
      row.className = 'lhd-attach-row';
      row.style.cssText = 'display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #e2e8f0;font-size:13px';
      row.nameEl = document.createElement('span');
      row.nameEl.style.flex = '1';
      row.appendChild(row.nameEl);
      list.appendChild(row);
      showUploading(row, label);
      return row;
    }

    function showUploading(row, label) {
      row.nameEl.textContent = '';
      var spin = document.createElement('span');
      spin.className = 'lhd-spin';
      row.nameEl.appendChild(spin);
      row.nameEl.appendChild(document.createTextNode(label + ' — subiendo…'));
    }

    function showDone(row, j, label) {
      row.nameEl.textContent = label + ' (' + sizeKb(j.size) + ' KB) — listo';
      addRemoveButton(row, function () {
        req(opts.uploadUrl + '?action=delete&id=' + j.id, 'POST', null, function () {
          ids = ids.filter(function (x) { return x !== j.id; });
          refresh();
          if (row.parentNode) { row.parentNode.removeChild(row); }
        });
      });
    }

    function showError(row, msg) {
      row.className += ' is-error';
      row.nameEl.textContent = msg;
      addRemoveButton(row, function () {
        if (row.parentNode) { row.parentNode.removeChild(row); }
      });
    }

    function enqueue(file) {
      if (file.size > maxBytes) {
        notify('error', '"' + file.name + '" (' + sizeLabel(file.size) + ') supera el tamaño máximo permitido (' + sizeLabel(maxBytes) + ').');
        return;
      }
      if (pending === 0) { batchTotal = 0; batchDone = 0; }
      pending++;
      batchTotal++;
      queue.push({ file: file, row: makeRow(file.name) });
      updateStatus();
      if (!busy) { next(); }
    }

    // Sube el siguiente archivo de la cola; al terminar pasa al siguiente.
    function next() {
      if (queue.length === 0) {
        busy = false;
        updateStatus();
        return;
      }
      busy = true;
      var item = queue.shift();
      var fd = new FormData();
      fd.append('file', item.file);
      req(opts.uploadUrl + '?token=' + encodeURIComponent(opts.token), 'POST', fd, function (err, j) {
        pending--;
        batchDone++;
        if (err || !j || !j.ok) {
          showError(item.row, '"' + item.file.name + '": ' + ((j && j.error) || 'no se pudo subir.'));
          notify('error', '"' + item.file.name + '": ' + ((j && j.error) || 'no se pudo subir.'));
        } else {
          ids.push(j.id);
          refresh();
          showDone(item.row, j, j.name || item.file.name);
          notify('success', 'Archivo adjuntado: ' + (j.name || item.file.name));
        }
        updateStatus();
        next();
      });
    }

    setupDropzone();

    input.addEventListener('change', function () {
      var files = Array.prototype.slice.call(input.files || []);
      input.value = '';
      files.forEach(enqueue);
    });

    // No permitir enviar el ticket mientras un adjunto aun se esta subiendo
    if (input.form) {
      input.form.addEventListener('submit', function (e) {
        if (pending > 0) {
          e.preventDefault();
          notify('warning', 'Espere a que terminen de subir los adjuntos antes de guardar.');
        }
      });
    }

    refresh();
  }

  // Al enviar el formulario, valida que la descripcion no este vacia (solo texto o imagenes)
  function guardSubmit(formId, fieldId, message) {
    var form = document.getElementById(formId);
    var field = document.getElementById(fieldId);
    if (!form || !field) { return; }
    form.addEventListener('submit', function (e) {
      var html = field.value || '';
      var text = html.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').trim();
      if (text.length === 0 && html.indexOf('<img') === -1) {
        e.preventDefault();
        notify('warning', message || 'Escriba la descripción.');
      }
    });
  }

  window.LhdEditor = {
    createEditor: createEditor,
    attachments: attachments,
    guardSubmit: guardSubmit
  };
})(window, document);
