<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: reports.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Form for admins to submit report queries.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    ' VBScript no trae IIf
    Function IIf(cond, a, b)
      If cond Then
        IIf = a
      Else
        IIf = b
      End If
    End Function

    ' Escribe un campo desplegable con buscador.
    ' fieldName: nombre del campo oculto que viaja en el POST (0 = Todos)
    Sub WriteSearchSelect(fieldName, labelText, sqlText, idCol, nameCol, placeholderText)
      Dim rs
      Response.Write("<div class=""rp-field""><label class=""rp-label"">" & labelText & "</label>")
      Response.Write("<div class=""ss"" data-name=""" & fieldName & """>")
      Response.Write("<input type=""hidden"" name=""" & fieldName & """ value=""0"">")
      Response.Write("<input type=""text"" class=""ss-input"" autocomplete=""off"" placeholder=""" & placeholderText & """>")
      Response.Write("<span class=""ss-clear"" title=""Limpiar"">&times;</span><span class=""ss-arrow""><svg viewBox=""0 0 24 24"" width=""14"" height=""14"" fill=""none"" stroke=""currentColor"" stroke-width=""2.4"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true""><path d=""m6 9 6 6 6-6""/></svg></span>")
      Response.Write("<ul class=""ss-list""><li data-id=""0"" class=""ss-all"">Todos</li>")
      Set rs = SQLQuery(cnnDB, sqlText)
      Do While Not rs.EOF
        Response.Write("<li data-id=""" & rs(idCol) & """>" & Server.HTMLEncode(rs(nameCol) & "") & "</li>")
        rs.MoveNext
      Loop
      rs.Close
      Response.Write("<li class=""ss-empty"">Sin coincidencias</li></ul></div></div>")
    End Sub
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Reports")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style>
      .rp-card { max-width: 860px; margin: 0 auto; text-align: left; }
      .rp-section { margin: 0 0 18px; }
      .rp-section .section-header { margin-top: 0; }
      .rp-section h3 { margin: 0 0 10px; font-size: 13px; text-transform: uppercase; letter-spacing: .05em; color: var(--color-muted); }
      .rp-types { display: flex; gap: 10px; flex-wrap: wrap; }
      .rp-types label { flex: 1 1 150px; border: 1px solid var(--color-border); border-radius: var(--radius); padding: 10px 12px; cursor: pointer; background: var(--color-surface); transition: var(--transition); }
      .rp-types label:hover { border-color: var(--color-primary); }
      .rp-types input { margin-right: 6px; }
      .rp-types label.on { border-color: var(--color-primary); background: var(--color-primary-light); font-weight: 600; }
      .rp-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 12px; }
      .rp-label { display: block; font-size: 12px; font-weight: 600; color: var(--color-text-secondary); margin-bottom: 4px; }
      .rp-dates { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
      .rp-dates select { padding: 6px 4px; border: 1px solid var(--color-border); border-radius: var(--radius-sm); background: var(--color-surface); }
      .ss { position: relative; }
      .ss-input { width: 100%; box-sizing: border-box; padding: 9px 52px 9px 12px; border: 1px solid var(--color-border); border-radius: var(--radius); background: var(--color-surface); font-size: 14px; transition: var(--transition); }
      .ss-input:focus { outline: none; border-color: var(--color-primary); box-shadow: 0 0 0 3px var(--color-primary-light); }
      .ss.has-value .ss-input { border-color: var(--color-primary); background: var(--color-primary-light); font-weight: 600; }
      .ss-arrow { position: absolute; right: 12px; top: 50%; transform: translateY(-50%); color: var(--color-muted); pointer-events: none; }
      .ss-clear { position: absolute; right: 30px; top: 50%; transform: translateY(-50%); color: var(--color-muted); cursor: pointer; font-size: 18px; display: none; }
      .ss.has-value .ss-clear { display: block; }
      .ss.open { z-index: 1000; }
      .rp-field { position: relative; }
      .rp-field:focus-within { z-index: 1000; }
      .ss-list { display: none; position: absolute; z-index: 1000; left: 0; right: 0; top: calc(100% + 4px); max-height: 220px; overflow-y: auto; margin: 0; padding: 4px; list-style: none; background: #FFFFFF !important; opacity: 1 !important; border: 1px solid #CBD5E0; border-radius: 8px; box-shadow: 0 8px 24px rgba(0,0,0,.18); }
      .ss.open .ss-list { display: block; }
      .ss-list li { display: block; padding: 8px 10px; margin: 0; border-radius: 4px; cursor: pointer; font-size: 14px; background: #FFFFFF; color: #1A202C; text-align: left; }
      .ss-list li.active, .ss-list li:hover { background: #E8F0FB; color: #003F8A; }
      .ss-list li.ss-empty { display: none; color: var(--color-muted); cursor: default; }
      .ss-list li.ss-empty:hover { background: none; color: var(--color-muted); }
      .ss-list mark { background: var(--color-warning-bg); color: inherit; padding: 0; }
      .rp-actions { text-align: center; margin-top: 8px; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

    %>
    <div class="lhd-wrapper">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128202;</span>
          <%=lang(cnnDB, "Reports")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-primary" href="reportans.asp">&#128202; Cumplimiento de ANS</a>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>
      <form method="post" action="viewreports.asp" class="rp-card">
              <div class="lhd-card rp-section">
                <div class="section-header">Tipo de reporte</div>
                <div class="rp-types">
                  <label class="on"><input type="radio" name="type" value="0" checked="checked"><%=lang(cnnDB, "Department")%>&nbsp;<%=lang(cnnDB, "Report")%></label>
                  <label><input type="radio" name="type" value="1"><%=lang(cnnDB, "Category")%>&nbsp;<%=lang(cnnDB, "Report")%></label>
                  <label><input type="radio" name="type" value="2"><%=lang(cnnDB, "SupportRep")%>&nbsp;<%=lang(cnnDB, "Report")%></label>
                </div>
              </div>

              <div class="lhd-card rp-section">
                <div class="section-header">Filtros (escribe para buscar)</div>
                <div class="rp-grid">
                  <%
                    Call WriteSearchSelect("f_dept", lang(cnnDB, "Department"), "SELECT department_id, dname FROM departments WHERE department_id > 0 ORDER BY dname ASC", "department_id", "dname", "Todos los departamentos")
                    Call WriteSearchSelect("f_cat", lang(cnnDB, "Category"), "SELECT category_id, cname FROM categories WHERE category_id > 0 ORDER BY cname ASC", "category_id", "cname", "Todas las categor&iacute;as")
                    Call WriteSearchSelect("f_rep", lang(cnnDB, "SupportRep"), "SELECT sid, uid FROM tblUsers WHERE sid>0 AND IsRep=1 ORDER BY uid ASC", "sid", "uid", "Todos los representantes")
                  %>
                </div>
              </div>

              <div class="lhd-card rp-section">
                <div class="section-header">Periodo</div>
                <div class="rp-dates">
                <%
                  Dim count, selDay, selMonth, selYear
                  ' Por defecto: desde el 1 de enero del a&ntilde;o actual hasta hoy
                  Response.Write("<div><label class=""rp-label"">" & lang(cnnDB, "StartDate") & "</label><select name=""s_day"">")
                  For count = 1 to 31
                    Response.Write("<option value=""" & count & """" & IIf(count = 1, " selected", "") & ">" & count & "</option>")
                  Next
                  Response.Write("</select> / <select name=""s_month"">")
                  For count = 1 to 12
                    Response.Write("<option value=""" & count & """" & IIf(count = 1, " selected", "") & ">" & count & "</option>")
                  Next
                  Response.Write("</select> / <select name=""s_year"">")
                  For count = 2000 to Year(Now)
                    Response.Write("<option value=""" & count & """" & IIf(count = Year(Now), " selected", "") & ">" & count & "</option>")
                  Next
                  Response.Write("</select></div>")

                  Response.Write("<div><label class=""rp-label"">" & lang(cnnDB, "EndDate") & "</label><select name=""e_day"">")
                  For count = 1 to 31
                    Response.Write("<option value=""" & count & """" & IIf(count = Day(Now), " selected", "") & ">" & count & "</option>")
                  Next
                  Response.Write("</select> / <select name=""e_month"">")
                  For count = 1 to 12
                    Response.Write("<option value=""" & count & """" & IIf(count = Month(Now), " selected", "") & ">" & count & "</option>")
                  Next
                  Response.Write("</select> / <select name=""e_year"">")
                  For count = 2000 to Year(Now)
                    Response.Write("<option value=""" & count & """" & IIf(count = Year(Now), " selected", "") & ">" & count & "</option>")
                  Next
                  Response.Write("</select></div>")
                %>
                </div>
              </div>

              <div class="rp-actions">
                <input type="submit" class="btn btn-primary" value="<%=lang(cnnDB, "ViewReport")%>">
              </div>
      </form>
    </div>

    <script type="text/javascript">
    (function () {
      // Resalta el tipo de reporte elegido
      var radios = document.querySelectorAll('.rp-types input');
      function syncTypes() {
        for (var i = 0; i < radios.length; i++) radios[i].parentNode.className = radios[i].checked ? 'on' : '';
      }
      for (var r = 0; r < radios.length; r++) radios[r].onchange = syncTypes;

      function norm(s) {
        s = (s || '').toLowerCase();
        return s.normalize ? s.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : s;
      }
      function esc(s) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;'); }

      var boxes = document.querySelectorAll('.ss');
      for (var b = 0; b < boxes.length; b++) (function (box) {
        var hidden = box.querySelector('input[type=hidden]');
        var input = box.querySelector('.ss-input');
        var clear = box.querySelector('.ss-clear');
        var items = box.querySelectorAll('.ss-list li[data-id]');
        var empty = box.querySelector('.ss-empty');
        var active = -1;

        function visible() {
          var v = [];
          for (var i = 0; i < items.length; i++) if (items[i].style.display !== 'none') v.push(items[i]);
          return v;
        }
        function setActive(n) {
          var v = visible();
          for (var i = 0; i < items.length; i++) items[i].className = items[i].className.replace(' active', '');
          if (!v.length) { active = -1; return; }
          active = (n + v.length) % v.length;
          v[active].className += ' active';
          if (v[active].scrollIntoView) v[active].scrollIntoView({ block: 'nearest' });
        }
        function filter() {
          var q = norm(input.value), shown = 0;
          for (var i = 0; i < items.length; i++) {
            var txt = items[i].getAttribute('data-txt');
            if (txt === null) { txt = items[i].textContent; items[i].setAttribute('data-txt', txt); }
            var pos = q ? norm(txt).indexOf(q) : 0;
            var ok = (pos > -1) && (items[i].getAttribute('data-id') !== '0' || !q);
            items[i].style.display = ok ? '' : 'none';
            if (ok) {
              shown++;
              items[i].innerHTML = (q && pos > -1)
                ? esc(txt.substr(0, pos)) + '<mark>' + esc(txt.substr(pos, q.length)) + '</mark>' + esc(txt.substr(pos + q.length))
                : esc(txt);
            }
          }
          empty.style.display = shown ? 'none' : 'block';
          setActive(0);
        }
        function choose(li) {
          var id = li.getAttribute('data-id');
          hidden.value = id;
          input.value = id === '0' ? '' : li.getAttribute('data-txt') || li.textContent;
          box.className = 'ss' + (id === '0' ? '' : ' has-value');
        }
        function open() { box.className = 'ss open' + (hidden.value !== '0' ? ' has-value' : ''); filter(); }
        function close() {
          // Si lo escrito no coincide con una opcion, se restaura la seleccion vigente
          var cur = null;
          for (var i = 0; i < items.length; i++) if (items[i].getAttribute('data-id') === hidden.value) cur = items[i];
          input.value = (cur && hidden.value !== '0') ? (cur.getAttribute('data-txt') || cur.textContent) : '';
          box.className = 'ss' + (hidden.value !== '0' ? ' has-value' : '');
        }

        input.onfocus = function () { input.select(); open(); };
        input.onclick = function () { if (box.className.indexOf('open') < 0) open(); };
        input.oninput = function () { if (box.className.indexOf('open') < 0) box.className += ' open'; filter(); };
        input.onkeydown = function (e) {
          e = e || window.event;
          var k = e.keyCode;
          if (k === 40) { setActive(active + 1); e.preventDefault && e.preventDefault(); }
          else if (k === 38) { setActive(active - 1); e.preventDefault && e.preventDefault(); }
          else if (k === 13) {
            var v = visible();
            if (box.className.indexOf('open') > -1) { if (v[active]) choose(v[active]); close(); }
            if (e.preventDefault) e.preventDefault();
          }
          else if (k === 27) { close(); input.blur(); }
        };
        for (var i = 0; i < items.length; i++) (function (li) {
          li.onmousedown = function (e) { choose(li); close(); if (e && e.preventDefault) e.preventDefault(); input.blur(); };
        })(items[i]);
        clear.onmousedown = function (e) { hidden.value = '0'; input.value = ''; box.className = 'ss'; if (e && e.preventDefault) e.preventDefault(); };
        input.onblur = function () { setTimeout(close, 120); };
      })(boxes[b]);
    })();
    </script>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close

    %>
  </body>
</html>
