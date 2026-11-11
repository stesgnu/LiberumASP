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

  Filename: viewusers.asp
  Purpose:  Lista de usuarios con busqueda, filtro por rol y paginado (10 por pagina).
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ManageUsers")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .usr-stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin-bottom: 18px; }
      .usr-stat { background: var(--color-surface, #fff); border: 1px solid var(--color-border-light, #E2E8F0); border-left: 4px solid var(--color-primary, #0057B8); border-radius: 10px; padding: 14px 16px; box-shadow: var(--shadow-sm, 0 1px 3px rgba(0,0,0,.08)); }
      .usr-stat .n { font-size: 26px; font-weight: 700; color: #1a202c; line-height: 1.1; }
      .usr-stat .l { font-size: 12px; color: #64748b; text-transform: uppercase; letter-spacing: .04em; margin-top: 2px; }
      .usr-toolbar { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between; padding: 14px 16px; border-bottom: 1px solid var(--color-border-light, #E2E8F0); }
      .usr-search-form { display: flex; gap: 8px; flex: 1 1 320px; max-width: 480px; margin: 0; }
      .usr-search { flex: 1; min-width: 0; padding: 10px 14px; border: 1px solid var(--color-border, #CBD5E0); border-radius: 10px; font-size: 14px; box-sizing: border-box; }
      .usr-search:focus { outline: none; border-color: var(--color-primary, #0057B8); box-shadow: 0 0 0 3px var(--color-primary-light, #E8F0FB); }
      .usr-chips { display: flex; gap: 8px; flex-wrap: wrap; }
      .usr-chip { border: 1px solid var(--color-border, #CBD5E0); background: #fff; color: #475569; border-radius: 999px; padding: 6px 14px; font-size: 13px; text-decoration: none !important; transition: all .15s ease; }
      .usr-chip:hover { border-color: var(--color-primary, #0057B8); color: var(--color-primary, #0057B8); }
      .usr-chip.on { background: var(--color-primary, #0057B8); border-color: var(--color-primary, #0057B8); color: #fff; }
      .usr-av { width: 36px; height: 36px; border-radius: 50%; background: linear-gradient(135deg, var(--color-primary, #0057B8), var(--color-accent, #0284C7)); color: #fff; font-weight: 700; font-size: 13px; display: inline-flex; align-items: center; justify-content: center; flex-shrink: 0; }
      .usr-cell { display: flex; align-items: center; gap: 12px; }
      .usr-name { font-weight: 600; color: #1a202c; }
      .usr-sub { font-size: 12px; color: #64748b; }
      .usr-role { display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; }
      .usr-role.rep { background: var(--color-primary-light, #E8F0FB); color: var(--color-primary-dark, #003F8A); }
      .usr-role.usr { background: #F1F5F9; color: #475569; }
      .usr-access { font-size: 12px; color: #64748b; }
      .usr-empty { padding: 36px 16px; text-align: center; color: #64748b; }
      .usr-tbl-wrap { overflow-x: auto; }
      .usr-edit-form { margin: 0; display: inline; }
      .usr-count { font-size: 13px; color: #64748b; }
      tr.usr-row { cursor: pointer; }
      .usr-pager { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; justify-content: space-between; padding: 14px 16px; border-top: 1px solid var(--color-border-light, #E2E8F0); }
      .usr-pages { display: flex; gap: 4px; flex-wrap: wrap; }
      .usr-pg { min-width: 34px; height: 34px; padding: 0 10px; display: inline-flex; align-items: center; justify-content: center; border: 1px solid var(--color-border, #CBD5E0); border-radius: 8px; background: #fff; color: #334155; font-size: 13px; font-weight: 600; text-decoration: none !important; }
      a.usr-pg:hover { border-color: var(--color-primary, #0057B8); color: var(--color-primary, #0057B8); }
      .usr-pg.on { background: var(--color-primary, #0057B8); border-color: var(--color-primary, #0057B8); color: #fff; }
      .usr-pg.off { color: #A0AEC0; background: #F7FAFC; cursor: default; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      ' ---- Parametros: busqueda, rol y pagina (10 registros por pagina) ----
      Const PER_PAGE = 10
      Dim q, qSql, roleF, pageNo, totalRows, totalPages, skipRows, whereSql
      q = Left(Trim(Request.QueryString("q")), 50)
      qSql = Replace(q, "'", "''")
      qSql = Replace(qSql, "[", "")
      qSql = Replace(qSql, "]", "")
      qSql = Replace(qSql, "%", "")
      roleF = LCase(Trim(Request.QueryString("role")))
      If roleF <> "rep" And roleF <> "usr" Then roleF = "all"

      whereSql = "sid > 0"
      If roleF = "rep" Then whereSql = whereSql & " AND IsRep <> 0"
      If roleF = "usr" Then whereSql = whereSql & " AND (IsRep = 0 OR IsRep IS NULL)"
      If Len(qSql) > 0 Then
        whereSql = whereSql & " AND (uid LIKE '%" & qSql & "%' OR fname LIKE '%" & qSql & "%' OR email1 LIKE '%" & qSql & "%')"
      End If

      ' Resumen (siempre sobre todos los usuarios, sin importar el filtro)
      Dim rstSt, nUsers, nReps, nUsrs
      Set rstSt = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM tblUsers WHERE sid > 0")
      nUsers = SafeCLng(rstSt("total"))
      rstSt.Close
      Set rstSt = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM tblUsers WHERE sid > 0 AND IsRep <> 0")
      nReps = SafeCLng(rstSt("total"))
      rstSt.Close
      nUsrs = nUsers - nReps

      ' Total del resultado filtrado y pagina actual
      Set rstSt = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM tblUsers WHERE " & whereSql)
      totalRows = SafeCLng(rstSt("total"))
      rstSt.Close

      totalPages = Int((totalRows + PER_PAGE - 1) / PER_PAGE)
      If totalPages < 1 Then totalPages = 1
      pageNo = SafeCLng(Request.QueryString("page"))
      If pageNo < 1 Then pageNo = 1
      If pageNo > totalPages Then pageNo = totalPages
      skipRows = (pageNo - 1) * PER_PAGE

      ' Solo se traen las filas hasta el final de la pagina pedida
      Dim rstUserList
      Set rstUserList = SQLQuery(cnnDB, "SELECT TOP " & (skipRows + PER_PAGE) & " sid, uid, fname, email1, IsRep, RepAccess FROM tblUsers WHERE " & whereSql & " ORDER BY uid ASC")

      Dim skipI
      skipI = 0
      Do While Not rstUserList.EOF And skipI < skipRows
        rstUserList.MoveNext
        skipI = skipI + 1
      Loop

      ' Piezas para armar los enlaces conservando la busqueda y el rol
      Dim qParam, roleParam, baseQs, firstShown, lastShown, pgI, pgFrom, pgTo
      qParam = ""
      If Len(q) > 0 Then qParam = "&amp;q=" & Server.URLEncode(q)
      roleParam = ""
      If roleF <> "all" Then roleParam = "&amp;role=" & roleF
      baseQs = qParam & roleParam
      firstShown = 0
      If totalRows > 0 Then firstShown = skipRows + 1
      lastShown = skipRows + PER_PAGE
      If lastShown > totalRows Then lastShown = totalRows

      Dim shownRows, isRepRow, accessTxt, nm, ini, parts
      shownRows = 0
    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128101;</span>
          <%=lang(cnnDB, "ManageUsers")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <form name="addusr" method="POST" action="adduser.asp" style="margin:0">
            <button type="submit" class="btn btn-primary">&#10133; <%=lang(cnnDB, "AddNewUsers")%></button>
          </form>
        </div>
      </div>

      <div class="usr-stats">
        <div class="usr-stat"><div class="n"><% = nUsers %></div><div class="l">Usuarios en total</div></div>
        <div class="usr-stat"><div class="n"><% = nReps %></div><div class="l">Agentes de soporte</div></div>
        <div class="usr-stat"><div class="n"><% = nUsrs %></div><div class="l">Usuarios finales</div></div>
      </div>

      <div class="lhd-card lhd-card-flush">
        <div class="usr-toolbar">
          <form method="get" action="viewusers.asp" class="usr-search-form">
            <% If roleF <> "all" Then %><input type="hidden" name="role" value="<% = roleF %>"><% End If %>
            <input type="text" name="q" class="usr-search" value="<% = Server.HTMLEncode(q) %>" placeholder="Buscar por usuario, nombre o correo..." maxlength="50" autocomplete="off">
            <button type="submit" class="btn btn-primary">&#128269; Buscar</button>
            <% If Len(q) > 0 Then %><a class="btn btn-secondary" href="viewusers.asp<% If roleF <> "all" Then Response.Write("?role=" & roleF) %>">Limpiar</a><% End If %>
          </form>
          <div class="usr-chips">
            <a class="usr-chip<% If roleF = "all" Then Response.Write(" on") %>" href="viewusers.asp<% If Len(q) > 0 Then Response.Write("?q=" & Server.URLEncode(q)) %>">Todos</a>
            <a class="usr-chip<% If roleF = "rep" Then Response.Write(" on") %>" href="viewusers.asp?role=rep<% = qParam %>">Agentes</a>
            <a class="usr-chip<% If roleF = "usr" Then Response.Write(" on") %>" href="viewusers.asp?role=usr<% = qParam %>">Usuarios</a>
          </div>
          <span class="usr-count">
            <% If totalRows = 0 Then %>
              Sin resultados
            <% Else %>
              Mostrando <strong><% = firstShown %>&ndash;<% = lastShown %></strong> de <strong><% = totalRows %></strong>
            <% End If %>
          </span>
        </div>

        <div class="usr-tbl-wrap">
          <table class="lhd-table" id="usrTable">
            <thead>
              <tr>
                <th>Usuario</th>
                <th>Correo</th>
                <th>Rol</th>
                <th>Acceso</th>
                <th class="col-actions"></th>
              </tr>
            </thead>
            <tbody>
              <%
                Do While Not rstUserList.EOF And shownRows < PER_PAGE
                  isRepRow = (SafeCInt(rstUserList("IsRep")) <> 0)
                  nm = Trim(rstUserList("fname") & "")
                  If Len(nm) = 0 Then nm = rstUserList("uid") & ""
                  ' iniciales (maximo dos)
                  parts = Split(nm, " ")
                  If UBound(parts) >= 1 Then
                    ini = UCase(Left(parts(0), 1) & Left(parts(1), 1))
                  Else
                    ini = UCase(Left(nm, 2))
                  End If
                  accessTxt = "&mdash;"
                  If isRepRow Then
                    Select Case SafeCInt(rstUserList("RepAccess"))
                      Case 1: accessTxt = "Solo sus tickets"
                      Case 2: accessTxt = "Solo lectura"
                      Case Else: accessTxt = "Completo"
                    End Select
                  End If
              %>
              <tr class="usr-row">
                <td>
                  <div class="usr-cell">
                    <span class="usr-av"><% = Server.HTMLEncode(ini) %></span>
                    <div>
                      <div class="usr-name"><% = Server.HTMLEncode(nm) %></div>
                      <div class="usr-sub"><% = Server.HTMLEncode(rstUserList("uid") & "") %></div>
                    </div>
                  </div>
                </td>
                <td><% If Len(rstUserList("email1") & "") > 0 Then Response.Write(Server.HTMLEncode(rstUserList("email1") & "")) Else Response.Write("&mdash;") End If %></td>
                <td>
                  <% If isRepRow Then %>
                    <span class="usr-role rep">Agente</span>
                  <% Else %>
                    <span class="usr-role usr">Usuario</span>
                  <% End If %>
                </td>
                <td><span class="usr-access"><% = accessTxt %></span></td>
                <td class="col-actions">
                  <form method="POST" action="moduser.asp" class="usr-edit-form">
                    <input type="hidden" name="usersid" value="<% = rstUserList("sid") %>">
                    <button type="submit" class="btn btn-secondary btn-sm">&#9998; <%=lang(cnnDB, "EditUserAccount")%></button>
                  </form>
                </td>
              </tr>
              <%
                  shownRows = shownRows + 1
                  rstUserList.MoveNext
                Loop
              %>
            </tbody>
          </table>
          <% If totalRows = 0 Then %>
            <div class="usr-empty">
              <% If Len(q) > 0 Then %>No se encontraron usuarios que coincidan con "<strong><% = Server.HTMLEncode(q) %></strong>".<% Else %>No hay usuarios con este filtro.<% End If %>
            </div>
          <% End If %>
        </div>

        <% If totalPages > 1 Then
             pgFrom = pageNo - 2
             If pgFrom < 1 Then pgFrom = 1
             pgTo = pgFrom + 4
             If pgTo > totalPages Then
               pgTo = totalPages
               pgFrom = pgTo - 4
               If pgFrom < 1 Then pgFrom = 1
             End If %>
        <div class="usr-pager">
          <span class="usr-count">P&aacute;gina <strong><% = pageNo %></strong> de <strong><% = totalPages %></strong></span>
          <div class="usr-pages">
            <% If pageNo > 1 Then %>
              <a class="usr-pg" href="viewusers.asp?page=1<% = baseQs %>" title="Primera">&laquo;</a>
              <a class="usr-pg" href="viewusers.asp?page=<% = pageNo - 1 %><% = baseQs %>"><%=lang(cnnDB, "Previous")%></a>
            <% Else %>
              <span class="usr-pg off">&laquo;</span>
              <span class="usr-pg off"><%=lang(cnnDB, "Previous")%></span>
            <% End If %>
            <% For pgI = pgFrom To pgTo
                 If pgI = pageNo Then %>
              <span class="usr-pg on"><% = pgI %></span>
            <%   Else %>
              <a class="usr-pg" href="viewusers.asp?page=<% = pgI %><% = baseQs %>"><% = pgI %></a>
            <%   End If
               Next %>
            <% If pageNo < totalPages Then %>
              <a class="usr-pg" href="viewusers.asp?page=<% = pageNo + 1 %><% = baseQs %>"><%=lang(cnnDB, "Next")%></a>
              <a class="usr-pg" href="viewusers.asp?page=<% = totalPages %><% = baseQs %>" title="&Uacute;ltima">&raquo;</a>
            <% Else %>
              <span class="usr-pg off"><%=lang(cnnDB, "Next")%></span>
              <span class="usr-pg off">&raquo;</span>
            <% End If %>
          </div>
        </div>
        <% End If %>
      </div>

    </div>

    <script type="text/javascript">
    (function () {
      // clic en la fila = editar ese usuario
      var rows = document.querySelectorAll('#usrTable tbody tr.usr-row');
      for (var r = 0; r < rows.length; r++) rows[r].onclick = function (e) {
        var t = e.target || e.srcElement;
        if (t.tagName === 'BUTTON' || t.tagName === 'A' || t.tagName === 'INPUT') return;
        var f = this.querySelector('form');
        if (f) f.submit();
      };
    })();
    </script>

    <%
      rstUserList.Close
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
