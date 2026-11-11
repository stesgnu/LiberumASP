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

  Filename: viewdep.asp
  Purpose:  The list of departments, with search and pagination (10 per page).
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Departments")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .dep-toolbar { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between; padding: 14px 16px; border-bottom: 1px solid var(--color-border-light, #E2E8F0); }
      .dep-search-form { display: flex; gap: 8px; flex: 1 1 320px; max-width: 480px; margin: 0; }
      .dep-search { flex: 1; padding: 10px 14px; border: 1px solid var(--color-border, #CBD5E0); border-radius: 10px; font-size: 14px; box-sizing: border-box; min-width: 0; }
      .dep-search:focus { outline: none; border-color: var(--color-primary, #0057B8); box-shadow: 0 0 0 3px var(--color-primary-light, #E8F0FB); }
      .dep-count { font-size: 13px; color: #64748b; }
      .dep-pager { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; justify-content: space-between; padding: 14px 16px; border-top: 1px solid var(--color-border-light, #E2E8F0); }
      .dep-pages { display: flex; gap: 4px; flex-wrap: wrap; }
      .dep-pg { min-width: 34px; height: 34px; padding: 0 10px; display: inline-flex; align-items: center; justify-content: center; border: 1px solid var(--color-border, #CBD5E0); border-radius: 8px; background: #fff; color: #334155; font-size: 13px; font-weight: 600; text-decoration: none !important; }
      a.dep-pg:hover { border-color: var(--color-primary, #0057B8); color: var(--color-primary, #0057B8); }
      .dep-pg.on { background: var(--color-primary, #0057B8); border-color: var(--color-primary, #0057B8); color: #fff; }
      .dep-pg.off { color: #A0AEC0; background: #F7FAFC; cursor: default; }
      .dep-empty { padding: 34px 16px; text-align: center; color: #64748b; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      ' ---- Parametros: busqueda y pagina (10 registros por pagina) ----
      Const PER_PAGE = 10
      Dim q, qSql, pageNo, totalRows, totalPages, skipRows, whereSql
      q = Left(Trim(Request.QueryString("q")), 50)
      qSql = Replace(q, "'", "''")
      qSql = Replace(qSql, "[", "")
      qSql = Replace(qSql, "]", "")
      qSql = Replace(qSql, "%", "")
      whereSql = "department_id > 0"
      If Len(qSql) > 0 Then whereSql = whereSql & " AND dname LIKE '%" & qSql & "%'"

      Dim rstCnt
      Set rstCnt = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM departments WHERE " & whereSql)
      totalRows = SafeCLng(rstCnt("total"))
      rstCnt.Close

      totalPages = Int((totalRows + PER_PAGE - 1) / PER_PAGE)
      If totalPages < 1 Then totalPages = 1
      pageNo = SafeCLng(Request.QueryString("page"))
      If pageNo < 1 Then pageNo = 1
      If pageNo > totalPages Then pageNo = totalPages
      skipRows = (pageNo - 1) * PER_PAGE

      ' Solo se traen las filas hasta el final de la pagina pedida
      Dim rstDepts
      Set rstDepts = SQLQuery(cnnDB, "SELECT TOP " & (skipRows + PER_PAGE) & " department_id, dname FROM departments WHERE " & whereSql & " ORDER BY dname ASC")

      Dim skipI
      skipI = 0
      Do While Not rstDepts.EOF And skipI < skipRows
        rstDepts.MoveNext
        skipI = skipI + 1
      Loop

      Dim firstShown, lastShown, qParam, pgI, pgFrom, pgTo
      qParam = ""
      If Len(q) > 0 Then qParam = "&amp;q=" & Server.URLEncode(q)
      firstShown = 0
      If totalRows > 0 Then firstShown = skipRows + 1
      lastShown = skipRows + PER_PAGE
      If lastShown > totalRows Then lastShown = totalRows
    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#127970;</span>
          <%=lang(cnnDB, "Departments")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <form method="post" action="modify.asp?mtype=3" style="margin:0">
            <button type="submit" class="btn btn-primary">&#10133; <%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Department")%></button>
          </form>
        </div>
      </div>

      <div class="lhd-card lhd-card-flush">
        <div class="dep-toolbar">
          <form method="get" action="viewdep.asp" class="dep-search-form">
            <input type="text" name="q" class="dep-search" value="<% = Server.HTMLEncode(q) %>" placeholder="Buscar departamento..." maxlength="50" autocomplete="off">
            <button type="submit" class="btn btn-primary">&#128269; Buscar</button>
            <% If Len(q) > 0 Then %><a class="btn btn-secondary" href="viewdep.asp">Limpiar</a><% End If %>
          </form>
          <span class="dep-count">
            <% If totalRows = 0 Then %>
              Sin resultados
            <% Else %>
              Mostrando <strong><% = firstShown %>&ndash;<% = lastShown %></strong> de <strong><% = totalRows %></strong>
            <% End If %>
          </span>
        </div>

        <div style="overflow-x:auto">
          <table class="lhd-table">
            <thead>
              <tr>
                <th><%=lang(cnnDB, "Department")%></th>
                <th class="col-actions"><%=lang(cnnDB, "Modify")%></th>
              </tr>
            </thead>
            <tbody>
            <%
              Dim shownRows
              shownRows = 0
              Do While Not rstDepts.EOF And shownRows < PER_PAGE
            %>
              <tr>
                <td><strong><% = Server.HTMLEncode(rstDepts("dname") & "") %></strong></td>
                <td class="col-actions">
                  <a class="btn btn-secondary btn-sm" href="modify.asp?mtype=3&amp;id=<% = rstDepts("department_id") %>">&#9998; <%=lang(cnnDB, "Edit")%></a>
                  <a class="btn btn-secondary btn-sm" href="confdelete.asp?mtype=3&amp;id=<% = rstDepts("department_id") %>">&#128465; <%=lang(cnnDB, "Delete")%></a>
                </td>
              </tr>
            <%
                shownRows = shownRows + 1
                rstDepts.MoveNext
              Loop
            %>
            </tbody>
          </table>
          <% If totalRows = 0 Then %>
            <div class="dep-empty">
              <% If Len(q) > 0 Then %>No se encontraron departamentos que coincidan con "<strong><% = Server.HTMLEncode(q) %></strong>".<% Else %>Todav&iacute;a no hay departamentos registrados.<% End If %>
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
        <div class="dep-pager">
          <span class="dep-count">P&aacute;gina <strong><% = pageNo %></strong> de <strong><% = totalPages %></strong></span>
          <div class="dep-pages">
            <% If pageNo > 1 Then %>
              <a class="dep-pg" href="viewdep.asp?page=1<% = qParam %>" title="Primera">&laquo;</a>
              <a class="dep-pg" href="viewdep.asp?page=<% = pageNo - 1 %><% = qParam %>"><%=lang(cnnDB, "Previous")%></a>
            <% Else %>
              <span class="dep-pg off">&laquo;</span>
              <span class="dep-pg off"><%=lang(cnnDB, "Previous")%></span>
            <% End If %>
            <% For pgI = pgFrom To pgTo
                 If pgI = pageNo Then %>
              <span class="dep-pg on"><% = pgI %></span>
            <%   Else %>
              <a class="dep-pg" href="viewdep.asp?page=<% = pgI %><% = qParam %>"><% = pgI %></a>
            <%   End If
               Next %>
            <% If pageNo < totalPages Then %>
              <a class="dep-pg" href="viewdep.asp?page=<% = pageNo + 1 %><% = qParam %>"><%=lang(cnnDB, "Next")%></a>
              <a class="dep-pg" href="viewdep.asp?page=<% = totalPages %><% = qParam %>" title="&Uacute;ltima">&raquo;</a>
            <% Else %>
              <span class="dep-pg off"><%=lang(cnnDB, "Next")%></span>
              <span class="dep-pg off">&raquo;</span>
            <% End If %>
          </div>
        </div>
        <% End If %>
      </div>
    </div>

    <%
      rstDepts.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
