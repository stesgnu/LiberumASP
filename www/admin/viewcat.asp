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

  Filename: viewcat.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Displays a list of categories
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Categories")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .cat-toolbar { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between; padding: 14px 16px; border-bottom: 1px solid var(--color-border-light, #E2E8F0); }
      .cat-search { flex: 1 1 260px; max-width: 420px; padding: 10px 14px; border: 1px solid var(--color-border, #CBD5E0); border-radius: 10px; font-size: 14px; box-sizing: border-box; }
      .cat-search:focus { outline: none; border-color: var(--color-primary, #0057B8); box-shadow: 0 0 0 3px var(--color-primary-light, #E8F0FB); }
      .cat-av { width: 30px; height: 30px; border-radius: 50%; background: linear-gradient(135deg, var(--color-primary, #0057B8), var(--color-accent, #0284C7)); color: #fff; font-weight: 700; font-size: 12px; display: inline-flex; align-items: center; justify-content: center; margin-right: 10px; vertical-align: middle; }
      .cat-rep-sub { font-size: 12px; color: #64748b; }
      .cat-empty { padding: 30px 16px; text-align: center; color: #64748b; display: none; }
      .cat-count { font-size: 13px; color: #64748b; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      ' Categoria + nombre completo del representante asignado
      Dim rstCategories, repName, repUid, nm, ini, parts, nCats
      Set rstCategories = SQLQuery(cnnDB, "SELECT c.category_id, c.cname, c.rep_id, u.fname, u.uid " & _
        "FROM categories AS c LEFT JOIN tblUsers AS u ON c.rep_id = u.sid " & _
        "WHERE c.category_id > 0 ORDER BY c.cname ASC")
      nCats = 0
    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128193;</span>
          <%=lang(cnnDB, "Categories")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <form method="post" action="modify.asp?mtype=2" style="margin:0">
            <button type="submit" class="btn btn-primary">&#10133; <%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Category")%></button>
          </form>
        </div>
      </div>

      <div class="lhd-card lhd-card-flush">
        <div class="cat-toolbar">
          <input type="text" id="catSearch" class="cat-search" placeholder="Buscar categor&iacute;a o representante..." autocomplete="off">
          <span class="cat-count" id="catCount"></span>
        </div>
        <div style="overflow-x:auto">
          <table class="lhd-table" id="catTable">
            <thead>
              <tr>
                <th><%=lang(cnnDB, "Category")%></th>
                <th><%=lang(cnnDB, "Rep")%></th>
                <th class="col-actions"><%=lang(cnnDB, "Modify")%></th>
              </tr>
            </thead>
            <tbody>
            <%
              Do While Not rstCategories.EOF
                nCats = nCats + 1
                repName = Trim(rstCategories("fname") & "")
                repUid = Trim(rstCategories("uid") & "")
                nm = repName
                If Len(nm) = 0 Then nm = repUid
                ini = ""
                If Len(nm) > 0 Then
                  parts = Split(nm, " ")
                  If UBound(parts) >= 1 Then
                    ini = UCase(Left(parts(0), 1) & Left(parts(1), 1))
                  Else
                    ini = UCase(Left(nm, 2))
                  End If
                End If
            %>
              <tr data-text="<% = Server.HTMLEncode(LCase((rstCategories("cname") & "") & " " & repName & " " & repUid)) %>">
                <td><strong><% = Server.HTMLEncode(rstCategories("cname") & "") %></strong></td>
                <td>
                  <% If Len(nm) > 0 Then %>
                    <span class="cat-av"><% = Server.HTMLEncode(ini) %></span><strong><% = Server.HTMLEncode(nm) %></strong>
                    <% If Len(repName) > 0 And Len(repUid) > 0 Then %><span class="cat-rep-sub">&nbsp;(<% = Server.HTMLEncode(repUid) %>)</span><% End If %>
                  <% Else %>
                    <span class="cat-rep-sub">&mdash; sin representante &mdash;</span>
                  <% End If %>
                </td>
                <td class="col-actions">
                  <a class="btn btn-secondary btn-sm" href="modify.asp?mtype=2&amp;id=<% = rstCategories("category_id") %>">&#9998; <%=lang(cnnDB, "Edit")%></a>
                  <a class="btn btn-secondary btn-sm" href="confdelete.asp?mtype=2&amp;id=<% = rstCategories("category_id") %>">&#128465; <%=lang(cnnDB, "Delete")%></a>
                </td>
              </tr>
            <%
                rstCategories.MoveNext
              Loop
            %>
            </tbody>
          </table>
          <div class="cat-empty" id="catEmpty">No se encontraron categor&iacute;as con ese criterio.</div>
        </div>
      </div>
    </div>

    <script type="text/javascript">
    (function () {
      var rows = document.querySelectorAll('#catTable tbody tr');
      var search = document.getElementById('catSearch');
      var empty = document.getElementById('catEmpty');
      var count = document.getElementById('catCount');
      function norm(s) {
        s = (s || '').toLowerCase();
        return s.normalize ? s.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : s;
      }
      function apply() {
        var q = norm(search.value), shown = 0, i;
        for (i = 0; i < rows.length; i++) {
          var ok = !q || norm(rows[i].getAttribute('data-text')).indexOf(q) > -1;
          rows[i].style.display = ok ? '' : 'none';
          if (ok) shown++;
        }
        empty.style.display = shown ? 'none' : 'block';
        count.textContent = shown + ' de ' + rows.length;
      }
      search.oninput = apply;
      apply();
    })();
    </script>

    <%
      rstCategories.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
