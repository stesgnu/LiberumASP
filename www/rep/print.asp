<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
  Response.Expires = -1
%>
<!DOCTYPE html>
<html lang="es">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: print.asp
  Purpose:  Displays a printer friendly version of the problem details
            (estilo nuevo de la aplicacion).
  -->

  <!-- 	#include file = "../public.asp" -->
  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

  ' Look for the problem id in the query string, if it's not
  ' there display an error.
  If Len(Request.QueryString("id")) = 0 Then
    Response.Write("<title>Error</title></head><body>")
    Call DisplayError(3, "No valid problem ID was entered.")
  End If

  ' Sub que escribe un campo "etiqueta / valor" dentro de la cuadricula de datos
  Sub PrnField(strLabel, strValue)
    ' Las etiquetas de lang() ya traen entidades HTML (p. ej. &EACUTE;): no se vuelven a codificar
    Response.Write("<div><div class=""prn-label"">" & (strLabel & "") & "</div>" & _
      "<div class=""prn-value"">" & FixTildes(Server.HTMLEncode(strValue & "")) & "</div></div>" & vbNewLine)
  End Sub
  %>

  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "Problem")%> <% = Server.HTMLEncode(Request.QueryString("id")) %> <%=lang(cnnDB, "Details")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css?v=4">
    <style>
      .prn-grid { display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:14px 22px; }
      .prn-label { font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; color:#718096; }
      .prn-value { font-size:14px; color:#1a202c; word-break:break-word; margin-top:2px; }
      .prn-rich { background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; padding:14px; line-height:1.6; color:#2d3748; }
      @media print {
        body { background:#fff !important; }
        .no-print { display:none !important; }
        .lhd-card { box-shadow:none !important; break-inside:avoid; }
        .prn-rich { max-height:none !important; overflow:visible !important; }
      }
    </style>
  </head>
  <body>

    <%
      ' Check if user has permissions for this page
      Call CheckRep(cnnDB, sid)

      ' Get the problem ID
      Dim id
      id = Cint(Request.QueryString("id"))

      ' Generate a query, making sure to search for id and uid so a user
      ' cannot view someone else's problem.
      Dim queryStr, probRes, rstNotes

      queryStr = _
      "SELECT p.id, p.uid, p.uemail, p.uphone, p.ulocation, p.entered_by, d.dname, p.start_date, p.status, s.sname, " & _
      "p.close_date, c.cname, r.uid AS ruid, r.email1 As remail, r.fname As rname, p.title, p.solution, p.description, " & _
      "pri.pname FROM ((((problems AS p " & _
      "INNER JOIN departments AS d ON p.department = d.department_id) " & _
      "INNER JOIN status AS s ON p.status = s.status_id) " & _
      "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
      "INNER JOIN priority AS pri ON p.priority = pri.priority_id) " & _
      "INNER JOIN categories AS c ON p.category = c.category_id " & _
      "WHERE p.id=" & id
      Set probRes = SQLQuery(cnnDB, queryStr)
      Set rstNotes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " ORDER BY addDate ASC")

      ' If no results are returned, display an error
      If probRes.EOF Then
        Call DisplayError(3, lang(cnnDB, "ProblemID") & " " & id & " " & lang(cnnDB, "wasfoundinthedatabase") & ".")
      End If

      ' If it is a closed problem, get the solution (consulta aparte: SQL solo admite un blob por consulta)
      Dim solution, solRes, isClosed
      solution = ""
      isClosed = (probRes("status") = Cfg(cnnDB, "CloseStatus"))
      If isClosed Then
        Set solRes = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id)
        solution = solRes("solution") & ""
        solRes.Close
      End If
    %>

    <div class="lhd-wrapper">

      <!-- Cabecera -->
      <div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:12px">
        <div>
          <div style="font-size:12px;color:#718096;text-transform:uppercase;letter-spacing:.05em">Incidencia #<% = id %></div>
          <h2 style="margin:2px 0 8px;font-size:20px;font-weight:700;color:#1a202c"><% = FixTildes(Server.HTMLEncode(probRes("title") & "")) %></h2>
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            <% = PriorityBadge(probRes("pname") & "") %>
            <% = StatusBadge(probRes("sname") & "") %>
          </div>
        </div>
        <div class="no-print" style="display:flex;gap:8px;align-items:center">
          <button type="button" class="btn btn-primary btn-sm" onclick="window.print()">&#128424; Imprimir</button>
          <button type="button" class="btn btn-secondary btn-sm" onclick="window.close()"><%=lang(cnnDB, "CloseThisWindow")%></button>
        </div>
      </div>

      <!-- Datos de la solicitud -->
      <div class="lhd-card">
        <div class="section-header">&#128196; Datos de la Solicitud</div>
        <div class="prn-grid">
          <% Call PrnField(lang(cnnDB, "ProblemID"), id) %>
          <% Call PrnField(lang(cnnDB, "UserName"), probRes("uid")) %>
          <% Call PrnField(lang(cnnDB, "EMail"), probRes("uemail")) %>
          <% Call PrnField(lang(cnnDB, "Phone"), probRes("uphone")) %>
          <% Call PrnField(lang(cnnDB, "Location"), probRes("ulocation")) %>
          <% Call PrnField(lang(cnnDB, "EnteredBy"), Usr(cnnDB, probRes("entered_by"), "uid")) %>
          <% Call PrnField(lang(cnnDB, "StartDate"), DisplayDate(probRes("start_date"), lhdDateTime)) %>
          <% Call PrnField(lang(cnnDB, "CloseDate"), DisplayDate(probRes("close_date"), lhdDateTime)) %>
          <% Call PrnField(lang(cnnDB, "AssignedTo"), probRes("rname")) %>
          <% Call PrnField(lang(cnnDB, "Department"), probRes("dname")) %>
          <% Call PrnField(lang(cnnDB, "Category"), probRes("cname")) %>
          <% Call PrnField(lang(cnnDB, "Status"), probRes("sname")) %>
          <% Call PrnField(lang(cnnDB, "Priority"), probRes("pname")) %>
        </div>
      </div>

      <!-- Descripcion y adjuntos -->
      <div class="lhd-card">
        <div class="section-header">&#128221; <%=lang(cnnDB, "Description")%></div>
        <div class="prn-rich lhd-rich"><% = LhdRichToDisplay(probRes("description") & "") %></div>
        <% = LhdAttachmentList(cnnDB, id) %>
      </div>

      <!-- Historial de notas (contraidas por defecto; el agente se muestra por nombre) -->
      <div class="lhd-card">
        <div class="section-header">&#128483; <%=lang(cnnDB, "Notes")%></div>
        <% If rstNotes.EOF Then %>
          <div style="text-align:center;padding:16px;color:#a0aec0;font-style:italic"><%=lang(cnnDB, "NoAvailableNotes")%></div>
        <% Else %>
          <div class="no-print" style="display:flex;gap:8px;justify-content:flex-end;margin-bottom:10px">
            <button type="button" class="btn btn-secondary btn-sm" onclick="lhdNotesToggle(true)">Expandir todas</button>
            <button type="button" class="btn btn-secondary btn-sm" onclick="lhdNotesToggle(false)">Contraer todas</button>
          </div>
          <div style="display:flex;flex-direction:column;gap:10px">
            <% Do While Not rstNotes.EOF %>
              <details class="lhd-note" style="background:#f8fafc;border-left:4px solid <% If rstNotes("private") = 1 Then %>#D69E2E<% Else %>var(--color-primary,#0057B8)<% End If %>;border-radius:0 6px 6px 0">
                <summary style="cursor:pointer;padding:10px 16px;font-size:13px;color:#2d3748">
                  <strong>&#128100; <% = Server.HTMLEncode(NoteAuthorName(cnnDB, rstNotes("uid"))) %></strong>
                  &bull; <% = DisplayDate(rstNotes("addDate"), lhdDateTime) %>
                  <% If rstNotes("private") = 1 Then %> &bull; <% = lang(cnnDB, "PRIVATE") %><% End If %>
                </summary>
                <div style="padding:0 16px 12px 16px;line-height:1.5;color:#2d3748;white-space:pre-wrap"><% = FixTildes(Server.HTMLEncode(rstNotes("note"))) %></div>
              </details>
            <%
                rstNotes.MoveNext
              Loop
            %>
          </div>
          <script type="text/javascript">
            function lhdNotesToggle(abrir) {
              var items = document.querySelectorAll('details.lhd-note');
              for (var i = 0; i < items.length; i++) { items[i].open = abrir; }
            }
          </script>
        <% End If %>
      </div>

      <!-- Solucion (solo si el ticket esta cerrado) -->
      <% If isClosed Then %>
        <div class="lhd-card">
          <div class="section-header">&#9989; <%=lang(cnnDB, "Solution")%></div>
          <div class="prn-rich lhd-rich"><% = LhdRichToDisplay(solution) %></div>
        </div>
      <% End If %>

    </div>

    <%
      ' Close Results
      probRes.Close
      rstNotes.Close
      cnnDB.Close
    %>
  </body>
</html>
