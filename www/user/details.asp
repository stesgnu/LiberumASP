<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% 
  Option Explicit
  Response.Buffer = TRUE
  Response.Expires = -1
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: details.asp (user)
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Displays details of a problem.
  -->

  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
    Session.Timeout = 40
  %>

  <head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta charset="utf-8">
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemDetails")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style>
      @media print {
        .no-print { display: none !important; }
        .lhd-detail-layout { grid-template-columns: 1fr !important; }
      }
    </style>
  </head>
  <body>

    <%
      ' See if authenticated
      Call CheckUser(cnnDB, sid)

      Dim uid, ufname
      uid = Usr(cnnDB, sid, "uid")

      ' Get the problem ID
      Dim id
      id = Cint(Request.QueryString("id"))
      If id = 0 Then
        If Len(Request.Form("id")) = 0 Then
          Call DisplayError(3, lang(cnnDB, "AproblemIDnumberisrequired"))
        End If
        id = Cint(Request.Form("id"))
      End If

      Dim queryStr, rstProblem, rstSolution

      queryStr = _
      "SELECT p.id, p.uid, p.uemail, p.uphone, p.ulocation, d.dname, p.start_date, p.status, s.sname, " & _
      "p.close_date, c.cname, pri.pname, r.uid As ruid, r.email1 As remail, r.fname, p.title, p.solution, p.description " & _
      "FROM ((((problems AS p " & _
      "INNER JOIN departments AS d ON p.department = d.department_id) " & _
      "INNER JOIN status AS s ON p.status = s.status_id) " & _
      "INNER JOIN priority AS pri ON p.priority = pri.priority_id) " & _
      "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
      "INNER JOIN categories AS c ON p.category = c.category_id " & _
      "WHERE p.id=" & id & " AND p.uid='" & uid & "'"

      Set rstProblem = SQLQuery(cnnDB, queryStr)

      If rstProblem.EOF Then
        cnnDB.Close
        Call DisplayError(3, lang(cnnDB, "ProblemID") & "&nbsp;" & id & "&nbsp;" & lang(cnnDB, "wasnotfoundinthedatabase") & ".")
      End If

      ' Get solution if closed
      Set rstSolution = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id & " AND uid='" & uid & "'")

      ' Get Notes
      Dim rstNotes
      Set rstNotes = SQLQuery(cnnDB, "SELECT id, addDate, uid, private, [note] FROM tblNotes WHERE id=" & id & " AND private=0 ORDER BY addDate ASC")

      Dim isClosed
      isClosed = (rstProblem("status") >= 100)

      Call DisplayHeader(cnnDB, sid)
    %>

    <div class="lhd-wrapper">

      <!-- Breadcrumb y acciones -->
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px" class="no-print">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128203;</span>
          <%=lang(cnnDB, "DetailsforProblem")%>&nbsp;#<% = id %>
          &nbsp;<% = StatusBadge(rstProblem("sname")) %>
          &nbsp;<% = PriorityBadge(rstProblem("pname")) %>
        </div>
        <div style="display:flex;gap:8px">
          <a class="btn btn-secondary btn-sm" href="view.asp">&#8592; Volver</a>
          <a class="btn btn-secondary btn-sm" href="print.asp?id=<% = id %>" target="printwindow">&#128424; <%=lang(cnnDB, "PrinterFriendly")%></a>
          <% If isClosed Then %>
            <a class="btn btn-primary btn-sm" href="../rep/details.asp?id=<% = id %>&reopen=1">
              &#8635; <%=lang(cnnDB, "ReopenProblem")%>
            </a>
          <% End If %>
        </div>
      </div>

      <% Call ShowToast(cnnDB) %>

      <div class="lhd-detail-layout">

        <!-- Columna principal -->
        <div>
          <!-- Titulo y descripcion -->
          <div class="lhd-card">
            <div class="section-header">&#128221; <%=lang(cnnDB, "ProblemInformation")%></div>

            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "Title")%></label>
              <input type="text" class="form-control" value="<% = FixTildes(Server.HTMLEncode(rstProblem("title"))) %>" readonly>
            </div>

            <div class="form-group" style="margin-bottom:0">
              <label class="form-label"><%=lang(cnnDB, "Description")%></label>
              <div class="lhd-rich" style="background:#F7FAFC;border:1px solid #CBD5E0;border-radius:4px;padding:14px;
                          font-size:13px;line-height:1.7;word-break:break-word;min-height:120px">
                <% = LhdRichToDisplay(rstProblem("description")) %>
              </div>
              <% = LhdAttachmentList(cnnDB, id) %>
            </div>
          </div>

          <!-- Notas / Timeline -->
          <div class="lhd-card">
            <div class="section-header">&#128172; <%=lang(cnnDB, "Notes")%></div>

            <% If rstNotes.EOF Then %>
              <p style="color:#718096;font-size:13px"><%=lang(cnnDB, "NoAvailableNotes")%></p>
            <% Else %>
              <div style="display:flex;gap:8px;justify-content:flex-end;margin-bottom:10px">
                <button type="button" class="btn btn-secondary btn-sm" onclick="lhdNotesToggle(true)">Expandir todas</button>
                <button type="button" class="btn btn-secondary btn-sm" onclick="lhdNotesToggle(false)">Contraer todas</button>
              </div>
              <div style="display:flex;flex-direction:column;gap:10px">
                <%
                  Do While Not rstNotes.EOF
                %>
                  <details class="lhd-note" style="background:#f8fafc;border-left:4px solid var(--color-primary,#0057B8);border-radius:0 6px 6px 0">
                    <summary style="cursor:pointer;padding:10px 16px;font-size:13px;color:#2d3748">
                      &#128100; <% = Server.HTMLEncode(NoteAuthorName(cnnDB, rstNotes("uid"))) %>
                      &nbsp;&mdash;&nbsp;
                      &#128336; <% = DisplayDate(rstNotes("addDate"), lhdDateTime) %>
                    </summary>
                    <div style="padding:0 16px 12px 16px;line-height:1.5;white-space:pre-wrap;color:#2d3748">
                      <% = Server.HTMLEncode(rstNotes("note")) %>
                    </div>
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
            <% rstNotes.Close %>

            <!-- Solucion si cerrado -->
            <% If isClosed Then %>
              <div class="section-subheader">&#10004; <%=lang(cnnDB, "Solution")%></div>
              <div class="lhd-rich" style="background:#F0FFF4;border:1px solid #9AE6B4;border-radius:4px;padding:14px;
                          font-size:13px;line-height:1.7;word-break:break-word">
                <% = LhdRichToDisplay(rstSolution("solution")) %>
              </div>
            <% End If %>
            <% rstSolution.Close %>

            <!-- Formulario para agregar nota si no cerrado -->
            <% If Not isClosed Then %>
              <div class="section-subheader no-print">&#10133; <%=lang(cnnDB, "EnterAdditionalNotes")%></div>
              <form method="post" action="update.asp" class="no-print">
                <input type="hidden" name="id" value="<% = id %>">
                <div class="form-group">
                  <textarea class="form-control" name="notes" rows="5"
                            placeholder="Agrega mas informacion sobre este ticket..."></textarea>
                </div>
                <div style="display:flex;gap:8px;justify-content:flex-end">
                  <button type="reset" class="btn btn-secondary btn-sm">&#10006; <%=lang(cnnDB, "ClearNotes")%></button>
                  <button type="submit" class="btn btn-primary">&#128228; <%=lang(cnnDB, "UpdateProblem")%></button>
                </div>
              </form>
            <% End If %>
          </div>
        </div>

        <!-- Columna lateral: metadatos -->
        <div class="lhd-detail-meta">
          <div class="lhd-detail-meta-header">&#8505; Informacion del ticket</div>
          <table>
            <tr>
              <td>ID</td>
              <td><strong>#<% = id %></strong></td>
            </tr>
            <tr>
              <td><%=lang(cnnDB, "Status")%></td>
              <td><% = StatusBadge(rstProblem("sname")) %></td>
            </tr>
            <tr>
              <td><%=lang(cnnDB, "Priority")%></td>
              <td><% = PriorityBadge(rstProblem("pname")) %></td>
            </tr>
            <tr>
              <td><%=lang(cnnDB, "AssignedTo")%></td>
              <td>
                <a href="mailto:<% = rstProblem("remail") %>?Subject=Help Desk: Ticket #<% = id %>">
                  &#128100; <% = Server.HTMLEncode(rstProblem("fname")) %>
                </a>
              </td>
            </tr>
            <tr>
              <td><%=lang(cnnDB, "Department")%></td>
              <td><% = Server.HTMLEncode(rstProblem("dname")) %></td>
            </tr>
            <tr>
              <td><%=lang(cnnDB, "Category")%></td>
              <td><% = Server.HTMLEncode(rstProblem("cname")) %></td>
            </tr>
            <tr>
              <td><%=lang(cnnDB, "UserName")%></td>
              <td><% = Server.HTMLEncode(LhdUserDisplayName(cnnDB, uid)) %></td>
            </tr>
            <tr>
              <td><%=lang(cnnDB, "EMail")%></td>
              <td><a href="mailto:<% = rstProblem("uemail") %>"><% = rstProblem("uemail") %></a></td>
            </tr>
            <% If Len(Trim(rstProblem("uphone"))) > 0 Then %>
            <tr>
              <td><%=lang(cnnDB, "Phone")%></td>
              <td><% = rstProblem("uphone") %></td>
            </tr>
            <% End If %>
            <% If Len(Trim(rstProblem("ulocation"))) > 0 Then %>
            <tr>
              <td><%=lang(cnnDB, "Location")%></td>
              <td><% = rstProblem("ulocation") %></td>
            </tr>
            <% End If %>
            <tr>
              <td><%=lang(cnnDB, "StartDate")%></td>
              <td><% = DisplayDate(rstProblem("start_date"), lhdDateTime) %></td>
            </tr>
            <% If isClosed Then %>
            <tr>
              <td><%=lang(cnnDB, "CloseDate")%></td>
              <td>&#10004; <% = DisplayDate(rstProblem("close_date"), lhdDateTime) %></td>
            </tr>
            <% End If %>
          </table>
        </div>

      </div>

    </div>

    <%
      rstProblem.Close
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
