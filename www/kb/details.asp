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

  Filename: details.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  This page displays the details of a problem.  The problem id is taken
  is an input from the URL. (details.asp?id=xx)  Modernized UX. Con sesion usa el
  encabezado del sistema; sin sesion muestra un encabezado simple sin barra lateral.
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
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "ProblemDetails")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style>
      .kb-simple-bar { background: linear-gradient(135deg, var(--color-primary,#0057B8), var(--color-accent,#0082C8)); color: #fff; padding: 14px 24px; display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
      .kb-simple-bar .kb-brand { font-size: 17px; font-weight: 700; }
      .kb-simple-bar a { color: #fff; font-size: 13px; text-decoration: underline; }
      .kb-simple-wrap { max-width: 1000px; margin: 0 auto; padding: 24px 16px; }
      .kb-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px 18px; }
      .kb-meta-l { font-size: 12px; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: .03em; }
      .kb-meta-v { font-size: 14px; color: #1a202c; margin-top: 2px; word-break: break-word; }
      .kb-rich { background: #F7FAFC; border: 1px solid #CBD5E0; border-radius: 4px; padding: 14px; font-size: 13px; line-height: 1.7; word-break: break-word; min-height: 60px; }
      .kb-sol { background: #F0FFF4; border-color: #9AE6B4; }
      .kb-note { background: #f8fafc; border-left: 4px solid var(--color-primary,#0057B8); border-radius: 0 6px 6px 0; }
      .kb-note summary { cursor: pointer; padding: 10px 16px; font-size: 13px; color: #2d3748; }
      .kb-note div { padding: 0 16px 12px 16px; line-height: 1.5; white-space: pre-wrap; color: #2d3748; }
    </style>
  </head>
  <body>

    <%
      ' See if authenticated
      Call CheckKB(cnnDB, sid)

      Dim hasSession
      hasSession = (sid > 0)

      ' Get the problem ID
      Dim id
      id = Cint(Request.QueryString("id"))
      If id = 0 Then
        If Len(Request.Form("id")) = 0 Then
          Call DisplayError(3, "A problem ID number is required.")
        End If
        id = Cint(Request.Form("id"))
      End If

      ' Generate a query, making sure to search for id and uid so a user
      ' cannot view someone else's problem.
      Dim queryStr, rstProblem, rstSolution

      queryStr = _
      "SELECT p.id, p.uid, p.uemail, p.uphone, p.ulocation, d.dname, p.start_date, p.status, s.sname, " & _
      "p.close_date, c.cname, r.uid As ruid, r.email1 As remail, r.fname, p.title, p.solution, p.description " & _
      "FROM (((problems AS p " & _
      "INNER JOIN departments AS d ON p.department = d.department_id) " & _
      "INNER JOIN status AS s ON p.status = s.status_id) " & _
      "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
      "INNER JOIN categories AS c ON p.category = c.category_id " & _
      "WHERE p.id=" & id & " AND p.kb=1 AND p.status=" & Cfg(cnnDB, "CloseStatus")

      Set rstProblem = SQLQuery(cnnDB, queryStr)

      ' If no results are returned, display an error
      If rstProblem.EOF Then
        cnnDB.Close
        Call DisplayError(3, "Problem ID " & id & " was not found in the database.")
      End If

      ' If it is a closed problem, get the solution
      Set rstSolution = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id)

      ' Get the Notes for this problem
      Dim rstNotes
      Set rstNotes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " AND private=0 ORDER BY addDate ASC")

      If hasSession Then
        Call DisplayHeader(cnnDB, sid)
        Response.Write "<div class=""lhd-wrapper"">" & vbNewLine
      Else
        Response.Write "<div class=""kb-simple-bar""><span class=""kb-brand"">&#128218; " & Server.HTMLEncode(Cfg(cnnDB, "SiteName") & "") & " &mdash; " & lang(cnnDB, "HelpDesk") & "</span>" & _
          "<a href=""../logon.asp"">" & lang(cnnDB, "Logon") & "</a></div>" & vbNewLine
        Response.Write "<div class=""kb-simple-wrap"">" & vbNewLine
      End If
    %>

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0"><span class="title-icon">&#128218;</span> <%=lang(cnnDB, "DetailsforProblem")%>&nbsp;#<% = id %></div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "SearchAgain")%></a>
          <a class="btn btn-secondary" href="print.asp?id=<% = id %>" target="printwindow">&#128424; <%=lang(cnnDB, "PrinterFriendly")%></a>
        </div>
      </div>

      <div class="lhd-card">
        <div class="section-header">&#8505; <%=lang(cnnDB, "ProblemInformation")%></div>
        <div class="kb-grid" style="margin-bottom:16px">
          <div><div class="kb-meta-l"><%=lang(cnnDB, "ProblemID")%></div><div class="kb-meta-v">#<% = id %></div></div>
          <div><div class="kb-meta-l"><%=lang(cnnDB, "StartDate")%></div><div class="kb-meta-v"><% = DisplayDate(rstProblem("start_date"), lhdDateTime) %></div></div>
          <div><div class="kb-meta-l"><%=lang(cnnDB, "CloseDate")%></div><div class="kb-meta-v"><% = DisplayDate(rstProblem("close_date"), lhdDateTime) %></div></div>
          <div><div class="kb-meta-l"><%=lang(cnnDB, "Category")%></div><div class="kb-meta-v"><% = Server.HTMLEncode(rstProblem("cname") & "") %></div></div>
        </div>

        <div class="form-group">
          <label class="form-label"><%=lang(cnnDB, "Title")%></label>
          <input type="text" class="form-control" value="<% = Server.HTMLEncode(rstProblem("title") & "") %>" readonly>
        </div>

        <div class="form-group" style="margin-bottom:0">
          <label class="form-label"><%=lang(cnnDB, "Description")%></label>
          <div class="kb-rich lhd-rich"><% = LhdRichToDisplay(rstProblem("description")) %></div>
        </div>
      </div>

      <div class="lhd-card">
        <div class="section-header">&#128172; <%=lang(cnnDB, "Notes")%></div>
        <% If rstNotes.EOF Then %>
          <p style="color:#718096;font-size:13px;margin:0"><%=lang(cnnDB, "NoAvailableNotes")%></p>
        <% Else %>
          <div style="display:flex;flex-direction:column;gap:10px">
            <% Do While Not rstNotes.EOF %>
              <details class="kb-note" open>
                <summary>&#128336; <% = DisplayDate(rstNotes("addDate"), lhdDateTime) %></summary>
                <div><% = Server.HTMLEncode(rstNotes("note") & "") %></div>
              </details>
            <%
                rstNotes.MoveNext
              Loop
            %>
          </div>
        <% End If %>
      </div>

      <div class="lhd-card">
        <div class="section-header">&#10004; <%=lang(cnnDB, "Solution")%></div>
        <div class="kb-rich kb-sol lhd-rich"><% = LhdRichToDisplay(rstSolution("solution")) %></div>
      </div>

    </div>

    <%
      ' Close Results
      rstProblem.Close
      rstNotes.Close
      rstSolution.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
