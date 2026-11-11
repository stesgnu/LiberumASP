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

  Filename: print.asp (kb)
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Vista de impresion de un problema de la base de conocimiento.
    Tarjeta limpia, sin barra lateral.
  -->

  <!-- 	#include file = "../public.asp" -->
  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta charset="utf-8">
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "ProblemDetails")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style>
      body { background: #F1F5F9; margin: 0; }
      .pr-page { max-width: 900px; margin: 0 auto; padding: 24px 16px; }
      .pr-bar { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
      .pr-card { background: #fff; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,.08); padding: 28px 32px; }
      .pr-title { font-size: 20px; font-weight: 700; margin: 0 0 4px; color: #1a202c; }
      .pr-sub { font-size: 13px; color: #64748b; margin: 0 0 18px; }
      .pr-meta { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px 24px; padding: 14px 0; border-top: 1px solid #E2E8F0; border-bottom: 1px solid #E2E8F0; margin-bottom: 18px; }
      .pr-l { font-size: 11px; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; }
      .pr-v { font-size: 14px; color: #1a202c; margin-top: 2px; word-break: break-word; }
      .pr-h { font-size: 14px; font-weight: 700; color: var(--color-primary,#0057B8); margin: 18px 0 6px; text-transform: uppercase; letter-spacing: .03em; }
      .pr-text { font-size: 13px; line-height: 1.7; color: #2d3748; word-break: break-word; }
      .pr-note { border-left: 3px solid var(--color-primary,#0057B8); padding: 4px 0 4px 12px; margin-bottom: 10px; page-break-inside: avoid; }
      .pr-note-h { font-size: 12px; font-weight: 700; color: #2d3748; }
      .pr-note-b { font-size: 13px; line-height: 1.5; white-space: pre-wrap; color: #2d3748; }
      @media print {
        body { background: #fff !important; }
        .no-print { display: none !important; }
        .pr-page { max-width: none; padding: 0; }
        .pr-card { box-shadow: none; border-radius: 0; padding: 0; }
      }
    </style>
  </head>
  <body>
    <%
      ' Check permissions
      Call CheckKB(cnnDB, sid)

      ' Get the problem ID
      Dim id
      id = Cint(Request.QueryString("id"))

      ' Generate a query, making sure to search for id and uid so a user
      ' cannot view someone else's problem.
      Dim queryStr, rstProbList, rstNotes

      queryStr = _
      "SELECT p.id, p.uid, p.uemail, p.uphone, p.ulocation, d.dname, p.start_date, p.status, s.sname, " & _
      "p.close_date, c.cname, r.uid As ruid, r.email1 As remail, r.fname, p.title, p.solution, p.description " & _
      "FROM (((problems AS p " & _
      "INNER JOIN departments AS d ON p.department = d.department_id) " & _
      "INNER JOIN status AS s ON p.status = s.status_id) " & _
      "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
      "INNER JOIN categories AS c ON p.category = c.category_id " & _
      "WHERE p.id=" & id & " AND p.kb=1 AND p.status=" & Cfg(cnnDB, "CloseStatus")
      Set rstProbList = SQLQuery(cnnDB, queryStr)
      Set rstNotes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " AND private=0 ORDER BY addDate ASC")

      ' If no results are returned, display an error
      If rstProbList.EOF Then
        cnnDB.Close
        Call DisplayError(3, "Problem ID " & id & " was not found in the database.")
      Else

      Dim solution, solRes
      solution = ""

      ' If it is a closed problem, get the solution
      If rstProbList("status") = Cfg(cnnDB, "CloseStatus") Then
        Set solRes = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id)
        solution = LhdRichToDisplay(solRes("solution"))
        solRes.Close
      End If
    %>

    <div class="pr-page">
      <div class="pr-bar no-print">
        <div class="lhd-page-title" style="margin:0"><span class="title-icon">&#128424;</span> <%=lang(cnnDB, "PrinterFriendly")%></div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button type="button" class="btn btn-secondary" onclick="window.close();"><%=lang(cnnDB, "CloseThisWindow")%></button>
          <button type="button" class="btn btn-primary" onclick="window.print();">&#128424; Imprimir</button>
        </div>
      </div>

      <div class="pr-card">
        <h1 class="pr-title"><% = Server.HTMLEncode(rstProbList("title") & "") %></h1>
        <p class="pr-sub"><%=lang(cnnDB, "DetailsforProblem")%>&nbsp;#<% = id %></p>

        <div class="pr-meta">
          <div><div class="pr-l"><%=lang(cnnDB, "ProblemID")%></div><div class="pr-v">#<% = id %></div></div>
          <div><div class="pr-l"><%=lang(cnnDB, "Category")%></div><div class="pr-v"><% = Server.HTMLEncode(rstProbList("cname") & "") %></div></div>
          <div><div class="pr-l"><%=lang(cnnDB, "StartDate")%></div><div class="pr-v"><% = DisplayDate(rstProbList("start_date"), lhdDateTime) %></div></div>
          <div><div class="pr-l"><%=lang(cnnDB, "CloseDate")%></div><div class="pr-v"><% = DisplayDate(rstProbList("close_date"), lhdDateTime) %></div></div>
        </div>

        <div class="pr-h"><%=lang(cnnDB, "Description")%></div>
        <div class="pr-text lhd-rich"><% = LhdRichToDisplay(rstProbList("description")) %></div>

        <div class="pr-h"><%=lang(cnnDB, "Notes")%></div>
        <% If rstNotes.EOF Then %>
          <div class="pr-text"><%=lang(cnnDB, "NoAvailableNotes")%></div>
        <% Else
             Do While Not rstNotes.EOF %>
          <div class="pr-note">
            <div class="pr-note-h"><% = Server.HTMLEncode(NoteAuthorName(cnnDB, rstNotes("uid"))) %> &mdash; <% = DisplayDate(rstNotes("addDate"), lhdDateTime) %></div>
            <div class="pr-note-b"><% = Server.HTMLEncode(rstNotes("note") & "") %></div>
          </div>
        <%   rstNotes.MoveNext
             Loop
           End If %>

        <% If rstProbList("status") = Cfg(cnnDB, "CloseStatus") Then %>
          <div class="pr-h"><%=lang(cnnDB, "Solution")%></div>
          <div class="pr-text lhd-rich"><% = solution %></div>
        <% End If %>
      </div>
    </div>

    <%
      End If

      ' Close Results
      rstProbList.Close

      cnnDB.Close

    %>
  </body>
</html>
