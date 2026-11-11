<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  Response.Buffer = TRUE

  ' Helpers de este archivo (no existen en public.asp)
  ' HTML-encode seguro: convierte Null/Empty a cadena vacia antes de codificar
  Function ViewHE(v)
    ViewHE = Server.HTMLEncode(v & "")
  End Function

  ' Encabezado de columna ordenable. Clic en la columna activa invierte el orden;
  ' clic en una columna nueva usa su orden por defecto (texto ascendente, resto descendente).
  Function SortHref(col, curSort, curOrder, numVal)
    Dim nextOrder
    If curSort = col Then
      nextOrder = 1 - curOrder
    ElseIf col = 2 Or col = 3 Then
      nextOrder = 1
    Else
      nextOrder = 0
    End If
    SortHref = "view.asp?sort=" & col & "&order=" & nextOrder & "&num=" & numVal & "&start=1&st=" & Server.URLEncode(Request.QueryString("st"))
  End Function

  Function SortIcon(col, curSort, curOrder)
    If curSort = col Then
      If curOrder = 1 Then
        SortIcon = "&#9650;"
      Else
        SortIcon = "&#9660;"
      End If
    Else
      SortIcon = "&#8597;"
    End If
  End Function
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Filename: view.asp (user)
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Lists problems for the user.
  -->

  <!-- #include file = "../public.asp" -->
  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta charset="utf-8" />
  <title><% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemList")%></title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="stylesheet" type="text/css" href="../default.css?v=4" />
<style type="text/css">
  .lhd-tabs { display: inline-flex; gap: 4px; padding: 4px; background: #E2E8F0; border-radius: 10px; margin-bottom: 14px; flex-wrap: wrap; }
  .lhd-tab { padding: 7px 16px; border-radius: 8px; font-size: 13px; font-weight: 600; color: #475569; text-decoration: none !important; display: inline-flex; align-items: center; gap: 8px; transition: all .15s ease; }
  .lhd-tab:hover { color: var(--color-primary, #0057B8); }
  .lhd-tab.on { background: #fff; color: var(--color-primary, #0057B8); box-shadow: 0 1px 3px rgba(0,0,0,.12); }
  .lhd-tab .n { background: #CBD5E0; color: #334155; border-radius: 9999px; font-size: 11px; padding: 1px 8px; }
  .lhd-tab.on .n { background: var(--color-primary, #0057B8); color: #fff; }
</style>
</head>
<body>

<%
  ' See if authenticated (redirects to logon if not)
  Call CheckUser(cnnDB, sid)
  Call DisplayHeader(cnnDB, sid)

  Dim uid, uidSql
  uid = Usr(cnnDB, sid, "uid")
  If IsNull(uid) Then uid = ""
  uidSql = Replace(uid, "'", "''")

  ' Parametros
  Dim intSort, intOrder, numToDisplay, start
  intSort = SafeCInt(Request.QueryString("sort"))
  intOrder = SafeCInt(Request.QueryString("order"))
  If intOrder <> 1 Then intOrder = 0

  numToDisplay = SafeCInt(Request.QueryString("num"))
  If numToDisplay < 1 Then numToDisplay = 25
  If numToDisplay > 200 Then numToDisplay = 200

  start = SafeCInt(Request.QueryString("start"))
  If start < 1 Then start = 1

  ' Consulta
  Dim whereSql, sortCol, orderSql
  whereSql = " WHERE p.uid='" & uidSql & "'"

  ' Filtro de estado: todos (por omision), abiertos o cerrados
  Dim stFilter
  stFilter = LCase(Trim(Request.QueryString("st")))
  If stFilter <> "open" And stFilter <> "closed" Then stFilter = "all"
  If stFilter = "open" Then whereSql = whereSql & " AND p.status < 100"
  If stFilter = "closed" Then whereSql = whereSql & " AND p.status >= 100"
  If Len(Request.QueryString("id")) > 0 And IsNumeric(Request.QueryString("id")) Then
    whereSql = whereSql & " AND p.id=" & SafeCInt(Request.QueryString("id"))
  End If

  Select Case intSort
    Case 1 : sortCol = "p.id"
    Case 2 : sortCol = "p.title"
    Case 3 : sortCol = "r.fname"
    Case 4 : sortCol = "p.start_date"
    Case 5 : sortCol = "p.status"
    Case Else : sortCol = ""
  End Select

  If sortCol = "" Then
    orderSql = " ORDER BY p.id DESC"
  ElseIf intOrder = 1 Then
    orderSql = " ORDER BY " & sortCol & " ASC"
  Else
    orderSql = " ORDER BY " & sortCol & " DESC"
  End If

  ' Conteos para las pestanas
  Dim cntU, cntUOpen, cntUClosed
  Set cntU = SQLQuery(cnnDB, "SELECT count(*) AS total FROM problems WHERE uid='" & uidSql & "' AND status < 100")
  cntUOpen = SafeCInt(cntU("total"))
  cntU.Close
  Set cntU = SQLQuery(cnnDB, "SELECT count(*) AS total FROM problems WHERE uid='" & uidSql & "' AND status >= 100")
  cntUClosed = SafeCInt(cntU("total"))
  cntU.Close

  Dim listStr, rstProbList
  listStr = "SELECT TOP " & (start + numToDisplay) & " p.id, p.title, p.start_date, r.fname, r.email1 AS remail, s.sname, pri.pname " & _
    "FROM problems AS p " & _
    "LEFT JOIN tblUsers AS r ON p.rep = r.sid " & _
    "INNER JOIN status AS s ON p.status = s.status_id " & _
    "INNER JOIN priority AS pri ON p.priority = pri.priority_id" & _
    whereSql & orderSql

  Set rstProbList = SQLQuery(cnnDB, listStr)

  ' Saltar las filas anteriores a "start" (start es un indice 1-based)
  Dim rowIdx, firstRowExists, shown
  rowIdx = 1
  Do While Not rstProbList.EOF And rowIdx < start
    rstProbList.MoveNext
    rowIdx = rowIdx + 1
  Loop
  firstRowExists = Not rstProbList.EOF

  Dim prevStart, nextStart
  prevStart = start - numToDisplay
  If prevStart < 1 Then prevStart = 1
  nextStart = start + numToDisplay
%>

<div class="lhd-wrapper">

  <div style="margin-bottom:20px">
    <div class="lhd-page-title">
      <span class="title-icon">&#128203;</span>
      <%=lang(cnnDB, "ProblemListingfor")%>&nbsp;<%=ViewHE(uid)%>
    </div>
  </div>

  <% Call ShowToast(cnnDB) %>

  <!-- Pestanas de estado: Abiertos / Cerrados / Todos -->
  <div class="lhd-tabs">
    <a class="lhd-tab<% If stFilter = "open" Then Response.Write(" on") %>" href="view.asp?st=open">Abiertos <span class="n"><% = cntUOpen %></span></a>
    <a class="lhd-tab<% If stFilter = "closed" Then Response.Write(" on") %>" href="view.asp?st=closed">Cerrados <span class="n"><% = cntUClosed %></span></a>
    <a class="lhd-tab<% If stFilter = "all" Then Response.Write(" on") %>" href="view.asp?st=all">Todos <span class="n"><% = cntUOpen + cntUClosed %></span></a>
  </div>

  <div class="lhd-table-filter">
    <span style="font-size:13px;color:#4A5568">&#128269;</span>
    <input type="text" id="tbl-filter" placeholder="Filtrar tickets..." autocomplete="off" aria-label="Filtrar tickets">
    <a class="btn btn-primary btn-sm" href="new.asp">&#10133; Nuevo ticket</a>
  </div>

  <div class="lhd-card" style="padding:0;overflow:hidden">
    <div style="overflow-x:auto">
      <table class="lhd-table" id="tickets-table">
        <thead>
          <tr>
            <th class="col-id">
              <a href="<%=SortHref(1, intSort, intOrder, numToDisplay)%>"><%=lang(cnnDB, "ID")%> <%=SortIcon(1, intSort, intOrder)%></a>
            </th>
            <th>
              <a href="<%=SortHref(2, intSort, intOrder, numToDisplay)%>"><%=lang(cnnDB, "Title")%> <%=SortIcon(2, intSort, intOrder)%></a>
            </th>
            <th>
              <a href="<%=SortHref(3, intSort, intOrder, numToDisplay)%>"><%=lang(cnnDB, "AssignedTo")%> <%=SortIcon(3, intSort, intOrder)%></a>
            </th>
            <th class="hide-mobile">
              <a href="<%=SortHref(4, intSort, intOrder, numToDisplay)%>"><%=lang(cnnDB, "DateSubmitted")%> <%=SortIcon(4, intSort, intOrder)%></a>
            </th>
            <th>
              <a href="<%=SortHref(5, intSort, intOrder, numToDisplay)%>"><%=lang(cnnDB, "Status")%> <%=SortIcon(5, intSort, intOrder)%></a>
            </th>
            <th>Prioridad</th>
          </tr>
        </thead>
        <tbody>
<%
  If Not firstRowExists Then
%>
          <tr>
            <td colspan="6" style="text-align:center;padding:32px;color:#718096">
              <div style="font-size:40px;margin-bottom:10px">&#128203;</div>
              <strong><%=lang(cnnDB, "Noresultsfound")%></strong><br>
              <a class="btn btn-primary btn-sm" href="new.asp" style="margin-top:12px;display:inline-flex">&#10133; Crear primer ticket</a>
            </td>
          </tr>
<%
  End If

  shown = 0
  Do While Not rstProbList.EOF And shown < numToDisplay
    Dim remail, fname
    remail = rstProbList("remail") & ""
    fname  = rstProbList("fname") & ""
%>
          <tr>
            <td class="col-id">
              <a href="details.asp?id=<%=rstProbList("id")%>"><%=rstProbList("id")%></a>
            </td>
            <td>
              <a href="details.asp?id=<%=rstProbList("id")%>"><%=ViewHE(rstProbList("title"))%></a>
            </td>
            <td>
<%
    If Len(remail) > 0 Then
%>
              <a href="mailto:<%=ViewHE(remail)%>?Subject=HELPDESK: Ticket <%=rstProbList("id")%>"><%=ViewHE(fname)%></a>
<%
    Else
%>
              <%=ViewHE(fname)%>
<%
    End If
%>
            </td>
            <td class="col-date hide-mobile">
              <%=DisplayDate(rstProbList("start_date"), lhdDateOnly)%>
            </td>
            <td><%=StatusBadge(rstProbList("sname") & "")%></td>
            <td><%=PriorityBadge(rstProbList("pname") & "")%></td>
          </tr>
<%
    shown = shown + 1
    rstProbList.MoveNext
  Loop
%>
        </tbody>
      </table>
    </div>

    <% If start > 1 Or Not rstProbList.EOF Then %>
    <div class="lhd-pagination" style="padding:16px">
      <% If start > 1 Then %>
        <a href="view.asp?start=<%=prevStart%>&num=<%=numToDisplay%>&sort=<%=intSort%>&order=<%=intOrder%>&st=<%=stFilter%>">&laquo; <%=lang(cnnDB, "Previous")%></a>
      <% Else %>
        <span class="pg-disabled">&laquo; <%=lang(cnnDB, "Previous")%></span>
      <% End If %>
      <% If Not rstProbList.EOF Then %>
        <a href="view.asp?start=<%=nextStart%>&num=<%=numToDisplay%>&sort=<%=intSort%>&order=<%=intOrder%>&st=<%=stFilter%>"><%=lang(cnnDB, "Next")%> &raquo;</a>
      <% Else %>
        <span class="pg-disabled"><%=lang(cnnDB, "Next")%> &raquo;</span>
      <% End If %>
    </div>
    <% End If %>
  </div>

</div>

<%
  rstProbList.Close
  Call DisplayFooter(cnnDB, sid)
  cnnDB.Close
%>
</body>
</html>
