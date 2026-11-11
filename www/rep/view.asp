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

  Filename: view.asp (rep)
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.2.2.2.1 $
  Purpose:  Technician view of open problems.
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
  <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "OpenProblems")%></title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" type="text/css" href="../default.css">
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
	' Check if user has permissions for this page
	Call CheckRep(cnnDB, sid)
	Call DisplayHeader(cnnDB, sid)

	' El tecnico solo ve sus propios tickets: no se acepta rep_id desde la URL o el formulario
	Dim rep_id
	rep_id = SafeCInt(sid)

	' Get the rep's name
	Dim ruid, repRes
  Set repRes = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & rep_id)
  If Not repRes.EOF Then
	  ruid = repRes("uid")
  Else
    ruid = ""
  End If
  
	' Query database
	Dim listStr, cntStr, listRes
	listStr = "SELECT TOP 100 p.id, p.title, p.start_date, p.uid, p.uemail, r.uid AS ruid, pri.pname, s.sname, " & _
	"u.sid AS usid, u.statuscode AS ustatus " & _
	"FROM problems AS p " & _
	"LEFT JOIN tblUsers AS r ON p.rep = r.sid " & _
	"LEFT JOIN tblUsers AS u ON p.uid = u.uid " & _
	"INNER JOIN priority AS pri ON p.priority = pri.priority_id " & _
	"INNER JOIN status AS s ON p.status = s.status_id " & _
	"WHERE"

	cntStr = "SELECT count(*) AS total FROM problems WHERE"
	Dim disp_total

	' Filtro de estado: abiertos (por omision), cerrados o todos
	Dim stFilter
	stFilter = LCase(Trim(Request.QueryString("st")))
	If stFilter <> "closed" And stFilter <> "all" Then stFilter = "open"

	' If problem ID entered, search only for that
	If Len(Request.QueryString("id")) > 0 And IsNumeric(Request.QueryString("id")) Then
		listStr = listStr & " p.id=" & SafeCInt(Request.QueryString("id"))
		disp_total = FALSE
	Else
		Select Case stFilter
			Case "closed"
				listStr = listStr & " p.status >= 100 AND p.rep=" & rep_id
				cntStr = cntStr & " status >= 100 AND rep=" & rep_id
			Case "all"
				listStr = listStr & " p.rep=" & rep_id
				cntStr = cntStr & " rep=" & rep_id
			Case Else
				listStr = listStr & " p.status < 100 AND p.rep=" & rep_id
				cntStr = cntStr & " status < 100 AND rep=" & rep_id
		End Select
		disp_total = TRUE
	End If

  ' Determine Sort Order
  Dim intSort, intOrder, intIDOrder, intTitleOrder, intUIDOrder, intDateOrder, intPriOrder, intStatusOrder
  intSort = SafeCInt(Request.QueryString("sort"))
  If Len(Request.QueryString("order")) > 0 And IsNumeric(Request.QueryString("order")) Then
    intOrder = SafeCInt(Request.QueryString("order"))
  Else
    intOrder = 0
  End If
  Select Case intSort
    Case 1
      listStr = listStr & " ORDER BY p.id"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intIDOrder = 1
      Else
        listStr = listStr & " ASC"
        intIDOrder = 0
      End If
    Case 2
      listStr = listStr & " ORDER BY p.title"
      If intOrder = 0 Then
        listStr = listStr & " ASC"
        intTitleOrder = 1
      Else
        listStr = listStr & " DESC"
        intTitleOrder = 0
      End If
    Case 3
      listStr = listStr & " ORDER BY p.uid"
      If intOrder = 0 Then
        listStr = listStr & " ASC"
        intUIDOrder = 1
      Else
        listStr = listStr & " DESC"
        intUIDOrder = 0
      End If
    Case 4
      listStr = listStr & " ORDER BY p.start_date"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intDateOrder = 1
      Else
        listStr = listStr & " ASC"
        intDateOrder = 0
      End If
    Case 5
      listStr = listStr & " ORDER BY p.priority"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intPriOrder = 1
      Else
        listStr = listStr & " ASC"
        intPriOrder = 0
      End If
    Case 6
      listStr = listStr & " ORDER BY p.status"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intStatusOrder = 1
      Else
        listStr = listStr & " ASC"
        intStatusOrder = 0
      End If
    Case Else
      listStr = listStr & " ORDER BY p.priority DESC, p.id DESC"
      intIDOrder = 1
  End Select

  Set listRes = SQLQuery(cnnDB, listStr)

	' Total count
  Dim cntRes, totalCount
  totalCount = 0
	If disp_total Then
		Set cntRes = SQLQuery(cnnDB, cntStr)
    totalCount = cntRes("total")
    cntRes.Close
	End If

	Dim Counter, numToDisplay, start
	Counter = 1
	If Len(Request.QueryString("num")) > 0 Then
		numToDisplay = SafeCInt(Request.QueryString("num"))
	Else
		numToDisplay = 25
	End if
	If Len(Request.QueryString("start")) > 0 Then
		start = SafeCInt(Request.QueryString("start"))
	Else
		start = 1
	End if

  Dim intUseInoutBoard
  intUseInoutBoard = Cfg(cnnDB, "UseInoutBoard")

  ' Conteos para las pestanas Abiertos / Cerrados
  Dim cntTabRes, cntTabOpen, cntTabClosed
  Set cntTabRes = SQLQuery(cnnDB, "SELECT count(*) AS total FROM problems WHERE status < 100 AND rep=" & rep_id)
  cntTabOpen = SafeCInt(cntTabRes("total"))
  cntTabRes.Close
  Set cntTabRes = SQLQuery(cnnDB, "SELECT count(*) AS total FROM problems WHERE status >= 100 AND rep=" & rep_id)
  cntTabClosed = SafeCInt(cntTabRes("total"))
  cntTabRes.Close
%>

<div class="lhd-wrapper">

  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:12px">
    <div class="lhd-page-title" style="margin:0">
      <span class="title-icon">&#128203;</span>
      <%=lang(cnnDB, "Problemsfor")%>&nbsp;<strong><% = Server.HTMLEncode(ruid) %></strong>
      <% If disp_total Then %>
        <span class="badge badge-open" style="font-size:13px;margin-left:8px"><% = totalCount %> <% If stFilter = "closed" Then %>cerrados<% ElseIf stFilter = "all" Then %>en total<% Else %>pendientes<% End If %></span>
      <% End If %>
    </div>
    <div style="display:flex;gap:8px">
      <a class="btn btn-secondary btn-sm" href="default.asp">&#8592; Panel</a>
      <a class="btn btn-primary btn-sm" href="new.asp">&#10133; <%=lang(cnnDB, "SubmitNewProblem")%></a>
    </div>
  </div>

  <% Call ShowToast(cnnDB) %>

  <!-- Pestanas de estado: Abiertos / Cerrados / Todos -->
  <div class="lhd-tabs">
    <a class="lhd-tab<% If stFilter = "open" Then Response.Write(" on") %>" href="view.asp?st=open">Abiertos <span class="n"><% = cntTabOpen %></span></a>
    <a class="lhd-tab<% If stFilter = "closed" Then Response.Write(" on") %>" href="view.asp?st=closed">Cerrados <span class="n"><% = cntTabClosed %></span></a>
    <a class="lhd-tab<% If stFilter = "all" Then Response.Write(" on") %>" href="view.asp?st=all">Todos <span class="n"><% = cntTabOpen + cntTabClosed %></span></a>
  </div>

  <!-- Barra de filtro de tabla en cliente y selector de rep -->
  <div class="lhd-table-filter" style="justify-content:space-between">
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <span style="font-size:14px;color:#718096">&#128269;</span>
      <input type="text" id="tbl-filter" placeholder="Filtrar en esta p&aacute;gina..." autocomplete="off">
    </div>
  </div>

  <div class="lhd-card" style="padding:0;overflow:hidden">
    <div style="overflow-x:auto">
      <table class="lhd-table" id="rep-tickets-table">
        <thead>
          <tr>
            <th class="col-id">
              <a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=1&order=<% = intIDOrder %>&st=<% = stFilter %>">
                <%=lang(cnnDB, "ID")%> &#8597;
              </a>
            </th>
            <th>
              <a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=2&order=<% = intTitleOrder %>&st=<% = stFilter %>">
                <%=lang(cnnDB, "Title")%> &#8597;
              </a>
            </th>
            <th>
              <a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=3&order=<% = intUIDOrder %>&st=<% = stFilter %>">
                <%=lang(cnnDB, "UserName")%> &#8597;
              </a>
            </th>
            <% If intUseInoutBoard = 1 Then %>
              <th style="width:40px;text-align:center" title="Estado In/Out">I/O</th>
            <% End If %>
            <th class="hide-mobile">
              <a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=4&order=<% = intDateOrder %>&st=<% = stFilter %>">
                <%=lang(cnnDB, "DateSubmitted")%> &#8597;
              </a>
            </th>
            <th>
              <a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=5&order=<% = intPriOrder %>&st=<% = stFilter %>">
                <%=lang(cnnDB, "Priority")%> &#8597;
              </a>
            </th>
            <th>
              <a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=6&order=<% = intStatusOrder %>&st=<% = stFilter %>">
                <%=lang(cnnDB, "Status")%> &#8597;
              </a>
            </th>
          </tr>
        </thead>
        <tbody>
          <%
            If Not listRes.EOF Then
              Do While Not (listRes.EOF) AND (Counter <= (numToDisplay + start - 1))
                If Counter >= start Then
          %>
                  <tr>
                    <td class="col-id">
                      <a href="details.asp?id=<% = listRes("id") %>">#<% = listRes("id") %></a>
                    </td>
                    <td>
                      <a href="details.asp?id=<% = listRes("id") %>">
                        <% = FixTildes(Server.HTMLEncode(listRes("title"))) %>
                      </a>
                    </td>
                    <td>
                      <a href="mailto:<% = listRes("uemail") %>?Subject=HELPDESK: Problem <% = listRes("id") %>">
                        <% = FixTildes(Server.HTMLEncode(listRes("uid"))) %>
                      </a>
                    </td>
                    <% If intUseInoutBoard = 1 Then %>
                      <td style="text-align:center">
                        <%
                          If Not IsNull(listRes("usid")) Then
                            Response.Write "<a href=""../inout/details.asp?id=" & listRes("usid") & """>" & InoutIcon(listRes("ustatus")) & "</a>"
                          Else
                            Response.Write "&nbsp;"
                          End If
                        %>
                      </td>
                    <% End If %>
                    <td class="col-date hide-mobile"><% = DisplayDate(listRes("start_date"), lhdDateOnly) %></td>
                    <td><% = PriorityBadge(FixTildes(listRes("pname"))) %></td>
                    <td><% = StatusBadge(FixTildes(listRes("sname"))) %></td>
                  </tr>
          <%
                End If
                Counter = Counter + 1
                listRes.MoveNext
              Loop
            Else
          %>
              <tr>
                <td colspan="7" style="text-align:center;padding:36px;color:#718096">
                  <div style="font-size:36px;margin-bottom:8px">&#10004;</div>
                  <strong><%=lang(cnnDB, "Noresultsfound")%></strong>
                </td>
              </tr>
          <% End If %>
        </tbody>
      </table>
    </div>

    <!-- Paginacion -->
    <%
      If start > 1 Or Not listRes.EOF Then
        Dim startP, startN
        startP = start - numToDisplay
        If startP < 1 Then startP = 1
        startN = start + numToDisplay
    %>
      <div class="lhd-pagination" style="padding:16px">
        <% If start > 1 Then %>
          <a href="view.asp?start=<% = startP %>&num=<% = numToDisplay %>&sort=<% = intSort %>&order=<% = intOrder %>&st=<% = stFilter %>">&laquo; <%=lang(cnnDB, "Previous")%></a>
        <% End If %>
        <% If Not listRes.EOF Then %>
          <a href="view.asp?start=<% = startN %>&num=<% = numToDisplay %>&sort=<% = intSort %>&order=<% = intOrder %>&st=<% = stFilter %>"><%=lang(cnnDB, "Next")%> &raquo;</a>
        <% End If %>
      </div>
    <% End If %>
  </div>

</div>

<script>
document.getElementById('tbl-filter').addEventListener('input', function() {
  var q = this.value.toLowerCase();
  var rows = document.querySelectorAll('#rep-tickets-table tbody tr');
  rows.forEach(function(row) {
    row.style.display = row.textContent.toLowerCase().indexOf(q) >= 0 ? '' : 'none';
  });
});
</script>

<%
	listRes.Close
	repRes.Close

	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close
%>

</body>
</html>
