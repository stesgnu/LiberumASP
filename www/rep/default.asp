<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% 
  Option Explicit
  Response.Buffer = TRUE
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Filename: default.asp (rep)
  Purpose:  Technician / Representative Dashboard
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
  <title><% = Cfg(cnnDB, "SiteName") %> <%=lang(cnnDB, "HelpDesk")%></title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="stylesheet" type="text/css" href="../default.css?v=4" />
</head>
<body>

<%
  Call CheckRep(cnnDB, sid)
  Call DisplayHeader(cnnDB, sid)

  Dim repUname, repFname
  repUname = Usr(cnnDB, sid, "uid")
  repFname = Usr(cnnDB, sid, "fname")
  If IsNull(repFname) Or Len(Trim(repFname)) = 0 Then repFname = repUname

  ' KPI 1: Assigned open tickets
  Dim rstOpen, countOpen
  Set rstOpen = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM problems WHERE rep=" & SafeCInt(sid) & " AND status < 100")
  countOpen = rstOpen("total")
  rstOpen.Close

  ' KPI 2: Urgent / High priority open tickets
  Dim rstUrgent, countUrgent
  Set rstUrgent = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM problems WHERE rep=" & SafeCInt(sid) & " AND status < 100 AND priority >= 3")
  countUrgent = rstUrgent("total")
  rstUrgent.Close

  ' KPI 3: Normal / Standard priority open tickets
  Dim rstNormal, countNormal
  Set rstNormal = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM problems WHERE rep=" & SafeCInt(sid) & " AND status < 100 AND priority < 3")
  countNormal = rstNormal("total")
  rstNormal.Close

  ' KPI 4: Closed / Resolved tickets
  Dim rstClosed, countClosed
  Set rstClosed = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM problems WHERE rep=" & SafeCInt(sid) & " AND status >= 100")
  countClosed = rstClosed("total")
  rstClosed.Close

  ' Query: open tickets from system (up to 15 to ensure at least 5 tickets are displayed)
  Dim rstRecent
  Set rstRecent = SQLQuery(cnnDB, "SELECT TOP 15 p.id, p.title, p.uid, p.start_date, p.rep, s.sname, pri.pname " & _
    "FROM problems AS p " & _
    "LEFT JOIN status AS s ON p.status = s.status_id " & _
    "LEFT JOIN priority AS pri ON p.priority = pri.priority_id " & _
    "WHERE p.status < 100 AND p.rep=" & SafeCInt(sid) & " " & _
    "ORDER BY p.id DESC")
%>

<div class="lhd-wrapper">

  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
    <div class="lhd-page-title" style="margin:0">
      <span class="title-icon">&#128202;</span>
      Panel de Soporte &mdash; <% = FixTildes(Server.HTMLEncode(repFname)) %>
    </div>
    <div style="display:flex;gap:8px">
      <a class="btn btn-primary" href="new.asp">&#10133; Nuevo Ticket</a>
      <a class="btn btn-secondary" href="view.asp">&#128203; Todos Mis Tickets</a>
    </div>
  </div>

  <% Call ShowToast(cnnDB) %>

  <!-- KPI Grid con Iconos y Diseno Moderno -->
  <div class="lhd-kpi-grid">
    <div class="lhd-kpi-card kpi-open" style="display:flex;align-items:center;justify-content:space-between">
      <div>
        <div class="kpi-value"><% = countOpen %></div>
        <div class="kpi-label">Tickets Abiertos</div>
      </div>
      <div style="width:44px;height:44px;border-radius:10px;background:var(--color-primary-light,#EFF6FF);color:var(--color-primary,#0057B8);display:flex;align-items:center;justify-content:center;font-size:22px">&#128194;</div>
    </div>
    <div class="lhd-kpi-card kpi-urgent" style="display:flex;align-items:center;justify-content:space-between">
      <div>
        <div class="kpi-value" style="color:var(--color-danger)"><% = countUrgent %></div>
        <div class="kpi-label">Prioridad Alta / Cr&iacute;tica</div>
      </div>
      <div style="width:44px;height:44px;border-radius:10px;background:#FEF2F2;color:#EF4444;display:flex;align-items:center;justify-content:center;font-size:22px">&#128293;</div>
    </div>
    <div class="lhd-kpi-card kpi-pending" style="display:flex;align-items:center;justify-content:space-between">
      <div>
        <div class="kpi-value" style="color:#D69E2E"><% = countNormal %></div>
        <div class="kpi-label">Prioridad Normal</div>
      </div>
      <div style="width:44px;height:44px;border-radius:10px;background:#FFFBEB;color:#F59E0B;display:flex;align-items:center;justify-content:center;font-size:22px">&#9201;</div>
    </div>
    <div class="lhd-kpi-card kpi-closed" style="display:flex;align-items:center;justify-content:space-between">
      <div>
        <div class="kpi-value" style="color:var(--color-success)"><% = countClosed %></div>
        <div class="kpi-label">Resueltos Totales</div>
      </div>
      <div style="width:44px;height:44px;border-radius:10px;background:#ECFDF5;color:#10B981;display:flex;align-items:center;justify-content:center;font-size:22px">&#127942;</div>
    </div>
  </div>

  <!-- Acciones rapidas con diseno enriquecido -->
  <div class="lhd-action-grid">
    <a class="lhd-action-card" href="new.asp">
      <span class="action-icon">&#10133;</span>
      <span class="action-label"><%=lang(cnnDB, "SubmitNewProblem")%></span>
      <span class="action-desc">Crear incidencia</span>
    </a>
    <a class="lhd-action-card" href="view.asp">
      <span class="action-icon">&#128203;</span>
      <span class="action-label"><%=lang(cnnDB, "ViewProblemList")%></span>
      <span class="action-desc">Tickets pendientes</span>
    </a>
    <a class="lhd-action-card" href="search.asp">
      <span class="action-icon">&#128269;</span>
      <span class="action-label"><%=lang(cnnDB, "SearchProblems")%></span>
      <span class="action-desc">B&uacute;squeda avanzada</span>
    </a>
    <% If SafeCInt(Cfg(cnnDB, "EnableKB")) >= 1 Then %>
    <a class="lhd-action-card" href="../kb/default.asp">
      <span class="action-icon">&#128218;</span>
      <span class="action-label"><%=lang(cnnDB, "KnowledgeBase")%></span>
      <span class="action-desc">Base de conocimiento</span>
    </a>
    <% End If %>
    <% If SafeCInt(Cfg(cnnDB, "UseInoutBoard")) = 1 Then %>
    <a class="lhd-action-card" href="../inout/default.asp">
      <span class="action-icon">&#128100;</span>
      <span class="action-label"><%=lang(cnnDB, "InOutBoard")%></span>
      <span class="action-desc">Tablero In/Out</span>
    </a>
    <% End If %>
    <a class="lhd-action-card" href="../register.asp?edit=1">
      <span class="action-icon">&#9881;</span>
      <span class="action-label"><%=lang(cnnDB, "EditInformation")%></span>
      <span class="action-desc">Mi cuenta</span>
    </a>
  </div>

  <!-- Mis Tickets Pendientes -->
  <div class="lhd-card lhd-card-flush">
    <div class="lhd-card-header">
      <span>&#128203;</span> Mis Tickets Pendientes
    </div>

    <% If rstRecent.EOF Then %>
      <div style="padding:20px">
        <div class="alert alert-success" style="margin:0">
          <span class="alert-icon">&#10004;</span>
          <span>&iexcl;Excelente! No tienes tickets pendientes asignados en este momento.</span>
        </div>
      </div>
    <% Else %>
      <div style="overflow-x:auto">
        <table class="lhd-table">
          <thead>
            <tr>
              <th class="col-id">ID</th>
              <th><%=lang(cnnDB, "Title")%></th>
              <th><%=lang(cnnDB, "UserName")%></th>
              <th>Prioridad</th>
              <th>Estado</th>
              <th class="col-date">Fecha Registro</th>
            </tr>
          </thead>
          <tbody>
            <% Do While Not rstRecent.EOF %>
              <tr>
                <td class="col-id">
                  <a href="details.asp?id=<% = rstRecent("id") %>" style="font-weight:700">#<% = rstRecent("id") %></a>
                </td>
                <td>
                  <a href="details.asp?id=<% = rstRecent("id") %>" style="font-weight:500">
                    <% = FixTildes(Server.HTMLEncode(rstRecent("title"))) %>
                  </a>
                </td>
                <td><% = FixTildes(Server.HTMLEncode(rstRecent("uid"))) %></td>
                <td><% = PriorityBadge(rstRecent("pname")) %></td>
                <td><% = StatusBadge(rstRecent("sname")) %></td>
                <td class="col-date">
                  <% = DisplayDate(rstRecent("start_date"), lhdDateOnly) %>
                </td>
              </tr>
            <% 
                rstRecent.MoveNext
               Loop 
            %>
          </tbody>
        </table>
      </div>
      <div style="text-align:right;padding:14px 20px">
        <a class="btn btn-secondary btn-sm" href="view.asp">Ver listado completo &rarr;</a>
      </div>
    <% 
        rstRecent.Close 
      End If 
    %>
  </div>

</div>

<%
  Call DisplayFooter(cnnDB, sid)
  cnnDB.Close
%>

</body>
</html>
