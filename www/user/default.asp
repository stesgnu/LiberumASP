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
  Filename: default.asp (user)
  Purpose:  Main menu/dashboard for users.
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
  <title><% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "HelpDesk")%></title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="stylesheet" type="text/css" href="../default.css?v=4" />
</head>
<body>

<%
  Call CheckUser(cnnDB, sid)
  Call DisplayHeader(cnnDB, sid)

  Dim uid, fname
  uid   = Usr(cnnDB, sid, "uid")
  fname = Usr(cnnDB, sid, "fname")
  If IsNull(fname) Or Len(Trim(fname)) = 0 Then fname = uid

  ' Query: last 5 tickets for this user
  Dim rstRecent, recentStr
  recentStr = "SELECT TOP 5 p.id, p.title, p.start_date, s.sname, pri.pname " & _
    "FROM problems AS p " & _
    "LEFT JOIN status AS s ON p.status = s.status_id " & _
    "LEFT JOIN priority AS pri ON p.priority = pri.priority_id " & _
    "WHERE p.uid='" & Replace(uid, "'", "''") & "' " & _
    "ORDER BY p.id DESC"
  Set rstRecent = SQLQuery(cnnDB, recentStr)

  ' Count open tickets for this user
  Dim rstCountOpen, countOpen
  Set rstCountOpen = SQLQuery(cnnDB, "SELECT count(*) AS total FROM problems WHERE uid='" & Replace(uid, "'", "''") & "' AND status < 100")
  countOpen = rstCountOpen("total")
  rstCountOpen.Close

  ' Count total tickets for this user
  Dim rstCountTotal, countTotal
  Set rstCountTotal = SQLQuery(cnnDB, "SELECT count(*) AS total FROM problems WHERE uid='" & Replace(uid, "'", "''") & "'")
  countTotal = rstCountTotal("total")
  rstCountTotal.Close
%>

<div class="lhd-wrapper">

  <!-- Bienvenida -->
  <div style="margin-bottom:20px">
    <div class="lhd-page-title">
      <span class="title-icon">&#127968;</span>
      Bienvenido, <%=FixTildes(Server.HTMLEncode(fname)) %>
    </div>
  </div>

  <!-- Toast de confirmacion -->
  <% Call ShowToast(cnnDB) %>

  <!-- KPI Cards con Iconos y Diseno Moderno -->
  <div class="lhd-kpi-grid">
    <div class="lhd-kpi-card kpi-open" style="display:flex;align-items:center;justify-content:space-between">
      <div>
        <div class="kpi-value"><%= countOpen %></div>
        <div class="kpi-label">Tickets Abiertos</div>
      </div>
      <div style="width:44px;height:44px;border-radius:10px;background:var(--color-primary-light,#EFF6FF);color:var(--color-primary,#0057B8);display:flex;align-items:center;justify-content:center;font-size:22px">&#128194;</div>
    </div>
    <div class="lhd-kpi-card kpi-closed" style="display:flex;align-items:center;justify-content:space-between">
      <div>
        <div class="kpi-value" style="color:var(--color-success)"><%= countTotal - countOpen %></div>
        <div class="kpi-label">Resueltos</div>
      </div>
      <div style="width:44px;height:44px;border-radius:10px;background:#ECFDF5;color:#10B981;display:flex;align-items:center;justify-content:center;font-size:22px">&#127942;</div>
    </div>
    <div class="lhd-kpi-card kpi-pending" style="display:flex;align-items:center;justify-content:space-between">
      <div>
        <div class="kpi-value" style="color:#D69E2E"><%= countTotal %></div>
        <div class="kpi-label">Total Tickets</div>
      </div>
      <div style="width:44px;height:44px;border-radius:10px;background:#FFFBEB;color:#F59E0B;display:flex;align-items:center;justify-content:center;font-size:22px">&#128203;</div>
    </div>
  </div>

  <!-- Acciones rapidas -->
  <div class="lhd-action-grid">
    <a class="lhd-action-card" href="new.asp">
      <span class="action-icon">&#10133;</span>
      <span class="action-label"><%=lang(cnnDB, "SubmitNewProblem")%></span>
      <span class="action-desc">Reportar un nuevo problema</span>
    </a>
    <a class="lhd-action-card" href="view.asp">
      <span class="action-icon">&#128203;</span>
      <span class="action-label"><%=lang(cnnDB, "ViewProblemList")%></span>
      <span class="action-desc">Ver todos mis tickets</span>
    </a>
    <% If SafeCInt(Cfg(cnnDB, "EnableKB")) >= 2 Then %>
    <a class="lhd-action-card" href="../kb/">
      <span class="action-icon">&#128218;</span>
      <span class="action-label"><%=lang(cnnDB, "KnowledgeBase")%></span>
      <span class="action-desc">Buscar soluciones conocidas</span>
    </a>
    <% End If %>
    <% If SafeCInt(Cfg(cnnDB, "UseInoutBoard")) = 1 Then %>
    <a class="lhd-action-card" href="../inout/default.asp">
      <span class="action-icon">&#128100;</span>
      <span class="action-label"><%=lang(cnnDB, "InOutBoard")%></span>
      <span class="action-desc">Ver disponibilidad del equipo</span>
    </a>
    <% End If %>
    <a class="lhd-action-card" href="../register.asp?edit=1">
      <span class="action-icon">&#9881;</span>
      <span class="action-label"><%=lang(cnnDB, "EditInformation")%></span>
      <span class="action-desc">Actualizar mi perfil</span>
    </a>
  </div>

  <!-- Buscar ticket especifico -->
  <div class="lhd-card" style="padding:16px 20px">
    <form method="POST" action="details.asp" style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <span style="font-size:13px;font-weight:600;color:#4A5568">Ir a ticket #:</span>
      <input type="text" name="id" class="form-control" size="8" placeholder="ID..." style="max-width:120px" maxlength="10" />
      <button type="submit" class="btn btn-secondary btn-sm">&#128269; Abrir</button>
    </form>
  </div>

  <!-- Tickets recientes -->
  <div class="lhd-card">
    <div class="section-header">
      &#128203; Mis tickets recientes
    </div>

    <% If rstRecent.EOF Then %>
      <div class="alert alert-info" style="margin-top:0">
        <span class="alert-icon">&#8505;</span>
        <span>No tienes tickets registrados. <a href="new.asp">&#10133; Crear el primero</a></span>
      </div>
    <% Else %>
      <div style="overflow-x:auto">
        <table class="lhd-table">
          <thead>
            <tr>
              <th class="col-id">ID</th>
              <th>Titulo</th>
              <th>Estado</th>
              <th>Prioridad</th>
              <th class="col-date">Fecha</th>
            </tr>
          </thead>
          <tbody>
            <% Do While Not rstRecent.EOF %>
              <tr>
                <td class="col-id">
                  <a href="details.asp?id=<% = rstRecent("id") %>">#<% = rstRecent("id") %></a>
                </td>
                <td>
                  <a href="details.asp?id=<% = rstRecent("id") %>">
                    <%= FixTildes(Server.HTMLEncode(rstRecent("title"))) %>
                  </a>
                </td>
                <td>
                  <%= StatusBadge(rstRecent("sname")) %>
                </td>
                <td>
                  <%= PriorityBadge(rstRecent("pname")) %>
                </td>
                <td class="col-date">
                  <%= DisplayDate(rstRecent("start_date"), lhdDateOnly) %>
                </td>
              </tr>
            <% 
                rstRecent.MoveNext 
               Loop 
            %>
          </tbody>
        </table>
      </div>
      <div style="text-align:right;margin-top:12px">
        <a class="btn btn-secondary btn-sm" href="view.asp">Ver todos los tickets &rarr;</a>
      </div>
    <% End If %>
    <% rstRecent.Close %>
  </div>

</div>

<% 
  Call DisplayFooter(cnnDB, sid)
  cnnDB.Close 
%>

</body>
</html>
