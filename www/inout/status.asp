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

  Filename: status.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Page to update the status information for the In/Out Board.
  -->

  <!-- 	#include file = "../public.asp" -->
  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    ' Look for the problem id in the query string, if it's not
    ' there display an error.

    If Len(Request.QueryString("id")) = 0 Then
      Dim noIdMsg
      noIdMsg = lang(cnnDB, "NovalidIDgiven")
      cnnDB.Close
      Call DisplayError(3, noIdMsg)
    End If
  %>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta charset="utf-8">
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "InOutBoard")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style>
      .io-wrap { max-width: 640px; margin: 0 auto; }
      .io-switch { display: flex; align-items: center; gap: 12px; padding: 12px 14px; border: 1px solid var(--color-border-light,#E2E8F0); border-radius: 10px; background: #F8FAFC; cursor: pointer; margin-bottom: 16px; }
      .io-switch input { position: absolute; opacity: 0; pointer-events: none; }
      .io-track { position: relative; width: 42px; height: 22px; border-radius: 9999px; background: #CBD5E0; flex-shrink: 0; transition: background .2s ease; }
      .io-knob { position: absolute; top: 2px; left: 2px; width: 18px; height: 18px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,.3); transition: transform .2s ease; }
      .io-switch input:checked + .io-track { background: var(--color-primary,#0057B8); }
      .io-switch input:checked + .io-track .io-knob { transform: translateX(20px); }
      .io-switch input:focus-visible + .io-track { outline: 2px solid var(--color-primary,#0057B8); outline-offset: 2px; }
      .io-switch .t { font-weight: 600; color: #1a202c; font-size: 14px; }
      .io-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
    </style>
  </head>
  <body>
  <%
    ' See if authenticated
    Call CheckUser(cnnDB, sid)

    'Check if user is In/Out Board admin
    Dim inoutadmin
    inoutadmin = SafeCInt(Usr(cnnDB, sid, "inoutadmin"))

    Dim strSql, rstUser, rstUpdate
    Dim usid
    usid = SafeCLng(Request.QueryString("id"))

    ' Un estado "Licencia" (2) solo lo puede cambiar el administrador del tablero In/Out.
    ' Se comprueba ANTES de guardar (antes el UPDATE se hacia primero y saltaba la restriccion).
    Dim rstPre
    Set rstPre = SQLQuery(cnnDB, "SELECT statuscode FROM tblUsers WHERE sid=" & usid)
    If Not rstPre.EOF Then
      If SafeCInt(rstPre("statuscode")) = 2 And inoutadmin <> 1 Then
        rstPre.Close
        Call DisplayError(3, lang(cnnDB, "OnlyAdministratorscanchangethisstatus"))
      End If
    End If
    rstPre.Close

    If Request.Form("save") = "1" Then
      Dim sStatus, Statustext, StatusDate
      Dim frm_status, frm_statustext
      If Request.Form("frm_status") = "on" Then
        sStatus = 1
        Statustext = Request.Form("frm_statustext")
        Statustext = Replace(Statustext, "'", "''")
      Else
        sStatus = 0
        Statustext = ""
      End If
      StatusDate = SQLDate(Now, lhdAddSQLDelim)
      strSql = "UPDATE tblUsers SET " &_
        "statuscode = " & sStatus & ", " &_
        "statustext = '" & Statustext & "', " &_
        "statusdate = " & StatusDate & " " &_
        "WHERE sid = " & usid
      Set rstUpdate = SQLQuery(cnnDB, strSql)
    End If

    strSql = "SELECT * FROM tblUsers WHERE sid=" & usid
    Set rstUser = SQLQuery(cnnDB, strSql)

    If rstUser.EOF Then
      Call DisplayError(3, lang(cnnDB, "Noresultsfound"))
    End If

    Dim uname, check
    uname = rstUser("fname") & ""
    If rstUser("statuscode") > 0 Then
      check = "checked"
    Else
      check = ""
    End If

    If rstUser("statuscode") = 2 And inoutadmin <> 1 Then
      Call DisplayError(3, lang(cnnDB, "OnlyAdministratorscanchangethisstatus"))
    End If

    If Request.Form("save") = "1" Then
      Call LhdToast("success", lang(cnnDB, "Status") & " " & lang(cnnDB, "isupdated"))
    End If

    Call DisplayHeader(cnnDB, sid)
  %>
  <div class="lhd-wrapper">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
      <div class="lhd-page-title" style="margin:0"><span class="title-icon">&#8644;</span> <%=lang(cnnDB, "StatusInformation")%></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <a class="btn btn-secondary" href="details.asp?id=<% = usid %>">&larr; <%=lang(cnnDB, "Details")%></a>
      </div>
    </div>

    <div class="io-wrap">
      <div class="lhd-card">
        <div class="section-header">&#128100; <% = Server.HTMLEncode(uname) %></div>
        <form method="POST" action="status.asp?id=<% = usid %>">
          <input type="hidden" name="save" value="1">

          <div class="form-group">
            <label class="form-label"><%=lang(cnnDB, "Fullname")%></label>
            <input type="text" class="form-control" value="<% = Server.HTMLEncode(uname) %>" readonly>
          </div>

          <label class="io-switch">
            <input type="checkbox" name="frm_status" <% = check %>>
            <span class="io-track"><span class="io-knob"></span></span>
            <span class="t"><%=lang(cnnDB, "Out")%></span>
          </label>

          <div class="form-group">
            <label class="form-label" for="frm_statustext"><%=lang(cnnDB, "Text")%></label>
            <textarea class="form-control" id="frm_statustext" name="frm_statustext" rows="4"><% = Server.HTMLEncode(rstUser("statustext") & "") %></textarea>
          </div>

          <div class="io-actions">
            <button type="submit" class="btn btn-primary"><%=lang(cnnDB, "Submit")%></button>
          </div>
        </form>
      </div>
    </div>
  </div>
  <%
    rstUser.Close
    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close
  %>
</body>
</html>
