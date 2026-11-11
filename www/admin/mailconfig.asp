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

  Filename: mailconfig.asp
  Purpose:  Configuracion del correo entrante ("correo a caso"): buzon IMAP, SMTP, reglas y registro.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;Correo entrante</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .mc-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 0 18px; }
      .mc-card { margin-bottom: 18px; }
      .mc-hint { font-size: 12px; color: #64748b; margin-top: 4px; }
      .mc-switch { display: flex; align-items: center; gap: 12px; padding: 12px 14px; border: 1px solid var(--color-border-light, #E2E8F0); border-radius: 10px; background: #F8FAFC; cursor: pointer; }
      .mc-switch input { position: absolute; opacity: 0; pointer-events: none; }
      .mc-track { position: relative; width: 42px; height: 22px; border-radius: 9999px; background: #CBD5E0; flex-shrink: 0; transition: background .2s ease; }
      .mc-knob { position: absolute; top: 2px; left: 2px; width: 18px; height: 18px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,.3); transition: transform .2s ease; }
      .mc-switch input:checked + .mc-track { background: var(--color-success, #1A7F45); }
      .mc-switch input:checked + .mc-track .mc-knob { transform: translateX(20px); }
      .mc-switch .t { font-weight: 600; color: #1a202c; font-size: 14px; }
      .mc-switch .d { font-size: 12px; color: #64748b; }
      .mc-status { display: flex; flex-wrap: wrap; gap: 14px; margin-bottom: 18px; }
      .mc-stat { flex: 1 1 200px; background: #fff; border: 1px solid var(--color-border-light, #E2E8F0); border-left: 4px solid #94A3B8; border-radius: 10px; padding: 12px 16px; }
      .mc-stat.ok { border-left-color: var(--color-success, #1A7F45); }
      .mc-stat.bad { border-left-color: var(--color-danger, #B02A37); }
      .mc-stat .l { font-size: 12px; color: #64748b; text-transform: uppercase; letter-spacing: .04em; }
      .mc-stat .v { font-size: 15px; font-weight: 700; color: #1a202c; margin-top: 3px; word-break: break-word; }
      .mc-badge { display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; }
      .mc-badge.created { background: var(--color-success-bg, #D4EDDA); color: var(--color-success, #1A7F45); }
      .mc-badge.note { background: var(--color-primary-light, #E8F0FB); color: var(--color-primary-dark, #003F8A); }
      .mc-badge.ignored, .mc-badge.duplicate { background: #F1F5F9; color: #64748b; }
      .mc-badge.error, .mc-badge.failed { background: var(--color-danger-bg, #F8D7DA); color: var(--color-danger, #B02A37); }
      .mc-code { display: block; background: #0F172A; color: #E2E8F0; border-radius: 8px; padding: 12px 14px; font-family: Consolas, "Courier New", monospace; font-size: 12px; word-break: break-all; white-space: pre-wrap; }
      .mc-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; margin-bottom: 18px; }
      .mc-token { display: flex; gap: 8px; flex-wrap: wrap; }
      .mc-token .form-control { flex: 1 1 280px; font-family: Consolas, "Courier New", monospace; font-size: 13px; }
      .mc-empty { padding: 28px 16px; text-align: center; color: #64748b; }
    </style>
  </head>
  <body>

    <%
      Call CheckAdmin

      Function MailSelectOptions(sqlText, idCol, nameCol, selId, firstLabel)
        Dim rsO, outO
        outO = "<option value=""0"">" & firstLabel & "</option>"
        Set rsO = SQLQuery(cnnDB, sqlText)
        Do While Not rsO.EOF
          If SafeCLng(rsO(idCol)) = SafeCLng(selId) Then
            outO = outO & "<option value=""" & rsO(idCol) & """ selected>" & Server.HTMLEncode(rsO(nameCol) & "") & "</option>"
          Else
            outO = outO & "<option value=""" & rsO(idCol) & """>" & Server.HTMLEncode(rsO(nameCol) & "") & "</option>"
          End If
          rsO.MoveNext
        Loop
        rsO.Close
        MailSelectOptions = outO
      End Function

      Function SecOptions(cur)
        Dim o
        o = "<option value=""ssl""" & IIf2(cur = "ssl", " selected", "") & ">SSL/TLS directo (puertos 993 / 465)</option>"
        o = o & "<option value=""starttls""" & IIf2(cur = "starttls", " selected", "") & ">STARTTLS (puertos 143 / 587)</option>"
        o = o & "<option value=""none""" & IIf2(cur = "none", " selected", "") & ">Sin cifrado (no recomendado)</option>"
        SecOptions = o
      End Function

      ' Estado de las tablas
      Dim mcReady, rsCfg, mcMsg, mcErr
      Dim cEnabled, cImapHost, cImapPort, cImapSec, cImapUser, cImapPassSet, cImapFolder, cImapDone
      Dim cSmtpHost, cSmtpPort, cSmtpSec, cSmtpUser, cSmtpPassSet, cSmtpFromAddr, cSmtpFromName
      Dim cDefCat, cDefDept, cDefPrio, cUnknown, cReopen, cAckSubj, cAckBody, cToken, cLastRun, cLastStatus
      mcReady = False
      mcErr = ""

      ' Guardar
      If Request.Form("save") = "1" Then
        Dim sqlSave, newTok, p1, p2, newImapPass, newSmtpPass
        newTok = Trim(Request.Form("api_token"))
        newImapPass = Request.Form("imap_pass")
        newSmtpPass = Request.Form("smtp_pass")
        If Len(newTok) > 0 And (Len(newTok) < 24 Or Len(newTok) > 80) Then
          mcErr = "El token debe tener entre 24 y 80 caracteres."
        End If
        If Len(mcErr) = 0 Then
          sqlSave = "UPDATE tblMailConfig SET " & _
            "enabled=" & IIf2(Request.Form("enabled") = "on", 1, 0) & ", " & _
            "imap_host='" & Replace(Left(Trim(Request.Form("imap_host")), 150), "'", "''") & "', " & _
            "imap_port=" & SafeCLng(Request.Form("imap_port")) & ", " & _
            "imap_security='" & Replace(Left(Request.Form("imap_security"), 10), "'", "''") & "', " & _
            "imap_user='" & Replace(Left(Trim(Request.Form("imap_user")), 150), "'", "''") & "', " & _
            "imap_folder='" & Replace(Left(Trim(Request.Form("imap_folder")), 100), "'", "''") & "', " & _
            "imap_done_folder='" & Replace(Left(Trim(Request.Form("imap_done_folder")), 100), "'", "''") & "', " & _
            "smtp_host='" & Replace(Left(Trim(Request.Form("smtp_host")), 150), "'", "''") & "', " & _
            "smtp_port=" & SafeCLng(Request.Form("smtp_port")) & ", " & _
            "smtp_security='" & Replace(Left(Request.Form("smtp_security"), 10), "'", "''") & "', " & _
            "smtp_user='" & Replace(Left(Trim(Request.Form("smtp_user")), 150), "'", "''") & "', " & _
            "smtp_from_addr='" & Replace(Left(Trim(Request.Form("smtp_from_addr")), 150), "'", "''") & "', " & _
            "smtp_from_name='" & Replace(Left(Trim(Request.Form("smtp_from_name")), 100), "'", "''") & "', " & _
            "def_category=" & SafeCLng(Request.Form("def_category")) & ", " & _
            "def_department=" & SafeCLng(Request.Form("def_department")) & ", " & _
            "def_priority=" & SafeCLng(Request.Form("def_priority")) & ", " & _
            "unknown_policy=" & SafeCLng(Request.Form("unknown_policy")) & ", " & _
            "reopen_closed=" & IIf2(Request.Form("reopen_closed") = "on", 1, 0) & ", " & _
            "ack_subject='" & Replace(Left(Trim(Request.Form("ack_subject")), 200), "'", "''") & "', " & _
            "ack_body='" & Replace(Request.Form("ack_body") & "", "'", "''") & "'"
          If Len(Request.Form("imap_pass")) > 0 Then sqlSave = sqlSave & ", imap_pass='" & Replace(Left(Request.Form("imap_pass"), 200), "'", "''") & "'"
          If Len(Request.Form("smtp_pass")) > 0 Then sqlSave = sqlSave & ", smtp_pass='" & Replace(Left(Request.Form("smtp_pass"), 200), "'", "''") & "'"
          If Len(newTok) > 0 Then sqlSave = sqlSave & ", api_token='" & Replace(newTok, "'", "''") & "'"
          Call SQLQuery(cnnDB, sqlSave)
          Call LhdToast("success", "Configuraci" & ChrW(243) & "n del correo entrante guardada.")
        Else
          Call LhdToast("error", mcErr)
        End If
      End If

      ' Correo de prueba (se envia por el script, desde la cola de salida)
      If Request.Form("sendtest") = "1" Then
        Dim testTo
        testTo = Left(Trim(Request.Form("test_to")), 200)
        If InStr(testTo, "@") > 1 And InStr(testTo, " ") = 0 Then
          Call SQLQuery(cnnDB, "INSERT INTO tblMailOutbox (to_addr, subject, body) VALUES ('" & Replace(testTo, "'", "''") & "', " & _
            "'Prueba de la mesa de servicio', 'Este es un correo de prueba enviado desde la aplicaci" & ChrW(243) & "n de la mesa de servicio. Si lo recibiste, el env" & ChrW(237) & "o por SMTP funciona.')")
          Call LhdToast("info", "Correo de prueba en cola para " & testTo & ". Se enviar" & ChrW(225) & " en la pr" & ChrW(243) & "xima ejecuci" & ChrW(243) & "n del script.")
        Else
          Call LhdToast("error", "Escribe una direcci" & ChrW(243) & "n de correo v" & ChrW(225) & "lida para la prueba.")
        End If
      End If

      ' Direcciones de entrada por categoria
      Dim aliasReady, rsAT
      aliasReady = False
      On Error Resume Next
      Set rsAT = cnnDB.Execute("SELECT TOP 1 id FROM tblMailAlias")
      If Err.Number = 0 Then
        aliasReady = True
        rsAT.Close
      End If
      Err.Clear
      On Error Goto 0

      If aliasReady And Request.Form("addalias") = "1" Then
        Dim aAddr, aCat, aDep, aPri, rsDup
        aAddr = LCase(Left(Trim(Request.Form("alias_address")), 150))
        aCat = SafeCLng(Request.Form("alias_category"))
        aDep = SafeCLng(Request.Form("alias_department"))
        aPri = SafeCLng(Request.Form("alias_priority"))
        If InStr(aAddr, "@") < 2 Or InStr(aAddr, " ") > 0 Or InStr(aAddr, "'") > 0 Or InStr(aAddr, """") > 0 Or InStr(aAddr, "<") > 0 Or InStr(Mid(aAddr, InStr(aAddr, "@")), ".") = 0 Then
          Call LhdToast("error", "Escribe una direcci" & ChrW(243) & "n de correo v" & ChrW(225) & "lida.")
        ElseIf aCat <= 0 Then
          Call LhdToast("error", "Elige la categor" & ChrW(237) & "a de la direcci" & ChrW(243) & "n.")
        Else
          Set rsDup = SQLQuery(cnnDB, "SELECT id FROM tblMailAlias WHERE LOWER(address)='" & Replace(aAddr, "'", "''") & "'")
          If Not rsDup.EOF Then
            Call LhdToast("warning", "Esa direcci" & ChrW(243) & "n ya est" & ChrW(225) & " registrada.")
          Else
            Call SQLQuery(cnnDB, "INSERT INTO tblMailAlias (address, category_id, def_department, def_priority, enabled) VALUES ('" & _
              Replace(aAddr, "'", "''") & "', " & aCat & ", " & IIf2(aDep > 0, aDep, "NULL") & ", " & IIf2(aPri > 0, aPri, "NULL") & ", 1)")
            Call LhdToast("success", "Direcci" & ChrW(243) & "n " & aAddr & " agregada.")
          End If
          rsDup.Close
        End If
      End If
      If aliasReady And SafeCLng(Request.Form("delalias")) > 0 Then
        Call SQLQuery(cnnDB, "DELETE FROM tblMailAlias WHERE id=" & SafeCLng(Request.Form("delalias")))
        Call LhdToast("success", "Direcci" & ChrW(243) & "n eliminada.")
      End If
      If aliasReady And SafeCLng(Request.Form("togglealias")) > 0 Then
        Call SQLQuery(cnnDB, "UPDATE tblMailAlias SET enabled = CASE WHEN enabled=1 THEN 0 ELSE 1 END WHERE id=" & SafeCLng(Request.Form("togglealias")))
        Call LhdToast("info", "Direcci" & ChrW(243) & "n actualizada.")
      End If

      ' Leer configuracion
      On Error Resume Next
      Set rsCfg = cnnDB.Execute("SELECT * FROM tblMailConfig")
      If Err.Number = 0 Then
        If Not rsCfg.EOF Then
          mcReady = True
          cEnabled = (SafeCInt(rsCfg("enabled")) <> 0)
          cImapHost = rsCfg("imap_host") & "" : cImapPort = SafeCLng(rsCfg("imap_port")) : cImapSec = rsCfg("imap_security") & ""
          cImapUser = rsCfg("imap_user") & "" : cImapPassSet = (Len(rsCfg("imap_pass") & "") > 0)
          cImapFolder = rsCfg("imap_folder") & "" : cImapDone = rsCfg("imap_done_folder") & ""
          cSmtpHost = rsCfg("smtp_host") & "" : cSmtpPort = SafeCLng(rsCfg("smtp_port")) : cSmtpSec = rsCfg("smtp_security") & ""
          cSmtpUser = rsCfg("smtp_user") & "" : cSmtpPassSet = (Len(rsCfg("smtp_pass") & "") > 0)
          cSmtpFromAddr = rsCfg("smtp_from_addr") & "" : cSmtpFromName = rsCfg("smtp_from_name") & ""
          cDefCat = SafeCLng(rsCfg("def_category")) : cDefDept = SafeCLng(rsCfg("def_department")) : cDefPrio = SafeCLng(rsCfg("def_priority"))
          cUnknown = SafeCLng(rsCfg("unknown_policy")) : cReopen = (SafeCInt(rsCfg("reopen_closed")) <> 0)
          cAckSubj = rsCfg("ack_subject") & "" : cAckBody = rsCfg("ack_body") & ""
          cToken = rsCfg("api_token") & ""
          cLastRun = rsCfg("last_run") : cLastStatus = rsCfg("last_status") & ""
        End If
        rsCfg.Close
      End If
      Err.Clear
      On Error Goto 0

      Call DisplayHeader(cnnDB, sid)

      Dim appUrl, scheme, pendingOut, errOut, errLog, rsX
      scheme = "http"
      If LCase(Request.ServerVariables("HTTPS") & "") = "on" Then scheme = "https"
      appUrl = scheme & "://" & Request.ServerVariables("SERVER_NAME") & GetAppPath()
      pendingOut = 0 : errOut = 0 : errLog = 0
      If mcReady Then
        Set rsX = SQLQuery(cnnDB, "SELECT COUNT(*) AS n FROM tblMailOutbox WHERE status=0")
        pendingOut = SafeCLng(rsX("n")) : rsX.Close
        Set rsX = SQLQuery(cnnDB, "SELECT COUNT(*) AS n FROM tblMailOutbox WHERE status=2")
        errOut = SafeCLng(rsX("n")) : rsX.Close
        Set rsX = SQLQuery(cnnDB, "SELECT COUNT(*) AS n FROM tblMailLog WHERE result IN ('error','failed') AND received >= " & SQLDate(DateAdd("d", -1, Now), lhdAddSQLDelim))
        errLog = SafeCLng(rsX("n")) : rsX.Close
      End If
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128231;</span>
          Correo entrante de la mesa de servicio
        </div>
        <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
      </div>

      <% If Not mcReady Then %>
        <div class="alert alert-error">
          <span class="alert-icon">&#9888;</span>
          <span>Las tablas del correo entrante no est&aacute;n instaladas. Ejecute el script <strong>db/mail_inbound_sqlserver.sql</strong> sobre la base de datos del Help Desk y recargue esta p&aacute;gina.</span>
        </div>
      <% Else %>

      <div class="mc-status">
        <div class="mc-stat <% If cEnabled Then Response.Write("ok") %>"><div class="l">Estado</div><div class="v"><% If cEnabled Then Response.Write("Activo") Else Response.Write("Desactivado") End If %></div></div>
        <div class="mc-stat"><div class="l">&Uacute;ltima ejecuci&oacute;n del script</div><div class="v"><% If IsDate(cLastRun) Then Response.Write(DisplayDate(cLastRun, lhdDateTime)) Else Response.Write("Nunca") End If %></div>
          <% If Len(cLastStatus) > 0 Then %><div class="mc-hint"><% = Server.HTMLEncode(Left(cLastStatus, 160)) %></div><% End If %></div>
        <div class="mc-stat <% If pendingOut = 0 And errOut = 0 Then %>ok<% ElseIf errOut > 0 Then %>bad<% End If %>"><div class="l">Respuestas por enviar</div><div class="v"><% = pendingOut %><% If errOut > 0 Then Response.Write(" &middot; " & errOut & " con error") %></div></div>
        <div class="mc-stat <% If errLog > 0 Then Response.Write("bad") Else Response.Write("ok") End If %>"><div class="l">Errores (24 h)</div><div class="v"><% = errLog %></div></div>
      </div>

      <form method="post" action="mailconfig.asp" id="frmMail" autocomplete="off">
        <input type="hidden" name="save" value="1">

        <div class="lhd-card mc-card">
          <div class="section-header">&#9881; Activaci&oacute;n</div>
          <label class="mc-switch">
            <input type="checkbox" name="enabled" <% If cEnabled Then Response.Write("checked") %>>
            <span class="mc-track"><span class="mc-knob"></span></span>
            <span><span class="t">Crear casos a partir de los correos recibidos</span><br><span class="d">Cuando est&aacute; desactivado el script no lee el buz&oacute;n ni env&iacute;a respuestas.</span></span>
          </label>
        </div>

        <div class="lhd-card mc-card">
          <div class="section-header">&#128229; Buz&oacute;n de entrada (IMAP)</div>
          <div class="mc-grid">
            <div class="form-group"><label class="form-label" for="imap_host">Servidor</label>
              <input type="text" class="form-control" id="imap_host" name="imap_host" maxlength="150" value="<% = Server.HTMLEncode(cImapHost) %>" placeholder="imap.empresa.com"></div>
            <div class="form-group"><label class="form-label" for="imap_port">Puerto</label>
              <input type="number" class="form-control" id="imap_port" name="imap_port" min="1" max="65535" value="<% If cImapPort > 0 Then Response.Write(cImapPort) %>" placeholder="993"></div>
            <div class="form-group"><label class="form-label" for="imap_security">Seguridad</label>
              <select class="form-control" id="imap_security" name="imap_security"><% = SecOptions(cImapSec) %></select></div>
            <div class="form-group"><label class="form-label" for="imap_user">Usuario (correo del buz&oacute;n)</label>
              <input type="text" class="form-control" id="imap_user" name="imap_user" maxlength="150" value="<% = Server.HTMLEncode(cImapUser) %>" placeholder="mesa@empresa.com" autocomplete="off"></div>
            <div class="form-group"><label class="form-label" for="imap_pass">Contrase&ntilde;a</label>
              <input type="password" class="form-control" id="imap_pass" name="imap_pass" maxlength="200" autocomplete="new-password" placeholder="<% If cImapPassSet Then Response.Write("&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull; (sin cambios)") %>"></div>
            <div class="form-group"><label class="form-label" for="imap_folder">Carpeta a leer</label>
              <input type="text" class="form-control" id="imap_folder" name="imap_folder" maxlength="100" value="<% = Server.HTMLEncode(cImapFolder) %>" placeholder="INBOX"></div>
            <div class="form-group"><label class="form-label" for="imap_done_folder">Mover procesados a (opcional)</label>
              <input type="text" class="form-control" id="imap_done_folder" name="imap_done_folder" maxlength="100" value="<% = Server.HTMLEncode(cImapDone) %>" placeholder="Procesados">
              <div class="mc-hint">Si queda vac&iacute;o, los correos solo se marcan como le&iacute;dos.</div></div>
          </div>
          <div class="mc-hint">Con Gmail o Microsoft 365 use una contrase&ntilde;a de aplicaci&oacute;n (no se admite OAuth).</div>
        </div>

        <div class="lhd-card mc-card">
          <div class="section-header">&#128228; Env&iacute;o de respuestas (SMTP)</div>
          <div class="mc-grid">
            <div class="form-group"><label class="form-label" for="smtp_host">Servidor</label>
              <input type="text" class="form-control" id="smtp_host" name="smtp_host" maxlength="150" value="<% = Server.HTMLEncode(cSmtpHost) %>" placeholder="smtp.empresa.com"></div>
            <div class="form-group"><label class="form-label" for="smtp_port">Puerto</label>
              <input type="number" class="form-control" id="smtp_port" name="smtp_port" min="1" max="65535" value="<% If cSmtpPort > 0 Then Response.Write(cSmtpPort) %>" placeholder="587"></div>
            <div class="form-group"><label class="form-label" for="smtp_security">Seguridad</label>
              <select class="form-control" id="smtp_security" name="smtp_security"><% = SecOptions(cSmtpSec) %></select></div>
            <div class="form-group"><label class="form-label" for="smtp_user">Usuario (vac&iacute;o si no requiere)</label>
              <input type="text" class="form-control" id="smtp_user" name="smtp_user" maxlength="150" value="<% = Server.HTMLEncode(cSmtpUser) %>" autocomplete="off"></div>
            <div class="form-group"><label class="form-label" for="smtp_pass">Contrase&ntilde;a</label>
              <input type="password" class="form-control" id="smtp_pass" name="smtp_pass" maxlength="200" autocomplete="new-password" placeholder="<% If cSmtpPassSet Then Response.Write("&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull; (sin cambios)") %>"></div>
            <div class="form-group"><label class="form-label" for="smtp_from_addr">Direcci&oacute;n remitente</label>
              <input type="text" class="form-control" id="smtp_from_addr" name="smtp_from_addr" maxlength="150" value="<% = Server.HTMLEncode(cSmtpFromAddr) %>" placeholder="mesa@empresa.com">
              <div class="mc-hint">Debe ser el correo del buz&oacute;n: as&iacute; las respuestas del usuario regresan a &eacute;l.</div></div>
            <div class="form-group"><label class="form-label" for="smtp_from_name">Nombre del remitente</label>
              <input type="text" class="form-control" id="smtp_from_name" name="smtp_from_name" maxlength="100" value="<% = Server.HTMLEncode(cSmtpFromName) %>" placeholder="Mesa de Servicio"></div>
          </div>
        </div>

        <div class="lhd-card mc-card">
          <div class="section-header">&#128221; Reglas de creaci&oacute;n de casos</div>
          <div class="mc-grid">
            <div class="form-group"><label class="form-label" for="def_category"><%=lang(cnnDB, "Category")%> por omisi&oacute;n <span class="form-required">*</span></label>
              <select class="form-control" id="def_category" name="def_category"><% = MailSelectOptions("SELECT category_id, cname FROM categories WHERE category_id > 0 ORDER BY cname ASC", "category_id", "cname", cDefCat, "&mdash; Elegir &mdash;") %></select>
              <div class="mc-hint">Se usa cuando el correo no lleg&oacute; por ninguna direcci&oacute;n registrada abajo. El caso se asigna al t&eacute;cnico de la categor&iacute;a.</div></div>
            <div class="form-group"><label class="form-label" for="def_department"><%=lang(cnnDB, "Department")%> por omisi&oacute;n <span class="form-required">*</span></label>
              <select class="form-control" id="def_department" name="def_department"><% = MailSelectOptions("SELECT department_id, dname FROM departments WHERE department_id > 0 ORDER BY dname ASC", "department_id", "dname", cDefDept, "&mdash; Elegir &mdash;") %></select>
              <div class="mc-hint">Se usa cuando el remitente no tiene departamento.</div></div>
            <div class="form-group"><label class="form-label" for="def_priority"><%=lang(cnnDB, "Priority")%> por omisi&oacute;n <span class="form-required">*</span></label>
              <select class="form-control" id="def_priority" name="def_priority"><% = MailSelectOptions("SELECT priority_id, pname FROM priority ORDER BY priority_id ASC", "priority_id", "pname", cDefPrio, "&mdash; Elegir &mdash;") %></select></div>
            <div class="form-group"><label class="form-label" for="unknown_policy">Remitentes no registrados</label>
              <select class="form-control" id="unknown_policy" name="unknown_policy">
                <option value="1"<% If cUnknown = 1 Then Response.Write(" selected") %>>Crear el caso igualmente</option>
                <option value="0"<% If cUnknown = 0 Then Response.Write(" selected") %>>Ignorar el correo</option>
              </select>
              <div class="mc-hint">Se reconoce al usuario por su correo registrado en la aplicaci&oacute;n.</div></div>
          </div>
          <label class="mc-switch" style="margin-top:6px">
            <input type="checkbox" name="reopen_closed" <% If cReopen Then Response.Write("checked") %>>
            <span class="mc-track"><span class="mc-knob"></span></span>
            <span><span class="t">Reabrir el caso si el usuario responde cuando ya est&aacute; cerrado</span><br><span class="d">La respuesta se agrega como nota; si est&aacute; desactivado, la nota se agrega sin cambiar el estado.</span></span>
          </label>
        </div>

        <div class="lhd-card mc-card">
          <div class="section-header">&#9993; Mensaje de confirmaci&oacute;n al usuario</div>
          <div class="form-group"><label class="form-label" for="ack_subject">Asunto</label>
            <input type="text" class="form-control" id="ack_subject" name="ack_subject" maxlength="200" value="<% = Server.HTMLEncode(cAckSubj) %>" placeholder="Re: [title]">
            <div class="mc-hint">Siempre se agrega al final la etiqueta <strong>[Caso #n]</strong>: con ella el sistema reconoce las respuestas del usuario.</div></div>
          <div class="form-group" style="margin-bottom:0"><label class="form-label" for="ack_body">Cuerpo</label>
            <textarea class="form-control" id="ack_body" name="ack_body" rows="9" placeholder="Si lo deja vac&iacute;o se usa el mensaje por omisi&oacute;n."><% = Server.HTMLEncode(cAckBody) %></textarea>
            <div class="mc-hint">Variables: <code>[nombre]</code> <code>[problemid]</code> <code>[title]</code> <code>[category]</code> <code>[status]</code> <code>[adjuntos]</code> <code>[sitename]</code> <code>[baseurl]</code></div></div>
        </div>

        <div class="lhd-card mc-card">
          <div class="section-header">&#128274; Token de la API y tarea programada</div>
          <div class="form-group">
            <label class="form-label" for="api_token">Token</label>
            <div class="mc-token">
              <input type="text" class="form-control" id="api_token" name="api_token" maxlength="80" value="<% = Server.HTMLEncode(cToken) %>" autocomplete="off" placeholder="Genere un token para activar la API">
              <button type="button" class="btn btn-secondary" id="btnTok">&#8635; Generar</button>
              <button type="button" class="btn btn-secondary" id="btnCopy">Copiar comando</button>
            </div>
            <div class="mc-hint">El script <code>tools/mailpoll.ps1</code> usa este token para entregar los correos. Gu&aacute;rdelo con la configuraci&oacute;n para que quede registrado.</div>
          </div>
          <label class="form-label">Comando para el Programador de tareas (cada 1 o 2 minutos)</label>
          <code class="mc-code" id="cmdLine">powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Liberum\tools\mailpoll.ps1" -Url "<% = Server.HTMLEncode(appUrl) %>" -Token "<span id="cmdTok"><% If Len(cToken) > 0 Then Response.Write(Server.HTMLEncode(cToken)) Else Response.Write("TOKEN") End If %></span>" -LogFile "C:\Liberum\tools\mailpoll.log"</code>
          <div class="mc-hint">Agregue <code>-SkipCertCheck</code> si el sitio usa un certificado autofirmado y <code>-Test</code> para probar la conexi&oacute;n IMAP/SMTP sin procesar correos.</div>
        </div>

        <div class="mc-actions">
          <a class="btn btn-secondary" href="default.asp">Cancelar</a>
          <button type="submit" class="btn btn-primary">&#128190; Guardar configuraci&oacute;n</button>
        </div>
      </form>

      <div class="lhd-card lhd-card-flush mc-card">
        <div class="section-header" style="margin:0;padding:14px 16px">&#128236; Direcciones de entrada por categor&iacute;a</div>
        <div style="padding:0 16px 12px">
          <p class="mc-hint" style="margin:0 0 6px">Todas las direcciones deben entregar al <strong>mismo buz&oacute;n</strong> configurado arriba (alias o reenv&iacute;o en su proveedor de correo).
            Al llegar un correo, el sistema mira a qu&eacute; direcci&oacute;n fue enviado (cabeceras <em>To</em>, <em>Cc</em>, <em>Delivered-To</em>, <em>X-Original-To</em>) y le asigna esa categor&iacute;a.
            Si no coincide ninguna, se usa la categor&iacute;a por omisi&oacute;n. No hay l&iacute;mite de direcciones.</p>
        </div>
        <% If Not aliasReady Then %>
          <div class="alert alert-error" style="margin:0 16px 16px"><span class="alert-icon">&#9888;</span><span>Falta la tabla <strong>tblMailAlias</strong>: vuelva a ejecutar <strong>db/mail_inbound_sqlserver.sql</strong>.</span></div>
        <% Else %>
        <div style="overflow-x:auto">
          <table class="lhd-table">
            <thead><tr><th>Direcci&oacute;n</th><th><%=lang(cnnDB, "Category")%></th><th><%=lang(cnnDB, "Department")%></th><th><%=lang(cnnDB, "Priority")%></th><th>Estado</th><th class="col-actions"></th></tr></thead>
            <tbody>
              <%
                Dim rsAl, nAl
                nAl = 0
                Set rsAl = SQLQuery(cnnDB, "SELECT a.id, a.address, a.enabled, c.cname, d.dname, p.pname " & _
                  "FROM ((tblMailAlias AS a LEFT JOIN categories AS c ON a.category_id = c.category_id) " & _
                  "LEFT JOIN departments AS d ON a.def_department = d.department_id) " & _
                  "LEFT JOIN priority AS p ON a.def_priority = p.priority_id ORDER BY a.address ASC")
                Do While Not rsAl.EOF
                  nAl = nAl + 1
              %>
              <tr>
                <td><strong><% = Server.HTMLEncode(rsAl("address") & "") %></strong></td>
                <td><% If Len(rsAl("cname") & "") > 0 Then Response.Write(Server.HTMLEncode(rsAl("cname") & "")) Else Response.Write("<em>(categor&iacute;a eliminada)</em>") End If %></td>
                <td><% If Len(rsAl("dname") & "") > 0 Then Response.Write(Server.HTMLEncode(rsAl("dname") & "")) Else Response.Write("&mdash;") End If %></td>
                <td><% If Len(rsAl("pname") & "") > 0 Then Response.Write(Server.HTMLEncode(rsAl("pname") & "")) Else Response.Write("&mdash;") End If %></td>
                <td><% If SafeCInt(rsAl("enabled")) <> 0 Then %><span class="mc-badge created">Activa</span><% Else %><span class="mc-badge ignored">Pausada</span><% End If %></td>
                <td class="col-actions">
                  <form method="post" action="mailconfig.asp" style="display:inline;margin:0"><input type="hidden" name="togglealias" value="<% = rsAl("id") %>"><button type="submit" class="btn btn-secondary btn-sm"><% If SafeCInt(rsAl("enabled")) <> 0 Then Response.Write("Pausar") Else Response.Write("Activar") End If %></button></form>
                  <form method="post" action="mailconfig.asp" style="display:inline;margin:0" onsubmit="return confirm('&iquest;Eliminar esta direcci&oacute;n?');"><input type="hidden" name="delalias" value="<% = rsAl("id") %>"><button type="submit" class="btn btn-secondary btn-sm">&#128465; Eliminar</button></form>
                </td>
              </tr>
              <%
                  rsAl.MoveNext
                Loop
                rsAl.Close
              %>
            </tbody>
          </table>
          <% If nAl = 0 Then %><div class="mc-empty">Todav&iacute;a no hay direcciones: todos los correos usan la categor&iacute;a por omisi&oacute;n.</div><% End If %>
        </div>
        <form method="post" action="mailconfig.asp" style="padding:16px;border-top:1px solid var(--color-border-light,#E2E8F0);margin:0">
          <input type="hidden" name="addalias" value="1">
          <div class="mc-grid">
            <div class="form-group"><label class="form-label" for="alias_address">Nueva direcci&oacute;n</label>
              <input type="text" class="form-control" id="alias_address" name="alias_address" maxlength="150" placeholder="redes@empresa.com" autocomplete="off"></div>
            <div class="form-group"><label class="form-label" for="alias_category"><%=lang(cnnDB, "Category")%> <span class="form-required">*</span></label>
              <select class="form-control" id="alias_category" name="alias_category"><% = MailSelectOptions("SELECT category_id, cname FROM categories WHERE category_id > 0 ORDER BY cname ASC", "category_id", "cname", 0, "&mdash; Elegir &mdash;") %></select></div>
            <div class="form-group"><label class="form-label" for="alias_department"><%=lang(cnnDB, "Department")%> (opcional)</label>
              <select class="form-control" id="alias_department" name="alias_department"><% = MailSelectOptions("SELECT department_id, dname FROM departments WHERE department_id > 0 ORDER BY dname ASC", "department_id", "dname", 0, "&mdash; Por omisi&oacute;n &mdash;") %></select></div>
            <div class="form-group"><label class="form-label" for="alias_priority"><%=lang(cnnDB, "Priority")%> (opcional)</label>
              <select class="form-control" id="alias_priority" name="alias_priority"><% = MailSelectOptions("SELECT priority_id, pname FROM priority ORDER BY priority_id ASC", "priority_id", "pname", 0, "&mdash; Por omisi&oacute;n &mdash;") %></select></div>
          </div>
          <button type="submit" class="btn btn-primary">&#10133; Agregar direcci&oacute;n</button>
        </form>
        <% End If %>
      </div>

      <div class="lhd-card mc-card">
        <div class="section-header">&#128228; Correo de prueba (SMTP)</div>
        <form method="post" action="mailconfig.asp" style="display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end;margin:0">
          <input type="hidden" name="sendtest" value="1">
          <div class="form-group" style="flex:1 1 280px;margin:0"><label class="form-label" for="test_to">Enviar a</label>
            <input type="email" class="form-control" id="test_to" name="test_to" maxlength="200" placeholder="tu.correo@empresa.com"></div>
          <button type="submit" class="btn btn-secondary">Poner en cola</button>
        </form>
        <div class="mc-hint">El correo sale en la siguiente ejecuci&oacute;n del script, con el servidor SMTP configurado arriba.</div>
      </div>

      <div class="lhd-card lhd-card-flush mc-card">
        <div class="section-header" style="margin:0;padding:14px 16px">&#128203; &Uacute;ltimos correos procesados</div>
        <div style="overflow-x:auto">
          <table class="lhd-table">
            <thead><tr><th>Recibido</th><th>Remitente</th><th>Asunto</th><th>Resultado</th><th>Caso</th><th>Detalle</th></tr></thead>
            <tbody>
              <%
                Dim rsLog, nLog, resTxt
                nLog = 0
                Set rsLog = SQLQuery(cnnDB, "SELECT TOP 30 received, from_addr, subject, result, problem_id, detail FROM tblMailLog ORDER BY id DESC")
                Do While Not rsLog.EOF
                  nLog = nLog + 1
                  Select Case LCase(rsLog("result") & "")
                    Case "created": resTxt = "Caso creado"
                    Case "note": resTxt = "Nota agregada"
                    Case "ignored": resTxt = "Ignorado"
                    Case "duplicate": resTxt = "Duplicado"
                    Case "error": resTxt = "Error"
                    Case "failed": resTxt = "Fallido"
                    Case Else: resTxt = rsLog("result") & ""
                  End Select
              %>
              <tr>
                <td style="white-space:nowrap"><% = DisplayDate(rsLog("received"), lhdDateTime) %></td>
                <td><% = Server.HTMLEncode(rsLog("from_addr") & "") %></td>
                <td><% = Server.HTMLEncode(Left(rsLog("subject") & "", 70)) %></td>
                <td><span class="mc-badge <% = LCase(rsLog("result") & "") %>"><% = resTxt %></span></td>
                <td><% If SafeCLng(rsLog("problem_id")) > 0 Then %><a href="../rep/details.asp?id=<% = rsLog("problem_id") %>">#<% = rsLog("problem_id") %></a><% Else Response.Write("&mdash;") End If %></td>
                <td style="font-size:12px;color:#64748b"><% = Server.HTMLEncode(Left(rsLog("detail") & "", 120)) %></td>
              </tr>
              <%
                  rsLog.MoveNext
                Loop
                rsLog.Close
              %>
            </tbody>
          </table>
          <% If nLog = 0 Then %><div class="mc-empty">Todav&iacute;a no se ha procesado ning&uacute;n correo.</div><% End If %>
        </div>
      </div>

      <% End If %>
    </div>

    <script type="text/javascript">
    (function () {
      var tok = document.getElementById('api_token');
      var out = document.getElementById('cmdTok');
      if (!tok) { return; }
      function sync() { out.textContent = tok.value || 'TOKEN'; }
      tok.addEventListener('input', sync);
      document.getElementById('btnTok').onclick = function () {
        var a = new Uint8Array(24), h = '', i;
        (window.crypto || window.msCrypto).getRandomValues(a);
        for (i = 0; i < a.length; i++) { h += ('0' + a[i].toString(16)).slice(-2); }
        tok.value = h; sync();
        if (window.LhdToast) { LhdToast.info('Token generado: recuerde guardar la configuraci\u00f3n.'); }
      };
      document.getElementById('btnCopy').onclick = function () {
        var txt = document.getElementById('cmdLine').textContent;
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(txt).then(function () { if (window.LhdToast) { LhdToast.success('Comando copiado.'); } });
        } else {
          var r = document.createRange(); r.selectNode(document.getElementById('cmdLine'));
          window.getSelection().removeAllRanges(); window.getSelection().addRange(r);
          if (window.LhdToast) { LhdToast.info('Comando seleccionado: presione Ctrl+C.'); }
        }
      };
    })();
    </script>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
