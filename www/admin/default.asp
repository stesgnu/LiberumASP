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

  Filename: default.asp (admin)
  Date:     $Date: 2002/01/24 14:57:50 $
  Version:  $Revision: 1.52.2.1 $
  Purpose:  Main administrative menu.
  -->
  
  <!--  #include file = "../settings.asp" -->
  <!-- 	#include file = "../public.asp" -->

  <% 
    Call SetAppVariables

    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "AdministrativeMenu")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style>
      .admin-card-icon {
        font-size: 18px;
        line-height: 1;
      }
    </style>
  </head>
  <body>

    <%
      ' Acceso administrativo: pide la clave de administrador.
      ' Proteccion contra fuerza bruta: despues de 5 claves incorrectas desde la misma
      ' direccion IP el acceso se bloquea 10 minutos.
      If Not Session("lhd_IsAdmin") Then

        Dim admIp, admKeyF, admKeyT, admLocked, admWait, admBad, admMax, admLockMin, admLeft, admWho
        admMax = 5
        admLockMin = 10
        admIp = Request.ServerVariables("REMOTE_ADDR") & ""
        admKeyF = "lhd_admfail_" & admIp
        admKeyT = "lhd_admlock_" & admIp
        admLocked = False
        admWait = 0
        admBad = False

        ' Bloqueo vigente?
        If Len(Application(admKeyT) & "") > 0 Then
          If IsDate(Application(admKeyT)) Then
            If Now < CDate(Application(admKeyT)) Then
              admLocked = True
              admWait = DateDiff("n", Now, CDate(Application(admKeyT))) + 1
            Else
              Application.Lock
              Application(admKeyT) = ""
              Application(admKeyF) = 0
              Application.UnLock
            End If
          End If
        End If

        ' Intento de acceso
        If Not admLocked And Len(Request.Form("password")) > 0 Then
          If Trim(Request.Form("password")) = Trim(Cfg(cnnDB, "AdminPass") & "") Then
            Session("lhd_IsAdmin") = TRUE
            Application.Lock
            Application(admKeyF) = 0
            Application(admKeyT) = ""
            Application.UnLock
          Else
            admBad = True
            Application.Lock
            Application(admKeyF) = SafeCInt(Application(admKeyF)) + 1
            If SafeCInt(Application(admKeyF)) >= admMax Then
              Application(admKeyT) = DateAdd("n", admLockMin, Now)
              Application(admKeyF) = 0
              admLocked = True
              admWait = admLockMin
            End If
            Application.UnLock
          End If
        End If

        If Not Session("lhd_IsAdmin") Then
          admLeft = admMax - SafeCInt(Application(admKeyF))
          admWho = ""
          If SafeCInt(sid) > 0 Then admWho = Trim(Usr(cnnDB, sid, "fname") & "")
          If Len(admWho) = 0 And SafeCInt(sid) > 0 Then admWho = Trim(Usr(cnnDB, sid, "uid") & "")
        %>
        <style type="text/css">
          .al-wrap { min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 24px; box-sizing: border-box;
            font-family: "Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif;
            background: radial-gradient(circle at 15% 10%, rgba(255,255,255,.16), transparent 42%), radial-gradient(circle at 90% 90%, rgba(255,255,255,.10), transparent 40%),
              linear-gradient(135deg, var(--color-primary,#0057B8) 0%, var(--color-accent,#0082C8) 100%); }
          .al-card { width: 100%; max-width: 880px; display: flex; background: #fff; border-radius: 24px; overflow: hidden; box-shadow: 0 30px 80px rgba(8,25,60,.35); }
          .al-side { flex: 1 1 45%; position: relative; padding: 44px 36px; color: #fff; display: flex; flex-direction: column; justify-content: space-between; overflow: hidden;
            background: linear-gradient(160deg, rgba(0,0,0,.18), rgba(0,0,0,.34)), linear-gradient(135deg, var(--color-primary,#0057B8), var(--color-accent,#0082C8)); }
          .al-side:before, .al-side:after { content: ""; position: absolute; border-radius: 50%; background: rgba(255,255,255,.10); }
          .al-side:before { width: 260px; height: 260px; right: -90px; top: -80px; }
          .al-side:after { width: 200px; height: 200px; left: -70px; bottom: -70px; }
          .al-side > * { position: relative; z-index: 1; }
          .al-logo { width: 58px; height: 58px; border-radius: 16px; background: rgba(255,255,255,.18); display: flex; align-items: center; justify-content: center; backdrop-filter: blur(4px); }
          .al-side h2 { margin: 22px 0 8px; font-size: 26px; line-height: 1.2; font-weight: 700; color: #fff; }
          .al-side p { margin: 0; font-size: 14px; line-height: 1.55; color: rgba(255,255,255,.86); }
          .al-list { list-style: none; margin: 26px 0 0; padding: 0; font-size: 13px; color: rgba(255,255,255,.92); }
          .al-list li { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
          .al-list svg { flex: 0 0 auto; }
          .al-main { flex: 1 1 55%; padding: 44px 44px 30px; background: #fff; display: flex; flex-direction: column; justify-content: center; }
          .al-main h1 { margin: 0 0 6px; font-size: 24px; font-weight: 700; color: #0F172A; letter-spacing: -.2px; }
          .al-sub { margin: 0 0 26px; font-size: 14px; color: #64748B; }
          .al-who { display: inline-flex; align-items: center; gap: 8px; margin: -12px 0 20px; padding: 6px 12px; border-radius: 999px; background: #F1F5F9; font-size: 12px; color: #475569; }
          .al-msg { display: flex; gap: 10px; align-items: flex-start; padding: 12px 14px; margin-bottom: 18px; border-radius: 12px; font-size: 13px; line-height: 1.45; background: #FEF2F2; color: #991B1B; border: 1px solid #FECACA; }
          .al-msg svg { flex: 0 0 auto; margin-top: 1px; }
          .al-label { display: block; margin-bottom: 8px; font-size: 13px; font-weight: 600; color: #334155; }
          .al-field { position: relative; }
          .al-field .al-ico { position: absolute; left: 16px; top: 50%; transform: translateY(-50%); color: #94A3B8; pointer-events: none; display: flex; transition: color .2s; }
          .al-field input { width: 100%; box-sizing: border-box; height: 54px; padding: 0 52px 0 48px; font-size: 15px; font-family: inherit; color: #0F172A;
            background: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; outline: none; transition: border-color .2s, box-shadow .2s, background .2s; }
          .al-field input::placeholder { color: #94A3B8; }
          .al-field input:hover { border-color: #CBD5E1; }
          .al-field input:focus { background: #fff; border-color: var(--color-primary,#0057B8); box-shadow: 0 0 0 4px rgba(0,87,184,.14); }
          .al-field:focus-within .al-ico { color: var(--color-primary,#0057B8); }
          .al-field input:disabled { opacity: .6; cursor: not-allowed; }
          .al-eye { position: absolute; right: 8px; top: 50%; transform: translateY(-50%); width: 38px; height: 38px; border: 0; background: transparent; cursor: pointer; color: #64748B; border-radius: 10px; display: flex; align-items: center; justify-content: center; transition: background .2s, color .2s; }
          .al-eye:hover { background: #E2E8F0; color: #0F172A; }
          .al-caps { display: none; margin-top: 10px; font-size: 12px; color: #92400E; background: #FEF3C7; border-radius: 10px; padding: 7px 12px; }
          .al-btn { margin-top: 22px; width: 100%; height: 54px; border: 0; border-radius: 14px; cursor: pointer; font-family: inherit; font-size: 15px; font-weight: 600; letter-spacing: .2px; color: #fff;
            display: flex; align-items: center; justify-content: center; gap: 10px;
            background: linear-gradient(135deg, var(--color-primary,#0057B8), var(--color-accent,#0082C8)); box-shadow: 0 10px 24px rgba(0,87,184,.32);
            transition: transform .15s, box-shadow .2s, filter .2s; }
          .al-btn:hover { transform: translateY(-2px); box-shadow: 0 14px 30px rgba(0,87,184,.40); filter: brightness(1.05); }
          .al-btn:active { transform: translateY(0); box-shadow: 0 6px 14px rgba(0,87,184,.30); }
          .al-btn:focus-visible { outline: 3px solid rgba(0,87,184,.35); outline-offset: 3px; }
          .al-btn:disabled { opacity: .55; cursor: not-allowed; transform: none; box-shadow: none; }
          .al-btn .al-spin { display: none; width: 18px; height: 18px; border: 2.5px solid rgba(255,255,255,.45); border-top-color: #fff; border-radius: 50%; animation: alspin .7s linear infinite; }
          .al-btn.loading .al-spin { display: block; }
          .al-btn.loading .al-arrow { display: none; }
          @keyframes alspin { to { transform: rotate(360deg); } }
          .al-foot { display: flex; justify-content: space-between; align-items: center; gap: 10px; margin-top: 26px; padding-top: 18px; border-top: 1px solid #EEF2F7; font-size: 13px; }
          .al-foot a { color: var(--color-primary,#0057B8); text-decoration: none; font-weight: 600; }
          .al-foot a:hover { text-decoration: underline; }
          .al-note { margin-top: 14px; font-size: 12px; color: #94A3B8; }
          @media (max-width: 760px) {
            .al-card { flex-direction: column; max-width: 440px; border-radius: 20px; }
            .al-side { padding: 26px 26px 22px; }
            .al-side h2 { font-size: 20px; margin-top: 14px; }
            .al-list { display: none; }
            .al-main { padding: 28px 24px 22px; }
          }
        </style>

        <div class="al-wrap">
          <div class="al-card">

            <div class="al-side">
              <div>
                <div class="al-logo">
                  <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l8 3v6c0 5-3.4 9.3-8 11-4.6-1.7-8-6-8-11V5l8-3z"/><path d="M9 12l2.2 2.2L15.5 10"/></svg>
                </div>
                <h2>Panel de administraci&oacute;n</h2>
                <p><% = Server.HTMLEncode(Cfg(cnnDB, "SiteName") & "") %></p>
              </div>
              <ul class="al-list">
                <li><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>Configuraci&oacute;n y mantenimiento del sistema</li>
                <li><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>Usuarios, categor&iacute;as y departamentos</li>
                <li><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>Informes y cumplimiento de ANS</li>
              </ul>
            </div>

            <div class="al-main">
              <h1><%=lang(cnnDB, "AdministrativeLogon")%></h1>
              <p class="al-sub">Ingresa la clave de administrador para continuar.</p>
              <% If Len(admWho) > 0 Then %><div class="al-who"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-7 8-7s8 2.6 8 7"/></svg>Sesi&oacute;n iniciada como <strong><% = Server.HTMLEncode(admWho) %></strong></div><% End If %>

              <% If admLocked Then %>
                <div class="al-msg"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="10" width="16" height="11" rx="2"/><path d="M8 10V7a4 4 0 018 0v3"/></svg>
                  <span>Demasiados intentos fallidos. Por seguridad el acceso est&aacute; bloqueado. Intenta de nuevo en <strong><% = admWait %> minuto<% If admWait <> 1 Then Response.Write("s") %></strong>.</span></div>
              <% ElseIf admBad Then %>
                <div class="al-msg"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v6M12 16.5v.5"/></svg>
                  <span><%=lang(cnnDB, "Passwordisincorrect")%>. Te quedan <strong><% = admLeft %></strong> intento<% If admLeft <> 1 Then Response.Write("s") %>.</span></div>
              <% End If %>

              <form method="post" action="default.asp" id="frmAdmLogin" autocomplete="off">
                <label class="al-label" for="admin-pwd"><%=lang(cnnDB, "Password")%></label>
                <div class="al-field">
                  <span class="al-ico"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="10" width="16" height="11" rx="2.5"/><path d="M8 10V7a4 4 0 018 0v3"/></svg></span>
                  <input type="password" id="admin-pwd" name="password" maxlength="50" autocomplete="current-password" placeholder="Clave de administrador" autofocus<% If admLocked Then Response.Write(" disabled") %>>
                  <button type="button" class="al-eye" id="admEye" title="Mostrar u ocultar" aria-label="Mostrar u ocultar la clave">
                    <svg id="admEyeOn" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M1.5 12S5.5 5 12 5s10.5 7 10.5 7-4 7-10.5 7S1.5 12 1.5 12z"/><circle cx="12" cy="12" r="3"/></svg>
                    <svg id="admEyeOff" style="display:none" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3l18 18M10.6 5.1A10.7 10.7 0 0112 5c6.5 0 10.5 7 10.5 7a17 17 0 01-3.2 4M6.5 6.6A17 17 0 001.5 12S5.5 19 12 19c1.8 0 3.4-.5 4.8-1.2"/></svg>
                  </button>
                </div>
                <div class="al-caps" id="admCaps">Las may&uacute;sculas est&aacute;n activadas.</div>
                <button class="al-btn" id="admBtn" type="submit"<% If admLocked Then Response.Write(" disabled") %>>
                  <span class="al-spin"></span>
                  <span><%=lang(cnnDB, "Logon")%></span>
                  <svg class="al-arrow" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg>
                </button>
              </form>

              <div class="al-foot">
                <a href="../default.asp">&larr; Volver a la aplicaci&oacute;n</a>
                <a href="../logoff.asp"><%=lang(cnnDB, "LogOff")%></a>
              </div>
              <div class="al-note">Despu&eacute;s de <% = admMax %> intentos fallidos el acceso se bloquea <% = admLockMin %> minutos.</div>
            </div>

          </div>
        </div>

        <script type="text/javascript">
        (function () {
          var pw = document.getElementById('admin-pwd');
          var eye = document.getElementById('admEye');
          var caps = document.getElementById('admCaps');
          var on = document.getElementById('admEyeOn');
          var off = document.getElementById('admEyeOff');
          var btn = document.getElementById('admBtn');
          eye.onclick = function () {
            var show = (pw.type === 'password');
            pw.type = show ? 'text' : 'password';
            on.style.display = show ? 'none' : '';
            off.style.display = show ? '' : 'none';
            pw.focus();
          };
          function capsCheck(e) {
            var c = e.getModifierState && e.getModifierState('CapsLock');
            caps.style.display = c ? 'block' : 'none';
          }
          pw.addEventListener('keyup', capsCheck);
          pw.addEventListener('keydown', capsCheck);
          pw.addEventListener('blur', function () { caps.style.display = 'none'; });
          document.getElementById('frmAdmLogin').addEventListener('submit', function () {
            if (pw.value.length > 0) { btn.classList.add('loading'); }
          });
        })();
        </script>

        <%
          Call DisplayFooter(cnnDB, sid)
          Response.Expires = -1
          Response.End
        End If
      End If

      Call DisplayHeader(cnnDB, sid)
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#9881;</span>
          <%=lang(cnnDB, "AdministrativeMenu")%>
        </div>
        <div>
          <a class="btn btn-secondary btn-sm" href="../rep/">&#8592; Panel de Soporte</a>
        </div>
      </div>

      <% Call ShowToast(cnnDB) %>

      <style type="text/css">
        .adm-toolbar { margin: 0 0 22px; }
        .adm-search { width: 100%; max-width: 440px; padding: 11px 16px; border: 1px solid #cbd5e0; border-radius: 10px; font-size: 14px; background: #fff; color: #1e293b; box-sizing: border-box; }
        .adm-search:focus { outline: none; border-color: var(--lhd-primary, #0057B8); box-shadow: 0 0 0 3px rgba(0,87,184,.15); }
        .adm-section { margin-bottom: 30px; }
        .adm-section h3 { font-size: 12px; text-transform: uppercase; letter-spacing: .8px; color: #64748b; margin: 0 0 12px; font-weight: 700; }
        .adm-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 14px; }
        .adm-card { display: flex; align-items: center; gap: 14px; background: #fff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px; text-decoration: none; color: #1e293b; box-shadow: 0 1px 3px rgba(15,23,42,.06); transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease; }
        .adm-card:hover { transform: translateY(-2px); box-shadow: 0 8px 18px rgba(15,23,42,.10); border-color: var(--accent, #0057B8); text-decoration: none; }
        .adm-ico { flex: 0 0 44px; width: 44px; height: 44px; border-radius: 11px; display: flex; align-items: center; justify-content: center; color: var(--accent, #0057B8); background: var(--accent-soft, #EFF6FF); }
        .adm-ico svg { width: 22px; height: 22px; }
        .adm-body { display: flex; flex-direction: column; min-width: 0; }
        .adm-body strong { font-size: 14px; font-weight: 600; line-height: 1.3; }
        .adm-body small { font-size: 12px; color: #64748b; margin-top: 3px; line-height: 1.4; }
        .adm-arrow { margin-left: auto; color: #94a3b8; font-size: 22px; line-height: 1; }
        .adm-empty { display: none; color: #64748b; padding: 24px 0; }
        @media (max-width: 640px) { .adm-grid { grid-template-columns: 1fr; } .adm-search { max-width: none; } }
      </style>

      <div class="adm-toolbar">
        <input type="search" id="admSearch" class="adm-search" placeholder="Buscar en el men&uacute; administrativo..." aria-label="Buscar en el men&uacute; administrativo" />
      </div>

      <%
        Function AdmIcon(name)
          Dim path
          Select Case name
            Case "gear": path = "<circle cx=""12"" cy=""12"" r=""3""/><path d=""M12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1""/>"
            Case "palette": path = "<circle cx=""12"" cy=""12"" r=""9""/><circle cx=""8"" cy=""10"" r=""1.2""/><circle cx=""12"" cy=""7.5"" r=""1.2""/><circle cx=""16"" cy=""10"" r=""1.2""/>"
            Case "mail": path = "<rect x=""3"" y=""5"" width=""18"" height=""14"" rx=""2""/><path d=""M3 7l9 6 9-6""/>"
            Case "key": path = "<circle cx=""8"" cy=""15"" r=""4""/><path d=""M11 12l9-9M17 6l3 3M14 9l2 2""/>"
            Case "check": path = "<path d=""M5 12l5 5 9-10""/>"
            Case "users": path = "<circle cx=""9"" cy=""8"" r=""3.5""/><path d=""M2.5 20c.8-3.3 3.3-5 6.5-5s5.7 1.7 6.5 5""/><path d=""M16 4.5a3.5 3.5 0 010 7M18 15c1.7.6 2.8 2.2 3 5""/>"
            Case "userplus": path = "<circle cx=""9"" cy=""8"" r=""3.5""/><path d=""M2.5 20c.8-3.3 3.3-5 6.5-5s5.7 1.7 6.5 5M19 8v6M16 11h6""/>"
            Case "globe": path = "<circle cx=""12"" cy=""12"" r=""9""/><path d=""M3 12h18M12 3c3 3.5 3 14.5 0 18M12 3c-3 3.5-3 14.5 0 18""/>"
            Case "folder": path = "<path d=""M3 6a2 2 0 012-2h4l2 2h8a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2z""/>"
            Case "building": path = "<rect x=""5"" y=""3"" width=""14"" height=""18"" rx=""1""/><path d=""M9 7h2M13 7h2M9 11h2M13 11h2M9 15h2M13 15h2""/>"
            Case "flag": path = "<path d=""M5 21V4M5 4h11l-2 4 2 4H5""/>"
            Case "status": path = "<circle cx=""12"" cy=""12"" r=""9""/><path d=""M12 7v5l3 2""/>"
            Case "chart": path = "<path d=""M4 20V10M10 20V4M16 20v-7M22 20H2""/>"
            Case "info": path = "<circle cx=""12"" cy=""12"" r=""9""/><path d=""M12 11v6M12 7.5v.5""/>"
            Case "clip": path = "<path d=""M21 11l-8.5 8.5a5 5 0 01-7-7L14 4a3.5 3.5 0 015 5l-8.5 8.5a2 2 0 01-3-3L15 7""/>"
            Case "text": path = "<path d=""M4 7V4h16v3M9 20h6M12 4v16""/>"
            Case "database": path = "<ellipse cx=""12"" cy=""5"" rx=""8"" ry=""3""/><path d=""M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3""/>"
            Case Else: path = "<circle cx=""12"" cy=""12"" r=""9""/>"
          End Select
          AdmIcon = "<svg viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & path & "</svg>"
        End Function

        Function AdmCard(icon, href, title, desc, accent)
          AdmCard = "<a class=""adm-card"" href=""" & href & """ style=""--accent:" & accent & ";--accent-soft:" & accent & "1f"">" & _
            "<span class=""adm-ico"">" & AdmIcon(icon) & "</span>" & _
            "<span class=""adm-body""><strong>" & title & "</strong><small>" & desc & "</small></span>" & _
            "<span class=""adm-arrow"">&rsaquo;</span></a>" & vbNewLine
        End Function

        Function AdmSection(title, cards)
          AdmSection = "<section class=""adm-section""><h3>" & title & "</h3><div class=""adm-grid"">" & cards & "</div></section>" & vbNewLine
        End Function

        Dim admHtml
        admHtml = ""
        admHtml = admHtml & AdmSection("Configuraci&oacute;n general", _
          AdmCard("gear", "config.asp", lang(cnnDB, "Configure") & "&nbsp;" & lang(cnnDB, "Site"), "Opciones generales del sitio y base de datos", GetThemeColor(cnnDB, "primary", "#0057B8")) & _
          AdmCard("palette", "colors.asp", "Personalizar colores y apariencia", "Paleta de colores, barra lateral y tema", "#7C3AED") & _
          AdmCard("mail", "cfgemail.asp", lang(cnnDB, "Configure") & "&nbsp;" & lang(cnnDB, "EmailMessages"), "Servidor SMTP y plantillas de correo", "#D97706") & _
          AdmCard("mail", "mailconfig.asp", "Correo entrante (mesa de servicio)", "Crear y actualizar casos desde un buz&oacute;n IMAP/SMTP", "#0284C7") & _
          AdmCard("key", "adminpass.asp", lang(cnnDB, "ChangeAdminPassword"), "Actualizar la contrase&ntilde;a administrativa", "#DC2626") & _
          AdmCard("check", "test.asp", lang(cnnDB, "TestConfiguration"), "Diagn&oacute;stico de conexi&oacute;n y componentes", "#16A34A"))

        admHtml = admHtml & AdmSection("Usuarios e idiomas", _
          AdmCard("users", "viewusers.asp", lang(cnnDB, "Manage") & "&nbsp;" & lang(cnnDB, "Users"), "Ver, editar y asignar roles a usuarios", GetThemeColor(cnnDB, "primary", "#0057B8")) & _
          AdmCard("userplus", "adduser.asp", "Nuevo usuario", "Registrar t&eacute;cnico o usuario manualmente", "#0F766E") & _
          AdmCard("globe", "viewlang.asp", lang(cnnDB, "Manage") & "&nbsp;" & lang(cnnDB, "Languages"), "Idiomas disponibles en la plataforma", "#0284C7"))

        admHtml = admHtml & AdmSection("Clasificaci&oacute;n de tickets", _
          AdmCard("folder", "viewcat.asp", lang(cnnDB, "Manage") & "&nbsp;" & lang(cnnDB, "Categories"), "Categor&iacute;as y asignaci&oacute;n autom&aacute;tica a t&eacute;cnicos", "#D97706") & _
          AdmCard("building", "viewdep.asp", lang(cnnDB, "Manage") & "&nbsp;" & lang(cnnDB, "Departments"), "Departamentos y &aacute;reas de la organizaci&oacute;n", "#7C3AED") & _
          AdmCard("flag", "viewpri.asp", lang(cnnDB, "Manage") & "&nbsp;" & lang(cnnDB, "Priorities"), "Niveles de urgencia e impacto", "#DC2626") & _
          AdmCard("status", "viewstatus.asp", lang(cnnDB, "Manage") & "&nbsp;" & lang(cnnDB, "Statuses"), "Estados del ticket y su flujo", "#0369A1"))

        admHtml = admHtml & AdmSection("Mantenimiento", _
          AdmCard("database", "install_attachments.asp", "Base de datos de adjuntos", "Crear o actualizar la tabla de adjuntos", "#0F766E") & _
          AdmCard("text", "repair_text.asp", "Reparar texto acentuado", "Corrige tildes y e&ntilde;es da&ntilde;adas en la base", "#7C3AED") & _
          AdmCard("info", "sysinfo.asp", "Informaci&oacute;n del sistema", "Versi&oacute;n de IIS, VBScript y componentes COM", "#64748b"))

        ' Informes: aqui se agregan los reportes nuevos (una tarjeta AdmCard por informe)
        admHtml = admHtml & AdmSection("Informes", _
          AdmCard("chart", "reports.asp", lang(cnnDB, "Reports"), "Informes de tickets por periodo", GetThemeColor(cnnDB, "primary", "#0057B8")) & _
          AdmCard("chart", "reportans.asp", "Cumplimiento de ANS", "Casos que cumplen y no cumplen el ANS", "#16A34A"))

        admHtml = admHtml & "<p class=""adm-empty"" id=""admEmpty"">No hay opciones que coincidan con la b&uacute;squeda.</p>"
        Response.Write admHtml
      %>

      <script type="text/javascript">
        (function () {
          var q = document.getElementById('admSearch');
          if (!q) { return; }
          q.addEventListener('input', function () {
            var v = q.value.toLowerCase();
            var visibles = 0;
            document.querySelectorAll('.adm-card').forEach(function (c) {
              var ok = c.textContent.toLowerCase().indexOf(v) > -1;
              c.style.display = ok ? '' : 'none';
            });
            document.querySelectorAll('.adm-section').forEach(function (s) {
              var n = 0;
              s.querySelectorAll('.adm-card').forEach(function (c) { if (c.style.display !== 'none') { n++; } });
              s.style.display = n ? '' : 'none';
              visibles += n;
            });
            document.getElementById('admEmpty').style.display = visibles ? 'none' : 'block';
          });
        })();
      </script>
    </div>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
