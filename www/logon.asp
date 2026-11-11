<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
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

  Filename: logon.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.52.4.1 $
  Purpose:  Get username and password from user.
  -->
  <!--	#include file = "settings.asp" -->
  <!-- 	#include file = "public.asp" -->

  <%
  Call SetAppVariables

  Dim cnnDB, sid
  Set cnnDB = CreateCon
  sid = GetSid

  %>

  <head>
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "HelpDesk")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="default.css">
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
        
          .al-links { display: flex; flex-wrap: wrap; gap: 8px 18px; }
          .al-links a { display: inline-flex; align-items: center; gap: 6px; }
          .al-btn { margin-top: 24px; }
        </style>
  </head>
  <body style="margin:0;padding:0">

    <%
      ' **************************************
      ' Get logon info.
      ' **************************************

      ' By default, the user will not have admin access.
      ' User must enter password on admin page.
      Dim username, url, invalid

      ' Solo se admiten rutas internas como destino tras el login: nada de http://..., //host,
      ' barras invertidas, saltos de linea ni esquemas (javascript:).
      Function LhdSafeReturnUrl(u)
        Dim q, c, bad
        bad = False
        If Len(u) = 0 Then bad = True
        If InStr(u, vbCr) > 0 Or InStr(u, vbLf) > 0 Or InStr(u, "\") > 0 Or InStr(u, "<") > 0 Or InStr(u, Chr(34)) > 0 Then bad = True
        If Left(u, 2) = "//" Then bad = True
        q = InStr(u, "?")
        c = InStr(u, ":")
        If c > 0 And (q = 0 Or c < q) Then bad = True
        If bad Then
          LhdSafeReturnUrl = "default.asp"
        Else
          LhdSafeReturnUrl = u
        End If
      End Function

      url = LhdSafeReturnUrl(Trim(Request.QueryString("URL")))

      Select Case Cfg(cnnDB, "AuthType")

        Case 1	' NT Authentication
          Dim domainLen
          username = Request.ServerVariables("AUTH_USER")

          If Len(username) = 0 Then
            Call DisplayError(3, lang(cnnDB, "UnabletoobtainusernamewithNTauthentication"))
          End If

          domainLen = InStr(username, "\")
          username = Mid(username, (domainLen+1), (Len(username)-domainLen))
          username = Lcase(username)

          If CBool(InStr(username, "'")) Then
            Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & Lang(cnnDB, "containsinvalidcharacters") & ".")
          End If

          Dim ntUserRes
          Set ntUserRes = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE (uid='" & username & "') AND (fname IS NOT NULL) OR (uid='" & username & "') AND (email1 IS NOT NULL)")
          If ntUserRes.EOF Then
            Dim ntUpdRes, ntSidRes
            ntUpdRes = SQLQuery(cnnDB, "INSERT INTO tblUsers (sid, uid) VALUES (" & GetUnique(cnnDB, "users") & ", '" & username & "')")
            ntSidRes = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & username & "'")
            Session("lhd_sid") = ntSidRes("sid")
            url = "register.asp?edit=1&new=1"
          Else
            Session("lhd_sid") = ntUserRes("sid")
          End If
          ntUserRes.Close

        Case 2	' DB Authentication
          If Request.Form("logon") = 1 Then
            Dim password
            username = Left(Lcase(Trim(Request.Form("uid"))), 50)
            username = Replace(username, "'", "''")
            password = Left(Trim(Request.Form("password")), 50)
            
            Dim userRes
            Set userRes = SQLQuery(cnnDB, "SELECT sid, password FROM tblUsers WHERE uid='" & username & "'")
            If userRes.EOF Then
              invalid = TRUE
              url=""
            ElseIf userRes("password") <> password Then
              invalid = TRUE
              frm_url = url
              url=""
            Else
              Session("lhd_sid") = userRes("sid")
            End If
            userRes.Close
          Else
            Dim frm_url
            frm_url = url
            url = ""
          End If

        Case 3	' External Authentication
          If Len(Session("lhd_ext_uid")) > 0 Then
            username = Lcase(Trim(Session("lhd_ext_uid")))
          ElseIf Len(Request.Form("lhd_ext_uid")) > 0 Then
            username = Lcase(Trim(Request.Form("lhd_ext_uid")))
          ElseIf Len(Request.QueryString("lhd_ext_uid")) > 0 Then
            username = Lcase(Trim(Request.QueryString("lhd_ext_uid")))
          Else
            Call DisplayError (3, lang(cnnDB, "Nousernamewasspecifiedbytheexternalauthenication") & ".")
          End If

          If CBool(InStr(username, "'")) Then
            Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & Lang(cnnDB, "containsinvalidcharacters") & ".")
          End If

          Dim extUserRes
          Set extUserRes = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & username & "'")
          If extUserRes.EOF Then
            Dim extUpdRes, extSidRes
            extUpdRes = SQLQuery(cnnDB, "INSERT INTO tblUsers (sid, uid) VALUES (" & GetUnique(cnnDB, "users") & ", '" & username & "')")
            extSidRes = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & username & "'")
            Session("lhd_sid") = extSidRes("sid")
            url = "register.asp?edit=1&new=1"
          Else
            Session("lhd_sid") = extUserRes("sid")
          End If
          extUserRes.Close

      End Select

      Session("lhd_IsAdmin") = FALSE
      Session("lhd_ViewMode") = Empty   ' al iniciar sesion, un tecnico entra con su vista de Representante

      If Len(url) > 0 Then
        ' Update logon time
        Dim updTimeRes
        Set updTimeRes = SQLQuery(cnnDB, "UPDATE tblUsers SET dtLastAccess=" & SQLDate(Now, lhdAddSQLDelim) & " WHERE sid=" & Session("lhd_sid"))
        cnnDB.Close
        Response.Redirect url
      End If
    %>

    <div class="al-wrap">
      <div class="al-card">

        <div class="al-side">
          <div>
            <div class="al-logo">
              <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5h16v11H9l-5 4V5z"/><path d="M8 9h8M8 12.5h5"/></svg>
            </div>
            <h2><% = Server.HTMLEncode(Cfg(cnnDB, "SiteName") & "") %></h2>
            <p><%=lang(cnnDB, "HelpDesk")%></p>
          </div>
          <ul class="al-list">
            <li><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>Registra y consulta tus solicitudes</li>
            <li><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>Seguimiento de cada caso en l&iacute;nea</li>
            <li><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>Respuestas y notificaciones por correo</li>
          </ul>
        </div>

        <div class="al-main">
          <h1><%=lang(cnnDB, "Logon")%></h1>
          <p class="al-sub">Ingresa con tu usuario y contrase&ntilde;a para continuar.</p>

          <% If invalid Then %>
            <div class="al-msg"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v6M12 16.5v.5"/></svg>
              <span><%=lang(cnnDB, "Invalidusernameorpassword")%>.</span></div>
          <% End If %>

          <form action="logon.asp?URL=<% = Server.URLEncode(frm_url & "") %>" method="POST" autocomplete="on" novalidate id="frmLogon">
            <input type="hidden" name="logon" value="1">

            <label class="al-label" for="uid"><%=lang(cnnDB, "UserName")%></label>
            <div class="al-field">
              <span class="al-ico"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-7 8-7s8 2.6 8 7"/></svg></span>
              <input type="text" id="uid" name="uid" autocomplete="username" placeholder="tu.usuario" autofocus maxlength="50" style="padding-right:16px">
            </div>

            <label class="al-label" for="password" style="margin-top:18px"><%=lang(cnnDB, "Password")%></label>
            <div class="al-field">
              <span class="al-ico"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="10" width="16" height="11" rx="2.5"/><path d="M8 10V7a4 4 0 018 0v3"/></svg></span>
              <input type="password" id="password" name="password" autocomplete="current-password" placeholder="Tu contrase&ntilde;a" maxlength="50">
              <button type="button" class="al-eye" id="admEye" title="Mostrar u ocultar" aria-label="Mostrar u ocultar la contrase&ntilde;a">
                <svg id="admEyeOn" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M1.5 12S5.5 5 12 5s10.5 7 10.5 7-4 7-10.5 7S1.5 12 1.5 12z"/><circle cx="12" cy="12" r="3"/></svg>
                <svg id="admEyeOff" style="display:none" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3l18 18M10.6 5.1A10.7 10.7 0 0112 5c6.5 0 10.5 7 10.5 7a17 17 0 01-3.2 4M6.5 6.6A17 17 0 001.5 12S5.5 19 12 19c1.8 0 3.4-.5 4.8-1.2"/></svg>
              </button>
            </div>
            <div class="al-caps" id="admCaps">Las may&uacute;sculas est&aacute;n activadas.</div>

            <button class="al-btn" id="admBtn" type="submit">
              <span class="al-spin"></span>
              <span><%=lang(cnnDB, "Logon")%></span>
              <svg class="al-arrow" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg>
            </button>
          </form>

          <div class="al-foot">
            <div class="al-links">
              <% If Cint(Cfg(cnnDB, "EnableKB")) = 3 Then %>
                <a href="kb/"><%=lang(cnnDB, "SearchtheKnowledgeBase")%></a>
              <% End If %>
              <a href="register.asp"><%=lang(cnnDB, "NewUser")%></a>
              <%
              If Cfg(cnnDB, "EmailType") <> 0 Then
                Response.Write "<a href=""forgotpass.asp"">" & lang(cnnDB, "EmailMyPassword") & "</a>"
              End If
              %>
            </div>
          </div>
        </div>

      </div>
    </div>

    <script type="text/javascript">
    (function () {
      var pw = document.getElementById('password');
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
      document.getElementById('frmLogon').addEventListener('submit', function () {
        btn.classList.add('loading');
      });
    })();
    </script>

    <%
      cnnDB.Close
    %>
  </body>
</html>
