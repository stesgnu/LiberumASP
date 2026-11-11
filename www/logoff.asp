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

  Filename: logoff.asp
  Date:     $Date: 2002/01/24 15:10:05 $
  Version:  $Revision: 1.50.2.1 $
  Purpose:  Logs off the user by removing session variables.

  -->
  <!-- 	#include file = "public.asp" -->
  <%

    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "LogOff")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="default.css"> 
    <style type="text/css">
      .lo-wrap { min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 24px; box-sizing: border-box;
        background: radial-gradient(circle at 20% 15%, rgba(255,255,255,.14), transparent 45%), linear-gradient(135deg, var(--color-primary,#0057B8) 0%, var(--color-accent,#0082C8) 100%); }
      .lo-card { width: 100%; max-width: 420px; background: #fff; border-radius: 16px; box-shadow: 0 24px 70px rgba(0,0,0,.28); overflow: hidden; text-align: center; }
      .lo-head { padding: 30px 32px 6px; }
      .lo-ico { width: 64px; height: 64px; margin: 0 auto 14px; border-radius: 18px; display: flex; align-items: center; justify-content: center; font-size: 30px; color: #fff;
        background: linear-gradient(135deg, var(--color-success,#1A7F45), #2FA463); box-shadow: 0 10px 24px rgba(0,0,0,.2); }
      .lo-head h1 { margin: 0; font-size: 20px; font-weight: 700; color: #1a202c; }
      .lo-site { margin: 4px 0 0; font-size: 13px; color: #64748b; }
      .lo-body { padding: 14px 32px 28px; }
      .lo-body p { color: #475569; font-size: 14px; margin: 0 0 18px; }
    </style>
  </head>
  <body style="margin:0;padding:0">

    <div class="lo-wrap">
      <div class="lo-card">
        <div class="lo-head">
          <div class="lo-ico">&#10004;</div>
          <h1><%=lang(cnnDB, "Yoursessionhasbeenloggedoff")%></h1>
          <p class="lo-site"><% = Server.HTMLEncode(Cfg(cnnDB, "SiteName") & "") %> &middot; <%=lang(cnnDB, "HelpDesk")%></p>
        </div>
        <div class="lo-body">
          <p>Tu sesi&oacute;n se cerr&oacute; correctamente. Por seguridad, cierra la ventana del navegador si est&aacute;s en un equipo compartido.</p>
          <a class="btn btn-primary btn-full btn-lg" href="<% = Server.HTMLEncode(Cfg(cnnDB, "BaseURL") & "") %>/logon.asp">&rarr; <%=lang(cnnDB, "Clickheretologin")%></a>
        </div>
      </div>
    </div>

    <%
      Session("lhd_LanguageID") = Empty
      Session("lhd_IsAdmin") = False
      Session("lhd_ViewMode") = Empty
      Session("lhd_sid") = 0
      sid = 0

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
