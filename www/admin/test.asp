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

  Filename: test.asp
  Date:     $Date: 2002/01/24 14:57:50 $
  Version:  $Revision: 1.1.2.1 $
  Purpose:  Misc commands for testing purpose
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=Lang(cnnDB, "TestConfiguration")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .tst-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 20px; align-items: start; }
      .tst-desc { color: #64748b; font-size: 13px; margin: 0 0 14px; }
      .tst-rows { margin: 0 0 16px; border: 1px solid var(--color-border-light, #E2E8F0); border-radius: 8px; overflow: hidden; }
      .tst-row { display: flex; justify-content: space-between; gap: 12px; padding: 9px 14px; font-size: 13px; }
      .tst-row + .tst-row { border-top: 1px solid var(--color-border-light, #E2E8F0); }
      .tst-row span:first-child { color: #64748b; }
      .tst-row span:last-child { font-weight: 600; color: #1a202c; text-align: right; word-break: break-all; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      ' Valores actuales de la configuracion de correo
      Dim cfgReply, cfgName, cfgSmtp
      cfgReply = Trim(Cfg(cnnDB, "HDReply") & "")
      cfgName = Trim(Cfg(cnnDB, "HDName") & "")
      cfgSmtp = Trim(Cfg(cnnDB, "SMTPServer") & "")

      ' Envio del correo de prueba (solo por POST: no se dispara con solo abrir un enlace)
      If Request.Form("doit") = "1" Then
        If Len(cfgReply) = 0 Then
          Call LhdToast("error", "No hay una direcci" & ChrW(243) & "n de respuesta configurada (HDReply). Config" & ChrW(250) & "rela en la configuraci" & ChrW(243) & "n del sitio.")
        Else
          Call SendMail(cfgReply, cfgReply, cfgName, "Test Message", "This is a test message from Liberum Help Desk", cnnDB)
          Call LhdToast("success", Lang(cnnDB, "Messagesentto") & " " & cfgReply)
        End If
      End If
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128295;</span>
          <%=Lang(cnnDB, "TestConfiguration")%>
        </div>
        <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
      </div>

      <div class="tst-grid">

        <div class="lhd-card">
          <div class="section-header">&#128231; Correo electr&oacute;nico</div>
          <p class="tst-desc">Env&iacute;a un mensaje de prueba a la direcci&oacute;n de respuesta del Help Desk para comprobar que el servidor de correo funciona.</p>
          <div class="tst-rows">
            <div class="tst-row"><span>Destinatario (HDReply)</span><span><% If Len(cfgReply) > 0 Then Response.Write(Server.HTMLEncode(cfgReply)) Else Response.Write("&mdash;") End If %></span></div>
            <div class="tst-row"><span>Nombre del remitente</span><span><% If Len(cfgName) > 0 Then Response.Write(Server.HTMLEncode(cfgName)) Else Response.Write("&mdash;") End If %></span></div>
            <div class="tst-row"><span>Servidor SMTP</span><span><% If Len(cfgSmtp) > 0 Then Response.Write(Server.HTMLEncode(cfgSmtp)) Else Response.Write("&mdash;") End If %></span></div>
          </div>
          <form method="post" action="test.asp" style="margin:0">
            <input type="hidden" name="doit" value="1">
            <button type="submit" class="btn btn-primary">&#128228; <%=Lang(cnnDB, "SendTestEmail")%></button>
          </form>
        </div>

        <div class="lhd-card">
          <div class="section-header">&#8505; Informaci&oacute;n del sistema</div>
          <p class="tst-desc">Consulta los datos del servidor y de la base de datos para diagnosticar problemas de instalaci&oacute;n o de conexi&oacute;n.</p>
          <a class="btn btn-secondary" href="sysinfo.asp">&#128269; <%=Lang(cnnDB, "ShowSystemInformation")%></a>
        </div>

      </div>
    </div>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
