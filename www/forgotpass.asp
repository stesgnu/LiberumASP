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

  Filename: forgotpass.asp
  Date:     $Date: 2002/01/03 15:40:49 $
  Version:  $Revision: 1.51 $
  Purpose:  Emails the user their password.

  -->
  <!-- 	#include file = "public.asp" -->
  <%

    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    Dim success, invalidUID
    success = FALSE
    invalidUID = FALSE

    If SafeCInt(CleanInputText(Request.Form("email"))) = 1 Then
      Dim uid, userRes
      uid = Lcase(Trim(CleanInputText(Request.Form("uid"))))
      uid = Replace(uid, "'", "''")

      Set userRes = SQLQuery(cnnDB, "SELECT email1, password FROM tblUsers WHERE uid='" & uid & "'")

      If userRes.EOF Then
        invalidUID = TRUE
      Else
        Dim strSubject, strBody
        ' Send password to the user
        strSubject = lang(cnnDB, "HELPDESK") & " : " & lang(cnnDB, "password")
        strBody = _
        "Username: " & uid & vbNewLine & _
        "Password: " & userRes("password") & vbNewLine & _
        vbNewLine & _
        "Log in to the help desk @ " & Cfg(cnnDB, "BaseURL")

        Call SendMail(userRes("email1"), Cfg(cnnDB, "HDReply"), Cfg(cnnDB, "HDName"), strSubject, strBody, cnnDB)

        success = TRUE
      End IF
    End If
  %>

  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta charset="utf-8">
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "EmailPassword")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="default.css">
    <style>
      .fp-wrap {
        min-height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
        background: linear-gradient(135deg, var(--color-primary,#0057B8) 0%, var(--color-accent,#0082C8) 100%);
        padding: 20px;
      }
      .fp-card {
        background: #fff;
        border-radius: 12px;
        box-shadow: 0 20px 60px rgba(0,0,0,0.2);
        width: 100%;
        max-width: 400px;
        overflow: hidden;
      }
      .fp-header {
        background: linear-gradient(135deg, var(--color-primary,#0057B8), var(--color-accent,#0082C8));
        color: #fff;
        padding: 28px 32px 24px;
        text-align: center;
      }
      .fp-header .logo-icon { font-size: 40px; margin-bottom: 8px; }
      .fp-header h1 { font-size: 20px; font-weight: 700; margin: 0; }
      .fp-header p  { font-size: 13px; opacity: .8; margin: 4px 0 0; }
      .fp-body { padding: 28px 32px 24px; }
      .fp-hint { font-size: 13px; color: #4a5568; margin: 0 0 16px; line-height: 1.5; }
      .fp-done { text-align: center; }
      .fp-done .fp-ok { font-size: 44px; color: var(--color-success,#16A34A); margin-bottom: 6px; }
      .fp-done h2 { font-size: 18px; margin: 0 0 8px; color: #1a202c; }
      .fp-footer { padding: 0 32px 24px; text-align: center; font-size: 12px; color: #718096; }
      .fp-footer a { color: var(--color-primary,#0057B8); }
      .fp-footer a:hover { text-decoration: underline; }
    </style>
  </head>
  <body style="margin:0;padding:0">

    <div class="fp-wrap">
      <div class="fp-card">
        <div class="fp-header">
          <div class="logo-icon">&#128274;</div>
          <h1><%=lang(cnnDB, "EmailPassword")%></h1>
          <p><% = Server.HTMLEncode(Cfg(cnnDB, "SiteName") & "") %>&nbsp;<%=lang(cnnDB, "HelpDesk")%></p>
        </div>

        <div class="fp-body">
          <% If success Then %>
            <div class="fp-done">
              <div class="fp-ok">&#10004;</div>
              <h2><%=lang(cnnDB, "PasswordSent")%></h2>
              <p class="fp-hint"><%=lang(cnnDB, "PasswordSentText")%>.</p>
              <a class="btn btn-primary btn-full" href="logon.asp"><%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Logon")%> &rarr;</a>
            </div>
          <% Else %>
            <% If invalidUID Then %>
              <div class="alert alert-error">
                <span class="alert-icon">&#9888;</span>
                <span><%=lang(cnnDB, "Invalidusername")%>.</span>
              </div>
            <% End If %>
            <form action="forgotpass.asp" method="POST" autocomplete="on" novalidate>
              <input type="hidden" name="email" value="1">
              <div class="form-group">
                <label class="form-label" for="uid"><%=lang(cnnDB, "EnterUsername")%></label>
                <input class="form-control" type="text" id="uid" name="uid" autocomplete="username" autofocus maxlength="50">
              </div>
              <button class="btn btn-primary btn-full btn-lg" type="submit" style="margin-top:4px">
                <%=lang(cnnDB, "EmailPassword")%> &rarr;
              </button>
            </form>
          <% End If %>
        </div>

        <div class="fp-footer">
          <a href="logon.asp">&larr; <%=lang(cnnDB, "Logon")%></a>
        </div>
      </div>
    </div>

    <%
      cnnDB.Close
    %>
  </body>
</html>
