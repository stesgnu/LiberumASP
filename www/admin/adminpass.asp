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

  Filename: adminpass.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Form to set the admin password.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ChangeAdminPassword")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .ap-wrap { max-width: 520px; margin: 0 auto; }
      .ap-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      Dim strMessage, passOk
      strMessage = ""
      passOk = False

      ' Save Results
      If Request.Form("save") = "1" Then
        Dim strSQL, AdminPass1, AdminPass2, CurrPass, OldPass, updRes
        AdminPass1 = Left(Trim(Request.Form("AdminPass1")), 50)
        AdminPass2 = Left(Trim(Request.Form("AdminPass2")), 50)
        CurrPass = Trim(Request.Form("CurrPass"))
        OldPass = Cfg(cnnDB, "AdminPass")

        strSQL = "UPDATE tblConfig SET " & _
          "AdminPass = '" & Replace(AdminPass1, "'", "''") & "'"

        If (AdminPass1 = AdminPass2) And (CurrPass = OldPass) Then
          Set updRes = SQLQuery(cnnDB, strSQL)
          strMessage = lang(cnnDB, "PasswordChanged")
          passOk = True
          Call LhdToast("success", strMessage)
        Else
          strMessage = lang(cnnDB, "PasswordChangeFailed")
          Call LhdToast("error", strMessage)
        End If

      End If

    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128273;</span>
          <%=lang(cnnDB, "ChangeAdminPassword")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>

      <div class="ap-wrap">
        <% If Request.Form("save") = "1" Then %>
          <% If passOk Then %>
          <div class="alert alert-success">
            <span class="alert-icon">&#10004;</span>
            <span><% = Server.HTMLEncode(strMessage & "") %></span>
          </div>
          <% Else %>
          <div class="alert alert-error">
            <span class="alert-icon">&#9888;</span>
            <span><% = Server.HTMLEncode(strMessage & "") %></span>
          </div>
          <% End If %>
        <% End If %>

        <form method="post" action="adminpass.asp" autocomplete="off">
          <input type="hidden" name="save" value="1">
          <div class="lhd-card">
            <div class="section-header">&#128274; <%=lang(cnnDB, "ChangeAdminPassword")%></div>
            <div class="form-group">
              <label class="form-label" for="CurrPass"><%=lang(cnnDB, "CurrentPassword")%></label>
              <input type="password" class="form-control" id="CurrPass" name="CurrPass" maxlength="50" autocomplete="current-password">
            </div>
            <div class="form-group">
              <label class="form-label" for="AdminPass1"><%=lang(cnnDB, "NewPassword")%></label>
              <input type="password" class="form-control" id="AdminPass1" name="AdminPass1" maxlength="50" autocomplete="new-password">
            </div>
            <div class="form-group">
              <label class="form-label" for="AdminPass2"><%=lang(cnnDB, "ConfirmPassword")%></label>
              <input type="password" class="form-control" id="AdminPass2" name="AdminPass2" maxlength="50" autocomplete="new-password">
            </div>
            <div class="ap-actions">
              <a class="btn btn-secondary" href="default.asp">Cancelar</a>
              <button type="submit" class="btn btn-primary">&#128190; <%=lang(cnnDB, "Save")%></button>
            </div>
          </div>
        </form>
      </div>
    </div>

    <%

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
