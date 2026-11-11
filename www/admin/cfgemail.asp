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

  Filename: cfgemailhtml.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Form to configure the email messages sent to users and reps.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "MessageConfiguration")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .ce-wrap { max-width: 860px; margin: 0 auto; }
      .ce-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; align-items: center; }
      .ce-body { font-family: Consolas, "Courier New", monospace; font-size: 13px; line-height: 1.5; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      Dim eType, displayMenu

      eType = Left(Trim(Request.Form("type")), 50)
      If Len(eType) > 0 Then
        displayMenu = False
      Else
        displayMenu = True
      End If

      ' Save Results
      If Request.Form("save") = "1" Then
        Dim strSQL, updRes
        Dim Subject, Body

        Subject = Left(Trim(CleanInputText(Request.Form("subject"))), 50)
        Body = Trim(CleanInputText(Request.Form("body")))

        strSQL = "UPDATE tblEmailMsg SET " & _
          "subject = '" & Subject & "', " & _
          "body = '" & Body & "' " & _
          "WHERE type='" & Replace(eType, "'", "''") & "'"

        Set updRes = SQLQuery(cnnDB, strSQL)
        Call LhdToast("success", lang(cnnDB, "MessageSaved"))

      End If

      If Not displayMenu Then
        Dim cfgRes
        ' Get current configuration
        Set cfgRes = SQLQuery(cnnDB, "Select Subject, Body From tblEmailMsg WHERE type='" & Replace(eType, "'", "''") & "'")

        If cfgRes.EOF Then
          Call DisplayError(3, lang(cnnDB, "Unable to read message from the database") & ".")
        End If
      End If
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#9993;</span>
          <%=lang(cnnDB, "EmailMessages")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <% If Not displayMenu Then %>
            <a class="btn btn-secondary" href="cfgemail.asp">&larr; <%=lang(cnnDB, "Chooseanothermessage")%></a>
          <% End If %>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>

      <div class="ce-wrap">
        <% If displayMenu Then %>
          <form method="post" action="cfgemail.asp">
            <div class="lhd-card">
              <div class="section-header">&#128221; <%=lang(cnnDB, "Message")%></div>
              <div class="form-group">
                <label class="form-label" for="type"><%=lang(cnnDB, "Message")%></label>
                <select class="form-control" id="type" name="type">
                  <option value="usernew"><%=lang(cnnDB, "User")%> - <%=lang(cnnDB, "New")%></option>
                  <option value="userupdate"><%=lang(cnnDB, "User")%> - <%=lang(cnnDB, "Update")%></option>
                  <option value="userclose"><%=lang(cnnDB, "User")%> - <%=lang(cnnDB, "Close")%></option>
                  <option value="repnew"><%=lang(cnnDB, "Rep")%> - <%=lang(cnnDB, "New")%></option>
                  <option value="repupdate"><%=lang(cnnDB, "Rep")%> - <%=lang(cnnDB, "Update")%></option>
                  <option value="repclose"><%=lang(cnnDB, "Rep")%> - <%=lang(cnnDB, "Close")%></option>
                  <option value="reppager"><%=lang(cnnDB, "Rep")%> - <%=lang(cnnDB, "Pager")%></option>
                </select>
              </div>
              <div class="ce-actions">
                <button type="submit" class="btn btn-primary">&#9998; <%=lang(cnnDB, "EditMessage")%></button>
              </div>
            </div>
          </form>
        <% Else %>
          <% If Request.Form("save") = "1" Then %>
            <div class="alert alert-success">
              <span class="alert-icon">&#10004;</span>
              <span><%=lang(cnnDB, "MessageSaved")%>.</span>
            </div>
          <% End If %>
          <form method="post" action="cfgemail.asp">
            <input type="hidden" name="save" value="1">
            <input type="hidden" name="type" value="<% = Server.HTMLEncode(eType) %>">
            <div class="lhd-card">
              <div class="section-header">&#128221; <%=lang(cnnDB, "Message")%>: <% = Server.HTMLEncode(eType) %></div>
              <div class="form-group">
                <label class="form-label" for="subject"><%=lang(cnnDB, "Subject")%></label>
                <input type="text" class="form-control" id="subject" name="subject" maxlength="50" value="<% = Server.HTMLEncode(cfgRes("Subject") & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="body"><%=lang(cnnDB, "Body")%> (Plain Text)</label>
                <textarea class="form-control ce-body" id="body" name="body" rows="12"><% = Server.HTMLEncode(cfgRes("body") & "") %></textarea>
              </div>
              <div class="ce-actions">
                <a class="btn btn-secondary" href="cfgemail_help.asp" target="_blank">&#8505; <%=lang(cnnDB, "SyntaxHelp")%></a>
                <button type="submit" class="btn btn-primary">&#128190; <%=lang(cnnDB, "Save")%></button>
              </div>
            </div>
          </form>
        <% End If %>
      </div>
    </div>
    <%
    '	cfgRes.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
