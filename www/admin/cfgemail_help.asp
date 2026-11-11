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

  Filename: cfgemail_help.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Help page for configuring email.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "MessageConfigurationHelp")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .eh-wrap { max-width: 960px; margin: 0 auto; }
      .eh-term { width: 220px; vertical-align: top; }
      .eh-term code { font-family: Consolas, "Courier New", monospace; font-weight: 600; color: var(--color-primary, #0057B8); }
    </style>
  </head>
  <body>

  <%
    ' Check for perms to view this page
    Call CheckAdmin
    Call DisplayHeader(cnnDB, sid)

    ' Row with a language-key definition
    Sub VarRow(varName, helpKey)
      Response.Write "<tr><td class=""eh-term""><code>" & varName & "</code></td><td>" & lang(cnnDB, helpKey) & "</td></tr>" & vbCrLf
    End Sub

    ' Row with a literal definition
    Sub VarRowText(varName, helpText)
      Response.Write "<tr><td class=""eh-term""><code>" & varName & "</code></td><td>" & helpText & "</td></tr>" & vbCrLf
    End Sub
  %>

  <div class="lhd-wrapper">

    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
      <div class="lhd-page-title" style="margin:0">
        <span class="title-icon">&#8505;</span>
        <%=lang(cnnDB, "MessageConfigurationHelp")%>
      </div>
      <div style="display:flex;gap:10px;flex-wrap:wrap">
        <a class="btn btn-secondary" href="cfgemail.asp">&larr; <%=lang(cnnDB, "EmailMessages")%></a>
        <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        <button type="button" class="btn btn-secondary" onclick="window.close();"><%=lang(cnnDB, "CloseThisWindow")%></button>
      </div>
    </div>

    <div class="eh-wrap">
      <div class="lhd-card" style="margin-bottom:18px">
        <div class="section-header">&#128221; <%=lang(cnnDB, "MessageConfigurationHelp")%></div>
        <p><%=lang(cnnDB, "MessageConfigurationHelpText")%></p>
        <p>
          <%=lang(cnnDB, "ProblemID")%>: <b>[problemid]</b><br>
          <%=lang(cnnDB, "Title")%>: <b>[title]</b>
        </p>
      </div>

      <div class="lhd-card lhd-card-flush">
        <table class="lhd-table">
          <thead>
            <tr>
              <th><%=lang(cnnDB, "Variable")%></th>
              <th><%=lang(cnnDB, "Definition")%></th>
            </tr>
          </thead>
          <tbody>
            <% Call VarRow("problemid", "problemidHelp") %>
            <% Call VarRow("title", "titleHelp") %>
            <% Call VarRow("category", "categoryHelp") %>
            <% Call VarRow("status", "statusHelp") %>
            <% Call VarRow("priority", "priorityHelp") %>
            <% Call VarRow("startdate", "startdateHelp") %>
            <% Call VarRow("closedate", "closedateHelp") %>
            <% Call VarRow("department", "departmentHelp") %>
            <% Call VarRow("description", "descriptionHelp") %>
            <% Call VarRow("notes", "notesHelp") %>
            <% Call VarRow("solution", "solutionHelp") %>
            <% Call VarRow("uid", "uidHelp") %>
            <% Call VarRow("ufname", "ufnameHelp") %>
            <% Call VarRow("uemail", "uemailHelp") %>
            <% Call VarRow("phone", "phoneHelp") %>
            <% Call VarRow("location", "locationHelp") %>
            <% Call VarRow("uurl", "uurlHelp") %>
            <% Call VarRow("rid", "ridHelp") %>
            <% Call VarRow("rfname", "rfnameHelp") %>
            <% Call VarRow("remail", "remailHelp") %>
            <% Call VarRow("rurl", "rurlHelp") %>
            <% Call VarRow("baseurl", "baseurlHelp") %>
            <% Call VarRowText("u_title", "URL Encoded version of the title.") %>
            <% Call VarRowText("u_rfname", "URL Encoded version of the support representative's full name.") %>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <%
    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close
  %>
  </body>
</html>
