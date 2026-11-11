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

  Filename: config_help.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Describes the various settings.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ConfigurationHelp")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .ch-wrap { max-width: 960px; margin: 0 auto; }
      .ch-term { width: 260px; font-weight: 600; vertical-align: top; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      ' Renders one row of the help table
      Sub HelpRow(termKey, helpKey)
        Response.Write "<tr><td class=""ch-term"">" & lang(cnnDB, termKey) & "</td><td>" & lang(cnnDB, helpKey) & "</td></tr>" & vbCrLf
      End Sub
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#8505;</span>
          <%=lang(cnnDB, "ConfigurationHelp")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="config.asp">&larr; <%=lang(cnnDB, "Configuration")%></a>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <button type="button" class="btn btn-secondary" onclick="window.close();"><%=lang(cnnDB, "CloseThisWindow")%></button>
        </div>
      </div>

      <div class="ch-wrap">
        <div class="lhd-card lhd-card-flush">
          <table class="lhd-table">
            <tbody>
              <% Call HelpRow("SiteName", "SiteNameHelp") %>
              <% Call HelpRow("BaseURL", "BaseURLHelp_2") %>
              <% Call HelpRow("AdministratorsName", "AdministratorsNameHelp") %>
              <% Call HelpRow("AdministratorsEmail", "AdministratorsEmailHelp") %>
              <% Call HelpRow("BaseEmail", "BaseEmailHelp") %>
              <% Call HelpRow("EmailType", "EmailTypeHelp") %>
              <% Call HelpRow("SMTPServer", "SMTPServerHelp") %>
              <% Call HelpRow("PagerPriorityLevel", "PagerPriorityLevelHelp") %>
              <% Call HelpRow("EmailUseronUpdate", "EmailUseronUpdateHelp") %>
              <% Call HelpRow("EnableUserKB", "EnableUserKBHelp") %>
              <% Call HelpRow("KBSQLFreeTextSearches", "KBSQLFreeTextSearchesHelp") %>
              <% Call HelpRow("DefaultPriority", "DefaultPriorityHelp") %>
              <% Call HelpRow("DefaultStatus", "DefaultStatusHelp") %>
              <% Call HelpRow("CloseStatus", "CloseStatusHelp") %>
              <% Call HelpRow("AuthenticationType", "AuthenticationTypeHelp") %>
              <% Call HelpRow("UseSelectUser", "UseSelectUserHelp") %>
              <% Call HelpRow("UseInOutBoard", "UseInOutBoardHelp") %>
              <% Call HelpRow("AllowImageUpload", "AllowImageUploadHelp") %>
              <% Call HelpRow("MaxImageSize", "MaxImageSizeHelp") %>
              <% Call HelpRow("Defaultlanguage", "DefaultlanguageHelp") %>
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
