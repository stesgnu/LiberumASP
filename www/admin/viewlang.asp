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

  Filename: viewlang.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Manage the list of availbable Languages
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Languages")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      Dim rstLanguage
      Set rstLanguage = SQLQuery(cnnDB, "SELECT * FROM tblLanguage ORDER BY id ASC")

    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#127760;</span>
          <%=lang(cnnDB, "Languages")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <form method="post" action="modify.asp?mtype=6" style="margin:0">
            <button type="submit" class="btn btn-primary">&#10133; <%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Language")%></button>
          </form>
        </div>
      </div>

      <div class="lhd-card lhd-card-flush">
        <div style="overflow-x:auto">
          <table class="lhd-table">
            <thead>
              <tr>
                <th><%=lang(cnnDB, "Language")%></th>
                <th><%=lang(cnnDB, "Localized")%></th>
                <th class="col-actions"><%=lang(cnnDB, "Modify")%></th>
              </tr>
            </thead>
            <tbody>
            <%
              Do While Not rstLanguage.EOF
            %>
              <tr>
                <td><strong><% = Server.HTMLEncode(rstLanguage("LangName") & "") %></strong></td>
                <td><% = Server.HTMLEncode(rstLanguage("Localized") & "") %></td>
                <td class="col-actions">
                  <% If rstLanguage("id") <> 1 Then %>
                    <a class="btn btn-secondary btn-sm" href="modify.asp?mtype=6&amp;id=<% = rstLanguage("id") %>">&#9998; <%=lang(cnnDB, "Edit")%></a>
                    <a class="btn btn-secondary btn-sm" href="confdelete.asp?mtype=6&amp;id=<% = rstLanguage("id") %>">&#128465; <%=lang(cnnDB, "Delete")%></a>
                  <% End If %>
                  <a class="btn btn-secondary btn-sm" href="viewlangstring.asp?lang_id=<% = rstLanguage("id") %>">&#128292; <%=lang(cnnDB, "Strings")%></a>
                </td>
              </tr>
            <%
              rstLanguage.MoveNext
              Loop
            %>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <%
      rstLanguage.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
