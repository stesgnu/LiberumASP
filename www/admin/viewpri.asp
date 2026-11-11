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

  Filename: viewpri.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  View the list of priorities.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Priorities")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      Dim rstPriList
      Set rstPriList = SQLQuery(cnnDB, "SELECT priority_id, pname FROM priority WHERE priority_id > 0 ORDER BY priority_id ASC")

    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#9888;</span>
          <%=lang(cnnDB, "Priorities")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <form method="post" action="modify.asp?mtype=4" style="margin:0">
            <button type="submit" class="btn btn-primary">&#10133; <%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Priority")%></button>
          </form>
        </div>
      </div>

      <div class="lhd-card lhd-card-flush">
        <div style="overflow-x:auto">
          <table class="lhd-table">
            <thead>
              <tr>
                <th><%=lang(cnnDB, "ID")%></th>
                <th><%=lang(cnnDB, "Priority")%></th>
                <th class="col-actions"><%=lang(cnnDB, "Modify")%></th>
              </tr>
            </thead>
            <tbody>
            <%
              Do While Not rstPriList.EOF
            %>
              <tr>
                <td><span class="badge"><% = Server.HTMLEncode(rstPriList("priority_id") & "") %></span></td>
                <td><strong><% = Server.HTMLEncode(rstPriList("pname") & "") %></strong></td>
                <td class="col-actions">
                  <a class="btn btn-secondary btn-sm" href="modify.asp?mtype=4&amp;id=<% = rstPriList("priority_id") %>">&#9998; <%=lang(cnnDB, "Edit")%></a>
                  <a class="btn btn-secondary btn-sm" href="confdelete.asp?mtype=4&amp;id=<% = rstPriList("priority_id") %>">&#128465; <%=lang(cnnDB, "Delete")%></a>
                </td>
              </tr>
            <%
              rstPriList.MoveNext
              Loop
            %>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <%
      rstPriList.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
