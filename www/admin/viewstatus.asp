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

  Filename: viewstatus.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  View the list of statuses.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Reports")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      Dim rstStat
      Set rstStat = SQLQuery(cnnDB, "SELECT status_id, sname FROM status WHERE status_id > 0 ORDER BY status_id ASC")

    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128204;</span>
          <%=lang(cnnDB, "Statuses")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <form method="post" action="modify.asp?mtype=5" style="margin:0">
            <button type="submit" class="btn btn-primary">&#10133; <%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Status")%></button>
          </form>
        </div>
      </div>

      <div class="lhd-card lhd-card-flush">
        <div style="overflow-x:auto">
          <table class="lhd-table">
            <thead>
              <tr>
                <th><%=lang(cnnDB, "ID")%></th>
                <th><%=lang(cnnDB, "Status")%></th>
                <th class="col-actions"><%=lang(cnnDB, "Modify")%></th>
              </tr>
            </thead>
            <tbody>
            <%
              Do While Not rstStat.EOF
            %>
              <tr>
                <td>
                  <span class="badge"><% = Server.HTMLEncode(rstStat("status_id") & "") %></span>
                  <% If Cint(rstStat("status_id")) = Cfg(cnnDB, "CloseStatus") Then
                    Response.Write("<em>*</em>")
                    End If
                  %>
                </td>
                <td><strong><% = Server.HTMLEncode(rstStat("sname") & "") %></strong></td>
                <td class="col-actions">
                  <a class="btn btn-secondary btn-sm" href="modify.asp?mtype=5&amp;id=<% = rstStat("status_id") %>">&#9998; <%=lang(cnnDB, "Edit")%></a>
                  <a class="btn btn-secondary btn-sm" href="confdelete.asp?mtype=5&amp;id=<% = rstStat("status_id") %>">&#128465; <%=lang(cnnDB, "Delete")%></a>
                </td>
              </tr>
            <%
              rstStat.MoveNext
              Loop
            %>
            </tbody>
          </table>
        </div>
        <div style="padding:12px 16px;font-size:13px;color:#64748b;border-top:1px solid var(--color-border-light,#E2E8F0)"><em>*</em> <%=lang(cnnDB, "ClosedStatusDonotdelete")%>.</div>
      </div>
    </div>

    <%
      rstStat.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
