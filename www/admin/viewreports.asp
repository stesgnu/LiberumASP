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

  Filename: viewreports.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Displays the report results with parameters from reports.asp.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Reports")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .vr-kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin-bottom: 18px; }
      .vr-kpi { background: var(--color-surface, #fff); border: 1px solid var(--color-border-light, #E2E8F0); border-left: 4px solid var(--color-primary, #0057B8); border-radius: 10px; padding: 14px 16px; box-shadow: var(--shadow-sm, 0 1px 3px rgba(0,0,0,.08)); }
      .vr-kpi .n { font-size: 26px; font-weight: 700; color: #1a202c; line-height: 1.1; }
      .vr-kpi .l { font-size: 12px; color: #64748b; text-transform: uppercase; letter-spacing: .04em; margin-top: 2px; }
      .vr-total td { font-weight: 700; background: #F1F5F9; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      Dim queryStr, start_date, end_date, intStartDay, intEndDay
      intStartDay = FixDay(Request.Form("s_month"), Request.Form("s_day"), Request.Form("s_year"))
      intEndDay = FixDay(Request.Form("e_month"), Request.Form("e_day"),  Request.Form("e_year"))
      start_date = Request.Form("s_month") & "/" & intStartDay & "/" & Request.Form("s_year")
      end_date = Request.Form("e_month") & "/" & intEndDay & "/" & Request.Form("e_year")
      start_date = start_date & " 00:00:00"
      end_date = end_date & " 23:59:59"

      ' Date criteria used in query.
      Dim dateCriteria
      dateCriteria = "p.start_date >= " & SQLDate(start_date, lhdAddSQLDelim) & " AND p.start_date <= " & SQLDate(end_date, lhdAddSQLDelim)

      ' Optional filters (0 = all)
      Dim fDept, fCat, fRep
      fDept = 0 : fCat = 0 : fRep = 0
      If IsNumeric(Request.Form("f_dept")) And Request.Form("f_dept") <> "" Then fDept = CLng(Request.Form("f_dept"))
      If IsNumeric(Request.Form("f_cat")) And Request.Form("f_cat") <> "" Then fCat = CLng(Request.Form("f_cat"))
      If IsNumeric(Request.Form("f_rep")) And Request.Form("f_rep") <> "" Then fRep = CLng(Request.Form("f_rep"))
      If fDept > 0 Then dateCriteria = dateCriteria & " AND p.department = " & fDept
      If fCat > 0 Then dateCriteria = dateCriteria & " AND p.category = " & fCat
      If fRep > 0 Then dateCriteria = dateCriteria & " AND p.rep = " & fRep

      ' Generate the sql queries
      Select Case Cint(Request.Form("type"))
        Case 0	' Departments
            queryStr = "SELECT d.dname AS name, Count(*) AS total, sum(p.time_spent) AS total_time " & _
              "FROM problems AS p INNER JOIN departments AS d ON p.department = d.department_id " & _
              "WHERE " & dateCriteria & " " & _
              "GROUP BY dname ORDER BY dname ASC"
        Case 1	' Categories
            queryStr = "SELECT c.cname AS name, Count(*) AS total, sum(p.time_spent) AS total_time " & _
              "FROM problems AS p INNER JOIN categories AS c ON p.category = c.category_id " & _
              "WHERE " & dateCriteria & " " & _
              "GROUP BY cname ORDER BY cname ASC"
        Case 2	' Reps
            queryStr = "SELECT r.uid AS name, Count(*) AS total, sum(p.time_spent) AS total_time " & _
              "FROM problems AS p INNER JOIN tblUsers AS r ON p.rep = r.sid " & _
              "WHERE " & dateCriteria & " AND r.sid>0 " & _
              "GROUP BY r.uid ORDER BY r.uid ASC"

      End Select ' end select
      
      Dim rstResults
      Set rstResults = SQLQuery(cnnDB, queryStr)

      Dim total, total_time
      total = 0
      total_time = 0
      If Not rstResults.EOF Then
        Do While NOT rstResults.EOF
          total = total + CLng(rstResults("total"))
          If Not IsNull(rstResults("total_time")) Then total_time = total_time + CLng(rstResults("total_time"))
          rstResults.MoveNext
        Loop
        rstResults.MoveFirst
      End If

    %>
    <div class="lhd-wrapper">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128202;</span>
          <%=lang(cnnDB, "Reports")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="reports.asp">&larr; <%=lang(cnnDB, "ReportsMenu")%></a>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>

      <% If Not rstResults.EOF Then %>
      <div class="vr-kpis">
        <div class="vr-kpi"><div class="n"><% = total %></div><div class="l"><%=lang(cnnDB, "Total")%></div></div>
        <div class="vr-kpi"><div class="n"><% = total_time %></div><div class="l"><%=lang(cnnDB, "Time")%></div></div>
        <div class="vr-kpi"><div class="n"><% If total > 0 Then Response.Write(FormatNumber((total_time / total), 1)) Else Response.Write("0") End If %></div><div class="l"><%=lang(cnnDB, "AvgTime")%></div></div>
        <div class="vr-kpi"><div class="n" style="font-size:15px;padding-top:6px"><% = Server.HTMLEncode(start_date) %><br>&ndash; <% = Server.HTMLEncode(end_date) %></div><div class="l">Periodo</div></div>
      </div>

      <div class="lhd-card lhd-card-flush">
        <div style="overflow-x:auto">
          <table class="lhd-table">
            <thead>
              <tr>
                <th>
                  <%
                    Select Case Cint(Request.Form("type"))
                      Case 0
                          Response.Write(lang(cnnDB, "Department"))
                      Case 1
                          Response.Write(lang(cnnDB, "Category"))
                      Case 2
                          Response.Write(lang(cnnDB, "Rep"))
                    End Select
                  %>
                </th>
                <th><%=lang(cnnDB, "Total")%></th>
                <th><%=lang(cnnDB, "Time")%></th>
                <th><%=lang(cnnDB, "AvgTime")%></th>
                <th><%=lang(cnnDB, "PofProblems")%></th>
                <th><%=lang(cnnDB, "PofTime")%></th>
              </tr>
            </thead>
            <tbody>
            <%
              Do While Not rstResults.EOF
                Dim rowTime
                rowTime = 0
                If Not IsNull(rstResults("total_time")) Then rowTime = CLng(rstResults("total_time"))
            %>
              <tr>
                <td><strong><% = Server.HTMLEncode(rstResults("name") & "") %></strong></td>
                <td><% = rstResults("total") %></td>
                <td><% = rowTime %></td>
                <td><% = FormatNumber((rowTime / CLng(rstResults("total"))), 1) %></td>
                <td><% = FormatNumber((CLng(rstResults("total")) / total * 100), 1) %>%</td>
                <td>
                  <% If total_time <> 0 Then
                    Response.Write(FormatNumber((rowTime / total_time * 100), 1))
                  Else
                    Response.Write("0")
                  End If
                  Response.Write("%") %>
                </td>
              </tr>
            <%
                rstResults.MoveNext
              Loop
              rstResults.Close
            %>
            </tbody>
            <tfoot>
              <tr class="vr-total">
                <td><%=lang(cnnDB, "Totals")%></td>
                <td><% = total %></td>
                <td><% = total_time %></td>
                <td><% If total > 0 Then Response.Write(FormatNumber((total_time / total), 1)) %></td>
                <td>100%</td>
                <td>100%</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
      <% Else ' no data %>
      <div class="lhd-card" style="text-align:center">
        <div style="font-size:15px;font-weight:600;margin-bottom:6px"><%=lang(cnnDB, "Noresultsfound")%></div>
        <div style="font-size:13px;color:#64748b">
          Periodo: <%=Server.HTMLEncode(start_date)%> &ndash; <%=Server.HTMLEncode(end_date)%>
          <% If fDept > 0 Or fCat > 0 Or fRep > 0 Then %>
            <br />Hay filtros activos (departamento, categor&iacute;a o representante). Prueba con "Todos" o ampl&iacute;a el periodo.
          <% End If %>
        </div>
      </div>
      <% End If %>
    </div>
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
