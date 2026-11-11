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

  Filename: search.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  This page displays a form of search paramaters which are posted
    to results.asp.

  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

<head>
  <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "ProblemSearch")%></title>
  <link rel="stylesheet" type="text/css" href="../default.css">
</head>
<body>

<%

	' Check if user has permissions for this page
	Call CheckRep(cnnDB, sid)
	Call DisplayHeader(cnnDB, sid)

%>
<div class="lhd-wrapper">

  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
    <div class="lhd-page-title" style="margin:0">
      <span class="title-icon">&#128269;</span>
      <%=lang(cnnDB, "ProblemSearch")%>
    </div>
    <a class="btn btn-secondary btn-sm" href="view.asp">&larr; Volver a la Lista</a>
  </div>

  <div class="lhd-card" style="max-width:820px;margin:0 auto">
    <div class="section-header">&#128269; <%=lang(cnnDB, "ProblemSpecifications")%></div>

    <form method="post" action="results.asp">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">

        <div class="form-group">
          <label class="form-label"><b><%=lang(cnnDB, "ReportedBy")%>:</b></label>
          <input type="text" name="uid" class="form-control" placeholder="Nombre de usuario">
        </div>

        <div class="form-group">
          <label class="form-label"><b><%=lang(cnnDB, "ProblemID")%>:</b></label>
          <input type="number" name="id" class="form-control" placeholder="N&uacute;mero de ticket">
        </div>

        <div class="form-group">
          <label class="form-label"><b><%=lang(cnnDB, "AssignedTo")%>:</b></label>
          <select name="rep" class="form-control">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <%
              Dim repRes
              Set repRes = SQLQuery(cnnDB, "SELECT sid, uid From tblUsers WHERE IsRep = 1 AND RepAccess <> 2 AND sid > 0 ORDER BY uid ASC")
              If Not repRes.EOF Then
                Do While Not repRes.EOF
                  Response.Write "<option value=""" & repRes("sid") & """>" & Server.HTMLEncode(repRes("uid")) & "</option>" & vbNewLine
                  repRes.MoveNext
                Loop
                repRes.Close
              End If
            %>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label"><b><%=lang(cnnDB, "Category")%>:</b></label>
          <select name="category" class="form-control">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <%
              Dim catRes
              Set catRes = SQLQuery(cnnDB, "SELECT category_id, cname From categories WHERE category_id > 0 ORDER BY cname ASC")
              If Not catRes.EOF Then
                Do While Not catRes.EOF
                  Response.Write "<option value=""" & catRes("category_id") & """>" & FixTildes(catRes("cname")) & "</option>" & vbNewLine
                  catRes.MoveNext
                Loop
                catRes.Close
              End If
            %>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label"><b><%=lang(cnnDB, "Department")%>:</b></label>
          <select name="department" class="form-control">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <%
              Dim depRes
              Set depRes = SQLQuery(cnnDB, "SELECT department_id, dname From departments WHERE department_id > 0 ORDER BY dname ASC")
              If Not depRes.EOF Then
                Do While Not depRes.EOF
                  Response.Write "<option value=""" & depRes("department_id") & """>" & FixTildes(depRes("dname")) & "</option>" & vbNewLine
                  depRes.MoveNext
                Loop
                depRes.Close
              End If
            %>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label"><b><%=lang(cnnDB, "Status")%>:</b></label>
          <select name="status" class="form-control">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <option value="-1"><%=lang(cnnDB, "OpenProblems")%></option>
            <%
              Dim statRes
              Set statRes = SQLQuery(cnnDB, "SELECT status_id, sname From status WHERE status_id > 0 ORDER BY status_id ASC")
              If Not statRes.EOF Then
                Do While Not statRes.EOF
                  Response.Write "<option value=""" & statRes("status_id") & """>" & FixTildes(statRes("sname")) & "</option>" & vbNewLine
                  statRes.MoveNext
                Loop
                statRes.Close
              End If
            %>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label"><b><%=lang(cnnDB, "Priority")%>:</b></label>
          <select name="priority" class="form-control">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <%
              Dim priRes
              Set priRes = SQLQuery(cnnDB, "SELECT priority_id, pname From priority WHERE priority_id > 0 ORDER BY priority_id ASC")
              If Not priRes.EOF Then
                Do While Not priRes.EOF
                  Response.Write "<option value=""" & priRes("priority_id") & """>" & FixTildes(priRes("pname")) & "</option>" & vbNewLine
                  priRes.MoveNext
                Loop
                priRes.Close
              End If
            %>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label"><b><%=lang(cnnDB, "SortBy")%>:</b></label>
          <select name="order" class="form-control">
            <option value="0"><%=lang(cnnDB, "ProblemID")%></option>
            <option value="1"><%=lang(cnnDB, "Priority")%></option>
            <option value="2"><%=lang(cnnDB, "Status")%></option>
            <option value="3"><%=lang(cnnDB, "Department")%></option>
            <option value="4"><%=lang(cnnDB, "Category")%></option>
          </select>
        </div>

      </div>

      <!-- Rango de Fechas -->
      <div style="margin-top:20px;padding-top:16px;border-top:1px solid #e2e8f0">
        <label class="form-label"><b><%=lang(cnnDB, "SearchBetweenDates")%>:</b></label>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
          <div>
            <span style="font-size:12px;color:#718096">Desde (D&iacute;a / Mes / A&ntilde;o):</span>
            <div style="display:flex;gap:6px;margin-top:4px">
              <select name="s_day" class="form-control" style="width:70px">
                <% Dim count: For count = 1 to 31: If count = 1 Then Response.Write "<option selected>" & count & "</option>" Else Response.Write "<option>" & count & "</option>" End If: Next %>
              </select>
              <select name="s_month" class="form-control" style="width:70px">
                <% For count = 1 to 12: If count = Month(now) Then Response.Write "<option selected>" & count & "</option>" Else Response.Write "<option>" & count & "</option>" End If: Next %>
              </select>
              <select name="s_year" class="form-control" style="width:90px">
                <% For count = 2000 to Year(now): If count = Year(now) Then Response.Write "<option selected>" & count & "</option>" Else Response.Write "<option>" & count & "</option>" End If: Next %>
              </select>
            </div>
          </div>
          <div>
            <span style="font-size:12px;color:#718096">Hasta (D&iacute;a / Mes / A&ntilde;o):</span>
            <div style="display:flex;gap:6px;margin-top:4px">
              <select name="e_day" class="form-control" style="width:70px">
                <% For count = 1 to 31: If count = Day(now) Then Response.Write "<option selected>" & count & "</option>" Else Response.Write "<option>" & count & "</option>" End If: Next %>
              </select>
              <select name="e_month" class="form-control" style="width:70px">
                <% For count = 1 to 12: If count = Month(now) Then Response.Write "<option selected>" & count & "</option>" Else Response.Write "<option>" & count & "</option>" End If: Next %>
              </select>
              <select name="e_year" class="form-control" style="width:90px">
                <% For count = 2000 to Year(now): If count = Year(now) Then Response.Write "<option selected>" & count & "</option>" Else Response.Write "<option>" & count & "</option>" End If: Next %>
              </select>
            </div>
          </div>
        </div>
      </div>

      <div style="text-align:center;margin-top:24px">
        <button type="submit" class="btn btn-primary" style="padding:10px 32px;font-size:14px">
          &#128269; <%=lang(cnnDB, "Search")%>
        </button>
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
