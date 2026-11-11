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

  Filename: results.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.51.4.1 $
  Purpose:  This page display a list of search results posted from search.asp.

  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>
  <head>
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "SearchResults")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

  <%
    ' Check if user has permissions for this page
    Call CheckRep(cnnDB, sid)
	Call DisplayHeader(cnnDB, sid)

    ' Get the search fields from the form.
    Dim uid, id, rep, category, department, status, priority, order

    uid = Trim(CleanInputText(Request.Form("uid")))
    rep = SafeCInt(Request.Form("rep"))
    category = SafeCInt(Request.Form("category"))
    department = SafeCInt(Request.Form("department"))
    status = SafeCInt(Request.Form("status"))
    priority = SafeCInt(Request.Form("priority"))
    order = SafeCInt(Request.Form("order"))

    id = Trim(Request.Form("id"))
    If IsNumeric(id) Then
      id = Cint(id)
    Else
      id = 0
    End If

    ' Convert to valid characters
    uid = Replace(uid,"'","''")

    ' Create the query string.  Only get 100 results. String is in two
    ' parts, listStr (here) and whereStr (below).
    Dim listStr, whereStr
    listStr = "SELECT TOP 100 p.id, p.title, p.start_date, p.uid, p.uemail, r.uid AS ruid, s.sname " & _
    "FROM (problems AS p " & _
    "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
    "INNER JOIN status AS s ON p.status = s.status_id " & _
    "WHERE "

    ' Convert the start and end dates to valid ODBC dates
    Dim start_date, end_date
    If Len(Request.Form("start_date")) > 0 Then
      start_date = Request.Form("start_date")
    Else
      Dim intStartDay
      intStartDay = FixDay(Request.Form("s_month"), Request.Form("s_day"), Request.Form("s_year"))
      start_date = Request.Form("s_month") & "/" & intStartDay & "/" & Request.Form("s_year")
      start_date = start_date & " 00:00:00"
    End If

    If Len(Request.Form("end_date")) > 0 Then
      end_date = Request.Form("end_date")
    Else
      Dim intEndDay
      intEndDay = FixDay(Request.Form("e_month"), Request.Form("e_day"), Request.Form("e_year"))
      end_date = Request.Form("e_month") & "/" & intEndDay & "/" & Request.Form("e_year")
      end_date = end_date & " 23:59:59"
    End If

    ' If a field was filled in on the search form, add it to the where string.

    ' Insert the dates to search between.
    whereStr = "p.start_date >" & SQLDate(start_date, lhdAddSQLDelim) & " AND p.start_date<" & SQLDate(end_date, lhdAddSQLDelim)

    If Len(uid) > 0 Then
      whereStr = whereStr & " AND p.uid='" & uid & "'"
    End If

    If id <> 0 Then
      whereStr = whereStr & " AND p.id=" & id
    End If

    If rep <> 0 Then
      whereStr = whereStr & " AND p.rep=" & rep
    End If

    If category <> 0 Then
      whereStr = whereStr & " AND p.category=" & category
    End If

    If department <> 0 Then
      whereStr = whereStr & " AND p.department=" & department
    End If

    If status <> 0 Then
      If status > 0 Then
        whereStr = whereStr & " AND p.status=" & status
      Else
        whereStr = whereStr & " AND p.status < 100"
      End If
    End If

    If priority <> 0 Then
      whereStr = whereStr & " AND p.priority=" & priority
    End If

    ' Add the kewords to the list
    If Len(CleanInputText(Request.Form("keywords"))) > 0 Then
      Dim strWhere2, keywords, srchRes, varWordList, blAllOff
      keywords = Trim(CleanInputText(Request.Form("keywords")))
      varWordList = Split(keywords, " ")

      If Request.Form("title") <> "on" And Request.Form("description") <> "on" And Request.Form("solution") <> "on" Then
        blAllOff = True
      Else
        blAllOff = False
      End If

      If Cfg(cnnDB, "KBFreeText") = 1 Then
        If Request.Form("title") = "on" Or blAllOff Then
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (FREETEXT(title, '" & keywords & "')"
          Else
            strWhere2 = strWhere2 & " OR FREETEXT(title, '" & keywords & "')"
          End If
        End If
        If Request.Form("description") = "on" Or blAllOff Then
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (FREETEXT(description, '" & keywords & "')"
          Else
            strWhere2 = strWhere2 & " OR FREETEXT(description, '" & keywords & "')"
          End If
        End If
        If Request.Form("solution") = "on" Or blAllOff Then
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (FREETEXT(solution, '" & keywords & "')"
          Else
            strWhere2 = strWhere2 & " OR FREETEXT(solution, '" & keywords & "')"
          End If
        End If
        strWhere2 = strWhere2 & ")"
      Else
        If Request.Form("title") = "on" Or blAllOff Then
          Dim strTitleKW, strWhereTitle
          For Each strTitleKW in varWordList
            If Len(strWhereTitle) < 1 Then
              strWhereTitle = "("
            Else
              strWhereTitle = strWhereTitle & " AND "
            End If
            strWhereTitle = strWhereTitle & "title LIKE '%" & strTitleKW & "%'"
          Next
          strWhereTitle = strWhereTitle & ")"
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (" & strWhereTitle
          End IF
        End If

        If Request.Form("description") = "on" Or blAllOff Then
          Dim strDescKW, strWhereDesc
          For Each strDescKW in varWordList
            If Len(strWhereDesc) < 1 Then
              strWhereDesc = "("
            Else
              strWhereDesc = strWhereDesc & " AND "
            End If
            strWhereDesc = strWhereDesc & "description LIKE '%" & strDescKW & "%'"
          Next
          strWhereDesc = strWhereDesc & ")"
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (" & strWhereDesc
          Else
            strWhere2 = strWhere2 & " OR " & strWhereDesc
          End IF
        End If

        If Request.Form("solution") = "on" Or blAllOff Then
          Dim strSolKW, strWhereSol
          For Each strSolKW in varWordList
            If Len(strWhereSol) < 1 Then
              strWhereSol = "("
            Else
              strWhereSol = strWhereSol & " AND "
            End If
            strWhereSol = strWhereSol & "solution LIKE '%" & strSolKW & "%'"
          Next
          strWhereSol = strWhereSol & ")"
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (" & strWhereSol
          Else
            strWhere2 = strWhere2 & " OR " & strWhereSol
          End IF
        End If
        strWhere2 = strWhere2 & ")"
      End If
      whereStr = whereStr & strWhere2
    End If

    Select Case order
      Case 1 whereStr = whereStr & " ORDER BY p.id ASC"
      Case 2 whereStr = whereStr & " ORDER BY p.uid ASC"
      Case 3 whereStr = whereStr & " ORDER BY r.uid ASC"
      Case 4 whereStr = whereStr & " ORDER BY p.status ASC"
    End Select


    ' Concatenate the two strings together.
    listStr = listStr & whereStr

    ' Query the database
    Dim rstProbList, start
    Set rstProbList = SQLQuery(cnnDB, listStr)

    ' If results are retuned, display them.  Only 10 results
    ' per page.
    If Not rstProbList.EOF Then
      Dim Counter, numToDisplay, startNum
      Counter = 1
      If Len(Request.Form("num")) > 0 Then
        numToDisplay = SafeCInt(Request.Form("num"))
      Else
        numToDisplay = 25
      End if
      If Len(Request.Form("start")) > 0 Then
        start = SafeCInt(Request.Form("start"))
      Else
        start = 1
      End if
    %>
    <div class="lhd-wrapper">

  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
    <div class="lhd-page-title" style="margin:0">
      <span class="title-icon">&#128269;</span>
      <%=lang(cnnDB, "SearchResults")%>
    </div>
    <div style="display:flex;gap:8px">
      <a class="btn btn-secondary btn-sm" href="search.asp">&larr; Nueva B&uacute;squeda</a>
      <a class="btn btn-secondary btn-sm" href="view.asp">&#128203; Todos los Tickets</a>
    </div>
  </div>

  <div class="lhd-card">
    <div class="section-header">&#128203; Resultados encontrados (m&aacute;x. 100)</div>

    <% If listRes.EOF Then %>
      <div style="text-align:center;padding:48px 20px;color:#718096">
        <div style="font-size:40px;margin-bottom:12px">&#128269;</div>
        <div style="font-size:16px;font-weight:600"><%=lang(cnnDB, "Noresultsfound")%></div>
        <div style="font-size:13px;margin-top:6px;color:#a0aec0">Intenta con otros criterios de b&uacute;squeda o un rango de fechas m&aacute;s amplio.</div>
        <div style="margin-top:16px">
          <a class="btn btn-primary btn-sm" href="search.asp">Volver a Buscar</a>
        </div>
      </div>
    <% Else %>
      <div style="overflow-x:auto">
        <table class="lhd-table">
          <thead>
            <tr>
              <th style="width:70px">ID</th>
              <th><%=lang(cnnDB, "Title")%></th>
              <th><%=lang(cnnDB, "UserName")%></th>
              <th><%=lang(cnnDB, "AssignedTo")%></th>
              <th>Estado</th>
              <th style="width:130px"><%=lang(cnnDB, "Date")%></th>
            </tr>
          </thead>
          <tbody>
            <% Do While Not listRes.EOF %>
              <tr>
                <td style="font-weight:700">
                  <a href="details.asp?id=<% = listRes("id") %>">#<% = listRes("id") %></a>
                </td>
                <td>
                  <a href="details.asp?id=<% = listRes("id") %>">
                    <% = FixTildes(Server.HTMLEncode(listRes("title"))) %>
                  </a>
                </td>
                <td><% = FixTildes(Server.HTMLEncode(listRes("uid"))) %></td>
                <td><% = Server.HTMLEncode(listRes("ruid")) %></td>
                <td><% = StatusBadge(FixTildes(listRes("sname"))) %></td>
                <td><% = DisplayDate(listRes("start_date"), lhdDateOnly) %></td>
              </tr>
            <% 
                listRes.MoveNext
               Loop 
            %>
          </tbody>
        </table>
      </div>
    <% End If %>
  </div>

</div>

<%	End If

    ' Close results
    rstProbList.Close

    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close
  %>
</body>
</html>


