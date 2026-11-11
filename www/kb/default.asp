<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% 
  Option Explicit
  Response.Buffer = TRUE
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: default.asp (kb)
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Knowledge Base Search with modern UI & highlight.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "KnowledgeBase")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check permissions
      Call CheckKB(cnnDB, sid)
      Call DisplayHeader(cnnDB, sid)

      Dim blnSearch, keywords
      blnSearch = False
      keywords = Trim(CleanInputText(Request.Form("keywords")))
      If Len(keywords) = 0 Then
        keywords = Trim(Request.QueryString("keywords"))
      End If

      If (Request.Form("search") = "1" Or Len(keywords) > 0) Then
        blnSearch = True
      End If

      Dim queryStr, strWhere1, strWhere2, srchRes, varWordList, blAllOff, kwSql
      ' Version de las palabras clave segura para SQL (comillas duplicadas, sin ; ni comentarios)
      kwSql = Replace(keywords, "'", "''")
      kwSql = Replace(kwSql, ";", " ")
      kwSql = Replace(kwSql, "--", " ")

      If blnSearch And Len(keywords) > 0 Then
        varWordList = Split(kwSql, " ")

        queryStr = "SELECT id, title, description, solution, start_date, close_date FROM problems"
        strWhere1 = "WHERE (kb=1 AND status=" & Cfg(cnnDB, "CloseStatus") & ")"

        If Request.Form("title") <> "on" And Request.Form("description") <> "on" And Request.Form("solution") <> "on" Then
          blAllOff = True
        Else
          blAllOff = False
        End If

        If Cfg(cnnDB, "KBFreeText") = 1 Then
          If Request.Form("title") = "on" Or blAllOff Then
            If Len(strWhere2) < 1 Then
              strWhere2 = " AND (FREETEXT(title, '" & kwSql & "')"
            Else
              strWhere2 = strWhere2 & " OR FREETEXT(title, '" & kwSql & "')"
            End If
          End If
          If Request.Form("description") = "on" Or blAllOff Then
            If Len(strWhere2) < 1 Then
              strWhere2 = " AND (FREETEXT(description, '" & kwSql & "')"
            Else
              strWhere2 = strWhere2 & " OR FREETEXT(description, '" & kwSql & "')"
            End If
          End If
          If Request.Form("solution") = "on" Or blAllOff Then
            If Len(strWhere2) < 1 Then
              strWhere2 = " AND (FREETEXT(solution, '" & kwSql & "')"
            Else
              strWhere2 = strWhere2 & " OR FREETEXT(solution, '" & kwSql & "')"
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

        queryStr = queryStr & " " & strWhere1 & strWhere2 & " ORDER BY start_date DESC"
        Set srchRes = SQLQuery(cnnDB, queryStr)
      End If
    %>

    <div class="lhd-wrapper">

      <!-- Hero Search Banner -->
      <div class="lhd-search-hero">
        <h2>&#128218; <%=lang(cnnDB, "KnowledgeBase")%></h2>
        <p>&iquest;C&oacute;mo podemos ayudarte hoy? Busca soluciones a problemas comunes</p>

        <form method="post" action="default.asp" class="lhd-search-bar">
          <input type="hidden" name="search" value="1">
          <input type="text"
                 name="keywords"
                 value="<% = Server.HTMLEncode(keywords) %>"
                 placeholder="Escribe palabras clave (ej: contrase&ntilde;a, vpn, impresora)..."
                 autofocus>
          <button type="submit">&#128269; <%=lang(cnnDB, "Search")%></button>
        </form>

        <div style="margin-top:14px;font-size:12px;opacity:.9;display:flex;justify-content:center;gap:16px;flex-wrap:wrap">
          <label style="cursor:pointer;display:flex;align-items:center;gap:4px">
            <input type="checkbox" name="title" checked> <%=lang(cnnDB, "Title")%>
          </label>
          <label style="cursor:pointer;display:flex;align-items:center;gap:4px">
            <input type="checkbox" name="description" checked> <%=lang(cnnDB, "Description")%>
          </label>
          <label style="cursor:pointer;display:flex;align-items:center;gap:4px">
            <input type="checkbox" name="solution" checked> <%=lang(cnnDB, "Solution")%>
          </label>
        </div>
      </div>

      <!-- Buscar por ID directo -->
      <div class="lhd-card" style="padding:14px 20px;margin-bottom:20px">
        <form method="POST" action="details.asp" style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
          <span style="font-size:13px;font-weight:600;color:var(--color-text-secondary)">&#128203; <%=lang(cnnDB, "LookupbyID")%>:</span>
          <input type="text" size="8" name="id" class="form-control" placeholder="ID..." style="width:110px">
          <button type="submit" class="btn btn-secondary btn-sm"><%=lang(cnnDB, "LookupbyID")%></button>
        </form>
      </div>

      <!-- Resultados de busqueda -->
      <% If blnSearch Then %>
        <% If srchRes.EOF Then %>
          <div class="lhd-card" style="text-align:center;padding:40px">
            <div style="font-size:42px;margin-bottom:12px">&#128269;</div>
            <h3><%=lang(cnnDB, "Noresultsfound")%></h3>
            <p style="color:var(--color-muted);font-size:13px">
              No encontramos art&iacute;culos que coincidan con "<strong><% = Server.HTMLEncode(keywords) %></strong>".
            </p>
            <p style="margin-top:16px">
              <a class="btn btn-primary btn-sm" href="../user/new.asp">&#10133; Crear una nueva solicitud</a>
            </p>
          </div>
        <% Else %>
          <div class="lhd-card" style="padding:0;overflow:hidden">
            <div style="padding:14px 20px;background:var(--color-surface);border-bottom:1px solid var(--color-border-light)">
              <strong>Art&iacute;culos de soluci&oacute;n encontrados</strong>
            </div>
            <div style="overflow-x:auto">
              <table class="lhd-table">
                <thead>
                  <tr>
                    <th class="col-id">ID</th>
                    <th><%=lang(cnnDB, "Title")%></th>
                    <th>Soluci&oacute;n previa</th>
                    <th class="col-date">Fecha resoluci&oacute;n</th>
                  </tr>
                </thead>
                <tbody>
                  <% Do While Not srchRes.EOF %>
                    <tr>
                      <td class="col-id">
                        <a href="details.asp?id=<% = srchRes("id") %>">#<% = srchRes("id") %></a>
                      </td>
                      <td>
                        <a href="details.asp?id=<% = srchRes("id") %>" style="font-weight:600">
                          <% = HighlightTerm(srchRes("title"), keywords) %>
                        </a>
                      </td>
                      <td style="color:var(--color-text-secondary);font-size:12px;max-width:320px">
                        <% 
                          Dim solSnippet
                          solSnippet = srchRes("solution")
                          If Len(solSnippet) > 120 Then solSnippet = Left(solSnippet, 120) & "..."
                          Response.Write HighlightTerm(solSnippet, keywords)
                        %>
                      </td>
                      <td class="col-date">
                        <% = DisplayDate(srchRes("close_date"), lhdDateOnly) %>
                      </td>
                    </tr>
                  <%
                      srchRes.MoveNext
                    Loop
                  %>
                </tbody>
              </table>
            </div>
          </div>
        <% End If %>
      <% End If %>

    </div>

    <%
      If blnSearch Then srchRes.Close
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
