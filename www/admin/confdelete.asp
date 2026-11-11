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

  Filename: confdelete.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Confirms the deletetion of a rep, category ,etc.  It looks for
  any open problem that use the item and gives a warning if
  any exist. If not, then displays a yes/no question to
  confirm the deletion.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ConfirmDeletion")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .cd-card { max-width: 560px; margin: 8px auto 0; text-align: center; padding: 34px 28px; }
      .cd-icon { width: 68px; height: 68px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 32px; margin-bottom: 14px; }
      .cd-icon.warn { background: #FEF3C7; color: #B45309; }
      .cd-icon.err { background: #FEE2E2; color: #B91C1C; }
      .cd-title { font-size: 22px; font-weight: 700; margin: 0 0 8px; color: #1E293B; }
      .cd-text { font-size: 15px; color: #475569; margin: 0 0 18px; line-height: 1.5; }
      .cd-actions { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; }
      .cd-list { text-align: left; margin: 0 auto 18px; max-width: 360px; padding: 0; list-style: none; }
      .cd-list li { padding: 8px 12px; border: 1px solid var(--color-border-light, #E2E8F0); border-radius: 8px; margin-bottom: 6px; }
      .cd-list a { font-weight: 600; text-decoration: none; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      ' get field/type and id number
      Dim mType, id, strLanguageName, intLangID

      intLangID = Request.QueryString("mLangID")
      strLanguageName = Request.QueryString("mLanguage")
      mType = Cint(Request.QueryString("mtype"))
      if mType = 7 then
        id = Request.QueryString("id")
      else
        id = Cint(Request.QueryString("id"))
      end if

      'Do not allow admins to delete english language
      If mType = 6 and id = 1 Then
        Call DisplayError(3, lang(cnnDB, "Sorryyoucannotdeletethislanguage"))
      End If

      'Only allow admins to delete language variables from default language
      If mType = 7 and id > 0 Then
        Call DisplayError(3, lang(cnnDB, "Youmustdeletevariablesfromenglishlanguage"))
      End If

      ' look for missing data
      If (mType < 1) OR (mType > 7)Then
        Call DisplayError(3, lang(cnnDB, "InvalidIDordatatype"))
      End If
      If mType < 7 Then
        If id = 0 Then
        Call DisplayError(3, lang(cnnDB, "InvalidIDordatatype"))
        End If
      End If

      Dim okToDelete, catExists, catRes, probRes
      okToDelete = TRUE
      catExists = FALSE

      ' look for OPEN problems using this
      ' data.  You cannot delete the rep, department, etc
      ' if an open problem is using it.
      Select Case mType
        Case 1	' rep - look for problems and categories
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
          Set catRes = SQLQuery(cnnDB, "SELECT category_id, cname FROM categories WHERE rep_id=" & id)
          If Not catRes.EOF Then
            catExists = TRUE
            okToDelete = FALSE
          End If

        Case 2	' category
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE category=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))

        Case 3	' department
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE department=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))

        Case 4	' priority
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE priority=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))

        Case 5	' status
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE status=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
          If id = Cfg(cnnDB, "CloseStatus") Then
            probRes.Close
            Call DisplayError(3, lang(cnnDB, "TheCLOSEDstatuscannotbedeleted"))
          End If

        Case 6  ' Language
        ' Does not effect problems

        Case 7  ' Language strings
        ' Does not effect problems

      End Select

      ' Problems exists, so can't delete.
      if mType < 6 then
        If Not probRes.EOF Then
          okToDelete = FALSE
        End If
      End If

      ' Validaciones terminadas (posibles DisplayError ya resueltos): se dibuja la cabecera
      Call DisplayHeader(cnnDB, sid)

    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128465;</span>
          <%=lang(cnnDB, "ConfirmDeletion")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <%
            ' Get the correct URL to return to
            Select Case mType
              Case 1 %>
                <a class="btn btn-secondary" href="viewrep.asp">&larr; <%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Support Reps")%></a>
          <%		Case 2 %>
                <a class="btn btn-secondary" href="viewcat.asp">&larr; <%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Categories")%></a>
          <%		Case 3 %>
                <a class="btn btn-secondary" href="viewdep.asp">&larr; <%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Departments")%></a>
          <%		Case 4 %>
                <a class="btn btn-secondary" href="viewpri.asp">&larr; <%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Priorities")%></a>
          <%		Case 5 %>
                <a class="btn btn-secondary" href="viewstatus.asp">&larr; <%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Statuses")%></a>
          <%		Case 6 %>
                <a class="btn btn-secondary" href="viewlang.asp">&larr; <%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Languages")%></a>
          <%		Case 7 %>
                <a class="btn btn-secondary" href="viewlangstring.asp?id=<% = Server.URLEncode(intLangID & "") %>">&larr; <%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "LanguageStrings")%></a>
          <%	End Select %>
        </div>
      </div>

      <div class="lhd-card cd-card">
        <% If okToDelete Then %>
          <div class="cd-icon warn">&#9888;</div>
          <h2 class="cd-title"><%=lang(cnnDB, "Areyousure")%></h2>
          <p class="cd-text">Esta acci&oacute;n eliminar&aacute; el elemento seleccionado de la base de datos y no se puede deshacer.</p>
          <div class="cd-actions">
            <a class="btn btn-secondary" href="<% = Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER") & "") %>"><%=lang(cnnDB, "NO")%></a>
            <a class="btn btn-primary" href="delete.asp?mtype=<% = mtype %>&amp;id=<% = Server.URLEncode(id & "") %>&amp;mLangID=<% = Server.URLEncode(intLangID & "") %>&amp;mLanguage=<% = Server.URLEncode(strLanguageName & "") %>"><%=lang(cnnDB, "YES")%></a>
          </div>
        <% Else  %>
          <div class="cd-icon err">&#9940;</div>
          <h2 class="cd-title"><%=lang(cnnDB, "UnableToDelete")%></h2>
          <%
          If catExists Then
          %>
            <p class="cd-text"><%=lang(cnnDB, "catExistsText")%></p>
            <ul class="cd-list">
              <% Do While Not catRes.EOF %>
                <li><a href="modify.asp?mtype=2&amp;id=<% = catRes("category_id") %>"><% = Server.HTMLEncode(catRes("cname") & "") %></a></li>
              <%   catRes.MoveNext
                 Loop
              %>
            </ul>
          <%
          End If
          If NOT probRes.EOF Then
          %>
            <p class="cd-text"><%=lang(cnnDB, "probResText")%></p>
            <ul class="cd-list">
              <% Do While Not probRes.EOF %>
                <li><a href="../rep/view.asp?id=<% = probRes("id") %>">#<% = Server.HTMLEncode(probRes("id") & "") %></a></li>
              <%   probRes.MoveNext
                 Loop
              %>
            </ul>
        <%
          End If
        %>
          <div class="cd-actions">
            <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          </div>
        <%
        End If
        %>
      </div>
    </div>

    <%
      If mType = 1 Then
        catRes.Close
      End If
      If mType < 6 then
        probRes.Close
      End If

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
