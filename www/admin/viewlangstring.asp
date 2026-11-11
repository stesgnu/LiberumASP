<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
  Server.ScriptTimeOut = 600  ' Wait 10 minute to time out script
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: viewlangstring.asp
  Date:     $Date: 2002/01/23 02:10:00 $
  Version:  $Revision: 1.51.2.1 $
  Purpose:  Manage the text strings for a selected language (with search and pagination, 10 per page)
  -->

  <!--  #include file = "../settings.asp" -->
  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "LanguageStrings")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .ls-toolbar { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between; padding: 14px 16px; border-bottom: 1px solid var(--color-border-light, #E2E8F0); }
      .ls-search-form { display: flex; gap: 8px; flex: 1 1 320px; max-width: 480px; margin: 0; }
      .ls-search { flex: 1; padding: 10px 14px; border: 1px solid var(--color-border, #CBD5E0); border-radius: 10px; font-size: 14px; box-sizing: border-box; min-width: 0; }
      .ls-search:focus { outline: none; border-color: var(--color-primary, #0057B8); box-shadow: 0 0 0 3px var(--color-primary-light, #E8F0FB); }
      .ls-count { font-size: 13px; color: #64748b; }
      .ls-pager { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; justify-content: space-between; padding: 14px 16px; border-top: 1px solid var(--color-border-light, #E2E8F0); }
      .ls-pages { display: flex; gap: 4px; flex-wrap: wrap; }
      .ls-pg { min-width: 34px; height: 34px; padding: 0 10px; display: inline-flex; align-items: center; justify-content: center; border: 1px solid var(--color-border, #CBD5E0); border-radius: 8px; background: #fff; color: #334155; font-size: 13px; font-weight: 600; text-decoration: none !important; }
      a.ls-pg:hover { border-color: var(--color-primary, #0057B8); color: var(--color-primary, #0057B8); }
      .ls-pg.on { background: var(--color-primary, #0057B8); border-color: var(--color-primary, #0057B8); color: #fff; }
      .ls-pg.off { color: #A0AEC0; background: #F7FAFC; cursor: default; }
      .ls-empty { padding: 34px 16px; text-align: center; color: #64748b; }
      .ls-var { font-family: Consolas, Menlo, monospace; font-size: 13px; word-break: break-all; }
      .ls-base { color: #475569; max-width: 360px; }
      .ls-savebar { display: flex; justify-content: flex-end; padding: 14px 16px; border-top: 1px solid var(--color-border-light, #E2E8F0); }
      .ls-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }
      @media (max-width: 720px) { .ls-grid { grid-template-columns: 1fr; } }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim intLangID, intCompLangID
      intLangID = request.querystring("lang_id")
      if len(intLangID) = 0 then
        Call DisplayError(3, lang(cnnDB, "NolanguageIDgiven"))
      end if
      intLangID = SafeCLng(intLangID)
      intCompLangID = 1 ' English

      Dim blnSaveSuccess, blnAddSuccess
      blnSaveSuccess = False
      blnAddSuccess = False

      If Request.Form("frm_save") = "1" Then
        ' Se guardan solo los campos enviados (la lista esta paginada): cada campo se llama como la variable
        Dim frmKey, rstVarOk, rstVarIsNull, rstUpdateLang, strUpdLangName, strVarName
        For Each frmKey In Request.Form
          If frmKey <> "frm_save" Then
            strVarName = Replace(frmKey, "'", "''")
            Set rstVarOk = SQLQuery(cnnDB, "SELECT id FROM tblLangStrings WHERE id=" & intCompLangID & " AND variable='" & strVarName & "'")
            If Not rstVarOk.EOF Then
              strUpdLangName = Request.Form(frmKey)
              strUpdLangName = Replace(strUpdLangName, "'", "''")
              Set rstVarIsNull = SQLQuery(cnnDB, "SELECT id FROM tblLangStrings WHERE id=" & intLangID & " AND variable='" & strVarName & "'")
              If rstVarISNull.EOF Then
                Set rstUpdateLang = SQLQuery(cnnDB, "INSERT INTO tblLangStrings (id, variable, LangText) VALUES " & _
                  "(" & intLangID & ", '" & strVarName & "', '" & strUpdLangName & "')")
              Else
                Set rstUpdateLang = SQLQuery(cnnDB, "UPDATE tblLangStrings SET " & _
                  "id = " & intLangID & ", " & _
                  "variable = '" & strVarName & "', " & _
                  "LangText = '" & strUpdLangName & "' " & _
                  "WHERE id=" & intLangID & " AND variable='" & strVarName & "'")
              End If
              rstVarIsNull.Close
            End If
            rstVarOk.Close
          End If
        Next
        blnSaveSuccess = True

        ' Remove cached language strings
        Call ClearLangCache(cnnDB)
        Call LhdToast("success", lang(cnnDB, "ChangesSaved"))
      End If

      If Request.Form("frm_add") = "1" Then
        Dim intLang1ID, intLang2ID, strNew1, strNew2, strNewVarName
        Dim rstCheckVarName, rstInsert1, rstInsert2
        strNewVarName = CleanInputText(Request.Form("varname"))
        Set rstCheckVarName = SQLQuery(cnnDB, "SELECT id FROM tblLangStrings WHERE variable='" & strNewVarName & "'")
        If Not rstCheckVarName.EOF Then
          rstCheckVarName.Close
          Call DisplayError(3, lang(cnnDB, "Variablenameisalreadyinuse"))
        End If
        intLang1ID = Cint(Request.Form("string1_id"))
        intLang2ID = Cint(Request.Form("string2_id"))
        strNew1 = CleanInputText(Request.Form("string1_value"))
        strNew2 = CleanInputText(Request.Form("string2_value"))

        strNewVarName = replace(strNewVarName, "'", "''")
        strNew1 = replace(strNew1, "'", "''")
        strNew2 = replace(strNew2, "'", "''")

        Set rstInsert1 = SQLQuery(cnnDB, "INSERT INTO tblLangStrings (id, variable, LangText) VALUES " & _
          "(" & intLang1ID & ", '" & strNewVarName & "', '" & strNew1 & "')")
        If intLang1ID <> intLang2ID Then
          Set rstInsert1 = SQLQuery(cnnDB, "INSERT INTO tblLangStrings (id, variable, LangText) VALUES " & _
            "(" & intLang2ID & ", '" & strNewVarName & "', '" & strNew2 & "')")
        End If

        blnAddSuccess = True

        ' Remove cached language strings
        Call ClearLangCache(cnnDB)
        Call LhdToast("success", lang(cnnDB, "StringAdded"))
      End If

      ' Todas las validaciones (y posibles DisplayError) terminaron: ahora si se dibuja la cabecera
      Call DisplayHeader(cnnDB, sid)

      ' ---- Parametros: busqueda y pagina (10 registros por pagina) ----
      Const PER_PAGE = 10
      Dim q, qSql, pageNo, totalRows, totalPages, skipRows, whereSql
      q = Left(Trim(Request.QueryString("q")), 50)
      qSql = Replace(q, "'", "''")
      qSql = Replace(qSql, "[", "")
      qSql = Replace(qSql, "]", "")
      qSql = Replace(qSql, "%", "")
      whereSql = "id=" & intCompLangID
      If Len(qSql) > 0 Then whereSql = whereSql & " AND (variable LIKE '%" & qSql & "%' OR LangText LIKE '%" & qSql & "%')"

      Dim rstCnt
      Set rstCnt = SQLQuery(cnnDB, "SELECT COUNT(*) AS total FROM tblLangStrings WHERE " & whereSql)
      totalRows = SafeCLng(rstCnt("total"))
      rstCnt.Close

      totalPages = Int((totalRows + PER_PAGE - 1) / PER_PAGE)
      If totalPages < 1 Then totalPages = 1
      pageNo = SafeCLng(Request.QueryString("page"))
      If pageNo < 1 Then pageNo = 1
      If pageNo > totalPages Then pageNo = totalPages
      skipRows = (pageNo - 1) * PER_PAGE

      Dim rstLangString, rstLangNameDefault, rstLangNameCurrent
      Set rstLangString = SQLQuery(cnnDB, "SELECT TOP " & (skipRows + PER_PAGE) & " variable, LangText FROM tblLangStrings WHERE " & whereSql & " ORDER BY variable ASC")
      Set rstLangNameDefault = SQLQuery(cnnDB, "SELECT * FROM tblLanguage WHERE id=" & intCompLangID)
      Set rstLangNameCurrent = SQLQuery(cnnDB, "SELECT * FROM tblLanguage WHERE id=" & intLangID)

      Dim skipI
      skipI = 0
      Do While Not rstLangString.EOF And skipI < skipRows
        rstLangString.MoveNext
        skipI = skipI + 1
      Loop

      Dim firstShown, lastShown, qParam, baseUrl, pgI, pgFrom, pgTo
      qParam = ""
      If Len(q) > 0 Then qParam = "&amp;q=" & Server.URLEncode(q)
      baseUrl = "viewlangstring.asp?lang_id=" & intLangID
      firstShown = 0
      If totalRows > 0 Then firstShown = skipRows + 1
      lastShown = skipRows + PER_PAGE
      If lastShown > totalRows Then lastShown = totalRows

      Function GetEditLangString(strEditVar)
        Dim rstEditString
        Set rstEditString = SQLQuery(cnnDB, "SELECT LangText FROM tblLangStrings WHERE id=" & intLangID & " AND variable='" & Replace(strEditVar & "", "'", "''") & "'")
        If rstEditString.EOF Then
          GetEditLangString=""
        Else
          GetEditLangString = Server.HTMLEncode(rstEditString("LangText") & "")
        End If
        rstEditString.Close
      End Function

    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128292;</span>
          <%=lang(cnnDB, "LanguageStrings")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="viewlang.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Languages")%></a>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>

      <div class="lhd-card lhd-card-flush">
        <div class="ls-toolbar">
          <form method="get" action="viewlangstring.asp" class="ls-search-form">
            <input type="hidden" name="lang_id" value="<% = intLangID %>">
            <input type="text" name="q" class="ls-search" value="<% = Server.HTMLEncode(q) %>" placeholder="Buscar por variable o texto..." maxlength="50" autocomplete="off">
            <button type="submit" class="btn btn-primary">&#128269; Buscar</button>
            <% If Len(q) > 0 Then %><a class="btn btn-secondary" href="<% = baseUrl %>">Limpiar</a><% End If %>
          </form>
          <span class="ls-count">
            <% If totalRows = 0 Then %>
              Sin resultados
            <% Else %>
              Mostrando <strong><% = firstShown %>&ndash;<% = lastShown %></strong> de <strong><% = totalRows %></strong>
            <% End If %>
          </span>
        </div>

        <form action="<% = baseUrl %>&amp;page=<% = pageNo %><% = qParam %>" method="POST" name="change" style="margin:0">
          <input type="hidden" name="frm_save" value="1">
          <div style="overflow-x:auto">
            <table class="lhd-table">
              <thead>
                <tr>
                  <th><%=lang(cnnDB, "Variable")%></th>
                  <th><% = Server.HTMLEncode(rstLangNameDefault("LangName") & "") %> (<% = Server.HTMLEncode(rstLangNameDefault("Localized") & "") %>)</th>
                  <th><% = Server.HTMLEncode(rstLangNameCurrent("LangName") & "") %> (<% = Server.HTMLEncode(rstLangNameCurrent("Localized") & "") %>)</th>
                </tr>
              </thead>
              <tbody>
              <%
                Dim shownRows
                shownRows = 0
                Do While Not rstLangString.EOF And shownRows < PER_PAGE
              %>
                <tr>
                  <td class="ls-var"><strong><% = Server.HTMLEncode(rstLangString("variable") & "") %></strong></td>
                  <td class="ls-base"><% = Server.HTMLEncode(rstLangString("LangText") & "") %></td>
                  <td>
                    <input type="text" class="form-control" style="min-width:220px" name="<% = Server.HTMLEncode(rstLangString("variable") & "") %>" value="<% = GetEditLangString(rstLangString("variable")) %>">
                  </td>
                </tr>
              <%
                  shownRows = shownRows + 1
                  rstLangString.MoveNext
                Loop
              %>
              </tbody>
            </table>
            <% If totalRows = 0 Then %>
              <div class="ls-empty">
                <% If Len(q) > 0 Then %>No se encontraron cadenas que coincidan con "<strong><% = Server.HTMLEncode(q) %></strong>".<% Else %>Todav&iacute;a no hay cadenas registradas.<% End If %>
              </div>
            <% End If %>
          </div>
          <% If totalRows > 0 Then %>
          <div class="ls-savebar">
            <button type="submit" class="btn btn-primary">&#128190; <%=lang(cnnDB, "Save")%></button>
          </div>
          <% End If %>
        </form>

        <% If totalPages > 1 Then
             pgFrom = pageNo - 2
             If pgFrom < 1 Then pgFrom = 1
             pgTo = pgFrom + 4
             If pgTo > totalPages Then
               pgTo = totalPages
               pgFrom = pgTo - 4
               If pgFrom < 1 Then pgFrom = 1
             End If %>
        <div class="ls-pager">
          <span class="ls-count">P&aacute;gina <strong><% = pageNo %></strong> de <strong><% = totalPages %></strong></span>
          <div class="ls-pages">
            <% If pageNo > 1 Then %>
              <a class="ls-pg" href="<% = baseUrl %>&amp;page=1<% = qParam %>" title="Primera">&laquo;</a>
              <a class="ls-pg" href="<% = baseUrl %>&amp;page=<% = pageNo - 1 %><% = qParam %>"><%=lang(cnnDB, "Previous")%></a>
            <% Else %>
              <span class="ls-pg off">&laquo;</span>
              <span class="ls-pg off"><%=lang(cnnDB, "Previous")%></span>
            <% End If %>
            <% For pgI = pgFrom To pgTo
                 If pgI = pageNo Then %>
              <span class="ls-pg on"><% = pgI %></span>
            <%   Else %>
              <a class="ls-pg" href="<% = baseUrl %>&amp;page=<% = pgI %><% = qParam %>"><% = pgI %></a>
            <%   End If
               Next %>
            <% If pageNo < totalPages Then %>
              <a class="ls-pg" href="<% = baseUrl %>&amp;page=<% = pageNo + 1 %><% = qParam %>"><%=lang(cnnDB, "Next")%></a>
              <a class="ls-pg" href="<% = baseUrl %>&amp;page=<% = totalPages %><% = qParam %>" title="&Uacute;ltima">&raquo;</a>
            <% Else %>
              <span class="ls-pg off"><%=lang(cnnDB, "Next")%></span>
              <span class="ls-pg off">&raquo;</span>
            <% End If %>
          </div>
        </div>
        <% End If %>
      </div>

      <div class="lhd-card">
        <div class="section-header">&#10133; <%=lang(cnnDB, "AddLanguageString")%></div>
        <form action="<% = baseUrl %>" method="POST" name="add" style="margin:0">
          <input type="hidden" name="frm_add" value="1">
          <input type="hidden" name="string1_id" value="<% = rstLangNameDefault("id") %>">
          <input type="hidden" name="string2_id" value="<% = rstLangNameCurrent("id") %>">
          <div class="ls-grid">
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "Variable")%></label>
              <input type="text" class="form-control" name="varname">
            </div>
            <div class="form-group">
              <label class="form-label"><% = Server.HTMLEncode(rstLangNameDefault("LangName") & "") %> (<% = Server.HTMLEncode(rstLangNameDefault("Localized") & "") %>)</label>
              <input type="text" class="form-control" name="string1_value">
            </div>
            <div class="form-group">
              <label class="form-label"><% = Server.HTMLEncode(rstLangNameCurrent("LangName") & "") %> (<% = Server.HTMLEncode(rstLangNameCurrent("Localized") & "") %>)</label>
              <input type="text" class="form-control" name="string2_value">
            </div>
          </div>
          <div style="display:flex;justify-content:flex-end">
            <button type="submit" class="btn btn-primary">&#10133; <%=lang(cnnDB, "AddNew")%></button>
          </div>
        </form>
      </div>
    </div>

    <%
      rstLangString.Close
      rstLangNameCurrent.Close
      rstLangNameDefault.Close
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
