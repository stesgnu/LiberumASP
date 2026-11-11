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

  Filename: config.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Allows modifications to the site's settings.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Configuration")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .cf-wrap { max-width: 960px; margin: 0 auto; }
      .cf-card { margin-bottom: 18px; }
      .cf-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 0 18px; }
      .cf-hint { font-size: 12px; color: #64748b; margin-top: 4px; }
      .cf-full { grid-column: 1 / -1; }
      .cf-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      ' Renders a Yes/No select
      Sub CfYesNo(nm, curVal)
        Response.Write "<select class=""form-control"" id=""" & nm & """ name=""" & nm & """>"
        If SafeCLng(curVal) = 0 Then
          Response.Write "<option value=""0"" selected>" & lang(cnnDB, "NO") & "</option><option value=""1"">" & lang(cnnDB, "YES") & "</option>"
        Else
          Response.Write "<option value=""0"">" & lang(cnnDB, "NO") & "</option><option value=""1"" selected>" & lang(cnnDB, "YES") & "</option>"
        End If
        Response.Write "</select>"
      End Sub

      ' Renders a select filled from a query
      Sub CfQuerySelect(nm, sqlText, idField, txtField, curVal, firstOpt)
        Dim rsOpt
        Response.Write "<select class=""form-control"" id=""" & nm & """ name=""" & nm & """>"
        If Len(firstOpt) > 0 Then
          If SafeCLng(curVal) = 0 Then
            Response.Write "<option value=""0"" selected>" & firstOpt & "</option>"
          Else
            Response.Write "<option value=""0"">" & firstOpt & "</option>"
          End If
        End If
        Set rsOpt = SQLQuery(cnnDB, sqlText)
        If Not rsOpt.EOF Then
          Do While Not rsOpt.EOF
            If SafeCLng(rsOpt(idField)) = SafeCLng(curVal) Then
              Response.Write "<option value=""" & Server.HTMLEncode(rsOpt(idField) & "") & """ selected>" & Server.HTMLEncode(rsOpt(txtField) & "") & "</option>"
            Else
              Response.Write "<option value=""" & Server.HTMLEncode(rsOpt(idField) & "") & """>" & Server.HTMLEncode(rsOpt(txtField) & "") & "</option>"
            End If
            rsOpt.MoveNext
          Loop
        End If
        rsOpt.Close
        Response.Write "</select>"
      End Sub

      ' Save Results
      If Request.Form("save") = "1" Then
        Dim strSQL, updRes
        Dim SiteName, BaseURL, HDName, HDReply, BaseEmail, EmailType, intNotifyUser
        Dim DefaultPriority, DefaultStatus, KBFreeText
        Dim CloseStatus, AuthType, DefaultLanguage
        Dim EnableKB, SMTPServer, EnablePager, UseSelectUser, UseInoutBoard
        Dim AllowImageUpload, MaxImageSize

        SiteName = Left(Trim(CleanInputText(Request.Form("sitename"))), 50)
        BaseURL = Left(Trim(CleanInputText(Request.Form("baseurl"))), 50)
        HDName = Left(Trim(CleanInputText(Request.Form("hdname"))), 50)
        HDReply = Left(Trim(CleanInputText(Request.Form("hdreply"))), 50)
        BaseEmail = Left(Trim(CleanInputText(Request.Form("baseemail"))), 50)
        EmailType = Cint(Request.Form("emailtype"))
        intNotifyUser = Cint(Request.Form("notifyuser"))
        DefaultPriority = Cint(Request.Form("defaultpriority"))
        DefaultStatus = Cint(Request.Form("defaultstatus"))
        CloseStatus = Cint(Request.Form("closestatus"))
        AuthType = Cint(Request.Form("authtype"))
        EnableKB = Cint(Request.Form("enablekb"))
        SMTPServer = Left(Trim(CleanInputText(Request.Form("smtpserver"))), 50)
        EnablePager = Cint(Request.Form("enablepager"))
        UseSelectUser = Cint(Request.Form("UseSelectUser"))
        UseInoutBoard = Cint(Request.Form("UseInoutBoard"))
        KBFreeText = Cint(Request.Form("KBFreeText"))
        DefaultLanguage = Cint(Request.Form("DefaultLanguage"))
        AllowImageUpload = Cint(Request.Form("AllowImageUpload"))
        MaxImageSize = Left(Trim(Request.Form("MaxImageSize")), 20)
        Dim AllowedAttachments, attachErr
        AllowedAttachments = LhdCleanExtList(Request.Form("AllowedAttachments"))
        If Len(AllowedAttachments) = 0 Then AllowedAttachments = LhdCleanExtList("jpg,jpeg,png,gif,webp,pdf,doc,docx,xls,xlsx,txt,zip")

        strSQL = "UPDATE tblConfig SET " & _
          "SiteName = '" & SiteName & "', " & _
          "BaseURL = '" & BaseURL & "', " & _
          "HDName = '" & HDName & "', " & _
          "HDReply = '" & HDReply & "', " & _
          "BaseEmail = '" & BaseEmail & "', " & _
          "NotifyUser = " & intNotifyUser & ", " & _
          "EmailType = " & EmailType & ", " & _
          "EnableKB = " & EnableKB & ", " & _
          "KBFreeText = " & KBFreeText & ", " & _
          "DefaultPriority = " & DefaultPriority & ", " & _
          "DefaultStatus = " & DefaultStatus & ", " & _
          "CloseStatus = " & CloseStatus & ", " & _
          "AuthType = " & AuthType & ", " & _
          "SMTPServer = '" & SMTPServer & "', " & _
          "UseSelectUser = " & UseSelectUser & ", " & _
          "UseInoutBoard = " & UseInoutBoard & ", " & _
          "DefaultLanguage = " & DefaultLanguage & ", " & _
          "AllowImageUpload = '" & AllowImageUpload & "', " & _
          "MaxImageSize = " & MaxImageSize & ", " & _
          "EnablePager = " & EnablePager

        Set updRes = SQLQuery(cnnDB, strSQL)
        attachErr = ""
        On Error Resume Next
        cnnDB.Execute "UPDATE tblConfig SET AllowedAttachments = '" & Replace(AllowedAttachments, "'", "''") & "'"
        If Err.Number <> 0 Then attachErr = Err.Description
        Err.Clear
        On Error Goto 0

      End If

      Dim rstConfig
      ' Get current configuration
      Set rstConfig = SQLQuery(cnnDB, "Select * From tblConfig")

      If rstConfig.EOF Then
        Call DisplayError(3, lang(cnnDB, "Unabletoreadconfigurationfromthedatabase"))
      End If

      Dim attachStored, attachReadErr, rstAT
      attachStored = ""
      attachReadErr = ""
      On Error Resume Next
      Set rstAT = cnnDB.Execute("SELECT TOP 1 AllowedAttachments FROM tblConfig")
      If Err.Number <> 0 Then
        attachReadErr = Err.Description
        Err.Clear
      ElseIf Not rstAT.EOF Then
        attachStored = rstAT(0) & ""
      End If
      On Error Goto 0

      If Request.Form("save") = "1" Then
        If attachErr <> "" Then
          Call LhdToast("error", "No se guardaron los tipos de adjuntos. Ejecute el instalador de adjuntos.")
        Else
          Call LhdToast("success", lang(cnnDB, "ConfigurationSaved"))
        End If
      End If
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#9881;</span>
          <%=lang(cnnDB, "Configuration")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <a class="btn btn-secondary" href="config_help.asp" target="_blank">&#8505; <%=lang(cnnDB, "Help")%></a>
        </div>
      </div>

      <div class="cf-wrap">

        <% If Request.Form("save") = "1" Then %>
        <div class="alert alert-success">
          <span class="alert-icon">&#10004;</span>
          <span><%=lang(cnnDB, "ConfigurationSaved")%></span>
        </div>
        <% End If %>
        <% If attachErr <> "" Then %>
        <div class="alert alert-error">
          <span class="alert-icon">&#9888;</span>
          <span>No se guardaron los tipos de adjuntos. Ejecute primero el instalador de adjuntos (admin/install_attachments.asp). Detalle: <% = Server.HTMLEncode(attachErr) %></span>
        </div>
        <% End If %>

        <form method="post" action="config.asp">
          <input type="hidden" name="save" value="1">

          <!-- General -->
          <div class="lhd-card cf-card">
            <div class="section-header">&#127970; General</div>
            <div class="cf-grid">
              <div class="form-group">
                <label class="form-label" for="sitename"><%=lang(cnnDB, "SiteName")%></label>
                <input type="text" class="form-control" id="sitename" name="sitename" maxlength="50" value="<% = Server.HTMLEncode(rstConfig("SiteName") & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="baseurl"><%=lang(cnnDB, "BaseURL")%></label>
                <input type="text" class="form-control" id="baseurl" name="baseurl" maxlength="50" value="<% = Server.HTMLEncode(rstConfig("BaseURL") & "") %>">
                <div class="cf-hint">(http://intranet/helpdesk)</div>
              </div>
              <div class="form-group">
                <label class="form-label" for="hdname"><%=lang(cnnDB, "AdministratorsName")%></label>
                <input type="text" class="form-control" id="hdname" name="hdname" maxlength="50" value="<% = Server.HTMLEncode(rstConfig("HDName") & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="hdreply"><%=lang(cnnDB, "AdministratorsEmail")%></label>
                <input type="text" class="form-control" id="hdreply" name="hdreply" maxlength="50" value="<% = Server.HTMLEncode(rstConfig("HDReply") & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="DefaultLanguage"><%=lang(cnnDB, "Defaultlanguage")%></label>
                <select class="form-control" id="DefaultLanguage" name="DefaultLanguage">
                <%
                  Dim optRes
                  Set optRes = SQLQuery(cnnDB, "SELECT * From tblLanguage")
                  If Not optRes.EOF Then
                    Do While Not optRes.EOF
                      If SafeCLng(optRes("id")) = SafeCLng(rstConfig("DefaultLanguage")) Then
                        Response.Write "<option value=""" & Server.HTMLEncode(optRes("id") & "") & """ selected>" & Server.HTMLEncode(optRes("LangName") & "") & " (" & Server.HTMLEncode(optRes("Localized") & "") & ")</option>"
                      Else
                        Response.Write "<option value=""" & Server.HTMLEncode(optRes("id") & "") & """>" & Server.HTMLEncode(optRes("LangName") & "") & " (" & Server.HTMLEncode(optRes("Localized") & "") & ")</option>"
                      End If
                      optRes.MoveNext
                    Loop
                  End If
                  optRes.Close
                %>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="authtype"><%=lang(cnnDB, "AuthenticationType")%></label>
                <% Call CfQuerySelect("authtype", "SELECT * From tblConfig_Auth", "id", "type", rstConfig("AuthType"), "") %>
              </div>
            </div>
          </div>

          <!-- Correo -->
          <div class="lhd-card cf-card">
            <div class="section-header">&#9993; Correo electr&oacute;nico</div>
            <div class="cf-grid">
              <div class="form-group">
                <label class="form-label" for="baseemail"><%=lang(cnnDB, "BaseEmail")%></label>
                <input type="text" class="form-control" id="baseemail" name="baseemail" maxlength="50" value="<% = Server.HTMLEncode(rstConfig("BaseEmail") & "") %>">
                <div class="cf-hint">(@company.com)</div>
              </div>
              <div class="form-group">
                <label class="form-label" for="emailtype"><%=lang(cnnDB, "EmailType")%></label>
                <% Call CfQuerySelect("emailtype", "SELECT * From tblConfig_Email", "id", "type", rstConfig("EmailType"), "") %>
              </div>
              <div class="form-group">
                <label class="form-label" for="smtpserver"><%=lang(cnnDB, "SMTPServer")%></label>
                <input type="text" class="form-control" id="smtpserver" name="smtpserver" maxlength="50" value="<% = Server.HTMLEncode(rstConfig("SMTPServer") & "") %>">
                <div class="cf-hint">(<%=lang(cnnDB, "JMailorASPEmail")%>)</div>
              </div>
              <div class="form-group">
                <label class="form-label" for="enablepager"><%=lang(cnnDB, "PagerPriorityLevel")%></label>
                <% Call CfQuerySelect("enablepager", "SELECT * FROM priority WHERE priority_id > 0", "priority_id", "pname", rstConfig("EnablePager"), lang(cnnDB, "Disabled")) %>
              </div>
              <div class="form-group">
                <label class="form-label" for="notifyuser"><%=lang(cnnDB, "EmailUseronUpdate")%></label>
                <% Call CfYesNo("notifyuser", rstConfig("NotifyUser")) %>
              </div>
            </div>
          </div>

          <!-- Tickets -->
          <div class="lhd-card cf-card">
            <div class="section-header">&#127915; Tickets</div>
            <div class="cf-grid">
              <div class="form-group">
                <label class="form-label" for="defaultpriority"><%=lang(cnnDB, "DefaultPriority")%></label>
                <% Call CfQuerySelect("defaultpriority", "SELECT * FROM priority WHERE priority_id > 0", "priority_id", "pname", rstConfig("DefaultPriority"), "") %>
              </div>
              <div class="form-group">
                <label class="form-label" for="defaultstatus"><%=lang(cnnDB, "DefaultStatus")%></label>
                <% Call CfQuerySelect("defaultstatus", "SELECT * FROM status WHERE status_id > 0 AND status_id <>" & SafeCLng(rstConfig("CloseStatus")) & " ORDER BY status_id ASC", "status_id", "sname", rstConfig("DefaultStatus"), "") %>
              </div>
              <div class="form-group">
                <label class="form-label" for="closestatus"><%=lang(cnnDB, "CloseStatus")%></label>
                <% Call CfQuerySelect("closestatus", "SELECT * FROM status WHERE status_id > 0 ORDER BY status_id ASC", "status_id", "sname", rstConfig("CloseStatus"), "") %>
              </div>
              <div class="form-group">
                <label class="form-label" for="UseSelectUser"><%=lang(cnnDB, "UseSelectUser")%></label>
                <% Call CfYesNo("UseSelectUser", rstConfig("useSelectUser")) %>
              </div>
              <div class="form-group">
                <label class="form-label" for="UseInoutBoard"><%=lang(cnnDB, "UseInOutBoard")%></label>
                <% Call CfYesNo("UseInoutBoard", rstConfig("useInoutBoard")) %>
              </div>
            </div>
          </div>

          <!-- Base de conocimiento -->
          <div class="lhd-card cf-card">
            <div class="section-header">&#128218; Base de conocimiento</div>
            <div class="cf-grid">
              <div class="form-group">
                <label class="form-label" for="enablekb"><%=lang(cnnDB, "EnableUserKB")%></label>
                <select class="form-control" id="enablekb" name="enablekb">
                  <option value="0"<% If SafeCLng(rstConfig("EnableKB")) = 0 Then Response.Write(" selected") %>><%=lang(cnnDB, "Disable")%></option>
                  <option value="1"<% If SafeCLng(rstConfig("EnableKB")) = 1 Then Response.Write(" selected") %>><%=lang(cnnDB, "RepsOnly")%></option>
                  <option value="2"<% If SafeCLng(rstConfig("EnableKB")) = 2 Then Response.Write(" selected") %>><%=lang(cnnDB, "UsersReps")%></option>
                  <option value="3"<% If SafeCLng(rstConfig("EnableKB")) = 3 Then Response.Write(" selected") %>><%=lang(cnnDB, "Anyone")%></option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="kbfreetext"><%=lang(cnnDB, "KBSQLFreeTextSearches")%></label>
                <select class="form-control" id="kbfreetext" name="kbfreetext">
                  <option value="0"<% If SafeCLng(rstConfig("KBFreeText")) = 0 Then Response.Write(" selected") %>><%=lang(cnnDB, "Disable")%></option>
                  <option value="1"<% If SafeCLng(rstConfig("KBFreeText")) <> 0 Then Response.Write(" selected") %>><%=lang(cnnDB, "Enable")%></option>
                </select>
              </div>
            </div>
          </div>

          <!-- Adjuntos -->
          <div class="lhd-card cf-card">
            <div class="section-header">&#128206; Archivos adjuntos</div>
            <div class="cf-grid">
              <div class="form-group">
                <label class="form-label" for="AllowImageUpload"><%=lang(cnnDB, "AllowImageUpload")%></label>
                <% Call CfYesNo("AllowImageUpload", rstConfig("AllowImageUpload")) %>
              </div>
              <div class="form-group">
                <label class="form-label" for="MaxImageSize"><%=lang(cnnDB, "MaxImageSize")%> (en bytes)</label>
                <input type="text" class="form-control" id="MaxImageSize" name="MaxImageSize" inputmode="numeric" pattern="[0-9]+" title="Solo n&uacute;mero de bytes" value="<% = Server.HTMLEncode(rstConfig("MaxImageSize") & "") %>">
                <div class="cf-hint">Tama&ntilde;o m&aacute;ximo por archivo adjunto. Equivalencias: 1 MB = 1048576 bytes, 5 MB = 5242880, 10 MB = 10485760.</div>
              </div>
              <div class="form-group cf-full">
                <label class="form-label" for="AllowedAttachments">Tipos de adjuntos permitidos</label>
                <input type="text" class="form-control" id="AllowedAttachments" name="AllowedAttachments" value="<% = Server.HTMLEncode(attachStored) %>">
                <div class="cf-hint">Separados por comas. Ej: pdf,doc,docx,xls,xlsx,png,jpg</div>
                <div class="cf-hint">Valor guardado: <b><% = Server.HTMLEncode(attachStored) %></b><% If attachReadErr <> "" Then %> &mdash; error al leer: <% = Server.HTMLEncode(attachReadErr) %><% End If %></div>
                <div class="cf-hint">No se aceptan ejecutables, scripts ni HTML (exe, js, html, svg...).</div>
              </div>
            </div>
          </div>

          <div class="cf-actions">
            <a class="btn btn-secondary" href="default.asp">Cancelar</a>
            <button type="submit" class="btn btn-primary">&#128190; <%=lang(cnnDB, "Save")%></button>
          </div>
        </form>
      </div>
    </div>
    <%
    	rstConfig.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
