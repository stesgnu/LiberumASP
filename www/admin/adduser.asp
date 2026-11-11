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

  Filename: adduser.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Form to add new users.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "AddNewUsers")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .au-wrap { max-width: 860px; margin: 0 auto; }
      .au-card { margin-bottom: 18px; }
      .au-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 0 18px; }
      .au-hint { font-size: 12px; color: #64748b; margin-top: 4px; }
      .au-pass { position: relative; }
      .au-pass .form-control { padding-right: 44px; }
      .au-eye { position: absolute; right: 6px; top: 50%; transform: translateY(-50%); background: none; border: 0; cursor: pointer; color: #64748b; padding: 6px 8px; border-radius: 6px; }
      .au-eye:hover { color: var(--color-primary, #0057B8); background: #F1F5F9; }
      .au-switch { display: flex; align-items: center; gap: 12px; padding: 12px 14px; border: 1px solid var(--color-border-light, #E2E8F0); border-radius: 10px; background: #F8FAFC; cursor: pointer; margin-bottom: 14px; }
      .au-switch input { position: absolute; opacity: 0; pointer-events: none; }
      .au-track { position: relative; width: 42px; height: 22px; border-radius: 9999px; background: #CBD5E0; flex-shrink: 0; transition: background .2s ease; }
      .au-knob { position: absolute; top: 2px; left: 2px; width: 18px; height: 18px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,.3); transition: transform .2s ease; }
      .au-switch input:checked + .au-track { background: var(--color-primary, #0057B8); }
      .au-switch input:checked + .au-track .au-knob { transform: translateX(20px); }
      .au-switch input:focus-visible + .au-track { outline: 2px solid var(--color-primary, #0057B8); outline-offset: 2px; }
      .au-switch .t { font-weight: 600; color: #1a202c; font-size: 14px; }
      .au-switch .d { font-size: 12px; color: #64748b; }
      .au-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
      .au-ok { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
    </style>
  </head>
  <body>

    <%
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      ' ---- Valores del formulario (se conservan si hay errores de validacion) ----
      Dim success, errList, createdUid
      Dim uid, email, fname, phone, location, department, pager
      Dim firstname, lastname, ListOnInoutboard, phone_home, phone_mobile
      Dim jobfunction, userresume, statuscode, statustext, statusdate
      Dim usrLanguage, IsRep, RepAccess, intNewSid, InoutAdmin, newpassword
      Dim optPager, optInout
      success = False
      errList = ""
      createdUid = ""
      optPager = (SafeCInt(Cfg(cnnDB, "EnablePager")) > 0)
      optInout = (SafeCInt(Cfg(cnnDB, "useinoutboard")) = 1)

      If Request.Form("save") = "1" Then
        uid = Left(Lcase(Trim(CleanInputText(Request.Form("uid")))), 50)
        email = Left(Lcase(Trim(CleanInputText(Request.Form("email")))), 50)
        pager = Left(Lcase(Trim(CleanInputText(Request.Form("pager")))), 50)
        phone = Left(Trim(CleanInputText(Request.Form("phone"))), 50)
        location = Left(Trim(CleanInputText(Request.Form("location"))), 50)
        department = SafeCLng(Request.Form("department"))
        ListOnInoutBoard = SafeCLng(Request.Form("ListOnInoutBoard"))
        firstname = Left(Trim(CleanInputText(Request.Form("firstname"))), 25)
        lastname = Left(Trim(CleanInputText(Request.Form("lastname"))), 24)
        phone_home = Left(Trim(CleanInputText(Request.Form("phone_home"))), 50)
        phone_mobile = Left(Trim(CleanInputText(Request.Form("phone_mobile"))), 50)
        jobfunction = Request.Form("jobfunction")
        userresume = Request.Form("userresume")
        statuscode = SafeCLng(Request.Form("statuscode"))
        statustext = Request.Form("statustext")
        usrLanguage = SafeCLng(Request.Form("usrLanguage"))
        RepAccess = SafeCLng(Request.Form("repaccess"))
        newpassword = Left(Trim(Request.Form("newpassword")), 50)
        If Request.Form("InoutAdmin") = "on" Then InoutAdmin = 1 Else InoutAdmin = 0
        If Request.Form("isrep") = "on" Then IsRep = 1 Else IsRep = 0

        ' ---- Validacion ----
        If Len(uid) = 0 Then
          errList = errList & lang(cnnDB, "Username") & " " & lang(cnnDB, "isarequiredfield") & ". "
        ElseIf InStr(uid, "'") > 0 Or InStr(uid, """") > 0 Or InStr(uid, " ") > 0 Then
          errList = errList & lang(cnnDB, "Username") & " " & lang(cnnDB, "containsinvalidcharacters") & ". "
        Else
          Dim rstDup
          Set rstDup = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & Replace(uid, "'", "''") & "'")
          If Not rstDup.EOF Then errList = errList & "Ya existe un usuario con ese nombre de usuario. "
          rstDup.Close
        End If
        If Len(firstname) = 0 Then errList = errList & lang(cnnDB, "FirstName") & " " & lang(cnnDB, "isarequiredfield") & ". "
        If Len(lastname) = 0 Then errList = errList & lang(cnnDB, "LastName") & " " & lang(cnnDB, "isarequiredfield") & ". "
        If Len(email) = 0 Then
          errList = errList & lang(cnnDB, "Emailaddress") & " " & lang(cnnDB, "isarequiredfield") & ". "
        ElseIf InStr(email, "@") = 0 Then
          errList = errList & "El correo electr" & ChrW(243) & "nico no es v" & ChrW(225) & "lido. "
        End If
        If Len(newpassword) = 0 Then errList = errList & lang(cnnDB, "NewPassword") & " " & lang(cnnDB, "isarequiredfield") & ". "

        If Len(errList) = 0 Then
          fname = firstname & " " & lastname
          intNewSid = GetUnique(cnnDB, "users")
          statusdate = SQLDate(Now, lhdAddSQLDelim)

          Dim sqlString, updRes
          sqlString = "INSERT INTO tblUsers (sid, uid, email1, email2, fname, firstname, lastname, phone, " & _
            "phone_home, phone_mobile, location1, department, InoutAdmin, IsRep, RepAccess, statuscode, " & _
            "statustext, statusdate, jobfunction, userresume, ListOnInoutboard, [language], [password])" & _
            " VALUES (" & intNewSid & ", '" & uid & "', '" & Replace(email, "'", "''") & "', '" & Replace(pager, "'", "''") & "', '" & _
            Replace(fname, "'", "''") & "', '" & Replace(firstname, "'", "''") & "', '" & Replace(lastname, "'", "''") & "', '" & _
            Replace(phone, "'", "''") & "', '" & Replace(phone_home, "'", "''") & "', '" & Replace(phone_mobile, "'", "''") & "', '" & _
            Replace(location, "'", "''") & "', " & department & ", " & InoutAdmin & ", " & IsRep & ", " & _
            RepAccess & ", " & statuscode & ", '" & Replace(statustext & "", "'", "''") & "', " & statusdate & ", '" & _
            Replace(jobfunction & "", "'", "''") & "', '" & Replace(userresume & "", "'", "''") & "', " & ListOnInoutboard & ", " & usrLanguage & ", '" & _
            Replace(newpassword, "'", "''") & "')"

          Set updRes = SQLQuery(cnnDB, sqlString)
          success = True
          createdUid = uid
          Call LhdToast("success", lang(cnnDB, "AccountCreated") & ": " & uid)
        Else
          Call LhdToast("error", errList)
        End If
      Else
        ' Primera carga: valores por omision
        RepAccess = 0
        statuscode = 0
        ListOnInoutboard = 1
        department = 0
        usrLanguage = 0
      End If

      ' Tras crear la cuenta el formulario vuelve a quedar vacio
      If success Then
        uid = "" : email = "" : pager = "" : phone = "" : location = "" : department = 0
        firstname = "" : lastname = "" : phone_home = "" : phone_mobile = ""
        jobfunction = "" : userresume = "" : statuscode = 0 : statustext = ""
        usrLanguage = 0 : RepAccess = 0 : newpassword = "" : IsRep = 0 : InoutAdmin = 0 : ListOnInoutboard = 1
      End If
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#10133;</span>
          <%=lang(cnnDB, "AddNewUsers")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="viewusers.asp">&larr; <%=lang(cnnDB, "ManageUsers")%></a>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>

      <div class="au-wrap">

        <% If success Then %>
        <div class="alert alert-success au-ok">
          <span><span class="alert-icon">&#10004;</span> <%=lang(cnnDB, "AccountCreated")%>: <strong><% = Server.HTMLEncode(createdUid) %></strong></span>
          <a class="btn btn-primary btn-sm" href="viewusers.asp?q=<% = Server.URLEncode(createdUid) %>">Ver en la lista</a>
        </div>
        <% ElseIf Len(errList) > 0 Then %>
        <div class="alert alert-error">
          <span class="alert-icon">&#9888;</span>
          <span><% = Server.HTMLEncode(errList) %></span>
        </div>
        <% End If %>

        <form name="upduser" id="frmAddUser" action="adduser.asp" method="POST" novalidate>
          <input type="hidden" name="save" value="1">

          <!-- Cuenta -->
          <div class="lhd-card au-card">
            <div class="section-header">&#128100; Datos de la cuenta <span style="float:right;font-weight:400;font-size:12px;color:#64748b"><em>*</em> = <%=lang(cnnDB, "Required")%></span></div>
            <div class="au-grid">
              <div class="form-group">
                <label class="form-label" for="uid"><%=lang(cnnDB, "Username")%> <span class="form-required">*</span></label>
                <input type="text" class="form-control" id="uid" name="uid" maxlength="50" value="<% = Server.HTMLEncode(uid & "") %>" autocomplete="off" autocapitalize="none">
                <div class="au-hint">Sin espacios. Se guarda en min&uacute;sculas.</div>
              </div>
              <div class="form-group">
                <label class="form-label" for="newpassword"><%=lang(cnnDB, "NewPassword")%> <span class="form-required">*</span></label>
                <div class="au-pass">
                  <input type="password" class="form-control" id="newpassword" name="newpassword" maxlength="50" autocomplete="new-password">
                  <button type="button" class="au-eye" id="btnEye" title="Mostrar u ocultar" aria-label="Mostrar u ocultar la contrase&ntilde;a">&#128065;</button>
                </div>
              </div>
              <div class="form-group">
                <label class="form-label" for="firstname"><%=lang(cnnDB, "FirstName")%> <span class="form-required">*</span></label>
                <input type="text" class="form-control" id="firstname" name="firstname" maxlength="25" value="<% = Server.HTMLEncode(firstname & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="lastname"><%=lang(cnnDB, "LastName")%> <span class="form-required">*</span></label>
                <input type="text" class="form-control" id="lastname" name="lastname" maxlength="24" value="<% = Server.HTMLEncode(lastname & "") %>">
              </div>
              <div class="form-group" style="grid-column:1/-1">
                <label class="form-label" for="email"><%=lang(cnnDB, "EmailAddress")%> <span class="form-required">*</span></label>
                <input type="email" class="form-control" id="email" name="email" maxlength="50" value="<% = Server.HTMLEncode(email & "") %>" placeholder="correo@empresa.com">
              </div>
            </div>
          </div>

          <!-- Contacto -->
          <div class="lhd-card au-card">
            <div class="section-header">&#128222; Contacto y ubicaci&oacute;n</div>
            <div class="au-grid">
              <div class="form-group">
                <label class="form-label" for="phone"><%=lang(cnnDB, "PhoneNumber")%></label>
                <input type="text" class="form-control" id="phone" name="phone" maxlength="50" value="<% = Server.HTMLEncode(phone & "") %>">
              </div>
              <% If optPager Then %>
              <div class="form-group">
                <label class="form-label" for="pager"><%=lang(cnnDB, "PagerAddress")%></label>
                <input type="text" class="form-control" id="pager" name="pager" maxlength="50" value="<% = Server.HTMLEncode(pager & "") %>">
              </div>
              <% Else %>
                <input type="hidden" name="pager" value="">
              <% End If %>
              <% If optInout Then %>
              <div class="form-group">
                <label class="form-label" for="phone_home"><%=lang(cnnDB, "HomePhone")%></label>
                <input type="text" class="form-control" id="phone_home" name="phone_home" maxlength="50" value="<% = Server.HTMLEncode(phone_home & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="phone_mobile"><%=lang(cnnDB, "MobilePhone")%></label>
                <input type="text" class="form-control" id="phone_mobile" name="phone_mobile" maxlength="50" value="<% = Server.HTMLEncode(phone_mobile & "") %>">
              </div>
              <% Else %>
                <input type="hidden" name="phone_home" value="">
                <input type="hidden" name="phone_mobile" value="">
              <% End If %>
              <div class="form-group">
                <label class="form-label" for="location"><%=lang(cnnDB, "Location")%></label>
                <input type="text" class="form-control" id="location" name="location" maxlength="50" value="<% = Server.HTMLEncode(location & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="department"><%=lang(cnnDB, "Department")%></label>
                <select class="form-control" id="department" name="department">
                  <option value="0"><%=lang(cnnDB, "NotSpecified")%></option>
                  <%
                    Dim depRes
                    Set depRes = SQLQuery(cnnDB, "SELECT department_id, dname From departments WHERE department_id > 0 ORDER BY dname ASC")
                    Do While Not depRes.EOF
                      If SafeCLng(depRes("department_id")) = SafeCLng(department) Then
                        Response.Write "<option value=""" & depRes("department_id") & """ selected>" & Server.HTMLEncode(depRes("dname") & "") & "</option>"
                      Else
                        Response.Write "<option value=""" & depRes("department_id") & """>" & Server.HTMLEncode(depRes("dname") & "") & "</option>"
                      End If
                      depRes.MoveNext
                    Loop
                    depRes.Close
                  %>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="usrLanguage"><%=lang(cnnDB, "Language")%></label>
                <select class="form-control" id="usrLanguage" name="usrLanguage">
                  <%
                    Dim rstLang, defLang, selLang
                    defLang = SafeCLng(Cfg(cnnDB, "DefaultLanguage"))
                    selLang = SafeCLng(usrLanguage)
                    If selLang = 0 Then selLang = defLang
                    Set rstLang = SQLQuery(cnnDB, "SELECT id, LangName, Localized From tblLanguage")
                    Do While Not rstLang.EOF
                      If SafeCLng(rstLang("id")) = selLang Then
                        Response.Write "<option value=""" & rstLang("id") & """ selected>" & Server.HTMLEncode(rstLang("LangName") & "") & " (" & Server.HTMLEncode(rstLang("Localized") & "") & ")</option>"
                      Else
                        Response.Write "<option value=""" & rstLang("id") & """>" & Server.HTMLEncode(rstLang("LangName") & "") & " (" & Server.HTMLEncode(rstLang("Localized") & "") & ")</option>"
                      End If
                      rstLang.MoveNext
                    Loop
                    rstLang.Close
                  %>
                </select>
              </div>
            </div>
          </div>

          <!-- Rol de soporte -->
          <div class="lhd-card au-card">
            <div class="section-header">&#128295; <%=lang(cnnDB, "SupportRep")%></div>
            <label class="au-switch">
              <input type="checkbox" name="isrep" id="isrep" <% If IsRep = 1 Then Response.Write("checked") %>>
              <span class="au-track"><span class="au-knob"></span></span>
              <span>
                <span class="t"><%=lang(cnnDB, "Enable")%> como agente de soporte</span><br>
                <span class="d">Podr&aacute; atender y gestionar tickets seg&uacute;n el nivel de acceso.</span>
              </span>
            </label>
            <div class="form-group" style="margin-bottom:0;max-width:340px">
              <label class="form-label" for="repaccess"><%=lang(cnnDB, "AccessLevel")%></label>
              <select class="form-control" id="repaccess" name="repaccess">
                <option value="0"<% If SafeCLng(RepAccess) = 0 Then Response.Write(" selected") %>><%=lang(cnnDB, "Normal")%></option>
                <option value="1"<% If SafeCLng(RepAccess) = 1 Then Response.Write(" selected") %>><%=lang(cnnDB, "Restricted")%></option>
                <option value="2"<% If SafeCLng(RepAccess) = 2 Then Response.Write(" selected") %>><%=lang(cnnDB, "ReadOnly")%></option>
              </select>
            </div>
          </div>

          <% If optInout Then %>
          <!-- Tablero In/Out -->
          <div class="lhd-card au-card">
            <div class="section-header">&#128203; <%=lang(cnnDB, "InOutBoard")%></div>
            <div class="au-grid">
              <div class="form-group">
                <label class="form-label" for="statuscode"><%=lang(cnnDB, "Status")%></label>
                <select class="form-control" id="statuscode" name="statuscode">
                  <option value="0"<% If SafeCLng(statuscode) = 0 Then Response.Write(" selected") %>><%=lang(cnnDB, "In")%></option>
                  <option value="1"<% If SafeCLng(statuscode) = 1 Then Response.Write(" selected") %>><%=lang(cnnDB, "Out")%></option>
                  <option value="2"<% If SafeCLng(statuscode) = 2 Then Response.Write(" selected") %>><%=lang(cnnDB, "Leave")%></option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="statustext"><%=lang(cnnDB, "Status")%>&nbsp;<%=lang(cnnDB, "text")%></label>
                <input type="text" class="form-control" id="statustext" name="statustext" value="<% = Server.HTMLEncode(statustext & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="listoninoutboard"><%=lang(cnnDB, "ListonBoard")%></label>
                <select class="form-control" id="listoninoutboard" name="listoninoutboard">
                  <option value="0"<% If SafeCLng(ListOnInoutboard) = 0 Then Response.Write(" selected") %>><%=lang(cnnDB, "NO")%></option>
                  <option value="1"<% If SafeCLng(ListOnInoutboard) = 1 Then Response.Write(" selected") %>><%=lang(cnnDB, "YES")%></option>
                </select>
              </div>
              <div class="form-group" style="display:flex;align-items:flex-end">
                <label class="au-switch" style="margin-bottom:0;width:100%">
                  <input type="checkbox" name="inoutadmin" <% If InoutAdmin = 1 Then Response.Write("checked") %>>
                  <span class="au-track"><span class="au-knob"></span></span>
                  <span class="t"><%=lang(cnnDB, "Administrator")%></span>
                </label>
              </div>
              <div class="form-group" style="grid-column:1/-1">
                <label class="form-label" for="jobfunction"><%=lang(cnnDB, "JobFunction")%></label>
                <textarea class="form-control" id="jobfunction" name="jobfunction" rows="3"><% = Server.HTMLEncode(jobfunction & "") %></textarea>
              </div>
              <div class="form-group" style="grid-column:1/-1">
                <label class="form-label" for="userresume"><%=lang(cnnDB, "Resume")%></label>
                <textarea class="form-control" id="userresume" name="userresume" rows="3"><% = Server.HTMLEncode(userresume & "") %></textarea>
              </div>
            </div>
          </div>
          <% Else %>
            <input type="hidden" name="statuscode" value="0">
            <input type="hidden" name="listoninoutboard" value="0">
            <input type="hidden" name="statustext" value="">
            <input type="hidden" name="jobfunction" value="">
            <input type="hidden" name="userresume" value="">
            <input type="hidden" name="inoutadmin" value="0">
          <% End If %>

          <div class="au-actions">
            <a class="btn btn-secondary" href="viewusers.asp">Cancelar</a>
            <button type="submit" class="btn btn-primary">&#128190; <%=lang(cnnDB, "Save")%></button>
          </div>
        </form>
      </div>
    </div>

    <script type="text/javascript">
    (function () {
      var pw = document.getElementById('newpassword');
      var eye = document.getElementById('btnEye');
      eye.onclick = function () { pw.type = (pw.type === 'password') ? 'text' : 'password'; pw.focus(); };

      // Validacion en el navegador antes de enviar (el servidor valida tambien)
      document.getElementById('frmAddUser').addEventListener('submit', function (e) {
        var req = [['uid', 'Usuario'], ['firstname', 'Nombre'], ['lastname', 'Apellido'], ['email', 'Correo'], ['newpassword', 'Contraseña']];
        var missing = [], i, el;
        for (i = 0; i < req.length; i++) {
          el = document.getElementById(req[i][0]);
          if (el) { el.className = el.className.replace(' is-invalid', ''); }
          if (el && !el.value.replace(/^\s+|\s+$/g, '')) { missing.push(req[i][1]); el.className += ' is-invalid'; }
        }
        var u = document.getElementById('uid').value;
        if (u && /\s/.test(u)) { missing.push('Usuario sin espacios'); document.getElementById('uid').className += ' is-invalid'; }
        if (missing.length) {
          e.preventDefault();
          var msg = 'Revisa: ' + missing.join(', ') + '.';
          if (window.LhdToast) { LhdToast.error(msg); } else { alert(msg); }
        }
      });
    })();
    </script>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
