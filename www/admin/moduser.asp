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

  Filename: moduser.asp
  Date:     $Date: 2002/08/28 15:30:07 $
  Version:  $Revision: 1.52.4.2 $
  Purpose:  Form to modify user account info.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ModifyUser")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .mu-wrap { max-width: 860px; margin: 0 auto; }
      .mu-card { margin-bottom: 18px; }
      .mu-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 0 18px; }
      .mu-hint { font-size: 12px; color: #64748b; margin-top: 4px; }
      .mu-pass { position: relative; }
      .mu-pass .form-control { padding-right: 44px; }
      .mu-eye { position: absolute; right: 6px; top: 50%; transform: translateY(-50%); background: none; border: 0; cursor: pointer; color: #64748b; padding: 6px 8px; border-radius: 6px; }
      .mu-eye:hover { color: var(--color-primary, #0057B8); background: #F1F5F9; }
      .mu-switch { display: flex; align-items: center; gap: 12px; padding: 12px 14px; border: 1px solid var(--color-border-light, #E2E8F0); border-radius: 10px; background: #F8FAFC; cursor: pointer; margin-bottom: 14px; }
      .mu-switch input { position: absolute; opacity: 0; pointer-events: none; }
      .mu-track { position: relative; width: 42px; height: 22px; border-radius: 9999px; background: #CBD5E0; flex-shrink: 0; transition: background .2s ease; }
      .mu-knob { position: absolute; top: 2px; left: 2px; width: 18px; height: 18px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,.3); transition: transform .2s ease; }
      .mu-switch input:checked + .mu-track { background: var(--color-primary, #0057B8); }
      .mu-switch input:checked + .mu-track .mu-knob { transform: translateX(20px); }
      .mu-switch input:focus-visible + .mu-track { outline: 2px solid var(--color-primary, #0057B8); outline-offset: 2px; }
      .mu-switch .t { font-weight: 600; color: #1a202c; font-size: 14px; }
      .mu-switch .d { font-size: 12px; color: #64748b; }
      .mu-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
      .mu-user { display: inline-block; font-weight: 600; padding: 8px 12px; background: #F1F5F9; border-radius: 8px; }
      .mu-danger { border-left: 4px solid var(--color-danger, #B02A37); }
      .mu-danger-row { display: flex; align-items: center; justify-content: space-between; gap: 14px; flex-wrap: wrap; }
    </style>
  </head>
  <body>

    <%
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      Dim modSid
      modSid = SafeCLng(Request.Form("usersid"))

      Dim success, deleted, blnRepProbs, catsBlock
      success = FALSE
      deleted = FALSE
      blnRepProbs = False
      catsBlock = False

      Dim uid, email, fname, phone, location, department, pager
      Dim firstname, lastname, ListOnInoutboard, phone_home, phone_mobile
      Dim jobfunction, userresume, statuscode, statustext, statusdate
      Dim usrLanguage, IsRep, RepAccess, InoutAdmin
      Dim rstRepProbs, sqlString, updRes, newpassword

      If SafeCLng(Request.Form("save")) = 1 Then
        uid = Left(Lcase(Trim(CleanInputText(Request.Form("uid")))), 50)
        uid = Replace(uid, "'", "''")

        email = Left(Lcase(Trim(CleanInputText(Request.Form("email")))), 50)
        email = Replace(email, "'", "''")

        pager = Left(Lcase(Trim(CleanInputText(Request.Form("pager")))), 50)
        pager = Replace(pager, "'", "''")

        phone = Left(Trim(Replace(CleanInputText(Request.Form("phone")), "'", "''")), 50)
        location = Left(Trim(Replace(CleanInputText(Request.Form("location")), "'", "''")), 50)
        department = SafeCLng(Request.Form("department"))

        ListOnInoutBoard = SafeCLng(Request.Form("ListOnInoutBoard"))
        firstname = Left(Trim(Replace(CleanInputText(Request.Form("firstname")), "'", "''")), 25)
        lastname = Left(Trim(Replace(CleanInputText(Request.Form("lastname")), "'", "''")), 24)
        fname = firstname & " " & lastname

        phone_home = Left(Trim(Replace(CleanInputText(Request.Form("phone_home")), "'", "''")), 50)
        phone_mobile = Left(Trim(Replace(CleanInputText(Request.Form("phone_mobile")), "'", "''")), 50)
        jobfunction = Request.form("jobfunction")
        jobfunction = Replace(jobfunction , "'", "''")
        userresume = Request.form("userresume")
        userresume = Replace(userresume , "'", "''")
        statuscode = SafeCLng(Request.Form("statuscode"))
        statustext = Request.form("statustext")
        statustext = Replace(statustext , "'", "''")
        statusdate = SQLDate(Now, lhdAddSQLDelim)
        usrLanguage = SafeCLng(Request.form("usrLanguage"))
        RepAccess = SafeCLng(Request.Form("repaccess"))

        If Request.Form("InoutAdmin") = "on" Then
          InoutAdmin = 1
        Else
          InoutAdmin = 0
        End If

        If Len(email) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Emailaddress") & lang(cnnDB, "is a required field") & ".")
        End If
        If Len(firstname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "FirstName") & lang(cnnDB, "is a required field") & ".")
        End IF
        If Len(lastname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "LastName") & lang(cnnDB, "is a required field") & ".")
        End IF

        If Request.Form("isrep") = "on" Then
          IsRep = 1
        Else
          If Usr(cnnDB, modSid, "IsRep") = 1 Then
            Set rstRepProbs = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & modSid & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
            If Not rstRepProbs.EOF Then
              blnRepProbs = True
              IsRep = 1
            Else
              IsRep = 0
            End If
            rstRepProbs.Close
          Else
            IsRep = 0
          End If
        End If

        newpassword = Left(Trim(Request.Form("newpassword")), 50)
        newpassword = Replace(newpassword, "'", "''")
        If Len(newpassword) > 0 Then
          sqlString = "UPDATE tblUsers SET " & _
            "email1 = '" & email & "', " & _
            "email2 = '" & pager & "', " & _
            "fname = '" & fname & "', " & _
            "firstname = '" & firstname & "', " & _
            "lastname = '" & lastname & "', " & _
            "phone = '" & phone & "', " & _
            "phone_home = '" & phone_home & "', " & _
            "phone_mobile = '" & phone_mobile & "', " & _
            "location1 = '" & location & "', " & _
            "department = " & department & ", " & _
            "IsRep = " & IsRep & ", " & _
            "RepAccess = " & RepAccess & ", " & _
            "InoutAdmin = " & InoutAdmin & ", " & _
            "statuscode = " & statuscode & ", " & _
            "statustext = '" & statustext & "', " & _
            "statusdate = " & statusdate & ", " & _
            "jobfunction = '" & jobfunction & "', " & _
            "userresume = '" & userresume & "', " & _
            "ListOnInoutboard = " & ListOnInoutboard & ", " & _
            "[Language] = " & usrLanguage & ", " & _
            "[password] = '" & newpassword & "'"
        Else
          sqlString = "UPDATE tblUsers SET " & _
            "email1 = '" & email & "', " & _
            "email2 = '" & pager & "', " & _
            "fname = '" & fname & "', " & _
            "firstname = '" & firstname & "', " & _
            "lastname = '" & lastname & "', " & _
            "phone = '" & phone & "', " & _
            "phone_home = '" & phone_home & "', " & _
            "phone_mobile = '" & phone_mobile & "', " & _
            "location1 = '" & location & "', " & _
            "IsRep = " & IsRep & ", " & _
            "RepAccess = " & RepAccess & ", " & _
            "InoutAdmin = " & InoutAdmin & ", " & _
            "statuscode = " & statuscode & ", " & _
            "statustext = '" & statustext & "', " & _
            "statusdate = " & statusdate & ", " & _
            "jobfunction = '" & jobfunction & "', " & _
            "userresume = '" & userresume & "', " & _
            "ListOnInoutboard = " & ListOnInoutboard & ", " & _
            "[Language] = " & usrLanguage & ", " & _
            "department = " & department
        End If

        sqlString = sqlString & " WHERE sid=" & modSid
        Set updRes = SQLQuery(cnnDB, sqlString)
        success = True
      End If

      Dim frm_email, frm_fname, frm_phone, frm_location, frm_department, frm_pager
      Dim frm_firstname, frm_lastname, frm_phone_home, frm_phone_mobile
      Dim frm_jobfunction, frm_userresume, frm_listoninoutboard, frm_usrLanguage
      Dim frm_statuscode, frm_statustext, frm_IsRep, frm_RepAccess, frm_InoutAdmin, frm_uid

      If SafeCLng(Request.Form("delete")) = 1 Then
        Dim delRes, rstProblemUpd1, rstProblemUpd2, strUserId
        strUserId = usr(cnnDB, modSid, "uid")
        If Usr(cnnDB, modSid, "IsRep") = 0 Then
          Set delRes = SQLQuery(cnnDB, "DELETE FROM tblUsers WHERE sid = " & modSid)
          Set rstProblemUpd1 = SQLQuery(cnnDB, "UPDATE problems SET entered_by=0 WHERE entered_by = " & modSid)
          Set rstProblemUpd2 = SQLQuery(cnnDB, "UPDATE problems SET rep=0 WHERE rep = " & modSid)
          success = True
          deleted = True
        Else
          Dim rstRepCats
          Set rstRepCats = SQLQuery(cnnDB, "SELECT category_id FROM categories WHERE rep_id=" & modSid)
          If Not rstRepCats.EOF Then
            Call DisplayError(3, "Please reassign categories to a different support rep.")
          End If
          rstRepCats.Close
          Set rstRepProbs = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & modSid & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
          If Not rstRepProbs.EOF Then
            blnRepProbs = True
          Else
            Set delRes = SQLQuery(cnnDB, "DELETE FROM tblUsers WHERE sid = " & modSid)
            Set rstProblemUpd1 = SQLQuery(cnnDB, "UPDATE problems SET entered_by=0 WHERE entered_by = " & modSid)
            success = True
            deleted = True
          End If
          rstRepProbs.Close
        End If
        'Code to delete user image file if In/Out Board is activated
        If cfg(cnnDB, "UseInoutBoard") = 1 Then
          If success = True Then
            Dim objFSO, strUserImage
            strUserImage = Server.MapPath("..\image\" & strUserId & ".jpg")
            Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
            If objFSO.FileExists(strUserImage) Then
              objFSO.DeleteFile strUserImage, False
            End If
            Set objFSO = Nothing
          End If
        End If
      End If

      ' Datos para el formulario (se leen de la base, tambien despues de guardar)
      If Not deleted Then
        frm_uid = Usr(cnnDB, modSid, "uid")
        frm_email = Usr(cnnDB, modSid, "email1")
        frm_firstname = Usr(cnnDB, modSid, "firstname")
        frm_lastname = Usr(cnnDB, modSid, "lastname")
        frm_phone = Usr(cnnDB, modSid, "phone")
        frm_phone_home = Usr(cnnDB, modSid, "phone_home")
        frm_phone_mobile = Usr(cnnDB, modSid, "phone_mobile")
        frm_location = Usr(cnnDB, modSid, "location1")
        frm_department = Usr(cnnDB, modSid, "department")
        frm_pager = Usr(cnnDB, modSid, "email2")
        frm_statuscode = Usr(cnnDB, modSid, "statuscode")
        frm_statustext = Usr(cnnDB, modSid, "statustext")
        frm_jobfunction = Usr(cnnDB, modSid, "jobfunction")
        frm_userresume = Usr(cnnDB, modSid, "userresume")
        frm_listoninoutboard = Usr(cnnDB, modSid, "listoninoutboard")
        frm_usrLanguage = Usr(cnnDB, modSid, "[language]")
        frm_IsRep = Usr(cnnDB, modSid, "IsRep")
        frm_RepAccess = Usr(cnnDB, modSid, "RepAccess")
        frm_InoutAdmin = Usr(cnnDB, modSid, "InoutAdmin")
      End If

      Dim optPager, optInout
      optPager = (SafeCInt(Cfg(cnnDB, "EnablePager")) > 0)
      optInout = (SafeCInt(Cfg(cnnDB, "useinoutboard")) = 1)

      ' Resultado de la operacion
      If blnRepProbs Then
        Call LhdToast("error", lang(cnnDB, "ErrorUpdatingAccount"))
      ElseIf deleted Then
        Call LhdToast("success", lang(cnnDB, "AccountDeleted"))
      ElseIf success Then
        Call LhdToast("success", lang(cnnDB, "AccountUpdated"))
      End If
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#9999;</span>
          <%=lang(cnnDB, "ModifyUser")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="viewusers.asp">&larr; <%=lang(cnnDB, "ManageUsers")%></a>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>

      <div class="mu-wrap">

        <% If blnRepProbs Then %>
        <div class="alert alert-error">
          <span class="alert-icon">&#9888;</span>
          <span><strong><%=lang(cnnDB, "ErrorUpdatingAccount")%></strong> &mdash; <%=lang(cnnDB, "ErrorUpdatingAccountText")%>.</span>
        </div>
        <% ElseIf deleted Then %>
        <div class="alert alert-success">
          <span class="alert-icon">&#10004;</span>
          <span><strong><%=lang(cnnDB, "AccountDeleted")%></strong> &mdash; <%=lang(cnnDB, "Theaccounthasbeenremoved")%>.</span>
        </div>
        <% ElseIf success Then %>
        <div class="alert alert-success">
          <span class="alert-icon">&#10004;</span>
          <span><%=lang(cnnDB, "AccountUpdated")%></span>
        </div>
        <% End If %>

        <% If Not deleted Then %>
        <form name="upduser" id="frmModUser" action="moduser.asp" method="POST" novalidate>
          <input type="hidden" name="usersid" value="<% = modSid %>">
          <input type="hidden" name="save" value="1">

          <!-- Cuenta -->
          <div class="lhd-card mu-card">
            <div class="section-header">&#128100; Datos de la cuenta <span style="float:right;font-weight:400;font-size:12px;color:#64748b"><em>*</em> = <%=lang(cnnDB, "Required")%></span></div>
            <div class="mu-grid">
              <div class="form-group">
                <label class="form-label"><%=lang(cnnDB, "Username")%></label>
                <div><span class="mu-user"><% = Server.HTMLEncode(frm_uid & "") %></span></div>
              </div>
              <div class="form-group"></div>
              <div class="form-group">
                <label class="form-label" for="firstname"><%=lang(cnnDB, "FirstName")%> <span class="form-required">*</span></label>
                <input type="text" class="form-control" id="firstname" name="firstname" maxlength="25" value="<% = Server.HTMLEncode(frm_firstname & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="lastname"><%=lang(cnnDB, "LastName")%> <span class="form-required">*</span></label>
                <input type="text" class="form-control" id="lastname" name="lastname" maxlength="24" value="<% = Server.HTMLEncode(frm_lastname & "") %>">
              </div>
              <div class="form-group" style="grid-column:1/-1">
                <label class="form-label" for="email"><%=lang(cnnDB, "EmailAddress")%> <span class="form-required">*</span></label>
                <input type="text" class="form-control" id="email" name="email" maxlength="50" value="<% = Server.HTMLEncode(frm_email & "") %>">
              </div>
            </div>
          </div>

          <!-- Contacto -->
          <div class="lhd-card mu-card">
            <div class="section-header">&#128222; Contacto y ubicaci&oacute;n</div>
            <div class="mu-grid">
              <div class="form-group">
                <label class="form-label" for="phone"><%=lang(cnnDB, "PhoneNumber")%></label>
                <input type="text" class="form-control" id="phone" name="phone" maxlength="50" value="<% = Server.HTMLEncode(frm_phone & "") %>">
              </div>
              <% If optPager And (SafeCLng(Usr(cnnDB, modSid, "IsRep")) = 1) Then %>
              <div class="form-group">
                <label class="form-label" for="pager"><%=lang(cnnDB, "PagerAddress")%></label>
                <input type="text" class="form-control" id="pager" name="pager" maxlength="50" value="<% = Server.HTMLEncode(frm_pager & "") %>">
              </div>
              <% Else %>
                <input type="hidden" name="pager" value="">
              <% End If %>
              <% If optInout Then %>
              <div class="form-group">
                <label class="form-label" for="phone_home"><%=lang(cnnDB, "HomePhone")%></label>
                <input type="text" class="form-control" id="phone_home" name="phone_home" maxlength="50" value="<% = Server.HTMLEncode(frm_phone_home & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="phone_mobile"><%=lang(cnnDB, "MobilePhone")%></label>
                <input type="text" class="form-control" id="phone_mobile" name="phone_mobile" maxlength="50" value="<% = Server.HTMLEncode(frm_phone_mobile & "") %>">
              </div>
              <% Else %>
                <input type="hidden" name="phone_home" value="">
                <input type="hidden" name="phone_mobile" value="">
              <% End If %>
              <div class="form-group">
                <label class="form-label" for="location"><%=lang(cnnDB, "Location")%></label>
                <input type="text" class="form-control" id="location" name="location" maxlength="50" value="<% = Server.HTMLEncode(frm_location & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="department"><%=lang(cnnDB, "Department")%></label>
                <select class="form-control" id="department" name="department">
                  <option value="0"><%=lang(cnnDB, "NotSpecified")%></option>
                  <%
                    Dim depRes
                    Set depRes = SQLQuery(cnnDB, "SELECT department_id, dname From departments WHERE department_id > 0 ORDER BY dname ASC")
                    Do While Not depRes.EOF
                      If SafeCLng(depRes("department_id")) = SafeCLng(frm_department) Then
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
                    Dim rstLang
                    Set rstLang = SQLQuery(cnnDB, "SELECT id, LangName, Localized From tblLanguage")
                    Do While Not rstLang.EOF
                      If SafeCLng(rstLang("id")) = SafeCLng(frm_usrLanguage) Then
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
          <div class="lhd-card mu-card">
            <div class="section-header">&#128295; <%=lang(cnnDB, "SupportRep")%></div>
            <label class="mu-switch">
              <input type="checkbox" name="isrep" id="isrep" <% If SafeCLng(frm_IsRep) <> 0 Then Response.Write("checked") %>>
              <span class="mu-track"><span class="mu-knob"></span></span>
              <span>
                <span class="t"><%=lang(cnnDB, "Enable")%> como agente de soporte</span><br>
                <span class="d">Podr&aacute; atender y gestionar tickets seg&uacute;n el nivel de acceso.</span>
              </span>
            </label>
            <div class="form-group" style="margin-bottom:0;max-width:340px">
              <label class="form-label" for="repaccess"><%=lang(cnnDB, "AccessLevel")%></label>
              <select class="form-control" id="repaccess" name="repaccess">
                <option value="0"<% If SafeCLng(frm_RepAccess) = 0 Then Response.Write(" selected") %>><%=lang(cnnDB, "Normal")%></option>
                <option value="1"<% If SafeCLng(frm_RepAccess) = 1 Then Response.Write(" selected") %>><%=lang(cnnDB, "Restricted")%></option>
                <option value="2"<% If SafeCLng(frm_RepAccess) = 2 Then Response.Write(" selected") %>><%=lang(cnnDB, "ReadOnly")%></option>
              </select>
            </div>
          </div>

          <% If optInout Then %>
          <!-- Tablero In/Out -->
          <div class="lhd-card mu-card">
            <div class="section-header">&#128203; <%=lang(cnnDB, "InOutBoard")%></div>
            <div class="mu-grid">
              <div class="form-group">
                <label class="form-label" for="statuscode"><%=lang(cnnDB, "Status")%></label>
                <select class="form-control" id="statuscode" name="statuscode">
                  <option value="0"<% If SafeCLng(frm_statuscode) <> 1 And SafeCLng(frm_statuscode) <> 2 Then Response.Write(" selected") %>><%=lang(cnnDB, "In")%></option>
                  <option value="1"<% If SafeCLng(frm_statuscode) = 1 Then Response.Write(" selected") %>><%=lang(cnnDB, "Out")%></option>
                  <option value="2"<% If SafeCLng(frm_statuscode) = 2 Then Response.Write(" selected") %>><%=lang(cnnDB, "Leave")%></option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="statustext"><%=lang(cnnDB, "Status")%>&nbsp;<%=lang(cnnDB, "text")%></label>
                <input type="text" class="form-control" id="statustext" name="statustext" value="<% = Server.HTMLEncode(frm_statustext & "") %>">
              </div>
              <div class="form-group">
                <label class="form-label" for="listoninoutboard"><%=lang(cnnDB, "ListonBoard")%></label>
                <select class="form-control" id="listoninoutboard" name="listoninoutboard">
                  <option value="0"<% If CStr(frm_listoninoutboard & "") = "0" Then Response.Write(" selected") %>><%=lang(cnnDB, "NO")%></option>
                  <option value="1"<% If CStr(frm_listoninoutboard & "") <> "0" Then Response.Write(" selected") %>><%=lang(cnnDB, "YES")%></option>
                </select>
              </div>
              <div class="form-group" style="display:flex;align-items:flex-end">
                <label class="mu-switch" style="margin-bottom:0;width:100%">
                  <input type="checkbox" name="inoutadmin" <% If SafeCLng(frm_InoutAdmin) <> 0 Then Response.Write("checked") %>>
                  <span class="mu-track"><span class="mu-knob"></span></span>
                  <span class="t"><%=lang(cnnDB, "Administrator")%></span>
                </label>
              </div>
              <div class="form-group" style="grid-column:1/-1">
                <label class="form-label" for="jobfunction"><%=lang(cnnDB, "JobFunction")%></label>
                <textarea class="form-control" id="jobfunction" name="jobfunction" rows="3"><% = Server.HTMLEncode(frm_jobfunction & "") %></textarea>
              </div>
              <div class="form-group" style="grid-column:1/-1">
                <label class="form-label" for="userresume"><%=lang(cnnDB, "Resume")%></label>
                <textarea class="form-control" id="userresume" name="userresume" rows="3"><% = Server.HTMLEncode(frm_userresume & "") %></textarea>
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

          <!-- Contrasena -->
          <div class="lhd-card mu-card">
            <div class="section-header">&#128274; <%=lang(cnnDB, "Password")%></div>
            <div class="form-group" style="margin-bottom:0;max-width:340px">
              <label class="form-label" for="newpassword"><%=lang(cnnDB, "NewPassword")%></label>
              <div class="mu-pass">
                <input type="password" class="form-control" id="newpassword" name="newpassword" maxlength="50" autocomplete="new-password">
                <button type="button" class="mu-eye" id="btnEye" title="Mostrar u ocultar" aria-label="Mostrar u ocultar la contrase&ntilde;a">&#128065;</button>
              </div>
              <div class="mu-hint">D&eacute;jala vac&iacute;a para conservar la contrase&ntilde;a actual.</div>
            </div>
          </div>

          <div class="mu-actions">
            <a class="btn btn-secondary" href="viewusers.asp">Cancelar</a>
            <button type="submit" class="btn btn-primary">&#128190; <%=lang(cnnDB, "Save")%></button>
          </div>
        </form>

        <!-- Eliminar cuenta -->
        <div class="lhd-card mu-card mu-danger" style="margin-top:18px">
          <div class="section-header">&#128465; <%=lang(cnnDB, "DeleteAccount")%></div>
          <form name="deluser" id="frmDelUser" action="moduser.asp" method="POST">
            <input type="hidden" name="usersid" value="<% = modSid %>">
            <input type="hidden" name="delete" value="1">
            <div class="mu-danger-row">
              <span class="mu-hint" style="margin:0">Esta acci&oacute;n elimina la cuenta de forma permanente y no se puede deshacer.</span>
              <button type="submit" class="btn btn-danger">&#128465; <%=lang(cnnDB, "DeleteAccount")%></button>
            </div>
          </form>
        </div>
        <% Else %>
        <div class="mu-actions">
          <a class="btn btn-primary" href="viewusers.asp">&larr; <%=lang(cnnDB, "ManageUsers")%></a>
        </div>
        <% End If %>
      </div>
    </div>

    <script type="text/javascript">
    (function () {
      var pw = document.getElementById('newpassword');
      var eye = document.getElementById('btnEye');
      if (pw && eye) { eye.onclick = function () { pw.type = (pw.type === 'password') ? 'text' : 'password'; pw.focus(); }; }

      var f = document.getElementById('frmModUser');
      if (f) {
        f.addEventListener('submit', function (e) {
          var req = [['firstname', 'Nombre'], ['lastname', 'Apellido'], ['email', 'Correo']];
          var missing = [], i, el;
          for (i = 0; i < req.length; i++) {
            el = document.getElementById(req[i][0]);
            if (el) { el.className = el.className.replace(' is-invalid', ''); }
            if (el && !el.value.replace(/^\s+|\s+$/g, '')) { missing.push(req[i][1]); el.className += ' is-invalid'; }
          }
          if (missing.length) {
            e.preventDefault();
            var msg = 'Revisa: ' + missing.join(', ') + '.';
            if (window.LhdToast) { LhdToast.error(msg); } else { alert(msg); }
          }
        });
      }

      var d = document.getElementById('frmDelUser');
      if (d) {
        d.addEventListener('submit', function (e) {
          if (!window.confirm('¿Eliminar esta cuenta de usuario? Esta acción no se puede deshacer.')) { e.preventDefault(); }
        });
      }
    })();
    </script>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
