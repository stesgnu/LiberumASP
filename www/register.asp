<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% 
  Option Explicit
  Response.Buffer = TRUE
  Response.Expires = -1
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: register.asp
  Date:     $Date: 2002/01/21 20:35:22 $
  Version:  $Revision: 1.50.2.1 $
  Purpose:  User registration & profile management.
  -->
  <!-- 	#include file = "public.asp" -->
  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>
  <head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta charset="utf-8">
    <title>
      <%=lang(cnnDB, "HelpDesk")%> &nbsp;-&nbsp; <%=lang(cnnDB, "Register")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="default.css">  
  </head>
  <body>

  <%
    Dim success, edit
    success = FALSE
    If (Request.QueryString("edit") = 1) OR (Request.Form("edit") = 1) Then
      edit = True
    Else
      edit = False
    End If

    If Request.Form("create") = 1 Then
      Dim uid, email, password1, password2, fname, phone, location, department, pager
      Dim firstname, lastname, phone_home, phone_mobile, usrLanguage, dateformat
      uid = Left(Lcase(Trim(CleanInputText(Request.Form("uid")))), 50)

      email = Left(Lcase(Trim(CleanInputText(Request.Form("email")))), 50)
      email = Replace(email, "'", "''")

      pager = Left(Lcase(Trim(CleanInputText(Request.Form("pager")))), 50)
      pager = Replace(pager, "'", "''")

      password1 = Left(Trim(Replace(Request.Form("password1"), "'", "''")), 50)
      password2 = Left(Trim(Replace(Request.Form("password2"), "'", "''")), 50)

      phone = Left(Trim(Replace(CleanInputText(Request.Form("phone")), "'", "''")), 50)
      location = CleanInputText(Left(Trim(Replace(CleanInputText(Request.Form("location")), "'", "''")), 50))
      department = Cint(Request.Form("department"))
      usrLanguage = Cint(Request.Form("usrLanguage"))

      firstname = CleanInputText(Left(Trim(Replace(CleanInputText(Request.Form("firstname")), "'", "''")), 25))
      lastname = CleanInputText(Left(Trim(Replace(CleanInputText(Request.Form("lastname")), "'", "''")), 24))
      fname = firstname & " " & lastname
      dateformat = Left(Trim(Lcase(Request.Form("dateformat"))), 12)
      dateformat = Replace(dateformat, "'", "''")

      If cfg(cnnDB,"useinoutboard") = 1 Then
        phone_home = Left(Trim(Replace(CleanInputText(Request.Form("phone_home")), "'", "''")), 50)
        phone_mobile = Left(Trim(Replace(CleanInputText(Request.Form("phone_mobile")), "'", "''")), 50)
      End If

      If Len(email) = 0 Then
        Call DisplayError (3, lang(cnnDB, "Emailaddress") & " " & lang(cnnDB, "isarequiredfield") & ".")
      End If
      If Len(firstname) = 0 Then
        Call DisplayError (3, lang(cnnDB, "FirstName") & " " & lang(cnnDB, "isarequiredfield") & ".")
      End IF
      If Len(lastname) = 0 Then
        Call DisplayError (3, lang(cnnDB, "LastName") & " " & lang(cnnDB, "isarequiredfield") & ".")
      End IF

      Dim sqlString

      If Edit Then
        Dim oldpassword
        oldpassword = Left(Trim(Request.Form("oldpassword")), 50)
        If (Len(oldpassword) > 0) Or (Len(password1) > 0) Then
          If password1 <> password2 Then
            Call DisplayError (3, lang(cnnDB, "Passwordsdonotmatch") & ".")
          End If
          If oldpassword <> Usr(cnnDB, sid, "password") Then
            Call DisplayError (3, lang(cnnDB, "Passwordisincorrect") & ".")
          End If
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
            "[language] = " & usrLanguage & ", " & _
            "[password] = '" & password1 & "'"
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
            "[language] = " & usrLanguage & ", " & _
            "department = " & department
        End If

        sqlString = sqlString & " WHERE sid=" & sid

        Dim updRes
        Set updRes = SQLQuery(cnnDB, sqlString)
        success = True
        Session("lhd_LanguageID") = Empty
      Else
        If Len(uid) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & lang(cnnDB, "isarequiredfield") & ".")
        End If
        If CBool(InStr(uid, "'")) Then
          Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & Lang(cnnDB, "containsinvalidcharacters") & ".")
        End If
        If password1 <> password2 Then
          Call DisplayError (3, lang(cnnDB, "Passwordsdonotmatch") & ".")
        End If

        If Len(password1) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Password") &"&nbsp;" & lang(cnnDB, "isarequiredfield") & ".")
        End IF

        Dim checkNameRes
        Set checkNameRes = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE uid='" & uid & "'")
        If Not checkNameRes.EOF Then
          checkNameRes.Close
          Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & lang(cnnDB, "alreadyinuse") & ".")
        Else
          checkNameRes.Close
        End If

        Dim newSid
        newSid = GetUnique(cnnDB, "users")

        sqlString = "INSERT INTO tblUsers " & _
          "(sid, uid, [password], email1, email2, fname, firstname, lastname, phone, phone_home, phone_mobile, location1, [language], department) VALUES (" & _
          newSid & ", " & _
          "'" & uid & "', " & _
          "'" & password1 & "', " & _
          "'" & email & "', " & _
          "'" & pager & "', " & _
          "'" & fname & "', " & _
          "'" & firstname & "', " & _
          "'" & lastname & "', " & _
          "'" & phone & "', " & _
          "'" & phone_home & "', " & _
          "'" & phone_mobile & "', " & _
          "'" & location & "', " & _
          usrLanguage & ", " & _
          department & _
          ")"

        Dim insertRes
        Set insertRes = SQLQuery(cnnDB, sqlString)
        success = TRUE
      End If
    End If

    Dim frm_email, frm_fname, frm_phone, frm_location, frm_department, frm_pager
    Dim frm_firstname, frm_lastname, frm_phone_home, frm_phone_mobile, frm_usrLanguage, frm_dateformat
    If Edit Then
      frm_email = Usr(cnnDB, sid, "email1")
      frm_phone = Usr(cnnDB, sid, "phone")
      frm_location = Usr(cnnDB, sid, "location1")
      frm_department = Usr(cnnDB, sid, "department")
      frm_pager = Usr(cnnDB, sid, "email2")
      frm_firstname = Usr(cnnDB, sid, "firstname")
      frm_lastname = Usr(cnnDB, sid, "lastname")
      frm_dateformat = "yyyy-mm-dd"
      frm_usrLanguage = Usr(cnnDB, sid, "[language]")
      If cfg(cnnDB,"useinoutboard") = 1 Then
        frm_phone_home = Usr(cnnDB, sid, "phone_home")
        frm_phone_mobile = Usr(cnnDB, sid, "phone_mobile")
      End If
    End If

    Dim varDateFormatList
    varDateFormatList = Array("mm/dd/yyyy", "dd/mm/yyyy", "yyyy-mm-dd")

    If sid > 0 Then
      Call DisplayHeader(cnnDB, sid)
    End If
  %>

  <div class="lhd-wrapper" style="max-width:700px;margin:30px auto">

    <div class="lhd-card">
      <div class="section-header">
        <% If Edit Then %>
          &#9881; <%=lang(cnnDB, "UpdateInformation")%>
        <% Else %>
          &#128100; <%=lang(cnnDB, "Registration")%> &mdash; <%=lang(cnnDB, "fornewusers")%>
        <% End If %>
      </div>

      <% If success Then %>
        <% If Edit Then %>
          <div style="text-align:center;padding:30px 10px">
            <div style="font-size:42px;color:var(--color-success);margin-bottom:12px">&#10004;</div>
            <h3><%=lang(cnnDB, "AccountUpdated")%></h3>
            <p style="color:var(--color-muted);font-size:14px;margin-bottom:20px"><%=lang(cnnDB, "accountupdatedtext")%>.</p>
            <a class="btn btn-primary" href="default.asp"><%=lang(cnnDB, "MainMenu")%> &rarr;</a>
          </div>
        <% Else %>
          <div style="text-align:center;padding:30px 10px">
            <div style="font-size:42px;color:var(--color-success);margin-bottom:12px">&#10004;</div>
            <h3><%=lang(cnnDB, "AccountCreated")%></h3>
            <p style="color:var(--color-muted);font-size:14px;margin-bottom:20px"><%=lang(cnnDB, "AccountCreatedText")%></p>
            <a class="btn btn-primary" href="logon.asp"><%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Logon")%> &rarr;</a>
          </div>
        <% End If %>
      <% Else %>

        <form action="register.asp" method="POST" id="reg-form" novalidate>
          <% If Edit Then %>
            <input type="hidden" name="edit" value="1">
            <input type="hidden" name="create" value="1">
          <% Else %>
            <input type="hidden" name="create" value="1">
          <% End If %>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "Username")%> <span class="form-required">*</span></label>
              <% If Edit Then %>
                <input type="text" class="form-control" value="<% = Server.HTMLEncode(Usr(cnnDB, sid, "uid")) %>" readonly disabled>
              <% Else %>
                <input class="form-control" type="text" name="uid" placeholder="nombre.usuario" required autofocus>
              <% End If %>
            </div>

            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "Emailaddress")%> <span class="form-required">*</span></label>
              <input class="form-control" type="email" name="email" value="<% = frm_email %>" placeholder="correo@empresa.com" required>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "FirstName")%> <span class="form-required">*</span></label>
              <input class="form-control" type="text" name="firstname" value="<% = frm_firstname %>" placeholder="Nombre" required>
            </div>
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "LastName")%> <span class="form-required">*</span></label>
              <input class="form-control" type="text" name="lastname" value="<% = frm_lastname %>" placeholder="Apellido" required>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "PhoneNumber")%></label>
              <input class="form-control" type="tel" name="phone" value="<% = frm_phone %>" placeholder="Tel&eacute;fono / Extensi&oacute;n">
            </div>
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "Location")%></label>
              <input class="form-control" type="text" name="location" value="<% = frm_location %>" placeholder="Oficina, edificio o piso">
            </div>
          </div>

          <% If (Cfg(cnnDB, "EnablePager") > 0) And (sid > 0 And Usr(cnnDB, sid, "IsRep") = 1) Then %>
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "PagerAddress")%></label>
              <input class="form-control" type="text" name="pager" value="<% = frm_pager %>">
            </div>
          <% Else %>
            <input type="hidden" name="pager" value="">
          <% End If %>

          <% If Cfg(cnnDB, "useinoutboard") = 1 Then %>
            <div class="form-row">
              <div class="form-group">
                <label class="form-label"><%=lang(cnnDB, "HomePhone")%></label>
                <input class="form-control" type="tel" name="phone_home" value="<% = frm_phone_home %>">
              </div>
              <div class="form-group">
                <label class="form-label"><%=lang(cnnDB, "MobilePhone")%></label>
                <input class="form-control" type="tel" name="phone_mobile" value="<% = frm_phone_mobile %>">
              </div>
            </div>
          <% Else %>
            <input type="hidden" name="phone_home" value="">
            <input type="hidden" name="phone_mobile" value="">
          <% End If %>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "Department")%></label>
              <select name="department" class="form-control">
                <option value="0"><%=lang(cnnDB, "NotSpecified")%></option>
                <%
                  Dim depRes
                  Set depRes = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
                  If Not depRes.EOF Then
                    Do While Not depRes.EOF
                      If depRes("department_id") = frm_department Then
                %>
                        <option value="<% = depRes("department_id") %>" selected><% = depRes("dname") %></option>
                <%    Else %>
                        <option value="<% = depRes("department_id") %>"><% = depRes("dname") %></option>
                <%    End If
                      depRes.MoveNext
                    Loop
                  End If
                  depRes.Close
                %>
              </select>
            </div>

            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "Language")%></label>
              <select name="usrLanguage" class="form-control">
                <%
                  Dim rstLang
                  Set rstLang = SQLQuery(cnnDB, "SELECT * From tblLanguage")
                  If Not rstLang.EOF Then
                    Do While Not rstLang.EOF
                      If rstLang("id") = frm_usrLanguage Then
                %>
                        <option value="<% = rstLang("id") %>" selected><% = rstLang("LangName") %> (<% = rstLang("Localized") %>)</option>
                <%    Else %>
                        <option value="<% = rstLang("id") %>"><% = rstLang("LangName") %> (<% = rstLang("Localized") %>)</option>
                <%    End If
                      rstLang.MoveNext
                    Loop
                  End If
                  rstLang.Close
                %>
              </select>
            </div>
          </div>


          <% If Cfg(cnnDB, "AuthType") = 2 Then %>
            <div class="section-subheader" style="margin-top:24px">&#128274; Seguridad y Contrase&ntilde;a</div>

            <% If Edit Then %>
              <div class="form-group">
                <label class="form-label"><%=lang(cnnDB, "OldPassword")%></label>
                <input class="form-control" type="password" name="oldpassword" placeholder="Contrase&ntilde;a actual">
              </div>
            <% End If %>

            <div class="form-row">
              <div class="form-group">
                <label class="form-label"><%=lang(cnnDB, "Password")%> <span class="form-required">*</span></label>
                <input class="form-control" type="password" name="password1" placeholder="Nueva contrase&ntilde;a">
              </div>
              <div class="form-group">
                <label class="form-label"><%=lang(cnnDB, "ConfirmPassword")%> <span class="form-required">*</span></label>
                <input class="form-control" type="password" name="password2" placeholder="Confirmar contrase&ntilde;a">
              </div>
            </div>
          <% Else %>
            <input type="hidden" name="oldpassword" value="">
            <input type="hidden" name="password1" value="">
            <input type="hidden" name="password2" value="">
          <% End If %>

          <div style="display:flex;justify-content:space-between;align-items:center;margin-top:20px;flex-wrap:wrap;gap:12px">
            <div style="font-size:12px;color:var(--color-muted)">
              <span class="form-required">*</span> = <%=lang(cnnDB, "Required")%>
            </div>
            <div style="display:flex;gap:10px">
              <% If Not Edit Then %>
                <a class="btn btn-secondary" href="logon.asp">&#8592; <%=lang(cnnDB, "Logon")%></a>
              <% End If %>
              <button class="btn btn-primary btn-lg" type="submit">
                &#10004; <%=lang(cnnDB, "Submit")%>
              </button>
            </div>
          </div>

        </form>
      <% End If %>
    </div>

  </div>

  <%
    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close
  %>
  </body>
</html>
