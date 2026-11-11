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

  Filename: update.asp
  Date:     $Date: 2001/12/13 09:38:21 $
  Version:  $Revision: 1.51 $
  Purpose:  Update user information in the In/Out Board.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta charset="utf-8">
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "InOutBoard")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style>
      .io-wrap { max-width: 860px; margin: 0 auto; }
      .io-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 0 18px; }
      .io-top { display: grid; grid-template-columns: 1fr 170px; gap: 20px; align-items: start; }
      .io-photo { width: 140px; height: 170px; object-fit: cover; border-radius: 10px; border: 1px solid var(--color-border-light,#E2E8F0); background: #F8FAFC; }
      .io-photo-box { text-align: center; }
      .io-hint { font-size: 12px; color: #64748b; margin-top: 4px; }
      .io-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
      @media (max-width: 640px) {
        .io-top { grid-template-columns: 1fr; }
        .io-photo-box { order: -1; }
      }
    </style>
  </head>
  <body>

  <%
    ' See is user is validated
    Call CheckUser(cnnDB, sid)

    'Check if user is In/Out Board admin
    Dim inoutadmin
    inoutadmin = SafeCInt(Usr(cnnDB, sid, "inoutadmin"))

    ' Look for the problem id in the query string, if it's not
    ' there display an error.
    Dim mId
    mId = Cint(Request.QueryString("id"))

    If (Trim(mId) <> Trim(sid)) And (inoutadmin <> 1) Or (Len(mId) = 0) Then
      Dim deniedMsg
      deniedMsg = lang(cnnDB, "AccessDenied")
      cnnDB.Close
      Call DisplayError(3, deniedMsg)
    Else

      Dim strSql, rstUser, rstUpdate
      Dim phone_home, phone_mobile, jobfunction, userresume
      Dim frm_phone_home, frm_phone_mobile, frm_jobfunction, frm_userresume
      If Request.Form("save") = "1" Then
        phone_home = Request.Form("frm_phone_home")
        phone_home = Replace(phone_home, "'", "''")
        phone_mobile = Request.Form("frm_phone_mobile")
        phone_mobile = Replace(phone_mobile, "'", "''")
        jobfunction = Request.Form("frm_jobfunction")
        jobfunction = Replace(jobfunction, "'", "''")
        userresume = Request.Form("frm_userresume")
        userresume = Replace(userresume, "'", "''")

        strSql = "UPDATE tblUsers SET " &_
          "phone_home = '" & phone_home & "', " &_
          "phone_mobile = '" & phone_mobile & "', " &_
          "jobfunction = '" & jobfunction & "', " &_
          "userresume = '" & userresume & "' " &_
          "WHERE sid = " & mId
        Set rstUpdate = SQLQuery(cnnDB, strSql)
      End If
      ' Fetch userinformation from the database
      strSql = "SELECT tblUsers.*, departments.dname" &_
        " FROM tblUsers, departments " &_
        " WHERE (sid = " & mId & ") AND (department = department_id)"

      Set rstUser = SQLQuery(cnnDB, strSql)

      If rstUser.EOF Then
        Call DisplayError(3, lang(cnnDB, "Noresultsfound"))
      End If

      Dim uname, uid, check
      uname = rstUser("fname") & ""
      uid = rstUser("uid") & ""

      If rstUser("statuscode") > 0 Then
        check = "checked"
      Else
        check = ""
      End If

      If Request.Form("save") = "1" Then
        Call LhdToast("success", lang(cnnDB, "Information") & " " & lang(cnnDB, "isUpdated"))
      End If

      Call DisplayHeader(cnnDB, sid)

      Dim fsFileSys, strFullFileName, photoSrc
      Set fsFileSys = Server.CreateObject("Scripting.FileSystemObject")
      strFullFileName = Server.MapPath("../image/" & uid & ".jpg")
      If fsFileSys.FileExists(strFullFileName) Then
        photoSrc = "../image/" & Server.HTMLEncode(uid) & ".jpg"
      Else
        photoSrc = "../image/nopicture.gif"
      End If
  %>

  <div class="lhd-wrapper">

    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
      <div class="lhd-page-title" style="margin:0"><span class="title-icon">&#9998;</span> <%=lang(cnnDB, "UserDetails")%></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <a class="btn btn-secondary" href="details.asp?id=<% = mId %>">&larr; <%=lang(cnnDB, "Details")%></a>
      </div>
    </div>

    <div class="io-wrap">

      <form method="POST" action="update.asp?id=<% = mId %>">
        <input type="hidden" name="save" value="1">
        <div class="lhd-card">
          <div class="section-header">&#128100; <% = Server.HTMLEncode(uname) %></div>
          <div class="io-top">
            <div>
              <div class="io-grid">
                <div class="form-group">
                  <label class="form-label"><%=lang(cnnDB, "Fullname")%></label>
                  <input type="text" class="form-control" value="<% = Server.HTMLEncode(uname) %>" readonly>
                </div>
                <div class="form-group">
                  <label class="form-label"><%=lang(cnnDB, "Department")%></label>
                  <input type="text" class="form-control" value="<% = Server.HTMLEncode(rstUser("dname") & "") %>" readonly>
                </div>
                <div class="form-group">
                  <label class="form-label"><%=lang(cnnDB, "Username")%></label>
                  <input type="text" class="form-control" value="<% = Server.HTMLEncode(uid) %>" readonly>
                </div>
                <div class="form-group">
                  <label class="form-label"><%=lang(cnnDB, "Email")%></label>
                  <input type="text" class="form-control" value="<% = Server.HTMLEncode(rstUser("email1") & "") %>" readonly>
                </div>
                <div class="form-group">
                  <label class="form-label"><%=lang(cnnDB, "Phone")%></label>
                  <input type="text" class="form-control" value="<% = Server.HTMLEncode(rstUser("phone") & "") %>" readonly>
                </div>
                <div class="form-group">
                  <label class="form-label" for="frm_phone_home"><%=lang(cnnDB, "HomePhone")%></label>
                  <input type="text" class="form-control" id="frm_phone_home" name="frm_phone_home" value="<% = Server.HTMLEncode(rstUser("phone_home") & "") %>">
                </div>
                <div class="form-group">
                  <label class="form-label" for="frm_phone_mobile"><%=lang(cnnDB, "MobilePhone")%></label>
                  <input type="text" class="form-control" id="frm_phone_mobile" name="frm_phone_mobile" value="<% = Server.HTMLEncode(rstUser("phone_mobile") & "") %>">
                </div>
              </div>
            </div>
            <div class="io-photo-box">
              <img class="io-photo" src="<% = photoSrc %>" alt="<% = Server.HTMLEncode(uname) %>">
            </div>
          </div>

          <div class="form-group">
            <label class="form-label" for="frm_jobfunction"><%=lang(cnnDB, "Function")%></label>
            <textarea class="form-control" id="frm_jobfunction" name="frm_jobfunction" rows="4"><% = Server.HTMLEncode(rstUser("jobfunction") & "") %></textarea>
          </div>
          <div class="form-group">
            <label class="form-label" for="frm_userresume"><%=lang(cnnDB, "Resume")%></label>
            <textarea class="form-control" id="frm_userresume" name="frm_userresume" rows="4"><% = Server.HTMLEncode(rstUser("userresume") & "") %></textarea>
          </div>

          <div class="io-actions">
            <button type="submit" class="btn btn-primary"><%=lang(cnnDB, "Submit")%></button>
          </div>
        </div>
      </form>

      <form method="post" action="savefile.asp?uid=<% = mId %>" enctype="multipart/form-data">
        <div class="lhd-card">
          <div class="section-header">&#128247; <%=lang(cnnDB, "Uploadimage")%></div>
          <div class="form-group">
            <label class="form-label" for="blob"><%=lang(cnnDB, "Image")%></label>
            <input type="file" class="form-control" id="blob" name="blob">
            <div class="io-hint">(<%=lang(cnnDB, "MaxImageSize")%>:&nbsp;<% = cfg(cnnDB, "MaxImageSize") %>&nbsp;<%=lang(cnnDB, "Bytes")%>)</div>
          </div>
          <div class="io-actions">
            <button type="submit" class="btn btn-primary"><%=lang(cnnDB, "Upload")%></button>
          </div>
        </div>
      </form>

    </div>
  </div>

<%
      rstUser.Close
      Set fsFileSys = Nothing
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    End If
%>
</body>
</html>
