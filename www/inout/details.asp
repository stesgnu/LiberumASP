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

  Filename: details.asp
  Date:     $Date: 2001/12/13 21:47:38 $
  Version:  $Revision: 1.51 $
  Purpose:  This is the details form for the In/Out Board.
  -->

  <!-- 	#include file = "../public.asp" -->
  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    ' Look for the problem id in the query string, if it's not
    ' there display an error.

    If Len(Request.QueryString("id")) = 0 Then
      Dim noIdMsg
      noIdMsg = lang(cnnDB, "NovalidIDgiven")
      cnnDB.Close
      Call DisplayError(3, noIdMsg)
    End If
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
      .io-grid { display: grid; grid-template-columns: 1fr 170px; gap: 20px; align-items: start; }
      .io-photo { width: 140px; height: 170px; object-fit: cover; border-radius: 10px; border: 1px solid var(--color-border-light,#E2E8F0); background: #F8FAFC; }
      .io-photo-box { text-align: center; }
      .io-dl { display: grid; grid-template-columns: 160px 1fr; gap: 10px 16px; margin: 0; font-size: 14px; }
      .io-dl dt { color: #64748b; font-weight: 600; }
      .io-dl dd { margin: 0; color: #1a202c; word-break: break-word; }
      .io-text { white-space: pre-wrap; line-height: 1.6; font-size: 14px; color: #2d3748; word-break: break-word; }
      @media (max-width: 640px) {
        .io-grid { grid-template-columns: 1fr; }
        .io-photo-box { order: -1; }
        .io-dl { grid-template-columns: 1fr; gap: 2px 0; }
        .io-dl dd { margin-bottom: 8px; }
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

      Dim usid
      usid = SafeCLng(Request.QueryString("id"))

      Dim strSql, rstUser
      strSql = "SELECT tblUsers.*, departments.dname" &_
        " FROM tblUsers, departments " &_
        " WHERE (sid = " & usid & ")" &_
        " AND (tblUsers.department = departments.department_id) "

      Set rstUser = SQLQuery(cnnDB, strSql)

      ' If no results are returned, display an error
      If rstUser.EOF Then
        rstUser.Close
        Dim noResMsg
        noResMsg = lang(cnnDB, "Noresultsfound")
        cnnDB.Close
        Call DisplayError(3, noResMsg)
      End If

      Dim uname, uid
      uname = rstUser("fname") & ""
      uid = rstUser("uid") & ""

      Call DisplayHeader(cnnDB, sid)

      Dim fsFileSys, strFullFileName, photoSrc
      Set fsFileSys = Server.CreateObject("Scripting.FileSystemObject")
      strFullFileName = Server.MapPath("../image/" & uid & ".jpg")
      If fsFileSys.FileExists(strFullFileName) Then
        photoSrc = "../image/" & Server.HTMLEncode(uid) & ".jpg"
      Else
        photoSrc = "../image/nopicture.gif"
      End If

      Dim str, phHome, phMobile
      str = rstUser("phone_home") & ""
      If Len(str) = 8 Then
        str = Mid(str, 1, 2) & " " & Mid(str, 3, 2) & " " & Mid(str, 5, 2) & " " & Mid(str, 7, 2)
      End If
      phHome = str
      str = rstUser("phone_mobile") & ""
      If Len(str) = 8 Then
        str = Mid(str, 1, 3) & " " & Mid(str, 4, 2) & " " & Mid(str, 6, 3)
      End If
      phMobile = str

      Dim canEdit, canStatus
      canEdit = (Trim(usid) = Trim(sid)) Or (inoutadmin = 1)
      canStatus = (rstUser("statuscode") <> 2) Or (inoutadmin = 1)
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0"><span class="title-icon">&#128100;</span> <%=lang(cnnDB, "UserDetails")%></div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <% If canEdit Then %>
            <a class="btn btn-secondary" href="update.asp?id=<% = usid %>">&#9998; <%=lang(cnnDB, "Edit")%></a>
          <% End If %>
          <% If canStatus Then %>
            <a class="btn btn-primary" href="status.asp?id=<% = usid %>">&#8644; <%=lang(cnnDB, "ChangeStatus")%></a>
          <% End If %>
        </div>
      </div>

      <div class="lhd-card">
        <div class="section-header">&#128100; <% = Server.HTMLEncode(uname) %></div>
        <div class="io-grid">
          <dl class="io-dl">
            <dt><%=lang(cnnDB, "Fullname")%></dt>
            <dd><strong><% = Server.HTMLEncode(uname) %></strong></dd>
            <dt><%=lang(cnnDB, "Department")%></dt>
            <dd><strong><% = Server.HTMLEncode(rstUser("dname") & "") %></strong></dd>
            <dt><%=lang(cnnDB, "UserName")%></dt>
            <dd><strong><% = Server.HTMLEncode(uid) %></strong></dd>
            <dt><%=lang(cnnDB, "Email")%></dt>
            <dd><strong><% = Server.HTMLEncode(rstUser("email1") & "") %></strong></dd>
            <dt><%=lang(cnnDB, "Phone")%></dt>
            <dd><strong><% = Server.HTMLEncode(rstUser("phone") & "") %></strong></dd>
            <dt><%=lang(cnnDB, "HomePhone")%></dt>
            <dd><strong><% = Server.HTMLEncode(phHome) %></strong></dd>
            <dt><%=lang(cnnDB, "MobilePhone")%></dt>
            <dd><strong><% = Server.HTMLEncode(phMobile) %></strong></dd>
            <dt><%=lang(cnnDB, "Status")%></dt>
            <dd>
              <%
                Select Case rstUser("statuscode")
                  Case 2
                    Response.Write "<img src=""../image/yellow_pin.gif"" alt="""" style=""vertical-align:middle"">&nbsp;" & Server.HTMLEncode(rstUser("statustext") & "")
                  Case 1
                    Response.Write "<img src=""../image/red_pin.gif"" alt="""" style=""vertical-align:middle"">&nbsp;" & Server.HTMLEncode(rstUser("statustext") & "")
                  Case 0
                    Response.Write "<img src=""../image/green_pin.gif"" alt="""" style=""vertical-align:middle"">"
                End Select
              %>
            </dd>
          </dl>
          <div class="io-photo-box">
            <img class="io-photo" src="<% = photoSrc %>" alt="<% = Server.HTMLEncode(uname) %>">
          </div>
        </div>
      </div>

      <div class="lhd-card">
        <div class="section-header">&#128188; <%=lang(cnnDB, "Function")%></div>
        <div class="io-text"><% = Server.HTMLEncode(rstUser("jobfunction") & "") %></div>
      </div>

      <div class="lhd-card">
        <div class="section-header">&#128196; <%=lang(cnnDB, "Resume")%></div>
        <div class="io-text"><% = Server.HTMLEncode(rstUser("userresume") & "") %></div>
      </div>

    </div>
    <%
      rstUser.Close
      Set fsFileSys = Nothing
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
</body>
</html>
