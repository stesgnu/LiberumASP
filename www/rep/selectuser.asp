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

  Filename: selectuser.asp
  Date:     $Date: 2002/08/28 15:30:54 $
  Version:  $Revision: 1.51.2.1.2.2 $
  Purpose:  Ventana emergente para elegir un usuario (nuevo problema). Sin barra lateral.
  -->

  <!-- 	#include file = "../public.asp" -->
  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    Dim strSearchName, strSearchSql, blnPost
    strSearchName = Trim(Request.Form("searchname"))
    strSearchSql = Replace(strSearchName, "'", "''")
    blnPost = False
    If (SafeCInt(Request.Form("postform")) = 1) And (Len(strSearchName) >= 1) Then
      blnPost = True
      Dim rstUsers
      Set rstUsers = SQLQuery(cnnDB, "SELECT sid, uid, firstname, lastname, location1 FROM tblUsers WHERE " & _
        "uid LIKE '" & strSearchSql & "%' OR " & _
        "firstname LIKE '" & strSearchSql & "%' OR " & _
        "lastname LIKE '" & strSearchSql & "%' " & _
        "ORDER BY lastname ASC")
    End If
  %>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta charset="utf-8">
<title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "SelectUser")%></title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" type="text/css" href="../default.css">
<style>
  body { background: #F1F5F9; margin: 0; }
  .su-page { padding: 16px; max-width: 720px; margin: 0 auto; }
  .su-title { font-size: 17px; font-weight: 700; color: #1a202c; margin: 0 0 12px; }
  .su-search { display: flex; gap: 8px; }
  .su-search .form-control { flex: 1; }
  .su-filter { margin: 12px 0 0; }
  .su-empty { color: #718096; font-size: 13px; text-align: center; padding: 18px 0; }
</style>
<script language="JavaScript">
  function updateParent(formval, uid) {
     opener.document.newProbForm.uselectid.value = formval;
     var element = opener.document.getElementById("selectUserText");
     if (element.hasChildNodes())
       element.firstChild.nodeValue = uid;
     else
       element.appendChild(opener.document.createTextNode(uid));
     opener.document.newProbForm.uid.disabled = true;
     opener.document.newProbForm.uid.value = '';
     opener.document.newProbForm.uemail.disabled = true;
     opener.document.newProbForm.uemail.value = '';
     opener.document.newProbForm.department.disabled = true;
     opener.document.newProbForm.department.value = 0;
     opener.document.newProbForm.ulocation.disabled = true;
     opener.document.newProbForm.ulocation.value = '';
     opener.document.newProbForm.uphone.disabled = true;
     opener.document.newProbForm.uphone.value = '';
     self.close();
     return false;
  }
  function suFilter(q) {
    q = (q || '').toLowerCase();
    var rows = document.querySelectorAll('#suBody tr');
    for (var i = 0; i < rows.length; i++) {
      var t = (rows[i].textContent || rows[i].innerText || '').toLowerCase();
      rows[i].style.display = (t.indexOf(q) > -1) ? '' : 'none';
    }
  }
</script>
</head>
<body>

<%
  ' Check if user has permissions for this page
  Call CheckRep(cnnDB, sid)
%>
<div class="su-page">
  <div class="lhd-card">
    <div class="section-header">&#128269; <%=lang(cnnDB, "SelectUser")%></div>
    <form method="post" action="selectuser.asp" name="selectUserForm">
      <input type="hidden" name="postform" value="1">
      <div class="su-search">
        <input type="text" class="form-control" name="searchname" autofocus value="<% = Server.HTMLEncode(strSearchName) %>">
        <button type="submit" class="btn btn-primary"><%=lang(cnnDB, "Search")%></button>
      </div>
    </form>

    <% If blnPost Then
         If rstUsers.EOF Then %>
      <div class="su-empty"><%=lang(cnnDB, "Noresultsfound")%>.</div>
    <%   Else %>
      <input type="text" class="form-control su-filter" placeholder="&#128269; <%=lang(cnnDB, "Search")%>..." oninput="suFilter(this.value)">
      <div style="overflow-x:auto;margin-top:12px">
        <table class="lhd-table">
          <thead>
            <tr>
              <th><%=lang(cnnDB, "LastName")%></th>
              <th><%=lang(cnnDB, "FirstName")%></th>
              <th><%=lang(cnnDB, "UserName")%></th>
              <th><%=lang(cnnDB, "Location")%></th>
              <th>&nbsp;</th>
            </tr>
          </thead>
          <tbody id="suBody">
    <%     Do While Not rstUsers.EOF %>
            <tr>
              <td><% = Server.HTMLEncode(rstUsers("lastname") & "") %></td>
              <td><% = Server.HTMLEncode(rstUsers("firstname") & "") %></td>
              <td><% = Server.HTMLEncode(rstUsers("uid") & "") %></td>
              <td><% = Server.HTMLEncode(rstUsers("location1") & "") %></td>
              <td style="text-align:right"><button type="button" class="btn btn-primary btn-sm" onclick="return updateParent('<% = SafeCLng(rstUsers("sid")) %>', '<% = Server.HTMLEncode(LhdJsEsc(rstUsers("uid") & "")) %>')">Select</button></td>
            </tr>
    <%       rstUsers.MoveNext
           Loop %>
          </tbody>
        </table>
      </div>
    <%   End If
         rstUsers.Close
       End If %>
  </div>
</div>
<%
  cnnDB.Close
%>

</body>

</html>
