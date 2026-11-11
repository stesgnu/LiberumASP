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

  Filename: default.asp (inout)
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Main In/Out Board.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "InOutBoard")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>
  <%
    Call CheckUser(cnnDB, sid)
    Call DisplayHeader(cnnDB, sid)

    Dim inoutadmin
    inoutadmin = Usr(cnnDB, sid, "inoutadmin")

    Dim sqlWhere, orderby
    Dim mFirstname, mLastname, mUid, mDept, mInoutStatus, mPhone, check
    
    sqlWhere = "(tblUsers.department = departments.department_id)"
    sqlWhere = sqlWhere & " AND (ListOnInoutBoard = 1)"
    sqlWhere = sqlWhere & " AND (sid > 0)"
    
    If Request.Form("button") = lang(cnnDB, "Search") Then
    	mFirstname = CleanInputText(Request.Form("mFirstname"))
    	mLastname= CleanInputText(Request.Form("mLastname"))
    	mUid = CleanInputText(Request.Form("mUid"))
    	mDept = CleanInputText(Request.Form("mDept"))
    	mPhone = CleanInputText(Request.Form("mPhone"))
    	mInoutStatus = Request.Form("mInoutStatus")
    Else
    	mFirstname = ""
    	mLastname = ""
    	mUid = ""
    	mDept = ""
    	mPhone = ""
    	mInoutStatus = 0
    End If
  
    Dim intSort, intOrder, intFirstNameOrder, intLastNameOrder
    Dim intStatusOrder, intPhoneOrder, intUserIDOrder, intDeptOrder
    intSort = Cint(Request.QueryString("sort"))
    If Len(Request.QueryString("order")) > 0 Then
      intOrder = Cint(Request.QueryString("order"))
    Else
      intOrder = 0
    End If
    Select Case intSort
    	Case 1
        If intOrder = 0 Then
      	  orderby = "firstname ASC, lastname ASC"
          intFirstNameOrder = 1
        Else
      	  orderby = "firstname DESC, lastname DESC"
          intFirstNameOrder = 0
        End If
    	Case 2
        If intOrder = 0 Then
      		orderby = "lastname ASC, firstname ASC"
          intLastNameOrder = 1
        Else
      		orderby = "lastname DESC, firstname DESC"
          intLastNameOrder = 0
        End If
    	Case 3
        If intOrder = 0 Then
      	  orderby = "statuscode ASC, firstname ASC, lastname ASC"
          intStatusOrder = 1
        Else
      	  orderby = "statuscode DESC, firstname DESC, lastname DESC"
          intStatusOrder = 0
        End If
    	Case 4
    		orderby = "phone"
        If intOrder = 0 Then
          orderby = orderby & " ASC"
          intPhoneOrder = 1
        Else
          orderby = orderby & " DESC"
          intPhoneOrder = 0
        End If
    	Case 5
    		orderby = "uid"
        If intOrder = 0 Then
          orderby = orderby & " ASC"
          intUserIDOrder = 1
        Else
          orderby = orderby & " DESC"
          intUserIDOrder = 0
        End If
    	Case 6
        If intOrder = 0 Then
    		  orderby = "dname ASC, firstname ASC, lastname ASC"
          intDeptOrder = 1
        Else
    		  orderby = "dname DESC, firstname DESC, lastname DESC"
          intDeptOrder = 0
        End If
      Case Else
        orderby = "firstname ASC, lastname ASC"
        intFirstNameOrder = 1
    End Select
    
    If Len(mFirstname) > 0 Then
    	sqlWhere = sqlWhere & " AND (Firstname LIKE '" & Replace(mFirstname,"'","''") & "%')"
    End If
    If Len(mLastname) > 0 Then
    	sqlWhere = sqlWhere & " AND (Lastname LIKE '" & Replace(mLastname,"'","''") & "%')"
    End If
    If Len(mUid) > 0 Then
    	sqlWhere = sqlWhere & " AND (uid LIKE '" & Replace(mUid,"'","''") & "%')"
    End If
    If Len(mPhone) > 0 Then
    	sqlWhere = sqlWhere & " AND (phone LIKE '" & Replace(mPhone,"'","''") & "%')"
    End If
    If mInoutStatus = 1 Then
    	check = "checked"
    	sqlWhere = sqlWhere & " AND (statuscode >= 1)"
    End If
    If Len(mDept) > 0 Then
    	sqlWhere = sqlWhere & " AND (departments.dname LIKE '" & Replace(mDept,"'","''") & "%')"
    End If

    Dim strSql
    strSql = "SELECT tblUsers.*, departments.dname" &_
    	" FROM tblUsers, departments " &_
    	" WHERE " & sqlWhere &_
    	" ORDER BY " & orderby
  %>

  <div class="lhd-wrapper">

    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:12px">
      <div class="lhd-page-title" style="margin:0">
        <span class="title-icon">&#128101;</span>
        <%=lang(cnnDB, "InOutBoard")%>
      </div>
      <div>
        <a class="btn btn-secondary btn-sm" href="update.asp?id=<% = sid %>">&#9998; Mi Estado</a>
      </div>
    </div>

    <!-- Leyenda de estados -->
    <div style="display:flex;gap:20px;margin-bottom:14px;font-size:12px;color:var(--color-muted);flex-wrap:wrap;align-items:center">
      <div style="display:flex;align-items:center;gap:6px">
        <span class="inout-dot inout-in"></span>
        <span><%=lang(cnnDB, "In")%> (Disponible)</span>
      </div>
      <div style="display:flex;align-items:center;gap:6px">
        <span class="inout-dot inout-out"></span>
        <span><%=lang(cnnDB, "Out")%> (Ausente)</span>
      </div>
      <div style="display:flex;align-items:center;gap:6px">
        <span class="inout-dot inout-leave"></span>
        <span><%=lang(cnnDB, "Leave")%> (Permiso / Vacaciones)</span>
      </div>
    </div>

    <!-- Filtro rapido en cliente -->
    <div class="lhd-table-filter">
      <span style="font-size:13px;color:var(--color-muted)">&#128269;</span>
      <input type="text" id="tbl-filter" placeholder="Filtrar por nombre, usuario, dpto..." autocomplete="off">
    </div>

    <div class="lhd-card" style="padding:0;overflow:hidden">
      <div style="overflow-x:auto">
        <table class="lhd-table" id="inout-table">
          <thead>
            <tr>
              <th style="width:50px;text-align:center">Estado</th>
              <th>
                <a href="default.asp?sort=1&order=<% = intFirstNameOrder %>">
                  <%=lang(cnnDB, "FirstName")%> &#8597;
                </a>
              </th>
              <th>
                <a href="default.asp?sort=2&order=<% = intLastNameOrder %>">
                  <%=lang(cnnDB, "LastName")%> &#8597;
                </a>
              </th>
              <th>
                <a href="default.asp?sort=4&order=<% = intPhoneOrder %>">
                  <%=lang(cnnDB, "Phone")%> &#8597;
                </a>
              </th>
              <th>
                <a href="default.asp?sort=5&order=<% = intUserIDOrder %>">
                  <%=lang(cnnDB, "UserName")%> &#8597;
                </a>
              </th>
              <th>
                <a href="default.asp?sort=6&order=<% = intDeptOrder %>">
                  <%=lang(cnnDB, "Department")%> &#8597;
                </a>
              </th>
              <th class="col-actions">Acci&oacute;n</th>
            </tr>
          </thead>
          <tbody>
            <%
              Dim rstUser, counter
              counter = 0
              Set rstUser = SQLQuery(cnnDB, strSql)
              If Not rstUser.EOF And Request.Form("button") <> lang(cnnDB, "ClearForm") Then
                Do While Not rstUser.EOF
                  counter = counter + 1
            %>
                  <tr>
                    <td style="text-align:center">
                      <a href="details.asp?id=<% = rstUser("sid") %>" title="<% = Server.HTMLEncode(rstUser("statustext") & "") %>">
                        <% = InoutIcon(rstUser("statuscode")) %>
                      </a>
                    </td>
                    <td><strong><% = Server.HTMLEncode(rstUser("firstname") & "") %></strong></td>
                    <td><% = Server.HTMLEncode(rstUser("lastname") & "") %></td>
                    <td><% = Server.HTMLEncode(rstUser("phone") & "") %></td>
                    <td><% = Server.HTMLEncode(rstUser("uid") & "") %></td>
                    <td><% = Server.HTMLEncode(rstUser("dname") & "") %></td>
                    <td class="col-actions">
                      <div style="display:flex;gap:6px;justify-content:center">
                        <a class="btn btn-secondary btn-sm" href="details.asp?id=<% = rstUser("sid") %>" title="Ver detalles">&#128065;</a>
                        <% If inoutadmin = 1 Or (sid = rstUser("sid")) Then %>
                          <a class="btn btn-primary btn-sm" href="update.asp?id=<% = rstUser("sid") %>" title="Editar estado">&#9998;</a>
                        <% End If %>
                      </div>
                    </td>
                  </tr>
            <%
                  rstUser.MoveNext
                Loop
              Else
            %>
                <tr>
                  <td colspan="7" style="text-align:center;padding:32px;color:var(--color-muted)">
                    <%=lang(cnnDB, "Noresultsfound")%>.
                  </td>
                </tr>
            <%
              End If
              rstUser.Close
            %>
          </tbody>
        </table>
      </div>
      <% If counter > 0 Then %>
        <div style="padding:12px 18px;font-size:12px;color:var(--color-muted);border-top:1px solid var(--color-border-light)">
          <% = counter %> &nbsp;<%=lang(cnnDB, "recordsfound")%>.
        </div>
      <% End If %>
    </div>

  </div>

  <script>
  document.getElementById('tbl-filter').addEventListener('input', function() {
    var q = this.value.toLowerCase();
    var rows = document.querySelectorAll('#inout-table tbody tr');
    rows.forEach(function(row) {
      row.style.display = row.textContent.toLowerCase().indexOf(q) >= 0 ? '' : 'none';
    });
  });
  </script>

  <%
    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close
  %>
  </body>
</html>
