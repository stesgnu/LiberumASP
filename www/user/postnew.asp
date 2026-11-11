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

  Filename: postnew.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.51.4.1 $
  Purpose:  Takes input from new.asp, validates, inserts problem, shows modern feedback.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemSubmitted")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' See if authenticated
      Call CheckUser(cnnDB, sid)

      ' Get form fields
      Dim uid, uemail, uphone, ulocation, category, department, title, description, entered_by
      Dim kb, duedate, priority
      
      uid = CleanInputText(Request.Form("uid"))
      uemail = CleanInputText(Request.Form("uemail"))
      uphone = CleanInputText(Request.Form("uphone"))
      ulocation = CleanInputText(Request.Form("ulocation"))
      category = Cint(Request.Form("category"))
      department = SafeCLng(Request.Form("department"))
      title = CleanInputText(Request.Form("title"))
      description = LhdSanitizeHtml(CleanInputText(Request.Form("description")))
      Dim attachToken
      attachToken = Request.Form("attachToken")
      entered_by = sid
      kb = 0
      duedate = ConvertFormattedDate(Request.Form("duedate"))
      priority = Cint(Request.Form("priority"))

      ' Validation
      If InStr(uemail, "@") = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Emailaddress"))
      End If

      If category = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Category"))
      End If

      If department = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Department"))
      End If

      If priority = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Priority"))
      End If

      If Not IsDate(duedate) Then
        duedate = DateAdd("d", 1, Now())
      End If

      If Len(title) = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Title"))
      ElseIf Len(title) > 255 Then
        title = Trim(title)
        title = Left(title, 255)
      End If

      If Not LhdRichHasContent(description) Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Description"))
      End If

      Dim id, status, start_date, rep, time_spent
      status = Cfg(cnnDB, "DefaultStatus")
      time_spent = 0
      start_date = Now

      Dim dname, rstDept
      Set rstDept = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & department)
      dname = rstDept("dname")

      Dim cname, catRes
      Set catRes = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & category)
      rep = catRes("rep_id")
      cname = catRes("cname")

      Dim pname, priRes
      Set priRes = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & priority)
      pname = priRes("pname")
      priRes.Close

      id = GetUnique(cnnDB, "problems")

      uemail = Left(Trim(uemail), 50)
      uemail = Replace(uemail, "'", "''")
      uphone = Left(Trim(uphone), 50)
      uphone = Replace(uphone, "'", "''")
      ulocation = Left(Trim(ulocation), 50)
      ulocation = Replace(ulocation, "'", "''")
      title = Replace(title, "'", "''")
      description = Replace(description, "'", "''")

      Dim strProblemQry, rstProbInsert
      strProblemQry = "INSERT INTO problems (id, uid, uemail, uphone, ulocation, entered_by, " & _
        "category, department, title, description, priority, status, start_date, rep, time_spent, kb) " & _
        "VALUES (" & id & ",'" & uid & "','" & uemail & "','" & uphone & "','" & _
        ulocation & "'," & entered_by & "," & category & "," & department & ",'" & title & "','" & _
        description & "'," & priority & "," & status & "," & SQLDate(start_date, lhdAddSQLDelim) & "," & rep & "," & time_spent & _
        "," & kb & ")"

      Set rstProbInsert = SQLQuery(cnnDB, strProblemQry)
      Call LhdLinkAttachments(cnnDB, attachToken, id)

      ' Notifications
      Call eMessage(cnnDB, "usernew", id, uemail)
      Call eMessage(cnnDB, "repnew", id, Usr(cnnDB, rep, "email1"))

      If (priority >= Cfg(cnnDB, "EnablePager")) And (Len(Usr(cnnDB, rep, "email2")) > 0) Then
        Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
      End If

      title = Replace(title, "''", "'")
      description = Replace(description, "''", "'")

      Call LhdToast("success", lang(cnnDB, "Problem") & " #" & id & " " & lang(cnnDB, "Submitted"))
      Call DisplayHeader(cnnDB, sid)
    %>

    <div class="lhd-wrapper">

      <div class="lhd-card" style="text-align:center;padding:36px;margin-bottom:24px">
        <div style="font-size:48px;color:var(--color-success);margin-bottom:12px">&#10004;</div>
        <h2 style="margin:0 0 8px"><%=lang(cnnDB, "Problem")%>&nbsp;#<% = id %>&nbsp;<%=lang(cnnDB, "Submitted")%></h2>
        <p style="color:var(--color-muted);font-size:14px;max-width:500px;margin:0 auto 20px">
          Tu solicitud fue registrada con &eacute;xito y asignada a nuestro equipo de soporte.
        </p>
        <div style="display:flex;gap:12px;justify-content:center;flex-wrap:wrap">
          <a class="btn btn-primary" href="details.asp?id=<% = id %>">
            &#128065; Ver Detalles del Ticket
          </a>
          <a class="btn btn-secondary" href="new.asp">
            &#10133; Enviar Otro Problema
          </a>
          <a class="btn btn-secondary" href="view.asp">
            &#128203; Mis Tickets
          </a>
        </div>
      </div>

      <div class="lhd-detail-layout">
        <div class="lhd-card">
          <div class="section-header">&#128221; <%=lang(cnnDB, "ProblemInformation")%></div>
          <div class="form-group">
            <label class="form-label"><%=lang(cnnDB, "Title")%></label>
            <input type="text" class="form-control" value="<% = Server.HTMLEncode(title) %>" readonly>
          </div>
          <div class="form-group" style="margin-bottom:0">
            <label class="form-label"><%=lang(cnnDB, "Description")%></label>
            <div style="background:#F7FAFC;border:1px solid var(--color-border-light);border-radius:var(--radius-sm);padding:14px;font-size:13px;line-height:1.7;white-space:pre-wrap">
              <% = LhdRichToDisplay(description) %>
            </div>
          </div>
        </div>

        <div class="lhd-detail-meta">
          <div class="lhd-detail-meta-header">&#8505; Resumen del Ticket</div>
          <table>
            <tr>
              <td>Ticket ID</td>
              <td><strong>#<% = id %></strong></td>
            </tr>
            <tr>
              <td>Prioridad</td>
              <td><% = PriorityBadge(pname) %></td>
            </tr>
            <tr>
              <td>Departamento</td>
              <td><% = Server.HTMLEncode(dname) %></td>
            </tr>
            <tr>
              <td>Categor&iacute;a</td>
              <td><% = Server.HTMLEncode(cname) %></td>
            </tr>
            <tr>
              <td>Asignado a</td>
              <td><% = Server.HTMLEncode(Usr(cnnDB, rep, "fname")) %></td>
            </tr>
            <tr>
              <td>Fecha ingreso</td>
              <td><% = DisplayDate(start_date, lhdDateTime) %></td>
            </tr>
          </table>
        </div>
      </div>

    </div>

    <%
      rstDept.Close
      catRes.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>

  </body>
</html>
