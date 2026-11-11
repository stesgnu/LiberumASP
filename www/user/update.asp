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
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Updates the notes in the database.
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
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemUpdated")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style>
      .uu-card { max-width: 560px; margin: 40px auto; text-align: center; padding: 36px 32px; }
      .uu-icon { width: 64px; height: 64px; border-radius: 50%; background: #DCFCE7; color: #16A34A; font-size: 32px; display: inline-flex; align-items: center; justify-content: center; margin-bottom: 14px; }
      .uu-card h2 { font-size: 20px; margin: 0 0 8px; color: #1a202c; }
      .uu-card p { font-size: 14px; color: #4a5568; margin: 0 0 22px; }
      .uu-actions { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; }
    </style>
  </head>
  <body>
    <%
      ' See if authenticated
      Call CheckUser(cnnDB, sid)

      ' Get the username, problem id and the additional
      ' notes that will be entered.
      Dim uid, id, notes

      uid = 	Usr(cnnDB, sid, "uid")
      id = SafeCLng(Request.Form("id"))
      notes = CleanInputText(Request.Form("notes"))

    ' Check for required fields (uemail, category, department, title, description)

      if Len(notes)=0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "AdditionalNotes"))
      End if


    ' Get old description
      Dim description, rep, rstOldRep, rstUpdateNotes, dtNoteDate
      ' Solo se pueden agregar notas a tickets propios
      Set rstOldRep = SQLquery(cnnDB, "SELECT rep FROM problems WHERE id=" & id & " AND uid='" & Replace(uid & "", "'", "''") & "'")
      If rstOldRep.EOF Then
        rstOldRep.Close
        cnnDB.Close
        Response.Clear
        Response.Status = "403 Forbidden"
        Response.Write "Acceso denegado."
        Response.End
      End If
      rep = rstOldRep("rep")

      dtNoteDate = SQLDate(Now, lhdAddSQLDelim)

    ' Update description, making sure only valid characters are in the description
      notes = Replace(notes, "'", "''")

      Set rstUpdateNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, uid, private) " & _
        "VALUES (" & id & ", '" &  notes & "', " & dtNoteDate & ", '" & _
        Usr(cnnDB, sid, "uid") & "', 0)")

    Call eMessage(cnnDB, "repupdate", id, Usr(cnnDB, rep, "email1"))
    Call LhdToast("success", lang(cnnDB, "Problem") & " " & id & " " & lang(cnnDB, "isUpdated"))

      Call DisplayHeader(cnnDB, sid)
    %>

    <div class="lhd-wrapper">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0"><span class="title-icon">&#128172;</span> <%=lang(cnnDB, "ProblemUpdated")%></div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="view.asp">&larr; Volver</a>
        </div>
      </div>

      <div class="lhd-card uu-card">
        <div class="uu-icon">&#10004;</div>
        <h2><%=lang(cnnDB, "Problem")%>&nbsp;#<% = id %>&nbsp;<%=lang(cnnDB, "isUpdated")%></h2>
        <p><%=lang(cnnDB, "Viewthedetailsof")%> <%=lang(cnnDB, "Problem")%>&nbsp;#<% = id %>.</p>
        <div class="uu-actions">
          <a class="btn btn-primary" href="details.asp?id=<% = id %>"><%=lang(cnnDB, "Problem")%>&nbsp;#<% = id %> &rarr;</a>
          <a class="btn btn-secondary" href="view.asp">&larr; Volver</a>
        </div>
      </div>
    </div>

    <%
      ' Close records
      rstOldRep.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>

  </body>
</html>
