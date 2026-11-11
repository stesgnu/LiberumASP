<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
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

  Filename: details.asp
  Date:     $Date: 2002/08/28 15:31:15 $
  Version:  $Revision: 1.50.4.2 $
  Purpose:  This page displays the problem details and allows reps to
            modify them.
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
  <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "EditProblem")%></title>
  <link rel="stylesheet" type="text/css" href="../default.css">
<link rel="stylesheet" type="text/css" href="https://cdnjs.cloudflare.com/ajax/libs/quill/1.3.7/quill.snow.css" />
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/quill/1.3.7/quill.min.js"></script>
<link rel="stylesheet" type="text/css" href="../js/dropzone/dropzone.min.css?v=1">
<script type="text/javascript" src="../js/dropzone/dropzone.min.js?v=1"></script>
<script type="text/javascript" src="../js/lhd-editor.js?v=4"></script>
</head>
<body>

<%
	' Check if user has permissions for this page (unless using reopen link)
	If Cint(Request.QueryString("reopen")) <> 1 Then
	  Call CheckRep(cnnDB, sid)
	End if

	' Get the problem ID
	Dim id, blnUpdate, strUpdateMessage
	id = Cint(Request.QueryString("id"))
  If Cint(Request.Form("update")) = 1 Then
    blnUpdate = True
  Else 
    blnUpdate = False
  End If

  Dim attachToken
  attachToken = Request.Form("attachToken")
  If Len(attachToken) = 0 Then attachToken = "d" & sid & "x" & CLng(Rnd * 1000000000) & "x" & CLng(Timer * 1000)
  Dim uid, uemail, uphone, ulocation, category, department, title, description, kb
  Dim priority, status, rep, time_spent, solution, notes, clost_date, start_date, close_date, duedate
  Dim blnNotifyRep

  ' ===============================
  ' Update the problem
  If blnUpdate Then
    Dim oldrep, blnFirstResponse

    ' Get the problem data from the form fields.
    id = Cint(Request.Form("id"))
    uid = CleanInputText(Request.Form("uid"))
    uemail = CleanInputText(Request.Form("uemail"))
    uphone = CleanInputText(Request.Form("uphone"))
    ulocation = CleanInputText(Request.Form("ulocation"))
    category = Cint(Request.Form("category"))
    department = Cint(Request.Form("department"))
    title = CleanInputText(Request.Form("title"))
    priority = Cint(Request.Form("priority"))
    status = Cint(Request.Form("status"))
    rep = Cint(Request.Form("rep"))
    oldrep = Cint(Request.Form("oldrep"))
    time_spent = SafeCLng(Request.Form("time_spent"))
    solution = LhdSanitizeHtml(CleanInputText(Request.Form("solution")))
    notes = CleanInputText(Request.Form("notes"))
    duedate = ConvertFormattedDate(Request.Form("duedate"))

    If Request.Form("kb") = "on" Then
      kb = 1
    Else
      kb = 0
    End If
	

	
  ' Check for required fields (uemail, category, department, title, description)

    if Len(uid)=0 Then
      Call DisplayError(1, Lang(cnnDB, "UserName"))
    End if

    if (uemail = Cfg(cnnDB, "BaseEmail")) OR (InStr(uemail, "@")=0) Then
      Call DisplayError(1, Lang(cnnDB, "EMailAddress"))
    End if

    if Len(title)=0 Then
      Call DisplayError(1, Lang(cnnDB, "Title"))
    End if

    If Not IsDate(duedate) Then
      duedate = DateAdd("d", 1, Now())
    End If

    if (status >= 100) and (Not LhdRichHasContent(solution)) Then
      Call DisplayError(1, Lang(cnnDB, "Solution"))
    End if
    
  ' Clean up fields
    uemail = Left(Trim(uemail), 50)
    uemail = Replace(uemail, "'", "''")
    uphone = Left(Trim(uphone), 50)
    uphone = Replace(uphone, "'", "''")
    ulocation = Left(Trim(ulocation), 50)
    ulocation = Replace(ulocation, "'", "''")
    title = Left(Trim(title), 50)
    title = Replace(title, "'", "''")

    time_spent = Trim(time_spent)
    If IsNumeric(time_spent) Then
      time_spent = Cint(time_spent)
    Else
      time_spent = 0
    End If

    If (Len(notes) > 0 And Request.Form("hidenotes") <> "on") Or (status >= 100) Then
      blnFirstResponse = True
    Else
      blnFirstResponse = False
    End If

    If (Len(notes) > 0) And (sid <> rep) And (rep = oldrep) Then
      blnNotifyRep = True
    Else
      blnNotifyRep = False
    End If


  ' Grab original description
    Dim rstDesc, strChangeNotes
    Set rstDesc = SQLQuery(cnnDB, "SELECT category, department, rep, status, priority, first_response FROM problems WHERE id=" & id)

  ' Insert actions into strChangeNotes
    If (category <> rstDesc("category")) OR (department <> rstDesc("department")) OR _
      (rep <> rstDesc("rep")) OR (status <> rstDesc("status")) OR _
      (priority <> rstDesc("priority")) Then

        If priority <> rstDesc("priority") Then
          Dim newPri, oldPri
          Set newPri = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & priority)
          Set oldPri = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & rstDesc("priority"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "PRIORITY_2") & ": " & oldPri("pname") & " => " & newPri("pname") & vbNewLine
          newPri.Close
          oldPri.Close
        End If

        If rep <> rstDesc("rep") Then
          Dim newRep, oldRep2
          Set newRep = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & rep)
          Set oldRep2 = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & rstDesc("rep"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "TRANSFERREPS") & ": " & oldRep2("uid") & " => " & newRep("uid") & vbNewLine
          newRep.Close
          oldRep2.Close
        End If

        If category <> rstDesc("category") Then
          Dim newCat, oldCat
          Set newCat = SQLQuery(cnnDB, "SELECT cname FROM categories WHERE category_id=" & category)
          Set oldCat = SQLQuery(cnnDB, "SELECT cname FROM categories WHERE category_id=" & rstDesc("category"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "CATEGORY_2") & ": " & oldCat("cname") & " => " & newCat("cname") & vbNewLine
          newCat.Close
          oldCat.Close
        End If

        if department <> rstDesc("department") Then
          Dim newDep, oldDep
          Set newDep = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & department)
          Set oldDep = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & rstDesc("department"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "DEPARTMENT_2") & ": " & oldDep("dname") & " => " & newDep("dname") & vbNewLine
          newDep.Close
          oldDep.Close
        End If

        If status <> rstDesc("status") Then
          Dim newStat, oldStat
          Set newStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & status)
          Set oldStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & rstDesc("status"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "STATUS_2") & ": " & oldStat("sname") & " => " & newStat("sname") & vbNewLine
          newStat.Close
          oldStat.Close
        End If
    End If


    ' Update the Notes
    Dim rstUpdateNotes, intPrivate, dtNoteDate, blnSendUpdateMsg
    blnSendUpdateMsg = False

    dtNoteDate = SQLDate(Now, lhdAddSQLDelim)

    If Len(notes)>0 Then
      notes = Replace(notes, "'", "''")
      If Request.Form("hidenotes") = "on" Then
        intPrivate = 1
      Else
        intPrivate = 0
        blnSendUpdateMsg = True
      End If    
      Set rstUpdateNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, uid, private) " & _
        "VALUES (" & id & ", '" &  notes & "', " & dtNoteDate & ", '" & _
        Usr(cnnDB, sid, "uid") & "', " & intPrivate & ")")
    End If

    If Len(strChangeNotes)> 0 Then
      Set rstUpdateNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, uid, private) " & _
        "VALUES (" & id & ", '" &  strChangeNotes & "', " & dtNoteDate & ", '" & _
        Usr(cnnDB, sid, "uid") & "', 1)")
    End If

  ' Get missing variables to enter problem

    Dim cname, catRes
    Set catRes = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & category)
    cname = catRes("cname")

    Dim probRes
    Set probRes = SQLQuery(cnnDB, "SELECT start_date FROM problems WHERE id=" & id)
    start_date = probRes("start_date")

  ' Get the old priority
    Dim old_priority, oldPriRes
    Set oldPriRes = SQLQuery(cnnDB, "SELECT priority FROM problems WHERE id=" & id)
    old_priority = oldPriRes("priority")
    oldPriRes.Close

  ' Remove apostrophes
    description = Replace(description, "'", "''")
    solution = Replace(solution, "'", "''")

  ' All data is present
  ' Write problem into database

    Dim probStr
    probStr = "UPDATE problems SET " & _
      "uid='" & uid & "', " & _
      "uemail='" & uemail & "', " & _
      "uphone='" & uphone & "', " & _
      "ulocation='" & ulocation & "', " & _
      "category=" & category & ", " & _
      "department=" & department & ", " & _
      "title='" & title & "', " & _
      "priority=" & priority & ", " & _
      "status=" & status & ", " & _
      "rep=" & rep & ", " & _
      "kb=" & kb & ", " & _
      "time_spent=" & time_spent & ", " & _
      "solution='" & solution & "'"

    ' Add the closed date/time if the problem is closed
    If status >= 100 Then
      probStr = probStr & ", close_date=" & SQLDate(Now, lhdAddSQLDelim)
	  If (Request.Form("noemail")<>"on") Then
	    'Set email send field to true
		probStr = probStr & ", emailsent=1"
	  End If
    End If
    strUpdateMessage = Lang(cnnDB, "Theproblemhasbeensaved") & "."

    probStr = probStr & " WHERE id=" & id

    Set probRes = SQLQuery(cnnDB, probStr)
    Call LhdLinkAttachments(cnnDB, attachToken, id)

    If blnFirstResponse And Not IsDate(rstDesc("first_response")) Then
      Dim firstResponseRes
      Set firstResponseRes = SQLQuery(cnnDB, "UPDATE problems SET first_response=" & SQLDate(Now, lhdAddSQLDelim) & " WHERE id = " & id)
      Set firstResponseRes = Nothing
    End If
    rstDesc.Close
    Set rstDesc = Nothing

    If status >= 100 Then
      If Not (Request.Form("noemail")="on") Then
      ' Send mail to the user'
        Call eMessage(cnnDB, "userclose", id, uemail)
      End If
    Else

      ' Notify user of update
      If (Cfg(cnnDB, "Notifyuser") = 1) And Not (Request.Form("noemail")="on") And blnSendUpdateMsg Then
        Call eMessage(cnnDB, "userupdate", id, uemail)
      End If

      'Notify rep if problem updated by another rep
      If blnNotifyRep Then
        Call eMessage(cnnDB, "repupdate", id, Usr(cnnDB, rep, "email1"))
      End If

      'Send mail to the appropriate rep for transfered problems
      If (rep <> oldrep) Then
        Call eMessage(cnnDB, "repnew", id, Usr(cnnDB, rep, "email1"))

        'Page Rep if enabled
        If (priority >= Cfg(cnnDB, "EnablePager")) And (Len(Usr(cnnDB, rep, "email2")) > 0) Then
          Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
        End If

      ElseIf priority <> old_priority Then
        'Page Rep if enabled
        If (priority >= Cfg(cnnDB, "EnablePager")) And (Len(Usr(cnnDB, rep, "email2")) > 0) Then
          Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
        End If
      End If
    End If
  End If

  ' ===============================
  
  If Cint(Request.QueryString("reopen")) = 1 Then
    Dim strSQLOpen, rstOpenProbUpd, rstOpenNotes, dtOpenNoteDate, strOpenNote
    Dim rstOpenOldStat, rstOpenNewStat, rstOpenUserEmail

    strSQLOpen = "UPDATE problems SET " & _
      "status = " & Cfg(cnnDB, "DefaultStatus") & ", " & _
      "close_date = NULL"
	If Usr(cnnDB, sid, "IsRep")=1 Then  ' If rep, then reassign
	  strSQLOpen = strSQLOpen & ", rep = " & sid 
	End If
    strSQLOpen = strSQLOpen & " WHERE id = " & id
    Set rstOpenProbUpd = SQLQuery(cnnDB, strSQLOpen)
    Set rstOpenNewStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & Cfg(cnnDB, "DefaultStatus"))
    Set rstOpenOldStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & Cfg(cnnDB, "CloseStatus"))
    strOpenNote = Lang(cnnDB, "STATUS_2") & ": " & rstOpenOldStat("sname") & " => " & rstOpenNewStat("sname") & vbNewLine
    dtOpenNoteDate = SQLDate(Now, lhdAddSQLDelim)
    Set rstOpenNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, uid, private) " & _
        "VALUES (" & id & ", '" &  strOpenNote & "', " & dtOpenNoteDate & ", '" & _
        Usr(cnnDB, sid, "uid") & "', 1)")
    Set rstOpenUserEmail = SQLQuery(cnnDB, "SELECT uemail FROM problems WHERE id = " & id)
    Call eMessage(cnnDB, "usernew", id, rstOpenUserEmail("uemail"))
    rstOpenUserEmail.Close
    rstOpenNewStat.Close
    rstOpenOldStat.Close
	
	' If not a support rep, redirect to the user details page
	If Usr(cnnDB, sid, "IsRep")=0 Then
	  Dim repRes, queryStr
	  queryStr = "SELECT r.email1 FROM (problems As p INNER JOIN tblUsers As r On r.sid = p.rep) WHERE p.id = " & id
	  Set repRes = SQLQuery(cnnDB, queryStr)
	  Call eMessage(cnnDB, "repupdate", id, repRes("email1"))
	  repRes.Close
	  Response.Clear()
	  Call LhdToast("success", "Ticket #" & id & " reabierto")
		  Response.Redirect("../user/details.asp?id=" & id)
	End if

  End If

  ' ===============================

	' Query the database for the problem info
  Dim rstProb, rstSol, rstNotes, entered_by, strProbQuery
  strProbQuery = "SELECT uid, uemail, uphone, ulocation, time_spent, department, " & _
		"category, status, priority, entered_by, rep, kb, start_date, close_date, title, description " & _
		"FROM problems WHERE id=" & id

  ' Make sure rep is viewing own problems.
  If Usr(cnnDB, sid, "RepAccess")=1 Then
    strProbQuery = strProbQuery & " AND rep=" & sid
  End If
  
  Set rstProb = SQLQuery(cnnDB, strProbQuery)
  
  If rstProb.EOF Then
    Call DisplayError(3, "Problem " & id & " could not be found in the database.")
  End If

	' Query for the solution seperately becuase SQL only
	' supports 1 blob per query
	Set rstSol = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id)

	uid = rstProb("uid")
	uemail = rstProb("uemail")
	uphone = rstProb("uphone")
	ulocation = rstProb("ulocation")
	time_spent = Cint(rstProb("time_spent"))
	department = Cint(rstProb("department"))
	category = Cint(rstProb("category"))
	status = Cint(rstProb("status"))
	priority = Cint(rstProb("priority"))
	rep = Cint(rstProb("rep"))
  kb = Cint(rstProb("kb"))
  entered_by = Cint(rstProb("entered_by"))
	start_date = rstProb("start_date")
	close_date = rstProb("close_date")
  duedate = DateAdd("d", 1, Now())
	title = rstProb("title")
	description = rstProb("description")

  ' Get the Notes for this problem
  ' Nombres de prioridad y estado para los badges del encabezado
  Dim pname, sname, rstBadgePri, rstBadgeSta
  pname = ""
  sname = ""
  Set rstBadgePri = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & priority)
  If Not rstBadgePri.EOF Then pname = rstBadgePri("pname")
  rstBadgePri.Close
  Set rstBadgeSta = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & status)
  If Not rstBadgeSta.EOF Then sname = rstBadgeSta("sname")
  rstBadgeSta.Close

  Set rstNotes = SQLQuery(cnnDB, "SELECT id, addDate, uid, private, [note] FROM tblNotes WHERE id=" & id & " ORDER BY addDate ASC")
  ' Get the solution and replace characters to make
	' it more readable.
	solution = rstSol("solution")

  ' If the rep has readonly access, or the problem is closed
  ' disable the fields
  Dim strTextDisable, strListDisable
  If Usr(cnnDB, sid, "RepAccess")=2 Or status >= 100 Then
    strTextDisable = "readonly"
    strListDisable = "disabled"
  Else
    strTextDisable = ""
    strListDisable = ""
  End If
%>


<%
  Call DisplayHeader(cnnDB, sid)
%>
<div class="lhd-wrapper">

  <!-- Breadcrumb y Cabecera de la Incidencia -->
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:12px">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <a class="btn btn-secondary btn-sm" href="view.asp">&larr; Volver a la Lista</a>
      <h2 style="margin:0;font-size:20px;font-weight:700;color:#1a202c">
        Incidencia #<% = id %> &mdash; <% = FixTildes(Server.HTMLEncode(title)) %>
      </h2>
    </div>
    <div style="display:flex;align-items:center;gap:8px">
      <% = PriorityBadge(pname) %>
      <% = StatusBadge(sname) %>
      <a class="btn btn-secondary btn-sm" href="print.asp?id=<% = id %>" target="_blank">&#128424; Imprimir</a>
      <% If status >= 100 Then %>
        <a class="btn btn-secondary btn-sm" href="details.asp?id=<% = id %>&reopen=1">&#8635; Reabrir Ticket</a>
      <% End If %>
    </div>
  </div>

  <% If blnUpdate Then Call LhdToast("success", FixTildes(strUpdateMessage)) %>

  <form action="details.asp" method="POST" id="frmTicketDetails">
    <input type="hidden" name="uid" value="<% = Server.HTMLEncode(uid) %>">
    <input type="hidden" name="id" value="<% = id %>">
    <input type="hidden" name="oldrep" value="<% = rep %>">
    <input type="hidden" name="update" value="1">

    <div style="display:grid;grid-template-columns:2fr 1fr;gap:20px;align-items:start">

      <!-- COLUMNA IZQUIERDA (Principal) -->
      <div>

        <!-- Tarjeta: Detalle del Problema -->
        <div class="lhd-card">
          <div class="section-header">&#128196; Detalle de la Solicitud</div>

          <div class="form-group" style="margin-bottom:16px">
            <label class="form-label"><b><%=lang(cnnDB, "Title")%>:</b> <em>*</em></label>
            <input type="text" name="title" class="form-control" value="<% = FixTildes(Server.HTMLEncode(title)) %>" <% = strTextDisable %> required>
          </div>

          <div class="form-group">
            <label class="form-label"><b><%=lang(cnnDB, "Description")%>:</b></label>
            <div class="lhd-rich" style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:14px;line-height:1.6;color:#2d3748;max-height:420px;overflow-y:auto">
              <% = LhdRichToDisplay(description) %>
            </div>
            <% = LhdAttachmentList(cnnDB, id) %>
          </div>
        </div>

        <!-- Tarjeta: Historial de Notas -->
        <div class="lhd-card">
          <div class="section-header">&#128483; Historial de Notas y Seguimiento</div>

          <% If rstNotes.EOF Then %>
            <div style="text-align:center;padding:20px;color:#a0aec0;font-style:italic">
              <%=lang(cnnDB, "NoAvailableNotes")%>
            </div>
          <% Else %>
            <div style="display:flex;gap:8px;justify-content:flex-end;margin-bottom:10px">
              <button type="button" class="btn btn-secondary btn-sm" onclick="lhdNotesToggle(true)">Expandir todas</button>
              <button type="button" class="btn btn-secondary btn-sm" onclick="lhdNotesToggle(false)">Contraer todas</button>
            </div>
            <div style="display:flex;flex-direction:column;gap:10px">
              <% Do While Not rstNotes.EOF %>
                <details class="lhd-note" style="background:#f8fafc;border-left:4px solid <% If rstNotes("private") = 1 Then %>#D69E2E<% Else %>var(--color-primary,#0057B8)<% End If %>;border-radius:0 6px 6px 0">
                  <summary style="cursor:pointer;padding:10px 16px;font-size:13px;color:#2d3748">
                    <strong>&#128100; <% = Server.HTMLEncode(NoteAuthorName(cnnDB, rstNotes("uid"))) %></strong>
                    &bull; <% = DisplayDate(rstNotes("addDate"), lhdDateTime) %>
                  </summary>
                  <div style="padding:0 16px 12px 16px">
                    <% If rstNotes("private") = 1 Then %>
                      <span class="badge" style="background:#FEFCBF;color:#744210;font-size:10px;display:inline-block;margin-bottom:6px">&#128274; NOTA INTERNA</span>
                    <% End If %>
                    <div style="color:#2d3748;line-height:1.5;white-space:pre-wrap"><% = FixTildes(Server.HTMLEncode(rstNotes("note"))) %></div>
                  </div>
                </details>
              <%
                  rstNotes.MoveNext
                 Loop
              %>
            </div>
            <script type="text/javascript">
              function lhdNotesToggle(abrir) {
                var items = document.querySelectorAll('details.lhd-note');
                for (var i = 0; i < items.length; i++) { items[i].open = abrir; }
              }
            </script>
          <% End If %>
        </div>

        <!-- Tarjeta: Agregar Nota (solo si ticket no esta cerrado) -->
        <% If status < 100 Then %>
          <div class="lhd-card">
            <div class="section-header">&#10133; <%=lang(cnnDB, "EnterAdditionalNotes")%></div>
            <div class="form-group" style="margin-bottom:12px">
              <textarea rows="4" class="form-control" name="notes" placeholder="Escribe aqu&iacute; notas de seguimiento o avances..." <% = strTextDisable %>></textarea>
            </div>
            <label style="display:inline-flex;align-items:center;gap:6px;font-size:13px;cursor:pointer">
              <input type="checkbox" name="hidenotes" <% = strListDisable %>>
              <%=lang(cnnDB, "HideFromEndUser")%> (Nota solo visible para t&eacute;cnicos)
            </label>
          </div>
        <% End If %>

        <!-- Tarjeta: Solucion -->
        <div class="lhd-card">
          <div class="section-header">&#9989; <%=lang(cnnDB, "Solution")%></div>
          <div class="form-group" style="margin-bottom:12px">
            <% If strTextDisable = "" Then %>
            <div id="solEditor" style="background:#fff;min-height:140px"></div>
            <% Else %>
            <div class="lhd-rich" style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:14px;line-height:1.6">
              <% = LhdRichToDisplay(solution) %>
            </div>
            <% End If %>
            <textarea name="solution" id="solField" style="display:none"><% = FixTildes(Server.HTMLEncode(solution)) %></textarea>
            <% If strTextDisable = "" Then %>
            <div class="form-group" style="margin-top:12px">
              <label class="form-label"><b>Agregar archivos:</b></label>
              <input type="file" id="attachFile" multiple class="form-control" accept="<% = LhdAcceptAttr(cnnDB) %>" />
              <div id="attachList" style="margin-top:8px"></div>
              <input type="hidden" name="attachIds" id="attachIds" value="" />
              <input type="hidden" name="attachToken" value="<% = attachToken %>" />
            </div>
            <% Else %>
            <input type="hidden" name="attachToken" value="<% = attachToken %>" />
            <% End If %>
          </div>
          <% If Cfg(cnnDB, "EnableKB") <> 0 Then %>
            <label style="display:inline-flex;align-items:center;gap:6px;font-size:13px;cursor:pointer">
              <input type="checkbox" name="kb" <% If kb = 1 Then %>checked<% End If %> <% = strListDisable %>>
              <%=lang(cnnDB, "EnterinKnowledgeBase")%> (Publicar soluci&oacute;n en la Base de Conocimiento)
            </label>
          <% End If %>
        </div>

      </div>

      <!-- COLUMNA DERECHA (Metadatos y Acciones) -->
      <div>

        <!-- Tarjeta: Acciones de Guardado -->
        <% If status < 100 And Usr(cnnDB, sid, "RepAccess") <> 2 Then %>
          <div class="lhd-card" style="border:2px solid var(--color-primary,#0057B8)">
            <div style="font-weight:700;font-size:14px;color:var(--color-primary,#0057B8);margin-bottom:12px">&#9889; Acciones</div>
            <% If Cfg(cnnDB, "EmailType") <> 0 Then %>
              <div style="margin-bottom:12px">
                <label style="display:inline-flex;align-items:center;gap:6px;font-size:12px;color:#4a5568;cursor:pointer">
                  <input type="checkbox" name="noemail" id="noEmailCheck">
                  <%=lang(cnnDB, "Dontsendemailtouser")%>
                </label>
              </div>
            <% End If %>
            <button type="submit" class="btn btn-primary" style="width:100%;justify-content:center;padding:12px;font-size:14px">
              &#128190; <%=lang(cnnDB, "SaveProblem")%>
            </button>
          </div>
        <% End If %>

        <!-- Tarjeta: Gestion de la Incidencia -->
        <div class="lhd-card">
          <div class="section-header">&#9881; <%=lang(cnnDB, "Classification")%></div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "Status")%>:</b></label>
            <select name="status" id="selStatus" class="form-control" <% = strListDisable %>>
              <%
                Dim rstStat
                Set rstStat = SQLQuery(cnnDB, "SELECT * From status WHERE status_id > 0 ORDER BY status_id ASC")
                If Not rstStat.EOF Then
                  Do While Not rstStat.EOF
                    If rstStat("status_id") = status Then
                      Response.Write "<option value=""" & rstStat("status_id") & """ selected>" & FixTildes(rstStat("sname")) & "</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rstStat("status_id") & """>" & FixTildes(rstStat("sname")) & "</option>" & vbNewLine
                    End If
                    rstStat.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "Priority")%>:</b></label>
            <select name="priority" class="form-control" <% = strListDisable %>>
              <%
                Dim rstPri
                Set rstPri = SQLQuery(cnnDB, "SELECT * From priority WHERE priority_id > 0 ORDER BY priority_id ASC")
                If Not rstPri.EOF Then
                  Do While Not rstPri.EOF
                    If rstPri("priority_id") = priority Then
                      Response.Write "<option value=""" & rstPri("priority_id") & """ selected>" & FixTildes(rstPri("pname")) & "</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rstPri("priority_id") & """>" & FixTildes(rstPri("pname")) & "</option>" & vbNewLine
                    End If
                    rstPri.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "AssignTo")%>:</b></label>
            <select name="rep" class="form-control" <% = strListDisable %>>
              <%
                Dim rstRes
                Set rstRes = SQLQuery(cnnDB, "SELECT * From tblUsers WHERE IsRep=1 AND RepAccess<>2 ORDER BY uid ASC")
                If Not rstRes.EOF Then
                  Do While Not rstRes.EOF
                    Dim repLabel
                    repLabel = rstRes("uid") & ""
                    If Len(Trim(rstRes("fname") & "")) > 0 Then repLabel = Trim(rstRes("fname") & "") & " (" & rstRes("uid") & ")"
                    If rstRes("sid") = rep Then
                      Response.Write "<option value=""" & rstRes("sid") & """ selected>" & Server.HTMLEncode(repLabel) & "</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rstRes("sid") & """>" & Server.HTMLEncode(repLabel) & "</option>" & vbNewLine
                    End If
                    rstRes.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "Category")%>:</b></label>
            <select name="category" class="form-control" <% = strListDisable %>>
              <%
                Dim rstCat
                Set rstCat = SQLQuery(cnnDB, "SELECT * From categories WHERE category_id > 0 ORDER BY cname ASC")
                If Not rstCat.EOF Then
                  Do While Not rstCat.EOF
                    If rstCat("category_id") = category Then
                      Response.Write "<option value=""" & rstCat("category_id") & """ selected>" & FixTildes(rstCat("cname")) & "</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rstCat("category_id") & """>" & FixTildes(rstCat("cname")) & "</option>" & vbNewLine
                    End If
                    rstCat.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "TimeSpent")%> (<%=lang(cnnDB, "minutes")%>):</b></label>
            <input type="number" name="time_spent" class="form-control" value="<% = time_spent %>" <% = strTextDisable %>>
          </div>

          <div style="font-size:12px;color:#718096;border-top:1px solid #e2e8f0;padding-top:10px;display:flex;flex-direction:column;gap:4px">
            <div><strong>Registro:</strong> <% = DisplayDate(start_date, lhdDateTime) %></div>
            <% If Len(Trim(close_date)) > 0 Then %>
              <div><strong>Cierre:</strong> <% = DisplayDate(close_date, lhdDateTime) %></div>
            <% End If %>
          </div>

        </div>

        <!-- Tarjeta: Informacion del Solicitante -->
        <div class="lhd-card">
          <div class="section-header">&#128100; <%=lang(cnnDB, "ContactInformation")%></div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "UserName")%>:</b></label>
            <div style="font-weight:600;color:#1a202c"><% = Server.HTMLEncode(LhdUserDisplayName(cnnDB, uid)) %></div>
          </div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "EMail")%>:</b></label>
            <input type="email" name="uemail" class="form-control" value="<% = Server.HTMLEncode(uemail) %>" <% = strTextDisable %> required>
          </div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "Department")%>:</b></label>
            <select name="department" class="form-control" <% = strListDisable %>>
              <%
                Dim rstDep
                Set rstDep = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
                If Not rstDep.EOF Then
                  Do While Not rstDep.EOF
                    If rstDep("department_id") = department Then
                      Response.Write "<option value=""" & rstDep("department_id") & """ selected>" & FixTildes(rstDep("dname")) & "</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rstDep("department_id") & """>" & FixTildes(rstDep("dname")) & "</option>" & vbNewLine
                    End If
                    rstDep.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "Location")%>:</b></label>
            <input type="text" name="ulocation" class="form-control" value="<% = FixTildes(Server.HTMLEncode(ulocation)) %>" <% = strTextDisable %>>
          </div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "Phone")%>:</b></label>
            <input type="text" name="uphone" class="form-control" value="<% = Server.HTMLEncode(uphone) %>" <% = strTextDisable %>>
          </div>

          <div style="font-size:12px;color:#718096;border-top:1px solid #e2e8f0;padding-top:10px">
            <strong><%=lang(cnnDB, "EnteredBy")%>:</strong> <% = Server.HTMLEncode(Usr(cnnDB, entered_by, "uid")) %>
          </div>

        </div>

      </div>

    </div>
  </form>
<script type="text/javascript">
  (function () {
    var tokenEl = document.querySelector('input[name=attachToken]');
    if (!tokenEl) { return; }
    var uploadUrl = '../attach/upload.asp';
    if (document.getElementById('solEditor')) {
      LhdEditor.createEditor({ editorId: 'solEditor', fieldId: 'solField', uploadUrl: uploadUrl, token: tokenEl.value, placeholder: 'Describa la solucion aplicada...' });
    }
    if (document.getElementById('attachFile')) {
      LhdEditor.attachments({ fileInputId: 'attachFile', listId: 'attachList', idsId: 'attachIds', uploadUrl: uploadUrl, token: tokenEl.value, maxBytes: <% = LhdMaxAttachBytes(cnnDB) %> });
    }
  })();
</script>

</div>

<%
	' close record sets
	rstCat.Close
	rstDep.Close
	rstStat.Close
	rstPri.Close
	rstRes.Close

  rstProb.Close
  rstSol.Close
  rstNotes.Close

	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close
%>

</body>

</html>
