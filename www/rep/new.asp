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

  Filename: new.asp
  Date:     $Date: 2002/08/28 15:30:54 $
  Version:  $Revision: 1.51.2.1.2.2 $
  Purpose:  This page displays the form used for entering new problems.
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
<title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "NewProblem")%></title>
<link rel="stylesheet" type="text/css" href="../default.css">
<script language="JavaScript">
  // Muestra la solucion solo cuando el estado es Cerrado/Resuelto (>= 100)
  function lhdToggleSolution(sel) {
    var show = parseInt(sel.value, 10) >= 100;
    var sol = document.getElementById('solutionBlock');
    var kb = document.getElementById('kbBlock');
    if (sol) sol.style.display = show ? '' : 'none';
    if (kb) kb.style.display = show ? '' : 'none';
    if (!show) {
      var solEd = document.querySelector('#solEditor .ql-editor');
      if (solEd) solEd.innerHTML = '';
      var ta = document.querySelector('textarea[name=solution]');
      if (ta) ta.value = '';
      var cb = document.querySelector('input[name=kb]');
      if (cb) cb.checked = false;
    }
  }
  document.addEventListener('DOMContentLoaded', function () {
    var sel = document.querySelector('select[name=status]');
    if (sel) lhdToggleSolution(sel);
  });
  function openChild(file,window) {
      childWindow=open(file,window,'resizable=yes,width=560,height=300');
      if (childWindow.opener == null) childWindow.opener = self;
      
  }
</script>
<link rel="stylesheet" type="text/css" href="https://cdnjs.cloudflare.com/ajax/libs/quill/1.3.7/quill.snow.css" />
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/quill/1.3.7/quill.min.js"></script>
<link rel="stylesheet" type="text/css" href="../js/dropzone/dropzone.min.css?v=1">
<script type="text/javascript" src="../js/dropzone/dropzone.min.js?v=1"></script>
<script type="text/javascript" src="../js/lhd-editor.js?v=4"></script>
</head>
<body>

<%
	' Check if user has permissions for this page
	Call CheckRep(cnnDB, sid)

  Dim blnSubmitNew, strSubmitResults
  Dim attachToken
  attachToken = Request.Form("attachToken")
  If Len(attachToken) = 0 Then attachToken = "n" & sid & "x" & CLng(Rnd * 1000000000) & "x" & CLng(Timer * 1000)
  If Cint(Request.Form("save")) = 1 Then
    blnSubmitNew = True
  Else
    blnSubmitNew = False
  End If

  ' ==============================================
  ' Save problem
  If blnSubmitNew Then
    ' Get the information from the form fields
    ' JsEsc: convierte un valor de BD en un literal de cadena seguro para JavaScript
    Function JsEsc(v)
      Dim s
      s = CStr(v & "")
      s = Replace(s, "\", "\\")
      s = Replace(s, """", "\""")
      s = Replace(s, vbCr, " ")
      s = Replace(s, vbLf, " ")
      s = Replace(s, "<", "\u003C")
      s = Replace(s, ">", "\u003E")
      JsEsc = """" & s & """"
    End Function

    Dim uid, uemail, uphone, ulocation, category, department, title, description
    Dim priority, status, rep, time_spent, solution, entered_by, uselectid, kb, duedate, intEmailSent

    uselectid = SafeCLng(Request.Form("uselectid"))
    uid = CleanInputText(Request.Form("uid"))
    uemail = CleanInputText(Request.Form("uemail"))
    uphone = CleanInputText(Request.Form("uphone"))
    ulocation = CleanInputText(Request.Form("ulocation"))
    category = Cint(Request.Form("category"))
    department = SafeCLng(Request.Form("department"))
    title = CleanInputText(Request.Form("title"))
    description = LhdSanitizeHtml(CleanInputText(Request.Form("description")))
    priority = Cint(Request.Form("priority"))
    status = Cint(Request.Form("status"))
    rep = Cint(Request.Form("rep"))
    time_spent = Cint(Request.Form("time_spent"))
    solution = LhdSanitizeHtml(CleanInputText(Request.Form("solution")))
    duedate = ConvertFormattedDate(Request.Form("duedate"))

    If Request.Form("kb") = "on" Then
      kb = 1
    Else
      kb = 0
    End If
	
	If (Request.Form("noemail")<>"on") And (status=Cfg(cnnDB, "CloseStatus")) Then
	  intEmailSent = 1
	Else
	  intEmailSent = 0
	End If

    ' Check for required fields (uemail, category, department, title, description)
    If uselectid <> 0 then 
      uid = usr(cnnDB, uselectid, "uid") 
      uemail = usr(cnnDB, uselectid, "email1") 
      uphone = usr(cnnDB, uselectid, "phone") 
      ulocation = usr(cnnDB, uselectid, "location1") 
      department = usr(cnnDB, uselectid, "department") 
    Else 
      If Len(uid)=0 Then
        Call DisplayError(1, lang(cnnDB, "UserName"))
      End if

      if uemail = Cfg(cnnDB, "BaseEmail") Then
        Call DisplayError(1, lang(cnnDB, "EMailAddress"))
      End if
    End If

    if category = 0 Then
      Call DisplayError(1, lang(cnnDB, "Category"))
    End if

    if (department = 0) And (uselectid = 0) Then
      Call DisplayError(1, lang(cnnDB, "Department"))
    End if

    if Len(title)=0 Then
      Call DisplayError(1, lang(cnnDB, "Title"))
    Elseif Len(title) > 255 Then
      title = Trim(title)
      title = Left(title, 255)
    End if

    If Not IsDate(duedate) Then
      duedate = DateAdd("d", 1, Now())
    End If

    if Not LhdRichHasContent(description) Then
      Call DisplayError(1, lang(cnnDB, "Description"))
    End if

    if (status >= 100) and (Not LhdRichHasContent(solution)) Then
      Call DisplayError(1, lang(cnnDB, "Solution"))
    End if

    ' Get missing variables to enter problem
    Dim id

    Dim dname, depRes
    Set depRes = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & department)
    dname = depRes("dname")

    Dim cname, catRes
    Set catRes = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & category)
    cname = catRes("cname")

    entered_by = sid

    ' Get the problem ID number then immediately update it
    id = GetUnique(cnnDB, "problems")

    ' Convert strings to valid SQL strings
    On Error Resume Next
    uphone = Replace(uphone,"'","''")
    ulocation = Replace(ulocation,"'","''")
    title = Replace(title,"'","''")
    description = Replace(description,"'","''")
    solution = Replace(solution,"'","''")
    On Error Goto 0

    ' All data is present
    ' Write problem into database
    Dim probStr

    ' If status is closed, then include the closed date/time
      If status >= 100 Then
        probStr = "INSERT INTO problems (id, uid, uemail, uphone, ulocation, " & _
        "category, department, title, description, priority, status, start_date, rep, time_spent, " & _
        "close_date, first_response, entered_by, solution, kb, emailsent) " & _
        "VALUES (" & id & ",'" & uid & "','" & uemail & "','" & uphone & "','" & _
        ulocation & "'," & category & "," & department & ",'" & title & "','" & _
        description & "'," & priority & "," & status & "," & SQLDate(Now, lhdAddSQLDelim) & "," & rep & "," & time_spent & _
        "," & SQLDate(Now, lhdAddSQLDelim) & "," & SQLDate(Now, lhdAddSQLDelim) & ","& entered_by & ",'" & solution & "', " & kb & ", " & intEmailSent & ")"
      Else
        probStr = "INSERT INTO problems (id, uid, uemail, uphone, ulocation, " & _
        "category, department, title, description, priority, status, start_date, rep, time_spent, " & _
        "entered_by, solution, kb) " & _
        "VALUES (" & id & ",'" & uid & "','" & uemail & "','" & uphone & "','" & _
        ulocation & "'," & category & "," & department & ",'" & title & "','" & _
        description & "'," & priority & "," & status & "," & SQLDate(Now, lhdAddSQLDelim) & "," & rep & "," & time_spent & _
        "," & entered_by & ",'" & solution & "'," & kb & ")"
      End If

      Dim problemRes
      Set problemRes = SQLQuery(cnnDB, probStr)
      Call LhdLinkAttachments(cnnDB, attachToken, id)

    ' Get support rep information for later
      Dim remail, repRes
      Set repRes = SQLQuery(cnnDB, "SELECT * FROM tblUsers WHERE sid=" & rep)
      remail = repRes("email1")

    ' Send mail to the user and support rep if the problem
    ' was not closed at the time of being entered.
    Dim strSubject, strBody

    If status < 100 Then
      ' Send mail to the user
      If Not (Request.Form("noemail")="on") Then
        Call eMessage(cnnDB, "usernew", id, uemail)
      End If

      ' Send mail to the Rep
      Call eMessage(cnnDB, "repnew", id, remail)

      'Page Rep if enabled
      If (priority >= Cfg(cnnDB, "EnablePager")) And (Len(Usr(cnnDB, rep, "email2")) > 0) Then
        Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
      End If

    End If

    ' If the problem is closed when being entered, send
    ' a different email to the user.
    If status >= 100 Then
      If Not (Request.Form("noemail")="on") Then
        ' Send mail to the user
        Call eMessage(cnnDB, "userclose", id, uemail)
      End If
    End If
    strSubmitResults = Lang(cnnDB, "Problem") & " " & id & " " & Lang(cnnDB, "hasbeenentered") & "."
  End If
  ' ==============================================

  Dim rstDepList, rstCatList, rstStatList, rstPriList, rstRepList
  Call DisplayHeader(cnnDB, sid)
%>
<div class="lhd-wrapper">

  <!-- Breadcrumb y Titulo -->
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <a class="btn btn-secondary btn-sm" href="view.asp">&larr; Volver a la Lista</a>
      <h2 style="margin:0;font-size:20px;font-weight:700;color:#1a202c">
        &#10133; <%=lang(cnnDB, "SubmitANewProblem")%>
      </h2>
    </div>
  </div>

  <% If blnSubmitNew Then Call LhdToast("success", FixTildes(strSubmitResults)) %>
  <% If blnSubmitNew Then %>
    <div class="alert alert-success" style="margin-bottom:24px">
      <span class="alert-icon">&#10004;</span>
      <div style="flex:1">
        <strong><% = FixTildes(strSubmitResults) %></strong>
      </div>
      <a class="btn btn-primary btn-sm" href="details.asp?id=<% = id %>">Ver Incidencia #<% = id %> &rarr;</a>
    </div>
  <% End If %>

  <form action="new.asp" method="POST" name="newProbForm" id="frmNewTicket">
    <input type="hidden" name="save" value="1">

    <div style="display:grid;grid-template-columns:2fr 1fr;gap:20px;align-items:start">

      <!-- COLUMNA IZQUIERDA: Detalle del Problema -->
      <div>

        <div class="lhd-card">
          <div class="section-header">&#128196; Detalle de la Incidencia</div>

          <div class="form-group" style="margin-bottom:16px">
            <label class="form-label"><b><%=lang(cnnDB, "Title")%>:</b> <em>*</em></label>
            <input type="text" name="title" class="form-control" placeholder="Resumen breve del problema o solicitud" required>
          </div>

          <div class="form-group" style="margin-bottom:16px">
            <label class="form-label"><b><%=lang(cnnDB, "Description")%>:</b> <em>*</em></label>
            <div id="descEditor" style="background:#fff;min-height:180px"></div>
            <textarea name="description" id="descField" style="display:none"></textarea>
          </div>

          <div class="form-group" style="margin-bottom:16px">
            <label class="form-label"><b>Archivos adjuntos:</b></label>
            <input type="file" id="attachFile" multiple class="form-control" accept="<% = LhdAcceptAttr(cnnDB) %>" />
            <div id="attachList" style="margin-top:8px"></div>
            <input type="hidden" name="attachIds" id="attachIds" value="" />
            <input type="hidden" name="attachToken" value="<% = attachToken %>" />
          </div>

          <div class="form-group" id="solutionBlock" style="display:none">
            <label class="form-label"><b><%=lang(cnnDB, "Solution")%>:</b> <em>*</em></label>
            <div id="solEditor" style="background:#fff;min-height:120px"></div>
            <textarea name="solution" id="solField" style="display:none"></textarea>
          </div>

          <% If Cfg(cnnDB, "EnableKB") <> 0 Then %>
            <div id="kbBlock" style="margin-top:10px;display:none">
              <label style="display:inline-flex;align-items:center;gap:6px;font-size:13px;cursor:pointer">
                <input type="checkbox" name="kb">
                <%=lang(cnnDB, "EnterinKnowledgeBase")%> (Publicar soluci&oacute;n en la Base de Conocimiento)
              </label>
            </div>
          <% End If %>

        </div>

      </div>

      <!-- COLUMNA DERECHA: Solicitante y Clasificacion -->
      <div>

        <!-- Tarjeta: Informacion del Solicitante -->
        <div class="lhd-card">
          <div class="section-header">&#128100; <%=lang(cnnDB, "ContactInformation")%></div>

          <!-- Selector de Usuario Existente (al elegir uno se llenan los campos) -->
          <%
            Dim rstUserList, jsUsers, rowSid, rowFname
            Set rstUserList = SQLQuery(cnnDB, "SELECT sid, uid, fname, email1, phone, location1, department FROM tblUsers ORDER BY uid ASC")
            jsUsers = ""
            If Not rstUserList.EOF Then
          %>
            <div class="form-group" style="margin-bottom:12px;background:#f8fafc;padding:10px;border-radius:6px;border:1px dashed #cbd5e0">
              <label class="form-label" style="font-size:12px"><b><%=lang(cnnDB, "SelectAnExistingUser")%>:</b></label>
              <select name="uselectid" class="form-control" style="font-size:12px" onchange="lhdFillUser(this.value)">
                <option value="0">&mdash; Ingresar datos manualmente &mdash;</option>
                <%
                  Do While Not rstUserList.EOF
                    rowSid = SafeCInt(rstUserList("sid"))
                    rowFname = rstUserList("fname") & ""
                    If Len(Trim(rowFname)) > 0 Then
                      Response.Write "<option value=""" & rowSid & """>" & Server.HTMLEncode(Trim(rowFname)) & " (" & Server.HTMLEncode(rstUserList("uid") & "") & ")</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rowSid & """>" & Server.HTMLEncode(rstUserList("uid") & "") & "</option>" & vbNewLine
                    End If
                    jsUsers = jsUsers & rowSid & ": {uid: " & JsEsc(rstUserList("uid")) & ", email: " & JsEsc(rstUserList("email1")) & _
                      ", phone: " & JsEsc(rstUserList("phone")) & ", location: " & JsEsc(rstUserList("location1")) & _
                      ", dept: " & SafeCInt(rstUserList("department")) & "}," & vbNewLine
                    rstUserList.MoveNext
                  Loop
                  rstUserList.Close
                %>
              </select>
            </div>
            <script type="text/javascript">
              var lhdUsers = {
<% = jsUsers %>
              };
              function lhdFillUser(sid) {
                var u = lhdUsers[sid];
                if (!u) return;
                var f = document.getElementById('frmNewTicket');
                function setVal(name, val) {
                  var el = f.elements[name];
                  if (el && val) el.value = val;
                }
                setVal('uid', u.uid);
                setVal('uemail', u.email);
                setVal('uphone', u.phone);
                setVal('ulocation', u.location);
                if (u.dept > 0) setVal('department', u.dept);
              }
            </script>
          <% End If %>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "UserName")%>:</b> <em>*</em></label>
            <input type="text" name="uid" class="form-control" placeholder="Nombre de usuario del solicitante">
          </div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "EMail")%>:</b> <em>*</em></label>
            <input type="email" name="uemail" class="form-control" value="<% = Cfg(cnnDB, "BaseEmail") %>">
          </div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "Department")%>:</b> <em>*</em></label>
            <select name="department" class="form-control">
              <option value="0"><%=lang(cnnDB, "SelectDepartment")%></option>
              <%
                Set rstDepList = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
                If Not rstDepList.EOF Then
                  Do While Not rstDepList.EOF
                    Response.Write "<option value=""" & rstDepList("department_id") & """>" & FixTildes(rstDepList("dname")) & "</option>" & vbNewLine
                    rstDepList.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "Location")%>:</b></label>
            <input type="text" name="ulocation" class="form-control" placeholder="Oficina o ubicaci&oacute;n">
          </div>

          <div class="form-group" style="margin-bottom:10px">
            <label class="form-label"><b><%=lang(cnnDB, "Phone")%>:</b></label>
            <input type="text" name="uphone" class="form-control" placeholder="Tel&eacute;fono o extensi&oacute;n">
          </div>

        </div>

        <!-- Tarjeta: Clasificacion y Asignacion -->
        <div class="lhd-card">
          <div class="section-header">&#9881; <%=lang(cnnDB, "Classification")%></div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "Category")%>:</b> <em>*</em></label>
            <select name="category" class="form-control" required>
              <option value="0"><%=lang(cnnDB, "SelectCategory")%></option>
              <%
                Set rstCatList = SQLQuery(cnnDB, "SELECT * From categories WHERE category_id > 0 ORDER BY cname ASC")
                If Not rstCatList.EOF Then
                  Do While Not rstCatList.EOF
                    Response.Write "<option value=""" & rstCatList("category_id") & """>" & FixTildes(rstCatList("cname")) & "</option>" & vbNewLine
                    rstCatList.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "Priority")%>:</b> <em>*</em></label>
            <select name="priority" class="form-control">
              <%
                Set rstPriList = SQLQuery(cnnDB, "SELECT * From priority WHERE priority_id > 0 ORDER BY priority_id ASC")
                If Not rstPriList.EOF Then
                  Do While Not rstPriList.EOF
                    If rstPriList("priority_id") = Cfg(cnnDB, "DefaultPriority") Then
                      Response.Write "<option value=""" & rstPriList("priority_id") & """ selected>" & FixTildes(rstPriList("pname")) & "</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rstPriList("priority_id") & """>" & FixTildes(rstPriList("pname")) & "</option>" & vbNewLine
                    End If
                    rstPriList.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "Status")%>:</b> <em>*</em></label>
            <select name="status" class="form-control" onchange="lhdToggleSolution(this)">
              <%
                Set rstStatList = SQLQuery(cnnDB, "SELECT * From status WHERE status_id > 0 ORDER BY status_id ASC")
                If Not rstStatList.EOF Then
                  Do While Not rstStatList.EOF
                    If rstStatList("status_id") = Cfg(cnnDB, "DefaultStatus") Then
                      Response.Write "<option value=""" & rstStatList("status_id") & """ selected>" & FixTildes(rstStatList("sname")) & "</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rstStatList("status_id") & """>" & FixTildes(rstStatList("sname")) & "</option>" & vbNewLine
                    End If
                    rstStatList.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:12px">
            <label class="form-label"><b><%=lang(cnnDB, "AssignTo")%>:</b></label>
            <select name="rep" class="form-control">
              <%
                Set rstRepList = SQLQuery(cnnDB, "SELECT * From tblUsers WHERE IsRep=1 AND RepAccess<>2 ORDER BY uid ASC")
                If Not rstRepList.EOF Then
                  Do While Not rstRepList.EOF
                    Dim repLabel
                    repLabel = rstRepList("uid") & ""
                    If Len(Trim(rstRepList("fname") & "")) > 0 Then repLabel = Trim(rstRepList("fname") & "") & " (" & rstRepList("uid") & ")"
                    If rstRepList("sid") = sid Then
                      Response.Write "<option value=""" & rstRepList("sid") & """ selected>" & Server.HTMLEncode(repLabel) & " (Yo)</option>" & vbNewLine
                    Else
                      Response.Write "<option value=""" & rstRepList("sid") & """>" & Server.HTMLEncode(repLabel) & "</option>" & vbNewLine
                    End If
                    rstRepList.MoveNext
                  Loop
                End If
              %>
            </select>
          </div>

          <div class="form-group" style="margin-bottom:16px">
            <label class="form-label"><b><%=lang(cnnDB, "TimeSpent")%> (<%=lang(cnnDB, "minutes")%>):</b></label>
            <input type="number" name="time_spent" class="form-control" value="0">
          </div>

          <% If Cfg(cnnDB, "EmailType") <> 0 Then %>
            <div style="margin-bottom:16px">
              <label style="display:inline-flex;align-items:center;gap:6px;font-size:12px;color:#4a5568;cursor:pointer">
                <input type="checkbox" name="noemail" id="noEmailCheck">
                <%=lang(cnnDB, "Dontsendemailtouser")%>
              </label>
            </div>
          <% End If %>

          <button type="submit" class="btn btn-primary" style="width:100%;justify-content:center;padding:12px;font-size:14px">
            &#10133; <%=lang(cnnDB, "SubmitProblem")%>
          </button>

        </div>

      </div>

    </div>
  </form>
<script type="text/javascript">
  (function () {
    var token = document.querySelector('input[name=attachToken]').value;
    var uploadUrl = '../attach/upload.asp';
    LhdEditor.createEditor({ editorId: 'descEditor', fieldId: 'descField', uploadUrl: uploadUrl, token: token, placeholder: 'Describa el problema con el mayor detalle posible...' });
    LhdEditor.createEditor({ editorId: 'solEditor', fieldId: 'solField', uploadUrl: uploadUrl, token: token, placeholder: 'Solucion aplicada...' });
    LhdEditor.attachments({ fileInputId: 'attachFile', listId: 'attachList', idsId: 'attachIds', uploadUrl: uploadUrl, token: token, maxBytes: <% = LhdMaxAttachBytes(cnnDB) %> });
    LhdEditor.guardSubmit('frmNewTicket', 'descField', 'Escriba la descripcion del problema.');
  })();
</script>

</div>

<%
	' close record sets
	rstCatList.Close
	rstDepList.Close
	rstStatList.Close
	rstPriList.Close
	rstRepList.Close

	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close

%>

</body>

</html>
