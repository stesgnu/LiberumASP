<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
  <% Option Explicit 'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: new.asp (user)
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Form for users to enter new problems.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
    Session.timeout = 40
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "NewProblem")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
      <link rel="stylesheet" type="text/css" href="https://cdnjs.cloudflare.com/ajax/libs/quill/1.3.7/quill.snow.css" />
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/quill/1.3.7/quill.min.js"></script>
    <link rel="stylesheet" type="text/css" href="../js/dropzone/dropzone.min.css?v=1">
<script type="text/javascript" src="../js/dropzone/dropzone.min.js?v=1"></script>
<script type="text/javascript" src="../js/lhd-editor.js?v=4"></script>
  </head>
  <body>

    <%
      ' See if authenticated Call CheckUser(cnnDB, sid) ' Determine username and email
      Dim uid, uemail
      uid    = Usr(cnnDB, sid, "uid")
      uemail = Usr(cnnDB, sid, "email1")
      Dim attachToken
      attachToken = "u" & sid & "x" & CLng(Rnd * 1000000000) & "x" & CLng(Timer * 1000)

      Call DisplayHeader(cnnDB, sid)
    %>

    <div class="lhd-wrapper">
      <div class="lhd-page-title">
        <span class="title-icon">&#10133;</span>
        <%=lang(cnnDB, "SubmitANewProblem")%>
      </div>

      <!-- Error de validacion (oculto por defecto) -->
      <div class="alert alert-error" id="form-error" style="display:none">
        <span class="alert-icon">&#9888;</span>
        <span>Por favor completa todos los campos obligatorios (<em>*</em>).</span>
      </div>

      <form action="postnew.asp" method="POST" id="new-ticket-form" novalidate>
        <input type="hidden" name="uid" value="<% = uid %>">

        <div class="lhd-card">
          <div class="section-header">&#128100; <%=lang(cnnDB, "ContactInformation")%></div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label"><%=lang(cnnDB, "UserName")%></label>
              <input type="text" class="form-control" value="<% = Server.HTMLEncode(LhdUserDisplayName(cnnDB, uid)) %>" readonly disabled>
            </div>
            <div class="form-group">
              <label class="form-label" for="uemail">
                <%=lang(cnnDB, "EMail")%> <span class="form-required">*</span>
              </label>
              <input class="form-control" type="email" id="uemail" name="uemail"
                     value="<% = uemail %>"
                     placeholder="correo@empresa.com"
                     autocomplete="email">
              <div class="form-invalid-msg" id="err-uemail">Ingresa un email valido.</div>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label" for="ulocation"><%=lang(cnnDB, "Location")%></label>
              <input class="form-control" type="text" id="ulocation" name="ulocation"
                     value="<% = Usr(cnnDB, sid, "location1") %>"
                     placeholder="Ej: Oficina 3B, Piso 2">
            </div>
            <div class="form-group">
              <label class="form-label" for="uphone"><%=lang(cnnDB, "Phone")%></label>
              <input class="form-control" type="tel" id="uphone" name="uphone"
                     value="<% = Usr(cnnDB, sid, "phone") %>"
                     placeholder="Ext. 1234"
                     autocomplete="tel">
            </div>
          </div>
        </div>

        <div class="lhd-card">
          <div class="section-header">&#128202; <%=lang(cnnDB, "ProblemClassification")%></div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label" for="department">
                <%=lang(cnnDB, "Department")%> <span class="form-required">*</span>
              </label>
              <select class="form-control" id="department" name="department">
                <option value="0">-- Seleccionar --</option>
                <%
                  Dim rstDepList
                  Set rstDepList = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
                  If Not rstDepList.EOF Then
                    Do While Not rstDepList.EOF
                      Dim depSel
                      If rstDepList("department_id") = Usr(cnnDB, sid, "department") Then
                        depSel = " selected"
                      Else
                        depSel = ""
                      End If
                %>
                      <option value="<% = rstDepList("department_id") %>"<% = depSel %>>
                        <% = rstDepList("dname") %>
                      </option>
                <%
                      rstDepList.MoveNext
                    Loop
                  End If
                  rstDepList.Close
                %>
              </select>
              <div class="form-invalid-msg" id="err-department">Selecciona un departamento.</div>
            </div>

            <div class="form-group">
              <label class="form-label" for="category">
                <%=lang(cnnDB, "Category")%> <span class="form-required">*</span>
              </label>
              <select class="form-control" id="category" name="category">
                <option value="0" selected><%=lang(cnnDB, "SelectCategory")%></option>
                <%
                  Dim rstCatList
                  Set rstCatList = SQLQuery(cnnDB, "SELECT * From categories WHERE category_id > 0 ORDER BY cname ASC")
                  If Not rstCatList.EOF Then
                    Do While Not rstCatList.EOF
                %>
                      <option value="<% = rstCatList("category_id") %>">
                        <% = rstCatList("cname") %>
                      </option>
                <%
                      rstCatList.MoveNext
                    Loop
                  End If
                  rstCatList.Close
                %>
              </select>
              <div class="form-invalid-msg" id="err-category">Selecciona una categoria.</div>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label" for="priority">
                <%=lang(cnnDB, "Priority")%> <span class="form-required">*</span>
              </label>
              <select class="form-control" id="priority" name="priority">
                <%
                  Dim rstPriList
                  Set rstPriList = SQLQuery(cnnDB, "SELECT * From priority WHERE priority_id > 0 ORDER BY priority_id ASC")
                  If Not rstPriList.EOF Then
                    Do While Not rstPriList.EOF
                      Dim priSel
                      If rstPriList("priority_id") = Cfg(cnnDB, "DefaultPriority") Then
                        priSel = " selected"
                      Else
                        priSel = ""
                      End If
                %>
                      <option value="<% = rstPriList("priority_id") %>"<% = priSel %>>
                        <% = rstPriList("pname") %>
                      </option>
                <%
                      rstPriList.MoveNext
                    Loop
                  End If
                  rstPriList.Close
                %>
              </select>
            </div>

            <div class="form-group">
              <label class="form-label" for="duedate">
                <%=lang(cnnDB, "DueDate")%> <span class="form-required">*</span>
              </label>
              <input class="form-control"
                     type="date"
                     id="duedate"
                     name="duedate"
                     value="<% = ISODate(DateAdd("d", 1, Now())) %>"
                     min="<% = ISODate(Now()) %>">
              <div class="form-hint">Formato: YYYY-MM-DD</div>
            </div>
          </div>
        </div>

        <div class="lhd-card">
          <div class="section-header">&#128221; <%=lang(cnnDB, "ProblemInformation")%></div>

          <div class="form-group">
            <label class="form-label" for="title">
              <%=lang(cnnDB, "Title")%> <span class="form-required">*</span>
            </label>
            <input class="form-control" type="text" id="title" name="title"
                   placeholder="Breve descripcion del problema..."
                   maxlength="255">
            <div class="form-invalid-msg" id="err-title">El titulo es requerido.</div>
          </div>

          <div class="form-group">
            <label class="form-label" for="description">
              <%=lang(cnnDB, "Description")%> <span class="form-required">*</span>
            </label>
            <div id="descEditor" style="background:#fff;min-height:200px"></div>
            <textarea id="description" name="description" style="display:none"></textarea>
            <div class="char-counter" id="char-count">0 caracteres</div>
            <div class="form-invalid-msg" id="err-description">La descripcion es requerida.</div>
          </div>

          <div class="form-group">
            <label class="form-label"><b>Archivos adjuntos:</b></label>
            <input type="file" id="attachFile" multiple class="form-control" accept="<% = LhdAcceptAttr(cnnDB) %>" />
            <div id="attachList" style="margin-top:8px"></div>
            <input type="hidden" name="attachIds" id="attachIds" value="" />
            <input type="hidden" name="attachToken" value="<% = attachToken %>" />
          </div>

          <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:4px">
            <button type="reset" class="btn btn-secondary">
              &#10006; <%=lang(cnnDB, "ClearForm")%>
            </button>
            <button type="submit" class="btn btn-primary btn-lg" id="submit-btn">
              &#128228; <%=lang(cnnDB, "SubmitProblem")%>
            </button>
          </div>
        </div>

      </form>
    <script type="text/javascript">
      (function () {
        var token = document.querySelector('input[name=attachToken]').value;
        var uploadUrl = '../attach/upload.asp';
        LhdEditor.createEditor({ editorId: 'descEditor', fieldId: 'description', uploadUrl: uploadUrl, token: token, placeholder: 'Describa el problema con el mayor detalle posible...' });
        LhdEditor.attachments({ fileInputId: 'attachFile', listId: 'attachList', idsId: 'attachIds', uploadUrl: uploadUrl, token: token, maxBytes: <% = LhdMaxAttachBytes(cnnDB) %> });
      })();
    </script>
    </div>

    <script>
      // === Validacion de formulario lado cliente ===
      document.getElementById('new-ticket-form').addEventListener('submit', function (e) {
        var valid = true; // Reset
        document.querySelectorAll('.form-control').forEach(function (el) { el.classList.remove('is-invalid'); });
        document.querySelectorAll('.form-invalid-msg').forEach(function (el) { el.style.display = 'none'; }); function
          invalidate(fieldId, errId) {
            var el = document.getElementById(fieldId); if (el) el.classList.add('is-invalid'); var
              err = document.getElementById(errId); if (err) err.style.display = 'block'; valid = false;
        } // Email var
        email = document.getElementById('uemail').value.trim(); if (!email || email.indexOf('@') < 0)
          invalidate('uemail', 'err-uemail'); // Department var dept=document.getElementById('department').value; if (!dept ||
        dept === '0' ) invalidate('department', 'err-department'); // Category var
      cat = document.getElementById('category').value; if (!cat || cat === '0') invalidate('category', 'err-category'); //
    Title var title = document.getElementById('title').value.trim(); if (!title) invalidate('title', 'err-title'); //
    Description var desc = document.getElementById('description').value.trim(); if (!desc)
        invalidate('description', 'err-description'); if (!valid) {
          e.preventDefault();
          document.getElementById('form-error').style.display = 'flex'; if (window.LhdToast) LhdToast.error('Por favor completa todos los campos obligatorios.'); window.scrollTo({ top: 0, behavior: 'smooth' });
        } });
      //===Contador de caracteres en textarea===var ta=document.getElementById('description'); var
      ctr = document.getElementById('char-count'); if (ta && ctr) {
        ta.addEventListener('input', function () {
          var
          n = ta.value.length; ctr.textContent = n + ' caracter' + (n !== 1 ? 'es' : ''); ctr.style.color = n > 2000 ? '#B02A37' :
            '#718096';
        });
      }
    </script>

    <% Call DisplayFooter(cnnDB, sid)
  cnnDB.Close %>
      </body>

      </html>