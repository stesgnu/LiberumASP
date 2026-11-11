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

  Filename: modify.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  This page is used to add/modify depts, categories, etc.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    ' ---- Columnas adicionales de categories (ANS / estado) ----
    Dim xNames, xLabels, xHas(5), xType(5), xVal(5), xSize(5)
    xNames  = Array("estado", "ANS", "TipoAns", "TipoIndicador", "Ans_Ini", "TipoAnsIni")
    xLabels = Array("Estado", "ANS", "Tipo de ANS", "Tipo de indicador", "ANS inicial", "Tipo de ANS inicial")

    ' Tipos ADO numericos / booleano
    Function XIsNum(t)
      Select Case t
        Case 2, 3, 4, 5, 6, 14, 16, 17, 18, 19, 20, 21, 131
          XIsNum = True
        Case Else
          XIsNum = False
      End Select
    End Function

    Function XIsBool(t)
      XIsBool = (t = 11)
    End Function

    ' Valor de un campo como texto para mostrarlo en el formulario
    Function XText(v, t)
      If IsNull(v) Or IsEmpty(v) Then
        XText = ""
      ElseIf XIsBool(t) Then
        If v Then
          XText = "1"
        Else
          XText = "0"
        End If
      Else
        XText = CStr(v)
      End If
    End Function

    ' <datalist> con los valores distintos que ya usan otras categorias
    Function XDistinct(col)
      Dim rsD, outD
      outD = ""
      On Error Resume Next
      Set rsD = cnnDB.Execute("SELECT DISTINCT [" & col & "] FROM categories WHERE [" & col & "] IS NOT NULL ORDER BY [" & col & "]")
      If Err.Number = 0 Then
        Do While Not rsD.EOF
          If Len(Trim(rsD(0) & "")) > 0 Then outD = outD & "<option value=""" & Server.HTMLEncode(Trim(rsD(0) & "")) & """>"
          rsD.MoveNext
        Loop
        rsD.Close
      End If
      Err.Clear
      On Error Goto 0
      XDistinct = outD
    End Function

    Function IIfSel(cond, dummy)
      If cond Then
        IIfSel = " selected"
      Else
        IIfSel = ""
      End If
    End Function

    ' Dibuja el control adecuado para la columna xi segun su tipo de dato
    Sub WriteExtra(xi)
      Dim nm, v, listHtml, selA, selB
      nm = "x_" & xNames(xi)
      v = XText(xVal(xi), xType(xi))
      Response.Write "<div class=""form-group""><label class=""form-label"" for=""" & nm & """><b>" & xLabels(xi) & ":</b></label>"
      If XIsBool(xType(xi)) Or (XIsNum(xType(xi)) And LCase(xNames(xi)) = "estado") Then
        ' Estado activo / inactivo (se conserva cualquier otro valor ya existente)
        Response.Write "<select class=""form-control"" id=""" & nm & """ name=""" & nm & """>"
        selA = ""
        selB = ""
        ' Una categoria nueva (o sin valor) queda Activa por defecto
        If v = "1" Or v = "" Then selA = " selected"
        If v = "0" Then selB = " selected"
        Response.Write "<option value=""1""" & selA & ">Activo</option><option value=""0""" & selB & ">Inactivo</option>"
        If v <> "" And v <> "0" And v <> "1" Then Response.Write "<option value=""" & Server.HTMLEncode(v) & """ selected>Valor actual (" & Server.HTMLEncode(v) & ")</option>"
        Response.Write "</select>"
      ElseIf XIsNum(xType(xi)) Then
        Response.Write "<input class=""form-control"" type=""number"" step=""any"" min=""0"" id=""" & nm & """ name=""" & nm & """ value=""" & Server.HTMLEncode(Replace(v, ",", ".")) & """>"
      ElseIf LCase(xNames(xi)) = "tipoans" Or LCase(xNames(xi)) = "tipoansini" Then
        ' Unidad del ANS: M = Minutos, H = Horas, D = Dias
        Dim uv
        uv = UCase(Trim(v))
        Response.Write "<select class=""form-control"" id=""" & nm & """ name=""" & nm & """>"
        Response.Write "<option value=""""" & IIfSel(uv = "", True) & ">&mdash; Sin definir &mdash;</option>"
        Response.Write "<option value=""M""" & IIfSel(uv = "M", True) & ">Minutos (M)</option>"
        Response.Write "<option value=""H""" & IIfSel(uv = "H", True) & ">Horas (H)</option>"
        Response.Write "<option value=""D""" & IIfSel(uv = "D", True) & ">D&iacute;as (D)</option>"
        If uv <> "" And uv <> "M" And uv <> "H" And uv <> "D" Then Response.Write "<option value=""" & Server.HTMLEncode(v) & """ selected>Valor actual (" & Server.HTMLEncode(v) & ")</option>"
        Response.Write "</select>"
      Else
        listHtml = XDistinct(xNames(xi))
        If LCase(xNames(xi)) = "estado" And Len(listHtml) = 0 Then listHtml = "<option value=""Activo""><option value=""Inactivo"">"
        Response.Write "<input class=""form-control"" type=""text"" list=""dl_" & nm & """ id=""" & nm & """ name=""" & nm & """ autocomplete=""off"" value=""" & Server.HTMLEncode(v) & """"
        If xSize(xi) > 0 And xSize(xi) < 4000 Then Response.Write " maxlength=""" & xSize(xi) & """"
        Response.Write "><datalist id=""dl_" & nm & """>" & listHtml & "</datalist>"
      End If
      Response.Write "</div>"
    End Sub
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "CreateModify")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .md-card { max-width: 780px; margin: 0 auto 18px; }
      .md-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 4px 18px; }
      .md-hint { font-size: 12px; color: #64748b; margin-top: 4px; }
      .md-lock { background: #F1F5F9 !important; color: #475569 !important; cursor: not-allowed; }
      .md-actions { max-width: 780px; margin: 0 auto; display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
      .md-note { background: var(--color-primary-light, #E8F0FB); border-left: 4px solid var(--color-primary, #0057B8); border-radius: 6px; padding: 10px 14px; font-size: 13px; color: #1e293b; margin-bottom: 14px; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      Call DisplayHeader(cnnDB, sid)

      ' Get the type of add/modify
      ' 1 - Rep
      ' 2 - Category
      ' 3 - Department
      ' 4 - Priority
      ' 5 - Status
      ' 6 - Language
      Dim mType, strLanguageName, intLangID
      intLangID = Request.QueryString("mLangID")
      strLanguageName = Request.QueryString("mLanguage")
      mType = Cint(Request.QueryString("Mtype"))
      If (mType < 1) or (mType > 6) Then
        Call DisplayError(3, "Invalid type to create/modify")
      End If

      Dim data_id, data1, data2, data3, data4, numDataFields
      Dim data1Name, data2Name, data3Name, data4Name, title

      If Not Len(Request.QueryString("id")) = 0 Then
        ' An ID is sent, so this is a modify of that ID
        data_id = Cint(Request.QueryString("id"))
        Select Case mType
          Case 1	' Modify Rep
            Dim repRes
            Set repRes = SQLQuery(cnnDB, "SELECT ruid, rname, remail, rdescription FROM reps WHERE rep_id=" & data_id)
            data1 = repRes("ruid")
            data2 = repRes("rname")
            data3 = repRes("remail")
            data4 = repRes("rdescription")
            repRes.Close
          Case 2	' Modify Category
            Dim catRes
            Set catRes = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & data_id)
            data1 = catRes("cname")
            data2 = catRes("rep_id")
            catRes.Close
          Case 3	' Modify Department
            Dim depRes
            Set depRes = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & data_id)
            data1 = depRes("dname")
            depRes.Close
          Case 4	' Modify Priority
            Dim priRes
            Set priRes = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & data_id)
            data1 = data_id
            data2 = priRes("pname")
            priRes.Close
          Case 5	' Modify Status
            Dim statRes
            Set statRes = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & data_id)
            data1 = data_id
            data2 = statRes("sname")
            statRes.Close
          Case 6  ' Modify Language
            Dim rstLanguage
            set rstLanguage = SQLQuery(cnnDB, "SELECT * FROM tblLanguage WHERE id=" & data_id)
            data1 = rstLanguage("LangName")
            data2 = rstLanguage("Localized")
            rstLanguage.Close
        End Select
      Else

        ' No ID is sent, so this is an add
        data_id = 0

        Select Case mType
          Case 1	' Modify Rep
            data3Name = Cfg(cnnDB, "BaseEmail")
        End Select
      End If

      ' Categorias: leer las columnas adicionales (si existen en la tabla) con su tipo de dato
      Dim rsX, xi, fi, xWhere, anyExtra
      anyExtra = False
      If mType = 2 Then
        If data_id > 0 Then
          xWhere = "category_id=" & data_id
        Else
          xWhere = "1=0"
        End If
        Set rsX = SQLQuery(cnnDB, "SELECT * FROM categories WHERE " & xWhere)
        For xi = 0 To 5
          xHas(xi) = False
          xVal(xi) = Null
          xSize(xi) = 0
          For fi = 0 To rsX.Fields.Count - 1
            If LCase(rsX.Fields(fi).Name) = LCase(xNames(xi)) Then
              xHas(xi) = True
              anyExtra = True
              xType(xi) = rsX.Fields(fi).Type
              xSize(xi) = rsX.Fields(fi).DefinedSize
              If Not rsX.EOF Then xVal(xi) = rsX.Fields(fi).Value
            End If
          Next
        Next
        rsX.Close
      End If

      'Fill in the names of the data fields
      Select Case mType
        Case 1	' Modify Rep
          title = lang(cnnDB, "SupportRepresentatives")
          data1Name = lang(cnnDB, "Username")
          data2Name = lang(cnnDB, "FullName")
          data3Name = lang(cnnDB, "Email")
          data4Name = lang(cnnDB, "Description")
          numDataFields = 4
        Case 2	' Modify Category
          title = lang(cnnDB, "Category")
          data1Name = lang(cnnDB, "CategoryName")
          data2Name = lang(cnnDB, "PrimaryRep")
          numDataFields = 2
        Case 3  ' Modify Departments
          title = lang(cnnDB, "Department")
          data1Name = lang(cnnDB, "Department")
          numDataFields = 1
        Case 4	' Modify Priority
          title = lang(cnnDB, "Priority")
          data1Name = lang(cnnDB, "PriorityNumber")
          data2Name = lang(cnnDB, "PriorityName")
          numDataFields = 2
        Case 5	' Modify Status
          title = lang(cnnDB, "Status")
          data1Name = lang(cnnDB, "StatusNumber")
          data2Name = lang(cnnDB, "StatusName")
          numDataFields = 2
        Case 6  ' Modify language
          title = lang(cnnDB, "Language")
          data1Name = lang(cnnDB, "LanguageName")
          data2Name = lang(cnnDB, "LocalizedName")
          numDataFields = 2
      End Select

      ' Enlace de retorno segun el tipo
      Dim backUrl, backText
      Select Case mType
        Case 1
          backUrl = "viewrep.asp"
          backText = lang(cnnDB, "SupportReps")
        Case 2
          backUrl = "viewcat.asp"
          backText = lang(cnnDB, "Categories")
        Case 3
          backUrl = "viewdep.asp"
          backText = lang(cnnDB, "Departments")
        Case 4
          backUrl = "viewpri.asp"
          backText = lang(cnnDB, "Priorities")
        Case 5
          backUrl = "viewstatus.asp"
          backText = lang(cnnDB, "Statuses")
        Case 6
          backUrl = "viewlang.asp"
          backText = lang(cnnDB, "Languages")
      End Select

      ' Campos que no se pueden cambiar al editar (conservan el historico)
      Dim lockData1
      lockData1 = False
      If data_id > 0 And (mType = 2 Or mType = 4 Or mType = 5) Then lockData1 = True
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#9998;</span>
          <%=lang(cnnDB, "CreateModify")%>&nbsp;<% = title %>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="<% = backUrl %>">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<% = backText %></a>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>

      <form method="post" action="postmods.asp" id="form1" name="form1">
        <input type="hidden" name="data_id" value="<% = data_id %>">
        <input type="hidden" name="numdatafields" value="<% = numDataFields %>">
        <input type="hidden" name="mtype" value="<% = mType %>">
        <input type="hidden" name="mLanguage" value="<% = Server.HTMLEncode(strLanguageName & "") %>">
        <input type="hidden" name="mLangID" value="<% = Server.HTMLEncode(intLangID & "") %>">

        <div class="lhd-card md-card">
          <div class="section-header">&#128221; Informaci&oacute;n general</div>

          <% If mType = 2 And data_id > 0 Then %>
            <div class="md-note">&#128274; El nombre de la categor&iacute;a no se puede modificar para conservar correctamente el hist&oacute;rico de los tickets.</div>
          <% End If %>

          <% If numDataFields >= 1 Then %>
            <div class="form-group">
              <label class="form-label" for="data1"><b><% = data1Name %>:</b></label>
              <input type="text" class="form-control<% If lockData1 Then Response.Write(" md-lock") %>" id="data1" name="data1" maxlength="50"
                     value="<% = Server.HTMLEncode(data1 & "") %>" <% If lockData1 Then Response.Write("readonly") %>>
            </div>
          <% End If %>

          <% If numDataFields >= 2 Then %>
            <div class="form-group">
              <label class="form-label" for="data2"><b><% = data2Name %>:</b></label>
              <% If mType <> 2 Then %>
                <input type="text" class="form-control" id="data2" name="data2" value="<% = Server.HTMLEncode(data2 & "") %>">
              <% Else %>
                <select name="data2" id="data2" class="form-control">
                  <%
                    Dim replRes, repLabel
                    Set replRes = SQLQuery(cnnDB, "SELECT * FROM tblUsers WHERE IsRep=1 AND RepAccess <> 2 ORDER BY fname ASC, uid ASC")
                    Do While Not replRes.EOF
                      ' Nombre completo (usuario); si no tiene nombre, solo el usuario
                      repLabel = replRes("uid") & ""
                      If Len(Trim(replRes("fname") & "")) > 0 Then repLabel = Trim(replRes("fname") & "") & " (" & replRes("uid") & ")"
                      If SafeCInt(replRes("sid")) = SafeCInt(data2) Then
                        Response.Write "<option value=""" & replRes("sid") & """ selected>" & Server.HTMLEncode(repLabel) & "</option>"
                      Else
                        Response.Write "<option value=""" & replRes("sid") & """>" & Server.HTMLEncode(repLabel) & "</option>"
                      End If
                      replRes.MoveNext
                    Loop
                    replRes.Close
                  %>
                </select>
                <div class="md-hint">Escribe para buscar por nombre o usuario.</div>
              <% End If %>
            </div>
          <% End If %>

          <% If numDataFields >= 3 Then %>
            <div class="form-group">
              <label class="form-label" for="data3"><b><% = data3Name %>:</b></label>
              <input type="text" class="form-control" id="data3" name="data3" value="<% = Server.HTMLEncode(data3 & "") %>">
            </div>
          <% End If %>

          <% If numDataFields >= 4 Then %>
            <div class="form-group">
              <label class="form-label" for="data4"><b><% = data4Name %>:</b></label>
              <input type="text" class="form-control" id="data4" name="data4" value="<% = Server.HTMLEncode(data4 & "") %>">
            </div>
          <% End If %>

          <% If (mType = 4 Or mType = 5) And data_id > 0 Then %>
            <div class="md-hint"><%=lang(cnnDB, "StatusandPrioritynumberscannotbechanged")%></div>
          <% End If %>
        </div>

        <% If mType = 2 And anyExtra Then %>
        <div class="lhd-card md-card">
          <div class="section-header">&#9201; Estado y acuerdos de nivel de servicio (ANS)</div>
          <div class="md-grid">
            <%
              For xi = 0 To 5
                If xHas(xi) Then Call WriteExtra(xi)
              Next
            %>
          </div>
        </div>
        <% End If %>

        <div class="md-actions">
          <a class="btn btn-secondary" href="<% = backUrl %>">Cancelar</a>
          <button type="submit" class="btn btn-primary">&#128190; <%=lang(cnnDB, "Save")%></button>
        </div>
      </form>

    </div>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
