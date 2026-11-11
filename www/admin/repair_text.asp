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
  Filename: repair_text.asp (admin)
  Purpose:  Repara en la base de datos los textos con tildes y enies que
            quedaron guardados como caracteres extranos (mojibake).
            Primero muestra una vista previa; los cambios solo se
            escriben cuando el administrador los confirma (POST).
            No modifica contrasenas ni la tabla de configuracion.
  -->

  <!-- #include file = "../public.asp" -->
  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
    Call CheckUser(cnnDB, sid)
    Call CheckAdmin
  %>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta charset="utf-8" />
  <title><% = Cfg(cnnDB, "SiteName") %> - Reparar texto acentuado</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="stylesheet" type="text/css" href="../default.css?v=4" />
</head>
<body>

<%
  ' Tablas a revisar: nombre | columnas de clave | columnas de texto
  Dim tableDefs
  tableDefs = Array( _
    "categories|category_id|cname", _
    "departments|department_id|dname", _
    "priority|priority_id|pname", _
    "status|status_id|sname", _
    "problems|id|uid,uemail,ulocation,uphone,title,description,solution", _
    "tblNotes|id,addDate|note,uid", _
    "tblUsers|sid|uid,fname,email1,email2,phone,location1,location2,firstname,lastname,phone_home,phone_mobile,jobfunction,userresume,statustext", _
    "tblEmailMsg|type|subject,body", _
    "tblLangStrings|id,variable|LangText" _
  )

  Dim mode, pendingSql, rowsScanned, changedRows, changedFields, errCount, previewHtml, previewCount, tableErrors
  tableErrors = ""
  mode = "preview"
  If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("mode") = "apply" Then mode = "apply"
  pendingSql = ""
  rowsScanned = 0
  changedRows = 0
  changedFields = 0
  errCount = 0
  previewHtml = ""
  previewCount = 0

  ' Literal SQL para un valor de clave
  Function KeyLiteral(v)
    Dim d
    If VarType(v) = vbDate Then
      d = v
      KeyLiteral = SQLDate(d, lhdAddSQLDelim)
    ElseIf VarType(v) >= vbInteger And VarType(v) <= vbDouble Then
      KeyLiteral = CStr(v)
    Else
      KeyLiteral = "'" & Replace(CStr(v & ""), "'", "''") & "'"
    End If
  End Function

  ' Agrega una fila a la vista previa (maximo 300 filas)
  Sub AddPreview(tblName, keyText, colName, oldText, newText)
    changedFields = changedFields + 1
    If previewCount < 300 Then
      previewCount = previewCount + 1
      previewHtml = previewHtml & "<tr><td>" & Server.HTMLEncode(tblName) & "</td>" & _
        "<td>" & Server.HTMLEncode(keyText) & "</td>" & _
        "<td>" & Server.HTMLEncode(colName) & "</td>" & _
        "<td>" & Server.HTMLEncode(Left(oldText, 120)) & "</td>" & _
        "<td>" & Server.HTMLEncode(Left(newText, 120)) & "</td></tr>" & vbNewLine
    End If
  End Sub

  ' Revisa una tabla y prepara los UPDATE necesarios
  Sub RepairTable(tblName, pkCols, txtCols)
    Dim selectCols, j, rst, whereSql, setSql, val, fixed, keyText, rowKey
    selectCols = ""
    For j = 0 To UBound(pkCols)
      selectCols = selectCols & pkCols(j) & ", "
    Next
    For j = 0 To UBound(txtCols)
      selectCols = selectCols & txtCols(j) & ", "
    Next
    selectCols = Left(selectCols, Len(selectCols) - 2)

    On Error Resume Next
    Set rst = cnnDB.Execute("SELECT " & selectCols & " FROM " & tblName)
    If Err.Number <> 0 Then
      tableErrors = tableErrors & "<li><b>" & tblName & "</b>: " & Server.HTMLEncode(Err.Description) & "</li>"
      Err.Clear
      On Error Goto 0
      Exit Sub
    End If
    On Error Goto 0
    Do While Not rst.EOF
      rowsScanned = rowsScanned + 1
      whereSql = ""
      keyText = ""
      For j = 0 To UBound(pkCols)
        If j > 0 Then
          whereSql = whereSql & " AND "
          keyText = keyText & ", "
        End If
        whereSql = whereSql & pkCols(j) & "=" & KeyLiteral(rst(pkCols(j)))
        keyText = keyText & pkCols(j) & "=" & CStr(rst(pkCols(j)) & "")
      Next

      setSql = ""
      For j = 0 To UBound(txtCols)
        val = CStr(rst(txtCols(j)) & "")
        fixed = LhdRepairText(val, False)
        If fixed <> val Then
          If Len(setSql) > 0 Then setSql = setSql & ", "
          setSql = setSql & txtCols(j) & "='" & Replace(fixed, "'", "''") & "'"
          Call AddPreview(tblName, keyText, txtCols(j), val, fixed)
        End If
      Next

      If Len(setSql) > 0 Then
        changedRows = changedRows + 1
        pendingSql = pendingSql & "UPDATE " & tblName & " SET " & setSql & " WHERE " & whereSql & Chr(1)
      End If
      rst.MoveNext
    Loop
    rst.Close
  End Sub

  Dim i, parts
  For i = 0 To UBound(tableDefs)
    parts = Split(tableDefs(i), "|")
    Call RepairTable(parts(0), Split(parts(1), ","), Split(parts(2), ","))
  Next

  ' Aplicar cambios solo cuando el administrador lo confirma
  Dim stmts, applied
  applied = 0
  If mode = "apply" And Len(pendingSql) > 0 Then
    stmts = Split(pendingSql, Chr(1))
    For i = 0 To UBound(stmts)
      If Len(stmts(i)) > 0 Then
        On Error Resume Next
        cnnDB.Execute stmts(i)
        If Err.Number <> 0 Then
          errCount = errCount + 1
          Err.Clear
        Else
          applied = applied + 1
        End If
        On Error Goto 0
      End If
    Next
  End If

  Call DisplayHeader(cnnDB, sid)
%>

<div class="lhd-wrapper">
  <div class="lhd-page-title" style="margin-bottom:16px">
    <span class="title-icon">&#128295;</span> Reparar texto acentuado
  </div>

  <div class="lhd-card">
    <% If tableErrors <> "" Then %>
    <div class="alert alert-error"><span class="alert-icon">&#9888;</span>
      <span>Estas tablas no se pudieron revisar:<ul style="margin:6px 0 0 18px"><% = tableErrors %></ul></span></div>
    <% End If %>
    <p>Filas revisadas: <b><% = rowsScanned %></b> &middot;
       Filas con texto da&ntilde;ado: <b><% = changedRows %></b> &middot;
       Campos a corregir: <b><% = changedFields %></b></p>

    <% If mode = "apply" Then %>
      <div class="alert alert-success">
        <span class="alert-icon">&#10004;</span>
        <span>Filas actualizadas: <% = applied %>. Errores: <% = errCount %>.
        Reinicie la aplicaci&oacute;n (IIS) para limpiar la cach&eacute; de textos.</span>
      </div>
    <% ElseIf changedRows = 0 Then %>
      <div class="alert alert-success">
        <span class="alert-icon">&#10004;</span>
        <span>No se encontr&oacute; texto da&ntilde;ado en la base de datos.</span>
      </div>
    <% Else %>
      <p>Revise la vista previa. Los cambios solo se aplican al confirmar.</p>
      <form method="post" action="repair_text.asp" onsubmit="return confirm('Se van a corregir <% = changedRows %> filas. Desea continuar?');">
        <input type="hidden" name="mode" value="apply" />
        <button type="submit" class="btn btn-primary">Aplicar cambios</button>
      </form>
    <% End If %>
  </div>

  <% If changedRows > 0 Then %>
  <div class="lhd-card" style="padding:0;overflow:hidden">
    <div style="overflow-x:auto">
      <table class="lhd-table">
        <thead>
          <tr>
            <th>Tabla</th><th>Registro</th><th>Campo</th><th>Antes</th><th>Despu&eacute;s</th>
          </tr>
        </thead>
        <tbody>
          <% = previewHtml %>
        </tbody>
      </table>
    </div>
  </div>
  <% End If %>
</div>

<%
  Call DisplayFooter(cnnDB, sid)
  cnnDB.Close
%>
</body>
</html>
