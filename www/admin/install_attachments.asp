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
  Filename: install_attachments.asp (admin)
  Purpose:  Crea la tabla tblAttachments (imagenes del editor y archivos
            adjuntos guardados en la base de datos). Es seguro ejecutarla
            varias veces: si la tabla ya existe no hace nada.
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
  <title><% = Cfg(cnnDB, "SiteName") %> - Instalar adjuntos</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="stylesheet" type="text/css" href="../default.css?v=4" />
</head>
<body>

<%
  Call DisplayHeader(cnnDB, sid)

  ' Funciones auxiliares
  Function TableExists(cnnDB, tblName)
    Dim r
    On Error Resume Next
    Set r = cnnDB.Execute("SELECT TOP 1 * FROM " & tblName)
    TableExists = (Err.Number = 0)
    Err.Clear
    On Error Goto 0
  End Function

  ' Tipo ADO de una columna (0 si no existe)
  Function ColumnType(cnnDB, sqlText, colName)
    Dim r
    On Error Resume Next
    Set r = Server.CreateObject("ADODB.Recordset")
    r.Open sqlText, cnnDB
    If Err.Number <> 0 Then
      ColumnType = 0
      Err.Clear
    Else
      ColumnType = r.Fields(colName).Type
      r.Close
    End If
    On Error Goto 0
  End Function

  Function CountRows(cnnDB, sqlText)
    Dim r
    On Error Resume Next
    Set r = cnnDB.Execute(sqlText)
    CountRows = 0
    If Err.Number = 0 Then CountRows = CLng(r(0))
    Err.Clear
    On Error Goto 0
  End Function

  Function RunSql(cnnDB, sqlText, label)
    On Error Resume Next
    cnnDB.Execute sqlText
    If Err.Number <> 0 Then
      RunSql = label & ": ERROR - " & Server.HTMLEncode(Err.Description)
    Else
      RunSql = label & ": OK"
    End If
    Err.Clear
    On Error Goto 0
  End Function

  ' Sentencias segun el tipo de base
  Dim isAccess, ddlCreate, ddlMigrate, ddlConfig, ddlDefault
  isAccess = (Application("DBType") = 3)
  If isAccess Then
    ddlCreate = "CREATE TABLE tblAttachments (id COUNTER PRIMARY KEY, problem_id INTEGER, token TEXT(64), uid TEXT(50), fname TEXT(255), ctype TEXT(100), fsize LONG, addDate DATETIME, data MEMO)"
    ddlMigrate = "ALTER TABLE tblAttachments ALTER COLUMN data MEMO"
    ddlConfig = "ALTER TABLE tblConfig ADD COLUMN AllowedAttachments TEXT(255)"
  Else
    ddlCreate = "CREATE TABLE tblAttachments (id INT IDENTITY(1,1) PRIMARY KEY, problem_id INT, token VARCHAR(64), uid VARCHAR(50), fname VARCHAR(255), ctype VARCHAR(100), fsize INT, addDate DATETIME, data VARCHAR(MAX))"
    ddlMigrate = "ALTER TABLE tblAttachments ALTER COLUMN data VARCHAR(MAX) NULL"
    ddlConfig = "ALTER TABLE tblConfig ADD AllowedAttachments VARCHAR(255) NULL"
  End If
  ddlDefault = "UPDATE tblConfig SET AllowedAttachments = 'jpg,jpeg,png,gif,webp,pdf,doc,docx,xls,xlsx,txt,zip' WHERE AllowedAttachments IS NULL"

  ' Estado actual
  Dim exists, dataType, cfgType, rowsBin, msg, confirmed
  exists = TableExists(cnnDB, "tblAttachments")
  dataType = 0
  If exists Then dataType = ColumnType(cnnDB, "SELECT data FROM tblAttachments WHERE 1=0", "data")
  cfgType = ColumnType(cnnDB, "SELECT AllowedAttachments FROM tblConfig WHERE 1=0", "AllowedAttachments")
  confirmed = (Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("confirm") = "1")

  msg = ""
  If confirmed Then
    If Not exists Then
      msg = msg & RunSql(cnnDB, ddlCreate, "Crear tabla tblAttachments") & "<br />"
      exists = True
      dataType = ColumnType(cnnDB, "SELECT data FROM tblAttachments WHERE 1=0", "data")
    End If
    ' Columna data en binario (205 = adLongVarBinary) -> texto base64
    If dataType = 205 Then
      rowsBin = CountRows(cnnDB, "SELECT COUNT(*) FROM tblAttachments WHERE data IS NOT NULL")
      If rowsBin > 0 Then
        msg = msg & "No se convirtio la columna data: hay " & rowsBin & " archivo(s) guardados en formato binario. Revise antes de continuar.<br />"
      Else
        msg = msg & RunSql(cnnDB, ddlMigrate, "Convertir columna data a texto base64") & "<br />"
      End If
    End If
    If cfgType = 0 Then
      msg = msg & RunSql(cnnDB, ddlConfig, "Agregar columna AllowedAttachments a tblConfig") & "<br />"
      cfgType = ColumnType(cnnDB, "SELECT AllowedAttachments FROM tblConfig WHERE 1=0", "AllowedAttachments")
    End If
    If cfgType <> 0 Then
      msg = msg & RunSql(cnnDB, ddlDefault, "Valor por defecto de tipos de adjuntos") & "<br />"
    End If
    dataType = ColumnType(cnnDB, "SELECT data FROM tblAttachments WHERE 1=0", "data")
  End If

  ' Resumen
  Dim okData, okCfg
  okData = (exists And (dataType = 203 Or dataType = 201))
  okCfg = (cfgType <> 0)
  If okData And okCfg Then
    msg = msg & "<b>Todo esta listo.</b> Tabla tblAttachments con columna data en texto base64 y columna AllowedAttachments en tblConfig."
  Else
    msg = msg & "<b>Pendiente:</b> "
    If Not exists Then msg = msg & "crear la tabla tblAttachments. "
    If exists And dataType = 205 Then msg = msg & "convertir la columna data a texto. "
    If Not okCfg Then msg = msg & "agregar la columna AllowedAttachments. "
  End If
%>

<div class="lhd-wrapper">
  <div class="lhd-page-title" style="margin-bottom:16px">
    <span class="title-icon">&#128190;</span> Adjuntos e im&aacute;genes en la base de datos
  </div>

  <div class="lhd-card">
    <p><% = msg %></p>
    <% If Not (okData And okCfg) Then %>
      <form method="post" action="install_attachments.asp">
        <input type="hidden" name="confirm" value="1" />
        <p><small>Sentencias que se ejecutar&aacute;n (seg&uacute;n el estado actual):</small></p>
        <pre style="white-space:pre-wrap;font-size:12px"><% = ddlCreate %>;
<% = ddlMigrate %>;
<% = ddlConfig %>;
<% = ddlDefault %>;</pre>
        <button type="submit" class="btn btn-primary">Aplicar cambios</button>
      </form>
    <% End If %>
  </div>
</div>

<%
  Call DisplayFooter(cnnDB, sid)
  cnnDB.Close
%>
</body>
</html>
