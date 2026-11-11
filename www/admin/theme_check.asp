<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  Response.Buffer = TRUE
%>
<!-- #include file = "../public.asp" -->
<%
  ' Diagnostico de la paleta de colores (tabla tblTheme). Solo administradores.
  Dim cnnDB, sid
  Set cnnDB = CreateCon
  sid = GetSid
  Call CheckUser(cnnDB, sid)
  Call CheckAdmin

  Dim out, r, n, errTxt
  out = ""

  Sub AddRow(ok, label, detail)
    Dim c
    If ok Then c = "#1A7F45" Else c = "#B02A37"
    out = out & "<tr><td style=""padding:6px 10px;color:" & c & ";font-weight:700"">" & IIf2(ok, "OK", "FALLA") & "</td>" & _
      "<td style=""padding:6px 10px"">" & Server.HTMLEncode(label) & "</td>" & _
      "<td style=""padding:6px 10px;font-family:monospace;font-size:12px"">" & Server.HTMLEncode(detail & "") & "</td></tr>"
  End Sub

  On Error Resume Next

  ' 1. Conexion
  Call AddRow(True, "Tipo de base (Application(DBType))", Application("DBType") & "  (1=SQL, 2=SQL integrada, 3=Access, 4=DSN)")
  Set r = cnnDB.Execute("SELECT DB_NAME() AS dbn, SUSER_SNAME() AS usr, USER_NAME() AS dbusr")
  If Err.Number = 0 Then
    Call AddRow(True, "Base de datos / login / usuario en la base", r("dbn") & " / " & r("usr") & " / " & r("dbusr"))
    r.Close
  Else
    Call AddRow(False, "Datos de conexion (solo SQL Server)", Err.Description)
    Err.Clear
  End If

  ' 2. Existe la tabla y es legible
  Set r = cnnDB.Execute("SELECT theme_key, theme_value FROM tblTheme")
  If Err.Number <> 0 Then
    Call AddRow(False, "SELECT sobre tblTheme", Err.Description)
    Err.Clear
  Else
    n = 0
    Dim rowsTxt
    rowsTxt = ""
    Do While Not r.EOF
      n = n + 1
      rowsTxt = rowsTxt & r("theme_key") & "=" & r("theme_value") & "  "
      r.MoveNext
    Loop
    r.Close
    Call AddRow(True, "SELECT sobre tblTheme (" & n & " filas)", rowsTxt)
  End If

  ' 3. Escritura de prueba (INSERT y DELETE de una fila temporal)
  cnnDB.Execute "INSERT INTO tblTheme (theme_key, theme_value) VALUES ('zz_test', '#000000')"
  If Err.Number <> 0 Then
    Call AddRow(False, "INSERT de prueba", Err.Description)
    Err.Clear
  Else
    Call AddRow(True, "INSERT de prueba", "fila zz_test insertada")
    cnnDB.Execute "DELETE FROM tblTheme WHERE theme_key='zz_test'"
    If Err.Number <> 0 Then
      Call AddRow(False, "DELETE de prueba", Err.Description)
      Err.Clear
    Else
      Call AddRow(True, "DELETE de prueba", "fila zz_test eliminada")
    End If
  End If

  ' 4. Escritura real con la rutina de la aplicacion
  Application("lhd_theme_err") = ""
  Call SetThemeColor(cnnDB, "primary", GetThemeColor(cnnDB, "primary", "#0057B8"))
  errTxt = Application("lhd_theme_err") & ""
  Call AddRow(Len(errTxt) = 0, "SetThemeColor('primary') con el valor actual", IIf2(Len(errTxt) = 0, "sin errores", errTxt))
  Application("lhd_theme_err") = ""

  On Error Goto 0
%>
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Diagn&oacute;stico de colores</title>
  <link rel="stylesheet" type="text/css" href="../default.css" />
</head>
<body>
<% Call DisplayHeader(cnnDB, sid) %>
<div class="lhd-wrapper">
  <div class="lhd-card">
    <div class="section-header">Diagn&oacute;stico de la paleta de colores (tblTheme)</div>
    <table style="width:100%;border-collapse:collapse"><% = out %></table>
    <p style="margin-top:14px"><a class="btn btn-secondary" href="colors.asp">&larr; Volver a colores</a></p>
  </div>
</div>
<% Call DisplayFooter(cnnDB, sid) %>
</body>
</html>
