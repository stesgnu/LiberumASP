<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
%>
<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: sysinfo.asp
  Purpose:  Informacion del servidor y de la base de datos para diagnostico. Solo administradores.
  -->

  <!-- #include file = "../public.asp" -->
  <%
    ' Solo administradores con sesion iniciada
    If Session("lhd_IsAdmin") <> True Then
      Response.Redirect "default.asp"
    End If

    On Error Resume Next
    Call SetAppVariables

    Dim strConn, cnnDB, sid
    sid = GetSid
    Select Case Application("DBType")
      Case 1	' SQL Sec
        strConn = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer") & _
          ";Initial Catalog=" & Application("SQLDBase") & _
          ";uid=" & Application("SQLUser") & ";pwd=" & Application("SQLPass")
      Case 2	' SQL Integrated Sec
        strConn = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer") & _
          ";Initial Catalog=" & Application("SQLDBase") & ";Integrated Security=SSPI"
      Case 3	' Access
        strConn = "Provider=Microsoft.Jet.OLEDB.4.0" & _
          ";Data Source=" & Application("AccessPath")
      Case 4	' DSN
        strConn = "DSN=" & Application("DSN_Name")
    End Select

    ' Esta pagina abre su propia conexion para poder diagnosticar aunque la base falle
    Dim blnConnError, connErrText
    blnConnError = False
    connErrText = ""
    Err.Clear
    Set cnnDB = Server.CreateObject("ADODB.Connection")
    cnnDB.Open(strConn)
    If Err.Number <> 0 Or cnnDB.State <> 1 Then
      blnConnError = True
      connErrText = Err.Description
      Err.Clear
    End If

    ' Existe el archivo de Access (solo aplica a DBType 3)
    Dim fsFileSys, blnDBExists
    blnDBExists = True
    If Application("DBType") = 3 Then
      Set fsFileSys = Server.CreateObject("Scripting.FileSystemObject")
      blnDBExists = fsFileSys.FileExists(Application("AccessPath"))
      Err.Clear
    End If

    ' Prueba de escritura SIN modificar datos: un UPDATE que no cambia nada, dentro de una
    ' transaccion que se revierte siempre (antes se insertaba y borraba una fila en tblConfig).
    Dim blnUpdateError, updErrText
    blnUpdateError = False
    updErrText = ""
    If blnConnError Then
      blnUpdateError = True
    Else
      Err.Clear
      cnnDB.BeginTrans
      cnnDB.Execute "UPDATE tblConfig SET version = version"
      If Err.Number <> 0 Then
        blnUpdateError = True
        updErrText = Err.Description
        Err.Clear
      End If
      cnnDB.RollbackTrans
      Err.Clear
    End If

    ' Lectura de un ajuste de tblConfig, ya codificado para HTML
    Function CfgLookup(strSetting)
      On Error Resume Next
      Err.Clear
      Dim confRes
      CfgLookup = "<em>Error</em>"
      If blnConnError Then Exit Function
      Set confRes = Server.CreateObject("ADODB.Recordset")
      confRes.ActiveConnection = cnnDB
      confRes.Open("SELECT " & strSetting & " From tblConfig")
      If Err.Number <> 0 Then
        Err.Clear
        CfgLookup = "<em>Error</em>"
      Else
        If confRes.EOF Then
          CfgLookup = "<em>No existe</em>"
        Else
          If Len(confRes(strSetting) & "") = 0 Then
            CfgLookup = "&mdash;"
          Else
            CfgLookup = Server.HTMLEncode(confRes(strSetting) & "")
          End If
        End If
        confRes.Close
      End If
    End Function

    ' Fila de resultado con distintivo OK / Error
    Function StatusBadgeHtml(ok, textOk, textBad)
      If ok Then
        StatusBadgeHtml = "<span class=""sys-badge ok"">&#10004; " & textOk & "</span>"
      Else
        StatusBadgeHtml = "<span class=""sys-badge bad"">&#10006; " & textBad & "</span>"
      End If
    End Function

    ' Valor de una variable de aplicacion codificado (o raya si esta vacio)
    Function AppVal(k)
      If Len(Application(k) & "") = 0 Then
        AppVal = "&mdash;"
      Else
        AppVal = Server.HTMLEncode(Application(k) & "")
      End If
    End Function

    ' Variables del servidor que no se muestran (cookies, claves y cabeceras completas)
    Function HiddenServerVar(k)
      Select Case UCase(k)
        Case "HTTP_COOKIE", "ALL_HTTP", "ALL_RAW", "AUTH_PASSWORD", "HTTP_AUTHORIZATION", "CERT_COOKIE", "HTTP_X_ORIGINAL_URL"
          HiddenServerVar = True
        Case Else
          HiddenServerVar = False
      End Select
    End Function

    Dim dbTypeName
    Select Case Application("DBType")
      Case 1: dbTypeName = "SQL Server (usuario y clave)"
      Case 2: dbTypeName = "SQL Server (seguridad integrada)"
      Case 3: dbTypeName = "Microsoft Access"
      Case 4: dbTypeName = "DSN (ODBC)"
      Case Else: dbTypeName = "Sin definir"
    End Select
  %>

  <head>
    <title>
      <% If Not blnConnError Then Response.Write(lang(cnnDB, "HelpDesk") & "&nbsp;-&nbsp;") %>Informaci&oacute;n del sistema
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .sys-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 20px; align-items: start; margin-bottom: 20px; }
      .sys-rows { border: 1px solid var(--color-border-light, #E2E8F0); border-radius: 8px; overflow: hidden; }
      .sys-row { display: flex; justify-content: space-between; align-items: center; gap: 14px; padding: 9px 14px; font-size: 13px; }
      .sys-row + .sys-row { border-top: 1px solid var(--color-border-light, #E2E8F0); }
      .sys-row:nth-child(even) { background: #F8FAFC; }
      .sys-k { color: #64748b; flex-shrink: 0; }
      .sys-v { font-weight: 600; color: #1a202c; text-align: right; word-break: break-all; }
      .sys-badge { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; }
      .sys-badge.ok { background: var(--color-success-bg, #D4EDDA); color: var(--color-success, #1A7F45); }
      .sys-badge.bad { background: var(--color-danger-bg, #F8D7DA); color: var(--color-danger, #B02A37); }
      .sys-err { margin: 10px 0 0; font-size: 12px; color: #B02A37; word-break: break-word; }
      .sys-toolbar { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between; margin-bottom: 12px; }
      .sys-search { flex: 1 1 260px; max-width: 420px; padding: 10px 14px; border: 1px solid var(--color-border, #CBD5E0); border-radius: 10px; font-size: 14px; box-sizing: border-box; }
      .sys-search:focus { outline: none; border-color: var(--color-primary, #0057B8); box-shadow: 0 0 0 3px var(--color-primary-light, #E8F0FB); }
      .sys-note { font-size: 12px; color: #64748b; }
      .sys-srv .sys-v { font-weight: 500; font-family: Consolas, "Courier New", monospace; font-size: 12px; }
    </style>
  </head>
  <body>

    <%
      On Error Goto 0
      If Not blnConnError Then Call DisplayHeader(cnnDB, sid)
    %>

    <div class="lhd-wrapper"<% If blnConnError Then Response.Write(" style=""padding:24px""") %>>

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#8505;</span>
          Informaci&oacute;n del sistema
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="test.asp">&larr; Pruebas de configuraci&oacute;n</a>
          <a class="btn btn-secondary" href="default.asp">Men&uacute; administrativo</a>
        </div>
      </div>

      <div class="sys-grid">

        <div class="lhd-card">
          <div class="section-header">&#9989; Estado de la base de datos</div>
          <div class="sys-rows">
            <div class="sys-row"><span class="sys-k">Conexi&oacute;n</span><span class="sys-v"><% = StatusBadgeHtml(Not blnConnError, "Correcta", "Error de conexi&oacute;n") %></span></div>
            <% If Application("DBType") = 3 Then %>
            <div class="sys-row"><span class="sys-k">Archivo Access (.mdb)</span><span class="sys-v"><% = StatusBadgeHtml(blnDBExists, "Existe", "No existe") %></span></div>
            <% End If %>
            <div class="sys-row"><span class="sys-k">Escritura en la base</span><span class="sys-v"><% = StatusBadgeHtml(Not blnUpdateError, "Correcta", "Error al escribir") %></span></div>
          </div>
          <% If blnConnError And Len(connErrText) > 0 Then %><p class="sys-err"><b>Detalle:</b> <% = Server.HTMLEncode(connErrText) %></p><% End If %>
          <% If blnUpdateError And Not blnConnError And Len(updErrText) > 0 Then %><p class="sys-err"><b>Detalle:</b> <% = Server.HTMLEncode(updErrText) %></p><% End If %>
          <p class="sys-note" style="margin:10px 0 0">La prueba de escritura no modifica datos: se revierte siempre.</p>
        </div>

        <div class="lhd-card">
          <div class="section-header">&#128295; Variables de aplicaci&oacute;n</div>
          <div class="sys-rows">
            <div class="sys-row"><span class="sys-k">Tipo de base (DBType)</span><span class="sys-v"><% = Application("DBType") %> &middot; <% = dbTypeName %></span></div>
            <div class="sys-row"><span class="sys-k">SQLServer</span><span class="sys-v"><% = AppVal("SQLServer") %></span></div>
            <div class="sys-row"><span class="sys-k">SQLDBase</span><span class="sys-v"><% = AppVal("SQLDBase") %></span></div>
            <div class="sys-row"><span class="sys-k">AccessPath</span><span class="sys-v"><% = AppVal("AccessPath") %></span></div>
            <div class="sys-row"><span class="sys-k">DSN_Name</span><span class="sys-v"><% = AppVal("DSN_Name") %></span></div>
          </div>
        </div>

        <div class="lhd-card">
          <div class="section-header">&#9881; Configuraci&oacute;n del sitio</div>
          <div class="sys-rows">
            <div class="sys-row"><span class="sys-k">BaseURL</span><span class="sys-v"><% = CfgLookup("BaseURL") %></span></div>
            <div class="sys-row"><span class="sys-k">EmailType</span><span class="sys-v"><% = CfgLookup("EmailType") %></span></div>
            <div class="sys-row"><span class="sys-k">SMTPServer</span><span class="sys-v"><% = CfgLookup("SMTPServer") %></span></div>
            <div class="sys-row"><span class="sys-k">EnablePager</span><span class="sys-v"><% = CfgLookup("EnablePager") %></span></div>
            <div class="sys-row"><span class="sys-k">AuthType</span><span class="sys-v"><% = CfgLookup("AuthType") %></span></div>
            <div class="sys-row"><span class="sys-k">Versi&oacute;n</span><span class="sys-v"><% = CfgLookup("Version") %></span></div>
            <div class="sys-row"><span class="sys-k">UseInOutBoard</span><span class="sys-v"><% = CfgLookup("UseInOutBoard") %></span></div>
            <div class="sys-row"><span class="sys-k">KBFreeText</span><span class="sys-v"><% = CfgLookup("KBFreeText") %></span></div>
          </div>
        </div>

        <div class="lhd-card">
          <div class="section-header">&#128221; Servidor</div>
          <div class="sys-rows">
            <div class="sys-row"><span class="sys-k">Hora del servidor</span><span class="sys-v"><% = Server.HTMLEncode(CStr(Now)) %></span></div>
            <div class="sys-row"><span class="sys-k">Motor de script</span><span class="sys-v"><% = ScriptEngine %> <% = ScriptEngineMajorVersion %>.<% = ScriptEngineMinorVersion %></span></div>
            <div class="sys-row"><span class="sys-k">Software del servidor</span><span class="sys-v"><% = Server.HTMLEncode(Request.ServerVariables("SERVER_SOFTWARE") & "") %></span></div>
            <div class="sys-row"><span class="sys-k">Nombre del servidor</span><span class="sys-v"><% = Server.HTMLEncode(Request.ServerVariables("SERVER_NAME") & "") %></span></div>
            <div class="sys-row"><span class="sys-k">Ruta de la aplicaci&oacute;n</span><span class="sys-v"><% = Server.HTMLEncode(Request.ServerVariables("APPL_PHYSICAL_PATH") & "") %></span></div>
          </div>
        </div>

      </div>

      <div class="lhd-card">
        <div class="section-header">&#128203; Variables del servidor</div>
        <div class="sys-toolbar">
          <input type="text" id="srvFilter" class="sys-search" placeholder="Filtrar variables..." autocomplete="off">
          <span class="sys-note">Se ocultan las cookies, cabeceras completas y credenciales.</span>
        </div>
        <div class="sys-rows sys-srv" id="srvRows">
          <%
            Dim strKey
            For Each strKey In Request.ServerVariables
              If Not HiddenServerVar(strKey) Then
                Response.Write "<div class=""sys-row"" data-t=""" & Server.HTMLEncode(LCase(strKey)) & """>" & _
                  "<span class=""sys-k"">" & Server.HTMLEncode(strKey) & "</span>" & _
                  "<span class=""sys-v"">" & Server.HTMLEncode(Request.ServerVariables(strKey) & "") & "</span></div>" & vbNewLine
              End If
            Next
          %>
        </div>
      </div>

    </div>

    <script type="text/javascript">
    (function () {
      var input = document.getElementById('srvFilter');
      var rows = document.querySelectorAll('#srvRows .sys-row');
      input.oninput = function () {
        var q = input.value.toLowerCase();
        for (var i = 0; i < rows.length; i++) {
          rows[i].style.display = (!q || rows[i].textContent.toLowerCase().indexOf(q) > -1) ? '' : 'none';
        }
      };
    })();
    </script>

    <%
      If Not blnConnError Then
        Call DisplayFooter(cnnDB, sid)
        cnnDB.Close
      End If
    %>
  </body>
</html>
