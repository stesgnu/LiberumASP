<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% 
  Option Explicit
  Response.Buffer = TRUE
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk - Color & Theme Customization
  Filename: colors.asp (admin)
  Purpose:  Allows administrators to customize the application color palette,
            sidebar color, and background theme with live preview.
  -->

  <!-- #include file = "../settings.asp" -->
  <!-- #include file = "../public.asp" -->

  <% 
    Call SetAppVariables

    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta charset="utf-8" />
  <title><%=lang(cnnDB, "HelpDesk")%> - Personalizar Colores y Apariencia</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="stylesheet" type="text/css" href="../default.css?v=4" />
  <style type="text/css">
    .color-picker-group {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 16px;
    }
    .color-picker-group input[type="color"] {
      width: 48px;
      height: 40px;
      padding: 0;
      border: 1px solid #CBD5E0;
      border-radius: 6px;
      cursor: pointer;
      background: none;
    }
    .color-picker-group input[type="text"] {
      width: 120px;
      font-family: monospace;
      font-weight: 600;
      text-transform: uppercase;
    }
    .color-preset-btn {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 8px 14px;
      border-radius: 6px;
      border: 1px solid #CBD5E0;
      background: #FFFFFF;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.15s ease;
      margin-right: 8px;
      margin-bottom: 8px;
    }
    .color-preset-btn:hover {
      border-color: #0057B8;
      box-shadow: 0 2px 6px rgba(0,0,0,0.1);
      transform: translateY(-1px);
    }
    .preset-swatch {
      width: 18px;
      height: 18px;
      border-radius: 50%;
      border: 1px solid rgba(0,0,0,0.15);
      display: inline-block;
    }
    .theme-preview-card {
      border: 1px solid #CBD5E0;
      border-radius: 12px;
      overflow: hidden;
      box-shadow: 0 4px 14px rgba(0,0,0,0.08);
      background: #F8FAFC;
    }
    .preview-header {
      padding: 12px 18px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      color: #FFFFFF;
      font-weight: 700;
      font-size: 14px;
    }
    .preview-body {
      padding: 20px;
      display: flex;
      gap: 20px;
    }
    .preview-mini-sidebar {
      width: 140px;
      border-radius: 8px;
      padding: 14px 10px;
      color: #FFFFFF;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    .preview-sidebar-item {
      padding: 6px 8px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 500;
      opacity: 0.85;
    }
    .preview-sidebar-item.active {
      opacity: 1;
      font-weight: 700;
    }
    .preview-content {
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 12px;
    }
    .preview-kpi {
      background: #FFFFFF;
      border-radius: 6px;
      padding: 10px 14px;
      border: 1px solid #E2E8F0;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }
    .preview-btn {
      padding: 6px 14px;
      border-radius: 4px;
      color: #FFFFFF;
      font-size: 12px;
      font-weight: 600;
      border: none;
      display: inline-block;
      width: fit-content;
    }
  </style>
</head>
<body>

  <%
    ' Check for administrative permissions
    Call CheckAdmin
    Call DisplayHeader(cnnDB, sid)

    Dim msgSuccess
    msgSuccess = False

    ' Process Reset to Defaults
    If Request.QueryString("reset") = "1" Then
      Call SetThemeColor(cnnDB, "primary", "#0057B8")
      Call SetThemeColor(cnnDB, "primary_hover", "#003F8A")
      Call SetThemeColor(cnnDB, "sidebar", "#0F172A")
      Call SetThemeColor(cnnDB, "accent", "#0284C7")
      Call SetThemeColor(cnnDB, "bg", "#F1F5F9")
      msgSuccess = True
    End If

    ' Process Form Submission
    If Request.Form("save") = "1" Then
      Dim fPrimary, fPrimaryHover, fSidebar, fAccent, fBg
      fPrimary      = Left(Trim(Request.Form("clr_primary")), 9)
      fPrimaryHover = Left(Trim(Request.Form("clr_primary_hover")), 9)
      fSidebar      = Left(Trim(Request.Form("clr_sidebar")), 9)
      fAccent       = Left(Trim(Request.Form("clr_accent")), 9)
      fBg           = Left(Trim(Request.Form("clr_bg")), 9)

      If Len(fPrimary) > 0 Then Call SetThemeColor(cnnDB, "primary", fPrimary)
      If Len(fPrimaryHover) > 0 Then Call SetThemeColor(cnnDB, "primary_hover", fPrimaryHover)
      If Len(fSidebar) > 0 Then Call SetThemeColor(cnnDB, "sidebar", fSidebar)
      If Len(fAccent) > 0 Then Call SetThemeColor(cnnDB, "accent", fAccent)
      If Len(fBg) > 0 Then Call SetThemeColor(cnnDB, "bg", fBg)
      msgSuccess = True
    End If

    ' Si no se pudo guardar en la base de datos, avisar con el motivo
    If Len(Application("lhd_theme_err") & "") > 0 Then
      Call LhdToast("error", "Los colores no se guardaron en la base de datos (tabla tblTheme). " & Application("lhd_theme_err"))
      msgSuccess = False
      Application("lhd_theme_err") = ""
    End If

    ' Read Current Colors
    Dim curPrimary, curPrimaryHover, curSidebar, curAccent, curBg
    curPrimary      = GetThemeColor(cnnDB, "primary", "#0057B8")
    curPrimaryHover = GetThemeColor(cnnDB, "primary_hover", "#003F8A")
    curSidebar      = GetThemeColor(cnnDB, "sidebar", "#0F172A")
    curAccent       = GetThemeColor(cnnDB, "accent", "#0284C7")
    curBg           = GetThemeColor(cnnDB, "bg", "#F1F5F9")
  %>

  <div class="lhd-wrapper">

    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:12px">
      <div class="lhd-page-title" style="margin:0">
        <span class="title-icon">&#127912;</span>
        Personalizar Colores y Apariencia
      </div>
      <div>
        <a class="btn btn-secondary btn-sm" href="default.asp">&#8592; <%=lang(cnnDB, "AdministrativeMenu")%></a>
      </div>
    </div>

    <% If msgSuccess Then %>
      <div class="alert alert-success">
        <span class="alert-icon">&#10004;</span>
        <span>&iexcl;Configuraci&oacute;n de colores guardada exitosamente! Los cambios se aplican de inmediato en toda la aplicaci&oacute;n.</span>
      </div>
    <% End If %>

    <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(360px, 1fr));gap:24px">

      <!-- Formulario de Configuracion de Colores -->
      <div class="lhd-card">
        <div class="section-header">
          &#9881; Paleta de Colores del Sistema
        </div>

        <form method="post" action="colors.asp" id="theme-form">
          <input type="hidden" name="save" value="1" />

          <!-- Paletas Predefinidas -->
          <div style="margin-bottom:24px">
            <label class="form-label" style="font-weight:700;margin-bottom:8px;display:block">
              Temas Predefinidos (1 Clic):
            </label>
            <div style="display:flex;flex-wrap:wrap">
              <button type="button" class="color-preset-btn" onclick="applyPreset('#0057B8', '#003F8A', '#0F172A', '#0284C7', '#F1F5F9')">
                <span class="preset-swatch" style="background:#0057B8"></span>
                Azul Socya (Oficial)
              </button>
              <button type="button" class="color-preset-btn" onclick="applyPreset('#4F46E5', '#3730A3', '#1E1B4B', '#6366F1', '#F8FAFC')">
                <span class="preset-swatch" style="background:#4F46E5"></span>
                &Iacute;ndigo Moderno
              </button>
              <button type="button" class="color-preset-btn" onclick="applyPreset('#059669', '#047857', '#064E3B', '#10B981', '#F0FDF4')">
                <span class="preset-swatch" style="background:#059669"></span>
                Verde Esmeralda
              </button>
              <button type="button" class="color-preset-btn" onclick="applyPreset('#2563EB', '#1D4ED8', '#18181B', '#3B82F6', '#F1F5F9')">
                <span class="preset-swatch" style="background:#18181B"></span>
                Pizarra / Modo Oscuro
              </button>
              <button type="button" class="color-preset-btn" onclick="applyPreset('#9F1239', '#881337', '#3B0715', '#E11D48', '#FFF1F2')">
                <span class="preset-swatch" style="background:#9F1239"></span>
                Vino Tinto / Carmes&iacute;
              </button>
            </div>
          </div>

          <hr style="border:0;border-top:1px solid #E2E8F0;margin:20px 0" />

          <!-- Selectores Personalizados -->
          <div class="form-group">
            <label class="form-label">Color Principal (Botones, Enlaces, Elementos activos):</label>
            <div class="color-picker-group">
              <input type="color" id="picker-primary" value="<%= curPrimary %>" oninput="syncColor('primary', this.value)" />
              <input type="text" class="form-control" id="text-primary" name="clr_primary" value="<%= curPrimary %>" oninput="syncPicker('primary', this.value)" />
              <span style="font-size:12px;color:var(--color-muted)">Utilizado en acciones principales y selecciones</span>
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Color Principal Hover (Efecto al pasar el cursor):</label>
            <div class="color-picker-group">
              <input type="color" id="picker-hover" value="<%= curPrimaryHover %>" oninput="syncColor('hover', this.value)" />
              <input type="text" class="form-control" id="text-hover" name="clr_primary_hover" value="<%= curPrimaryHover %>" oninput="syncPicker('hover', this.value)" />
              <span style="font-size:12px;color:var(--color-muted)">Versi&oacute;n ligeramente m&aacute;s oscura</span>
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Barra Lateral de Navegaci&oacute;n (Sidebar Background):</label>
            <div class="color-picker-group">
              <input type="color" id="picker-sidebar" value="<%= curSidebar %>" oninput="syncColor('sidebar', this.value)" />
              <input type="text" class="form-control" id="text-sidebar" name="clr_sidebar" value="<%= curSidebar %>" oninput="syncPicker('sidebar', this.value)" />
              <span style="font-size:12px;color:var(--color-muted)">Color de fondo del men&uacute; lateral izquierdo</span>
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Color de Acento (Badges, Destacados y Gradiantes):</label>
            <div class="color-picker-group">
              <input type="color" id="picker-accent" value="<%= curAccent %>" oninput="syncColor('accent', this.value)" />
              <input type="text" class="form-control" id="text-accent" name="clr_accent" value="<%= curAccent %>" oninput="syncPicker('accent', this.value)" />
              <span style="font-size:12px;color:var(--color-muted)">Color secundario para insignias y resaltados</span>
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Fondo de la Aplicaci&oacute;n (Canvas / Background):</label>
            <div class="color-picker-group">
              <input type="color" id="picker-bg" value="<%= curBg %>" oninput="syncColor('bg', this.value)" />
              <input type="text" class="form-control" id="text-bg" name="clr_bg" value="<%= curBg %>" oninput="syncPicker('bg', this.value)" />
              <span style="font-size:12px;color:var(--color-muted)">Color de fondo de las p&aacute;ginas</span>
            </div>
          </div>

          <div style="display:flex;gap:12px;margin-top:24px;flex-wrap:wrap">
            <button type="submit" class="btn btn-primary">
              &#128190; Guardar Colores
            </button>
            <a href="colors.asp?reset=1" class="btn btn-secondary" onclick="return confirm('&iquest;Restablecer todos los colores a los valores oficiales por defecto?')">
              &#8635; Restablecer Predeterminados
            </a>
          </div>
        </form>
      </div>

      <!-- Vista Previa en Vivo -->
      <div class="lhd-card">
        <div class="section-header">
          &#128065; Vista Previa en Tiempo Real
        </div>

        <p style="font-size:13px;color:var(--color-muted);margin-bottom:16px">
          Esta maqueta refleja de forma interactiva la apariencia que tendr&aacute; la barra lateral, las tarjetas KPI, los botones y las tablas con los colores seleccionados.
        </p>

        <div class="theme-preview-card" id="preview-container">
          <div class="preview-header" id="pv-header" style="background:<%= curPrimary %>">
            <span>&#128203; Socya Help Desk</span>
            <span style="font-size:11px;opacity:0.85">● En l&iacute;nea</span>
          </div>
          <div class="preview-body" id="pv-body" style="background:<%= curBg %>">
            <div class="preview-mini-sidebar" id="pv-sidebar" style="background:<%= curSidebar %>">
              <div style="font-size:10px;text-transform:uppercase;letter-spacing:1px;opacity:0.6;margin-bottom:4px">Men&uacute;</div>
              <div class="preview-sidebar-item active" id="pv-sb-item1" style="background:<%= curPrimary %>">&#128202; Panel</div>
              <div class="preview-sidebar-item">&#128203; Tickets</div>
              <div class="preview-sidebar-item">&#128269; Buscar</div>
              <div class="preview-sidebar-item">&#9881; Ajustes</div>
            </div>
            <div class="preview-content">
              <div class="preview-kpi" style="border-left:4px solid <%= curPrimary %>" id="pv-kpi1">
                <div>
                  <div style="font-size:11px;color:#718096;font-weight:600">ABIERTOS</div>
                  <div style="font-size:22px;font-weight:700;color:#1A202C">12</div>
                </div>
                <span style="font-size:20px">&#128194;</span>
              </div>
              <div class="preview-kpi" style="border-left:4px solid <%= curAccent %>" id="pv-kpi2">
                <div>
                  <div style="font-size:11px;color:#718096;font-weight:600">RESUELTOS</div>
                  <div style="font-size:22px;font-weight:700;color:#1A202C">491</div>
                </div>
                <span style="font-size:20px">&#10004;</span>
              </div>
              <div style="display:flex;gap:8px;margin-top:6px">
                <button type="button" class="preview-btn" id="pv-btn" style="background:<%= curPrimary %>">
                  + Nuevo Ticket
                </button>
                <span class="badge" id="pv-badge" style="background:<%= curAccent %>;color:#FFFFFF;padding:4px 8px;font-size:11px">
                  Prioridad Normal
                </span>
              </div>
            </div>
          </div>
        </div>

      </div>

    </div>

  </div>

  <script type="text/javascript">
    function syncColor(type, value) {
      var txt = document.getElementById("text-" + type);
      if (txt) txt.value = value;
      updatePreview();
    }
    function syncPicker(type, value) {
      if (value.length === 7 && value.indexOf("#") === 0) {
        var pkr = document.getElementById("picker-" + type);
        if (pkr) pkr.value = value;
        updatePreview();
      }
    }
    function applyPreset(p, h, s, a, b) {
      document.getElementById("picker-primary").value = p;
      document.getElementById("text-primary").value = p;

      document.getElementById("picker-hover").value = h;
      document.getElementById("text-hover").value = h;

      document.getElementById("picker-sidebar").value = s;
      document.getElementById("text-sidebar").value = s;

      document.getElementById("picker-accent").value = a;
      document.getElementById("text-accent").value = a;

      document.getElementById("picker-bg").value = b;
      document.getElementById("text-bg").value = b;

      updatePreview();
    }
    function updatePreview() {
      var p = document.getElementById("text-primary").value;
      var s = document.getElementById("text-sidebar").value;
      var a = document.getElementById("text-accent").value;
      var b = document.getElementById("text-bg").value;

      var pvHeader = document.getElementById("pv-header");
      if (pvHeader) pvHeader.style.background = p;

      var pvSidebar = document.getElementById("pv-sidebar");
      if (pvSidebar) pvSidebar.style.background = s;

      var pvItem = document.getElementById("pv-sb-item1");
      if (pvItem) pvItem.style.background = p;

      var pvBody = document.getElementById("pv-body");
      if (pvBody) pvBody.style.background = b;

      var pvKpi1 = document.getElementById("pv-kpi1");
      if (pvKpi1) pvKpi1.style.borderLeftColor = p;

      var pvKpi2 = document.getElementById("pv-kpi2");
      if (pvKpi2) pvKpi2.style.borderLeftColor = a;

      var pvBtn = document.getElementById("pv-btn");
      if (pvBtn) pvBtn.style.background = p;

      var pvBadge = document.getElementById("pv-badge");
      if (pvBadge) pvBadge.style.background = a;
    }
  </script>

  <%
    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close
  %>

</body>
</html>

