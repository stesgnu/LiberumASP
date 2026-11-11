<%@ LANGUAGE="VBScript" CODEPAGE="65001" EnableSessionState=False %>
<%
  Option Explicit
  Response.Buffer = True
  Response.ContentType = "text/css"
  Response.CharSet = "utf-8"
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>
<!-- #include file = "public.asp" -->
<%
  ' Paleta de colores configurada en admin/colors.asp, como hoja de estilos.
  ' default.css la importa, asi que TODA pagina que cargue default.css
  ' (login, usuario, representante, administracion...) usa la misma paleta.

  ' Mezcla un color #RRGGBB con blanco (pct = 0..100 de blanco)
  Function LhdTint(hex, pct)
    Dim r, g, b
    r = CLng("&H" & Mid(hex, 2, 2))
    g = CLng("&H" & Mid(hex, 4, 2))
    b = CLng("&H" & Mid(hex, 6, 2))
    r = r + (255 - r) * pct \ 100
    g = g + (255 - g) * pct \ 100
    b = b + (255 - b) * pct \ 100
    LhdTint = "#" & Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
  End Function

  ' Normalmente la paleta ya esta en memoria (la carga cualquier pagina con DisplayHeader),
  ' asi esta hoja de estilos no abre una conexion a la base en cada pagina.
  Dim cnnDB, tPrimary, tHover, tSidebar, tAccent, tBg
  If Application("lhd_theme_loaded") <> "1" Then
    Set cnnDB = CreateCon
    Call LhdThemeLoad(cnnDB)
    cnnDB.Close
  End If
  tPrimary = Application("lhd_theme_primary")
  tHover   = Application("lhd_theme_primary_hover")
  tSidebar = Application("lhd_theme_sidebar")
  tAccent  = Application("lhd_theme_accent")
  tBg      = Application("lhd_theme_bg")
  If Len(tPrimary & "") = 0 Then tPrimary = "#0057B8"
  If Len(tHover & "") = 0 Then tHover = "#003F8A"
  If Len(tSidebar & "") = 0 Then tSidebar = "#0F172A"
  If Len(tAccent & "") = 0 Then tAccent = "#0284C7"
  If Len(tBg & "") = 0 Then tBg = "#F1F5F9"

  If Not LhdIsHexColor(tPrimary) Then tPrimary = "#0057B8"
  If Not LhdIsHexColor(tHover)   Then tHover   = "#003F8A"
  If Not LhdIsHexColor(tSidebar) Then tSidebar = "#0F172A"
  If Not LhdIsHexColor(tAccent)  Then tAccent  = "#0284C7"
  If Not LhdIsHexColor(tBg)      Then tBg      = "#F1F5F9"

  Response.Write ":root {" & vbLf
  Response.Write "  --color-primary: " & tPrimary & " !important;" & vbLf
  Response.Write "  --color-primary-dark: " & tHover & " !important;" & vbLf
  Response.Write "  --color-primary-light: " & LhdTint(tPrimary, 90) & " !important;" & vbLf
  Response.Write "  --color-accent: " & tAccent & " !important;" & vbLf
  Response.Write "  --color-bg: " & tBg & " !important;" & vbLf
  Response.Write "  --lhd-primary: " & tPrimary & " !important;" & vbLf
  Response.Write "  --lhd-primary-hover: " & tHover & " !important;" & vbLf
  Response.Write "  --lhd-sidebar: " & tSidebar & " !important;" & vbLf
  Response.Write "  --lhd-accent: " & tAccent & " !important;" & vbLf
  Response.Write "  --lhd-bg: " & tBg & " !important;" & vbLf
  Response.Write "}" & vbLf
%>
