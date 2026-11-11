  <!--	#include file = "settings.asp" -->
<SCRIPT LANGUAGE=VBScript RUNAT=SERVER>

'  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
'  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
'  Please view the license.html file for the full GNU General Public License.

'  Filename: public.asp
'  Date:     $Date: 2002/08/28 15:32:12 $
'  Version:  $Revision: 1.53.2.4.2.2 $
'  Purpose:  This file holds the public procs and functions used
'            throughout the site.  Each page has an include for
'            this file.

  
' ################################
' PAGE START CODE
' ################################
' Declare Constants
CONST adOpenStatic = 3
CONST adOpenForwardOnly = 0

CONST lhdDateOnly = 0
CONST lhdDateTime = 1

CONST lhdAddSQLDelim = 0
CONST lhdNoSQLDelim = 1



' TEXTO ACENTUADO (UTF-8)
' Todo el texto que llega desde un formulario o sale de la base de
' datos pasa por estas funciones. Este archivo solo contiene ASCII:
' los caracteres acentuados se escriben con ChrW, para que no
' dependan de la pagina de codigos con la que se interprete el archivo.
'
' Cada fila de LhdCharTable: byte1, byte2, caracter correcto, nombre de entidad.
' Un texto UTF-8 que se leyo como ANSI (Windows-1252) queda como dos
' caracteres (por ejemplo a con tilde -> dos caracteres, uno de ellos U+00C3).
' La segunda columna tiene dos variantes en el rango 0x80-0x9F
' (Latin-1 y Windows-1252).
Function LhdCharTable()
  LhdCharTable = Array( _
    Array(&HC3, &HA1, &HE1, "aacute"), Array(&HC3, &HA9, &HE9, "eacute"), _
    Array(&HC3, &HAD, &HED, "iacute"), Array(&HC3, &HB3, &HF3, "oacute"), _
    Array(&HC3, &HBA, &HFA, "uacute"), Array(&HC3, &HB1, &HF1, "ntilde"), _
    Array(&HC3, &HBC, &HFC, "uuml"),   Array(&HC3, &H81, &HC1, "Aacute"), _
    Array(&HC3, &H89, &HC9, "Eacute"), Array(&HC3, &H2030, &HC9, "Eacute"), _
    Array(&HC3, &H8D, &HCD, "Iacute"), Array(&HC3, &H93, &HD3, "Oacute"), _
    Array(&HC3, &H201C, &HD3, "Oacute"), Array(&HC3, &H9A, &HDA, "Uacute"), _
    Array(&HC3, &H161, &HDA, "Uacute"), Array(&HC3, &H91, &HD1, "Ntilde"), _
    Array(&HC3, &H2018, &HD1, "Ntilde"), Array(&HC3, &H9C, &HDC, "Uuml"), _
    Array(&HC3, &H153, &HDC, "Uuml"),  Array(&HC2, &HBF, &HBF, "iquest"), _
    Array(&HC2, &HA1, &HA1, "iexcl") _
  )
End Function

' LhdRepairText: corrige texto UTF-8 mal interpretado.
' useEntities = True  -> salida HTML segura (tildes como entidades &aacute; ...)
' useEntities = False -> salida con los caracteres Unicode reales (para guardar en la base)
Function LhdRepairText(strInput, useEntities)
  Dim s, t, i, r
  If IsNull(strInput) Or IsEmpty(strInput) Then
    LhdRepairText = ""
    Exit Function
  End If
  s = CStr(strInput)
  t = LhdCharTable()

  ' 1) Secuencias UTF-8 leidas como ANSI
  For i = 0 To UBound(t)
    r = t(i)
    If useEntities Then
      s = Replace(s, ChrW(r(0)) & ChrW(r(1)), "&" & r(3) & ";")
    Else
      s = Replace(s, ChrW(r(0)) & ChrW(r(1)), ChrW(r(2)))
    End If
  Next

  If useEntities Then
    ' 2) Caracteres acentuados sueltos -> entidades HTML
    For i = 0 To UBound(t)
      r = t(i)
      s = Replace(s, ChrW(r(2)), "&" & r(3) & ";")
    Next
    ' 3) Evitar doble codificacion (&amp;aacute; -> &aacute;)
    For i = 0 To UBound(t)
      r = t(i)
      s = Replace(s, "&amp;" & r(3) & ";", "&" & r(3) & ";")
    Next
  End If

  LhdRepairText = s
End Function

' FixTildes: texto seguro para mostrar en HTML (texto, atributos, opciones)
Function FixTildes(strInput)
  FixTildes = LhdRepairText(strInput, True)
End Function

' CleanInputText: normaliza texto recibido de un formulario antes de guardarlo
Function CleanInputText(strInput)
  CleanInputText = LhdRepairText(strInput, False)
End Function

' HTML ENRIQUECIDO (editor WYSIWYG) Y ADJUNTOS
' Lo que escribe el editor se limpia con una lista blanca antes de
' guardarse y antes de mostrarse: no se permiten scripts, eventos
' (onclick...), ni enlaces javascript:.

' LhdTagAllowed: etiquetas HTML permitidas en descripciones y soluciones
Function LhdTagAllowed(tagName)
  LhdTagAllowed = (InStr(",p,br,strong,b,em,i,u,s,ul,ol,li,a,img,h1,h2,h3,blockquote,pre,code,span,div,table,thead,tbody,tr,td,th,", "," & LCase(tagName) & ",") > 0)
End Function

' LhdSafeStyle: conserva solo propiedades de estilo de texto inofensivas
Function LhdSafeStyle(styleText)
  Dim parts, i, decl, prop, val, out
  out = ""
  parts = Split(styleText, ";")
  For i = 0 To UBound(parts)
    decl = Trim(parts(i))
    If InStr(decl, ":") > 0 Then
      prop = LCase(Trim(Left(decl, InStr(decl, ":") - 1)))
      val = Trim(Mid(decl, InStr(decl, ":") + 1))
      If InStr(",color,background-color,text-align,font-size,font-weight,font-style,text-decoration,", "," & prop & ",") > 0 _
         And InStr(LCase(val), "url") = 0 And InStr(LCase(val), "expression") = 0 _
         And InStr(val, "<") = 0 And InStr(val, """") = 0 Then
        If Len(out) > 0 Then out = out & "; "
        out = out & prop & ": " & val
      End If
    End If
  Next
  LhdSafeStyle = out
End Function

' LhdCleanAttrs: devuelve solo los atributos permitidos, ya limpios
Function LhdCleanAttrs(tagName, attrs)
  Dim re, mc, m, name, val, out, t
  t = LCase(tagName)
  out = ""
  Set re = New RegExp
  re.Global = True
  re.IgnoreCase = True
  re.Pattern = "([a-zA-Z\-:]+)\s*=\s*(""([^""]*)""|'([^']*)'|([^\s""'>]+))"
  Set mc = re.Execute(attrs)
  For Each m In mc
    name = LCase(m.SubMatches(0))
    val = m.SubMatches(2) & m.SubMatches(3) & m.SubMatches(4)
    Select Case True
      Case (name = "href") And (t = "a")
        If LhdIsSafeUrl(val) Then out = out & " href=""" & LhdAttrEsc(val) & """"
      Case (name = "src") And (t = "img")
        If LhdIsSafeImageSrc(val) Then out = out & " src=""" & LhdAttrEsc(val) & """"
      Case (name = "alt") And (t = "img")
        out = out & " alt=""" & LhdAttrEsc(val) & """"
      Case (name = "width" Or name = "height") And (t = "img")
        If val Like "#*" Or IsNumeric(val) Then
          If IsNumeric(val) Then out = out & " " & name & "=""" & CStr(CLng(val)) & """"
        End If
      Case (name = "class") And (t = "span" Or t = "p" Or t = "div" Or t = "li" Or t = "ol" Or t = "ul" Or t = "h1" Or t = "h2" Or t = "h3" Or t = "pre")
        If LhdIsSafeClass(val) Then out = out & " class=""" & LhdAttrEsc(val) & """"
      Case (name = "style")
        val = LhdSafeStyle(val)
        If Len(val) > 0 Then out = out & " style=""" & LhdAttrEsc(val) & """"
      Case (name = "colspan" Or name = "rowspan") And (t = "td" Or t = "th")
        If IsNumeric(val) Then out = out & " " & name & "=""" & CStr(CLng(val)) & """"
    End Select
  Next
  If t = "a" And InStr(out, " href=") > 0 Then out = out & " rel=""noopener noreferrer"" target=""_blank"""
  LhdCleanAttrs = out
End Function

Function LhdAttrEsc(v)
  Dim s
  s = CStr(v & "")
  s = Replace(s, "&", "&amp;")
  s = Replace(s, """", "&quot;")
  s = Replace(s, "<", "&lt;")
  s = Replace(s, ">", "&gt;")
  LhdAttrEsc = s
End Function

Function LhdIsSafeUrl(v)
  Dim re
  Set re = New RegExp
  re.IgnoreCase = True
  re.Pattern = "^(https?://[^\s""'<>]+|mailto:[^\s""'<>]+)$"
  LhdIsSafeUrl = re.Test(v)
End Function

' Imagenes: solo las guardadas en la base (attach/file.asp?id=N) o data:image base64
Function LhdIsSafeImageSrc(v)
  Dim re
  Set re = New RegExp
  re.IgnoreCase = True
  re.Pattern = "^(([^\s""'<>?]*/)?attach/file\.asp\?id=[0-9]+|data:image/(png|jpeg|jpg|gif|webp);base64,[A-Za-z0-9+/=]+)$"
  LhdIsSafeImageSrc = re.Test(v)
End Function

Function LhdIsSafeClass(v)
  Dim re
  Set re = New RegExp
  re.IgnoreCase = True
  re.Pattern = "^(ql-[a-z0-9\-]+)( ql-[a-z0-9\-]+)*$"
  LhdIsSafeClass = re.Test(v)
End Function

' LhdSanitizeHtml: deja solo etiquetas y atributos de la lista blanca
Function LhdSanitizeHtml(ByVal s)
  Dim re, mc, m, out, pos, tagClose, tagName, tagAttrs, cleanTag, between
  If IsNull(s) Or IsEmpty(s) Then
    LhdSanitizeHtml = ""
    Exit Function
  End If
  s = CStr(s)

  ' 1) Quitar bloques peligrosos completos y comentarios
  Set re = New RegExp
  re.Global = True
  re.IgnoreCase = True
  re.Pattern = "<" & "!--[\s\S]*?" & "--" & ">"
  s = re.Replace(s, "")
  re.Pattern = "<(script|style|iframe|object|embed|form|textarea|select|button|svg|math)\b[^>]*>[\s\S]*?<\/\1\s*>"
  s = re.Replace(s, "")
  re.Pattern = "<\/?(script|style|iframe|object|embed|form|textarea|select|button|input|meta|link|svg|math)\b[^>]*>"
  s = re.Replace(s, "")
  s = re.Replace(s, "")

  ' 2) Recorrer etiquetas y reconstruirlas limpias
  Set re = New RegExp
  re.Global = True
  re.Pattern = "<(\/?)([a-zA-Z][a-zA-Z0-9]*)([^<>]*)>"
  Set mc = re.Execute(s)
  out = ""
  pos = 1
  For Each m In mc
    between = Mid(s, pos, m.FirstIndex + 1 - pos)
    ' Un "<" suelto (p. ej. una etiqueta sin cerrar) podria unirse con la siguiente etiqueta
    ' permitida y formar una etiqueta con eventos: se escapa siempre.
    between = Replace(between, "<", "&lt;")
    out = out & between
    tagClose = m.SubMatches(0)
    tagName = m.SubMatches(1)
    tagAttrs = m.SubMatches(2)
    cleanTag = ""
    If LhdTagAllowed(tagName) Then
      If tagClose = "/" Then
        cleanTag = "</" & LCase(tagName) & ">"
      ElseIf LCase(tagName) = "br" Then
        cleanTag = "<br />"
      ElseIf LCase(tagName) = "img" Then
        cleanTag = "<img" & LhdCleanAttrs(tagName, tagAttrs) & " />"
      Else
        cleanTag = "<" & LCase(tagName) & LhdCleanAttrs(tagName, tagAttrs) & ">"
      End If
    End If
    out = out & cleanTag
    pos = m.FirstIndex + m.Length + 1
  Next
  out = out & Replace(Mid(s, pos), "<", "&lt;")

  LhdSanitizeHtml = out
End Function

' LhdRichToDisplay: texto listo para mostrar en HTML (texto plano antiguo o HTML del editor)
Function LhdRichToDisplay(ByVal s)
  Dim t
  If IsNull(s) Or IsEmpty(s) Then
    LhdRichToDisplay = ""
    Exit Function
  End If
  t = CStr(s)
  If InStr(t, "<") = 0 Then
    t = Server.HTMLEncode(t)
    t = Replace(t, vbCrLf, "<br />")
    t = Replace(t, vbLf, "<br />")
    LhdRichToDisplay = LhdRepairText(t, True)
  Else
    LhdRichToDisplay = LhdRepairText(LhdSanitizeHtml(t), True)
  End If
End Function

' LhdRichHasContent: True si hay texto o alguna imagen
Function LhdRichHasContent(ByVal s)
  Dim re, t
  t = CStr(s & "")
  If InStr(LCase(t), "<img") > 0 Then
    LhdRichHasContent = True
    Exit Function
  End If
  Set re = New RegExp
  re.Global = True
  re.Pattern = "<[^>]*>"
  t = re.Replace(t, " ")
  t = Replace(t, "&nbsp;", " ")
  LhdRichHasContent = (Len(Trim(t)) > 0)
End Function

' LhdAttachCtype: tipo MIME permitido segun la extension (nunca HTML ni scripts)
Function LhdAttachCtype(ext)
  Select Case LCase(ext)
    Case "jpg", "jpeg": LhdAttachCtype = "image/jpeg"
    Case "png": LhdAttachCtype = "image/png"
    Case "gif": LhdAttachCtype = "image/gif"
    Case "webp": LhdAttachCtype = "image/webp"
    Case "pdf": LhdAttachCtype = "application/pdf"
    Case "txt": LhdAttachCtype = "text/plain"
    Case "doc": LhdAttachCtype = "application/msword"
    Case "docx": LhdAttachCtype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    Case "xls": LhdAttachCtype = "application/vnd.ms-excel"
    Case "xlsx": LhdAttachCtype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    Case "zip": LhdAttachCtype = "application/zip"
    Case "csv": LhdAttachCtype = "text/csv"
    Case "rtf": LhdAttachCtype = "application/rtf"
    Case "ppt": LhdAttachCtype = "application/vnd.ms-powerpoint"
    Case "pptx": LhdAttachCtype = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    Case "odt": LhdAttachCtype = "application/vnd.oasis.opendocument.text"
    Case "ods": LhdAttachCtype = "application/vnd.oasis.opendocument.spreadsheet"
    Case "7z": LhdAttachCtype = "application/x-7z-compressed"
    Case Else: LhdAttachCtype = "application/octet-stream"
  End Select
End Function

Function LhdIsImageCtype(ctype)
  LhdIsImageCtype = (Left(LCase(ctype), 6) = "image/")
End Function

' LhdLinkAttachments: asocia los adjuntos subidos con este token al ticket recien creado.
' Si la tabla tblAttachments aun no existe no interrumpe el guardado del ticket.
Sub LhdLinkAttachments(cnnDB, token, problemId)
  If Len(token) = 0 Or Not IsNumeric(problemId) Then Exit Sub
  On Error Resume Next
  cnnDB.Execute "UPDATE tblAttachments SET problem_id=" & CLng(problemId) & " WHERE token='" & Replace(token, "'", "''") & "' AND problem_id=0"
  Err.Clear
  On Error Goto 0
End Sub

' LhdAttachmentList: lista de adjuntos guardados en la base para un ticket.
' Si la tabla tblAttachments aun no existe devuelve una cadena vacia.
Function LhdAttachmentList(cnnDB, problemId)
  Dim rst, out, baseUrl, kb, sqlText
  out = ""
  If CLng(problemId) <= 0 Then
    LhdAttachmentList = ""
    Exit Function
  End If
  baseUrl = GetAppPath()
  sqlText = "SELECT id, fname, fsize FROM tblAttachments WHERE problem_id=" & CLng(problemId) & " ORDER BY id"
  On Error Resume Next
  Set rst = Server.CreateObject("ADODB.Recordset")
  rst.Open sqlText, cnnDB
  If Err.Number <> 0 Then
    Err.Clear
    On Error Goto 0
    LhdAttachmentList = ""
    Exit Function
  End If
  On Error Goto 0
  Do While Not rst.EOF
    kb = Round(Val(rst("fsize") & "") / 1024)
    out = out & "<li><a href=""" & baseUrl & "/attach/file.asp?id=" & rst("id") & """ target=""_blank"" rel=""noopener"">" & _
      LhdRepairText(Server.HTMLEncode(rst("fname") & ""), True) & "</a> <small>(" & kb & " KB)</small></li>" & vbNewLine
    rst.MoveNext
  Loop
  rst.Close
  If Len(out) > 0 Then out = "<ul class=""lhd-attach-list"">" & out & "</ul>"
  LhdAttachmentList = out
End Function

' LhdBlockedExt: extensiones que nunca se aceptan (ejecutables, scripts, HTML/SVG)
Function LhdBlockedExt(ext)
  LhdBlockedExt = (InStr(",exe,bat,cmd,com,scr,msi,msp,dll,vbs,vbe,js,jse,wsf,wsh,ps1,hta,asp,aspx,asa,asax,cer,php,jsp,html,htm,shtml,svg,xml,xhtml,cgi,pl,reg,lnk,cpl,jar,sh,mht,mhtml,url,ocx,sys,inf,application,gadget,pif,vb,", "," & LCase(ext) & ",") > 0)
End Function

' LhdCleanExtList: "PDF, .doc;xlsx" -> "pdf,doc,xlsx" (sin duplicados ni extensiones bloqueadas)
Function LhdCleanExtList(raw)
  Dim parts, i, e, out, re
  Set re = New RegExp
  re.Pattern = "^[a-z0-9]{1,10}$"
  parts = Split(Replace(LCase(raw & ""), ";", ","), ",")
  out = ""
  For i = 0 To UBound(parts)
    e = Replace(Trim(parts(i)), ".", "")
    If Len(e) > 0 Then
      If re.Test(e) And Not LhdBlockedExt(e) Then
        If InStr("," & out & ",", "," & e & ",") = 0 Then
          If Len(out) > 0 Then out = out & ","
          out = out & e
        End If
      End If
    End If
  Next
  LhdCleanExtList = out
End Function

' LhdAttachTypesRaw: valor guardado en admin (columna AllowedAttachments).
' Si la columna aun no existe devuelve cadena vacia.
Function LhdAttachTypesRaw(cnnDB)
  Dim rst, v
  v = ""
  On Error Resume Next
  Set rst = Server.CreateObject("ADODB.Recordset")
  rst.Open "SELECT AllowedAttachments FROM tblConfig", cnnDB
  If Err.Number = 0 Then
    If Not rst.EOF Then v = rst("AllowedAttachments") & ""
    rst.Close
  End If
  Err.Clear
  On Error Goto 0
  LhdAttachTypesRaw = v
End Function

' LhdAllowedExts: lista limpia de extensiones que acepta la aplicacion
Function LhdAllowedExts(cnnDB)
  Dim s
  s = LhdCleanExtList(LhdAttachTypesRaw(cnnDB))
  If Len(s) = 0 Then s = LhdCleanExtList("jpg,jpeg,png,gif,webp,pdf,doc,docx,xls,xlsx,txt,zip")
  LhdAllowedExts = s
End Function

' LhdMaxAttachBytes: tamano maximo de un adjunto (MaxImageSize, en bytes)
Function LhdMaxAttachBytes(cnnDB)
  Dim v
  v = SafeCLng(Cfg(cnnDB, "MaxImageSize"))
  If v <= 0 Then v = 2000000
  LhdMaxAttachBytes = v
End Function

' LhdAcceptAttr: valor del atributo accept de los selectores de archivo (.pdf,.doc,...)
Function LhdAcceptAttr(cnnDB)
  Dim parts, i, out
  parts = Split(LhdAllowedExts(cnnDB), ",")
  out = ""
  For i = 0 To UBound(parts)
    If Len(out) > 0 Then out = out & ","
    out = out & "." & parts(i)
  Next
  LhdAcceptAttr = out
End Function

' HasDateFormatColumn:
' Checks once per application lifetime if tblUsers has the dateformat column
Function HasDateFormatColumn(cnnDB)
  HasDateFormatColumn = False
End Function

' HasDueDateColumn:
' Checks once per application lifetime if problems has the due_date column
Function HasDueDateColumn(cnnDB)
  HasDueDateColumn = False
End Function

' GetAppPath:
' Dynamically resolves the virtual directory root (e.g. /Helpdesk_JS) from the request URL
Function GetAppPath()
  Dim scriptName, pos
  scriptName = Request.ServerVariables("SCRIPT_NAME")
  pos = InStr(LCase(scriptName), "/rep/")
  If pos = 0 Then pos = InStr(LCase(scriptName), "/user/")
  If pos = 0 Then pos = InStr(LCase(scriptName), "/admin/")
  If pos = 0 Then pos = InStr(LCase(scriptName), "/kb/")
  If pos = 0 Then pos = InStr(LCase(scriptName), "/inout/")
  If pos > 0 Then
    GetAppPath = Left(scriptName, pos - 1)
  Else
    pos = InStrRev(scriptName, "/")
    If pos > 1 Then
      GetAppPath = Left(scriptName, pos - 1)
    Else
      GetAppPath = ""
    End If
  End If
End Function


'#################################

' Cfg:
' Returns a configuration setting from tblConfig
Dim lhdCfgCache   ' cache por peticion: valor de cada ajuste de tblConfig
Dim lhdQryCount, lhdQryMs, lhdPageStart   ' medicion por peticion: consultas, milisegundos en BD, inicio

' LhdCacheReset: vacia las caches por peticion (se llama tras cualquier INSERT/UPDATE/DELETE)
Sub LhdCacheReset()
  If IsObject(lhdCfgCache) Then lhdCfgCache.RemoveAll
  If IsObject(lhdUsrCache) Then lhdUsrCache.RemoveAll
End Sub

Function Cfg(cnnDB, strSetting)
  Err.Clear
  Dim confRes, cfgKey, cfgVal
  cfgKey = LCase(Trim(strSetting))
  If Not IsObject(lhdCfgCache) Then
    Set lhdCfgCache = Server.CreateObject("Scripting.Dictionary")
  End If
  ' Una misma pagina lee varias veces el mismo ajuste (cabecera, pie, idioma...): una sola consulta
  If lhdCfgCache.Exists(cfgKey) Then
    Cfg = lhdCfgCache(cfgKey)
    Exit Function
  End If
  Set confRes = SQLQuery(cnnDB, "SELECT " & strSetting & " From tblConfig")
  If confRes.EOF Then
    Cfg = vbNull
    Call DisplayError(3, strSetting & " " & lang(cnnDB, "isaninvalidsetting") & ".")
  Else
    cfgVal = confRes(strSetting)
    Cfg = cfgVal
    lhdCfgCache(cfgKey) = cfgVal
  End If
  confRes.Close
End Function


' Usr:
' Returns selected user information
' Nombre del autor de una nota (tblNotes.uid guarda el login). Si no se encuentra, devuelve el login.
Function NoteAuthorName(cnnDB, authorUid)
  On Error Resume Next
  Dim rec, nm
  nm = ""
  Set rec = Server.CreateObject("ADODB.Recordset")
  rec.ActiveConnection = cnnDB
  rec.Open "SELECT fname FROM tblUsers WHERE uid='" & Replace(authorUid & "", "'", "''") & "'"
  If Err.Number = 0 Then
    If Not rec.EOF Then nm = rec("fname") & ""
    rec.Close
  End If
  Err.Clear
  On Error Goto 0
  If Len(Trim(nm)) > 0 Then
    NoteAuthorName = nm
  Else
    NoteAuthorName = authorUid & ""
  End If
End Function

Dim lhdUsrCache   ' cache por peticion: "sid|columna" -> valor

Function Usr(cnnDB, sid, strColumn)
  On Error Resume Next
  Usr = ""
  sid = Cint(sid)
  Dim usrRec, usrKey, usrVal
  usrKey = sid & "|" & LCase(Trim(strColumn))
  If Not IsObject(lhdUsrCache) Then
    Set lhdUsrCache = Server.CreateObject("Scripting.Dictionary")
  End If
  If lhdUsrCache.Exists(usrKey) Then
    Usr = lhdUsrCache(usrKey)
    Exit Function
  End If
  Set usrRec = Server.CreateObject("ADODB.Recordset")
  usrRec.ActiveConnection = cnnDB
  lhdQryCount = lhdQryCount + 1
  usrRec.Open "SELECT " & strColumn & " FROM tblUsers WHERE sid=" & sid
  If Err.Number = 0 Then
    usrVal = ""
    If Not usrRec.EOF Then
      usrVal = usrRec(0)
    End If
    Usr = usrVal
    usrRec.Close
    lhdUsrCache(usrKey) = usrVal
  Else
    Err.Clear
    Usr = ""
  End If
  Set usrRec = Nothing
  On Error Goto 0
End Function



' Lang:
' Returns select Language string
Dim lhdLangDbFallbacks   ' contador por peticion (vacio = 0)

Dim lhdLangDict, lhdLangDictId   ' diccionario por peticion (clave -> texto)

Function Lang(cnnDB, variable)
  Dim intLangSetting, sid
  sid = GetSid
  If sid = 0 Then
    intLangSetting = Cfg(cnnDB, "DefaultLanguage")
  Else
    intLangSetting = Session("lhd_LanguageID")
    If Cint(intLangSetting) < 1 Then
      intLangSetting = Usr(cnnDB, sid, "[language]")
      If IsNull(intLangSetting) Or IsEmpty(intLangSetting) Or intLangSetting = 0 Then
        intLangSetting = Cfg(cnnDB, "DefaultLanguage")
      End If
      Session("lhd_LanguageID") = intLangSetting
    End If
  End If

  ' Las cadenas del idioma se guardan en Application (matriz). Leer esa matriz copia todo su
  ' contenido, asi que se hace UNA vez por peticion y se pasa a un diccionario: cada lang()
  ' posterior es una busqueda directa (sin copiar la matriz ni recorrerla).
  Dim varLangCache, strAppVar, di, dKey
  If Not IsObject(lhdLangDict) Then
    lhdLangDictId = ""
  End If
  If lhdLangDictId <> CStr(intLangSetting) Then
    strAppVar = "lhd_LangCache_" & intLangSetting
    varLangCache = Application(strAppVar)
    If IsEmpty(varLangCache) Then
      Dim rstLangStr
      Set rstLangStr = SQLQuery(cnnDB, "SELECT variable, langText FROM tblLangStrings WHERE id=" & intLangSetting & " ORDER BY variable ASC")
      If Not rstLangStr.EOF Then
        varLangCache = rstLangStr.GetRows
      Else
        rstLangStr.Close
        Call DisplayError(3, "The language strings are missing.  Please run setup.asp to install them. (" & intLangSetting & ")")
      End If
      rstLangStr.Close
      Application(strAppVar) = varLangCache
      Application(strAppVar & "_Count") = Cdbl(UBound(varLangCache, 2))
    End If
    Set lhdLangDict = Server.CreateObject("Scripting.Dictionary")
    lhdLangDict.CompareMode = 1   ' sin distinguir mayusculas
    For di = 0 To UBound(varLangCache, 2)
      dKey = Trim(varLangCache(0, di) & "")
      If Not lhdLangDict.Exists(dKey) Then lhdLangDict.Add dKey, varLangCache(1, di) & ""
    Next
    lhdLangDictId = CStr(intLangSetting)
  End If

  Lang = variable

  Dim strValue
  strValue = "_0"
  dKey = Trim(variable & "")
  If lhdLangDict.Exists(dKey) Then strValue = lhdLangDict(dKey)

  If strValue = "_0" Then
    ' Cadena que no esta en la cache (p. ej. se agrego a la base despues de cargarla):
    ' consulta directa a la base, primero el idioma del usuario, luego el idioma por omision
    ' del sitio (tblConfig.DefaultLanguage) y por ultimo el ingles (id 1).
    ' Usa ADO directo (no SQLQuery) para que un fallo aqui nunca muestre la pagina de error,
    ' y solo si la conexion sigue abierta (algunas paginas llaman a lang() despues de cerrarla).
    Dim lkOpen
    lkOpen = False
    On Error Resume Next
    lkOpen = (cnnDB.State = 1)
    Err.Clear
    ' Tope por peticion: una pagina con muchas cadenas inexistentes no debe saturar la base
    If lhdLangDbFallbacks >= 6 Then lkOpen = False
    If lkOpen Then
      lhdLangDbFallbacks = lhdLangDbFallbacks + 1
      Dim lkRs, lkIds, lkTry, lkDef, lkSafe
      lkSafe = Replace(Trim(variable), "'", "''")
      lkDef = SafeCInt(Cfg(cnnDB, "DefaultLanguage"))
      Err.Clear
      lkIds = Array(CLng(intLangSetting), lkDef, 1)
      For Each lkTry In lkIds
        If strValue = "_0" And lkTry > 0 Then
          Set lkRs = cnnDB.Execute("SELECT langText FROM tblLangStrings WHERE id=" & lkTry & _
            " AND variable='" & lkSafe & "'")
          If Err.Number = 0 Then
            If Not lkRs.EOF Then strValue = lkRs("langText") & ""
            lkRs.Close
          End If
          Err.Clear
        End If
      Next
    End If
    On Error Goto 0
  End If

  If strValue = "_0" Then
    Lang = "@" & variable & "@"
  Else
    If Len(strValue) < 1 Then
      Lang = "!" & variable & "!"
    Else
      Lang = FixTildes(strValue)
    End If
  End If
End Function


' ArrayFind:
' Finds a value in a two demensional, sorted array
' Returns "_0" if the key is not found
Function ArrayFind(varArray, strKey, intLBound, intUBound)
  Dim intMid
  ArrayFind = "_0"
  Do While intLBound <= intUBound
    intMid = (intUBound + intLBound) \ 2
    Select Case StrComp(UCase(strKey), UCase(varArray(0, intMid)), 1)
      Case 0  ' Key Found
        ArrayFind = varArray(1, intMid)
        Exit Function
      Case -1 ' Key less than mid
        intUBound = intMid - 1
      Case 1  ' Key greater than mid
        intLBound = intMid + 1
    End Select
  Loop
End Function


' ClearLangCache:
' Removes all language strings from the lang cache application variables
' The cache application variables acre called lhd_LangCache_<lang id>
Sub ClearLangCache(cnnDB)
On Error Goto 0
  Dim strAppVar, rstLangID
  Set rstLangID = SQLQuery(cnnDB, "SELECT id FROM tblLanguage")
  While Not rstLangID.EOF 
    strAppVar = "lhd_LangCache_" & rstLangID("id")
    Application(strAppVar) = Empty
    rstLangID.MoveNext
  Wend
  rstLangID.Close
  Set rstLangID = Nothing
End Sub


' CreateCon:
' Returns a ADO Connection object
Function CreateCon
	' Ensure UTF-8 output encoding across all pages
	On Error Resume Next
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	Session.CodePage = 65001
	On Error Goto 0

	' Make sure application is initialized
	Call SetAppVariables
	
	Dim strConn, cnnDB
	' Check for usage of SQL securing or integrated security and
	' use the correct connection string

	' Connection strings with DRIVER are ODBC,
	' those with PROVIDER are OLE DB connections
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
	
	' Override the DBType & settings and manually set the connection string
	' strConn = "Provider=sqloledb;Data Source=yourdbserver;Initial Catalog=helpdesk;User Id=sa;Password=sapassword"

	' Keep Errors from occuring
	If Not Application("Debug") Then
		On Error Resume Next
	End If

	' Create and open the database connection and save it as a
	' session variable
	If Application("Debug") Then
		Response.Write("<p><b>Creating Connection...</b></p>")
	End If
	If lhdPageStart = 0 Then lhdPageStart = Timer
	Set cnnDB = Server.CreateObject("ADODB.Connection")
	cnnDB.Open(strConn)
	' Si no abrio, mostrar el motivo real (sin credenciales) en lugar de fallar mas adelante
	If cnnDB.State <> 1 Then
		Dim openNum, openDesc
		openNum = Err.Number
		openDesc = Err.Description
		If Len(openDesc) = 0 Then openDesc = "(sin descripcion)"
		Response.Clear
		Response.Write("<p><center><table width=""420""><tr><td bgcolor=""red"" align=""center""><b>No se pudo abrir la conexi&oacute;n a la base de datos</b></td></tr>" & _
			"<tr><td bgcolor=""#EEEEEE"" align=""left""><b>Error: </b>" & openNum & "<br /><b>Detalle: </b>" & Server.HTMLEncode(openDesc) & _
			"<br /><b>Tipo de base (DBType): </b>" & Application("DBType") & "<br /><b>Servidor / base / archivo: </b>" & _
			Server.HTMLEncode(Application("SQLServer") & " / " & Application("SQLDBase") & " / " & Application("AccessPath")) & _
			"<hr />Verifique que el servidor de base de datos est&eacute; disponible y que los datos de settings.asp sean correctos.</td></tr></table></center>")
		Response.End
	End If

	If IsEmpty(Application("lhd_cache_v4")) Then
		On Error Resume Next
		Call ClearLangCache(cnnDB)
		Application("lhd_cache_v4") = 1
		On Error Goto 0
	End If
	Set CreateCon = cnnDB
End Function


' GetSid:
' Returns the user's sid or 0

' SafeCInt: Safely converts any value to integer without crashing on empty or invalid inputs
Function SafeCInt(val)
  On Error Resume Next
  If IsNull(val) Or IsEmpty(val) Then
    SafeCInt = 0
  ElseIf IsNumeric(val) Then
    SafeCInt = CInt(val)
  Else
    SafeCInt = 0
  End If
  On Error Goto 0
End Function

' SafeCLng: como SafeCInt pero con Long (CInt desborda por encima de 32767,
' por eso un limite de bytes mayor se leia como 0).
Function SafeCLng(val)
  On Error Resume Next
  If IsNull(val) Or IsEmpty(val) Then
    SafeCLng = 0
  ElseIf IsNumeric(val) Then
    SafeCLng = CLng(val)
  Else
    SafeCLng = 0
  End If
  If Err.Number <> 0 Then SafeCLng = 0
  On Error Goto 0
End Function

Function GetSid
	If IsNull(Session("lhd_sid")) Then
		GetSid = 0
	Else
		If Session("lhd_sid") > 0 Then
			GetSid = Session("lhd_sid")
		Else
			GetSid = 0
		End If
	End if
End Function


' DisplayError:
' Procedure for creating error pages.
Sub DisplayError(eType, component)

	' Create the web page
  If Not Application("Debug") Then
    Response.Clear
    Response.Write("<html><head><title>ERROR</title></head><body>")
  End If
  Response.Write "<p><center><table width=""200""><tr><td bgcolor=""red"" align=""center"">" & _
	  "<b>ERROR</b></tr></td><tr><td bgcolor=""#eeeeee"" align=""center"">"

'	Error Types:
'	1: Missing required field
'	2: SQL error
'	3: Generic Error, just display full component string

	Select Case eType
	Case 1
		Response.Write "<b>" & component & "</b> " & _
		  "&nbsp;" & lang(cnnDB, "isarequiredfield") & ".<p>" & _
		  "<i>" & lang(cnnDB, "PleasepresstheBACKbutton") & "</i></p>"
	Case 2
		Response.Write(lang(cnnDB, "ASQLqueryhasfailed") & ". ")
	Case 3
		Response.Write(component)
	End Select


	' Finish off the table and page
	Response.Write("</tr></td></table></center><p>&nbsp;</p></body></html>")

	' Stop processing the .asp file
	Response.End()
End Sub


' TrapError:
' Gracefully trap errors and print message
' ** Do not translate this routine
Sub TrapError (intNum, strDescription, strSource)

	Dim strInformation, strHexNum

	' Format the error number in Hex, 8 characters long
	strHexNum = Right(String(8, "0") & Hex(intNum), 8)

	Response.Write("<p><center><table width=""300"">")
	Response.Write("<tr><td bgcolor=""red"" align=""center"">")
	Response.Write("<B>Application Error</b></td></tr>")
	Response.Write("<tr><td bgcolor=""#EEEEEE"" align=""left"">")

	Response.Write("<b>Number: </b>" & intNum & " (0x" & strHexNum & ")<br />")
	Response.Write("<b>Source: </b>" & strSource & "<br />")
	Response.Write("<b>Description: </b>" & strDescription & "<hr />")

	' Print extra information
	Select Case Err.Number
		Case 3709	' no db connection
			strInformation = "The database connection could not be opened.  Please check your" & _
				" configuration and make sure the database is accessible."
		Case -2147217865 ' sql: bad table reference
			strInformation = "There is an error in the SQL query string probably referencing a" & _
				" table.  Read the description above and check the code."
		Case -2147217900 ' sql: bad syntax
			strInformation = "The SQL query string is using bad syntax.  Read the description above" & _
				" and check the code."
		Case -2147217904 ' sql: invalid column
			strInformation = "The SQL query references a invalid column or object.  Read the description above" & _
				" and check the code."
		Case Else
			strInformation = "No more information is available."
	End Select

	strInformation = strInformation & vbNewLine & _
		"<p>Contact your administrator or visit the Liberum Help Desk " & _
		"<a href=""http://www.liberum.org"">website</a>."
	Response.Write(strInformation)

	Response.End
End Sub



' LhdThemeLoad:
' Carga la paleta guardada en la base de datos (tabla tblTheme) a memoria de
' aplicacion, una sola vez por arranque. Si la tabla no existe la crea.
Sub LhdThemeLoad(cnnDB)
  If Application("lhd_theme_loaded") = "1" Then Exit Sub
  ' Si la tabla no existe (o fallo la lectura), no reintentar en cada peticion: esperar 1 minuto.
  If Len(Application("lhd_theme_next") & "") > 0 Then
    If IsDate(Application("lhd_theme_next")) Then
      If Now < CDate(Application("lhd_theme_next")) Then Exit Sub
    End If
  End If
  On Error Resume Next
  Dim rs, arr, ti, loadedOk
  loadedOk = False
  Set rs = cnnDB.Execute("SELECT theme_key, theme_value FROM tblTheme")
  If Err.Number = 0 Then
    If Not rs Is Nothing Then
      If rs.State = 1 Then
        If Not rs.EOF Then
          ' GetRows lee todo de una vez: sin bucle sobre el cursor
          arr = rs.GetRows()
          If IsArray(arr) Then
            For ti = 0 To UBound(arr, 2)
              If Len(arr(1, ti) & "") > 0 Then Application("lhd_theme_" & Trim(arr(0, ti) & "")) = arr(1, ti) & ""
            Next
          End If
        End If
        rs.Close
        loadedOk = True
      End If
    End If
  End If
  Err.Clear
  If loadedOk Then
    Application("lhd_theme_loaded") = "1"
  Else
    ' La tabla tblTheme no existe (se crea con db/tblTheme_*.sql): colores por defecto
    Application("lhd_theme_next") = DateAdd("n", 1, Now)
  End If
  On Error Goto 0
End Sub

' GetThemeColor:
' Recupera un color de tema personalizado (memoria de aplicacion, respaldada en BD).
Function GetThemeColor(cnnDB, strKey, strDefault)
  On Error Resume Next
  Dim val
  Call LhdThemeLoad(cnnDB)
  val = Application("lhd_theme_" & strKey)
  If Len(val) > 0 Then
    GetThemeColor = val
  Else
    GetThemeColor = strDefault
  End If
  On Error Goto 0
End Function

' SetThemeColor:
' Guarda un color de tema en memoria de aplicacion y en la base de datos, para que
' se conserve al reiniciar la aplicacion.
Sub SetThemeColor(cnnDB, strKey, strValue)
  On Error Resume Next
  Dim k, v
  Call LhdThemeLoad(cnnDB)
  ' solo colores #RRGGBB
  strValue = Left(Trim(strValue), 20)
  If Not LhdIsHexColor(strValue) Then
    Application("lhd_theme_err") = "Color no valido para '" & strKey & "': " & strValue
    Exit Sub
  End If
  k = Replace(Left(strKey, 30), "'", "")
  v = Replace(strValue, "'", "")
  Application("lhd_theme_" & strKey) = strValue
  Err.Clear
  cnnDB.Execute "DELETE FROM tblTheme WHERE theme_key='" & k & "'"
  If Err.Number <> 0 Then
    Application("lhd_theme_err") = "DELETE: " & Err.Description
    Err.Clear
  End If
  cnnDB.Execute "INSERT INTO tblTheme (theme_key, theme_value) VALUES ('" & k & "', '" & v & "')"
  If Err.Number <> 0 Then
    Application("lhd_theme_err") = "INSERT: " & Err.Description
    Err.Clear
  End If
  On Error Goto 0
End Sub

' LhdIsHexColor: True si el texto es un color #RRGGBB valido.
Function LhdIsHexColor(strValue)
  Dim ci, ch
  LhdIsHexColor = False
  If Len(strValue) <> 7 Then Exit Function
  If Left(strValue, 1) <> "#" Then Exit Function
  For ci = 2 To 7
    ch = LCase(Mid(strValue, ci, 1))
    If InStr("0123456789abcdef", ch) = 0 Then Exit Function
  Next
  LhdIsHexColor = True
End Function

' GetUserInitials:
' Devuelve las iniciales del usuario para el avatar de forma segura
Function GetUserInitials(strName)
  On Error Resume Next
  Dim init, s, parts
  init = "U"
  If IsNull(strName) Or IsEmpty(strName) Then
    GetUserInitials = init
    Exit Function
  End If
  s = CStr(Trim(strName))
  If Len(s) > 0 Then
    parts = Split(s, " ")
    If UBound(parts) >= 1 Then
      init = UCase(Left(parts(0), 1) & Left(parts(1), 1))
    Else
      init = UCase(Left(s, 2))
    End If
  End If
  If Len(init) = 0 Then init = "U"
  GetUserInitials = init
  On Error Goto 0
End Function

' DisplayHeader:
' Procedure to write HTML for the header (modernized with Left Sidebar Navigation & Sticky Footer).
Sub DisplayHeader(cnnDB, sid)
  On Error Resume Next
  Response.CodePage = 65001
  Response.CharSet = "utf-8"
  Session.CodePage = 65001
  On Error Goto 0

  Session("lhd_HeaderRendered") = True

  Dim isRep, uname, fname, baseURL, isAdmin, curScript, userInitials
  isRep   = SafeCInt(Usr(cnnDB, sid, "IsRep"))
  uname   = Usr(cnnDB, sid, "uid")
  If IsNull(uname) Then uname = ""
  fname   = Usr(cnnDB, sid, "fname")
  If IsNull(fname) Or Len(Trim(fname)) = 0 Then fname = uname
  If Len(Trim(fname)) = 0 Then fname = "Usuario"
  isAdmin = False
  If Session("lhd_IsAdmin") = True Then isAdmin = True
  userInitials = GetUserInitials(fname)

  ' Vista activa: un tecnico puede alternar entre el menu de Representante y el de Usuario
  ' (switchview.asp guarda la eleccion en Session("lhd_ViewMode")).
  Dim viewMode, showRep
  viewMode = LCase(Session("lhd_ViewMode") & "")
  Dim closedView, actClosedLink
  closedView = (LCase(Request.QueryString("st")) = "closed")
  actClosedLink = ""
  showRep = (isRep = 1 And viewMode <> "user")

  baseURL = GetAppPath()
  If Len(baseURL) = 0 Then baseURL = Cfg(cnnDB, "BaseURL")
  If Right(baseURL, 1) = "/" Then baseURL = Left(baseURL, Len(baseURL) - 1)

  curScript = LCase(Request.ServerVariables("SCRIPT_NAME"))

  ' Theme colors from configuration
  Dim clrPrimary, clrPrimaryHover, clrSidebar, clrAccent, clrBg
  clrPrimary      = GetThemeColor(cnnDB, "primary", "#0057B8")
  clrPrimaryHover = GetThemeColor(cnnDB, "primary_hover", "#003F8A")
  clrSidebar      = GetThemeColor(cnnDB, "sidebar", "#0F172A")
  clrAccent       = GetThemeColor(cnnDB, "accent", "#0284C7")
  clrBg           = GetThemeColor(cnnDB, "bg", "#F1F5F9")

  ' Show open ticket count for rep
  Dim hdrNumProblems
  hdrNumProblems = 0
  If isRep = 1 And SafeCInt(sid) > 0 Then
    Dim hdrNumStr, hdrNumRes
    hdrNumStr = "SELECT count(*) AS total FROM problems WHERE rep=" & SafeCInt(sid) & " AND status < 100"
    Set hdrNumRes = SQLQuery(cnnDB, hdrNumStr)
    If Not hdrNumRes.EOF Then
      hdrNumProblems = SafeCInt(hdrNumRes("total"))
    End If
    hdrNumRes.Close
  End If

  Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" & vbNewLine
  Response.Write "<meta charset=""utf-8"" />" & vbNewLine
  Response.Write "<meta name=""viewport"" content=""width=device-width, initial-scale=1"" />" & vbNewLine
  Response.Write "<style type=""text/css"">" & vbNewLine
  Response.Write ":root {" & vbNewLine
  Response.Write "  --lhd-primary: " & clrPrimary & " !important;" & vbNewLine
  Response.Write "  --lhd-primary-hover: " & clrPrimaryHover & " !important;" & vbNewLine
  Response.Write "  --lhd-sidebar: " & clrSidebar & " !important;" & vbNewLine
  Response.Write "  --lhd-accent: " & clrAccent & " !important;" & vbNewLine
  Response.Write "  --lhd-bg: " & clrBg & " !important;" & vbNewLine
  Response.Write "  --color-primary: " & clrPrimary & " !important;" & vbNewLine
  Response.Write "  --color-primary-dark: " & clrPrimaryHover & " !important;" & vbNewLine
  Response.Write "  --color-accent: " & clrAccent & " !important;" & vbNewLine
  Response.Write "  --color-bg: " & clrBg & " !important;" & vbNewLine
  Response.Write "}" & vbNewLine
  Response.Write "html, body { min-height: 100vh !important; margin: 0 !important; padding: 0 !important; }" & vbNewLine
  Response.Write "body { font-family: -apple-system, BlinkMacSystemFont, ""Segoe UI"", Roboto, Arial, sans-serif !important; font-size: 14px; line-height: 1.5; color: #1e293b; background-color: var(--lhd-bg) !important; display: flex !important; flex-direction: row !important; overflow-x: hidden !important; }" & vbNewLine
  Response.Write "/* Left Sidebar Navigation */" & vbNewLine
  Response.Write ".lhd-sidebar { width: 260px !important; min-width: 260px !important; height: 100vh !important; position: fixed !important; top: 0 !important; left: 0 !important; bottom: 0 !important; background: var(--lhd-sidebar) !important; color: #f8fafc !important; display: flex !important; flex-direction: column !important; z-index: 1000 !important; overflow-y: auto !important; box-shadow: 2px 0 16px rgba(0,0,0,0.18) !important; transition: transform 0.25s ease !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write ".lhd-sidebar-brand { padding: 18px 20px 14px !important; display: flex !important; align-items: center !important; gap: 12px !important; text-decoration: none !important; color: #ffffff !important; border-bottom: 1px solid rgba(255,255,255,0.08) !important; background: rgba(0,0,0,0.15) !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write ".lhd-sidebar-brand .brand-logo { width: 36px; height: 36px; border-radius: 8px; background: linear-gradient(135deg, var(--lhd-primary), var(--lhd-accent)); display: flex; align-items: center; justify-content: center; font-size: 18px; color: #fff; box-shadow: 0 4px 10px rgba(0,0,0,0.3); flex-shrink: 0; }" & vbNewLine
  Response.Write ".lhd-sidebar-brand .brand-title { font-size: 14px; font-weight: 700; color: #ffffff; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; line-height: 1.2; }" & vbNewLine
  Response.Write ".lhd-sidebar-brand .brand-sub { font-size: 11px; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px; }" & vbNewLine
  Response.Write ".lhd-sidebar-user { padding: 12px 16px !important; margin: 12px 14px 6px !important; border-radius: 8px !important; background: rgba(255,255,255,0.06) !important; border: 1px solid rgba(255,255,255,0.08) !important; display: flex !important; align-items: center !important; gap: 10px !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write ".lhd-user-avatar { width: 36px; height: 36px; border-radius: 50%; background: linear-gradient(135deg, var(--lhd-accent), var(--lhd-primary)); color: #ffffff; font-weight: 700; font-size: 13px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }" & vbNewLine
  Response.Write ".lhd-user-meta { display: flex; flex-direction: column; min-width: 0; overflow: hidden; }" & vbNewLine
  Response.Write ".lhd-user-name { font-size: 13px; font-weight: 600; color: #f8fafc; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; line-height: 1.2; }" & vbNewLine
  Response.Write ".lhd-user-role { font-size: 11px; color: #94a3b8; display: flex; align-items: center; gap: 4px; margin-top: 2px; }" & vbNewLine
  Response.Write ".lhd-status-dot { width: 7px; height: 7px; border-radius: 50%; background: #10b981; display: inline-block; box-shadow: 0 0 6px #10b981; }" & vbNewLine
  Response.Write ".lhd-sidebar-nav { padding: 8px 12px !important; flex: 1 !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write ".lhd-nav-section-title { font-size: 10px !important; font-weight: 700 !important; text-transform: uppercase !important; color: #64748b !important; letter-spacing: 0.8px !important; padding: 14px 10px 4px !important; }" & vbNewLine
  Response.Write ".lhd-nav-link { display: flex !important; align-items: center !important; gap: 10px !important; padding: 8px 12px !important; margin-bottom: 2px !important; border-radius: 6px !important; color: #cbd5e1 !important; text-decoration: none !important; font-size: 13px !important; font-weight: 500 !important; transition: all 0.15s ease !important; background: transparent !important; }" & vbNewLine
  Response.Write ".lhd-nav-link:hover { background: rgba(255,255,255,0.09) !important; color: #ffffff !important; transform: translateX(3px) !important; }" & vbNewLine
  Response.Write ".lhd-nav-link.active { background: var(--lhd-primary) !important; color: #ffffff !important; font-weight: 600 !important; box-shadow: 0 4px 10px rgba(0,0,0,0.25) !important; }" & vbNewLine
  Response.Write ".lhd-nav-link .nav-icon { font-size: 15px; width: 18px; text-align: center; flex-shrink: 0; }" & vbNewLine
  Response.Write ".lhd-nav-badge { margin-left: auto; background: var(--lhd-primary); color: #ffffff; font-size: 11px; font-weight: 700; padding: 1px 7px; border-radius: 9999px; line-height: 1.3; }" & vbNewLine
  Response.Write ".lhd-nav-cta { background: linear-gradient(135deg, var(--lhd-primary), var(--lhd-accent)) !important; color: #ffffff !important; font-weight: 600 !important; margin: 4px 0 10px !important; box-shadow: 0 4px 10px rgba(0,0,0,0.2) !important; }" & vbNewLine
  Response.Write ".lhd-nav-cta:hover { filter: brightness(1.1) !important; transform: translateY(-1px) !important; }" & vbNewLine
  Response.Write "/* Main Content & Sticky Footer */" & vbNewLine
  Response.Write ".lhd-main-wrapper { margin-left: 260px !important; flex: 1 !important; display: flex !important; flex-direction: column !important; min-height: 100vh !important; min-width: 0 !important; background-color: var(--lhd-bg) !important; transition: margin-left 0.25s ease !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write ".lhd-topbar { height: 56px !important; background: #ffffff !important; border-bottom: 1px solid #e2e8f0 !important; display: flex !important; align-items: center !important; justify-content: space-between !important; padding: 0 28px !important; position: sticky !important; top: 0 !important; z-index: 900 !important; box-shadow: 0 1px 3px rgba(0,0,0,0.03) !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write ".lhd-topbar-left { display: flex !important; align-items: center !important; gap: 12px !important; }" & vbNewLine
  Response.Write ".lhd-sidebar-toggle { display: none !important; background: transparent !important; border: 1px solid #cbd5e0 !important; border-radius: 6px !important; padding: 5px 9px !important; font-size: 16px !important; cursor: pointer !important; color: #334155 !important; }" & vbNewLine
  Response.Write ".lhd-breadcrumb { font-size: 13px !important; font-weight: 600 !important; color: #334155 !important; display: flex !important; align-items: center !important; gap: 6px !important; }" & vbNewLine
  Response.Write ".lhd-topbar-right { display: flex !important; align-items: center !important; gap: 10px !important; }" & vbNewLine
  Response.Write ".lhd-content { flex: 1 0 auto !important; padding: 24px 28px !important; width: 100% !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write ".lhd-footer { flex-shrink: 0 !important; margin-top: auto !important; background: #ffffff !important; border-top: 1px solid #e2e8f0 !important; padding: 16px 28px !important; font-size: 13px !important; color: #64748b !important; text-align: center !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write "/* Full Width Adaptation */" & vbNewLine
  Response.Write ".lhd-wrapper { width: 100% !important; max-width: 100% !important; margin: 0 !important; padding: 0 !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write "table.Normal, table.Wide, table.Narrow { width: 100% !important; max-width: 100% !important; margin: 0 0 20px 0 !important; background: #fff !important; border-radius: 8px !important; box-shadow: 0 2px 8px rgba(0,0,0,0.06) !important; border-collapse: collapse !important; overflow: hidden !important; }" & vbNewLine
  Response.Write "/* Cards & Widgets */" & vbNewLine
  Response.Write ".lhd-card { background: #ffffff !important; border-radius: 10px !important; box-shadow: 0 2px 8px rgba(0,0,0,0.06) !important; padding: 24px !important; margin-bottom: 20px !important; box-sizing: border-box !important; border: 1px solid #e2e8f0 !important; }" & vbNewLine
  Response.Write ".lhd-kpi-grid { display: grid !important; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)) !important; gap: 16px !important; margin-bottom: 24px !important; }" & vbNewLine
  Response.Write ".lhd-kpi-card { background: #ffffff !important; border-radius: 10px !important; box-shadow: 0 2px 8px rgba(0,0,0,0.06) !important; padding: 20px !important; border: 1px solid #e2e8f0 !important; border-left: 4px solid var(--lhd-primary) !important; transition: transform 0.18s ease, box-shadow 0.18s ease !important; }" & vbNewLine
  Response.Write ".lhd-kpi-card:hover { transform: translateY(-2px) !important; box-shadow: 0 6px 16px rgba(0,0,0,0.1) !important; }" & vbNewLine
  Response.Write ".lhd-kpi-card.kpi-open { border-left-color: var(--lhd-primary) !important; }" & vbNewLine
  Response.Write ".lhd-kpi-card.kpi-urgent { border-left-color: #ef4444 !important; }" & vbNewLine
  Response.Write ".lhd-kpi-card.kpi-pending { border-left-color: #f59e0b !important; }" & vbNewLine
  Response.Write ".lhd-kpi-card.kpi-closed { border-left-color: #10b981 !important; }" & vbNewLine
  Response.Write ".lhd-kpi-card .kpi-value { font-size: 32px !important; font-weight: 700 !important; line-height: 1 !important; color: #1e293b !important; }" & vbNewLine
  Response.Write ".lhd-kpi-card .kpi-label { font-size: 12px !important; color: #64748b !important; font-weight: 600 !important; text-transform: uppercase !important; margin-top: 6px !important; letter-spacing: 0.5px !important; }" & vbNewLine
  Response.Write ".lhd-action-grid { display: grid !important; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)) !important; gap: 16px !important; margin-bottom: 24px !important; }" & vbNewLine
  Response.Write ".lhd-action-card { background: #ffffff !important; border: 1px solid #e2e8f0 !important; border-radius: 10px !important; padding: 20px 16px !important; text-align: center !important; text-decoration: none !important; color: #1e293b !important; display: flex !important; flex-direction: column !important; align-items: center !important; gap: 8px !important; transition: all 0.2s ease !important; box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important; }" & vbNewLine
  Response.Write ".lhd-action-card:hover { border-color: var(--lhd-primary) !important; box-shadow: 0 8px 20px rgba(0,87,184,0.14) !important; transform: translateY(-3px) !important; color: var(--lhd-primary) !important; }" & vbNewLine
  Response.Write ".lhd-action-card .action-icon { font-size: 28px !important; line-height: 1 !important; }" & vbNewLine
  Response.Write ".lhd-action-card .action-label { font-size: 13px !important; font-weight: 600 !important; }" & vbNewLine
  Response.Write ".lhd-action-card .action-desc { font-size: 11px !important; color: #64748b !important; }" & vbNewLine
  Response.Write "/* Tables */" & vbNewLine
  Response.Write ".lhd-table { width: 100% !important; border-collapse: collapse !important; font-size: 13px !important; background: #ffffff !important; }" & vbNewLine
  Response.Write ".lhd-table th { background: var(--lhd-primary) !important; color: #ffffff !important; padding: 12px 14px !important; text-align: left !important; font-weight: 600 !important; font-size: 12px !important; letter-spacing: 0.3px !important; }" & vbNewLine
  Response.Write ".lhd-table th a { color: #ffffff !important; text-decoration: none !important; }" & vbNewLine
  Response.Write ".lhd-table td { padding: 11px 14px !important; border-bottom: 1px solid #e2e8f0 !important; vertical-align: middle !important; }" & vbNewLine
  Response.Write ".lhd-table tr:hover td { background: #f8fafc !important; }" & vbNewLine
  Response.Write ".section-header { background: var(--lhd-primary) !important; color: #ffffff !important; padding: 12px 18px !important; font-weight: 600 !important; font-size: 14px !important; border-radius: 8px 8px 0 0 !important; margin: -24px -24px 18px !important; display: flex !important; align-items: center !important; gap: 8px !important; }" & vbNewLine
  Response.Write "/* Buttons & Badges */" & vbNewLine
  Response.Write ".btn { display: inline-flex !important; align-items: center !important; gap: 6px !important; padding: 8px 16px !important; border-radius: 6px !important; font-size: 13px !important; font-weight: 600 !important; cursor: pointer !important; border: 1px solid transparent !important; text-decoration: none !important; line-height: 1.4 !important; transition: all 0.15s ease !important; font-family: inherit !important; }" & vbNewLine
  Response.Write ".btn-primary { background: var(--lhd-primary) !important; color: #ffffff !important; border-color: var(--lhd-primary) !important; }" & vbNewLine
  Response.Write ".btn-primary:hover { background: var(--lhd-primary-hover) !important; border-color: var(--lhd-primary-hover) !important; color: #ffffff !important; text-decoration: none !important; }" & vbNewLine
  Response.Write ".btn-secondary { background: #ffffff !important; color: #475569 !important; border-color: #cbd5e1 !important; }" & vbNewLine
  Response.Write ".btn-secondary:hover { background: #f8fafc !important; border-color: #94a3b8 !important; color: #1e293b !important; text-decoration: none !important; }" & vbNewLine
  Response.Write ".btn-sm { padding: 5px 12px !important; font-size: 12px !important; }" & vbNewLine
  Response.Write ".badge { display: inline-flex !important; align-items: center !important; gap: 4px !important; padding: 3px 10px !important; border-radius: 9999px !important; font-size: 11px !important; font-weight: 600 !important; letter-spacing: 0.3px !important; white-space: nowrap !important; }" & vbNewLine
  Response.Write ".badge-open { background: #dbeafe !important; color: #1e40af !important; }" & vbNewLine
  Response.Write ".badge-pending { background: #fef3c7 !important; color: #92400e !important; }" & vbNewLine
  Response.Write ".badge-closed { background: #d1fae5 !important; color: #065f46 !important; }" & vbNewLine
  Response.Write ".badge-high { background: #fee2e2 !important; color: #991b1b !important; }" & vbNewLine
  Response.Write ".badge-medium { background: #fef3c7 !important; color: #92400e !important; }" & vbNewLine
  Response.Write ".badge-low { background: #d1fae5 !important; color: #065f46 !important; }" & vbNewLine
  Response.Write ".lhd-dot { display: inline-block !important; width: 8px !important; height: 8px !important; border-radius: 50% !important; background: currentColor !important; flex-shrink: 0 !important; }" & vbNewLine
  Response.Write ".lhd-card-flush { padding: 0 !important; overflow: hidden !important; }" & vbNewLine
  Response.Write ".lhd-table thead th { position: static !important; top: auto !important; }" & vbNewLine
  Response.Write ".lhd-card-header { background: var(--lhd-primary) !important; color: #ffffff !important; padding: 14px 20px !important; font-weight: 600 !important; font-size: 14px !important; display: flex !important; align-items: center !important; gap: 8px !important; }" & vbNewLine
  Response.Write ".alert { display: flex !important; align-items: center !important; gap: 10px !important; padding: 12px 16px !important; border-radius: 8px !important; margin-bottom: 16px !important; font-size: 13px !important; }" & vbNewLine
  Response.Write ".alert-success { background: #ecfdf5 !important; border: 1px solid #a7f3d0 !important; color: #065f46 !important; }" & vbNewLine
  Response.Write ".alert-error { background: #fef2f2 !important; border: 1px solid #fecaca !important; color: #991b1b !important; }" & vbNewLine
  Response.Write ".form-control { width: 100% !important; padding: 8px 12px !important; border: 1px solid #cbd5e0 !important; border-radius: 6px !important; font-size: 13px !important; font-family: inherit !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write ".form-control:focus { outline: none !important; border-color: var(--lhd-primary) !important; box-shadow: 0 0 0 3px rgba(0,87,184,0.15) !important; }" & vbNewLine
  Response.Write "input[type=""text""], input[type=""password""], select, textarea { font-family: inherit !important; font-size: 13px !important; border: 1px solid #cbd5e0 !important; border-radius: 6px !important; padding: 7px 10px !important; box-sizing: border-box !important; }" & vbNewLine
  Response.Write "input[type=""submit""], input[type=""button""] { background: var(--lhd-primary) !important; color: #fff !important; border: 1px solid var(--lhd-primary) !important; padding: 7px 16px !important; border-radius: 6px !important; font-weight: 600 !important; cursor: pointer !important; font-size: 13px !important; font-family: inherit !important; }" & vbNewLine
  Response.Write "input[type=""submit""]:hover, input[type=""button""]:hover { background: var(--lhd-primary-hover) !important; border-color: var(--lhd-primary-hover) !important; }" & vbNewLine
  Response.Write "tr.Head1 td, tr.Head1 th { background: linear-gradient(135deg, var(--lhd-primary), var(--lhd-accent)) !important; color: #ffffff !important; font-size: 15px !important; font-weight: 700 !important; padding: 13px 16px !important; text-align: center !important; border: none !important; }" & vbNewLine
  Response.Write "tr.Head2 td, tr.Head2 th { background: var(--lhd-primary-hover) !important; color: #ffffff !important; font-weight: 600 !important; padding: 10px 14px !important; border: none !important; font-size: 13px !important; }" & vbNewLine
  Response.Write "tr.Head2 a, a.HeadLink { color: #ffffff !important; text-decoration: none !important; }" & vbNewLine
  Response.Write "tr.Head2 a:hover, a.HeadLink:hover { text-decoration: underline !important; }" & vbNewLine
  Response.Write "tr.Body1 td, tr.Body2 td, tr.Body3 td { padding: 10px 14px !important; border-bottom: 1px solid #edf2f7 !important; font-size: 13px !important; color: #334155 !important; }" & vbNewLine
  Response.Write "tr.Body2 td { background: #f8fafc !important; }" & vbNewLine
  Response.Write "a { color: var(--lhd-primary); text-decoration: none; }" & vbNewLine
  Response.Write "a:hover { color: var(--lhd-primary-hover); text-decoration: underline; }" & vbNewLine
  Response.Write "@media (max-width: 991px) {" & vbNewLine
  Response.Write "  .lhd-sidebar { transform: translateX(-100%) !important; }" & vbNewLine
  Response.Write "  .lhd-sidebar.open { transform: translateX(0) !important; }" & vbNewLine
  Response.Write "  .lhd-main-wrapper { margin-left: 0 !important; }" & vbNewLine
  Response.Write "  .lhd-sidebar-toggle { display: inline-flex !important; }" & vbNewLine
  Response.Write "  .lhd-content { padding: 16px !important; }" & vbNewLine
  Response.Write "}" & vbNewLine
  Response.Write ".lhd-sidebar-overlay { display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 999; }" & vbNewLine
  Response.Write ".lhd-sidebar-overlay.open { display: block; }" & vbNewLine
  Response.Write ".lhd-viewswitch { margin: 6px 14px 4px; padding: 9px 12px; border-radius: 8px; background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.08); }" & vbNewLine
  Response.Write ".lhd-viewswitch .vs-title { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .8px; color: #64748b; margin-bottom: 7px; }" & vbNewLine
  Response.Write ".lhd-viewswitch a.vs { display: flex; align-items: center; justify-content: space-between; gap: 8px; text-decoration: none !important; color: #94a3b8 !important; font-size: 12px; font-weight: 600; }" & vbNewLine
  Response.Write ".lhd-viewswitch .vs-lbl.on { color: #ffffff; }" & vbNewLine
  Response.Write ".lhd-viewswitch .vs-track { position: relative; width: 38px; height: 20px; border-radius: 9999px; background: #475569; flex-shrink: 0; transition: background .2s ease; }" & vbNewLine
  Response.Write ".lhd-viewswitch .vs-track.rep { background: var(--lhd-primary); }" & vbNewLine
  Response.Write ".lhd-viewswitch .vs-knob { position: absolute; top: 2px; left: 2px; width: 16px; height: 16px; border-radius: 50%; background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,.35); transition: transform .2s ease; }" & vbNewLine
  Response.Write ".lhd-viewswitch .vs-track.rep .vs-knob { transform: translateX(18px); }" & vbNewLine
  Response.Write ".lhd-viewswitch a.vs:hover .vs-track { filter: brightness(1.15); }" & vbNewLine
  Response.Write ".lhd-viewswitch a.vs:focus-visible { outline: 2px solid #ffffff; outline-offset: 3px; border-radius: 6px; }" & vbNewLine
  Response.Write "</style>" & vbNewLine

  ' Render Left Sidebar
  Response.Write "<aside class=""lhd-sidebar"" id=""lhdSidebar"">"
  Response.Write "  <a class=""lhd-sidebar-brand"" href=""" & baseURL & "/"">"
  Response.Write "    <div class=""brand-logo"">&#128203;</div>"
  Response.Write "    <div class=""brand-info"">"
  Response.Write "      <div class=""brand-title"">" & FixTildes(Server.HTMLEncode(Cfg(cnnDB, "SiteName"))) & "</div>"
  Response.Write "      <div class=""brand-sub"">Mesa de Ayuda</div>"
  Response.Write "    </div>"
  Response.Write "  </a>"

  Response.Write "  <div class=""lhd-sidebar-user"">"
  Response.Write "    <div class=""lhd-user-avatar"">" & userInitials & "</div>"
  Response.Write "    <div class=""lhd-user-meta"">"
  Response.Write "      <div class=""lhd-user-name"" title=""" & FixTildes(Server.HTMLEncode(fname)) & """>" & FixTildes(Server.HTMLEncode(fname)) & "</div>"
  Response.Write "      <div class=""lhd-user-role"">"
  Response.Write "        <span class=""lhd-status-dot""></span>"
  If isRep = 1 Then
    Response.Write "        <span>T&eacute;cnico</span>"
  ElseIf isAdmin Then
    Response.Write "        <span>Administrador</span>"
  Else
    Response.Write "        <span>Usuario</span>"
  End If
  Response.Write "      </div>"
  Response.Write "    </div>"
  Response.Write "  </div>"

  If isRep = 1 Then
    Dim vsMode, vsTitle
    If showRep Then
      vsMode = "user"
      vsTitle = "Cambiar a la vista de Usuario"
    Else
      vsMode = "rep"
      vsTitle = "Cambiar a la vista de Representante"
    End If
    Response.Write "  <div class=""lhd-viewswitch"">"
    Response.Write "    <div class=""vs-title"">Vista del men&uacute;</div>"
    Response.Write "    <a class=""vs"" role=""switch"" aria-checked=""" & LCase(CStr(showRep)) & """ title=""" & vsTitle & """ href=""" & baseURL & "/switchview.asp?mode=" & vsMode & """>"
    Response.Write "      <span class=""vs-lbl" & IIf2(Not showRep, " on", "") & """>Usuario</span>"
    Response.Write "      <span class=""vs-track" & IIf2(showRep, " rep", "") & """><span class=""vs-knob""></span></span>"
    Response.Write "      <span class=""vs-lbl" & IIf2(showRep, " on", "") & """>Representante</span>"
    Response.Write "    </a>"
    Response.Write "  </div>"
  End If

  Response.Write "  <nav class=""lhd-sidebar-nav"">"

  If showRep Then
    ' Navegacion para Tecnico
    Response.Write "    <div class=""lhd-nav-section-title"">Principal</div>"
    Dim actRepHome, actRepNew, actRepTickets, actRepSearch, actKb, actInout, actAdmin
    actRepHome = "" : actRepNew = "" : actRepTickets = "" : actRepSearch = "" : actKb = "" : actInout = "" : actAdmin = ""

    If InStr(curScript, "/rep/default.asp") > 0 Or (InStr(curScript, "/rep/") > 0 And Right(curScript, 5) = "/rep/") Then actRepHome = " active"
    If InStr(curScript, "/rep/new.asp") > 0 Then actRepNew = " active"
    If InStr(curScript, "/rep/view.asp") > 0 Or InStr(curScript, "/rep/details.asp") > 0 Then actRepTickets = " active"
    If InStr(curScript, "/rep/search.asp") > 0 Or InStr(curScript, "/rep/results.asp") > 0 Then actRepSearch = " active"
    If InStr(curScript, "/kb/") > 0 Then actKb = " active"
    If InStr(curScript, "/inout/") > 0 Then actInout = " active"
    If InStr(curScript, "/admin/") > 0 Then actAdmin = " active"

    Response.Write "    <a class=""lhd-nav-link" & actRepHome & """ href=""" & baseURL & "/rep/"">"
    Response.Write "      <span class=""nav-icon"">&#128202;</span><span>Panel de Soporte</span>"
    Response.Write "    </a>"

    Response.Write "    <a class=""lhd-nav-link lhd-nav-cta" & actRepNew & """ href=""" & baseURL & "/rep/new.asp"">"
    Response.Write "      <span class=""nav-icon"">&#10133;</span><span>Nuevo Ticket</span>"
    Response.Write "    </a>"

    Response.Write "    <div class=""lhd-nav-section-title"">Gesti&oacute;n de Tickets</div>"

    If closedView And actRepTickets <> "" Then
      actRepTickets = ""
      actClosedLink = " active"
    End If
    Response.Write "    <a class=""lhd-nav-link" & actRepTickets & """ href=""" & baseURL & "/rep/view.asp"">"
    Response.Write "      <span class=""nav-icon"">&#128203;</span><span>Mis Tickets</span>"
    If hdrNumProblems > 0 Then
      Response.Write "      <span class=""lhd-nav-badge"">" & hdrNumProblems & "</span>"
    End If
    Response.Write "    </a>"

    Response.Write "    <a class=""lhd-nav-link" & actClosedLink & """ href=""" & baseURL & "/rep/view.asp?st=closed"">"
    Response.Write "      <span class=""nav-icon"">&#9989;</span><span>Tickets Cerrados</span>"
    Response.Write "    </a>"

    Response.Write "    <a class=""lhd-nav-link" & actRepSearch & """ href=""" & baseURL & "/rep/search.asp"">"
    Response.Write "      <span class=""nav-icon"">&#128269;</span><span>Buscar Incidencias</span>"
    Response.Write "    </a>"

    If SafeCInt(Cfg(cnnDB, "EnableKB")) >= 1 Or SafeCInt(Cfg(cnnDB, "UseInoutBoard")) = 1 Then
      Response.Write "    <div class=""lhd-nav-section-title"">Recursos</div>"
      If SafeCInt(Cfg(cnnDB, "EnableKB")) >= 1 Then
        Response.Write "    <a class=""lhd-nav-link" & actKb & """ href=""" & baseURL & "/kb/"">"
        Response.Write "      <span class=""nav-icon"">&#128218;</span><span>Base Conocimiento</span>"
        Response.Write "    </a>"
      End If
      If SafeCInt(Cfg(cnnDB, "UseInoutBoard")) = 1 Then
        Response.Write "    <a class=""lhd-nav-link" & actInout & """ href=""" & baseURL & "/inout/default.asp"">"
        Response.Write "      <span class=""nav-icon"">&#128100;</span><span>Tablero In/Out</span>"
        Response.Write "    </a>"
      End If
    End If

    If isAdmin Then
      Response.Write "    <div class=""lhd-nav-section-title"">Administraci&oacute;n</div>"
      Response.Write "    <a class=""lhd-nav-link" & actAdmin & """ href=""" & baseURL & "/admin/"">"
      Response.Write "      <span class=""nav-icon"">&#9881;</span><span>Panel Admin</span>"
      Response.Write "    </a>"
    End If

  Else
    ' Navegacion para Usuario Normal
    Response.Write "    <div class=""lhd-nav-section-title"">Principal</div>"
    Dim actUsrHome, actUsrNew, actUsrTickets, actUsrKb
    actUsrHome = "" : actUsrNew = "" : actUsrTickets = "" : actUsrKb = ""
    If InStr(curScript, "/user/default.asp") > 0 Or (InStr(curScript, "/user/") > 0 And Right(curScript, 6) = "/user/") Then actUsrHome = " active"
    If InStr(curScript, "/user/new.asp") > 0 Then actUsrNew = " active"
    If InStr(curScript, "/user/view.asp") > 0 Or InStr(curScript, "/user/details.asp") > 0 Then actUsrTickets = " active"
    If InStr(curScript, "/kb/") > 0 Then actUsrKb = " active"

    Response.Write "    <a class=""lhd-nav-link" & actUsrHome & """ href=""" & baseURL & "/user/"">"
    Response.Write "      <span class=""nav-icon"">&#127968;</span><span>Inicio</span>"
    Response.Write "    </a>"

    Response.Write "    <a class=""lhd-nav-link lhd-nav-cta" & actUsrNew & """ href=""" & baseURL & "/user/new.asp"">"
    Response.Write "      <span class=""nav-icon"">&#10133;</span><span>Nueva Solicitud</span>"
    Response.Write "    </a>"

    Response.Write "    <div class=""lhd-nav-section-title"">Mis Solicitudes</div>"

    If closedView And actUsrTickets <> "" Then
      actUsrTickets = ""
      actClosedLink = " active"
    End If
    Response.Write "    <a class=""lhd-nav-link" & actUsrTickets & """ href=""" & baseURL & "/user/view.asp"">"
    Response.Write "      <span class=""nav-icon"">&#128203;</span><span>Mis Tickets</span>"
    Response.Write "    </a>"

    Response.Write "    <a class=""lhd-nav-link" & actClosedLink & """ href=""" & baseURL & "/user/view.asp?st=closed"">"
    Response.Write "      <span class=""nav-icon"">&#9989;</span><span>Tickets Cerrados</span>"
    Response.Write "    </a>"

    If SafeCInt(Cfg(cnnDB, "EnableKB")) >= 2 Then
      Response.Write "    <div class=""lhd-nav-section-title"">Ayuda</div>"
      Response.Write "    <a class=""lhd-nav-link" & actUsrKb & """ href=""" & baseURL & "/kb/"">"
      Response.Write "      <span class=""nav-icon"">&#128218;</span><span>Base Conocimiento</span>"
      Response.Write "    </a>"
    End If

  End If

  Response.Write "    <div class=""lhd-nav-section-title"">Cuenta</div>"
  Response.Write "    <a class=""lhd-nav-link"" href=""" & baseURL & "/register.asp?edit=1"">"
  Response.Write "      <span class=""nav-icon"">&#9881;</span><span>Editar Datos</span>"
  Response.Write "    </a>"
  Response.Write "    <a class=""lhd-nav-link"" href=""" & baseURL & "/logoff.asp"" style=""color:#f87171 !important"">"
  Response.Write "      <span class=""nav-icon"">&#128682;</span><span>Cerrar Sesi&oacute;n</span>"
  Response.Write "    </a>"

  Response.Write "  </nav>"
  Response.Write "</aside>"

  ' Mobile Overlay
  Response.Write "<div class=""lhd-sidebar-overlay"" id=""lhdOverlay"" onclick=""toggleLhdSidebar()""></div>"

  ' Main Wrapper Opening
  Response.Write "<div class=""lhd-main-wrapper"" id=""lhdMainWrapper"">"

  ' Topbar
  Response.Write "  <header class=""lhd-topbar"">"
  Response.Write "    <div class=""lhd-topbar-left"">"
  Response.Write "      <button type=""button"" class=""lhd-sidebar-toggle"" onclick=""toggleLhdSidebar()"" aria-label=""Men&uacute;"">&#9776;</button>"
  Response.Write "      <div class=""lhd-breadcrumb"">"
  Response.Write "        <span>" & FixTildes(Server.HTMLEncode(Cfg(cnnDB, "SiteName"))) & "</span>"
  Response.Write "        <span style=""color:#cbd5e1"">&#8250;</span>"
  Response.Write "        <span style=""color:var(--lhd-primary);font-weight:700"">Mesa de Ayuda</span>"
  Response.Write "      </div>"
  Response.Write "    </div>"
  Response.Write "    <div class=""lhd-topbar-right"">"
  If showRep Then
    Response.Write "      <a class=""btn btn-primary btn-sm"" href=""" & baseURL & "/rep/new.asp"">&#10133; Nuevo Ticket</a>"
  Else
    Response.Write "      <a class=""btn btn-primary btn-sm"" href=""" & baseURL & "/user/new.asp"">&#10133; Nuevo Ticket</a>"
  End If
  Response.Write "      <a class=""btn btn-secondary btn-sm"" href=""" & baseURL & "/logoff.asp"" title=""Cerrar Sesi&oacute;n"">&#128682; Salir</a>"
  Response.Write "    </div>"
  Response.Write "  </header>"

  ' Main Content Container Opening
  Response.Write "  <main class=""lhd-content"" id=""lhdContent"">"

End Sub


' DisplayFooter:
' Procedure to write the HTML for a sticky footer (modernized).
Sub DisplayFooter(cnnDB, sid)
  On Error Resume Next

  If Session("lhd_HeaderRendered") = True Then
    Session("lhd_HeaderRendered") = False
    Response.Write "  </main>" & vbNewLine
    Response.Write "  <footer class=""lhd-footer"">" & vbNewLine
    Response.Write "    &copy; " & Year(Now) & " <a href=""http://www.liberum.org"">Liberum Help Desk</a>."
    Response.Write "    " & lang(cnnDB, "Pleaseviewthe") & " <a href=""" & Cfg(cnnDB, "BaseURL") & "/license.html"">" & lang(cnnDB, "license") & "</a>."
    Response.Write "  </footer>" & vbNewLine
    Response.Write "</div>" & vbNewLine
    Response.Write "<scr" & "ipt type=""text/javascript"">" & vbNewLine
    Response.Write "function toggleLhdSidebar() {" & vbNewLine
    Response.Write "  var sb = document.getElementById('lhdSidebar');" & vbNewLine
    Response.Write "  var ov = document.getElementById('lhdOverlay');" & vbNewLine
    Response.Write "  if (sb) sb.classList.toggle('open');" & vbNewLine
    Response.Write "  if (ov) ov.classList.toggle('open');" & vbNewLine
    Response.Write "}" & vbNewLine
    Response.Write "document.addEventListener('DOMContentLoaded', function() {" & vbNewLine
    Response.Write "  var filterInput = document.getElementById('tbl-filter');" & vbNewLine
    Response.Write "  if (filterInput) {" & vbNewLine
    Response.Write "    filterInput.addEventListener('keyup', function() {" & vbNewLine
    Response.Write "      var val = this.value.toLowerCase();" & vbNewLine
    Response.Write "      var rows = document.querySelectorAll('.lhd-table tbody tr, table.Normal tbody tr');" & vbNewLine
    Response.Write "      for (var i = 0; i < rows.length; i++) {" & vbNewLine
    Response.Write "        var text = rows[i].textContent || rows[i].innerText;" & vbNewLine
    Response.Write "        rows[i].style.display = text.toLowerCase().indexOf(val) > -1 ? '' : 'none';" & vbNewLine
    Response.Write "      }" & vbNewLine
    Response.Write "    });" & vbNewLine
    Response.Write "  }" & vbNewLine
    Response.Write "});" & vbNewLine
    Response.Write "</scr" & "ipt>" & vbNewLine
  Else
    ' Pagina sin cabecera del sistema (login, registro...): aplicar tambien la paleta de colores
    Dim ftP, ftH, ftA, ftB
    ftP = GetThemeColor(cnnDB, "primary", "#0057B8")
    ftH = GetThemeColor(cnnDB, "primary_hover", "#003F8A")
    ftA = GetThemeColor(cnnDB, "accent", "#0284C7")
    ftB = GetThemeColor(cnnDB, "bg", "#F1F5F9")
    If LhdIsHexColor(ftP) And LhdIsHexColor(ftH) And LhdIsHexColor(ftA) And LhdIsHexColor(ftB) Then
      Response.Write "<style type=""text/css"">:root{--color-primary:" & ftP & ";--color-primary-dark:" & ftH & ";--color-accent:" & ftA & ";--color-bg:" & ftB & ";}</style>" & vbNewLine
    End If
    Response.Write "<footer class=""lhd-footer"">" & vbNewLine
    Response.Write "  &copy; " & Year(Now) & " <a href=""http://www.liberum.org"">Liberum Help Desk</a>."
    Response.Write "  " & lang(cnnDB, "Pleaseviewthe") & " <a href=""" & Cfg(cnnDB, "BaseURL") & "/license.html"">" & lang(cnnDB, "license") & "</a>."
    Response.Write "</footer>" & vbNewLine
  End If
  ' Medicion de rendimiento (comentario HTML visible solo con sesion de administrador).
  ' El inicio y el fin del comentario HTML se escriben partidos: el analizador de ASP no debe verlos.
  If Session("lhd_IsAdmin") = True Then
    Response.Write "<" & "!-- lhd-perf: " & lhdQryCount & " consultas, " & FormatNumber(lhdQryMs, 0) & " ms en base de datos, " & _
      FormatNumber((Timer - lhdPageStart) * 1000, 0) & " ms desde la conexion hasta el pie, DefaultLang fallbacks: " & lhdLangDbFallbacks & " --" & ">" & vbNewLine
  End If
  ' Buscador en todas las listas desplegables (<select>) de la pagina
  ' Iconos SVG en lugar de emojis (los emojis se ven como un cuadro si la fuente no los tiene)
  Response.Write "<scr" & "ipt type=""text/javascript"" src=""" & GetAppPath() & "/js/lhd-icons.js?v=2""></scr" & "ipt>" & vbNewLine
  Response.Write "<scr" & "ipt type=""text/javascript"" src=""" & GetAppPath() & "/searchselect.js?v=20261110d""></scr" & "ipt>" & vbNewLine

  ' Notificaciones Toastify (libreria local en js/toastify) + cola de mensajes del servidor
  Dim tstBase, tstQueue, tstItems, tstI, tstParts, tstJs
  tstBase = GetAppPath()
  Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""" & tstBase & "/js/toastify/toastify.css?v=20261110b"">" & vbNewLine
  tstJs = ""
  tstQueue = Session("lhd_toasts") & ""
  If Len(tstQueue) > 0 Then
    tstItems = Split(tstQueue, vbLf)
    For tstI = 0 To UBound(tstItems)
      tstParts = Split(tstItems(tstI), "|", 2)
      If UBound(tstParts) = 1 Then
        If Len(tstJs) > 0 Then tstJs = tstJs & ","
        tstJs = tstJs & "[""" & LhdJsEsc(tstParts(0)) & """,""" & LhdJsEsc(tstParts(1)) & """]"
      End If
    Next
    Session("lhd_toasts") = ""
  End If
  Response.Write "<scr" & "ipt type=""text/javascript"">window.__lhdToastQueue=[" & tstJs & "];</scr" & "ipt>" & vbNewLine
  Response.Write "<scr" & "ipt type=""text/javascript"" src=""" & tstBase & "/js/toastify/toastify.js?v=20261110b""></scr" & "ipt>" & vbNewLine
  Response.Write "<scr" & "ipt type=""text/javascript"" src=""" & tstBase & "/js/lhd-toast.js?v=20261110b""></scr" & "ipt>" & vbNewLine
End Sub


' LhdUserDisplayName:
' Nombre completo (fname) del usuario con ese uid; si no tiene, devuelve el uid.
Function LhdUserDisplayName(cnnDB, uidVal)
  Dim rs, nm
  nm = ""
  On Error Resume Next
  Set rs = SQLQuery(cnnDB, "SELECT fname FROM tblUsers WHERE uid='" & Replace(uidVal & "", "'", "''") & "'")
  If Not rs.EOF Then nm = Trim(rs("fname") & "")
  rs.Close
  On Error Goto 0
  If Len(nm) = 0 Then nm = uidVal & ""
  LhdUserDisplayName = nm
End Function


' IIf2: equivalente a IIf (VBScript no lo trae)
Function IIf2(cond, a, b)
  If cond Then
    IIf2 = a
  Else
    IIf2 = b
  End If
End Function


' LhdToast:
' Deja una notificacion (Toastify) para mostrar en la siguiente pagina que
' llame a DisplayFooter (sirve tambien tras un Response.Redirect).
' kind: "success", "error", "warning" o "info".
Sub LhdToast(kind, msg)
  Dim k
  k = LCase(kind & "")
  If k <> "success" And k <> "error" And k <> "warning" And k <> "info" Then k = "info"
  msg = Replace(Replace(msg & "", vbCr, " "), vbLf, " ")
  Session("lhd_toasts") = Session("lhd_toasts") & k & "|" & msg & vbLf
End Sub


' LhdJsEsc:
' Escapa un texto para ponerlo dentro de una cadena JavaScript entre comillas dobles.
Function LhdJsEsc(s)
  s = s & ""
  s = Replace(s, "\", "\\")
  s = Replace(s, """", "\""")
  s = Replace(s, vbCr, "")
  s = Replace(s, vbLf, "\n")
  s = Replace(s, "<", "<")
  LhdJsEsc = s
End Function


' Funciones auxiliares de la interfaz

' PriorityBadge:
' Devuelve un badge HTML con color segun el nombre de prioridad.
Function PriorityBadge(pname)
  Dim cls, icon
  Select Case LCase(Trim(pname))
    Case "alta","high","urgent","critica","critical"
      cls  = "badge-high"
      icon = "<span class=""lhd-dot""></span>"
    Case "media","medium","normal"
      cls  = "badge-medium"
      icon = "<span class=""lhd-dot""></span>"
    Case "baja","low"
      cls  = "badge-low"
      icon = "<span class=""lhd-dot""></span>"
    Case Else
      cls  = "badge-normal"
      icon = "<span class=""lhd-dot""></span>"
  End Select
  PriorityBadge = "<span class=""badge " & cls & """>" & icon & " " & FixTildes(Server.HTMLEncode(pname)) & "</span>"
End Function


' StatusBadge:
' Devuelve un badge HTML con color segun el nombre de estado.
Function StatusBadge(sname)
  Dim cls, icon
  Select Case LCase(Trim(sname))
    Case "abierto","open","new","nuevo"
      cls  = "badge-open"
      icon = "<span class=""lhd-dot""></span>"
    Case "cerrado","closed","resuelto","resolved"
      cls  = "badge-closed"
      icon = "&#10004;"
    Case "reabierto","reopened"
      cls  = "badge-reopen"
      icon = "&#8635;"
    Case Else
      cls  = "badge-pending"
      icon = "&#9201;"
  End Select
  StatusBadge = "<span class=""badge " & cls & """>" & icon & " " & FixTildes(Server.HTMLEncode(sname)) & "</span>"
End Function


' InoutIcon:
' Devuelve un indicador de punto de color para el tablero In/Out.
Function InoutIcon(statuscode)
  Select Case CInt(statuscode)
    Case 0
      InoutIcon = "<span class=""inout-dot inout-in"" title=""Disponible""></span>"
    Case 1
      InoutIcon = "<span class=""inout-dot inout-out"" title=""Ausente""></span>"
    Case 2
      InoutIcon = "<span class=""inout-dot inout-leave"" title=""Permiso""></span>"
    Case Else
      InoutIcon = "<span class=""inout-dot inout-in"" title=""Disponible""></span>"
  End Select
End Function


' Paginate:
' Genera controles de paginacion HTML reutilizables.
Function Paginate(pageNum, totalPages, baseUrl)
  Dim html, i
  If totalPages <= 1 Then
    Paginate = ""
    Exit Function
  End If
  html = "<div class=""lhd-pagination"">"
  If pageNum > 1 Then
    html = html & "<a href=""" & baseUrl & "&page=" & (pageNum-1) & """>&laquo;</a>"
  Else
    html = html & "<span class=""pg-disabled"">&laquo;</span>"
  End If
  For i = 1 To totalPages
    If i = pageNum Then
      html = html & "<span class=""pg-current"">" & i & "</span>"
    Else
      html = html & "<a href=""" & baseUrl & "&page=" & i & """>" & i & "</a>"
    End If
  Next
  If pageNum < totalPages Then
    html = html & "<a href=""" & baseUrl & "&page=" & (pageNum+1) & """>&raquo;</a>"
  Else
    html = html & "<span class=""pg-disabled"">&raquo;</span>"
  End If
  html = html & "</div>"
  Paginate = html
End Function


' ISODate:
' Devuelve una fecha en formato YYYY-MM-DD para input type=date
Function ISODate(d)
  If IsDate(d) Then
    ISODate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
  Else
    ISODate = ""
  End If
End Function


' HighlightTerm:
' Envuelve el termino buscado en <mark> dentro de un texto.
Function HighlightTerm(sText, sTerm)
  If Len(Trim(sTerm)) = 0 Then
    HighlightTerm = Server.HTMLEncode(sText)
  Else
    HighlightTerm = Replace(Server.HTMLEncode(sText), _
                     Server.HTMLEncode(sTerm), _
                     "<mark style=""background:#FFF176;border-radius:2px;padding:0 2px"">" & _
                     Server.HTMLEncode(sTerm) & "</mark>", _
                     1, -1, 1)
  End If
End Function


' ShowToast:
' Muestra una alerta/toast basada en el parametro ?msg= de la URL.
Sub ShowToast(cnnDB)
  Dim msg
  msg = Request.QueryString("msg")
  Select Case LCase(Trim(msg))
    Case "created"
      Response.Write "<div class=""alert alert-success"">" & _
        "<span class=""alert-icon"">&#10004;</span>" & _
        "<span>Ticket creado con &eacute;xito. Recibir&aacute;s un correo de confirmaci&oacute;n.</span>" & _
        "<button class=""alert-dismiss"" onclick=""this.parentElement.style.display='none'"">&#10005;</button>" & _
        "</div>"
    Case "updated"
      Response.Write "<div class=""alert alert-success"">" & _
        "<span class=""alert-icon"">&#10004;</span>" & _
        "<span>Ticket actualizado correctamente.</span>" & _
        "<button class=""alert-dismiss"" onclick=""this.parentElement.style.display='none'"">&#10005;</button>" & _
        "</div>"
    Case "saved"
      Response.Write "<div class=""alert alert-success"">" & _
        "<span class=""alert-icon"">&#10004;</span>" & _
        "<span>Cambios guardados correctamente.</span>" & _
        "<button class=""alert-dismiss"" onclick=""this.parentElement.style.display='none'"">&#10005;</button>" & _
        "</div>"
    Case "deleted"
      Response.Write "<div class=""alert alert-warning"">" & _
        "<span class=""alert-icon"">&#9888;</span>" & _
        "<span>Elemento eliminado.</span>" & _
        "<button class=""alert-dismiss"" onclick=""this.parentElement.style.display='none'"">&#10005;</button>" & _
        "</div>"
  End Select
End Sub


' SQLQuery:
' Takes an input string which is the SQL query statment and
' returns the results of the query as a recordset if any
' results are returned.
Function SQLQuery(cnnDB, queryStr)

	If Application("Debug") Then
		If InStr(Lcase(queryStr), "config")=0 Then
			Response.Write("<p><b>SQL Query: </b>" & queryStr & "</p><p>")
		End If
	Else
		On Error Resume Next
	End If
  
  Dim adoRec
	' Un error pendiente de codigo anterior (On Error Resume Next en quien nos llama)
	' no debe confundirse con un fallo de esta consulta.
	Err.Clear
	Set adoRec = Server.CreateObject("ADODB.Recordset")

	adoRec.ActiveConnection = cnnDB

	Dim qT0
	qT0 = Timer
	adoRec.Open(queryStr)
	lhdQryCount = lhdQryCount + 1
	lhdQryMs = lhdQryMs + ((Timer - qT0) * 1000)

	Set SQLQuery = adoRec

	' Tras un INSERT/UPDATE/DELETE se vacian las caches por peticion (Cfg, Usr)
	Dim qHead
	qHead = LCase(Left(LTrim(queryStr), 6))
	If qHead = "insert" Or qHead = "update" Or qHead = "delete" Then Call LhdCacheReset()

	If Err.Number <> 0 Then
		' Para administradores se agrega la consulta que fallo, para poder diagnosticar.
		If Session("lhd_IsAdmin") = True Then
			Call TrapError(Err.Number, Err.Description & "<br /><b>SQL: </b>" & Server.HTMLEncode(queryStr), Err.Source)
		Else
			Call TrapError(Err.Number, Err.Description, Err.Source)
		End If
	End If

End Function


' GetUnique:
' Finds a unique key for a database.  The database name
' is given (dbname) and the lookup done in a table called
' db_keys.  This proc should be atomic, but isn't.  May cause
' some problems on a busy site.
Function GetUnique(cnnDB, dbname)

	Dim queryStr, key, adoRec

	queryStr = _
	"SELECT " & dbname & " FROM db_keys"

	Set adoRec = SQLQuery(cnnDB, queryStr)

	' Get the key from the results and return it
	key = Cint(adoRec(dbname))

	adoRec.Close

	GetUnique = key

	' Increment the key and update the database for next time.
	queryStr = _
	"UPDATE db_keys SET " & dbname & "=" & (key+1)

	Set adoRec = SQLQuery(cnnDB, queryStr)

End Function


' CheckUser:
' Checks to see if user is logged on
Sub CheckUser(cnnDB, sid)
  Dim userchkRes
	Set userchkRes = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & sid)
	If (userchkRes.EOF) OR (sid = 0) Then
		Dim reAddr

		reAddr = Cfg(cnnDB, "BaseURL") & "/logon.asp?URL=" & _
			Request.ServerVariables("PATH_INFO")
		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
			reAddr = reAddr & _
				"?" & Request.ServerVariables("QUERY_STRING")
		End If
		userchkRes.close
		cnnDB.Close
		Response.Clear
		Response.Redirect reAddr
	End If
	userchkRes.close
End Sub


'
' CheckRep:
' Checks to see if IsRep is TRUE, If not, returns
' an error.  Used for permission check on pages.
Sub CheckRep(cnnDB, sid)
  Dim userchkRes
	Set userchkRes = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & sid)
	If (userchkRes.EOF) OR (sid = 0) Then
		Dim reAddr

		reAddr = Cfg(cnnDB, "BaseURL") & "/logon.asp?URL=" & _
			Request.ServerVariables("PATH_INFO")
		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
			reAddr = reAddr & _
				"?" & Request.ServerVariables("QUERY_STRING")
		End If
		userchkRes.Close
		cnnDB.Close
		Response.Clear
		Response.Redirect reAddr
	End If
	userchkRes.Close

	If Usr(cnnDB, sid, "IsRep") <> 1 Then
		Call DisplayError(3, "Access denied.  You do not have permission to view this page.")
	End If
End Sub

' CheckKB:
' Checks to see if user has permissions to the KB
' EnableKB 0-Disable, 1-Rep, 2-User/Rep, 3-AnyOne
Sub CheckKB(cnnDB, sid)
  Dim blnOK
  blnOK = False
  Select Case Cfg(cnnDB, "EnableKB")
    Case 0
      blnOK = False
    Case 1
      Dim rstRep
      Set rstRep = SQLQuery(cnnDB, "SELECT IsRep FROM tblUsers WHERE sid=" & sid)
      If Cint(rstRep("IsRep")) = 1 Then
        blnOK = True
      Else
        blnOK = False
      End If
      rstRep.Close
    Case 2
      If sid > 0 Then
        blnOK = True
      Else
        blnOK = False
      End If
    Case 3
      blnOK = True
    Case Else
      blnOK = False
  End Select
	If Not blnOK Then
		cnnDB.Close
    Call DisplayError(3, lang(cnnDB, "Accessdenied") & " " & lang(cnnDB, "NoPermission")  & ".")
	End If
End Sub


' CheckAdmin:
' Same as CheckRep, except looks for admin privs
Sub CheckAdmin
	If Not Session("lhd_IsAdmin") Then
    Call DisplayError(3, lang(cnnDB, "Accessdenied") & " " & lang(cnnDB, "NoPermission") & ".")
	End If
End Sub


' SendMail:
' Sends mail using the supported system set in global.asa
Sub SendMail (strToAddr, strFromAddr, strFromName, strSubject, strBody, cnnDB)

Dim Mail

If Not Application("Debug") Then
	On Error Resume Next
End If

Select Case Cfg(cnnDB, "EmailType")
	Case 0	'No Email

	Case 1	'CDONTS
		Set Mail = Server.CreateObject("CDONTS.NewMail")
		Mail.BodyFormat = 1	' Text Only, 0 for HTML
		Mail.Subject = strSubject
		Mail.To = strToAddr
		Mail.Body = strBody
		Mail.Send(strFromName & "<" & strFromAddr & ">")
		Set Mail = Nothing

	Case 2	'Jmail
		On Error Resume Next ' Use Jmail logging
		Set Mail = Server.CreateObject("Jmail.Message")
		Mail.Logging = True
		Mail.From = strFromAddr
		Mail.FromName = strFromName
		Mail.AddRecipient strToAddr
		Mail.Subject = strSubject
		Mail.Body = strBody
		If Not Mail.Send(Cfg(cnnDB, "SMTPServer")) Then
			If Application("Debug") Then
				Call DisplayError(3, Mail.Log)
			End If
		End If
		Set Mail = Nothing

	Case 3	'ASPEmail
		Set Mail = Server.CreateObject("Persits.MailSender")
		Mail.Host = Cfg(cnnDB, "SMTPServer")
		Mail.From = strFromAddr
		Mail.FromName = strFromName
		Mail.AddAddress strToAddr
		Mail.Subject = strSubject
		Mail.Body = strBody
		Mail.Send
		Set Mail = Nothing

  Case 4  ' ASPMail
    Set Mail = Server.CreateObject("SMTPsvg.Mailer")
    Mail.FromName = strFromName
    Mail.FromAddress = strFromAddr
    Mail.RemoteHost = Cfg(cnnDB, "SMTPServer")
    Mail.AddRecipient "Help Desk User", strToAddr
    Mail.Subject = strSubject
    Mail.BodyText = strBody
    Mail.SendMail
  Case 5  ' CDOSYS
    DIM strIncomingMailServer, objCDOSYSMail, objCDOSYSCon
    'Create the e-mail server object
	strIncomingMailServer = Cfg(cnnDB, "SMTPServer")
	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
    Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

    'Set and update fields properties
	With objCDOSYSCon
        'Out going SMTP server
        .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strIncomingMailServer
        'SMTP port
        .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  = 25
        'CDO Port
        .Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        '.Timeout
        .Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		.Fields.Update
   End With
   'Update the CDOSYS Configuration 
   Set objCDOSYSMail.Configuration = objCDOSYSCon
   With objCDOSYSMail
	.From = """" & strFromName & """ <" & strFromAddr & ">"
	'.From = Cfg(cnnDB, "HDName") & " <" & Cfg(cnnDB, "HDReply") & ">"
	.To = strToAddr
	.Subject = strSubject
	'.HTMLBody = strBody 'Set the e-mail body format (HTMLBody=HTML TextBody=Plain)
	.TextBody = strBody
	'Send the e-mail
	If NOT strIncomingMailServer = "" Then .Send
   End with
  'Close the server mail object
   Set objCDOSYSMail = Nothing 
End Select

If Err.Number <> 0 Then
	Call TrapError(Err.Number, Err.Description, Err.Source)
End If

End Sub

' SendMailHTML:
' Sends mail using the supported system set in global.asa
Sub SendMailHTML (strToAddr, strFromAddr, strFromName, strSubject, strBody, cnnDB)

	Dim Mail

	If Not Application("Debug") Then
		On Error Resume Next
	End If

	Select Case Cfg(cnnDB, "EmailType")
	Case 0	'No Email

	Case 1	'CDONTS
		Set Mail = Server.CreateObject("CDONTS.NewMail")
		Mail.BodyFormat = 1	' Text Only, 0 for HTML
		Mail.Subject = strSubject
		Mail.To = strToAddr
		Mail.Body = strBody
		Mail.Send(strFromName & "<" & strFromAddr & ">")
		Set Mail = Nothing

	Case 2	'Jmail
		On Error Resume Next ' Use Jmail logging
		Set Mail = Server.CreateObject("Jmail.Message")
		Mail.Logging = True
		Mail.From = strFromAddr
		Mail.FromName = strFromName
		Mail.AddRecipient strToAddr
		Mail.Subject = strSubject
		Mail.Body = strBody
		If Not Mail.Send(Cfg(cnnDB, "SMTPServer")) Then
			If Application("Debug") Then
				Call DisplayError(3, Mail.Log)
			End If
		End If
		Set Mail = Nothing

	Case 3	'ASPEmail
		Set Mail = Server.CreateObject("Persits.MailSender")
		Mail.Host = Cfg(cnnDB, "SMTPServer")
		Mail.From = strFromAddr
		Mail.FromName = strFromName
		Mail.AddAddress strToAddr
		Mail.Subject = strSubject
		Mail.Body = strBody
		Mail.Send
		Set Mail = Nothing

	Case 4  ' ASPMail
		Set Mail = Server.CreateObject("SMTPsvg.Mailer")
		Mail.FromName = strFromName
		Mail.FromAddress = strFromAddr
		Mail.RemoteHost = Cfg(cnnDB, "SMTPServer")
		Mail.AddRecipient "Help Desk User", strToAddr
		Mail.Subject = strSubject
		Mail.BodyText = strBody
		Mail.SendMail
	Case 5  ' CDOSYS
		DIM strIncomingMailServer, objCDOSYSMail, objCDOSYSCon
		'Create the e-mail server object
		strIncomingMailServer = Cfg(cnnDB, "SMTPServer")
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

		'Set and update fields properties
		With objCDOSYSCon
			'Out going SMTP server
			.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strIncomingMailServer
			'SMTP port
			.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  = 25
			'CDO Port
			.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			'.Timeout
			.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
			.Fields.Update
		End With
   'Update the CDOSYS Configuration 
		Set objCDOSYSMail.Configuration = objCDOSYSCon
			With objCDOSYSMail
				.From = """" & strFromName & """ <" & strFromAddr & ">"
				.To = strToAddr
				.Subject = strSubject
				.TextBody = strBody
				'Send the e-mail
				If NOT strIncomingMailServer = "" Then .Send
			End with
			'Close the server mail object
		Set objCDOSYSMail = Nothing 
	End Select

	If Err.Number <> 0 Then
		Call TrapError(Err.Number, Err.Description, Err.Source)
	End If

End Sub


' Message:
' Parses and sends an email
Sub eMessage (cnnDB, eType, id, strToAddr)
    ' Get the email message from the database
	Dim emailRes
	Set emailRes = SQLQuery(cnnDB, "Select subject, body FROM tblEmailMsg WHERE type='" & eType & "'")
	If emailRes.EOF Then
		Call DisplayError(3, lang(cnnDB, "Nomessageoftype") & " " & eType & " " & lang(cnnDB, "wasfoundinthedatabase") & ".")
	End If

	Dim subject, body, strFromAddr, strFromName
	subject = emailRes("subject")
	body = emailRes("body")

	emailRes.Close

	Dim queryStr

	queryStr = _
		"SELECT p.id, p.uid, p.uemail, p.uphone, p.ulocation, d.dname, p.start_date, p.status, s.sname, " & _
		"p.close_date, pri.pname, c.cname, p.rep, p.title, p.solution, p.description " & _
		"FROM (((problems AS p " & _
		"INNER JOIN departments AS d ON p.department = d.department_id) " & _
		"INNER JOIN status AS s ON p.status = s.status_id) " & _
		"INNER JOIN priority AS pri ON p.priority = pri.priority_id) " & _
		"INNER JOIN categories AS c ON p.category = c.category_id " & _
		"WHERE p.id=" & id

	Dim probRes, repRes, userRes, notesRes, notes, notesHTML
	Set probRes = SQLQuery(cnnDB, queryStr)
	Set repRes = SQLQuery(cnnDB, "SELECT uid, email1, fname FROM tblUsers WHERE sid=" & probRes("rep"))
	Set userRes = SQLQuery(cnnDB, "SELECT fname FROM tblUsers WHERE uid='" & probRes("uid") & "'")
	If InStr(eType, "rep") > 0 Then
      Set notesRes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " ORDER BY addDate ASC")
	Else
	  Set notesRes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " AND private=0 ORDER BY addDate ASC")
	End If

	If InStr(eType, "user") > 0 Then
		strFromAddr = repRes("email1")
		strFromName = repRes("fname")
	Else
		strFromAddr = Cfg(cnnDB, "HDReply")
		strFromName = Cfg(cnnDB, "HDName")
	End If


	If probRes.EOF Then
		cnnDB.Close
		Call DisplayError(3, lang(cnnDB, "Problem") & " " & id & " " & lang(cnnDB, "doesnotexist") & ". " & lang(cnnDB, "Cannotsendmail") & ".")
	End If

		notes = ""
		notesHTML = ""
		
	If Not notesRes.EOF Then
		Do While Not notesRes.EOF
		If Len(notes) > 0 Then
			notes = notes & vbNewLine
		End If
		notes = notes & "[" & notesRes("addDate") & " - " & notesRes("uid") & "]"
		notes = notes & vbNewLine
		notes = notes & notesRes("note") & vbNewLine
		notesRes.MoveNext
		Loop
	Else
		notes = " "
	End If
	
	notesRes.Close
	
	' Parse out the notes array for HTML formatting

	If Not Application("Debug") Then
		On Error Resume Next
	End If
	body = Replace(body, "[problemid]", probRes("id"))
	body = Replace(body, "[title]", probRes("title"))
	body = Replace(body, "[description]", probRes("description"))
	body = Replace(body, "[status]", probRes("sname"))
	body = Replace(body, "[priority]", probRes("pname"))
	body = Replace(body, "[startdate]", DisplayDate(probRes("start_date"), lhdDateTime))
	body = Replace(body, "[closedate]", DisplayDate(probRes("close_date"), lhdDateTime))
	On Error Resume Next: body = Replace(body, "[duedate]", ""): On Error Goto 0
	body = Replace(body, "[category]", probRes("cname"))
	body = Replace(body, "[department]", probRes("dname"))
	body = Replace(body, "[phone]", probRes("uphone"))
	body = Replace(body, "[location]", probRes("ulocation"))
	body = Replace(body, "[solution]", probRes("solution"))
	body = Replace(body, "[baseurl]", Cfg(cnnDB, "BaseURL"))
	body = Replace(body, "[uid]", probRes("uid"))
	body = Replace(body, "[ufname]", userRes("fname"))
	body = Replace(body, "[uemail]", probRes("uemail"))
	body = Replace(body, "[rid]", repRes("uid"))
	body = Replace(body, "[rfname]", repRes("fname"))
	body = Replace(body, "[remail]", repRes("email1"))
	body = Replace(body, "[uurl]", Cfg(cnnDB, "BaseURL") & "/user/view.asp?id=" & id)
	body = Replace(body, "[rurl]", Cfg(cnnDB, "BaseURL") & "/rep/view.asp?id=" & id)
    body = Replace(body, "[notes]", notes)

    
	' urlencoded
	body = Replace(body, "[u_title]", Server.URLEncode(probRes("title")))
	body = Replace(body, "[u_rfname]", Server.URLEncode(repRes("fname")))
	
	subject = Replace(subject, "[problemid]", probRes("id"))
	subject = Replace(subject, "[title]", probRes("title"))
	subject = Replace(subject, "[description]", probRes("description"))
	subject = Replace(subject, "[status]", probRes("sname"))
	subject = Replace(subject, "[priority]", probRes("pname"))
	subject = Replace(subject, "[startdate]", DisplayDate(probRes("start_date"), lhdDateTime))
	subject = Replace(subject, "[closedate]", DisplayDate(probRes("close_date"), lhdDateTime))
  On Error Resume Next: subject = Replace(subject, "[duedate]", ""): On Error Goto 0
	subject = Replace(subject, "[category]", probRes("cname"))
	subject = Replace(subject, "[department]", probRes("dname"))
	subject = Replace(subject, "[phone]", probRes("uphone"))
	subject = Replace(subject, "[location]", probRes("ulocation"))
	subject = Replace(subject, "[solution]", probRes("solution"))
	subject = Replace(subject, "[baseurl]", Cfg(cnnDB, "BaseURL"))
	subject = Replace(subject, "[uid]", probRes("uid"))
	subject = Replace(subject, "[ufname]", userRes("fname"))
	subject = Replace(subject, "[uemail]", probRes("uemail"))
	subject = Replace(subject, "[rid]", repRes("uid"))
	subject = Replace(subject, "[rfname]", repRes("fname"))
	subject = Replace(subject, "[remail]", repRes("email1"))
	subject = Replace(subject, "[uurl]", Cfg(cnnDB, "BaseURL") & "/user/view.asp?id=" & id)
	subject = Replace(subject, "[rurl]", Cfg(cnnDB, "BaseURL") & "/rep/view.asp?id=" & id)

  Err.Clear
  On Error GoTo 0
	Call SendMailHTML (strToAddr, strFromAddr, strFromName, Subject, Body, cnnDB)
End Sub

' FixDay:
' Returns a day not greater than the last day of the month
Function FixDay (intMonth, intDay, intYear)
  FixDay = intDay
  If (intMonth=4) Or (intMonth=6) Or (intMonth=9) Or (intMonth=11) Then
    If intDay > 30 Then
      FixDay = 30
    End If
  End If
  If (intMonth=2) and (intDay>28) Then
    If (intYear Mod 4 = 0) Then
      FixDay=29
    Else
      FixDay=28
    End If
  End If
End Function
         

' Function to making it easy to make dropdown lists from database
' cnnDB = a open ADO Connection object
' dropdownlistname = text name of the form field
' keyfieldname = text fieldname with value returned when the form is executed
' selectedkey = value related to keyfieldname used to show initial selected value
' valuefieldname = text name of the field shown in the dropdownlist
' tablename = text name of the table or view containing the fields
' criteria = text with the criteria for selecting data ex. 'id > 1 ' (can be empty)
' sortorder = text to enable sorting data ex. ' id ASC ' can be empty

function dropdownlist(cnnDB, dropdownlistname, keyfieldname, selectedkey, _
														 valuefieldname, tablename, criteria, sortorder)

	dim tempstr, dropdownRes, sqlStr

	' build and execute SQL query
	
	sqlStr = "SELECT " & keyfieldname & ", " & valuefieldname & " FROM " & tablename
	if criteria <> "" then
		sqlStr = sqlStr & " WHERE " & criteria
	end if
	if  sortorder <> "" then
		sqlStr = sqlStr & " ORDER BY " & sortorder
	end if
  Set dropdownRes = SQLQuery(cnnDB, sqlStr)

	tempstr = "<SELECT NAME=""" & dropdownlistname & """>"

  If Not dropdownRes.EOF Then
    Do While Not dropdownRes.EOF
      tempstr = tempstr & "<OPTION VALUE=""" & dropdownRes(keyfieldname) & """"
      if dropdownRes(keyfieldname)=selectedkey then
	      tempstr = tempstr & " SELECTED "
	    end if  
      tempstr = tempstr & ">" & dropdownRes(valuefieldname) & "</OPTION>"
      dropdownRes.MoveNext
    Loop
  End If

	tempstr = tempstr & "</SELECT>"
  
	dropdownlist = tempstr

end function

         
' SQLDate
' Returns a correctly formated date for a SQL statement including
' the correct delimitations.
' dtDate = Date format recognized by CDate
Function SQLDate (dtDate, intUseDelim)
  If Len(Trim(dtDate)) > 0 Then
    dtDate = CDate(dtDate)
    SQLDate = Year(dtDate) & "-" & Month(dtDate) & "-" & Day(dtDate)
    SQLDate = SQLDate & " " & Hour(dtDate) & ":" & Minute(dtDate) & ":" & Second(dtDate)
    Dim strDeLim
    If Application("DBType") = 1 Or Application("DBType") = 2 Or Application("DBType") = 4 Then
      strDeLim = "'"
    Else
      strDeLim = "#"
    End If
    If intUseDelim = lhdAddSQLDelim Then
      SQLDate = strDelim & SQLDate & strDeLim
    End If
  Else
    SQLDate = ""
  End If
End Function

' DisplayDate
' Formats the date and time for the locale
' dtDate = Date format recognized by CDate
' intFormat = 0 - Date Only
'           = 1 - Date and Time
Function DisplayDate(dtDate, intFormat)
  If Len(Trim(dtDate)) > 0 Then
    On Error Resume Next
    dtDate = CDate(dtDate)
    If Err.Number <> 0 Then
      Err.Clear
      DisplayDate = ""
      Exit Function
    End If

    Dim yr, mo, dy, strLocalDate
    yr = CStr(Year(dtDate))
    mo = Right("0" & Month(dtDate), 2)
    dy = Right("0" & Day(dtDate), 2)
    strLocalDate = yr & "-" & mo & "-" & dy
    If intFormat = lhdDateTime Then
      strLocalDate = strLocalDate & " " & FormatDateTime(dtDate, vbLongTime)
    End If
    DisplayDate = strLocalDate
    On Error Goto 0
  Else
    DisplayDate = ""
  End If
End Function

' ConvertFormattedDate
' Takes a display formated date and returns
' a VB date
Function ConvertFormattedDate (strInputDate)
  Dim strUserFormat, strDay, strMonth, strYear, strSplitChar, varUser, varInput
  ConvertFormattedDate = Null
  strUserFormat = "yyyy-mm-dd"
  On Error Resume Next

  Dim intChar
  intChar = 1
  Do Until NOT ((Mid(strUserFormat, intChar, 1) >= "a") AND (Mid(strUserFormat, intChar, 1) <= "z"))
    intChar = intChar + 1
    If intChar > Len(strUserFormat) Then Exit Do
  Loop
  If intChar <= Len(strUserFormat) Then
    strSplitChar = Mid(strUserFormat, intChar, 1)
  Else
    strSplitChar = "-"
  End If
  varUser = Split(strUserFormat, strSplitChar)
  varInput = Split(strInputDate, strSplitChar)
  If Ubound(varInput) = 2 Then
    Dim strUserSection, intInputIndex
    intInputIndex = 0
    For Each strUserSection In varUser
      Select Case strUserSection
        Case "dd"
          strDay = varInput(intInputIndex)
        Case "mm"
          strMonth = varInput(intInputIndex)
        Case "yy"
          strYear = varInput(intInputIndex)
        Case "yyyy"
          strYear = varInput(intInputIndex)
       End Select
       intInputIndex = intInputIndex + 1
    Next
    Select Case Len(strYear)
      Case 2
        strYear = "20" & strYear
      Case 4
        ' do nothing
      Case Else
        strYear = Null
     End Select
    If IsNumeric(strYear) AND IsNumeric(strMonth) AND IsNumeric(strDay) Then
      ConvertFormattedDate = CDate(strYear & "-" & strMonth & "-" & strDay)
    End If
    If Err.Number <> 0 Then
      ConvertFormattedDate = DateAdd("d", 1, Now())
    End If
  End If
  On Error Goto 0
End Function
</SCRIPT>
