<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  Response.Buffer = True
  Response.ContentType = "application/json"
  Response.CharSet = "utf-8"
  Response.Expires = -1
%>
<!-- #include file = "../public.asp" -->
<%
  ' API del correo entrante ("correo a caso").
  ' La usa el script tools/mailpoll.ps1 (tarea programada) que habla IMAP/SMTP; aqui esta toda la logica:
  '   POST ?action=config       -> ajustes del buzon (IMAP / SMTP)
  '   POST ?action=ingest       -> cuerpo = correo completo (RFC 822). Crea el caso o agrega una nota.
  '   POST ?action=outbox       -> lista de correos pendientes de enviar (acuses / pruebas)
  '   POST ?action=outbox_done  -> marca un correo como enviado o con error (?id=&ok=1&msgid=...)
  '   POST ?action=status       -> el script informa el resultado de su ejecucion (cuerpo = texto)
  ' Autenticacion: cabecera X-Mail-Token (token configurado en admin/mailconfig.asp).

  ' ==== HELPERS BEGIN ====
  Function JsonStr(v)
    Dim s
    s = CStr(v & "")
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonStr = s
  End Function

  Function Esc(s)
    Esc = Replace(s & "", "'", "''")
  End Function

  ' Comparacion de textos que no se corta en la primera diferencia
  Function SafeEq(a, b)
    Dim i, diff
    a = a & "" : b = b & ""
    If Len(a) <> Len(b) Then
      SafeEq = False
      Exit Function
    End If
    diff = 0
    For i = 1 To Len(a)
      If Asc(Mid(a, i, 1)) <> Asc(Mid(b, i, 1)) Then diff = diff + 1
    Next
    SafeEq = (diff = 0)
  End Function

  ' Direccion de correo en "Nombre <usuario@dominio>" o "usuario@dominio"
  Function ExtractAddr(s)
    Dim re, m
    ExtractAddr = ""
    s = s & ""
    Set re = New RegExp
    re.Pattern = "<([^<>\s]+@[^<>\s]+)>"
    If re.Test(s) Then
      ExtractAddr = LCase(Trim(re.Execute(s)(0).SubMatches(0)))
      Exit Function
    End If
    re.Pattern = "([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})"
    If re.Test(s) Then ExtractAddr = LCase(Trim(re.Execute(s)(0).SubMatches(0)))
  End Function

  ' Nombre visible en "Nombre <usuario@dominio>"
  Function ExtractName(s)
    Dim p, n
    s = s & ""
    p = InStr(s, "<")
    If p > 1 Then
      n = Trim(Left(s, p - 1))
      n = Replace(n, """", "")
      ExtractName = Trim(n)
    Else
      ExtractName = ""
    End If
  End Function

  ' Valor de una cabecera del mensaje (cadena vacia si no existe)
  Function Hdr(msg, nombre)
    Dim v
    v = ""
    On Error Resume Next
    v = msg.Fields("urn:schemas:mailheader:" & nombre).Value & ""
    If Err.Number <> 0 Then
      v = ""
      Err.Clear
    End If
    On Error Goto 0
    Hdr = Trim(v)
  End Function

  ' Asunto sin "RE:", "RV:", "FW:" repetidos ni la etiqueta [Caso #n]
  Function CleanSubject(s)
    Dim re
    s = Trim(s & "")
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "\[\s*Caso\s*#\s*\d+\s*\]"
    s = Trim(re.Replace(s, ""))
    re.Global = False
    re.Pattern = "^\s*(re|rv|fw|fwd|reenv|enc|r)\s*(\[\d+\])?\s*:\s*"
    Do While re.Test(s)
      s = Trim(re.Replace(s, ""))
    Loop
    If Len(s) = 0 Then s = "(sin asunto)"
    CleanSubject = s
  End Function

  ' Numero de caso en el asunto: "[Caso #123]" (0 si no hay)
  Function SubjectCaseId(s)
    Dim re, m
    SubjectCaseId = 0
    Set re = New RegExp
    re.IgnoreCase = True
    re.Pattern = "\[\s*Caso\s*#\s*(\d+)\s*\]"
    If re.Test(s & "") Then SubjectCaseId = CLng(re.Execute(s & "")(0).SubMatches(0))
  End Function

  ' Quita espacios, tabuladores y saltos de linea al inicio y al final
  Function TrimNl(s)
    Dim t
    t = s & ""
    Do While Len(t) > 0
      If Left(t, 1) = vbLf Or Left(t, 1) = vbCr Or Left(t, 1) = " " Or Left(t, 1) = vbTab Then
        t = Mid(t, 2)
      Else
        Exit Do
      End If
    Loop
    Do While Len(t) > 0
      If Right(t, 1) = vbLf Or Right(t, 1) = vbCr Or Right(t, 1) = " " Or Right(t, 1) = vbTab Then
        t = Left(t, Len(t) - 1)
      Else
        Exit Do
      End If
    Loop
    TrimNl = t
  End Function

  ' Entidades HTML frecuentes (nombres y numericas) -> caracteres
  Function DecodeEntities(t)
    Dim re, mc, m, pares, i, par
    t = t & ""
    pares = Array("nbsp| ", "lt|<", "gt|>", "quot|""", "aacute|" & ChrW(225), "eacute|" & ChrW(233), "iacute|" & ChrW(237), _
      "oacute|" & ChrW(243), "uacute|" & ChrW(250), "ntilde|" & ChrW(241), "Aacute|" & ChrW(193), "Eacute|" & ChrW(201), _
      "Iacute|" & ChrW(205), "Oacute|" & ChrW(211), "Uacute|" & ChrW(218), "Ntilde|" & ChrW(209), "uuml|" & ChrW(252), _
      "Uuml|" & ChrW(220), "iquest|" & ChrW(191), "iexcl|" & ChrW(161), "ordm|" & ChrW(186), "ordf|" & ChrW(170))
    For i = 0 To UBound(pares)
      par = Split(pares(i), "|")
      t = Replace(t, "&" & par(0) & ";", par(1))
    Next
    Set re = New RegExp
    re.Global = True
    re.Pattern = "&#(\d{2,5});"
    Set mc = re.Execute(t)
    For Each m In mc
      t = Replace(t, m.Value, ChrW(CLng(m.SubMatches(0))))
    Next
    t = Replace(t, "&amp;", "&")
    DecodeEntities = t
  End Function

  ' HTML -> texto plano
  Function HtmlToText(h)
    Dim re, t
    t = h & ""
    Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<(style|script|head)[\s\S]*?</\1\s*>"
    t = re.Replace(t, "")
    re.Pattern = "<br\s*/?>|</p\s*>|</div\s*>|</tr\s*>|</li\s*>|</h[1-6]\s*>"
    t = re.Replace(t, vbLf)
    re.Pattern = "<[^>]*>"
    t = re.Replace(t, "")
    t = DecodeEntities(t)
    t = Replace(t, vbCr, "")
    re.Pattern = "\n{3,}"
    t = re.Replace(t, vbLf & vbLf)
    HtmlToText = TrimNl(t)
  End Function

  ' Quita el texto citado de una respuesta (lo que viene despues de "El ... escribio:", "-----Mensaje original-----",
  ' el bloque De:/Enviado: de Outlook, o las lineas que empiezan con ">")
  Function StripQuoted(txt)
    Dim lines, i, ln, out, cut, re1, re2, re3, re4, nxt
    txt = Replace(Replace(txt & "", vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(txt, vbLf)
    Set re1 = New RegExp : re1.IgnoreCase = True
    re1.Pattern = "^\s*(on|el)\s.{5,200}\s(wrote|escribi\S*)\s*:?\s*$"
    Set re2 = New RegExp : re2.IgnoreCase = True
    re2.Pattern = "^\s*-{2,}\s*(original message|mensaje original|forwarded message|mensaje reenviado)\s*-{2,}\s*$"
    Set re3 = New RegExp : re3.IgnoreCase = True
    re3.Pattern = "^\s*(de|from)\s*:\s*\S"
    Set re4 = New RegExp : re4.IgnoreCase = True
    re4.Pattern = "^\s*(enviado|sent|fecha|date|para|to|asunto|subject)\s*(el)?\s*:\s*\S"
    out = ""
    cut = False
    For i = 0 To UBound(lines)
      ln = lines(i)
      If re1.Test(ln) Or re2.Test(ln) Then
        cut = True
      ElseIf re3.Test(ln) And i < UBound(lines) Then
        nxt = lines(i + 1)
        If re4.Test(nxt) Then cut = True
      End If
      If cut Then Exit For
      If Left(Trim(ln), 1) <> ">" Then out = out & ln & vbLf
    Next
    out = TrimNl(out)
    If Len(out) = 0 Then out = TrimNl(txt)
    StripQuoted = out
  End Function

  ' Texto plano -> HTML seguro (parrafos y saltos de linea)
  Function TextToHtml(t)
    Dim paras, i, p, out
    t = Replace(Replace(t & "", vbCrLf, vbLf), vbCr, vbLf)
    paras = Split(t, vbLf & vbLf)
    out = ""
    For i = 0 To UBound(paras)
      p = Trim(paras(i))
      If Len(p) > 0 Then out = out & "<p>" & Replace(Server.HTMLEncode(p), vbLf, "<br />") & "</p>"
    Next
    If Len(out) = 0 Then out = "<p>(sin contenido)</p>"
    TextToHtml = out
  End Function

  ' Correo automatico (respuesta automatica, rebote, lista, no-reply...): no se procesa
  Function IsAutoMail(msg, addr, subj, ownAddr)
    Dim v, re
    IsAutoMail = False
    v = LCase(Hdr(msg, "auto-submitted"))
    If Len(v) > 0 And v <> "no" Then IsAutoMail = True : Exit Function
    v = LCase(Hdr(msg, "precedence"))
    If v = "bulk" Or v = "junk" Or v = "list" Or v = "auto_reply" Then IsAutoMail = True : Exit Function
    If Len(Hdr(msg, "x-autoreply")) > 0 Or Len(Hdr(msg, "x-autorespond")) > 0 Then IsAutoMail = True : Exit Function
    If Len(addr) = 0 Then IsAutoMail = True : Exit Function
    If Len(ownAddr) > 0 And LCase(ownAddr) = LCase(addr) Then IsAutoMail = True : Exit Function
    Set re = New RegExp
    re.IgnoreCase = True
    re.Pattern = "^(mailer-daemon|postmaster|no-?reply|do-?not-?reply|noresponder|no-responder|bounce|bounces)[@+._-]"
    If re.Test(addr) Then IsAutoMail = True : Exit Function
    re.Pattern = "^\s*(undeliverable|delivery status notification|delivery failure|returned mail|mail delivery failed|out of office|automatic reply|auto-?reply|respuesta autom|fuera de (la )?oficina|no se pudo entregar|ausente)"
    If re.Test(subj) Then IsAutoMail = True : Exit Function
  End Function

  ' Cabeceras del correo (texto plano, sin plegar) a partir de los bytes del mensaje
  Function RawHeaderBlock(bytes)
    Dim st, txt, p, re
    On Error Resume Next
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write bytes
    st.Position = 0
    st.Type = 2
    st.Charset = "iso-8859-1"
    txt = st.ReadText(40000)
    st.Close
    If Err.Number <> 0 Then
      Err.Clear
      txt = ""
    End If
    On Error Goto 0
    p = InStr(txt, vbCrLf & vbCrLf)
    If p = 0 Then p = InStr(txt, vbLf & vbLf)
    If p > 0 Then txt = Left(txt, p)
    Set re = New RegExp
    re.Global = True
    re.Pattern = "\r?\n[ \t]+"
    RawHeaderBlock = re.Replace(txt, " ")
  End Function

  ' Direcciones de destino del correo en orden de confiabilidad: entrega real (Delivered-To, X-Original-To,
  ' Envelope-To), luego To, Cc y por ultimo "for <direccion>" de las cabeceras Received. Una por linea.
  Function RecipientCandidates(block)
    Dim names, i, reH, reA, mH, mA, outTxt, a, mRe, reF, mF
    outTxt = ""
    names = Array("x-original-to", "delivered-to", "envelope-to", "x-envelope-to", "x-forwarded-to", "to", "cc")
    block = Replace(block & "", vbCrLf, vbLf)
    Set reA = New RegExp
    reA.Global = True
    reA.Pattern = "[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"
    For i = 0 To UBound(names)
      Set reH = New RegExp
      reH.Global = True
      reH.Multiline = True
      reH.IgnoreCase = True
      reH.Pattern = "^" & names(i) & ":[ \t]*(.*)$"
      For Each mH In reH.Execute(block)
        For Each mA In reA.Execute(mH.SubMatches(0))
          a = LCase(mA.Value)
          If InStr(vbLf & outTxt, vbLf & a & vbLf) = 0 Then outTxt = outTxt & a & vbLf
        Next
      Next
    Next
    Set reF = New RegExp
    reF.Global = True
    reF.IgnoreCase = True
    reF.Pattern = "^received:.*?\sfor\s+<?([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})"
    reF.Multiline = True
    For Each mF In reF.Execute(block)
      a = LCase(mF.SubMatches(0))
      If InStr(vbLf & outTxt, vbLf & a & vbLf) = 0 Then outTxt = outTxt & a & vbLf
    Next
    RecipientCandidates = outTxt
  End Function

  ' Binario -> base64 en una sola linea
  Function BytesToB64(b)
    Dim dom, el
    Set dom = Server.CreateObject("MSXML2.DOMDocument")
    Set el = dom.createElement("b64")
    el.dataType = "bin.base64"
    el.nodeTypedValue = b
    BytesToB64 = Replace(Replace(el.text, vbLf, ""), vbCr, "")
  End Function
  ' ==== HELPERS END ====

  Sub Fail(code, msg)
    Response.Clear
    Response.Status = code
    Response.Write "{""ok"":false,""error"":""" & JsonStr(msg) & """}"
    Response.End
  End Sub

  Dim cnnDB, action, tokenIn, tokenDb, cf, rsC, fld, cfgOk
  Dim gMsgId, gFrom, gSubj, gErr
  Set cnnDB = CreateCon

  If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then Call Fail("405 Method Not Allowed", "Use POST")
  action = LCase(Request.QueryString("action"))
  tokenIn = Request.ServerVariables("HTTP_X_MAIL_TOKEN") & ""

  ' Configuracion (primera fila de tblMailConfig)
  Set cf = Server.CreateObject("Scripting.Dictionary")
  cfgOk = False
  On Error Resume Next
  Set rsC = cnnDB.Execute("SELECT * FROM tblMailConfig")
  If Err.Number = 0 Then
    If Not rsC.EOF Then
      For Each fld In rsC.Fields
        cf(LCase(fld.Name)) = fld.Value
      Next
      cfgOk = True
    End If
    rsC.Close
  End If
  Err.Clear
  On Error Goto 0
  If Not cfgOk Then Call Fail("500 Internal Server Error", "Las tablas del correo entrante no estan instaladas (ejecute db/mail_inbound_sqlserver.sql)")

  tokenDb = cf("api_token") & ""
  If Len(tokenDb) < 24 Or Not SafeEq(tokenIn, tokenDb) Then Call Fail("403 Forbidden", "Token no valido")

  ' Si algo falla, se responde con un JSON de error (el correo queda sin marcar y se reintenta);
  ' tras 3 errores del mismo correo se deja de reintentar.
  On Error Resume Next
  Select Case action
    Case "config"
      Call ApiConfig()
    Case "ingest"
      Call ApiIngest()
    Case "outbox"
      Call ApiOutbox()
    Case "outbox_done"
      Call ApiOutboxDone()
    Case "status"
      Call ApiStatus()
    Case Else
      Call Fail("400 Bad Request", "Accion desconocida")
  End Select
  If Err.Number <> 0 Then
    gErr = Err.Description & " (" & Err.Number & ")"
    Err.Clear
    If action = "ingest" Then Call LogMail(gMsgId, gFrom, gSubj, "error", 0, gErr)
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.Write "{""ok"":false,""retry"":true,""error"":""" & JsonStr(gErr) & """}"
  End If
  On Error Goto 0

  cnnDB.Close
  Session.Abandon
  Response.End

  Sub ApiConfig()
    Dim enabled, s, enTxt
    enabled = (SafeCInt(cf("enabled")) <> 0 Or cf("enabled") = True)
    If enabled Then enTxt = "true" Else enTxt = "false"
    s = "{""ok"":true,""enabled"":" & enTxt & "," & _
      """imap"":{""host"":""" & JsonStr(cf("imap_host")) & """,""port"":" & SafeCLng(cf("imap_port")) & _
      ",""security"":""" & JsonStr(cf("imap_security")) & """,""user"":""" & JsonStr(cf("imap_user")) & _
      """,""pass"":""" & JsonStr(cf("imap_pass")) & """,""folder"":""" & JsonStr(cf("imap_folder")) & _
      """,""done_folder"":""" & JsonStr(cf("imap_done_folder")) & """}," & _
      """smtp"":{""host"":""" & JsonStr(cf("smtp_host")) & """,""port"":" & SafeCLng(cf("smtp_port")) & _
      ",""security"":""" & JsonStr(cf("smtp_security")) & """,""user"":""" & JsonStr(cf("smtp_user")) & _
      """,""pass"":""" & JsonStr(cf("smtp_pass")) & """,""from_addr"":""" & JsonStr(cf("smtp_from_addr")) & _
      """,""from_name"":""" & JsonStr(cf("smtp_from_name")) & """}}"
    Response.Write s
  End Sub

  Sub ApiOutbox()
    Dim rs, s, first
    s = "{""ok"":true,""items"":["
    first = True
    Dim hasReplyTo, rt
    hasReplyTo = True
    On Error Resume Next
    Set rs = cnnDB.Execute("SELECT TOP 20 id, to_addr, subject, body, in_reply_to, refs, reply_to FROM tblMailOutbox WHERE status=0 AND tries < 5 ORDER BY id ASC")
    If Err.Number <> 0 Then
      Err.Clear
      hasReplyTo = False
      Set rs = cnnDB.Execute("SELECT TOP 20 id, to_addr, subject, body, in_reply_to, refs FROM tblMailOutbox WHERE status=0 AND tries < 5 ORDER BY id ASC")
    End If
    On Error Goto 0
    Do While Not rs.EOF
      If Not first Then s = s & ","
      first = False
      rt = ""
      If hasReplyTo Then rt = rs("reply_to") & ""
      s = s & "{""id"":" & rs("id") & ",""to"":""" & JsonStr(rs("to_addr")) & """,""subject"":""" & JsonStr(rs("subject")) & _
        """,""body"":""" & JsonStr(rs("body")) & """,""in_reply_to"":""" & JsonStr(rs("in_reply_to")) & _
        """,""references"":""" & JsonStr(rs("refs")) & """,""reply_to"":""" & JsonStr(rt) & """}"
      rs.MoveNext
    Loop
    rs.Close
    Response.Write s & "]}"
  End Sub

  Sub ApiOutboxDone()
    Dim id, ok, msgid, errTxt
    id = SafeCLng(Request.QueryString("id"))
    ok = (Request.QueryString("ok") = "1")
    msgid = Left(Request.QueryString("msgid") & "", 250)
    errTxt = Left(Request.QueryString("err") & "", 480)
    If id <= 0 Then Call Fail("400 Bad Request", "id")
    If ok Then
      cnnDB.Execute "UPDATE tblMailOutbox SET status=1, sent=" & SQLDate(Now, lhdAddSQLDelim) & ", msg_id='" & Esc(msgid) & "', last_error=NULL WHERE id=" & id
    Else
      cnnDB.Execute "UPDATE tblMailOutbox SET tries=tries+1, last_error='" & Esc(errTxt) & "', status=CASE WHEN tries+1 >= 5 THEN 2 ELSE 0 END WHERE id=" & id
    End If
    Response.Write "{""ok"":true}"
  End Sub

  Sub ApiStatus()
    Dim bytes, txt, n
    n = Request.TotalBytes
    txt = ""
    If n > 0 And n < 4000 Then
      bytes = Request.BinaryRead(n)
      txt = BinToText(bytes)
    End If
    cnnDB.Execute "UPDATE tblMailConfig SET last_run=" & SQLDate(Now, lhdAddSQLDelim) & ", last_status='" & Esc(Left(txt, 480)) & "'"
    Response.Write "{""ok"":true}"
  End Sub

  ' Bytes UTF-8 -> texto
  Function BinToText(bytes)
    Dim st
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write bytes
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    BinToText = st.ReadText
    st.Close
  End Function

  Sub LogMail(msgId, fromAddr, subj, result, pid, detail)
    On Error Resume Next
    cnnDB.Execute "INSERT INTO tblMailLog (message_id, from_addr, subject, result, problem_id, detail) VALUES ('" & _
      Esc(Left(msgId, 250)) & "', '" & Esc(Left(fromAddr, 250)) & "', '" & Esc(Left(subj, 250)) & "', '" & Esc(result) & "', " & _
      SafeCLng(pid) & ", '" & Esc(Left(detail, 480)) & "')"
    Err.Clear
    On Error Goto 0
  End Sub

  Sub QueueMail(toAddr, subj, body, inReplyTo, refs, pid, replyTo)
    On Error Resume Next
    cnnDB.Execute "INSERT INTO tblMailOutbox (to_addr, subject, body, in_reply_to, refs, problem_id, reply_to) VALUES ('" & _
      Esc(Left(toAddr, 250)) & "', '" & Esc(Left(subj, 250)) & "', '" & Esc(body) & "', '" & Esc(Left(inReplyTo, 250)) & "', '" & _
      Esc(Left(refs, 1400)) & "', " & SafeCLng(pid) & ", '" & Esc(Left(replyTo, 250)) & "')"
    If Err.Number <> 0 Then
      ' bases que aun no tienen la columna reply_to
      Err.Clear
      cnnDB.Execute "INSERT INTO tblMailOutbox (to_addr, subject, body, in_reply_to, refs, problem_id) VALUES ('" & _
        Esc(Left(toAddr, 250)) & "', '" & Esc(Left(subj, 250)) & "', '" & Esc(body) & "', '" & Esc(Left(inReplyTo, 250)) & "', '" & _
        Esc(Left(refs, 1400)) & "', " & SafeCLng(pid) & ")"
    End If
    On Error Goto 0
  End Sub

  ' Correo entrante: crea un caso nuevo o agrega una nota a un caso existente
  Sub ApiIngest()
    Dim n, bytes, stm, msg, msgId, fromRaw, fromAddr, fromName, subj, subjClean, bodyText, bodyHtml, replyTxt
    Dim ownAddr, refsHdr, inReplyTo, caseId, rs, rsU, userSid, userUid, userDept, userPhone, userLoc, userIsRep
    Dim resultTxt, problemId, detailTxt, attInfo, errCount, aliasObj

    n = Request.TotalBytes
    If n <= 0 Then Call Fail("400 Bad Request", "Sin contenido")
    If n > 60000000 Then Call Fail("413 Request Entity Too Large", "Correo demasiado grande")
    bytes = Request.BinaryRead(n)

    On Error Resume Next
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write bytes
    stm.Position = 0
    Set msg = Server.CreateObject("CDO.Message")
    msg.DataSource.OpenObject stm, "_Stream"
    If Err.Number <> 0 Then
      Err.Clear
      On Error Goto 0
      Call Fail("422 Unprocessable Entity", "No se pudo interpretar el correo")
    End If

    msgId = Hdr(msg, "message-id")
    fromRaw = msg.From & ""
    If Len(fromRaw) = 0 Then fromRaw = Hdr(msg, "from")
    fromAddr = ExtractAddr(fromRaw)
    fromName = ExtractName(fromRaw)
    subj = msg.Subject & ""
    inReplyTo = Hdr(msg, "in-reply-to")
    refsHdr = Hdr(msg, "references")
    ownAddr = LCase(Trim(cf("smtp_from_addr") & ""))
    If Len(msgId) = 0 Then msgId = "noid-" & Replace(Replace(CStr(Now), " ", "_"), "/", "-") & "-" & fromAddr
    On Error Goto 0
    gMsgId = msgId
    gFrom = fromAddr
    gSubj = subj

    ' Duplicados / reintentos
    Set rs = cnnDB.Execute("SELECT result FROM tblMailLog WHERE message_id='" & Esc(Left(msgId, 250)) & "'")
    errCount = 0
    Do While Not rs.EOF
      If rs("result") = "error" Then
        errCount = errCount + 1
      Else
        rs.Close
        Response.Write "{""ok"":true,""action"":""duplicate""}"
        Exit Sub
      End If
      rs.MoveNext
    Loop
    rs.Close
    If errCount >= 3 Then
      Call LogMail(msgId, fromAddr, subj, "failed", 0, "Se dejo de reintentar despues de 3 errores")
      Response.Write "{""ok"":true,""action"":""failed""}"
      Exit Sub
    End If

    ' Correos que no se atienden
    If IsAutoMail(msg, fromAddr, subj, ownAddr) Then
      Call LogMail(msgId, fromAddr, subj, "ignored", 0, "Correo automatico o de sistema")
      Response.Write "{""ok"":true,""action"":""ignored""}"
      Exit Sub
    End If

    ' Cuerpo del mensaje
    bodyText = ""
    bodyHtml = ""
    On Error Resume Next
    bodyText = msg.TextBody & ""
    bodyHtml = msg.HTMLBody & ""
    Err.Clear
    On Error Goto 0
    If Len(Trim(bodyText)) = 0 And Len(Trim(bodyHtml)) > 0 Then bodyText = HtmlToText(bodyHtml)
    bodyText = Trim(bodyText)

    ' Remitente: usuario registrado?
    userSid = 0 : userUid = "" : userDept = 0 : userPhone = "" : userLoc = "" : userIsRep = False
    Set rsU = cnnDB.Execute("SELECT TOP 1 sid, uid, department, phone, location1, IsRep FROM tblUsers WHERE LOWER(email1)='" & Esc(fromAddr) & "' ORDER BY sid")
    If Not rsU.EOF Then
      userSid = SafeCLng(rsU("sid"))
      userUid = rsU("uid") & ""
      userDept = SafeCLng(rsU("department"))
      userPhone = rsU("phone") & ""
      userLoc = rsU("location1") & ""
      userIsRep = (SafeCInt(rsU("IsRep")) <> 0)
    End If
    rsU.Close

    ' Es la respuesta a un caso existente?
    caseId = SubjectCaseId(subj)
    If caseId = 0 Then caseId = CaseFromHeaders(inReplyTo, refsHdr)

    If caseId > 0 Then
      Call AddNoteFromMail(msg, caseId, msgId, fromAddr, subj, bodyText, userSid, userUid, userIsRep)
    Else
      Set aliasObj = ResolveAlias(bytes)
      Call CreateTicketFromMail(msg, msgId, fromAddr, fromName, subj, bodyText, inReplyTo, refsHdr, userSid, userUid, userDept, userPhone, userLoc, aliasObj)
    End If
  End Sub

  ' Direccion de entrada (alias) por la que llego el correo: devuelve un diccionario con
  ' address, category, department, priority (vacio si ninguna direccion de destino esta registrada)
  Function ResolveAlias(bytes)
    Dim d, cands, parts, i, rs
    Set d = Server.CreateObject("Scripting.Dictionary")
    On Error Resume Next
    cands = RecipientCandidates(RawHeaderBlock(bytes))
    parts = Split(cands, vbLf)
    For i = 0 To UBound(parts)
      If Len(parts(i)) > 3 And i < 25 Then
        Set rs = cnnDB.Execute("SELECT TOP 1 address, category_id, def_department, def_priority FROM tblMailAlias WHERE enabled=1 AND LOWER(address)='" & Esc(parts(i)) & "'")
        If Err.Number <> 0 Then
          Err.Clear
          Exit For
        End If
        If Not rs.EOF Then
          d("address") = rs("address") & ""
          d("category") = SafeCLng(rs("category_id"))
          d("department") = SafeCLng(rs("def_department"))
          d("priority") = SafeCLng(rs("def_priority"))
          rs.Close
          Exit For
        End If
        rs.Close
      End If
    Next
    On Error Goto 0
    Set ResolveAlias = d
  End Function

  ' Busca en la cola de salida el caso al que responde el usuario (por Message-ID de nuestro acuse)
  Function CaseFromHeaders(inReplyTo, refsHdr)
    Dim ids, i, rs, parts, one
    CaseFromHeaders = 0
    ids = Trim(inReplyTo & " " & refsHdr)
    If Len(ids) = 0 Then Exit Function
    parts = Split(Replace(ids, vbTab, " "), " ")
    For i = 0 To UBound(parts)
      one = Trim(parts(i))
      If Len(one) > 3 And Left(one, 1) = "<" Then
        Set rs = cnnDB.Execute("SELECT TOP 1 problem_id FROM tblMailOutbox WHERE msg_id='" & Esc(one) & "' AND problem_id > 0")
        If Not rs.EOF Then CaseFromHeaders = SafeCLng(rs("problem_id"))
        rs.Close
        If CaseFromHeaders > 0 Then Exit Function
      End If
    Next
  End Function

  ' Guarda los adjuntos del correo en tblAttachments. Devuelve "registrados|omitidos" (listas separadas por "; ").
  Function SaveAttachments(msg, problemId, uidText)
    Dim i, a, st, b, nm, ext, sz, ctype, b64, cmd, allowed, maxB, okList, skipList, cdisp, p, tinyInline
    okList = "" : skipList = ""
    allowed = LhdAllowedExts(cnnDB)
    maxB = LhdMaxAttachBytes(cnnDB)
    On Error Resume Next
    For i = 1 To msg.Attachments.Count
      Set a = msg.Attachments.Item(i)
      nm = ""
      nm = a.FileName & ""
      cdisp = LCase(a.Fields("urn:schemas:mailheader:content-disposition").Value & "")
      Err.Clear
      Set st = a.GetDecodedContentStream
      st.Type = 1
      st.Position = 0
      sz = st.Size
      ' Imagenes pequenas incrustadas (firmas, logos) no se registran
      tinyInline = (InStr(cdisp, "inline") > 0 And LCase(Left(a.ContentMediaType & "", 6)) = "image/" And sz < 30000)
      If Not tinyInline Then
        If Len(nm) = 0 Then nm = "adjunto_" & i
        nm = Replace(Replace(Replace(Replace(nm, "\", "_"), "/", "_"), """", "_"), "|", "_")
        p = InStrRev(nm, ".")
        ext = ""
        If p > 0 Then ext = LCase(Mid(nm, p + 1))
        If Len(ext) = 0 Or LhdBlockedExt(ext) Or InStr("," & allowed & ",", "," & ext & ",") = 0 Then
          skipList = skipList & nm & " (tipo no permitido); "
        ElseIf sz > maxB Then
          skipList = skipList & nm & " (supera " & Round(maxB / 1024) & " KB); "
        ElseIf sz <= 0 Then
          skipList = skipList & nm & " (vacio); "
        Else
          b = st.Read
          b64 = BytesToB64(b)
          ctype = LhdAttachCtype(ext)
          Set cmd = Server.CreateObject("ADODB.Command")
          Set cmd.ActiveConnection = cnnDB
          cmd.CommandText = "INSERT INTO tblAttachments (problem_id, token, uid, fname, ctype, fsize, addDate, data) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
          cmd.Parameters.Append cmd.CreateParameter("pid", 3, 1, 0, CLng(problemId))
          cmd.Parameters.Append cmd.CreateParameter("token", 202, 1, 64, "mail")
          cmd.Parameters.Append cmd.CreateParameter("uid", 202, 1, 50, Left(uidText, 50))
          cmd.Parameters.Append cmd.CreateParameter("fname", 202, 1, 255, Left(nm, 255))
          cmd.Parameters.Append cmd.CreateParameter("ctype", 202, 1, 100, ctype)
          cmd.Parameters.Append cmd.CreateParameter("fsize", 3, 1, 0, CLng(sz))
          cmd.Parameters.Append cmd.CreateParameter("addDate", 7, 1, 0, Now())
          cmd.Parameters.Append cmd.CreateParameter("data", 203, 1, Len(b64), b64)
          cmd.Execute
          If Err.Number <> 0 Then
            skipList = skipList & nm & " (error al guardar); "
            Err.Clear
          Else
            okList = okList & nm & "; "
          End If
          Set cmd = Nothing
        End If
      End If
      Err.Clear
    Next
    On Error Goto 0
    SaveAttachments = okList & "|" & skipList
  End Function

  ' Caso nuevo
  Sub CreateTicketFromMail(msg, msgId, fromAddr, fromName, subj, bodyText, inReplyTo, refsHdr, userSid, userUid, userDept, userPhone, userLoc, aliasObj)
    Dim rsCat, rep, catName, deptId, prioId, statusId, uidTxt, id, title, descr, sqlIns, attRes, parts, okAtt, skipAtt
    Dim catId, aliasAddr, replyTo
    Dim policy, ackSubj, ackBody, nameShow, rsSt, stName, recent, rsRecent

    policy = SafeCInt(cf("unknown_policy"))
    If userSid = 0 And policy = 0 Then
      Call LogMail(msgId, fromAddr, subj, "ignored", 0, "Remitente no registrado (politica: ignorar)")
      Response.Write "{""ok"":true,""action"":""ignored""}"
      Exit Sub
    End If

    ' Categoria: la de la direccion por la que llego el correo; si no coincide ninguna, la de por omision
    catId = SafeCLng(cf("def_category"))
    aliasAddr = ""
    If aliasObj.Count > 0 Then
      aliasAddr = aliasObj("address")
      If SafeCLng(aliasObj("category")) > 0 Then catId = SafeCLng(aliasObj("category"))
    End If

    ' Categoria -> tecnico asignado
    Set rsCat = cnnDB.Execute("SELECT cname, rep_id FROM categories WHERE category_id=" & catId)
    If rsCat.EOF Then
      rsCat.Close
      Call LogMail(msgId, fromAddr, subj, "error", 0, "Falta configurar la categoria por omision")
      Response.Write "{""ok"":false,""retry"":true,""error"":""Falta configurar la categoria por omision""}"
      Exit Sub
    End If
    rep = SafeCLng(rsCat("rep_id"))
    catName = rsCat("cname") & ""
    rsCat.Close

    deptId = userDept
    If deptId <= 0 And aliasObj.Count > 0 Then deptId = SafeCLng(aliasObj("department"))
    If deptId <= 0 Then deptId = SafeCLng(cf("def_department"))
    prioId = SafeCLng(cf("def_priority"))
    If aliasObj.Count > 0 Then
      If SafeCLng(aliasObj("priority")) > 0 Then prioId = SafeCLng(aliasObj("priority"))
    End If
    statusId = SafeCLng(Cfg(cnnDB, "DefaultStatus"))

    If userSid > 0 Then
      uidTxt = userUid
    Else
      uidTxt = Left(fromAddr, 50)
    End If

    ' Limite de seguridad: 5 casos por hora por remitente (evita bucles con otros sistemas automaticos)
    Set rsRecent = cnnDB.Execute("SELECT COUNT(*) AS n FROM tblMailLog WHERE from_addr='" & Esc(fromAddr) & "' AND result='created' AND received >= " & SQLDate(DateAdd("h", -1, Now), lhdAddSQLDelim))
    recent = SafeCLng(rsRecent("n"))
    rsRecent.Close
    If recent >= 5 Then
      Call LogMail(msgId, fromAddr, subj, "ignored", 0, "Demasiados casos por correo en la ultima hora")
      Response.Write "{""ok"":true,""action"":""ignored""}"
      Exit Sub
    End If

    title = Left(CleanSubject(subj), 255)
    If Len(Trim(bodyText)) = 0 Then bodyText = "(sin contenido)"
    descr = TextToHtml(bodyText)
    If Len(fromName) > 0 Then
      descr = "<p><em>Correo recibido de " & Server.HTMLEncode(fromName) & " &lt;" & Server.HTMLEncode(fromAddr) & "&gt;</em></p>" & descr
    Else
      descr = "<p><em>Correo recibido de " & Server.HTMLEncode(fromAddr) & "</em></p>" & descr
    End If

    id = GetUnique(cnnDB, "problems")
    sqlIns = "INSERT INTO problems (id, uid, uemail, uphone, ulocation, entered_by, category, department, title, description, priority, status, start_date, rep, time_spent, kb) VALUES (" & _
      id & ", '" & Esc(uidTxt) & "', '" & Esc(Left(fromAddr, 50)) & "', '" & Esc(Left(userPhone, 50)) & "', '" & Esc(Left(userLoc, 50)) & "', " & _
      IIf2(userSid > 0, userSid, rep) & ", " & catId & ", " & deptId & ", '" & Esc(title) & "', '" & Esc(descr) & "', " & prioId & ", " & _
      statusId & ", " & SQLDate(Now, lhdAddSQLDelim) & ", " & rep & ", 0, 0)"
    cnnDB.Execute sqlIns

    attRes = SaveAttachments(msg, id, uidTxt)
    parts = Split(attRes, "|")
    okAtt = parts(0)
    skipAtt = ""
    If UBound(parts) >= 1 Then skipAtt = parts(1)

    ' Acuse al usuario (respuesta al mismo correo, con la etiqueta [Caso #n] en el asunto)
    stName = ""
    Set rsSt = cnnDB.Execute("SELECT sname FROM status WHERE status_id=" & statusId)
    If Not rsSt.EOF Then stName = rsSt("sname") & ""
    rsSt.Close
    nameShow = fromName
    If Len(nameShow) = 0 Then nameShow = fromAddr
    ackSubj = BuildAck(cf("ack_subject") & "", "Re: [title]", id, title, catName, stName, nameShow, okAtt, skipAtt, True)
    ackBody = BuildAck(cf("ack_body") & "", DefaultAckBody(), id, title, catName, stName, nameShow, okAtt, skipAtt, False)
    If InStr(ackSubj, "[Caso #" & id & "]") = 0 Then ackSubj = ackSubj & " [Caso #" & id & "]"
    replyTo = ""
    If Len(aliasAddr) > 0 And LCase(aliasAddr) <> LCase(Trim(cf("smtp_from_addr") & "")) Then replyTo = aliasAddr
    Call QueueMail(fromAddr, Left(ackSubj, 250), ackBody, msgId, Trim(refsHdr & " " & msgId), id, replyTo)

    Call LogMail(msgId, fromAddr, subj, "created", id, "Caso creado" & IIf2(Len(aliasAddr) > 0, " (via " & aliasAddr & ")", " (categoria por omision)") & IIf2(Len(okAtt) > 0, ". Adjuntos: " & okAtt, "") & IIf2(Len(skipAtt) > 0, ". Omitidos: " & skipAtt, ""))

    ' Aviso al tecnico asignado con las plantillas del sistema (despues de guardar todo: si la plantilla
    ' falta o el envio falla, el caso y el acuse ya quedaron registrados)
    On Error Resume Next
    If TplExists("repnew") Then Call eMessage(cnnDB, "repnew", id, Usr(cnnDB, rep, "email1"))
    If (SafeCLng(Cfg(cnnDB, "EnablePager")) > 0) Then
      If (SafeCLng(Cfg(cnnDB, "EnablePager")) <= prioId) And (Len(Usr(cnnDB, rep, "email2")) > 0) And TplExists("reppager") Then Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
    End If
    Err.Clear
    On Error Goto 0

    Response.Clear
    Response.Write "{""ok"":true,""action"":""created"",""problem_id"":" & id & "}"
  End Sub

  ' Existe la plantilla de correo del sistema? (eMessage termina la peticion si no existe)
  Function TplExists(tipo)
    Dim rs
    Set rs = cnnDB.Execute("SELECT subject FROM tblEmailMsg WHERE type='" & Esc(tipo) & "'")
    TplExists = Not rs.EOF
    rs.Close
  End Function

  Function DefaultAckBody()
    DefaultAckBody = "Hola [nombre]," & vbLf & vbLf & _
      "Hemos recibido tu solicitud y se registro como el caso #[problemid]: [title]" & vbLf & vbLf & _
      "Categoria: [category]" & vbLf & "Estado: [status]" & vbLf & "[adjuntos]" & vbLf & _
      "Para agregar informacion o archivos a este caso, responde a este correo (conserva [Caso #[problemid]] en el asunto)." & vbLf & vbLf & _
      "Gracias," & vbLf & "Mesa de servicio - [sitename]"
  End Function

  Function BuildAck(tpl, defTpl, id, title, catName, stName, nameShow, okAtt, skipAtt, isSubject)
    Dim t, adj
    t = tpl
    If Len(Trim(t)) = 0 Then t = defTpl
    adj = ""
    If Len(okAtt) > 0 Then adj = adj & "Adjuntos registrados: " & okAtt & vbLf
    If Len(skipAtt) > 0 Then adj = adj & "Adjuntos NO registrados: " & skipAtt & vbLf
    t = Replace(t, "[problemid]", id)
    t = Replace(t, "[title]", title)
    t = Replace(t, "[category]", catName)
    t = Replace(t, "[status]", stName)
    t = Replace(t, "[nombre]", nameShow)
    t = Replace(t, "[adjuntos]", adj)
    t = Replace(t, "[sitename]", Cfg(cnnDB, "SiteName") & "")
    t = Replace(t, "[baseurl]", Cfg(cnnDB, "BaseURL") & "")
    If isSubject Then t = Replace(Replace(t, vbCr, " "), vbLf, " ")
    BuildAck = t
  End Function

  ' Nota en un caso existente
  Sub AddNoteFromMail(msg, caseId, msgId, fromAddr, subj, bodyText, userSid, userUid, userIsRep)
    Dim rsP, pUid, pEmail, pStatus, pRep, allowed, noteText, authorUid, attRes, parts, okAtt, skipAtt, reopened, repEmail, defStatus

    Set rsP = cnnDB.Execute("SELECT uid, uemail, status, rep FROM problems WHERE id=" & caseId)
    If rsP.EOF Then
      rsP.Close
      Call LogMail(msgId, fromAddr, subj, "ignored", 0, "El caso #" & caseId & " no existe")
      Response.Write "{""ok"":true,""action"":""ignored""}"
      Exit Sub
    End If
    pUid = rsP("uid") & ""
    pEmail = LCase(Trim(rsP("uemail") & ""))
    pStatus = SafeCLng(rsP("status"))
    pRep = SafeCLng(rsP("rep"))
    rsP.Close

    ' Solo el solicitante, el usuario dueno del caso o un tecnico pueden agregar notas
    allowed = False
    If LCase(fromAddr) = pEmail Then allowed = True
    If userSid > 0 And LCase(userUid) = LCase(pUid) Then allowed = True
    If userIsRep Then allowed = True
    If Not allowed Then
      Call LogMail(msgId, fromAddr, subj, "ignored", caseId, "Remitente no autorizado para el caso #" & caseId)
      Response.Write "{""ok"":true,""action"":""ignored""}"
      Exit Sub
    End If

    noteText = StripQuoted(bodyText)
    If Len(Trim(noteText)) = 0 Then noteText = "(correo sin texto)"
    If Len(noteText) > 7000 Then noteText = Left(noteText, 7000) & vbLf & "[...]"
    authorUid = pUid
    If userIsRep Then authorUid = userUid

    attRes = SaveAttachments(msg, caseId, authorUid)
    parts = Split(attRes, "|")
    okAtt = parts(0)
    skipAtt = ""
    If UBound(parts) >= 1 Then skipAtt = parts(1)
    If Len(okAtt) > 0 Then noteText = noteText & vbLf & vbLf & "[Adjuntos por correo: " & okAtt & "]"
    If Len(skipAtt) > 0 Then noteText = noteText & vbLf & "[Adjuntos no registrados: " & skipAtt & "]"

    cnnDB.Execute "INSERT INTO tblNotes (id, [note], addDate, uid, private) VALUES (" & caseId & ", '" & Esc(noteText) & "', " & _
      SQLDate(Now, lhdAddSQLDelim) & ", '" & Esc(Left(authorUid, 50)) & "', 0)"

    ' Un caso cerrado se reabre cuando el usuario responde (configurable)
    reopened = False
    If pStatus >= 100 And SafeCInt(cf("reopen_closed")) <> 0 And Not userIsRep Then
      defStatus = SafeCLng(Cfg(cnnDB, "DefaultStatus"))
      cnnDB.Execute "UPDATE problems SET status=" & defStatus & ", close_date=NULL WHERE id=" & caseId
      cnnDB.Execute "INSERT INTO tblNotes (id, [note], addDate, uid, private) VALUES (" & caseId & ", 'Caso reabierto: el usuario respondio por correo.', " & _
        SQLDate(Now, lhdAddSQLDelim) & ", '" & Esc(Left(authorUid, 50)) & "', 1)"
      reopened = True
    End If

    Call LogMail(msgId, fromAddr, subj, "note", caseId, "Nota agregada al caso #" & caseId & IIf2(reopened, " (reabierto)", "") & IIf2(Len(okAtt) > 0, ". Adjuntos: " & okAtt, ""))

    On Error Resume Next
    repEmail = Usr(cnnDB, pRep, "email1")
    If Not userIsRep Then
      If TplExists("repupdate") Then Call eMessage(cnnDB, "repupdate", caseId, repEmail)
    ElseIf SafeCInt(Cfg(cnnDB, "Notifyuser")) = 1 Then
      If TplExists("userupdate") Then Call eMessage(cnnDB, "userupdate", caseId, pEmail)
    End If
    Err.Clear
    On Error Goto 0

    Response.Clear
    Response.Write "{""ok"":true,""action"":""note"",""problem_id"":" & caseId & "}"
  End Sub
%>
