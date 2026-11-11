<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  Response.Buffer = TRUE
%>
<!-- #include file = "../public.asp" -->
<%
  ' Sirve un adjunto o imagen guardada en tblAttachments: GET attach/file.asp?id=N
  ' Solo lo ven el autor, el tecnico asignado al ticket y los tecnicos.

  Dim cnnDB, sid, id, rst, uidCur, probId, ownerUid, probUid, probRep, ctype, fname, fsize, isRep, allowed, b64, safeName

  ' Decodifica base64 a bytes (MSXML: bin.base64)
  Function LhdBase64ToBytes(txt)
    Dim dom, node
    Set dom = Server.CreateObject("MSXML2.DOMDocument")
    Set node = dom.createElement("b64")
    node.dataType = "bin.base64"
    node.text = txt
    LhdBase64ToBytes = node.nodeTypedValue
  End Function
  Set cnnDB = CreateCon
  sid = GetSid
  If sid = 0 Then
    Response.Clear
    Response.Status = "403 Forbidden"
    cnnDB.Close
    Response.End
  End If

  uidCur = Usr(cnnDB, sid, "uid") & ""
  isRep = SafeCInt(Usr(cnnDB, sid, "IsRep"))
  id = SafeCInt(Request.QueryString("id"))

  Set rst = SQLQuery(cnnDB, "SELECT problem_id, uid, ctype, fname, fsize FROM tblAttachments WHERE id=" & id)
  If id <= 0 Or rst.EOF Then
    Response.Clear
    Response.Status = "404 Not Found"
    cnnDB.Close
    Response.End
  End If

  probId = SafeCInt(rst("problem_id"))
  ownerUid = rst("uid") & ""
  ctype = rst("ctype") & ""
  fname = rst("fname") & ""
  fsize = rst("fsize")
  rst.Close

  allowed = False
  If ownerUid = uidCur Then allowed = True
  If isRep = 1 Then allowed = True
  If Not allowed And probId > 0 Then
    Set rst = SQLQuery(cnnDB, "SELECT uid, rep FROM problems WHERE id=" & probId)
    If Not rst.EOF Then
      probUid = rst("uid") & ""
      probRep = SafeCInt(rst("rep"))
      If probUid = uidCur Or probRep = sid Then allowed = True
    End If
    rst.Close
  End If

  If Not allowed Then
    Response.Clear
    Response.Status = "403 Forbidden"
    cnnDB.Close
    Response.End
  End If

  ' Solo tipos permitidos (nunca HTML ni scripts)
  If ctype = "" Or LCase(Left(ctype, 9)) = "text/html" Then ctype = "application/octet-stream"

  safeName = Replace(Replace(fname, """", "'"), vbCr & vbLf, " ")

  Set rst = SQLQuery(cnnDB, "SELECT data FROM tblAttachments WHERE id=" & id)
  If Not rst.EOF Then
    b64 = rst("data") & ""
    Response.Clear
    Response.ContentType = ctype
    If LhdIsImageCtype(ctype) Then
      Response.AddHeader "Content-Disposition", "inline; filename=""" & safeName & """"
    Else
      Response.AddHeader "Content-Disposition", "attachment; filename=""" & safeName & """"
    End If
    If Len(b64) > 0 Then Response.BinaryWrite LhdBase64ToBytes(b64)
  End If
  rst.Close
  cnnDB.Close
%>
