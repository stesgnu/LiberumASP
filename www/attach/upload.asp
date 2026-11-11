<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  Response.Buffer = TRUE
  Response.Expires = 0
  Response.CodePage = 65001
  Response.CharSet = "utf-8"
%>
<!-- #include file = "../public.asp" -->
<!-- #include file = "../uploadClass.asp" -->
<%
  ' Subida de adjuntos e imagenes del editor. Devuelve JSON.
  ' POST ?token=...            -> sube un archivo (multipart, campo "file")
  ' POST ?action=delete&id=N   -> borra un adjunto propio que todavia no esta ligado a un ticket

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

  Sub JsonFail(msg)
    Response.Write "{""ok"":false,""error"":""" & JsonStr(msg) & """}"
  End Sub

  Sub FailStep(stepName)
    Dim d
    d = Err.Description
    Err.Clear
    Response.Clear
    Call JsonFail("Error al " & stepName & ": " & d)
    Response.End
  End Sub

  Dim cnnDB, sid, uidCur, action, appRoot, scriptName
  Set cnnDB = CreateCon
  sid = GetSid
  uidCur = ""
  If sid > 0 Then uidCur = Usr(cnnDB, sid, "uid") & ""
  Response.ContentType = "application/json"

  scriptName = Request.ServerVariables("SCRIPT_NAME")
  appRoot = Left(scriptName, InStr(LCase(scriptName), "/attach/") - 1)

  If uidCur = "" Then
    Call JsonFail("Sesion no valida. Vuelva a iniciar sesion.")
    cnnDB.Close
    Response.End
  End If

  action = LCase(Request.QueryString("action"))

  If action = "delete" Then
    Dim delId
    delId = SafeCInt(Request.QueryString("id"))
    If delId > 0 Then
      Call SQLQuery(cnnDB, "DELETE FROM tblAttachments WHERE id=" & delId & _
        " AND uid='" & Replace(uidCur, "'", "''") & "' AND problem_id=0")
      Response.Write "{""ok"":true}"
    Else
      Call JsonFail("Adjunto no valido.")
    End If
    cnnDB.Close
    Response.End
  End If

  ' ---- Subida ----
  Dim token, maxBytes, objFO, objProps, objFile, fname, ext, ctype, rst, newId, nidRes, allowedList
  token = Request.QueryString("token")
  If Len(token) = 0 Or Len(token) > 64 Then
    Call JsonFail("Sesion de formulario no valida.")
    cnnDB.Close
    Response.End
  End If
  Dim tokenOk, ci, ch
  tokenOk = True
  For ci = 1 To Len(token)
    ch = Mid(token, ci, 1)
    If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_" Or ch = "-") Then tokenOk = False
  Next
  If Not tokenOk Then
    Call JsonFail("Sesion de formulario no valida.")
    cnnDB.Close
    Response.End
  End If

  allowedList = LhdAllowedExts(cnnDB)
  maxBytes = SafeCLng(Cfg(cnnDB, "MaxImageSize"))
  If maxBytes <= 0 Then maxBytes = 2000000

  Set objFO = New FileUpload
  Set objProps = objFO.GetUploadSettings
  With objProps
    .Extensions = Split(allowedList, ",")
    .UploadDirectory = Server.MapPath(appRoot & "/image/")
    .AllowOverWrite = True
    .MaximumFileSize = maxBytes
    .MininumFileSize = 1
    .UploadDisabled = False
  End With
  Set objProps = Nothing

  objFO.ProcessUpload
  Set objFile = objFO.File("file")
  If objFile Is Nothing Then Set objFile = objFO.File(1)

  If objFile Is Nothing Then
    Call JsonFail("No se recibio ningun archivo.")
    cnnDB.Close
    Response.End
  End If
  If objFile.ErrorMessage <> "" Then
    Call JsonFail(objFile.ErrorMessage)
    cnnDB.Close
    Response.End
  End If

  fname = objFile.FileName
  ext = LCase(objFile.FileExtension)
  ctype = LhdAttachCtype(ext)
  If LhdBlockedExt(ext) Or InStr("," & allowedList & ",", "," & ext & ",") = 0 Then
    Call JsonFail("Tipo de archivo no permitido: ." & ext & ". Tipos aceptados: " & allowedList)
    cnnDB.Close
    Response.End
  End If

  ' Guardar el archivo en la base como texto base64 (cualquier tipo de archivo).
  ' Cada paso informa su error para poder diagnosticar la falla.
  Dim stepName, b64, cmd
  On Error Resume Next
  stepName = "codificar el archivo en base64"
  b64 = objFile.SaveAsBase64EncodedStr
  If Err.Number <> 0 Or Len(b64 & "") = 0 Then Call FailStep(stepName)
  stepName = "guardar el archivo en la base de datos"
  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = cnnDB
  cmd.CommandText = "INSERT INTO tblAttachments (problem_id, token, uid, fname, ctype, fsize, addDate, data) " & _
    "VALUES (0, ?, ?, ?, ?, ?, ?, ?)"
  cmd.Parameters.Append cmd.CreateParameter("token", 202, 1, 64, token)
  cmd.Parameters.Append cmd.CreateParameter("uid", 202, 1, 50, uidCur)
  cmd.Parameters.Append cmd.CreateParameter("fname", 202, 1, 255, Left(fname, 255))
  cmd.Parameters.Append cmd.CreateParameter("ctype", 202, 1, 100, ctype)
  cmd.Parameters.Append cmd.CreateParameter("fsize", 3, 1, 0, CLng(objFile.ByteCount))
  cmd.Parameters.Append cmd.CreateParameter("addDate", 7, 1, 0, Now())
  cmd.Parameters.Append cmd.CreateParameter("data", 203, 1, Len(b64), b64)
  cmd.Execute
  If Err.Number <> 0 Then Call FailStep(stepName)
  Set cmd = Nothing
  Set rst = Nothing
  stepName = "leer el ID asignado"
  Set nidRes = SQLQuery(cnnDB, "SELECT MAX(id) AS nid FROM tblAttachments WHERE token='" & token & _
    "' AND uid='" & Replace(uidCur, "'", "''") & "' AND problem_id=0")
  If Err.Number <> 0 Then Call FailStep(stepName)
  newId = 0
  If Not nidRes.EOF Then newId = SafeCInt(nidRes("nid"))
  nidRes.Close
  On Error Goto 0

  If newId <= 0 Then
    Call JsonFail("No se pudo guardar el archivo en la base de datos (sin ID asignado).")
  Else
    Response.Write "{""ok"":true,""id"":" & newId & _
      ",""name"":""" & JsonStr(fname) & """" & _
      ",""size"":" & CLng(objFile.ByteCount) & _
      ",""image"":" & LCase(CStr(LhdIsImageCtype(ctype))) & _
      ",""url"":""" & JsonStr(appRoot & "/attach/file.asp?id=" & newId) & """}"
  End If

  Set objFile = Nothing
  Set objFO = Nothing
  cnnDB.Close
%>
