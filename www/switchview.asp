<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  Response.Buffer = True
%>
<!-- #include file = "public.asp" -->
<%
  ' Alterna el menu de un tecnico entre la vista de Representante y la de Usuario.
  ' Uso: switchview.asp?mode=user  |  switchview.asp?mode=rep
  Dim cnnDB, sid, mode, isRepUser
  Set cnnDB = CreateCon
  sid = GetSid
  Call CheckUser(cnnDB, sid)

  isRepUser = (SafeCInt(Usr(cnnDB, sid, "IsRep")) = 1)
  mode = LCase(Trim(Request.QueryString("mode")))
  cnnDB.Close

  If isRepUser And mode = "rep" Then
    Session("lhd_ViewMode") = "rep"
    Response.Redirect "rep/"
  Else
    ' Cualquier otro caso (o un usuario sin rol de tecnico) va a la vista de usuario
    If isRepUser Then Session("lhd_ViewMode") = "user"
    Response.Redirect "user/"
  End If
%>
