<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<%
  Option Explicit
  With Response
    .Buffer = True
    .Expires = 0
    .Clear
  End With
%>


<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: savefile.asp
  Date:     $Date: 2002/01/24 14:19:18 $
  Version:  $Revision: 1.51.2.1 $
  Purpose:  Save uploaded user image to disk.
  -->

  <!--  #include file = "../public.asp" -->
  <!--  #include file = "../uploadClass.asp" -->

  <%
    Dim cnnDB, sid, usid, uid
    Set cnnDB = CreateCon
    sid = GetSid
    usid = SafeCLng(Request.QueryString("uid"))
    uid = usr(cnnDB, usid, "uid")
  %>

  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta charset="utf-8">
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "Uploadimage")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

  <%
    ' See is user is validated
    Call CheckUser(cnnDB, sid)

    ' Solo el propio usuario, el administrador del tablero In/Out o el administrador del sitio
    ' pueden cambiar la foto de un usuario (igual que en inout/update.asp).
    If SafeCLng(sid) <> usid And SafeCInt(Usr(cnnDB, sid, "inoutadmin")) <> 1 And Session("lhd_IsAdmin") <> True Then
      cnnDB.Close
      Response.Clear
      Response.Status = "403 Forbidden"
      Response.Write "Acceso denegado."
      Response.End
    End If

    Call DisplayHeader(cnnDB, sid)

    Dim objFO, objProps, objFile, upOk, upMsg
    upOk = False
    upMsg = ""
    Set objFO = New FileUpload
    Set objProps = objFO.GetUploadSettings
    With objProps
      .Extensions = Array("jpg", "jpeg")
      .UploadDirectory = Server.Mappath("../image/")
      .AllowOverWrite = true
      .MaximumFileSize = cfg(cnnDB, "MaxImageSize")
      .MininumFileSize = 1000
      .UploadDisabled = false
    End With
    Set objProps = Nothing
    objFO.ProcessUpload
    Set objFile = objFO.File(1)
    If objFile.ErrorMessage <> "" Then
      upMsg = lang(cnnDB, "Anerroroccurreduploadingafile") & ": " & objFile.ErrorMessage
      Call LhdToast("error", upMsg)
    Else
      objFile.FileName = uid & ".jpg"
      objFile.SaveAsFile
      If objFile.UploadSuccessful Then
        upOk = True
        upMsg = lang(cnnDB, "fileuploadedsuccessfully")
        Call LhdToast("success", upMsg)
      Else
        upMsg = lang(cnnDB, "Anerroroccurredsavingfiletodisk") & ": " & objFile.ErrorMessage
        Call LhdToast("error", upMsg)
      End If
    End If
  %>

  <div class="lhd-wrapper">

    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
      <div class="lhd-page-title" style="margin:0"><span class="title-icon">&#128247;</span> <%=lang(cnnDB, "Uploadimage")%></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <a class="btn btn-secondary" href="details.asp?id=<% = usid %>">&larr; <%=lang(cnnDB, "Details")%></a>
      </div>
    </div>

    <div class="lhd-card" style="max-width:640px;margin:0 auto">
      <div class="section-header"><% = IIf2(upOk, "&#10004;", "&#9888;") %> <%=lang(cnnDB, "Uploadimage")%></div>
      <div class="alert <% = IIf2(upOk, "alert-success", "alert-error") %>">
        <span class="alert-icon"><% = IIf2(upOk, "&#10004;", "&#9888;") %></span>
        <span><% = Server.HTMLEncode(upMsg) %></span>
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap">
        <a class="btn btn-secondary" href="update.asp?id=<% = usid %>">&#9998; <%=lang(cnnDB, "Edit")%></a>
        <a class="btn btn-primary" href="details.asp?id=<% = usid %>"><%=lang(cnnDB, "Details")%></a>
      </div>
    </div>

  </div>

  <%
    Call DisplayFooter(cnnDB, sid)
    'clean up
    Set objFile = Nothing
    Set objFO = Nothing
    cnnDB.Close
  %>
  </body>
</html>
