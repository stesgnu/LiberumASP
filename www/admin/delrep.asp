<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: delrep.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Removed the selected support rep from viewrep.asp
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "RemoveSupportReps")%></title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .dr-card { max-width: 560px; margin: 8px auto 0; text-align: center; padding: 34px 28px; }
      .dr-icon { width: 68px; height: 68px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 32px; margin-bottom: 14px; }
      .dr-icon.ok { background: #DCFCE7; color: #15803D; }
      .dr-icon.err { background: #FEE2E2; color: #B91C1C; }
      .dr-title { font-size: 22px; font-weight: 700; margin: 0 0 8px; color: #1E293B; }
      .dr-text { font-size: 15px; color: #475569; margin: 0 0 18px; line-height: 1.5; }
      .dr-actions { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim sidList, blnActiveProblem, strActiveRepName, blnActiveCategory
      sidList = Split(Request.Form("rsid"), ",")
      blnActiveProblem = False
      blnActiveCategory = False

      Dim rep, probRes
      For Each Rep in sidList
        Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & Rep & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
        If Not probRes.EOF Then
          blnActiveProblem = True
          strActiveRepName = Usr(cnnDB, Rep, "uid")
        End If
        probRes.Close
      Next

      Dim rstProb
      If Not blnActiveProblem Then
        For Each Rep in sidList
          Set rstProb = SQLQuery(cnnDB, "SELECT category_id FROM categories WHERE rep_id = " & Rep)
          If Not rstProb.EOF Then
            blnActiveCategory = True
            strActiveRepName = Usr(cnnDB, Rep, "uid")
          End If
        Next
        rstProb.Close
      End If

      Dim oldRep, OldProbRes, repRes, rstEnteredBy
      If Not (blnActiveProblem Or blnActiveCategory) Then
        For Each oldRep In sidList
          Set OldProbRes = SQLQuery(cnnDB, "UPDATE problems SET rep=0 WHERE rep=" & oldRep)
          Set repRes = SQLQuery(cnnDB, "UPDATE tblUsers SET IsRep=0 WHERE sid=" & oldRep)
        Next
      End If
    %>
    <%
      Call DisplayHeader(cnnDB, sid)
    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128101;</span>
          <%=lang(cnnDB, "SupportRepresentatives")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <a class="btn btn-secondary" href="viewrep.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "SupportReps")%></a>
        </div>
      </div>

      <div class="lhd-card dr-card">
        <% If blnActiveProblem Or blnActiveCategory Then %>
          <div class="dr-icon err">&#9940;</div>
          <h2 class="dr-title"><%=lang(cnnDB, "UnableToDelete")%></h2>
          <p class="dr-text">
            <% If blnActiveProblem Then %>
              <%=lang(cnnDB, "Theuser")%>&nbsp;'<% = Server.HTMLEncode(strActiveRepName & "") %>'&nbsp;
              <%=lang(cnnDB, "DelrepOpenProblemsText")%>
            <% Else %>
              <%=lang(cnnDB, "Theuser")%>&nbsp;'<% = Server.HTMLEncode(strActiveRepName & "") %>'&nbsp;
              <%=lang(cnnDB, "DelrepCategoryAssignedText")%>
            <% End If %>
          </p>
        <% Else %>
          <div class="dr-icon ok">&#10004;</div>
          <h2 class="dr-title"><%=lang(cnnDB, "OperationComplete")%></h2>
          <p class="dr-text"><%=lang(cnnDB, "Thesupportrephavebeenremoved")%></p>
        <% End If %>
        <div class="dr-actions">
          <a class="btn btn-secondary" href="viewrep.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "SupportReps")%></a>
        </div>
      </div>
    </div>

    <%

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
