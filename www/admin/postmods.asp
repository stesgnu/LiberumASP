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

  Filename: postmods.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Adds or modifies an item.  Takes data from the form
  on MODIFY.ASP, checks to make sure valid before
  entering in the database.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    ' Columnas adicionales de categories (estado / ANS): se arman las partes del SQL segun el
    ' tipo de dato real de cada columna (solo las que existen y llegaron en el formulario).
    Dim catSet, catCols, catVals
    catSet = ""
    catCols = ""
    catVals = ""

    Function CatLiteral(t, raw, defSize, label)
      Dim v
      v = Trim(raw & "")
      Select Case t
        Case 11   ' booleano
          If v = "1" Or LCase(v) = "true" Or LCase(v) = "activo" Then
            CatLiteral = "1"
          Else
            CatLiteral = "0"
          End If
        Case 2, 3, 4, 5, 6, 14, 16, 17, 18, 19, 20, 21, 131   ' numericos
          If Len(v) = 0 Then
            CatLiteral = "NULL"
          ElseIf IsNumeric(v) Then
            CatLiteral = Replace(CStr(v), ",", ".")
          Else
            Call DisplayError(3, "El campo '" & label & "' debe ser num" & ChrW(233) & "rico.")
          End If
        Case Else   ' texto
          ' Unidad del ANS: solo M (minutos), H (horas) o D (dias)
          If (label = "Tipo de ANS" Or label = "Tipo de ANS inicial") And Len(v) > 0 Then
            v = UCase(v)
            If v <> "M" And v <> "H" And v <> "D" Then
              Call DisplayError(3, "El campo '" & label & "' solo admite M (minutos), H (horas) o D (d" & ChrW(237) & "as).")
            End If
          End If
          If Len(v) = 0 Then
            CatLiteral = "NULL"
          Else
            If defSize > 0 And defSize < 4000 Then v = Left(v, defSize)
            CatLiteral = "'" & Replace(v, "'", "''") & "'"
          End If
      End Select
    End Function

    Sub BuildCatExtras()
      Dim rsF, xi, fi, xNames, xLabels, lit, fname
      xNames  = Array("estado", "ANS", "TipoAns", "TipoIndicador", "Ans_Ini", "TipoAnsIni")
      xLabels = Array("Estado", "ANS", "Tipo de ANS", "Tipo de indicador", "ANS inicial", "Tipo de ANS inicial")
      Set rsF = SQLQuery(cnnDB, "SELECT * FROM categories WHERE 1=0")
      For xi = 0 To 5
        For fi = 0 To rsF.Fields.Count - 1
          If LCase(rsF.Fields(fi).Name) = LCase(xNames(xi)) Then
            If Request.Form("x_" & xNames(xi)).Count > 0 Then
              fname = rsF.Fields(fi).Name
              lit = CatLiteral(rsF.Fields(fi).Type, Request.Form("x_" & xNames(xi)), rsF.Fields(fi).DefinedSize, xLabels(xi))
              catSet = catSet & ", [" & fname & "]=" & lit
              catCols = catCols & ", [" & fname & "]"
              catVals = catVals & ", " & lit
            End If
          End If
        Next
      Next
      rsF.Close
    End Sub
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ModificationDone")%>
    </title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .pm-card { max-width: 560px; margin: 8px auto 0; text-align: center; padding: 34px 28px; }
      .pm-icon { width: 68px; height: 68px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 32px; margin-bottom: 14px; }
      .pm-icon.ok { background: #DCFCE7; color: #15803D; }
      .pm-icon.err { background: #FEE2E2; color: #B91C1C; }
      .pm-title { font-size: 22px; font-weight: 700; margin: 0 0 8px; color: #1E293B; }
      .pm-text { font-size: 15px; color: #475569; margin: 0; line-height: 1.5; }
    </style>
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      ' Get control values
      Dim mType, strLanguageName, numDataFields, data_id, intLangID
      numDataFields = Cint(Request.Form("numdatafields"))
      intLangID = Request.Form("mLangID")
      strLanguageName = Request.Form("mLanguage")
      mType = Cint(Request.Form("mtype"))
      data_id = Cint(Request.Form("data_id"))


      Dim data1, data2, data3, data4
      data1 = CleanInputText(Request.Form("data1"))
      data2 = CleanInputText(Request.Form("data2"))
      data3 = CleanInputText(Request.Form("data3"))
      data4 = CleanInputText(Request.Form("data4"))

      'Check required fields
      If (numDataFields >= 1) AND (Len(data1) = 0) Then
        Call DisplayError(1, lang(cnnDB, "Field") & " 1")
      End If
      If (numDataFields >= 2) AND (Len(data2) = 0) Then
        Call DisplayError(1, lang(cnnDB, "Field") & " 2")
      End If
      If (numDataFields >= 3) AND (Len(data3) = 0) Then
        Call DisplayError(1, lang(cnnDB, "Field") & " 3")
      End If
      If (numDataFields >= 4) AND (Len(data4) = 0) Then
        Call DisplayError(1, lang(cnnDB, "Field") & " 4")
      End If

      'Convert data fields to integers where needed
      Select Case mType
        Case 1	' rep
            data1 = Lcase(data1)
            data4 = Replace(data4, "'", "`")
        Case 2	' category
            data2 = Cint(data2)
        Case 3	' department
        Case 4	' priority
            data1 = Cint(data1)
        Case 5	' status
            data1 = Cint(data1)
        Case 6  ' Language
      End Select

      If mType = 2 Then Call BuildCatExtras()

      Dim sqlString, rstSQL, sqlString2, isNewCat
      isNewCat = (data_id = 0)
      ' If data_id = 0 then creating a new entry in the
      ' database, otherwise we are modifying an old one

      if data_id = 0 Then
        Select Case mType
          Case 1	' rep
            Dim repRes
            Set repRes = SQLQuery(cnnDB, "SELECT rep_id FROM reps WHERE ruid='" & data1 & "'")
            If (Not repRes.EOF) Then
              repRes.Close
              Call DisplayError(3, lang(cnnDB, "This username already exists")& ".")
            End If
            repRes.Close
            data_id = GetUnique(cnnDB, "reps")
            sqlString = "INSERT INTO reps (rep_id, ruid, rname, remail, rdescription) " & _
              "VALUES (" & data_id & ",'" & data1 & "','" & data2 & "','" & data3 & "','" & data4 & "')"

          Case 2	' category
            data_id = GetUnique(cnnDB, "categories")
            sqlString = "INSERT INTO categories (category_id, rep_id, cname" & catCols & ") " & _
              "VALUES (" & data_id & "," & data2 & ",'" & Replace(Left(Trim(data1), 50), "'", "''") & "'" & catVals & ")"

          Case 3	' department
            data_id = GetUnique(cnnDB, "departments")
            sqlString = "INSERT INTO departments (department_id, dname) " & _
              "VALUES (" & data_id & ",'" & data1 & "')"

          Case 4	'  priorities
            If data1 < 0 Then
              Call DisplayError(3, "Enter a positive priority number.")
            End If
            Dim priRes
            Set priRes = SQLQuery(cnnDB, "SELECT priority_id FROM priority WHERE priority_id=" & data1)
            If Not priRes.EOF Then
              priRes.Close
              Call DisplayError(3, "Enter a unique priority number.")
            End If
            priRes.Close
            sqlString = "INSERT INTO priority (priority_id, pname) " & _
              "VALUES (" & data1 & ",'" & data2 & "')"

          Case 5	' status
            If data1 < 0 Then
              Call DisplayError(3, lang(cnnDB, "Enterapositivestatusnumber") & ".")
            End If
            Dim statusRes
            Set statusRes = SQLQuery(cnnDB, "SELECT status_id FROM status WHERE status_id=" & data1)
            If Not statusRes.EOF Then
              statusRes.Close
              Call DisplayError(3, lang(cnnDB, "Enterauniquestatusnumber") & ".")
            End If
            statusRes.Close
            sqlString = "INSERT INTO status (status_id, sname) " & _
              "VALUES (" & data1 & ",'" & data2 & "')"
              
          Case 6  ' Language
            data_id = GetUnique(cnnDB, "lang")
            Dim rstLanguage
            Set rstLanguage = SQLQuery(cnnDB, "SELECT id FROM tblLanguage WHERE langname='" & data1 & "' AND localized='" & data2 & "'")
            If Not rstLanguage.EOF Then
              rstLanguage.Close
              Call DisplayError(3, lang(cnnDB, "Enterauniquelanguageandlocalizedname") & ".")
            End If
            rstLanguage.Close
            sqlString = "INSERT INTO tblLanguage (id, langname, localized) " & _
              "VALUES (" & data_id & ",'" & data1 & "','" & data2 & "')"
              
        End Select
      Else
          Select Case mType
            Case 1	' rep
              Set repRes = SQLQuery(cnnDB, "SELECT rep_id FROM reps WHERE ruid='" & data1 & "'")
              If (Not repRes.EOF) Then
                If Cint(repRes("rep_id")) <> data_id Then
                  repRes.Close
                  Call DisplayError(3,lang(cnnDB, "Thisusernamealreadyexists") & ".")
                End If
              End If
              repRes.Close
              sqlString = "UPDATE reps SET " & _
                "ruid='" & data1 & "', " & _
                "rname='" & data2 & "', " & _
                "remail='" & data3 & "', " & _
                "rdescription='" & data4 & "' " & _
                "WHERE rep_id=" & data_id
  
            Case 2	' category
              ' El nombre (cname) NO se modifica: se conserva para mantener correcto el historico.
              sqlString = "UPDATE categories SET " & _
                "rep_id=" & data2 & catSet & " " & _
                "WHERE category_id=" & data_id
  
            Case 3	' department
              sqlString = "UPDATE departments SET dname='" & data1 & "' " & _
                "WHERE department_id=" & data_id
  
            Case 4	' priority
              sqlString = "UPDATE priority SET pname='" & data2 & "' " & _
                "WHERE priority_id=" & data_id
  
            Case 5	' status
              sqlString = "UPDATE status SET sname='" & data2 & "' " & _
                "WHERE status_id=" & data_id
               
            Case 6	' Language
              sqlString = "UPDATE tblLanguage SET " & _
                "langname='" & data1 & "', " & _
                "localized='" & data2 & "' " & _
                "WHERE id=" & data_id
          End Select
        End If
      

      ' All data is present
      ' Write into database
      Dim saveOk
      saveOk = True
      Set rstSQL = SQLQuery(cnnDB, sqlString)

      ' Categorias: comprobar que el nombre y el representante realmente quedaron guardados
      If mType = 2 Then
        Dim rstChk, catSaved
        catSaved = False
        Set rstChk = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & data_id)
        If Not rstChk.EOF Then
          ' En una categoria nueva se compara tambien el nombre; al editar el nombre no cambia.
          catSaved = (SafeCInt(rstChk("rep_id")) = SafeCInt(data2))
          If isNewCat Then catSaved = catSaved And (Trim(rstChk("cname") & "") = Trim(Left(Trim(data1), 50)))
        End If
        rstChk.Close
        saveOk = catSaved
        If catSaved Then
          Call LhdToast("success", "Categor" & ChrW(237) & "a guardada correctamente.")
        Else
          Call LhdToast("error", "La base de datos no guard" & ChrW(243) & " el cambio de la categor" & ChrW(237) & "a. Revise restricciones, disparadores (triggers) o permisos de UPDATE sobre la tabla categories.")
        End If
      End If

      If mType = 6 and len(sqlString2) > 0 then
        Set rstSQL = SQLQuery(cnnDB, sqlString2)
      End If

    %>

    <%
      ' Todo el guardado/validacion termino (los DisplayError ya se resolvieron): se dibuja la cabecera
      Call DisplayHeader(cnnDB, sid)
    %>
    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128190;</span>
          <%=lang(cnnDB, "ModificationDone")%>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
          <%
            ' Get the correct URL to return to
            Select Case mType
              Case 1 %>
                <a class="btn btn-secondary" href="viewrep.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "SupportReps")%></a>
          <%		Case 2 %>
                <a class="btn btn-secondary" href="viewcat.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Categories")%></a>
          <%		Case 3 %>
                <a class="btn btn-secondary" href="viewdep.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Departments")%></a>
          <%		Case 4 %>
                <a class="btn btn-secondary" href="viewpri.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Priorities")%></a>
          <%		Case 5 %>
                <a class="btn btn-secondary" href="viewstatus.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Statuses")%></a>
          <%		Case 6 %>
                <a class="btn btn-secondary" href="viewlang.asp">&larr; <%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Languages")%></a>
          <%	End Select %>
        </div>
      </div>

      <div class="lhd-card pm-card">
        <% If saveOk Then %>
          <div class="pm-icon ok">&#10004;</div>
          <h2 class="pm-title"><%=lang(cnnDB, "OperationComplete")%></h2>
          <p class="pm-text"><%=lang(cnnDB, "Thedatabasehasbeenupdated")%>.</p>
        <% Else %>
          <div class="pm-icon err">&#9888;</div>
          <h2 class="pm-title">No se guard&oacute; el cambio</h2>
          <p class="pm-text">La base de datos no confirm&oacute; el cambio. Revise el aviso mostrado y vuelva a intentarlo.</p>
        <% End If %>
      </div>
    </div>
    
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
