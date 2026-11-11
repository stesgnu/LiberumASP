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

  Filename: reportans.asp
  Purpose:  Reporte de cumplimiento de ANS (acuerdos de nivel de servicio) por caso.
  -->

  <!-- 	#include file = "../public.asp" -->

  <%
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
    Call CheckAdmin

    ' Funciones auxiliares
    ' Minutos que equivale una unidad de ANS: M = minutos, H = horas, D = dias
    Function UnitMin(t)
      Select Case UCase(Trim(t & ""))
        Case "M"
          UnitMin = 1
        Case "H"
          UnitMin = 60
        Case "D"
          UnitMin = 1440
        Case Else
          UnitMin = 0
      End Select
    End Function

    ' Limite del ANS en minutos (0 = sin ANS definido)
    Function LimitMin(v, t)
      Dim n
      n = 0
      On Error Resume Next
      If Not IsNull(v) Then
        If Len(Trim(v & "")) > 0 Then n = CDbl(v & "")
      End If
      If Err.Number <> 0 Then
        n = 0
        Err.Clear
      End If
      On Error Goto 0
      If n <= 0 Then
        LimitMin = 0
      Else
        LimitMin = n * UnitMin(t)
      End If
    End Function

    ' Estado de un plazo: NA (sin ANS), OK (cumple), FAIL (no cumple o vencido), RUN (en curso, aun dentro del plazo)
    Function EvalState(limitM, startD, endD, hasEnd, ByRef elapsedM)
      elapsedM = -1
      If limitM <= 0 Then
        EvalState = "NA"
        Exit Function
      End If
      If hasEnd Then
        elapsedM = DateDiff("n", startD, endD)
        If elapsedM <= limitM Then
          EvalState = "OK"
        Else
          EvalState = "FAIL"
        End If
      Else
        elapsedM = DateDiff("n", startD, Now)
        If elapsedM > limitM Then
          EvalState = "FAIL"
        Else
          EvalState = "RUN"
        End If
      End If
    End Function

    ' Duracion legible a partir de minutos
    Function FmtDur(m)
      If m < 0 Then
        FmtDur = "&mdash;"
        Exit Function
      End If
      m = CLng(m)
      If m < 60 Then
        FmtDur = m & " min"
      ElseIf m < 1440 Then
        FmtDur = Int(m / 60) & " h " & (m Mod 60) & " min"
      Else
        FmtDur = Int(m / 1440) & " d " & Int((m Mod 1440) / 60) & " h"
      End If
    End Function

    ' Limite del ANS como texto (valor + unidad)
    Function FmtLimit(v, t)
      Dim u
      Select Case UCase(Trim(t & ""))
        Case "M"
          u = "min"
        Case "H"
          u = "h"
        Case "D"
          u = "d"
        Case Else
          u = ""
      End Select
      If LimitMin(v, t) <= 0 Then
        FmtLimit = "&mdash;"
      Else
        FmtLimit = Replace(CStr(CDbl(v & "")), ",", ".") & " " & u
      End If
    End Function

    Function Badge(st)
      Select Case st
        Case "OK"
          Badge = "<span class=""ans-badge ok"">&#10004; Cumple</span>"
        Case "FAIL"
          Badge = "<span class=""ans-badge bad"">&#10006; No cumple</span>"
        Case "RUN"
          Badge = "<span class=""ans-badge run"">&#128336; En curso</span>"
        Case Else
          Badge = "<span class=""ans-badge na"">N/A</span>"
      End Select
    End Function

    Function ParseIso(s, def)
      Dim d
      ParseIso = def
      s = Trim(s & "")
      If Len(s) = 10 Then
        If Mid(s, 5, 1) = "-" And Mid(s, 8, 1) = "-" Then
          If IsNumeric(Left(s, 4)) And IsNumeric(Mid(s, 6, 2)) And IsNumeric(Right(s, 2)) Then
            On Error Resume Next
            d = DateSerial(CInt(Left(s, 4)), CInt(Mid(s, 6, 2)), CInt(Right(s, 2)))
            If Err.Number = 0 Then ParseIso = d
            Err.Clear
            On Error Goto 0
          End If
        End If
      End If
    End Function

    Function IsoStr(d)
      IsoStr = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
    End Function

    Function CsvCell(v)
      CsvCell = """" & Replace(v & "", """", """""") & """"
    End Function

    ' Parametros y filtros
    Const PER_PAGE = 10
    Dim dFrom, dTo, fDep, fCat, fRep, fRes, pageNo, doExport
    dTo = ParseIso(Request.QueryString("to"), Date)
    dFrom = ParseIso(Request.QueryString("from"), DateAdd("d", -30, Date))
    If dFrom > dTo Then dFrom = dTo
    fDep = SafeCLng(Request.QueryString("dep"))
    fCat = SafeCLng(Request.QueryString("cat"))
    fRep = SafeCLng(Request.QueryString("rep"))
    fRes = LCase(Trim(Request.QueryString("res")))
    If fRes <> "ok" And fRes <> "fail" And fRes <> "run" Then fRes = "all"
    pageNo = SafeCLng(Request.QueryString("page"))
    doExport = (Request.QueryString("export") = "1")

    ' Las columnas de ANS deben existir en categories
    Dim ansReady, rsProbe, pi, pname, hasAns, hasTipo, hasIni, hasTipoIni
    hasAns = False : hasTipo = False : hasIni = False : hasTipoIni = False
    Set rsProbe = SQLQuery(cnnDB, "SELECT * FROM categories WHERE 1=0")
    For pi = 0 To rsProbe.Fields.Count - 1
      pname = LCase(rsProbe.Fields(pi).Name)
      If pname = "ans" Then hasAns = True
      If pname = "tipoans" Then hasTipo = True
      If pname = "ans_ini" Then hasIni = True
      If pname = "tipoansini" Then hasTipoIni = True
    Next
    rsProbe.Close
    ansReady = (hasAns And hasTipo And hasIni And hasTipoIni)

    ' La columna first_response (fecha de primera respuesta) puede no existir en bases antiguas.
    ' Si falta, la primera respuesta se estima con la primera nota de alguien distinto al solicitante.
    Dim hasFirstResp, rsProbe2, pj
    hasFirstResp = False
    Set rsProbe2 = SQLQuery(cnnDB, "SELECT * FROM problems WHERE 1=0")
    For pj = 0 To rsProbe2.Fields.Count - 1
      If LCase(rsProbe2.Fields(pj).Name) = "first_response" Then hasFirstResp = True
    Next
    rsProbe2.Close

    ' Calculo del cumplimiento
    Const NF = 14
    Dim allRows(), nRows, cap
    cap = 256
    nRows = 0
    ReDim allRows(NF - 1, cap - 1)
    Dim nOk, nFail, nRun, nSolOk, nSolFail, nIniOk, nIniFail
    nOk = 0 : nFail = 0 : nRun = 0 : nSolOk = 0 : nSolFail = 0 : nIniOk = 0 : nIniFail = 0

    If ansReady Then
      Dim sqlList, rsL, whereExtra
      whereExtra = ""
      If fDep > 0 Then whereExtra = whereExtra & " AND p.department = " & fDep
      If fCat > 0 Then whereExtra = whereExtra & " AND p.category = " & fCat
      If fRep > 0 Then whereExtra = whereExtra & " AND p.rep = " & fRep

      Dim firstRespSql
      If hasFirstResp Then
        firstRespSql = "p.first_response"
      Else
        firstRespSql = "(SELECT MIN(n.addDate) FROM tblNotes AS n WHERE n.id = p.id AND n.uid <> p.uid) AS first_response"
      End If
      sqlList = "SELECT p.id, p.title, p.start_date, p.close_date, " & firstRespSql & ", p.status, " & _
        "d.dname, c.cname, c.ANS, c.TipoAns, c.Ans_Ini, c.TipoAnsIni, r.fname, r.uid AS ruid " & _
        "FROM (((problems AS p INNER JOIN categories AS c ON p.category = c.category_id) " & _
        "LEFT JOIN departments AS d ON p.department = d.department_id) " & _
        "LEFT JOIN tblUsers AS r ON p.rep = r.sid) " & _
        "WHERE p.start_date >= " & SQLDate(dFrom, lhdAddSQLDelim) & _
        " AND p.start_date <= " & SQLDate(DateAdd("s", 86399, dTo), lhdAddSQLDelim) & _
        " AND (c.ANS > 0 OR c.Ans_Ini > 0)" & whereExtra & " ORDER BY p.start_date DESC"
      Set rsL = SQLQuery(cnnDB, sqlList)

      Dim stD, clD, frD, isClosed, limSol, limIni, elSol, elIni, solSt, iniSt, overall, repName
      Do While Not rsL.EOF
        If Not IsNull(rsL("start_date")) Then
          stD = rsL("start_date")
          isClosed = (SafeCLng(rsL("status")) >= 100)
          clD = rsL("close_date")
          frD = rsL("first_response")

          limSol = LimitMin(rsL("ANS"), rsL("TipoAns"))
          limIni = LimitMin(rsL("Ans_Ini"), rsL("TipoAnsIni"))

          ' Solucion: del inicio al cierre; si sigue abierto, hasta ahora
          If isClosed And Not IsNull(clD) Then
            solSt = EvalState(limSol, stD, clD, True, elSol)
          Else
            solSt = EvalState(limSol, stD, Now, False, elSol)
          End If

          ' Respuesta inicial: del inicio a la primera respuesta (o al cierre si cerro sin respuesta registrada)
          If Not IsNull(frD) Then
            iniSt = EvalState(limIni, stD, frD, True, elIni)
          ElseIf isClosed And Not IsNull(clD) Then
            iniSt = EvalState(limIni, stD, clD, True, elIni)
          Else
            iniSt = EvalState(limIni, stD, Now, False, elIni)
          End If

          If solSt = "FAIL" Or iniSt = "FAIL" Then
            overall = "FAIL"
          ElseIf solSt = "RUN" Or iniSt = "RUN" Then
            overall = "RUN"
          ElseIf solSt = "OK" Or iniSt = "OK" Then
            overall = "OK"
          Else
            overall = "NA"
          End If

          If overall <> "NA" Then
            ' Contadores (sobre todos los casos filtrados, sin importar el filtro de resultado)
            If overall = "OK" Then nOk = nOk + 1
            If overall = "FAIL" Then nFail = nFail + 1
            If overall = "RUN" Then nRun = nRun + 1
            If solSt = "OK" Then nSolOk = nSolOk + 1
            If solSt = "FAIL" Then nSolFail = nSolFail + 1
            If iniSt = "OK" Then nIniOk = nIniOk + 1
            If iniSt = "FAIL" Then nIniFail = nIniFail + 1

            If fRes = "all" Or (fRes = "ok" And overall = "OK") Or (fRes = "fail" And overall = "FAIL") Or (fRes = "run" And overall = "RUN") Then
              If nRows >= cap Then
                cap = cap * 2
                ReDim Preserve allRows(NF - 1, cap - 1)
              End If
              repName = Trim(rsL("fname") & "")
              If Len(repName) = 0 Then repName = Trim(rsL("ruid") & "")
              allRows(0, nRows) = rsL("id")
              allRows(1, nRows) = rsL("title") & ""
              allRows(2, nRows) = rsL("cname") & ""
              allRows(3, nRows) = rsL("dname") & ""
              allRows(4, nRows) = repName
              allRows(5, nRows) = stD
              If isClosed And Not IsNull(clD) Then allRows(6, nRows) = clD Else allRows(6, nRows) = Empty
              allRows(7, nRows) = FmtLimit(rsL("ANS"), rsL("TipoAns"))
              allRows(8, nRows) = elSol
              allRows(9, nRows) = solSt
              allRows(10, nRows) = FmtLimit(rsL("Ans_Ini"), rsL("TipoAnsIni"))
              allRows(11, nRows) = elIni
              allRows(12, nRows) = iniSt
              allRows(13, nRows) = overall
              nRows = nRows + 1
            End If
          End If
        End If
        rsL.MoveNext
      Loop
      rsL.Close
    End If

    ' Exportar a CSV (Excel): se responde antes de dibujar la pagina
    If doExport And ansReady Then
      Dim ri, csv, stTxt
      csv = ChrW(65279) & "ID;T" & ChrW(237) & "tulo;Categor" & ChrW(237) & "a;Departamento;Representante;Fecha inicio;Fecha cierre;" & _
        "ANS soluci" & ChrW(243) & "n;Tiempo soluci" & ChrW(243) & "n (min);Resultado soluci" & ChrW(243) & "n;" & _
        "ANS respuesta inicial;Tiempo respuesta (min);Resultado respuesta inicial;Resultado general" & vbCrLf
      For ri = 0 To nRows - 1
        csv = csv & allRows(0, ri) & ";" & CsvCell(allRows(1, ri)) & ";" & CsvCell(allRows(2, ri)) & ";" & CsvCell(allRows(3, ri)) & ";" & _
          CsvCell(allRows(4, ri)) & ";" & CsvCell(allRows(5, ri)) & ";"
        If IsEmpty(allRows(6, ri)) Then csv = csv & ";" Else csv = csv & CsvCell(allRows(6, ri)) & ";"
        csv = csv & CsvCell(Replace(allRows(7, ri), "&mdash;", "")) & ";" & allRows(8, ri) & ";"
        Select Case allRows(9, ri)
          Case "OK": stTxt = "Cumple"
          Case "FAIL": stTxt = "No cumple"
          Case "RUN": stTxt = "En curso"
          Case Else: stTxt = "N/A"
        End Select
        csv = csv & stTxt & ";" & CsvCell(Replace(allRows(10, ri), "&mdash;", "")) & ";" & allRows(11, ri) & ";"
        Select Case allRows(12, ri)
          Case "OK": stTxt = "Cumple"
          Case "FAIL": stTxt = "No cumple"
          Case "RUN": stTxt = "En curso"
          Case Else: stTxt = "N/A"
        End Select
        csv = csv & stTxt & ";"
        Select Case allRows(13, ri)
          Case "OK": stTxt = "Cumple"
          Case "FAIL": stTxt = "No cumple"
          Case Else: stTxt = "En curso"
        End Select
        csv = csv & stTxt & vbCrLf
      Next
      Response.Clear
      Response.ContentType = "text/csv"
      Response.AddHeader "Content-Disposition", "attachment; filename=reporte_ans_" & IsoStr(dFrom) & "_" & IsoStr(dTo) & ".csv"
      Response.Write csv
      cnnDB.Close
      Response.End
    End If

    ' Paginacion
    Dim totalPages, skipRows, firstShown, lastShown, nEval, pctOk
    totalPages = Int((nRows + PER_PAGE - 1) / PER_PAGE)
    If totalPages < 1 Then totalPages = 1
    If pageNo < 1 Then pageNo = 1
    If pageNo > totalPages Then pageNo = totalPages
    skipRows = (pageNo - 1) * PER_PAGE
    firstShown = 0
    If nRows > 0 Then firstShown = skipRows + 1
    lastShown = skipRows + PER_PAGE
    If lastShown > nRows Then lastShown = nRows

    nEval = nOk + nFail
    pctOk = 0
    If nEval > 0 Then pctOk = Round(nOk * 100 / nEval, 1)

    ' Parametros para conservar el filtro en los enlaces
    Dim qsBase
    qsBase = "from=" & IsoStr(dFrom) & "&amp;to=" & IsoStr(dTo)
    If fDep > 0 Then qsBase = qsBase & "&amp;dep=" & fDep
    If fCat > 0 Then qsBase = qsBase & "&amp;cat=" & fCat
    If fRep > 0 Then qsBase = qsBase & "&amp;rep=" & fRep
    If fRes <> "all" Then qsBase = qsBase & "&amp;res=" & fRes
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;Cumplimiento de ANS</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" href="../default.css">
    <style type="text/css">
      .ans-filters { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 0 14px; align-items: end; }
      .ans-kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin-bottom: 18px; }
      .ans-kpi { background: var(--color-surface, #fff); border: 1px solid var(--color-border-light, #E2E8F0); border-left: 4px solid #94A3B8; border-radius: 10px; padding: 14px 16px; box-shadow: var(--shadow-sm, 0 1px 3px rgba(0,0,0,.08)); }
      .ans-kpi.ok { border-left-color: var(--color-success, #1A7F45); }
      .ans-kpi.bad { border-left-color: var(--color-danger, #B02A37); }
      .ans-kpi.run { border-left-color: #D69E2E; }
      .ans-kpi.pct { border-left-color: var(--color-primary, #0057B8); }
      .ans-kpi .n { font-size: 26px; font-weight: 700; color: #1a202c; line-height: 1.1; }
      .ans-kpi .l { font-size: 12px; color: #64748b; text-transform: uppercase; letter-spacing: .04em; margin-top: 2px; }
      .ans-bar { height: 8px; border-radius: 9999px; background: #E2E8F0; margin-top: 8px; overflow: hidden; }
      .ans-bar > span { display: block; height: 100%; background: var(--color-success, #1A7F45); }
      .ans-badge { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; white-space: nowrap; }
      .ans-badge.ok { background: var(--color-success-bg, #D4EDDA); color: var(--color-success, #1A7F45); }
      .ans-badge.bad { background: var(--color-danger-bg, #F8D7DA); color: var(--color-danger, #B02A37); }
      .ans-badge.run { background: var(--color-warning-bg, #FFF3CD); color: var(--color-warning, #856404); }
      .ans-badge.na { background: #F1F5F9; color: #64748b; }
      .ans-sub { font-size: 12px; color: #64748b; margin-top: 3px; }
      .ans-note { font-size: 12px; color: #64748b; padding: 12px 16px; border-top: 1px solid var(--color-border-light, #E2E8F0); }
      .ans-pager { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; justify-content: space-between; padding: 14px 16px; border-top: 1px solid var(--color-border-light, #E2E8F0); }
      .ans-pages { display: flex; gap: 4px; flex-wrap: wrap; }
      .ans-pg { min-width: 34px; height: 34px; padding: 0 10px; display: inline-flex; align-items: center; justify-content: center; border: 1px solid var(--color-border, #CBD5E0); border-radius: 8px; background: #fff; color: #334155; font-size: 13px; font-weight: 600; text-decoration: none !important; }
      a.ans-pg:hover { border-color: var(--color-primary, #0057B8); color: var(--color-primary, #0057B8); }
      .ans-pg.on { background: var(--color-primary, #0057B8); border-color: var(--color-primary, #0057B8); color: #fff; }
      .ans-pg.off { color: #A0AEC0; background: #F7FAFC; cursor: default; }
      .ans-count { font-size: 13px; color: #64748b; }
      .ans-empty { padding: 36px 16px; text-align: center; color: #64748b; }
    </style>
  </head>
  <body>

    <%
      Call DisplayHeader(cnnDB, sid)
    %>

    <div class="lhd-wrapper">

      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:12px">
        <div class="lhd-page-title" style="margin:0">
          <span class="title-icon">&#128202;</span>
          Cumplimiento de ANS
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <a class="btn btn-secondary" href="reports.asp">&larr; <%=lang(cnnDB, "ReportsMenu")%></a>
          <a class="btn btn-secondary" href="default.asp">&larr; <%=lang(cnnDB, "AdministrativeMenu")%></a>
        </div>
      </div>

      <% If Not ansReady Then %>
        <div class="alert alert-error">
          <span class="alert-icon">&#9888;</span>
          <span>La tabla <strong>categories</strong> no tiene todas las columnas de ANS (<code>ANS</code>, <code>TipoAns</code>, <code>Ans_Ini</code>, <code>TipoAnsIni</code>). Agr&eacute;guelas para usar este reporte.</span>
        </div>
      <% Else %>

      <!-- Filtros -->
      <div class="lhd-card" style="margin-bottom:18px">
        <div class="section-header">&#128269; Filtros</div>
        <form method="get" action="reportans.asp">
          <div class="ans-filters">
            <div class="form-group">
              <label class="form-label" for="from">Desde</label>
              <input type="date" class="form-control" id="from" name="from" value="<% = IsoStr(dFrom) %>">
            </div>
            <div class="form-group">
              <label class="form-label" for="to">Hasta</label>
              <input type="date" class="form-control" id="to" name="to" value="<% = IsoStr(dTo) %>">
            </div>
            <div class="form-group">
              <label class="form-label" for="dep"><%=lang(cnnDB, "Department")%></label>
              <select class="form-control" id="dep" name="dep">
                <option value="0">Todos</option>
                <%
                  Dim rsD
                  Set rsD = SQLQuery(cnnDB, "SELECT department_id, dname FROM departments WHERE department_id > 0 ORDER BY dname ASC")
                  Do While Not rsD.EOF
                    If SafeCLng(rsD("department_id")) = fDep Then
                      Response.Write "<option value=""" & rsD("department_id") & """ selected>" & Server.HTMLEncode(rsD("dname") & "") & "</option>"
                    Else
                      Response.Write "<option value=""" & rsD("department_id") & """>" & Server.HTMLEncode(rsD("dname") & "") & "</option>"
                    End If
                    rsD.MoveNext
                  Loop
                  rsD.Close
                %>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label" for="cat"><%=lang(cnnDB, "Category")%></label>
              <select class="form-control" id="cat" name="cat">
                <option value="0">Todas</option>
                <%
                  Dim rsC
                  Set rsC = SQLQuery(cnnDB, "SELECT category_id, cname FROM categories WHERE category_id > 0 ORDER BY cname ASC")
                  Do While Not rsC.EOF
                    If SafeCLng(rsC("category_id")) = fCat Then
                      Response.Write "<option value=""" & rsC("category_id") & """ selected>" & Server.HTMLEncode(rsC("cname") & "") & "</option>"
                    Else
                      Response.Write "<option value=""" & rsC("category_id") & """>" & Server.HTMLEncode(rsC("cname") & "") & "</option>"
                    End If
                    rsC.MoveNext
                  Loop
                  rsC.Close
                %>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label" for="rep"><%=lang(cnnDB, "SupportRep")%></label>
              <select class="form-control" id="rep" name="rep">
                <option value="0">Todos</option>
                <%
                  Dim rsR, repLbl
                  Set rsR = SQLQuery(cnnDB, "SELECT sid, uid, fname FROM tblUsers WHERE sid > 0 AND IsRep <> 0 ORDER BY fname ASC, uid ASC")
                  Do While Not rsR.EOF
                    repLbl = rsR("uid") & ""
                    If Len(Trim(rsR("fname") & "")) > 0 Then repLbl = Trim(rsR("fname") & "") & " (" & rsR("uid") & ")"
                    If SafeCLng(rsR("sid")) = fRep Then
                      Response.Write "<option value=""" & rsR("sid") & """ selected>" & Server.HTMLEncode(repLbl) & "</option>"
                    Else
                      Response.Write "<option value=""" & rsR("sid") & """>" & Server.HTMLEncode(repLbl) & "</option>"
                    End If
                    rsR.MoveNext
                  Loop
                  rsR.Close
                %>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label" for="res">Resultado</label>
              <select class="form-control" id="res" name="res">
                <option value="all"<% If fRes = "all" Then Response.Write(" selected") %>>Todos</option>
                <option value="ok"<% If fRes = "ok" Then Response.Write(" selected") %>>Cumplen</option>
                <option value="fail"<% If fRes = "fail" Then Response.Write(" selected") %>>No cumplen</option>
                <option value="run"<% If fRes = "run" Then Response.Write(" selected") %>>En curso</option>
              </select>
            </div>
          </div>
          <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:6px">
            <button type="submit" class="btn btn-primary">&#128269; Ver reporte</button>
            <a class="btn btn-secondary" href="reportans.asp">Limpiar</a>
            <a class="btn btn-secondary" href="reportans.asp?<% = qsBase %>&amp;export=1">&#128228; Exportar CSV</a>
          </div>
        </form>
      </div>

      <!-- Resumen -->
      <div class="ans-kpis">
        <div class="ans-kpi"><div class="n"><% = nOk + nFail + nRun %></div><div class="l">Casos evaluados</div></div>
        <div class="ans-kpi ok"><div class="n"><% = nOk %></div><div class="l">Cumplen el ANS</div></div>
        <div class="ans-kpi bad"><div class="n"><% = nFail %></div><div class="l">No cumplen</div></div>
        <div class="ans-kpi run"><div class="n"><% = nRun %></div><div class="l">En curso (a tiempo)</div></div>
        <div class="ans-kpi pct">
          <div class="n"><% = Replace(CStr(pctOk), ",", ".") %>%</div>
          <div class="l">% de cumplimiento</div>
          <div class="ans-bar"><span style="width:<% = Replace(CStr(pctOk), ",", ".") %>%"></span></div>
        </div>
      </div>

      <div class="ans-kpis">
        <div class="ans-kpi"><div class="l" style="margin:0 0 6px">Soluci&oacute;n</div><div class="ans-sub" style="margin:0"><strong style="color:var(--color-success,#1A7F45)"><% = nSolOk %></strong> cumplen &middot; <strong style="color:var(--color-danger,#B02A37)"><% = nSolFail %></strong> no cumplen</div></div>
        <div class="ans-kpi"><div class="l" style="margin:0 0 6px">Respuesta inicial</div><div class="ans-sub" style="margin:0"><strong style="color:var(--color-success,#1A7F45)"><% = nIniOk %></strong> cumplen &middot; <strong style="color:var(--color-danger,#B02A37)"><% = nIniFail %></strong> no cumplen</div></div>
      </div>

      <!-- Detalle por caso -->
      <div class="lhd-card lhd-card-flush">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;padding:14px 16px;border-bottom:1px solid var(--color-border-light,#E2E8F0)">
          <strong>Detalle por caso</strong>
          <span class="ans-count">
            <% If nRows = 0 Then %>Sin resultados<% Else %>Mostrando <strong><% = firstShown %>&ndash;<% = lastShown %></strong> de <strong><% = nRows %></strong><% End If %>
          </span>
        </div>
        <div style="overflow-x:auto">
          <table class="lhd-table">
            <thead>
              <tr>
                <th class="col-id">ID</th>
                <th>Caso</th>
                <th>Representante</th>
                <th class="hide-mobile">Inicio</th>
                <th>Respuesta inicial</th>
                <th>Soluci&oacute;n</th>
                <th>Resultado</th>
              </tr>
            </thead>
            <tbody>
              <%
                Dim rw, lastRow
                lastRow = skipRows + PER_PAGE - 1
                If lastRow > nRows - 1 Then lastRow = nRows - 1
                For rw = skipRows To lastRow
              %>
              <tr>
                <td class="col-id"><a href="../rep/details.asp?id=<% = allRows(0, rw) %>">#<% = allRows(0, rw) %></a></td>
                <td>
                  <a href="../rep/details.asp?id=<% = allRows(0, rw) %>"><strong><% = Server.HTMLEncode(allRows(1, rw)) %></strong></a>
                  <div class="ans-sub"><% = Server.HTMLEncode(allRows(2, rw)) %><% If Len(allRows(3, rw)) > 0 Then Response.Write(" &middot; " & Server.HTMLEncode(allRows(3, rw))) %></div>
                </td>
                <td><% If Len(allRows(4, rw)) > 0 Then Response.Write(Server.HTMLEncode(allRows(4, rw))) Else Response.Write("&mdash;") End If %></td>
                <td class="hide-mobile"><% = DisplayDate(allRows(5, rw), lhdDateOnly) %></td>
                <td>
                  <% = Badge(allRows(12, rw)) %>
                  <div class="ans-sub">L&iacute;mite <% = allRows(10, rw) %> &middot; <% = FmtDur(allRows(11, rw)) %></div>
                </td>
                <td>
                  <% = Badge(allRows(9, rw)) %>
                  <div class="ans-sub">L&iacute;mite <% = allRows(7, rw) %> &middot; <% = FmtDur(allRows(8, rw)) %><% If IsEmpty(allRows(6, rw)) Then Response.Write(" (abierto)") %></div>
                </td>
                <td><% = Badge(allRows(13, rw)) %></td>
              </tr>
              <%
                Next
              %>
            </tbody>
          </table>
          <% If nRows = 0 Then %>
            <div class="ans-empty">No hay casos con ANS definido en el periodo y filtros elegidos.</div>
          <% End If %>
        </div>

        <% If totalPages > 1 Then
             Dim pgFrom, pgTo, pgI
             pgFrom = pageNo - 2
             If pgFrom < 1 Then pgFrom = 1
             pgTo = pgFrom + 4
             If pgTo > totalPages Then
               pgTo = totalPages
               pgFrom = pgTo - 4
               If pgFrom < 1 Then pgFrom = 1
             End If %>
        <div class="ans-pager">
          <span class="ans-count">P&aacute;gina <strong><% = pageNo %></strong> de <strong><% = totalPages %></strong></span>
          <div class="ans-pages">
            <% If pageNo > 1 Then %>
              <a class="ans-pg" href="reportans.asp?page=1&amp;<% = qsBase %>" title="Primera">&laquo;</a>
              <a class="ans-pg" href="reportans.asp?page=<% = pageNo - 1 %>&amp;<% = qsBase %>"><%=lang(cnnDB, "Previous")%></a>
            <% Else %>
              <span class="ans-pg off">&laquo;</span>
              <span class="ans-pg off"><%=lang(cnnDB, "Previous")%></span>
            <% End If %>
            <% For pgI = pgFrom To pgTo
                 If pgI = pageNo Then %>
              <span class="ans-pg on"><% = pgI %></span>
            <%   Else %>
              <a class="ans-pg" href="reportans.asp?page=<% = pgI %>&amp;<% = qsBase %>"><% = pgI %></a>
            <%   End If
               Next %>
            <% If pageNo < totalPages Then %>
              <a class="ans-pg" href="reportans.asp?page=<% = pageNo + 1 %>&amp;<% = qsBase %>"><%=lang(cnnDB, "Next")%></a>
              <a class="ans-pg" href="reportans.asp?page=<% = totalPages %>&amp;<% = qsBase %>" title="&Uacute;ltima">&raquo;</a>
            <% Else %>
              <span class="ans-pg off"><%=lang(cnnDB, "Next")%></span>
              <span class="ans-pg off">&raquo;</span>
            <% End If %>
          </div>
        </div>
        <% End If %>

        <div class="ans-note">
          <strong>C&oacute;mo se calcula:</strong> el tiempo se mide en horas y minutos corridos (calendario) desde la fecha de inicio del caso.
          <em>Soluci&oacute;n</em> compara contra la fecha de cierre (o contra la hora actual si el caso sigue abierto);
          <em>Respuesta inicial</em> compara contra la fecha de primera respuesta. Un caso abierto cuyo plazo ya pas&oacute; cuenta como <strong>No cumple</strong>.
          Solo se incluyen casos cuya categor&iacute;a tiene ANS definido.
          <% If Not hasFirstResp Then %>
            <br><strong>Nota:</strong> la tabla <code>problems</code> no tiene la columna <code>first_response</code>; la respuesta inicial se estima con la fecha de la primera nota de una persona distinta al solicitante.
          <% End If %>
        </div>
      </div>

      <% End If %>
    </div>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
