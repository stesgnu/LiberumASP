<#
.SYNOPSIS
  Lee el buzon de la mesa de servicio por IMAP, entrega cada correo a la aplicacion (que crea el caso o agrega
  una nota) y envia por SMTP las respuestas pendientes (acuses de recibo).

.DESCRIPTION
  Pensado para ejecutarse cada 1-2 minutos desde el Programador de tareas de Windows. No guarda datos propios:
  la configuracion (IMAP, SMTP, reglas) se administra en la aplicacion (Administracion > Correo entrante) y el
  script la consulta en cada ejecucion con el token del buzon.

  Requisitos: Windows PowerShell 5.1 (incluido en Windows Server 2012 R2 en adelante). No necesita modulos.

.PARAMETER Url
  Direccion base de la aplicacion, por ejemplo https://mesa.empresa.com/Helpdesk_JS

.PARAMETER Token
  Token de la API de correo (Administracion > Correo entrante > Token).

.PARAMETER MaxPerRun
  Maximo de correos que se procesan en cada ejecucion (por omision 25).

.PARAMETER SkipCertCheck
  No validar el certificado HTTPS de la aplicacion (solo para certificados autofirmados).

.PARAMETER Test
  Solo comprueba la conexion IMAP y SMTP con los datos configurados y muestra el resultado.

.PARAMETER LogFile
  Archivo donde se guarda el registro de cada ejecucion.

.EXAMPLE
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\tools\mailpoll.ps1 -Url https://mesa.empresa.com/Helpdesk_JS -Token ABC123 -LogFile C:\tools\mailpoll.log
#>
param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$Token,
  [int]$MaxPerRun = 25,
  [switch]$SkipCertCheck,
  [switch]$Test,
  [string]$LogFile
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Enum]::ToObject([Net.SecurityProtocolType], 3072 -bor 768 -bor 192)   # TLS 1.2 / 1.1 / 1.0
if ($SkipCertCheck) { [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }

$script:summary = New-Object System.Collections.ArrayList
$script:tagN = 0

function Write-Log([string]$msg) {
  $line = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $msg
  Write-Host $line
  if ($LogFile) { try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch { } }
  [void]$script:summary.Add($msg)
}

# API de la aplicacion
function Invoke-Api([string]$Action, [byte[]]$Body = $null, [hashtable]$Query = @{}) {
  $q = "action=$Action"
  foreach ($k in $Query.Keys) { $q += '&' + $k + '=' + [uri]::EscapeDataString([string]$Query[$k]) }
  $req = [System.Net.HttpWebRequest]::Create(($Url.TrimEnd('/') + '/mail/api.asp?' + $q))
  $req.Method = 'POST'
  $req.Headers.Add('X-Mail-Token', $Token)
  $req.ContentType = 'application/octet-stream'
  $req.Timeout = 180000
  $req.ReadWriteTimeout = 180000
  if ($null -eq $Body) { $Body = New-Object byte[] 0 }
  $req.ContentLength = $Body.Length
  $rs = $req.GetRequestStream()
  $rs.Write($Body, 0, $Body.Length)
  $rs.Close()
  $status = 0
  try {
    $resp = $req.GetResponse()
  } catch [System.Net.WebException] {
    $resp = $_.Exception.Response
    if (-not $resp) { throw }
  }
  $status = [int]$resp.StatusCode
  $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
  $txt = $sr.ReadToEnd()
  $sr.Close(); $resp.Close()
  try {
    $obj = $txt | ConvertFrom-Json
  } catch {
    $short = $txt; if ($short.Length -gt 200) { $short = $short.Substring(0, 200) }
    throw "Respuesta no valida de la aplicacion (HTTP $status): $short"
  }
  $obj | Add-Member -NotePropertyName http -NotePropertyValue $status -Force
  return $obj
}

# Conexion (TCP / TLS)
function Open-Connection([string]$HostName, [int]$Port, [string]$Security) {
  $tcp = New-Object System.Net.Sockets.TcpClient
  $tcp.ReceiveTimeout = 60000
  $tcp.SendTimeout = 60000
  $tcp.Connect($HostName, $Port)
  $stream = $tcp.GetStream()
  $stream.ReadTimeout = 60000
  $stream.WriteTimeout = 60000
  if ($Security -eq 'ssl') { $stream = Start-Tls $stream $HostName }
  return @{ Tcp = $tcp; Stream = $stream; Host = $HostName }
}

function Start-Tls($stream, [string]$HostName) {
  if ($SkipCertCheck) {
    $ssl = New-Object System.Net.Security.SslStream($stream, $false, ([System.Net.Security.RemoteCertificateValidationCallback]{ param($s, $c, $ch, $e) $true }))
  } else {
    $ssl = New-Object System.Net.Security.SslStream($stream, $false)
  }
  $ssl.AuthenticateAsClient($HostName)
  $ssl.ReadTimeout = 60000
  $ssl.WriteTimeout = 60000
  return $ssl
}

function Send-Text($stream, [string]$text) {
  $b = [Text.Encoding]::UTF8.GetBytes($text)
  $stream.Write($b, 0, $b.Length)
  $stream.Flush()
}

function Read-Line($stream) {
  $ms = New-Object IO.MemoryStream
  while ($true) {
    $b = $stream.ReadByte()
    if ($b -lt 0) { throw 'Conexion cerrada por el servidor' }
    if ($b -eq 10) { break }
    if ($b -ne 13) { $ms.WriteByte([byte]$b) }
  }
  return [Text.Encoding]::UTF8.GetString($ms.ToArray())
}

function Read-Exact($stream, [int]$n) {
  $buf = New-Object byte[] $n
  $off = 0
  while ($off -lt $n) {
    $r = $stream.Read($buf, $off, $n - $off)
    if ($r -le 0) { throw 'Conexion cerrada por el servidor' }
    $off += $r
  }
  return , $buf
}

# IMAP
function Quote-Imap([string]$s) { return '"' + ($s -replace '\\', '\\' -replace '"', '\"') + '"' }

function Invoke-Imap($conn, [string]$cmd) {
  $script:tagN++
  $tag = 'A{0:D4}' -f $script:tagN
  Send-Text $conn.Stream ("$tag $cmd`r`n")
  $lines = New-Object System.Collections.ArrayList
  $lits = New-Object System.Collections.ArrayList
  $final = ''
  while ($true) {
    $line = Read-Line $conn.Stream
    if ($line -match '^\*.*\{(\d+)\}$') {
      $bytes = Read-Exact $conn.Stream ([int]$Matches[1])
      [void]$lits.Add($bytes)
      [void]$lines.Add($line)
      continue
    }
    if ($line.StartsWith("$tag ")) { $final = $line; break }
    [void]$lines.Add($line)
  }
  $ok = $final -match "^$tag OK"
  return @{ Ok = $ok; Final = $final; Lines = $lines; Literals = $lits }
}

function Connect-Imap($c) {
  $security = ($c.security + '').ToLower()
  if (-not $security) { $security = 'ssl' }
  $conn = Open-Connection $c.host $c.port $security
  [void](Read-Line $conn.Stream)    # saludo
  if ($security -eq 'starttls') {
    $r = Invoke-Imap $conn 'STARTTLS'
    if (-not $r.Ok) { throw "El servidor IMAP rechazo STARTTLS: $($r.Final)" }
    $conn.Stream = Start-Tls $conn.Stream $c.host
  }
  $r = Invoke-Imap $conn ('LOGIN ' + (Quote-Imap $c.user) + ' ' + (Quote-Imap $c.pass))
  if (-not $r.Ok) { throw "No se pudo iniciar sesion en IMAP: $($r.Final)" }
  $folder = $c.folder; if (-not $folder) { $folder = 'INBOX' }
  $r = Invoke-Imap $conn ('SELECT ' + (Quote-Imap $folder))
  if (-not $r.Ok) { throw "No se pudo abrir la carpeta '$folder': $($r.Final)" }
  return $conn
}

function Close-Conn($conn) {
  try { $conn.Stream.Close() } catch { }
  try { $conn.Tcp.Close() } catch { }
}

function Read-Unseen($conn) {
  $r = Invoke-Imap $conn 'UID SEARCH UNSEEN'
  if (-not $r.Ok) { throw "Fallo la busqueda de correos nuevos: $($r.Final)" }
  $uids = @()
  foreach ($l in $r.Lines) {
    if ($l -match '^\* SEARCH\s*(.*)$') { $uids += ($Matches[1].Trim() -split '\s+' | Where-Object { $_ -match '^\d+$' }) }
  }
  return $uids
}

function Process-Mailbox($imapCfg) {
  $conn = Connect-Imap $imapCfg
  $processed = 0; $failed = 0
  try {
    $uids = @(Read-Unseen $conn)
    Write-Log ("IMAP: {0} correo(s) nuevo(s)" -f $uids.Count)
    $expunge = $false
    foreach ($uid in ($uids | Select-Object -First $MaxPerRun)) {
      try {
        $f = Invoke-Imap $conn "UID FETCH $uid (BODY.PEEK[])"
        if (-not $f.Ok -or $f.Literals.Count -eq 0) { throw "No se pudo leer el correo UID $uid" }
        $raw = [byte[]]$f.Literals[0]
        $markDone = $false
        try {
          $res = Invoke-Api 'ingest' $raw
          if ($res.ok) {
            $markDone = $true
            $extra = ''
            if ($res.problem_id) { $extra = " (caso #$($res.problem_id))" }
            Write-Log ("  UID {0}: {1}{2}" -f $uid, $res.action, $extra)
          } else {
            $failed++
            Write-Log ("  UID {0}: la aplicacion no pudo procesarlo ({1}); se reintentara" -f $uid, $res.error)
          }
        } catch [System.Net.WebException] {
          # correo demasiado grande para la aplicacion: no tiene sentido reintentar cada minuto
          if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 413) {
            $markDone = $true
            Write-Log "  UID ${uid}: demasiado grande para la aplicacion (HTTP 413); se marca como leido. Aumente 'maxRequestEntityAllowed' de ASP en IIS."
          } else { throw }
        }
        if ($markDone) {
          [void](Invoke-Imap $conn "UID STORE $uid +FLAGS (\Seen)")
          if ($imapCfg.done_folder) {
            $cp = Invoke-Imap $conn ("UID COPY $uid " + (Quote-Imap $imapCfg.done_folder))
            if (-not $cp.Ok) {
              [void](Invoke-Imap $conn ('CREATE ' + (Quote-Imap $imapCfg.done_folder)))
              $cp = Invoke-Imap $conn ("UID COPY $uid " + (Quote-Imap $imapCfg.done_folder))
            }
            if ($cp.Ok) { [void](Invoke-Imap $conn "UID STORE $uid +FLAGS (\Deleted)"); $expunge = $true }
          }
          $processed++
        }
      } catch {
        $failed++
        Write-Log ("  UID {0}: error: {1}" -f $uid, $_.Exception.Message)
      }
    }
    if ($expunge) { [void](Invoke-Imap $conn 'EXPUNGE') }
    [void](Invoke-Imap $conn 'LOGOUT')
  } finally { Close-Conn $conn }
  return @{ Processed = $processed; Failed = $failed }
}

# SMTP
function Read-SmtpReply($stream) {
  $code = 0; $text = New-Object System.Collections.ArrayList
  while ($true) {
    $l = Read-Line $stream
    [void]$text.Add($l)
    if ($l -match '^(\d{3})([ -])') {
      $code = [int]$Matches[1]
      if ($Matches[2] -eq ' ') { break }
    } else { break }
  }
  return @{ Code = $code; Text = $text }
}

function Invoke-Smtp($conn, [string]$cmd, [int[]]$expect) {
  Send-Text $conn.Stream ($cmd + "`r`n")
  $r = Read-SmtpReply $conn.Stream
  if ($expect -notcontains $r.Code) { throw "SMTP: se esperaba $($expect -join '/') y se recibio '$($r.Text -join ' | ')' al enviar '$($cmd.Split(' ')[0])'" }
  return $r
}

function Encode-Header([string]$s) {
  if ($s -match '^[\x20-\x7E]*$') { return $s }
  $parts = New-Object System.Collections.ArrayList
  $i = 0
  while ($i -lt $s.Length) {
    $n = [Math]::Min(24, $s.Length - $i)
    $chunk = $s.Substring($i, $n)
    [void]$parts.Add('=?UTF-8?B?' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($chunk)) + '?=')
    $i += $n
  }
  return ($parts -join "`r`n ")
}

function Build-Mime($smtpCfg, $item, [string]$messageId) {
  $nl = "`r`n"
  $date = [DateTimeOffset]::Now.ToString('ddd, dd MMM yyyy HH:mm:ss zzz', [Globalization.CultureInfo]::InvariantCulture)
  $date = $date -replace '([+-]\d\d):(\d\d)$', '$1$2'
  $fromName = $smtpCfg.from_name
  if ($fromName) { $from = (Encode-Header $fromName) + ' <' + $smtpCfg.from_addr + '>' } else { $from = '<' + $smtpCfg.from_addr + '>' }
  $body = ($item.body -replace "`r`n", "`n") -replace "`n", $nl
  $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($body))
  $b64lines = ($b64 -split '(.{1,76})' | Where-Object { $_ }) -join $nl
  $h = New-Object System.Collections.ArrayList
  [void]$h.Add("From: $from")
  [void]$h.Add("To: <$($item.to)>")
  [void]$h.Add('Subject: ' + (Encode-Header $item.subject))
  [void]$h.Add("Date: $date")
  [void]$h.Add("Message-ID: $messageId")
  if ($item.reply_to) { [void]$h.Add('Reply-To: <' + $item.reply_to + '>') }
  if ($item.in_reply_to) { [void]$h.Add('In-Reply-To: ' + $item.in_reply_to) }
  if ($item.references) { [void]$h.Add('References: ' + ($item.references -replace '\s+', ' ')) }
  [void]$h.Add('MIME-Version: 1.0')
  [void]$h.Add('Content-Type: text/plain; charset=utf-8')
  [void]$h.Add('Content-Transfer-Encoding: base64')
  [void]$h.Add('Auto-Submitted: auto-replied')
  [void]$h.Add('X-Auto-Response-Suppress: All')
  return (($h -join $nl) + $nl + $nl + $b64lines + $nl)
}

function Send-Mail($c, $item) {
  $security = ($c.security + '').ToLower()
  if (-not $security) { $security = 'none' }
  $conn = Open-Connection $c.host $c.port $security
  try {
    $r = Read-SmtpReply $conn.Stream
    if ($r.Code -ne 220) { throw "SMTP: saludo inesperado: $($r.Text -join ' | ')" }
    $ehloName = [Net.Dns]::GetHostName()
    $r = Invoke-Smtp $conn "EHLO $ehloName" @(250)
    if ($security -eq 'starttls') {
      [void](Invoke-Smtp $conn 'STARTTLS' @(220))
      $conn.Stream = Start-Tls $conn.Stream $c.host
      $r = Invoke-Smtp $conn "EHLO $ehloName" @(250)
    }
    if ($c.user) {
      $caps = ($r.Text -join ' ')
      if ($caps -match 'AUTH[ =].*PLAIN') {
        $cred = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("`0" + $c.user + "`0" + $c.pass))
        [void](Invoke-Smtp $conn "AUTH PLAIN $cred" @(235))
      } else {
        [void](Invoke-Smtp $conn 'AUTH LOGIN' @(334))
        [void](Invoke-Smtp $conn ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($c.user))) @(334))
        [void](Invoke-Smtp $conn ([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($c.pass))) @(235))
      }
    }
    $domain = ($c.from_addr -split '@')[-1]
    $messageId = '<' + [guid]::NewGuid().ToString('N') + '@' + $domain + '>'
    [void](Invoke-Smtp $conn ('MAIL FROM:<' + $c.from_addr + '>') @(250))
    [void](Invoke-Smtp $conn ('RCPT TO:<' + $item.to + '>') @(250, 251))
    [void](Invoke-Smtp $conn 'DATA' @(354))
    $mime = Build-Mime $c $item $messageId
    $mime = $mime -replace '(?m)^\.', '..'          # relleno de puntos (dot-stuffing)
    Send-Text $conn.Stream ($mime + ".`r`n")
    $r = Read-SmtpReply $conn.Stream
    if ($r.Code -ne 250) { throw "SMTP: el servidor no acepto el mensaje: $($r.Text -join ' | ')" }
    try { [void](Invoke-Smtp $conn 'QUIT' @(221)) } catch { }
    return $messageId
  } finally { Close-Conn $conn }
}

function Process-Outbox($smtpCfg) {
  $ob = Invoke-Api 'outbox'
  if (-not $ob.ok) { throw "No se pudo leer la cola de salida: $($ob.error)" }
  $items = @($ob.items)
  if ($items.Count -eq 0) { return @{ Sent = 0; Failed = 0 } }
  Write-Log ("SMTP: {0} correo(s) por enviar" -f $items.Count)
  $sent = 0; $failed = 0
  foreach ($it in $items) {
    try {
      $mid = Send-Mail $smtpCfg $it
      [void](Invoke-Api 'outbox_done' $null @{ id = $it.id; ok = 1; msgid = $mid })
      $sent++
      Write-Log ("  enviado a {0}: {1}" -f $it.to, $it.subject)
    } catch {
      $failed++
      $msg = $_.Exception.Message
      Write-Log ("  error al enviar a {0}: {1}" -f $it.to, $msg)
      try { [void](Invoke-Api 'outbox_done' $null @{ id = $it.id; ok = 0; err = $msg }) } catch { }
    }
  }
  return @{ Sent = $sent; Failed = $failed }
}

# Principal
$mutex = New-Object System.Threading.Mutex($false, 'Global\LiberumMailPoll')
if (-not $mutex.WaitOne(0)) { Write-Host 'Ya hay otra ejecucion en curso.'; exit 0 }
$exitCode = 0
try {
  try {
    $cfgResp = Invoke-Api 'config'
  } catch {
    Write-Log ("No se pudo consultar la aplicacion: " + $_.Exception.Message)
    exit 2
  }
  if (-not $cfgResp.ok) { Write-Log ("La aplicacion rechazo la consulta: " + $cfgResp.error); exit 2 }

  if ($Test) {
    try {
      $c = Connect-Imap $cfgResp.imap
      $n = @(Read-Unseen $c).Count
      [void](Invoke-Imap $c 'LOGOUT'); Close-Conn $c
      Write-Host "IMAP OK ($($cfgResp.imap.host):$($cfgResp.imap.port)). Correos sin leer: $n"
    } catch { Write-Host ("IMAP ERROR: " + $_.Exception.Message); $exitCode = 1 }
    try {
      $conn = Open-Connection $cfgResp.smtp.host $cfgResp.smtp.port ($cfgResp.smtp.security + '')
      $r = Read-SmtpReply $conn.Stream
      Close-Conn $conn
      if ($r.Code -eq 220) { Write-Host "SMTP OK ($($cfgResp.smtp.host):$($cfgResp.smtp.port))" } else { Write-Host ("SMTP respuesta inesperada: " + ($r.Text -join ' | ')); $exitCode = 1 }
    } catch { Write-Host ("SMTP ERROR: " + $_.Exception.Message); $exitCode = 1 }
    exit $exitCode
  }

  if (-not $cfgResp.enabled) { Write-Host 'El correo entrante esta deshabilitado en la aplicacion.'; exit 0 }

  try {
    $r = Process-Mailbox $cfgResp.imap
    if ($r.Processed -gt 0 -or $r.Failed -gt 0) { Write-Log ("IMAP: procesados {0}, con error {1}" -f $r.Processed, $r.Failed) }
  } catch {
    $exitCode = 1
    Write-Log ("IMAP ERROR: " + $_.Exception.Message)
  }

  try {
    [void](Process-Outbox $cfgResp.smtp)
  } catch {
    $exitCode = 1
    Write-Log ("SMTP ERROR: " + $_.Exception.Message)
  }

  # Resultado de la ejecucion para mostrarlo en la aplicacion
  try {
    $txt = if ($exitCode -eq 0) { 'OK' } else { 'CON ERRORES' }
    $detail = (($script:summary | Select-Object -Last 4) -join ' | ')
    [void](Invoke-Api 'status' ([Text.Encoding]::UTF8.GetBytes("$txt - $detail")))
  } catch { }
}
finally {
  $mutex.ReleaseMutex()
  $mutex.Dispose()
}
exit $exitCode
