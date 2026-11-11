# Correo entrante de la mesa de servicio ("correo a caso")

La aplicación lee un buzón de correo (IMAP), crea un caso por cada correo nuevo —con sus adjuntos— y responde al
remitente por SMTP confirmando el número de caso. Si el usuario responde ese correo, su respuesta se agrega como una
nota del mismo caso.

## Cómo funciona

```
 Buzón IMAP ──► tools\mailpoll.ps1 (tarea programada, cada 1-2 min) ──► mail/api.asp ──► caso / nota / adjuntos
                       ▲                                                       │
                       └──── envía por SMTP las respuestas en cola ◄───────────┘
```

* `mailpoll.ps1` solo hace la parte de red (IMAP y SMTP). Toda la lógica está en la aplicación (`mail/api.asp`).
* La configuración (servidores, usuario, reglas) se administra en **Administración → Correo entrante**.
* El script consulta esa configuración en cada ejecución; no guarda datos propios.

### Qué hace con cada correo

| Situación | Resultado |
|---|---|
| Correo nuevo (sin etiqueta de caso) | Crea un caso: asunto = título, cuerpo = descripción, adjuntos registrados. Asigna el técnico de la categoría por omisión y le avisa con la plantilla `repnew`. Pone en cola un correo de confirmación para el remitente. |
| Correo nuevo enviado a una **dirección de entrada** registrada (p. ej. `redes@`) | La categoría (y, si se indicaron, el departamento y la prioridad) salen de esa dirección. Si no coincide ninguna, se usa la categoría por omisión. |
| Respuesta con `[Caso #123]` en el asunto, o respuesta al correo de confirmación | Agrega una nota al caso (sin el texto citado), registra los adjuntos y avisa al técnico. Si el caso estaba cerrado y la opción está activa, lo reabre. |
| Respuesta automática, rebote, lista de distribución, `no-reply`, el propio buzón | Se ignora y queda en el registro. |
| Remitente no autorizado a escribir en ese caso | Se ignora y queda en el registro. |
| Mismo correo recibido dos veces | Se detecta por `Message-ID` y no se duplica. |

* Al remitente se le reconoce por el correo registrado en su usuario. Si no existe, el caso se crea igualmente con el
  correo como identificador (se puede cambiar a "Ignorar").
* **Adjuntos**: se registran los que cumplen los tipos y el tamaño permitidos en *Configuración del sitio*; los demás se
  omiten y se informan en el correo de confirmación. Las imágenes pequeñas incrustadas (firmas) no se registran.
* **Límite de seguridad**: máximo 5 casos por hora por remitente (evita bucles con otros sistemas automáticos).
* El correo de confirmación sale con las cabeceras `In-Reply-To`/`References` y `Auto-Submitted`, y su asunto lleva la
  etiqueta `[Caso #n]`.

## Direcciones de entrada por categoría

Todas las direcciones (`redes@`, `soporte@`, `nomina@`...) deben entregar al **mismo buzón** configurado (alias, lista de
distribución o regla de reenvío del proveedor de correo). En *Administración → Correo entrante → Direcciones de entrada
por categoría* se registra cada dirección con su categoría. Al llegar un correo se revisa a qué dirección fue enviado
(cabeceras `X-Original-To`, `Delivered-To`, `Envelope-To`, `To`, `Cc` y el `for <...>` de `Received`) y se asigna la
categoría correspondiente. No hay límite de direcciones: el script lee un único buzón. La confirmación al usuario sale
con `Reply-To` apuntando a la dirección por la que entró.

## Instalación

1. **Base de datos**: ejecute `db\mail_inbound_sqlserver.sql` sobre la base del Help Desk (crea `tblMailConfig`,
   `tblMailLog` y `tblMailOutbox`). El usuario de la aplicación debe poder leer y escribir en ellas.
2. **IIS** (una sola vez): la API recibe el correo completo en una petición. Suba el límite de ASP, por ejemplo a 50 MB
   (ajuste el nombre del sitio):

   ```
   %windir%\system32\inetsrv\appcmd set config "Default Web Site/Helpdesk_JS" -section:system.webServer/asp /limits.maxRequestEntityAllowed:52428800 /commit:apphost
   ```
3. **Script**: copie `tools\mailpoll.ps1` al servidor (por ejemplo `C:\Liberum\tools\`). Necesita Windows PowerShell 5.1.
4. **Aplicación**: en *Administración → Correo entrante* complete IMAP, SMTP, categoría/departamento/prioridad por
   omisión, genere un **token** y guarde. Active "Crear casos a partir de los correos recibidos".
5. **Prueba**: desde el servidor ejecute

   ```
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Liberum\tools\mailpoll.ps1 -Url https://servidor/Helpdesk_JS -Token SU_TOKEN -Test
   ```

   Debe mostrar `IMAP OK` y `SMTP OK`. Use `-SkipCertCheck` si el sitio tiene un certificado autofirmado.
6. **Tarea programada** (cada minuto, con una cuenta que pueda leer el script):

   ```
   schtasks /Create /TN "Liberum Correo entrante" /SC MINUTE /MO 1 /RU SYSTEM ^
     /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Liberum\tools\mailpoll.ps1 -Url https://servidor/Helpdesk_JS -Token SU_TOKEN -LogFile C:\Liberum\tools\mailpoll.log"
   ```

   El script evita ejecuciones simultáneas. La pantalla de administración muestra la última ejecución, las respuestas
   pendientes de enviar y los últimos correos procesados.

## Parámetros de `mailpoll.ps1`

| Parámetro | Descripción |
|---|---|
| `-Url` | Dirección base de la aplicación. |
| `-Token` | Token de la API (pantalla de administración). |
| `-MaxPerRun` | Máximo de correos por ejecución (25 por omisión). |
| `-SkipCertCheck` | No validar el certificado HTTPS de la aplicación ni de los servidores de correo. |
| `-Test` | Solo prueba la conexión IMAP y SMTP. |
| `-LogFile` | Archivo de registro. |

## Limitaciones y notas

* **Autenticación por contraseña** (IMAP `LOGIN` y SMTP `AUTH PLAIN/LOGIN`). Con Gmail/Google Workspace use una
  *contraseña de aplicación*. Microsoft 365 desactivó la contraseña simple en IMAP: requiere OAuth, que este script no
  implementa (habría que usar un buzón en otro proveedor o un conector de Microsoft Graph).
* Los nombres de carpeta IMAP con tildes no están soportados (use nombres sin acentos).
* Las contraseñas del buzón se guardan en `tblMailConfig` igual que el resto de la configuración de la aplicación.
  Limite el acceso a la base y a la pantalla de administración.
* La API `mail/api.asp` solo acepta POST con el token en la cabecera `X-Mail-Token`.
* El aviso al técnico asignado usa las plantillas y el envío de correo ya configurados en la aplicación
  (*Configuración del sitio*); la confirmación al usuario sale por el SMTP de esta pantalla.
