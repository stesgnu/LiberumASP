# Cambios realizados al Help Desk (Liberum ASP)

Resumen de lo que se ha modificado sobre la base original. Las rutas son relativas a la carpeta `www` salvo que se indique otra cosa.

## 1. Interfaz y experiencia de uso

- **Vistas modernizadas** (tarjetas, botones de acción arriba, paginación): menú de usuario y de técnico, listados, detalle de caso, nuevo caso, tablero In/Out, registro, recuperar contraseña, KB, y todo el menú de administración.
- **Íconos**: las vistas usaban emojis que en algunos equipos salían como cuadros. `js/lhd-icons.js` los reemplaza por íconos SVG.
- **Listas desplegables con buscador** (`searchselect.js`): todas las listas se pueden filtrar escribiendo. Se omiten las múltiples, las de fechas y las marcadas con `data-nosearch`.
- **Notificaciones** con Toastify (`js/lhd-toast.js`, `LhdToast` en `public.asp`) al guardar, subir archivos, errores, etc.
- **Adjuntos**: zona de arrastrar y soltar (Dropzone, `js/lhd-editor.js`), cola de subida con estado por archivo y validación del tamaño máximo que configura la aplicación.
- **Paleta de colores**: se guarda en `tblTheme` y se aplica a todas las vistas, no solo a `/rep/` (`theme.asp`, `admin/colors.asp`).
- **Cambio de vista** usuario / técnico en el menú lateral (`switchview.asp`). Se recuperó la opción de casos cerrados en ambas vistas.
- **Inicio de sesión** (`logon.asp`) y **acceso de administración** (`admin/default.asp`): rediseño de dos paneles, campos con ícono, botón de mostrar contraseña, aviso de mayúsculas e indicador de carga. El acceso de administración bloquea 10 minutos tras 5 claves incorrectas desde la misma IP.

## 2. Informes y administración

- **Reportes**: los filtros por departamento, categoría y técnico funcionan juntos; filtros con listas con buscador.
- **Informe de cumplimiento de ANS** (`admin/reportans.asp`): tiempo de respuesta inicial y de solución por caso frente al ANS de su categoría, con resumen, filtros, paginación y exportación a CSV. Requiere `db/problems_first_response_sqlserver.sql`.
- **Categorías** (`admin/viewcat.asp`, `modify.asp`): se muestra y edita el técnico; el nombre ya no se puede cambiar (para no alterar el histórico); columnas nuevas editables: estado, ANS, tipo de ANS (M/H/D), tipo de indicador, ANS inicial y su tipo.
- **Usuarios y departamentos**: filtro y paginación de 10 por página; "Asignar a" muestra el nombre del técnico; en los datos de contacto se muestra el nombre y no el uid.
- Menú de administración dividido en **Mantenimiento** e **Informes**. Se quitó la tarjeta de tipos de adjuntos (ya estaba en la configuración del sitio).
- Pantallas modernizadas: `admin/test.asp`, `sysinfo.asp`, `adduser.asp`, `viewusers.asp`, `viewdep.asp`, entre otras.

## 3. Correo a caso (nuevo)

La aplicación lee un buzón por IMAP, crea un caso por cada correo nuevo (con adjuntos) y responde por SMTP con el número de caso. Si el usuario responde ese correo, su respuesta se agrega como nota.

| Pieza | Archivo |
|---|---|
| Lógica (API) | `mail/api.asp` |
| Script de red (IMAP/SMTP, tarea programada) | `../tools/mailpoll.ps1` |
| Pantalla de configuración | `admin/mailconfig.asp` |
| Tablas | `../db/mail_inbound_sqlserver.sql` |
| Guía de instalación | `../tools/LEEME_correo_entrante.md` |

- **Categoría por dirección de entrada**: todas las direcciones son alias de un solo buzón; la categoría (y opcionalmente departamento y prioridad) sale de la dirección a la que se escribió. Si no coincide ninguna se usa la categoría por omisión.
- Detecta duplicados por `Message-ID`, ignora respuestas automáticas y rebotes, limita a 5 casos por hora por remitente y reabre casos cerrados cuando el usuario responde.
- Limitaciones: solo autenticación por contraseña (Microsoft 365 exige OAuth); no se ha probado contra un servidor IMAP/SMTP real.

## 4. Seguridad

- Revisión general y correcciones: validación de permisos en `inout/status.asp` y `inout/savefile.asp`, destino seguro tras el login (solo rutas internas), HTML del editor filtrado con lista blanca, consultas con valores escapados.
- `web.config` endurecido: cabeceras de seguridad, archivos y carpetas internas no se sirven, métodos limitados a GET/POST/HEAD, errores detallados solo en local.
- Pendientes recomendados: contraseñas con hash, tokens CSRF y límite de intentos para usuarios normales.

## 5. Rendimiento

- Cachés por petición para configuración, usuario e idioma en `public.asp`.
- Índices en `db/indexes_sqlserver.sql` (usuarios, notas, cadenas de idioma, casos).

## 6. Servidor (`web.config`)

- URLs sin la extensión `.asp` (redirección de `/pagina.asp` a `/pagina` y reescritura interna).
- Si algo falla con la reescritura hay una variante `web.config.sin-rewrite`; la original quedó en `web.config.original`.
- IIS no admite `<!--` ni `-->` dentro de cadenas de VBScript; hay que partirlas (`"<" & "!--"`).

## 7. Idioma

- Cadenas nuevas en los archivos de idioma y `db/lang_SelectAnExistingUser.sql` para bases que ya tenían la tabla de cadenas. El idioma por omisión es Spanish (Español).

## 8. Scripts SQL (ejecutar en la base de datos)

| Script | Para qué |
|---|---|
| `db/tblTheme_sqlserver.sql` (o `_access.sql`) | Guardar la paleta de colores |
| `db/tblAttachments_sqlserver.sql`, `attachments_migrate_sqlserver.sql` | Adjuntos guardados en la base |
| `db/problems_first_response_sqlserver.sql` | Fecha de primera respuesta (informe de ANS) |
| `db/indexes_sqlserver.sql` | Índices de rendimiento |
| `db/mail_inbound_sqlserver.sql` | Correo a caso y alias de entrada |

## 9. Despliegue

- No sobrescribir `settings.asp` del servidor (tiene la conexión real a SQL Server).
- Desactivar la depuración de ASP en el servidor si aparecen ventanas de Visual Studio.
- Para el correo a caso: subir el límite de tamaño de petición de ASP y crear la tarea programada (ver la guía en `tools`).
