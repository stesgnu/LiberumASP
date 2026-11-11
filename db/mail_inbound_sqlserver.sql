/*
  Liberum Help Desk - correo entrante (SQL Server)

  Crea las tablas de la funcionalidad "correo a caso":
    tblMailConfig  configuracion del buzon IMAP / SMTP y reglas (una sola fila)
    tblMailLog     registro de cada correo procesado (tambien evita duplicados)
    tblMailOutbox  cola de respuestas / acuses que envia el script de apoyo

  - Es seguro ejecutarlo varias veces.
  - Ejecutar sobre la base de datos de la mesa de ayuda.
*/

IF OBJECT_ID(N'dbo.tblMailConfig', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tblMailConfig (
        id               INT           NOT NULL CONSTRAINT PK_tblMailConfig PRIMARY KEY,
        enabled          BIT           NOT NULL CONSTRAINT DF_tblMailConfig_enabled DEFAULT (0),
        imap_host        VARCHAR(150)  NULL,
        imap_port        INT           NULL,
        imap_security    VARCHAR(10)   NULL,      -- ssl | starttls | none
        imap_user        VARCHAR(150)  NULL,
        imap_pass        VARCHAR(200)  NULL,
        imap_folder      VARCHAR(100)  NULL,
        imap_done_folder VARCHAR(100)  NULL,      -- carpeta a donde se mueven los correos procesados (opcional)
        smtp_host        VARCHAR(150)  NULL,
        smtp_port        INT           NULL,
        smtp_security    VARCHAR(10)   NULL,      -- ssl | starttls | none
        smtp_user        VARCHAR(150)  NULL,
        smtp_pass        VARCHAR(200)  NULL,
        smtp_from_addr   VARCHAR(150)  NULL,
        smtp_from_name   VARCHAR(100)  NULL,
        def_category     INT           NULL,
        def_department   INT           NULL,
        def_priority     INT           NULL,
        unknown_policy   INT           NOT NULL CONSTRAINT DF_tblMailConfig_unknown DEFAULT (1),  -- 1 = aceptar, 0 = ignorar
        reopen_closed    BIT           NOT NULL CONSTRAINT DF_tblMailConfig_reopen DEFAULT (1),
        ack_subject      VARCHAR(200)  NULL,
        ack_body         VARCHAR(MAX)  NULL,
        api_token        VARCHAR(80)   NULL,
        last_run         DATETIME      NULL,
        last_status      VARCHAR(500)  NULL
    );
    INSERT INTO dbo.tblMailConfig (id, enabled, imap_port, imap_security, imap_folder, smtp_port, smtp_security, unknown_policy, reopen_closed)
    VALUES (1, 0, 993, 'ssl', 'INBOX', 587, 'starttls', 1, 1);
    PRINT 'Tabla tblMailConfig creada.';
END
GO

IF OBJECT_ID(N'dbo.tblMailLog', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tblMailLog (
        id          INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_tblMailLog PRIMARY KEY,
        received    DATETIME      NOT NULL CONSTRAINT DF_tblMailLog_received DEFAULT (GETDATE()),
        message_id  VARCHAR(255)  NULL,
        from_addr   VARCHAR(255)  NULL,
        subject     VARCHAR(255)  NULL,
        result      VARCHAR(20)   NULL,    -- created | note | ignored | error | failed
        problem_id  INT           NULL,
        detail      VARCHAR(500)  NULL
    );
    CREATE NONCLUSTERED INDEX IX_tblMailLog_message_id ON dbo.tblMailLog (message_id);
    CREATE NONCLUSTERED INDEX IX_tblMailLog_received ON dbo.tblMailLog (received DESC);
    PRINT 'Tabla tblMailLog creada.';
END
GO

IF OBJECT_ID(N'dbo.tblMailOutbox', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tblMailOutbox (
        id          INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_tblMailOutbox PRIMARY KEY,
        created     DATETIME      NOT NULL CONSTRAINT DF_tblMailOutbox_created DEFAULT (GETDATE()),
        to_addr     VARCHAR(255)  NOT NULL,
        subject     VARCHAR(255)  NULL,
        body        VARCHAR(MAX)  NULL,
        in_reply_to VARCHAR(255)  NULL,
        refs        VARCHAR(1500) NULL,
        problem_id  INT           NULL,
        msg_id      VARCHAR(255)  NULL,    -- Message-ID con el que salio (para reconocer las respuestas del usuario)
        status      INT           NOT NULL CONSTRAINT DF_tblMailOutbox_status DEFAULT (0),  -- 0 pendiente, 1 enviado, 2 error
        tries       INT           NOT NULL CONSTRAINT DF_tblMailOutbox_tries DEFAULT (0),
        last_error  VARCHAR(500)  NULL,
        sent        DATETIME      NULL
    );
    CREATE NONCLUSTERED INDEX IX_tblMailOutbox_status ON dbo.tblMailOutbox (status, id);
    CREATE NONCLUSTERED INDEX IX_tblMailOutbox_msg_id ON dbo.tblMailOutbox (msg_id);
    PRINT 'Tabla tblMailOutbox creada.';
END
GO

-- Direcciones de entrada por categoria (alias que entregan al buzon unico)
IF OBJECT_ID(N'dbo.tblMailAlias', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tblMailAlias (
        id             INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_tblMailAlias PRIMARY KEY,
        address        VARCHAR(150) NOT NULL,     -- direccion a la que escribe el usuario (en minusculas)
        category_id    INT          NOT NULL,     -- categoria que se asigna al caso
        def_department INT          NULL,         -- departamento por omision para esta direccion (opcional)
        def_priority   INT          NULL,         -- prioridad por omision para esta direccion (opcional)
        enabled        BIT          NOT NULL CONSTRAINT DF_tblMailAlias_enabled DEFAULT (1),
        created        DATETIME     NOT NULL CONSTRAINT DF_tblMailAlias_created DEFAULT (GETDATE())
    );
    CREATE UNIQUE NONCLUSTERED INDEX UX_tblMailAlias_address ON dbo.tblMailAlias (address);
    PRINT 'Tabla tblMailAlias creada.';
END
GO

-- Direccion Reply-To de las respuestas (la direccion de entrada por la que llego el caso)
IF COL_LENGTH('dbo.tblMailOutbox', 'reply_to') IS NULL
    ALTER TABLE dbo.tblMailOutbox ADD reply_to VARCHAR(255) NULL;
GO

-- Permisos: el usuario con el que se conecta la aplicacion debe poder leer y escribir en estas tablas.
-- GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.tblMailConfig TO [usuario_de_la_aplicacion];
-- GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.tblMailLog    TO [usuario_de_la_aplicacion];
-- GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.tblMailOutbox TO [usuario_de_la_aplicacion];
-- GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.tblMailAlias  TO [usuario_de_la_aplicacion];
