/*
  Liberum Help Desk - tabla de adjuntos e imagenes del editor (SQL Server)

  Guarda los archivos adjuntos y las imagenes insertadas en las descripciones
  y soluciones de los tickets, directamente en la base de datos.

  - Es seguro ejecutarlo varias veces: si la tabla ya existe no hace nada.
  - Ejecutar sobre la base de datos de la mesa de ayuda (por ejemplo HelpDesk).
*/

IF OBJECT_ID(N'dbo.tblAttachments', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.tblAttachments (
        id          INT IDENTITY(1,1) NOT NULL,
        problem_id  INT           NOT NULL CONSTRAINT DF_tblAttachments_problem_id DEFAULT (0),
        token       VARCHAR(64)   NULL,
        uid         VARCHAR(50)   NULL,
        fname       VARCHAR(255)  NULL,
        ctype       VARCHAR(100)  NULL,
        fsize       INT           NULL,
        addDate     DATETIME      NULL CONSTRAINT DF_tblAttachments_addDate DEFAULT (GETDATE()),
        data        VARBINARY(MAX) NULL,
        CONSTRAINT PK_tblAttachments PRIMARY KEY CLUSTERED (id)
    );

    -- Busqueda por ticket (listas en detalle) y por token (vinculo al crear el ticket)
    CREATE NONCLUSTERED INDEX IX_tblAttachments_problem_id ON dbo.tblAttachments (problem_id);
    CREATE NONCLUSTERED INDEX IX_tblAttachments_token ON dbo.tblAttachments (token);

    PRINT 'Tabla tblAttachments creada correctamente.';
END
ELSE
BEGIN
    PRINT 'La tabla tblAttachments ya existe. No se hicieron cambios.';
END
GO
