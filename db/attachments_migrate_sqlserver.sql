/*
  Liberum Help Desk - migracion de adjuntos (SQL Server)

  1) Convierte la columna data de tblAttachments de binario (VARBINARY) a texto
     base64 (VARCHAR(MAX)), para aceptar cualquier tipo de archivo.
  2) Agrega la columna AllowedAttachments a tblConfig: lista de extensiones
     permitidas, separada por comas (la configura el administrador).

  Es seguro ejecutarlo varias veces. Si la tabla tiene archivos guardados en
  formato binario, la conversion se detiene para no perderlos.
*/

-- 1) Columna data: binario -> texto base64
IF EXISTS (SELECT 1 FROM sys.columns c JOIN sys.types t ON c.user_type_id = t.user_type_id
           WHERE c.object_id = OBJECT_ID(N'dbo.tblAttachments') AND c.name = N'data' AND t.name = N'varbinary')
BEGIN
    IF EXISTS (SELECT 1 FROM dbo.tblAttachments WHERE data IS NOT NULL)
        RAISERROR(N'tblAttachments tiene archivos en formato binario. Revise antes de convertir la columna data.', 16, 1);
    ELSE
    BEGIN
        ALTER TABLE dbo.tblAttachments ALTER COLUMN data VARCHAR(MAX) NULL;
        PRINT 'Columna tblAttachments.data convertida a VARCHAR(MAX) (base64).';
    END
END
ELSE
    PRINT 'Columna tblAttachments.data ya es texto. Sin cambios.';
GO

-- 2) Lista de tipos de adjuntos permitidos en la configuracion
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.tblConfig') AND name = N'AllowedAttachments')
BEGIN
    ALTER TABLE dbo.tblConfig ADD AllowedAttachments VARCHAR(255) NULL;
    PRINT 'Columna tblConfig.AllowedAttachments agregada.';
END
GO

UPDATE dbo.tblConfig
   SET AllowedAttachments = 'jpg,jpeg,png,gif,webp,pdf,doc,docx,xls,xlsx,txt,zip'
 WHERE AllowedAttachments IS NULL;
GO
