/*
  Liberum Help Desk - indices recomendados (SQL Server)

  En el esquema original, tblUsers, tblNotes, tblLangStrings y tblLanguage NO tienen
  llave primaria ni indices, y problems solo tiene el de id. Cada pagina hace varias
  consultas por tblUsers.sid, tblNotes.id y tblLangStrings.(id, variable), asi que sin
  indices SQL Server recorre la tabla completa en cada una.

  - Es seguro ejecutarlo varias veces (crea cada indice solo si no existe).
  - Ejecutar sobre la base de datos de la mesa de ayuda.
  - Son indices NO agrupados: no reorganizan las tablas.
*/

-- tblUsers: Usr(sid, ...) se llama varias veces por pagina; el login busca por uid
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_tblUsers_sid' AND object_id = OBJECT_ID('dbo.tblUsers'))
    CREATE NONCLUSTERED INDEX IX_tblUsers_sid ON dbo.tblUsers (sid);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_tblUsers_uid' AND object_id = OBJECT_ID('dbo.tblUsers'))
    CREATE NONCLUSTERED INDEX IX_tblUsers_uid ON dbo.tblUsers (uid);
GO

-- tblNotes: notas de un ticket (detalle, impresion, KB)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_tblNotes_id' AND object_id = OBJECT_ID('dbo.tblNotes'))
    CREATE NONCLUSTERED INDEX IX_tblNotes_id ON dbo.tblNotes (id);
GO

-- tblLangStrings: busqueda de una cadena por idioma y nombre
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_tblLangStrings_id_variable' AND object_id = OBJECT_ID('dbo.tblLangStrings'))
    CREATE NONCLUSTERED INDEX IX_tblLangStrings_id_variable ON dbo.tblLangStrings (id, variable);
GO

-- problems: listas (por tecnico y estado, por usuario), contadores y reportes
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_problems_rep_status' AND object_id = OBJECT_ID('dbo.problems'))
    CREATE NONCLUSTERED INDEX IX_problems_rep_status ON dbo.problems (rep, status) INCLUDE (priority, start_date);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_problems_uid_status' AND object_id = OBJECT_ID('dbo.problems'))
    CREATE NONCLUSTERED INDEX IX_problems_uid_status ON dbo.problems (uid, status) INCLUDE (start_date);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_problems_start_date' AND object_id = OBJECT_ID('dbo.problems'))
    CREATE NONCLUSTERED INDEX IX_problems_start_date ON dbo.problems (start_date) INCLUDE (department, category, rep, time_spent);
GO

-- Estadisticas al dia despues de crear los indices
UPDATE STATISTICS dbo.tblUsers;
UPDATE STATISTICS dbo.tblNotes;
UPDATE STATISTICS dbo.tblLangStrings;
UPDATE STATISTICS dbo.problems;
GO
