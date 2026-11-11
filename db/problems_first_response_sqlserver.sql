/*
  Liberum Help Desk - columna first_response en problems (SQL Server)

  rep/details.asp, rep/new.asp y el reporte de ANS (admin/reportans.asp) usan la fecha de la
  primera respuesta de un tecnico. En bases creadas antes de esa version la columna no existe
  y las consultas fallan con: Invalid column name 'first_response'.

  - Es seguro ejecutarlo varias veces.
  - Los casos existentes quedan con NULL (sin fecha de primera respuesta registrada).
*/
IF COL_LENGTH('dbo.problems', 'first_response') IS NULL
BEGIN
    ALTER TABLE dbo.problems ADD first_response DATETIME NULL;
    PRINT 'Columna first_response agregada a dbo.problems.';
END
ELSE
    PRINT 'La columna first_response ya existe.';
GO
