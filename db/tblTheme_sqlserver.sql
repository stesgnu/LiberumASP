-- Paleta de colores de la aplicacion (admin/colors.asp). SQL Server.
-- Ejecutar una vez sobre la base de datos del Help Desk.
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.tblTheme') AND type = N'U')
BEGIN
    CREATE TABLE dbo.tblTheme (
        theme_key   VARCHAR(30) NOT NULL PRIMARY KEY,
        theme_value VARCHAR(20) NULL
    );
END
GO

-- Opcional: valores iniciales (si no se insertan, la aplicacion usa los colores por defecto)
-- INSERT INTO dbo.tblTheme (theme_key, theme_value) VALUES ('primary',       '#0057B8');
-- INSERT INTO dbo.tblTheme (theme_key, theme_value) VALUES ('primary_hover', '#003F8A');
-- INSERT INTO dbo.tblTheme (theme_key, theme_value) VALUES ('sidebar',       '#0F172A');
-- INSERT INTO dbo.tblTheme (theme_key, theme_value) VALUES ('accent',        '#0284C7');
-- INSERT INTO dbo.tblTheme (theme_key, theme_value) VALUES ('bg',            '#F1F5F9');

-- Permisos: el usuario con que se conecta la aplicacion debe poder leer y escribir en la tabla.
-- GRANT SELECT, INSERT, UPDATE, DELETE ON dbo.tblTheme TO [usuario_de_la_aplicacion];
