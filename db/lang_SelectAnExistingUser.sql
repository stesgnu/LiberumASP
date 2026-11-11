-- Agrega la cadena SelectAnExistingUser (SQL Server o Access).
-- En ejecucion la aplicacion lee tblLangStrings, no los archivos lang/*.txt.
-- Es seguro ejecutarlo varias veces.

-- 1) Idioma por omision del sitio (tblConfig.DefaultLanguage = Spanish/Espanol)
INSERT INTO tblLangStrings (id, variable, LangText)
SELECT DefaultLanguage, 'SelectAnExistingUser', 'Seleccionar usuario existente'
FROM tblConfig
WHERE NOT EXISTS (SELECT 1 FROM tblLangStrings s
                  WHERE s.id = tblConfig.DefaultLanguage AND s.variable = 'SelectAnExistingUser');

-- 2) Todo idioma llamado Spanish (por si los usuarios usan otro id de espanol)
INSERT INTO tblLangStrings (id, variable, LangText)
SELECT id, 'SelectAnExistingUser', 'Seleccionar usuario existente'
FROM tblLanguage
WHERE LangName = 'Spanish'
  AND NOT EXISTS (SELECT 1 FROM tblLangStrings s
                  WHERE s.id = tblLanguage.id AND s.variable = 'SelectAnExistingUser');

-- 3) Ingles
INSERT INTO tblLangStrings (id, variable, LangText)
SELECT id, 'SelectAnExistingUser', 'Select an existing user'
FROM tblLanguage
WHERE LangName = 'English'
  AND NOT EXISTS (SELECT 1 FROM tblLangStrings s
                  WHERE s.id = tblLanguage.id AND s.variable = 'SelectAnExistingUser');

-- Verificacion: debe haber una fila por idioma
-- SELECT s.id, l.LangName, s.variable, s.LangText
-- FROM tblLangStrings s LEFT JOIN tblLanguage l ON l.id = s.id
-- WHERE s.variable = 'SelectAnExistingUser';
-- SELECT DefaultLanguage FROM tblConfig;
