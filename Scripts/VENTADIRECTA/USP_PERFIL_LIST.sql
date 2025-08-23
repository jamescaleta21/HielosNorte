IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'PROCEDURE'	-- Validación del tipo
            		AND ROUTINE_SCHEMA = 'dbo'		-- Validación del esquema
                    AND S.ROUTINE_NAME = 'USP_PERFIL_LIST' )		-- Validación del nombre
    BEGIN
        DROP PROC [dbo].[USP_PERFIL_LIST]
    END
GO
/*
USP_PERFIL_LIST
*/
CREATE PROCEDURE [dbo].[USP_PERFIL_LIST]

WITH ENCRYPTION
AS
BEGIN
SET NOCOUNT ON
SELECT ide = -1, nom = '.: SELECCIONE :.'
UNION
SELECT ide = p.IDPERFIL, nom = p.DENOMINACION FROM dbo.PERFIL p
WHERE p.ACTIVO = 1 AND p.ELIMINADO = 0
END
GO

