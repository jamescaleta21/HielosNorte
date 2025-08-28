IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'PROCEDURE'	-- Validación del tipo
            		AND ROUTINE_SCHEMA = 'dbo'		-- Validación del esquema
                    AND S.ROUTINE_NAME = 'USP_VENDEDOR_LIST' )		-- Validación del nombre
    BEGIN
        DROP PROC [dbo].[USP_VENDEDOR_LIST]
    END
GO
/*
USP_VENDEDOR_LIST '01'
*/
CREATE PROCEDURE [dbo].[USP_VENDEDOR_LIST] @CODCIA CHAR(2)
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT v.VEM_CODVEN AS cod,
       RTRIM(LTRIM(v.VEM_NOMBRE)) AS nom
FROM dbo.VEMAEST v
WHERE v.VEM_CODCIA = @CODCIA
      AND v.VEM_ACTIVO = 1
	  AND v.VEM_IDPERFIL = 2
ORDER BY v.VEM_NOMBRE;
GO

