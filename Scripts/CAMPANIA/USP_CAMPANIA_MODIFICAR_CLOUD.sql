IF EXISTS (
    SELECT TOP 1 s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_CAMPANIA_MODIFICAR_CLOUD'
)
BEGIN
    DROP PROCEDURE [dbo].[USP_CAMPANIA_MODIFICAR_CLOUD];
END;
GO
/*
USP_CAMPANIA_MODIFICAR_CLOUD '15',1
*/
CREATE PROCEDURE [dbo].[USP_CAMPANIA_MODIFICAR_CLOUD]
    @CODCIA CHAR(2),
    @IDCAMPANIA INT
WITH ENCRYPTION
AS
SET NOCOUNT ON;

DECLARE @IDEMPRESA INT;
DECLARE @LINKEDSERVER VARCHAR(20);
DECLARE @INSERTQUERY VARCHAR(100);
DECLARE @TSQL1 VARCHAR(MAX),
        @TSQL2 VARCHAR(MAX);

DECLARE @EXITO VARCHAR(300);

SET @EXITO = '0=Datos Enviados Correctamente a Cloud.';

BEGIN TRY
    SELECT TOP 1 @IDEMPRESA = e.IdEmpresa
    FROM dbo.EMPRESA e
    WHERE e.Activo = 1
          AND e.Eliminado = 0
          AND e.Defecto = 1;

    -- DATOS DE LA CAMPAÑA
    DECLARE @NOMBRE VARCHAR(100),
            @INI DATE,
            @FIN DATE,
            @MONTO MONEY,
            @ACTIVO BIT,
            @ELIMINADO BIT;

    SELECT TOP 1
           @NOMBRE = c.NOMBRE,
           @INI = c.FECHAINICIO,
           @FIN = c.FECHAFIN,
           @MONTO = c.MONTO,
           @ACTIVO = c.ACTIVO,
           @ELIMINADO = c.ELIMINADO
    FROM dbo.CAMPANIA c
    WHERE c.CODCIA = @CODCIA
          AND c.IDCAMPANIA = @IDCAMPANIA;

    SET @LINKEDSERVER = 'TICKETS';

    SET @INSERTQUERY = 'UPDATE OPENQUERY(' + @LINKEDSERVER + ', ''';

    SET @TSQL1 = 'SELECT idEmpresa, idCampania, nombre, fechaInicio, fechaFin, monto, activo, eliminado FROM CAMPANIA WHERE idEmpresa =' + CAST(@IDEMPRESA AS VARCHAR(10)) 
        + ' AND idCampania = ' + CAST(@IDCAMPANIA AS VARCHAR(10)) + ''')';

    SET @TSQL2 = ' SET nombre = ''' + @NOMBRE + ''', fechaInicio = ''' + CONVERT(CHAR(8), @INI, 112) + ''', fechaFin = ''' + CONVERT(CHAR(8), @FIN, 112) + ''', monto = ' 
        + CAST(@MONTO AS VARCHAR(20)) + ', activo = ' + CAST(@ACTIVO AS VARCHAR(1)) + ', eliminado = ' + CAST(@ELIMINADO AS VARCHAR(1));

    -- Ejecutar la consulta construida
    EXEC (@INSERTQUERY + @TSQL1 + @TSQL2);
	--SELECT @INSERTQUERY + @TSQL1 + @TSQL2;


END TRY
BEGIN CATCH
    SET @EXITO = RTRIM(LTRIM(STR(ERROR_NUMBER()))) + '=' + ERROR_MESSAGE();
    GOTO Terminar;
END CATCH;

Terminar:
SELECT @EXITO AS 'mensaje';
GO
