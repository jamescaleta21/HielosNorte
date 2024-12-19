IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_CAMPANIA_REGISTRAR_CLOUD'
)
BEGIN
    DROP PROC [dbo].[USP_CAMPANIA_REGISTRAR_CLOUD];
END;
GO
/*
USP_CAMPANIA_REGISTRAR_CLOUD '15',1
*/
CREATE PROCEDURE [dbo].[USP_CAMPANIA_REGISTRAR_CLOUD]
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

--BEGIN TRAN;
BEGIN TRY

    SELECT TOP 1
           @IDEMPRESA = e.IdEmpresa
    FROM dbo.EMPRESA e
    WHERE e.Activo = 1
          AND e.Eliminado = 0
          AND e.Defecto = 1;

    --DATOS DE LA CAMPAÑA
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

    SET @INSERTQUERY = 'INSERT INTO OPENQUERY(' + @LINKEDSERVER + ', ''';

    SET @TSQL1
        = 'select idEmpresa, idCampania, nombre, fechaInicio, fechafin, monto, activo, eliminado from CAMPANIA'')';


    SET @TSQL2
        = ' select ' + CAST(@IDEMPRESA AS VARCHAR(20)) + ',' + CAST(@IDCAMPANIA AS VARCHAR(20)) + ', ''' + @NOMBRE
          + ''', ''' + CONVERT(CHAR(8), @INI, 112) + ''',''' + CONVERT(CHAR(8), @FIN, 112) + ''','
          + CAST(@MONTO AS VARCHAR(20)) + ',' + CAST(@ACTIVO AS VARCHAR(20)) + ',' + CAST(@ELIMINADO AS VARCHAR(20))
          + '';

    EXEC (@INSERTQUERY + @TSQL1 + @TSQL2);
--SELECT @INSERTQUERY + @TSQL1 + @TSQL2;


END TRY
BEGIN CATCH
    SET @EXITO = RTRIM(LTRIM(STR(ERROR_NUMBER()))) + '=' + ERROR_MESSAGE();
    --ROLLBACK TRAN;
    GOTO Terminar;
END CATCH;


--IF @@TRANCOUNT > 0
--    COMMIT;

Terminar:
SELECT @EXITO AS 'mensaje';
GO

