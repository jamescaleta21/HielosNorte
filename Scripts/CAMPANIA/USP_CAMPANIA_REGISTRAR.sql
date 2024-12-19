IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_CAMPANIA_REGISTRAR'
)
BEGIN
    DROP PROC [dbo].[USP_CAMPANIA_REGISTRAR];
END;
GO
/*
USP_CAMPANIA_REGISTRAR '01','CAMPAÑA 1','20240501','20240701',100,'JAMES'
*/
CREATE PROCEDURE [dbo].[USP_CAMPANIA_REGISTRAR]
    @CODCIA CHAR(2),
    @NOMBRE VARCHAR(100),
    @INI DATE,
    @FIN DATE,
    @MONTO MONEY,
    @CURRENTUSER VARCHAR(20)
WITH ENCRYPTION
AS
SET NOCOUNT ON;
DECLARE @Exito VARCHAR(300);
DECLARE @IDCAMPANIA INT;
DECLARE @IDLOG INT;

SET @Exito = '0=Campaña Registrada Correctamente.';

IF EXISTS (SELECT TOP 1 'X' FROM dbo.CAMPANIA c WHERE c.NOMBRE = @NOMBRE)
BEGIN
    SET @Exito = '-1=Nombre de Campaña ya existe.';
    GOTO Terminar;
END;

IF EXISTS
(
    SELECT TOP 1
           'X'
    FROM dbo.CAMPANIA c
    WHERE @INI
    BETWEEN c.FECHAINICIO AND c.FECHAFIN
)
   OR EXISTS
(
    SELECT TOP 1
           'X'
    FROM dbo.CAMPANIA c
    WHERE @FIN
    BETWEEN c.FECHAINICIO AND c.FECHAFIN
)
BEGIN
    SET @Exito = '-2=Fecha se cruza con otras Campañas.';
    GOTO Terminar;
END;


BEGIN TRAN;
BEGIN TRY

    SELECT TOP 1
           @IDCAMPANIA = c.IDCAMPANIA
    FROM dbo.CAMPANIA c
    WHERE c.CODCIA = @CODCIA
    ORDER BY c.IDCAMPANIA DESC;

    IF @IDCAMPANIA IS NULL
    BEGIN
        SET @IDCAMPANIA = 0;
    END;

    SET @IDCAMPANIA = @IDCAMPANIA + 1;

    INSERT INTO dbo.CAMPANIA
    (
        IDCAMPANIA,
        CODCIA,
        NOMBRE,
        FECHAINICIO,
        FECHAFIN,
        MONTO,
        ACTIVO,
        CU_REGISTER
    )
    VALUES
    (   @IDCAMPANIA,    -- IDCAMPANIA - int
        @CODCIA,        -- CODCIA - char(2)
        @NOMBRE,        -- NOMBRE - varchar(100)
        @INI,           -- FECHAINICIO - date
        @FIN,           -- FECHAFIN - date
        @MONTO,         -- MONTO - money
        0, @CURRENTUSER -- CU_REGISTER - varchar(20)
        );

    --REGISTRANDO EN LOG

    INSERT INTO dbo.CAMPANIA_LOG
    (
        IDCAMPANIA,
        CODCIA,
        IDLOG,
        NOMBRE,
        FECHAINICIO,
        FECHAFIN,
        MONTO,
        ACTIVO,
        CU_REGISTER,
        FE_REGISTER,
        ELIMINADO,
        CU_DELETE,
        FE_DELETE
    )
    SELECT IDCAMPANIA,
           CODCIA,
           1,
           NOMBRE,
           FECHAINICIO,
           FECHAFIN,
           MONTO,
           ACTIVO,
           CU_REGISTER,
           FE_REGISTER,
           ELIMINADO,
           CU_DELETE,
           FE_DELETE
    FROM dbo.CAMPANIA c
    WHERE c.CODCIA = @CODCIA
          AND c.IDCAMPANIA = @IDCAMPANIA;



END TRY
BEGIN CATCH
    SET @Exito = RTRIM(LTRIM(STR(ERROR_NUMBER()))) + '=' + ERROR_MESSAGE();
    SET @IDCAMPANIA = 0;
    ROLLBACK TRAN;
    GOTO Terminar;
END CATCH;


IF @@TRANCOUNT > 0
    COMMIT;

Terminar:
SELECT @Exito AS 'mensaje',
       @IDCAMPANIA AS 'ide';
GO

