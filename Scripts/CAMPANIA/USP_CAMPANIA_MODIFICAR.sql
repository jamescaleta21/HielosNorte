IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_CAMPANIA_MODIFICAR'
)
BEGIN
    DROP PROC [dbo].[USP_CAMPANIA_MODIFICAR];
END;
GO
/*
USP_CAMPANIA_REGISTRAR '01','CAMPAÑA 1','20240501','20240701',100,'JAMES'
*/
CREATE PROCEDURE [dbo].[USP_CAMPANIA_MODIFICAR]
    @CODCIA CHAR(2),
    @NOMBRE VARCHAR(100),
    @INI DATE,
    @FIN DATE,
    @MONTO MONEY,
    @CURRENTUSER VARCHAR(20),
    @IDCAMPANIA INT
WITH ENCRYPTION
AS
SET NOCOUNT ON;
DECLARE @Exito VARCHAR(300);
DECLARE @IDLOG INT;

SET @Exito = '0=Campaña Modificada Correctamente.';


BEGIN TRAN;
BEGIN TRY

    UPDATE dbo.CAMPANIA
    SET FECHAINICIO = @INI,
        FECHAFIN = @FIN,
        MONTO = @MONTO,
        NOMBRE = @NOMBRE
    WHERE CODCIA = @CODCIA
          AND IDCAMPANIA = @IDCAMPANIA;


    SELECT TOP 1
           @IDLOG = cl.IDLOG + 1
    FROM dbo.CAMPANIA_LOG cl
    WHERE cl.CODCIA = @CODCIA
          AND cl.IDCAMPANIA = @IDCAMPANIA
    ORDER BY cl.IDLOG DESC;

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
           @IDLOG,
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
    ROLLBACK TRAN;
    GOTO Terminar;
END CATCH;


IF @@TRANCOUNT > 0
    COMMIT;

Terminar:
SELECT @Exito AS 'mensaje';
GO

