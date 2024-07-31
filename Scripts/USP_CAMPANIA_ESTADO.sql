IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_CAMPANIA_ESTADO'
)
BEGIN
    DROP PROC [dbo].[USP_CAMPANIA_ESTADO];
END;
GO
/*
USP_CAMPANIA_REGISTRAR '01','CAMPAÑA 1','20240501','20240701',100,'JAMES'
*/
CREATE PROCEDURE [dbo].[USP_CAMPANIA_ESTADO]
    @CODCIA CHAR(2),
    @IDCAMPANIA INT,
    @ESTADO BIT,
    @CURRENTUSER VARCHAR(20)
WITH ENCRYPTION
AS
SET NOCOUNT ON;
DECLARE @Exito VARCHAR(300);

DECLARE @IDLOG INT;

SET @Exito = '0=Campaña Actualizada Correctamente.';


IF @ESTADO = 1
BEGIN
    IF
    (
        SELECT COUNT(c.IDCAMPANIA)
        FROM dbo.CAMPANIA c
        WHERE c.CODCIA = @CODCIA
              AND c.ACTIVO = 1
              AND c.IDCAMPANIA <> @IDCAMPANIA
    ) <> 0
    BEGIN
        SET @Exito = '-1=Ya existe una Campaña activa.';
        GOTO Terminar;
    END;
END;

BEGIN TRAN;
BEGIN TRY

    UPDATE dbo.CAMPANIA
    SET ACTIVO = @ESTADO
    WHERE CODCIA = @CODCIA
          AND IDCAMPANIA = @IDCAMPANIA;

    IF @ESTADO = 0
    BEGIN
        UPDATE dbo.CAMPANIA
        SET ELIMINADO = 1,
            CU_DELETE = @CURRENTUSER,
            FE_DELETE = GETDATE()
        WHERE CODCIA = @CODCIA
              AND IDCAMPANIA = @IDCAMPANIA;
    END;
    ELSE
    BEGIN
        UPDATE dbo.CAMPANIA
        SET ELIMINADO = 0,
            CU_DELETE = NULL,
            FE_DELETE = NULL
        WHERE CODCIA = @CODCIA
              AND IDCAMPANIA = @IDCAMPANIA;
    END;


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

