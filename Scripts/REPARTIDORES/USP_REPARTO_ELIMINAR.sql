IF EXISTS
(
    SELECT TOP (1)
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_REPARTO_ELIMINAR'
)
BEGIN
    DROP PROC [dbo].[USP_REPARTO_ELIMINAR];
END;
GO
/*
USP_REPARTO_REGISTER '01','20241030',9,'NADA',100,'CALETA','[{"idp":1,"idv":10},{"idp":2,"idv":20}]'
USP_REPARTO_ELIMINAR '01',8,'demo2','caleta'
*/
CREATE PROCEDURE [dbo].[USP_REPARTO_ELIMINAR]
    @CODCIA CHAR(2),
    @IDREPARTO BIGINT,
    @MOTIVO VARCHAR(200),
    @CURRENTUSER VARCHAR(20)
WITH ENCRYPTION
AS
SET NOCOUNT ON;

DECLARE @EXITO VARCHAR(300);

SET @EXITO = '0=Eliminado Satisfactoriamente.';

IF
(
    SELECT TOP 1
           COALESCE(rc.onCloud, 0)
    FROM dbo.REPARTO_CAB rc
    WHERE rc.codCia = @CODCIA
          AND rc.idReparto = @IDREPARTO
) = 1
BEGIN
    SET @EXITO = '-1=No se puede eliminar el Reparto.';
    GOTO Terminar;
END;

IF
(
    SELECT TOP 1
           rc.motivoElimina
    FROM dbo.REPARTO_CAB rc
    WHERE rc.codCia = @CODCIA
          AND rc.idReparto = @IDREPARTO
) IS null
BEGIN
    SET @EXITO = '-2=El reparto ya fue eliminado.';
    GOTO Terminar;
END;

--BEGIN TRAN;
BEGIN TRY

    UPDATE dbo.REPARTO_CAB
    SET motivoElimina = @MOTIVO,
        eliminado = 1,
        cu_Delete = @CURRENTUSER,
        fe_Delete = GETDATE()
    WHERE codCia = @CODCIA
          AND idReparto = @IDREPARTO;

END TRY
BEGIN CATCH
    SET @EXITO = RTRIM(LTRIM(STR(ERROR_NUMBER()))) + '=' + ERROR_MESSAGE();
    --ROLLBACK TRAN;
    GOTO Terminar;
END CATCH;


--IF @@TRANCOUNT > 0
--    COMMIT;

Terminar:
SELECT @EXITO;

GO