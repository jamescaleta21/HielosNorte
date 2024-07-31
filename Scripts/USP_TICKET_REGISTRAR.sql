/*
SELECT * FROM dbo.CAMPANIAS c
SELECT * FROM dbo.TICKET t
DELETE FROM dbo.TICKET
*/

IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_TICKET_REGISTRAR'
)
BEGIN
    DROP PROC [dbo].[USP_TICKET_REGISTRAR];
END;
GO
/*
USP_TICKET_REGISTRAR '15','B','2',7182
*/
CREATE PROCEDURE [dbo].[USP_TICKET_REGISTRAR]
    @CODCIA CHAR(2),
	@FECHA DATE,
    @TIPODOCTO CHAR(1),
    @SERIE VARCHAR(3),
    @NUMERO BIGINT,
	@CURRENTUSER VARCHAR(20)
WITH ENCRYPTION
AS
SET NOCOUNT ON;


DECLARE @MONTOCAMPANIA MONEY,
        @IDCAMPANIA INT;
DECLARE @MIN INT,
        @MAX INT;
DECLARE @IDTICKET BIGINT;
DECLARE @IDCLIENTE BIGINT;
DECLARE @RUC CHAR(11),
        @IMPORTE MONEY;


-- OBTENIENDO EL TOTAL DEL DOCUMENTO
SELECT TOP 1
       @IMPORTE = a.ALL_IMPORTE_AMORT,
	   @IDCLIENTE = A.ALL_CODCLIE
FROM dbo.ALLOG a
WHERE a.ALL_FBG = @TIPODOCTO
      AND RTRIM(LTRIM(a.ALL_NUMSER)) = @SERIE
      AND a.ALL_NUMFAC = @NUMERO
      AND a.ALL_CODCIA = @CODCIA
      AND a.ALL_CODTRA = 2401;



SELECT TOP 1
       @RUC = p.PAR_RUC_EMP
FROM dbo.PARGEN p
WHERE p.PAR_CODCIA = @CODCIA;

--OBTENIENDO EL MONTO DE LA CAMPAÑA ACTIVA
SELECT TOP 1
       @MONTOCAMPANIA = COALESCE(c.MONTO, 0),
       @IDCAMPANIA = c.IDCAMPANIA
FROM dbo.CAMPANIA c
WHERE 
--c.CODCIA = @CODCIA AND 
	  c.ACTIVO = 1
      AND c.ELIMINADO = 0;

--SELECT @MONTOCAMPANIA;

DECLARE @RESULTADO INT;
SET @RESULTADO = 0;

SET @RESULTADO = FLOOR(@IMPORTE / @MONTOCAMPANIA);

--SELECT @RESULTADO;


IF @RESULTADO <> 0
BEGIN

    SET @MIN = 1;
    SET @MAX = @RESULTADO;

    WHILE @MIN <= @MAX
    BEGIN
        --OBTENIENDO EL IDTICKET SEGUN EL CODCIA - inicio
        SELECT TOP 1
               @IDTICKET = t.IDTICKET
        FROM dbo.TICKET t
        --WHERE t.CODCIA = @CODCIA
        ORDER BY t.IDTICKET DESC;

        IF @IDTICKET IS NULL
        BEGIN
            SET @IDTICKET = 0;
        END;

        SET @IDTICKET = @IDTICKET + 1;
        --OBTENIENDO EL IDTICKET SEGUN EL CODCIA - fin

        INSERT INTO dbo.TICKET
        (
            CODCIA,
            IDTICKET,
			IDCAMPANIA,
            RUCEMISOR,
            FECHAEMISION,
            TIPOCOMPROBANTE,
            SERIECOMPROBANTE,
            NUMEROCOMPROBANTE,
            CODIGOCLIENTE,
            NUMEROTICKET,
			CU_REGISTER
        )
        VALUES
        (   @CODCIA,                                                                                           -- CODCIA - char(2)
            @IDTICKET,                                                                                         -- IDTICKET - bigint
			@IDCAMPANIA,
            @RUC,                                                                                              -- RUCEMISOR - varchar(11)
            @FECHA,	                                                                                           -- FECHAEMISION - date
            @TIPODOCTO,                                                                                        -- TIPOCOMPROBANTE - char(1)
            @SERIE,                                                                                            -- SERIECOMPROBANTE - varchar(3)
            @NUMERO,                                                                                           -- NUMEROCOMPROBANTE - bigint
            @IDCLIENTE,                                                                                                 -- CODIGOCLIENTE - bigint
            'TK' + CAST(@IDCAMPANIA AS VARCHAR(20)) + RIGHT('0000000000' + CAST(@IDTICKET AS VARCHAR(10)), 10) -- NUMEROTICKET - varchar(20)
			,@CURRENTUSER
            );

        SET @MIN = @MIN + 1;
    END;


END;



GO