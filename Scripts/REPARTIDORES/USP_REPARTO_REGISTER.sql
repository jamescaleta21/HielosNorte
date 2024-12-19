IF EXISTS
(
    SELECT TOP (1)
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_REPARTO_REGISTER'
)
BEGIN
    DROP PROC [dbo].[USP_REPARTO_REGISTER];
END;
GO
/*
USP_REPARTO_REGISTER '01','20241030',9,'NADA',100,'CALETA','[{"idp":1,"idv":10},{"idp":2,"idv":20}]'
USP_REPARTO_REGISTER '01','20241030',0,'NADA',100,'CALETA','[{"idp":1,"idv":10},{"idp":2,"idv":20}]'
*/
CREATE PROCEDURE [dbo].[USP_REPARTO_REGISTER]
    @CODCIA CHAR(2),
    @FECHA DATE,
    @IDREPARTIDOR INT,
    @PESO DECIMAL(16, 2),
    @CURRENTUSER VARCHAR(20),
    @DET NVARCHAR(MAX),
    @OBS VARCHAR(300) = NULL
WITH ENCRYPTION
AS
SET NOCOUNT ON;

DECLARE @EXITO VARCHAR(300);

SET @EXITO = '0=Registrado Satisfactoriamente.';

IF @IDREPARTIDOR = 0
BEGIN
    SET @EXITO = '-1=Debe ingresar el Repartidor';
    GOTO Terminar;
END;

BEGIN TRAN;
BEGIN TRY

    DECLARE @posInicio INT;
    DECLARE @posFin INT;
    DECLARE @parObjeto NVARCHAR(MAX);
    DECLARE @idpedido INT;
    DECLARE @idvendedor INT;

    DECLARE @IDREPARTO INT;

    SELECT TOP (1)
           @IDREPARTO = rc.idReparto
    FROM dbo.REPARTO_CAB rc
    WHERE rc.codCia = @CODCIA
    ORDER BY rc.idReparto DESC;

    IF @IDREPARTO IS NULL
    BEGIN
        SET @IDREPARTO = 1;
    END;
    ELSE
    BEGIN
        SET @IDREPARTO = @IDREPARTO + 1;
    END;

    INSERT INTO dbo.REPARTO_CAB
    (
        codCia,
        idReparto,
        fecha,
        idRepartidor,
        observaciones,
        peso,
        cu_Register
    )
    VALUES
    (   @CODCIA,       -- codCia - char(2)
        @IDREPARTO,    -- idReparto - bigint
        @FECHA,        -- fecha - date
        @IDREPARTIDOR, -- idRepartidor - int
        @OBS,          -- observaciones - varchar(300)
        @PESO,         -- peso - decimal(16, 2)
        @CURRENTUSER   -- cu_Register - varchar(20)
        );

    --REGISTRANDO EN DETALLE

    -- Inicializar la posición inicial
    SET @posInicio = CHARINDEX('{', @DET);

    -- Iterar mientras existan objetos en el JSON
    WHILE @posInicio > 0
    BEGIN
        -- Buscar el final del objeto
        SET @posFin = CHARINDEX('}', @DET, @posInicio);
        SET @parObjeto = SUBSTRING(@DET, @posInicio, @posFin - @posInicio + 1);

        -- Extraer idpedido
        SET @idpedido
            = CAST(SUBSTRING(
                                @parObjeto,
                                CHARINDEX('idp', @parObjeto) + 5,
                                CHARINDEX(',', @parObjeto) - CHARINDEX('idp', @parObjeto) - 5
                            ) AS INT);

        -- Extraer idvendedor
        SET @idvendedor
            = CAST(SUBSTRING(
                                @parObjeto,
                                CHARINDEX('idv', @parObjeto) + 5,
                                CHARINDEX('}', @parObjeto) - CHARINDEX('idv', @parObjeto) - 5
                            ) AS INT);


        INSERT INTO dbo.REPARTO_DET
        (
            codCia,
            idReparto,
            idPedido,
            idVendedor
        )
        VALUES
        (   @CODCIA,    -- codCia - char(2)
            @IDREPARTO, -- idReparto - bigint
            @idpedido,  -- idPedido - bigint
            @idvendedor -- idVendedor - bigint
            );

		--REGISTRANDO EN REPARTO_DET_ITEM
		INSERT INTO dbo.REPARTO_DET_ITEM
		(
		    codCia,
		    idReparto,
		    idPedido,
		    idProducto,
		    cantidad
		)
		SELECT @CODCIA,@IDREPARTO, PD.idpedido,pd.idproducto,pd.cantidad FROM dbo.PEDIDO_DETALLE pd WITH (NOLOCK)
		WHERE pd.idpedido = @idpedido

        --actualizar pedido
        UPDATE dbo.PEDIDO
        SET idRepartidor = @IDREPARTIDOR
        WHERE idpedido = @idpedido;

        -- Actualizar el JSON eliminando el objeto procesado
        SET @DET = SUBSTRING(@DET, @posFin + 1, LEN(@DET) - @posFin);

        -- Buscar el siguiente objeto
        SET @posInicio = CHARINDEX('{', @DET);
    END;
END TRY
BEGIN CATCH
    SET @EXITO = RTRIM(LTRIM(STR(ERROR_NUMBER()))) + '=' + ERROR_MESSAGE();
    ROLLBACK TRAN;
    GOTO Terminar;
END CATCH;


IF @@TRANCOUNT > 0
    COMMIT;

Terminar:
SELECT @EXITO;

GO