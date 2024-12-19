IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_REPARTO_SYNC'
)
BEGIN
    DROP PROC [dbo].[USP_REPARTO_SYNC];
END;
GO
/*
USP_REPARTO_SYNC
*/
CREATE PROCEDURE [dbo].[USP_REPARTO_SYNC]
WITH ENCRYPTION
AS
SET NOCOUNT ON;

DECLARE @CODCIA CHAR(2);
SET @CODCIA = '01';

DECLARE @MINC INT,
        @MAXC INT;
DECLARE @EXITO VARCHAR(300);
SET @EXITO = '0=Repartos cerrados Correctamente.';
--VARIABLES CABECERA
DECLARE @IDREPARTO INT,
        @FECHA DATE,
        @IDREPARTIDOR INT,
        @OBS VARCHAR(300),
        @PESO DECIMAL(16, 2);
--VARIABLES DETALLE
DECLARE @TBLREPARTODET TABLE
(
    INDICE INT IDENTITY,
    IDPEDIDO BIGINT,
    IDVENDEDOR INT
);
DECLARE @IDPEDIDO BIGINT,
        @IDVENDEDOR INT;
DECLARE @MIND INT,
        @MAXD INT;
--TABLA PARA VALIDACION SI EXISTE IDREPARTO PREVIAMENTE
DECLARE @TBLVALIDA TABLE
(
    DATO CHAR(1)
);


DECLARE @TBLREPARTO TABLE
(
    INDICE INT IDENTITY,
    IDREPARTO INT,
    FECHA DATE
);


IF (  SELECT COUNT( rc.idReparto)
    FROM dbo.REPARTO_CAB rc
    WHERE ISNULL(rc.onCloud, 0)  = 0) =0
	BEGIN
	SET @EXITO='-1=No hay nada que enviar a Cloud.'
	GOTO terminar
    end



--BEGIN TRAN
BEGIN TRY

    DECLARE @IDEMPRESA INT;
    SELECT TOP 1
           @IDEMPRESA = e.IdEmpresa
    FROM dbo.EMPRESA e
    WHERE e.Activo = 1
          AND e.Defecto = 1;

    DECLARE @tsql VARCHAR(4000),
            @openquery VARCHAR(4000),
            @linkedserver VARCHAR(20),
            @insertquery VARCHAR(4000),
            @tsql2 VARCHAR(4000);
    SET @linkedserver = 'SMARTSOFT';
    SET @openquery = 'SELECT * FROM OPENQUERY(' + @linkedserver + ', ''';

	--RECUPERANDO BLOQUE DE REPARTOS PENDIENTES DE SUBIR
    INSERT INTO @TBLREPARTO
    (
        IDREPARTO,
        FECHA
    )
    SELECT rc.idReparto,
           rc.fecha
    FROM dbo.REPARTO_CAB rc
    WHERE ISNULL(rc.onCloud, 0) = 0 AND ISNULL(rc.eliminado,0) = 0
          AND rc.codCia = @CODCIA;

	--RECUPERANDO LOS VALORES MINIMO Y MAXIMO PARA REALIZAR EL RECORRIDO
    SELECT @MINC = MIN(t.INDICE)
    FROM @TBLREPARTO t;
    SELECT @MAXC = MAX(t.INDICE)
    FROM @TBLREPARTO t;

	--RECORRIENDO LOS REPARTOS PENDIENTES DE SUBIR
    WHILE @MINC <= @MAXC
    BEGIN
        SELECT @IDREPARTO = t.IDREPARTO,
               @FECHA = t.FECHA
        FROM @TBLREPARTO t
        WHERE t.INDICE = @MINC;

        SELECT @IDREPARTIDOR = rc.idRepartidor,
               @OBS = COALESCE(rc.observaciones, ''),
               @PESO = rc.peso
        FROM dbo.REPARTO_CAB rc WITH (NOLOCK)
        WHERE rc.idReparto = @IDREPARTO
              AND rc.codCia = @CODCIA
              AND rc.fecha = @FECHA;

        --VALIDANDO SI EL REPARTO SE ENCUENTRA EN CLOUD
        SET @tsql
            = 'SELECT TOP 1 ''''X'''' FROM dbo.REPARTO_CAB rc with (nolock) where rc.idReparto = '
              + CAST(@IDREPARTO AS VARCHAR(10)) + ' and idEmpresa = ' + CAST(@IDEMPRESA AS VARCHAR(10)) + ''');';

        INSERT INTO @TBLVALIDA
        (
            DATO
        )
        EXEC (@openquery + @tsql);


        IF NOT EXISTS (SELECT TOP 1 'X' FROM @TBLVALIDA t)--NO EXISTE EN CLOUD, CONTINUA CON EL FLUJO
        BEGIN
            SET @insertquery = 'INSERT INTO OPENQUERY(' + @linkedserver + ', ''';
            SET @tsql
                = 'SELECT idReparto, idEmpresa, fecha, idRepartidor, observaciones, peso FROM dbo.REPARTO_CAB'')';
            SET @tsql2
                = ' SELECT ' + CAST(@IDREPARTO AS VARCHAR(20)) + ', ' + CAST(@IDEMPRESA AS VARCHAR(10)) + ', '''
                  + CONVERT(VARCHAR(8), @FECHA, 112) + ''',' + CAST(@IDREPARTIDOR AS VARCHAR(10)) + ','''
                  + RTRIM(LTRIM(@OBS)) + ''',' + CAST(@PESO AS VARCHAR(20)) + ';';

            EXEC (@insertquery + @tsql + @tsql2);

            --SUBIENDO TABLA REPARTO_DET
            INSERT INTO @TBLREPARTODET
            (
                IDPEDIDO,
                IDVENDEDOR
            )
            SELECT p.idpedido_cloud,
                   rd.idVendedor
            FROM dbo.REPARTO_DET rd WITH (NOLOCK)
                INNER JOIN dbo.PEDIDO p WITH (NOLOCK)
                    ON rd.idPedido = p.idpedido
                       AND rd.codCia = @CODCIA
            WHERE rd.codCia = @CODCIA
                  AND rd.idReparto = @IDREPARTO;

            SELECT @MIND = MIN(t.INDICE)
            FROM @TBLREPARTODET t;
            SELECT @MAXD = MAX(t.INDICE)
            FROM @TBLREPARTODET t;


            --SELECT * FROM @TBLREPARTODET
            WHILE @MIND <= @MAXD
            BEGIN
                SELECT TOP 1
                       @IDPEDIDO = t.IDPEDIDO,
                       @IDVENDEDOR = t.IDVENDEDOR
                FROM @TBLREPARTODET t
                WHERE t.INDICE = @MIND;

                SET @insertquery = 'INSERT INTO OPENQUERY(' + @linkedserver + ', ''';
                SET @tsql = 'SELECT idReparto, idEmpresa, idPedido, idVendedor FROM dbo.REPARTO_DET'')';
                SET @tsql2
                    = ' SELECT ' + CAST(@IDREPARTO AS VARCHAR(20)) + ', ' + CAST(@IDEMPRESA AS VARCHAR(10)) + ', '
                      + CAST(@IDPEDIDO AS VARCHAR(10)) + ',' + CAST(@IDVENDEDOR AS VARCHAR(10)) + ';';

                EXEC (@insertquery + @tsql + @tsql2);
                --SELECT @insertquery + @tsql + @tsql2;

                SET @MIND = @MIND + 1;
            END;
        END;
        ELSE
        BEGIN
            INSERT INTO dbo.REPARTO_SYNC_LOG
            (
                codCia,
                Mensaje
            )
            VALUES
            (   @CODCIA,                                                          -- codCia - char(2)
                'idReparto ya existe en Cloud=>' + CAST(@IDREPARTO AS VARCHAR(10))
                + ' se procede a aislar el reparto para no volver a sincronizar.' -- Mensaje - varchar(500)
                );
        END;

        UPDATE dbo.REPARTO_CAB SET onCloud = 1 WHERE codCia = @CODCIA AND idReparto = @IDREPARTO

        DELETE FROM @TBLVALIDA;
        DELETE FROM @TBLREPARTODET;

        SET @MINC = @MINC + 1;
    END;


END TRY
BEGIN CATCH
    SET @EXITO = RTRIM(LTRIM(STR(ERROR_NUMBER()))) + '=' + ERROR_MESSAGE();
    --ROLLBACK TRAN
    GOTO Terminar;
END CATCH;


--IF @@TRANCOUNT > 0
--    COMMIT;

Terminar:
SELECT @EXITO;
/*
USP_REPARTO_SYNC
*/




GO