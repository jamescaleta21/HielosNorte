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
USP_REPARTO_SYNC '01',6
*/
CREATE PROCEDURE [dbo].[USP_REPARTO_SYNC]
    @CODCIA CHAR(2),
    @IDREPARTO INT
WITH ENCRYPTION
AS
SET NOCOUNT ON;

--DECLARE @CODCIA CHAR(2);
--SET @CODCIA = '01';

DECLARE @MINC INT,
        @MAXC INT;
DECLARE @EXITO VARCHAR(300);
SET @EXITO = '0=Repartos cerrados Correctamente.';
--VARIABLES CABECERA
DECLARE
    --@IDREPARTO INT,
    @FECHA DATE,
    @IDREPARTIDOR INT,
    @OBS VARCHAR(300),
    @PESO DECIMAL(16, 2);
--VARIABLES DETALLE
DECLARE @TBLREPARTODET TABLE
(
    INDICE INT IDENTITY,
    IDPEDIDO BIGINT,
    IDPEDIDOCLOUD BIGINT,
    IDVENDEDOR INT
);
DECLARE @IDPEDIDO BIGINT,@IDPEDIDOCLOUD BIGINT,
        @IDVENDEDOR INT;
DECLARE @MIND INT,
        @MAXD INT;
--TABLA PARA VALIDACION SI EXISTE IDREPARTO PREVIAMENTE
DECLARE @TBLVALIDA TABLE
(
    DATO CHAR(1)
);
DECLARE @TBLREPARTOCANT TABLE(NUMERO INT);


DECLARE @TBLREPARTO TABLE
(
    INDICE INT IDENTITY,
    IDREPARTO INT,
    FECHA DATE
);


IF
(
    SELECT COUNT(rc.idReparto)
    FROM dbo.REPARTO_CAB rc
    WHERE ISNULL(rc.onCloud, 0) = 0
          AND rc.codCia = @CODCIA
          AND rc.idReparto = @IDREPARTO
) = 0
BEGIN
    SET @EXITO = '-1=No hay nada que enviar a Cloud.';
    GOTO terminar;
END;



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
    WHERE ISNULL(rc.onCloud, 0) = 0
          AND ISNULL(rc.eliminado, 0) = 0
          AND rc.codCia = @CODCIA
          AND rc.idReparto = @IDREPARTO;

    --RECUPERANDO LOS VALORES MINIMO Y MAXIMO PARA REALIZAR EL RECORRIDO
    SELECT @MINC = MIN(t.INDICE),@MAXC = MAX(t.INDICE)
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

     

            --SUBIENDO TABLA REPARTO_DET
            INSERT INTO @TBLREPARTODET
            (
                IDPEDIDO,
                IDPEDIDOCLOUD,
                IDVENDEDOR
            )
            SELECT p.idpedido, p.idpedido_cloud,
                   rd.idVendedor
            FROM dbo.REPARTO_DET rd WITH (NOLOCK)
                INNER JOIN dbo.PEDIDO p WITH (NOLOCK)
                    ON rd.idPedido = p.idpedido
                       AND rd.codCia = @CODCIA
            WHERE rd.codCia = @CODCIA
                  AND rd.idReparto = @IDREPARTO AND COALESCE(rd.onCloud,0) = 0;

            SELECT @MIND = MIN(t.INDICE),@MAXD = MAX(t.INDICE)
            FROM @TBLREPARTODET t;
            
           

    /*
USP_REPARTO_SYNC '01',6
*/
            --SELECT * FROM @TBLREPARTODET
            WHILE @MIND <= @MAXD
            BEGIN
    
                SELECT TOP 1
                       @IDPEDIDO = t.IDPEDIDO,
                       @IDPEDIDOCLOUD = t.IDPEDIDOCLOUD,
                       @IDVENDEDOR = t.IDVENDEDOR
                FROM @TBLREPARTODET t
                WHERE t.INDICE = @MIND;

                SET @insertquery = 'INSERT INTO OPENQUERY(' + @linkedserver + ', ''';
                SET @tsql = 'SELECT idReparto, idEmpresa, idPedido, idVendedor FROM dbo.REPARTO_DET'')';
                SET @tsql2
                    = ' SELECT ' + CAST(@IDREPARTO AS VARCHAR(20)) + ', ' + CAST(@IDEMPRESA AS VARCHAR(10)) + ', '
                      + CAST(@IDPEDIDOCLOUD AS VARCHAR(10)) + ',' + CAST(@IDVENDEDOR AS VARCHAR(10)) + ';';

                EXEC (@insertquery + @tsql + @tsql2);
                --SELECT @insertquery + @tsql + @tsql2;
                
               

				UPDATE dbo.REPARTO_DET SET onCloud = 1 WHERE codCia = @CODCIA AND idReparto = @IDREPARTO AND idPedido = @IDPEDIDO;
			

                SET @MIND = @MIND + 1;
            END;
			---compara
			


			   --VALIDANDO SI EL REPARTO SE ENCUENTRA EN CLOUD
			SET @tsql
				= 'SELECT COUNT(rd.idVendedor) FROM dbo.REPARTO_DET rd with (nolock) where rd.idReparto = '
				  + CAST(@IDREPARTO AS VARCHAR(10)) + ' and rd.idEmpresa = ' + CAST(@IDEMPRESA AS VARCHAR(10)) + ''');';

		 INSERT INTO @TBLREPARTOCANT
		 (
			 NUMERO
		 )
			EXEC (@openquery + @tsql);

			IF (SELECT COUNT(rd.idVendedor) FROM dbo.REPARTO_DET rd WHERE rd.codCia = @CODCIA AND rd.idReparto = @IDREPARTO) = (SELECT TOP 1 t.NUMERO FROM @TBLREPARTOCANT t)
			BEGIN

			  SET @insertquery = 'INSERT INTO OPENQUERY(' + @linkedserver + ', ''';
            SET @tsql
                = 'SELECT idReparto, idEmpresa, fecha, idRepartidor, observaciones, peso FROM dbo.REPARTO_CAB'')';
            SET @tsql2
                = ' SELECT ' + CAST(@IDREPARTO AS VARCHAR(20)) + ', ' + CAST(@IDEMPRESA AS VARCHAR(10)) + ', '''
                  + CONVERT(VARCHAR(8), @FECHA, 112) + ''',' + CAST(@IDREPARTIDOR AS VARCHAR(10)) + ','''
                  + RTRIM(LTRIM(@OBS)) + ''',' + CAST(@PESO AS VARCHAR(20)) + ';';

            EXEC (@insertquery + @tsql + @tsql2);


				UPDATE dbo.REPARTO_CAB
				SET onCloud = 1
				WHERE codCia = @CODCIA
				AND idReparto = @IDREPARTO;
				
				
				 DELETE FROM @TBLVALIDA;
        DELETE FROM @TBLREPARTODET;
        delete from @TBLREPARTOCANT
		

        END;
 

       

       

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