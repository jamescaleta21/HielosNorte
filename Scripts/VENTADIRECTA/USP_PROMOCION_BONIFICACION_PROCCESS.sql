IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE' -- Validación del tipo
          AND ROUTINE_SCHEMA = 'dbo' -- Validación del esquema
          AND s.ROUTINE_NAME = 'USP_PROMOCION_BONIFICACION_PROCCESS'
) -- Validación del nombre
BEGIN
    DROP PROC [dbo].[USP_PROMOCION_BONIFICACION_PROCCESS];
END;
GO
/*
USP_PROMOCION_BONIFICACION_PROCCESS '01',1007,'AGUA FIEL PAQUETE 650 ML - 15 UNID.','<r><d idprecio="1" ini="1" fin="10" pre="7" /></r>','<r><d idboni="1649" cant="10" boni="1" pre="0" /></r>'
*/
CREATE PROCEDURE [dbo].[USP_PROMOCION_BONIFICACION_PROCCESS]
    @CODCIA CHAR(2),
    @IDPRODUCTO BIGINT,
    @PRODUCTO VARCHAR(100),
    @CURRENTUSER VARCHAR(20),
    @xPROMOCION XML,
    @xBONIFICACION XML = NULL
WITH ENCRYPTION
AS
SET NOCOUNT ON;

DECLARE @IDOC INT;
DECLARE @EXITO VARCHAR(300);

SET @EXITO = '0=Datos Almacenados correctamente';

DECLARE @TBLPROMOCION TABLE
(
    idprecio INT,
    ini INT,
    fin INT,
    pre MONEY,
    fini DATE,
    ffin DATE
);
DECLARE @TBLBONIFICACION TABLE
(
    idboni BIGINT,
    cant INT,
    boni INT,
    pre MONEY,
    tope INT,
    fini DATE,
    ffin DATE
);
DECLARE @RESULTADOS TABLE
(
    ACCION VARCHAR(50),
    CODCIA CHAR(2),
    IDPRODUCTO BIGINT,
    IDPRECIO INT,
    DESCRIPCION VARCHAR(200),
    RNG_INI INT,
    RNG_FIN INT,
    PRECIO MONEY,
    CU_REGISTERED VARCHAR(20),
    FE_REGISTERED DATETIME,
    CU_UPDATED VARCHAR(20),
    FE_UPDATED DATETIME,
	FECINI DATE,
	FECFIN DATE,
    CODCIA_D CHAR(2),
    IDPRODUCTO_D BIGINT,
    IDPRECIO_D INT,
    DESCRIPCION_D VARCHAR(200),
    RNG_INI_D INT,
    RNG_FIN_D INT,
    PRECIO_D MONEY,
    CU_REGISTERED_D VARCHAR(20),
    FE_REGISTERED_D DATETIME,
    CU_UPDATED_D VARCHAR(20),
    FE_UPDATED_D DATETIME,
    FECINI_D DATE,
    FECFIN_D DATE
);
DECLARE @RESULTADOSBONI TABLE
(
    ACCION VARCHAR(50),
    CODCIA CHAR(2),
    IDPRODUCTO BIGINT,
    IDBONIFICACION BIGINT,
    CANTIDAD INT,
    BONIFICACION INT,
    PRECIO MONEY,
    TOPE INT,
    CU_REGISTERED VARCHAR(20),
    FE_REGISTERED DATETIME,
    CU_UPDATED VARCHAR(20),
    FE_UPDATED DATETIME,
	FECINI DATE,
	FECFIN DATE,
    CODCIA_D CHAR(2),
    IDPRODUCTO_D BIGINT,
    IDBONIFICACION_D BIGINT,
    CANTIDAD_D INT,
    BONIFICACION_D INT,
    PRECIO_D MONEY,
    TOPE_D INT,
    CU_REGISTERED_D VARCHAR(20),
    FE_REGISTERED_D DATETIME,
    CU_UPDATED_D VARCHAR(20),
    FE_UPDATED_D DATETIME,
    FECINI_D DATE,
    FECFIN_D DATE
);

--ALMACENANDO PROMOCIONES
EXEC sp_xml_preparedocument @IDOC OUTPUT, @xPROMOCION;
INSERT INTO @TBLPROMOCION
SELECT *
FROM
    OPENXML(@IDOC, '/r/d', 1)
    WITH
    (
        idprecio INT,
        ini INT,
        fin INT,
        pre MONEY,
        fini DATE,
        ffin DATE
    );

EXEC sp_xml_removedocument @IDOC;

--ALMACENANDO BONIFICACIONES
EXEC sp_xml_preparedocument @IDOC OUTPUT, @xBONIFICACION;
INSERT INTO @TBLBONIFICACION
SELECT *
FROM
    OPENXML(@IDOC, '/r/d', 1)
    WITH
    (
        idboni BIGINT,
        cant INT,
        boni INT,
        pre MONEY,
        tope INT,
        fini DATE,
        ffin DATE
    );

EXEC sp_xml_removedocument @IDOC;

BEGIN TRAN;
BEGIN TRY
    MERGE dbo.PRODUCTO_PROMOCION AS tgt
    USING @TBLPROMOCION AS src
    ON (
           tgt.IDPRECIO = src.idprecio
           AND tgt.IDPRODUCTO = @IDPRODUCTO
       )
    WHEN NOT MATCHED THEN
        INSERT
        (
            CODCIA,
            IDPRODUCTO,
            IDPRECIO,
            RNG_INI,
            RNG_FIN,
            PRECIO,
            DESCRIPCION,
            CU_REGISTERED,
            FECINI,
            FECFIN
        )
        VALUES
        (   @CODCIA, @IDPRODUCTO, src.idprecio, src.ini, CASE
                                                             WHEN src.fin = 0 THEN
                                                                 NULL
                                                             ELSE
                                                                 src.fin
                                                         END, src.pre, @PRODUCTO, @CURRENTUSER, src.fini, src.ffin)
    WHEN MATCHED AND tgt.CODCIA = @CODCIA
                     AND
                     (
                         tgt.RNG_INI <> src.ini
                         OR tgt.RNG_FIN <> CASE
                                               WHEN src.fin = 0 THEN
                                                   NULL
                                               ELSE
                                                   src.fin
                                           END
                         OR tgt.PRECIO <> src.pre
                         OR ISNULL(tgt.FECINI,'') <> ISNULL(src.fini,'')
                         OR ISNULL(tgt.FECFIN,'') <> ISNULL(src.ffin,'')
                     ) THEN
        UPDATE SET tgt.RNG_INI = src.ini,
                   tgt.RNG_FIN = CASE
                                     WHEN src.fin = 0 THEN
                                         NULL
                                     ELSE
                                         src.fin
                                 END,
                   tgt.PRECIO = src.pre,
                   tgt.FE_UPDATED = GETDATE(),
                   tgt.CU_UPDATED = @CURRENTUSER,
                   tgt.FECINI = src.fini,
                   tgt.FECFIN = src.ffin
    WHEN NOT MATCHED BY SOURCE AND tgt.CODCIA = @CODCIA
                                   AND tgt.IDPRODUCTO = @IDPRODUCTO THEN
        DELETE
    OUTPUT $action,
           Inserted.*,
           Deleted.*
    INTO @RESULTADOS
    --OUTPUT $action,Deleted.*
    ;

    --GRABANDO EN LOG
    DECLARE @IDLOG BIGINT;

    SELECT TOP 1
           @IDLOG = apl.IDLOG
    FROM dbo.PRODUCTO_PROMOCION_LOG apl
    WHERE apl.CODCIA = @CODCIA
          AND apl.IDPRODUCTO = @IDPRODUCTO
    ORDER BY apl.IDLOG DESC;

    IF @IDLOG IS NULL
    BEGIN
        SET @IDLOG = 0;
    END;

    SET @IDLOG = @IDLOG + 1;

    INSERT INTO dbo.PRODUCTO_PROMOCION_LOG
    (
        CODCIA,
        IDPRODUCTO,
        IDPRECIO,
        IDLOG,
        DESCRIPCION,
        RNG_INI,
        RNG_FIN,
        PRECIO,
        CU_REGISTERED,
        FE_REGISTERED,
        CU_UPDATED,
        FE_UPDATED,
        FECINI,
        FECFIN
    )
    SELECT r.CODCIA,
           r.IDPRODUCTO,
           r.IDPRECIO,
           @IDLOG,
           r.DESCRIPCION,
           r.RNG_INI,
           r.RNG_FIN,
           r.PRECIO,
           r.CU_REGISTERED,
           r.FE_REGISTERED,
           r.CU_UPDATED,
           r.FE_UPDATED,
           r.FECINI,
           r.FECFIN
    FROM @RESULTADOS r
    WHERE r.ACCION = 'INSERT';


    INSERT INTO dbo.PRODUCTO_PROMOCION_LOG
    (
        CODCIA,
        IDPRODUCTO,
        IDPRECIO,
        IDLOG,
        DESCRIPCION,
        RNG_INI,
        RNG_FIN,
        PRECIO,
        CU_REGISTERED,
        FE_REGISTERED,
        CU_UPDATED,
        FE_UPDATED,
        FECINI,
        FECFIN
    )
    SELECT r.CODCIA,
           r.IDPRODUCTO,
           r.IDPRECIO,
           @IDLOG,
           r.DESCRIPCION,
           r.RNG_INI,
           r.RNG_FIN,
           r.PRECIO,
           r.CU_REGISTERED,
           r.FE_REGISTERED,
           r.CU_UPDATED,
           r.FE_UPDATED,
           r.FECINI,
           r.FECFIN
    FROM @RESULTADOS r
    WHERE r.ACCION = 'UPDATE';

    INSERT INTO dbo.PRODUCTO_PROMOCION_LOG
    (
        CODCIA,
        IDPRODUCTO,
        IDPRECIO,
        IDLOG,
        DESCRIPCION,
        RNG_INI,
        RNG_FIN,
        PRECIO,
        CU_REGISTERED,
        FE_REGISTERED,
        CU_UPDATED,
        FE_UPDATED,
        CU_REMOVED,
        FE_REMOVED,
        FECINI,
        FECFIN
    )
    SELECT r.CODCIA_D,
           r.IDPRODUCTO_D,
           r.IDPRECIO_D,
           @IDLOG,
           r.DESCRIPCION_D,
           r.RNG_INI_D,
           r.RNG_FIN_D,
           r.PRECIO_D,
           r.CU_REGISTERED_D,
           r.FE_REGISTERED_D,
           r.CU_UPDATED_D,
           r.FE_UPDATED_D,
           @CURRENTUSER,
           GETDATE(),
           r.FECINI_D,
           r.FECFIN_D
    FROM @RESULTADOS r
    WHERE r.ACCION = 'DELETE';

    --SELECT * FROM @TBLBONIFICACION t
    --SELECT * FROM dbo.PRODUCTO_BONIFICACION ab

    MERGE dbo.PRODUCTO_BONIFICACION tgt
    USING @TBLBONIFICACION src
    ON (
           tgt.CODCIA = @CODCIA
           AND tgt.IDPRODUCTO = @IDPRODUCTO
           AND tgt.IDBONIFICACION = src.idboni
       )
    WHEN NOT MATCHED THEN
        INSERT
        (
            CODCIA,
            IDPRODUCTO,
            IDBONIFICACION,
            CANTIDAD,
            BONIFICACION,
            PRECIO,
            TOPE,
            CU_REGISTERED,
            FECINI,
            FECFIN
        )
        VALUES
        (   @CODCIA, @IDPRODUCTO, src.idboni, src.cant, src.boni, src.pre, CASE
                                                                               WHEN src.tope = 0 THEN
                                                                                   NULL
                                                                               ELSE
                                                                                   src.tope
                                                                           END, @CURRENTUSER, src.fini, src.ffin)
    WHEN MATCHED AND tgt.CODCIA = @CODCIA
                     AND tgt.IDPRODUCTO = @IDPRODUCTO
                     AND
                     (
                         tgt.CANTIDAD <> src.cant
                         OR tgt.BONIFICACION <> src.boni
						 OR ISNULL(tgt.FECINI,'') <> ISNULL(src.fini,'')
						 OR ISNULL(tgt.FECFIN,'') <> ISNULL(src.ffin,'')
                     ) THEN
        UPDATE SET tgt.IDBONIFICACION = src.idboni,
                   tgt.CANTIDAD = src.cant,
                   tgt.BONIFICACION = src.boni,
                   tgt.CU_UPDATED = @CURRENTUSER,
                   tgt.FE_UPDATED = GETDATE(),
                   tgt.TOPE = src.tope,
                   tgt.FECINI = src.fini,
                   tgt.FECFIN = src.ffin
    WHEN NOT MATCHED BY SOURCE AND tgt.IDPRODUCTO = @IDPRODUCTO
                                   AND tgt.CODCIA = @CODCIA THEN
        DELETE
    OUTPUT $action,
           Inserted.*,
           Deleted.*
    INTO @RESULTADOSBONI;

    --GRABANDO EN LOG
    SET @IDLOG = NULL;

    SELECT TOP 1
           @IDLOG = apl.IDLOG
    FROM dbo.PRODUCTO_BONIFICACION_LOG apl
    WHERE apl.CODCIA = @CODCIA
          AND apl.IDPRODUCTO = @IDPRODUCTO
    ORDER BY apl.IDLOG DESC;

    IF @IDLOG IS NULL
    BEGIN
        SET @IDLOG = 0;
    END;

    SET @IDLOG = @IDLOG + 1;

    INSERT INTO dbo.PRODUCTO_BONIFICACION_LOG
    (
        CODCIA,
        IDPRODUCTO,
        IDBONIFICACION,
        IDLOG,
        CANTIDAD,
        BONIFICACION,
        PRECIO,
        TOPE,
        CU_REGISTERED,
        FE_REGISTERED,
        CU_UPDATED,
        FE_UPDATED,
        FECINI,
        FECFIN
    )
    SELECT r.CODCIA,
           r.IDPRODUCTO,
           r.IDBONIFICACION,
           @IDLOG,
           r.CANTIDAD,
           r.BONIFICACION,
           r.PRECIO,
           r.TOPE,
           r.CU_REGISTERED,
           r.FE_REGISTERED,
           r.CU_UPDATED,
           r.FE_UPDATED,
           r.FECINI,
           r.FECFIN
    FROM @RESULTADOSBONI r
    WHERE r.ACCION = 'INSERT';

    INSERT INTO dbo.PRODUCTO_BONIFICACION_LOG
    (
        CODCIA,
        IDPRODUCTO,
        IDBONIFICACION,
        IDLOG,
        CANTIDAD,
        BONIFICACION,
        PRECIO,
        TOPE,
        CU_REGISTERED,
        FE_REGISTERED,
        CU_UPDATED,
        FE_UPDATED,
        FECINI,
        FECFIN
    )
    SELECT r.CODCIA,
           r.IDPRODUCTO,
           r.IDBONIFICACION,
           @IDLOG,
           r.CANTIDAD,
           r.BONIFICACION,
           r.PRECIO,
           r.TOPE,
           r.CU_REGISTERED,
           r.FE_REGISTERED,
           r.CU_UPDATED,
           r.FE_UPDATED,
           r.FECINI,
           r.FECFIN
    FROM @RESULTADOSBONI r
    WHERE r.ACCION = 'UPDATE';

    INSERT INTO dbo.PRODUCTO_BONIFICACION_LOG
    (
        CODCIA,
        IDPRODUCTO,
        IDBONIFICACION,
        IDLOG,
        CANTIDAD,
        BONIFICACION,
        PRECIO,
        TOPE,
        CU_REGISTERED,
        FE_REGISTERED,
        CU_UPDATED,
        FE_UPDATED,
        CU_REMOVED,
        FE_REMOVED,
        FECINI,
        FECFIN
    )
    SELECT r.CODCIA_D,
           r.IDPRODUCTO_D,
           r.IDBONIFICACION_D,
           @IDLOG,
           r.CANTIDAD_D,
           r.BONIFICACION_D,
           r.PRECIO_D,
           r.TOPE_D,
           r.CU_REGISTERED_D,
           r.FE_REGISTERED_D,
           r.CU_UPDATED_D,
           r.FE_UPDATED_D,
           @CURRENTUSER,
           GETDATE(),
           r.FECINI_D,
           r.FECFIN_D
    FROM @RESULTADOSBONI r
    WHERE r.ACCION = 'DELETE';
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