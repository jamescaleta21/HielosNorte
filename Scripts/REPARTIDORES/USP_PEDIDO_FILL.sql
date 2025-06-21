IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_PEDIDO_FILL'
)
BEGIN
    DROP PROC [dbo].[USP_PEDIDO_FILL];
END;
GO
/*
USP_PEDIDO_FILL '01',51
*/
CREATE PROCEDURE [dbo].[USP_PEDIDO_FILL]
    @CODCIA CHAR(2),
    @IDPEDIDO INT
WITH ENCRYPTION
AS
SET NOCOUNT ON;


SELECT p.idpedido AS ide,
       p.idcliente AS idcliente,
       RTRIM(LTRIM(c.CLI_NOMBRE)) AS cliente,
       p.fecha AS fecha,
       p.total AS total,
       p.idvendedor AS idven,
       RTRIM(LTRIM(v.VEM_NOMBRE)) AS vendedor
	   ,COALESCE(p.observacion,'') AS obs
	   ,COALESCE(c.CLI_CASA_DIREC,'') AS 'dir'
FROM dbo.PEDIDO p
    INNER JOIN dbo.CLIENTES c
        ON p.idcliente = c.CLI_CODCLIE
           AND c.CLI_CODCIA = @CODCIA
    INNER JOIN dbo.VEMAEST v
        ON p.idvendedor = v.VEM_CODVEN
           AND v.VEM_CODCIA = @CODCIA
WHERE p.idpedido = @IDPEDIDO;

SELECT pd.cantidad AS cant,
       pd.idproducto AS ideproducto,
       RTRIM(LTRIM(a.ART_NOMBRE)) AS producto,
       pd.precio AS pre,
       pd.importe AS imp
FROM dbo.PEDIDO_DETALLE pd
    INNER JOIN dbo.ARTI a
        ON pd.idproducto = a.ART_KEY
           AND a.ART_CODCIA = @CODCIA
WHERE pd.idpedido = @IDPEDIDO
ORDER BY pd.secuencia;

/*
USP_PEDIDO_FILL '01',51
*/
GO