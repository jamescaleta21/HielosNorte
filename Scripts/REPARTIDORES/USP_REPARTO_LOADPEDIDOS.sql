IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_REPARTO_LOADPEDIDOS'
)
BEGIN
    DROP PROC [dbo].[USP_REPARTO_LOADPEDIDOS];
END;
GO
/*
USP_REPARTO_LOADPEDIDOS '20241030',5,'01'
*/
CREATE PROCEDURE [dbo].[USP_REPARTO_LOADPEDIDOS]
    @FECHA DATE,
    @IDVENDEDOR INT,
    @CODCIA CHAR(2)
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT p.idpedido,
       RTRIM(LTRIM(c.CLI_NOMBRE)) AS 'nombre',
       --p.total AS 'total'
	   SUM(p2.PRE_PESO * pd.cantidad) AS 'peso',
	   COALESCE(p.observacion,'') AS 'obs'
FROM dbo.PEDIDO p WITH (NOLOCK)
INNER JOIN dbo.PEDIDO_DETALLE pd WITH (NOLOCK) ON p.idpedido = pd.idpedido
	INNER JOIN dbo.PRECIOS p2 WITH (NOLOCK) ON pd.idproducto = p2.PRE_CODART AND p2.PRE_CODCIA = @CODCIA
    INNER JOIN dbo.CLIENTES c WITH (NOLOCK)
        ON p.idcliente = c.CLI_CODCLIE
           AND c.CLI_CODCIA = @CODCIA
WHERE p.fecha = @FECHA
      AND p.idvendedor = @IDVENDEDOR AND p.idRepartidor IS null
	  GROUP BY c.CLI_NOMBRE,p.idpedido,p.observacion
	  ;

GO