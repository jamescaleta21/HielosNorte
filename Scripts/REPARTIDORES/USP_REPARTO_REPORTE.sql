IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_REPARTO_REPORTE'
)
BEGIN
    DROP PROC [dbo].[USP_REPARTO_REPORTE];
END;
GO
/*
exec [dbo].[USP_REPARTO_REPORTE] '01','20241030',9
select * from pedido
*/
CREATE PROCEDURE [dbo].[USP_REPARTO_REPORTE]
    @CODCIA CHAR(2),
    @FECHA DATE,
    @IDREPARTO INT
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT COUNT(p.idpedido) AS cant,
       SUM(p.total) AS total
FROM PEDIDO p
    INNER JOIN dbo.REPARTO_DET rd
        ON rd.codCia = @CODCIA
           AND rd.idPedido = p.idpedido
    INNER JOIN dbo.REPARTO_CAB rc
        ON rd.idReparto = rc.idReparto
           AND rd.codCia = rc.codCia
WHERE rd.idReparto = @IDREPARTO
      AND COALESCE(p.ANULADO, 0) = 0
      AND rc.fecha = @FECHA;

SELECT COALESCE(CAST(c.CLI_PRECIOS AS VARCHAR(10)), '') AS orden,
       p.idpedido AS ide,
       RTRIM(LTRIM(c.CLI_NOMBRE)) + CHAR(13) + c.CLI_CASA_DIREC + CHAR(13) + CAST(c.CLI_CODCLIE AS VARCHAR(20))
       + ' HORA: ' + COALESCE(dbo.UFN_DEVUELVE_HORA(p.fe_register_cloud), '') --(@#)2-A
       + CHAR(13) + RTRIM(LTRIM(p.observacion)) + '-' + 'ZONA:' + RTRIM(LTRIM(t.TAB_NOMLARGO)) AS 'cli',
       RTRIM(LTRIM(a.ART_NOMBRE)) AS prod,
       --SUM(pd.cantidad) AS cant,
       SUM(rdi.cantidad) AS cant,
       SUM(pd.importe) AS imp,
       RTRIM(LTRIM(v.VEM_NOMBRE)) AS ven,
       RTRIM(LTRIM(c.CLI_PRENDA)) AS dia,
       RTRIM(LTRIM(p.IDSUBRUTA)) AS subruta
FROM dbo.REPARTO_CAB rc WITH (NOLOCK)
    INNER JOIN dbo.REPARTO_DET rd WITH (NOLOCK)
        ON rd.codCia = rc.codCia
           AND rd.idReparto = rc.idReparto
    INNER JOIN dbo.REPARTO_DET_ITEM rdi WITH (NOLOCK)
        ON rd.codCia = rdi.codCia
           AND rd.idReparto = rdi.idReparto
           AND rd.idPedido = rdi.idPedido
    INNER JOIN dbo.ARTI a WITH (NOLOCK)
        ON rdi.codCia = a.ART_CODCIA
           AND rdi.idProducto = a.ART_KEY
    INNER JOIN dbo.PEDIDO p WITH (NOLOCK)
        ON rd.idPedido = p.idpedido
    INNER JOIN dbo.CLIENTES c WITH (NOLOCK)
        ON p.idcliente = c.CLI_CODCLIE
           AND c.CLI_CODCIA = @CODCIA
    INNER JOIN dbo.PEDIDO_DETALLE pd WITH (NOLOCK)
        ON rdi.codCia = @CODCIA
           AND pd.idpedido = rdi.idPedido
           AND pd.idproducto = rdi.idProducto
    INNER JOIN dbo.VEMAEST v
        ON rc.codCia = v.VEM_CODCIA
           AND v.VEM_CODVEN = rc.idRepartidor
    INNER JOIN TABLAS t WITH (NOLOCK)
        ON c.CLI_ZONA_NEW = t.TAB_NUMTAB
           AND t.TAB_TIPREG = 35
           AND t.TAB_CODCIA = '00'
WHERE rc.idReparto = @IDREPARTO
      --p.idvendedor = @IDVENDEDOR
      AND COALESCE(p.ANULADO, 0) = 0
      AND p.fecha = @FECHA
GROUP BY p.idpedido,
         c.CLI_NOMBRE,
         a.ART_NOMBRE,
         c.CLI_CASA_DIREC,
         c.CLI_CODCLIE,
         p.observacion,
         p.fe_register_cloud,
         COALESCE(CAST(c.CLI_PRECIOS AS VARCHAR(10)), ''),
         c.CLI_PRENDA,
         p.IDSUBRUTA,
         t.TAB_NOMLARGO,
         v.VEM_NOMBRE
ORDER BY p.idpedido;

--RESUMEN
declare @TBLZONAS TABLE(DESCRIPCION VARCHAR(1000),IDSUBRUTA TINYINT)
DECLARE @INI TINYINT, @FIN TINYINT, @STRFINAL VARCHAR(1000), @IDSUBRUTA INT
DECLARE @TBLDISTINCT TABLE(IDSUBRUTA TINYINT, DESCRIPCION VARCHAR(100), INDICE TINYINT IDENTITY)


SET @STRFINAL = ''


	INSERT INTO @TBLZONAS(DESCRIPCION,IDSUBRUTA)
	select distinct RTRIM(LTRIM(t.TAB_NOMLARGO)),p.IDSUBRUTA
	FROM PEDIDO p with (nolock)
	INNER JOIN dbo.CLIENTES c WITH (NOLOCK)
		ON p.idcliente = c.CLI_CODCLIE
		   AND c.CLI_CODCIA = @CODCIA
		 INNER JOIN TABLAS t WITH (NOLOCK) ON c.CLI_ZONA_NEW = t.TAB_NUMTAB AND t.TAB_TIPREG = 35 AND t.TAB_CODCIA = '00'      
		 INNER JOIN dbo.REPARTO_DET rd WITH (NOLOCK) ON RD.codCia = @CODCIA AND RD.idPedido = P.idpedido
	  WHERE RD.idReparto = @IDREPARTO
	  AND COALESCE(p.ANULADO, 0) = 0
	  AND p.fecha = @FECHA

	  /*
	  exec [dbo].[USP_REPARTO_REPORTE] '01','20241030',7
	  */
	 
              
     INSERT INTO @TBLDISTINCT(IDSUBRUTA)
     SELECT DISTINCT IDSUBRUTA FROM @TBLZONAS

	SELECT @INI = MIN(INDICE) FROM @TBLDISTINCT
	SELECT @FIN = mAX(INDICE) FROM @TBLDISTINCT

	WHILE @INI <= @FIN
	BEGIN
		SELECT TOP 1 @IDSUBRUTA = IDSUBRUTA FROM @TBLDISTINCT WHERE INDICE = @INI

		SELECT @STRFINAL = STUFF((SELECT ', ' + descripcion FROM @TBLZONAS WHERE IDSUBRUTA = @IDSUBRUTA
		FOR XML PATH('')),1,1, '')

		UPDATE @TBLDISTINCT SET DESCRIPCION = @STRFINAL WHERE INDICE = @INI

		SET @INI = @INI + 1
	END
              
    SELECT 
           RTRIM(LTRIM(a.ART_NOMBRE)) AS prod,
           SUM(pd.cantidad) AS cant,
           SUM(pd.importe) AS imp,
           p.fecha AS fecha,
           RTRIM(LTRIM(v.VEM_NOMBRE)) AS ven,
           RTRIM(ltrim(c.CLI_PRENDA)) as dia,
           RTRIM(LTRIM(p.IDSUBRUTA)) AS subruta,
           td.DESCRIPCION as zona
    FROM
	dbo.REPARTO_CAB rc WITH (NOLOCK) 
	INNER JOIN dbo.REPARTO_DET rd WITH (NOLOCK) ON rc.idReparto = rd.idReparto AND rc.codCia = rd.codCia
	INNER JOIN dbo.PEDIDO p WITH (NOLOCK) ON rd.idPedido = p.idpedido AND rd.codCia = @CODCIA
        INNER JOIN dbo.CLIENTES c WITH (NOLOCK)
            ON p.idcliente = c.CLI_CODCLIE
               AND c.CLI_CODCIA = @CODCIA
        INNER JOIN dbo.PEDIDO_DETALLE pd WITH (NOLOCK)
            ON pd.idpedido = p.idpedido
        INNER JOIN dbo.ARTI a WITH (NOLOCK)
            ON pd.idproducto = a.ART_KEY
               AND a.ART_CODCIA = @CODCIA
        INNER JOIN dbo.VEMAEST v WITH (NOLOCK)
            ON rc.idRepartidor = v.VEM_CODVEN
               AND v.VEM_CODCIA = @CODCIA   
        INNER JOIN TABLAS t WITH (NOLOCK) ON c.CLI_ZONA_NEW = t.TAB_NUMTAB AND t.TAB_TIPREG = 35 AND t.TAB_CODCIA = '00'           
        INNER JOIN @TBLDISTINCT td ON p.IDSUBRUTA = td.IDSUBRUTA
		
    WHERE RD.idReparto = @IDREPARTO
          AND COALESCE(p.ANULADO, 0) = 0
          AND p.fecha = @FECHA
		  AND rc.codCia = @CODCIA
    GROUP BY a.ART_NOMBRE,
             p.fecha
            ,v.VEM_NOMBRE
            ,c.CLI_PRENDA
            ,p.IDSUBRUTA
            ,t.TAB_NOMLARGO
            ,td.DESCRIPCION
GO