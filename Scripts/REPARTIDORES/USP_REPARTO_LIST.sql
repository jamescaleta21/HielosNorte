IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_REPARTO_LIST'
)
BEGIN
    DROP PROC [dbo].[USP_REPARTO_LIST];
END;
GO
/*
USP_REPARTO_LIST '01','20241030'
*/
CREATE PROCEDURE [dbo].[USP_REPARTO_LIST]
    @CODCIA CHAR(2),
    @FECHA DATE
WITH ENCRYPTION
AS
SET NOCOUNT ON;

SELECT rc.idReparto,
       'REPARTO NRO ' + CAST(rc.idReparto AS VARCHAR(20)) AS 'reparto',
       RTRIM(LTRIM(v.VEM_NOMBRE)) AS 'repartidor',
       CONVERT(VARCHAR(10), rc.fecha, 103) AS 'fecha',
	   CASE WHEN ISNULL(rc.onCloud,0) = 0 THEN 'NO'ELSE 'SI' END AS 'cloud',
	   COALESCE(rc.observaciones,'') AS 'obs'
FROM dbo.REPARTO_CAB rc WITH (NOLOCK)
    INNER JOIN dbo.VEMAEST v WITH (NOLOCK)
        ON rc.codCia = v.VEM_CODCIA
           AND rc.idRepartidor = v.VEM_CODVEN
WHERE rc.codCia = @CODCIA AND rc.eliminado = 0
      AND rc.fecha = @FECHA;
GO