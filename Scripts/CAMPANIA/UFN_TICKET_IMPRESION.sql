IF EXISTS ( SELECT TOP 1
                    S.SPECIFIC_NAME
            FROM    information_schema.routines s
            WHERE   s.ROUTINE_TYPE = 'FUNCTION'
                    AND S.ROUTINE_NAME = 'UFN_TICKET_IMPRESION' )
    BEGIN
        DROP FUNCTION [dbo].[UFN_TICKET_IMPRESION]
    END
GO
/*
SELECT * FROM UFN_TICKET_IMPRESION()
*/
CREATE FUNCTION [dbo].[UFN_TICKET_IMPRESION]
()
RETURNS TABLE
AS
RETURN
(
    SELECT t1.CODCIA,
           t1.FECHAEMISION,
           t1.TIPOCOMPROBANTE,
           t1.SERIECOMPROBANTE,
           t1.NUMEROCOMPROBANTE,
           t1.CODIGOCLIENTE,
           STUFF(
           (
               SELECT ',' + t2.NUMEROTICKET
               FROM TICKET t2
               WHERE t2.CODCIA = t1.CODCIA
                     AND t2.TIPOCOMPROBANTE = t1.TIPOCOMPROBANTE
                     AND t2.SERIECOMPROBANTE = t1.SERIECOMPROBANTE
                     AND t2.NUMEROCOMPROBANTE = t1.NUMEROCOMPROBANTE
                     AND t2.CODIGOCLIENTE = t1.CODIGOCLIENTE
               FOR XML PATH('')
           ),
           1,
           1,
           ''
                ) AS listaTickets
    FROM TICKET t1
    GROUP BY t1.CODCIA,
             t1.FECHAEMISION,
             t1.TIPOCOMPROBANTE,
             t1.SERIECOMPROBANTE,
             t1.NUMEROCOMPROBANTE,
             t1.CODIGOCLIENTE
);
