INSERT INTO dbo.PERFIL
(
    IDPERFIL,
    DENOMINACION,
    ACTIVO,
    FE_REGISTER,
    CU_REGISTER
   
)
VALUES
(   4,       -- IDPERFIL - int
    'Cliente',      -- DENOMINACION - varchar(50)
    1,    -- ACTIVO - bit
    GETDATE(), -- FE_REGISTER - datetime
    'SYSTEM'      -- CU_REGISTER - varchar(20)

    )
INSERT INTO dbo.PERFIL
(
    IDPERFIL,
    DENOMINACION,
    ACTIVO,
    FE_REGISTER,
    CU_REGISTER
   
)
VALUES
(   5,       -- IDPERFIL - int
    'Venta Directa',      -- DENOMINACION - varchar(50)
    1,    -- ACTIVO - bit
    GETDATE(), -- FE_REGISTER - datetime
    'SYSTEM'      -- CU_REGISTER - varchar(20)

    )