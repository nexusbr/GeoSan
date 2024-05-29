USE [geosan]
GO
/****** Object:  StoredProcedure [dbo].[atualizaConsumoMedio]    Script Date: 02/24/2022 17:38:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		José Maria Villac Pinheiro
-- Create date: 28/05/2024
-- Description:	Atualiza a tabela com os dados so sistema comercial, com o objetivo do sistema rodar maisr rapidamente
-- =============================================
CREATE PROCEDURE [dbo].[atualizaConsumoMedio]
AS
BEGIN
	IF OBJECT_ID ('dbo.NXGS_V_LIG_COM_CONSUMO_MEDIO') IS NOT NULL
	    BEGIN
			DELETE FROM [dbo].[NXGS_V_LIG_COM_CONSUMO_MEDIO]
	        INSERT INTO [dbo].[NXGS_V_LIG_COM_CONSUMO_MEDIO]
						([NRO_LIGACAO],[NRO_LIGACAO_SEM_DV],[CONSUMO_MEDIO])
	        SELECT CAST(CODLIGACAO+'0' AS INT) AS CODLIGACAO,CAST(CAST(CODLIGACAO AS INT) AS VARCHAR(20)) AS CODLIGACAO2, MEDIA 
	        FROM
						OPENQUERY(FIREBIRD25, 
							'SELECT CODLIGACAO, MEDIA FROM VW_CONSUMOS where EXERCICIO = 2024 and MES = 3');
	    END
	ELSE
	    BEGIN
	    	INSERT INTO [dbo].[NXGS_V_LIG_COM_CONSUMO_MEDIO]
						([NRO_LIGACAO],[NRO_LIGACAO_SEM_DV],[CONSUMO_MEDIO])
	        SELECT CAST(CODLIGACAO+'0' AS INT) AS CODLIGACAO,CAST(CAST(CODLIGACAO AS INT) AS VARCHAR(20)) AS CODLIGACAO2, MEDIA 
	        FROM
						OPENQUERY(FIREBIRD25, 
							'SELECT CODLIGACAO, MEDIA FROM VW_CONSUMOS where EXERCICIO = 2024 and MES = 3');
	    END
	SET NOCOUNT ON;
END