USE [geosan]
GO
/****** Object:  StoredProcedure [dbo].[atualizaConsumo]    Script Date: 02/24/2022 17:38:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Jos� Maria Villac Pinheiro
-- Create date: 28/05/2024
-- Description:	Atualiza a tabela com os dados so sistema comercial, com o objetivo do sistema rodar mais rapidamente
-- =============================================
CREATE PROCEDURE [dbo].[atualizaConsumo]
AS
BEGIN
	IF OBJECT_ID ('dbo.NXGS_V_LIG_COMERCIAL_CONSUMO') IS NOT NULL
	    BEGIN
			DELETE FROM [dbo].[NXGS_V_LIG_COMERCIAL_CONSUMO]
			INSERT INTO [dbo].[NXGS_V_LIG_COMERCIAL_CONSUMO]
				([NRO_LIGACAO],[NRO_LIGACAO_SEM_DV],[ANO], [MES], [CONSUMO_MEDIDO], [CONSUMO_FATURADO])
			SELECT CAST(CODLIGACAO+'0' AS INT) AS CODLIGACAO,CAST(CAST(CODLIGACAO AS INT) AS VARCHAR(20)) AS CODLIGACAO2, EXERCICIO, MES, CONSUMOMEDIDO, CONSUMOFATURADO 
			FROM
				OPENQUERY(FIREBIRD25, 
					'SELECT CODLIGACAO, EXERCICIO, MES, CONSUMOMEDIDO, CONSUMOFATURADO FROM VW_CONSUMOS');
	END
	ELSE
	    BEGIN
	    INSERT INTO [dbo].[NXGS_V_LIG_COMERCIAL_CONSUMO]
				([NRO_LIGACAO],[NRO_LIGACAO_SEM_DV],[ANO], [MES], [CONSUMO_MEDIDO], [CONSUMO_FATURADO])
			SELECT CAST(CODLIGACAO+'0' AS INT) AS CODLIGACAO,CAST(CAST(CODLIGACAO AS INT) AS VARCHAR(20)) AS CODLIGACAO2, EXERCICIO, MES, CONSUMOMEDIDO, CONSUMOFATURADO 
			FROM
				OPENQUERY(FIREBIRD25, 
					'SELECT CODLIGACAO, EXERCICIO, MES, CONSUMOMEDIDO, CONSUMOFATURADO FROM VW_CONSUMOS');
	    END
	SET NOCOUNT ON;
END