USE [geosan]
GO
/****** Object:  StoredProcedure [dbo].[atualizaConsumoMedio]    Script Date: 02/24/2022 17:38:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		José Maria Villac Pinheiro
-- Create date: 24/02/2022
-- Description:	Atualiza a tabela com os dados so sistema comercial, com o objetivo do sistema rodar maisr rapidamente
-- =============================================
ALTER PROCEDURE [dbo].[atualizaConsumoMedio]
AS
BEGIN
	IF OBJECT_ID ('dbo.NXGS_V_LIG_COM_CONSUMO_MEDIO') IS NOT NULL
	    BEGIN
	        DROP TABLE [dbo].[NXGS_V_LIG_COM_CONSUMO_MEDIO];
	        SELECT * INTO [dbo].[NXGS_V_LIG_COM_CONSUMO_MEDIO] FROM [dbo].[NXGS_V_LIG_COM_CONSUMO_MEDIO_CONEXAO];
	    END
	ELSE
	    BEGIN
	        SELECT * INTO [dbo].[NXGS_V_LIG_COM_CONSUMO_MEDIO] FROM [dbo].[NXGS_V_LIG_COM_CONSUMO_MEDIO_CONEXAO];
	    END
	SET NOCOUNT ON;
END