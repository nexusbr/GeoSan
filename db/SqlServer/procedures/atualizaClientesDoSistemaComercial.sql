USE [geosan]
GO
/****** Object:  StoredProcedure [dbo].[atualizaClientesDoSistemaComercial]    Script Date: 02/24/2022 17:33:50 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		José Maria Villac Pinheiro
-- Create date: 24/02/2022
-- Description:	Atualiza a tabela com os dados so sistema comercial, com o objetivo do sistema rodar maisr rapidamente
-- NXGS_V_LIG_COMERCIAL_CONEXAO - é a vista conectada diretamente ao sistema comercial
-- =============================================
ALTER PROCEDURE [dbo].[atualizaClientesDoSistemaComercial]
AS
BEGIN
	IF OBJECT_ID ('dbo.NXGS_V_LIG_COMERCIAL') IS NOT NULL
	    BEGIN
	        DROP TABLE [dbo].[NXGS_V_LIG_COMERCIAL];
	        SELECT * INTO [dbo].[NXGS_V_LIG_COMERCIAL] FROM [dbo].[NXGS_V_LIG_COMERCIAL_CONEXAO];
	    END
	ELSE
	    BEGIN
	        SELECT * INTO [dbo].[NXGS_V_LIG_COMERCIAL] FROM [dbo].[NXGS_V_LIG_COMERCIAL_CONEXAO];
	    END
	SET NOCOUNT ON;
END
