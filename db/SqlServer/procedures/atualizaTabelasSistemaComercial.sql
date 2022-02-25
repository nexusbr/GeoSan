USE [geosan]
GO
/****** Object:  StoredProcedure [dbo].[atualizaTabelasSistemaComercial]    Script Date: 02/24/2022 23:02:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		José Maria Villac Pinheiro
-- Create date: 24/02/2022
-- Description:	Atualiza as tabelas com os dados so sistema comercial, com o objetivo do sistema rodar maisr rapidamente
-- =============================================
ALTER PROCEDURE [dbo].[atualizaTabelasSistemaComercial]
AS
BEGIN
	EXEC dbo.atualizaClientesDoSistemaComercial;
	EXEC dbo.atualizaConsumoMedio;
	EXEC dbo.atualizaConsumo;
	SET NOCOUNT ON;
END