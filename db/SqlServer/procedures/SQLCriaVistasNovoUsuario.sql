USE [geosan]
GO
/****** Object:  StoredProcedure [dbo].[SQLCriaVistasNovoUsuario]    Script Date: 20/07/2022 12:16:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =================================================================
-- Author:		<José  Maria Villac Pinheiro>
-- Create date: <26/06/2022>
-- Description:	<Copia vistas de um usuário para novo usuário>
-- nomeUsuarioAtualUSRNom - nome do usuário de onde serão copiadas as vistas existentes
-- nomeNovoUsuarioUSRNom - nome do usuário para onde serão copiadas as vistas
-- permissaoDoUsuario - 1 - Administrador, 2 - Usuário, 3 - Visitante, 4 - Visualizador
-- =================================================================
ALTER PROCEDURE [dbo].[SQLCriaVistasNovoUsuario] @nomeUsuarioAtualUSRNom nvarchar(50) = NULL, @nomeNovoUsuarioUSRNom nvarchar(50) = NULL, @permissaoDoUsuario int = NULL
AS
BEGIN
	DECLARE
	@view_id int,
	@projection_id int,
	@name nvarchar(255),
	@visibility int,
	@lower_x float,
	@lower_y float,
	@upper_x float,
	@upper_y float,
	@view_id_daVistaRecemInserida int,

	@theme_id int,
	@layer_id int,
	@view_id_te_theme int,
	@name_te_theme varchar(255),
	@parent_id int,
	@priority int,
	@node_type int,
	@min_scale float,
	@max_scale float,
	@generate_attribute_where varchar(255),
	@generate_spatial_where varchar(255),
	@generate_temporal_where varchar(255),
	@collection_table varchar(255),
	@visible_rep int,
	@enable_visibility int,
	@lower_x_te_them float,
	@lower_y_te_them float,
	@upper_x_te_them float,
	@upper_y_te_them float,
	@creation_time datetime,
	@theme_id_doTemaRecemInserido int,

	@legend_id int,
	@legendId_theme_id int,
	@legendId_group_id int,
	@legendId_num_objs int,
	@legend_id_daLegendaRecemInserida int,

	@geom_type int,
    @symb_id int,
    @red int,
    @green int,
    @blue int,
    @transparency int,
    @width int,
    @contour_symb_id int,
    @contour_red int,
    @contour_green int,
    @contour_blue int,
    @contour_transp int,
    @contour_width int,
    @size_value int,
    @pt_angle int,
    @family varchar(255),
    @bold int,
    @italic int,
    @alignment_vert float,
    @alignment_horiz float,
    @tab_size int,
    @line_space int,
    @fixed_size int,
	@legend_idDaTeVisual int

	BEGIN
		INSERT INTO dbo.SystemUsers (USRLog, USRNom, USRFun, USRDep, USRPwd, USRExp, USRBrk, USRDATA)
			SELECT @nomeNovoUsuarioUSRNom, @nomeNovoUsuarioUSRNom, @permissaoDoUsuario, 1, @nomeNovoUsuarioUSRNom, 0, 0, COALESCE(Right('0'+Convert(varchar(10), day(GETDATE())),2),'') + COALESCE(Right('0'+Convert(varchar(10), month(GETDATE())),2),'') + COALESCE(Right('0'+Convert(varchar(10), year(GETDATE())),4),'')
	END

	DECLARE cursorVistasDoUsuario CURSOR FOR
		SELECT view_id, projection_id, name, visibility, lower_x, lower_y, upper_x, upper_y FROM dbo.te_view where user_name = @nomeUsuarioAtualUSRNom
	OPEN cursorVistasDoUsuario
	FETCH NEXT FROM cursorVistasDoUsuario INTO @view_id, @projection_id, @name, @visibility, @lower_x, @lower_y, @upper_x, @upper_y 
	WHILE @@FETCH_STATUS = 0
	BEGIN
		BEGIN
			INSERT INTO dbo.te_view (projection_id, name, user_name, visibility, lower_x, lower_y, upper_x, upper_y) 
				VALUES (@projection_id, @name, @nomeNovoUsuarioUSRNom, @visibility, @lower_x, @lower_y, @upper_x, @upper_y)
			SET @view_id_daVistaRecemInserida = SCOPE_IDENTITY()
			PRINT (STR(@view_id_daVistaRecemInserida) + ' ' + STR(@view_id) + ' ' + STR(@projection_id) + ' ' + @name + ' ' + @nomeUsuarioAtualUSRNom + ' ' + STR(@visibility))
			DECLARE cursorTemaDaVista CURSOR FOR
				SELECT theme_id ,layer_id ,view_id ,name ,parent_id ,priority ,node_type ,min_scale ,max_scale ,generate_attribute_where ,generate_spatial_where ,generate_temporal_where ,collection_table ,visible_rep ,enable_visibility ,lower_x ,lower_y ,upper_x ,upper_y ,creation_time
  					FROM dbo.te_theme
  					WHERE view_id = @view_id
			OPEN cursorTemaDaVista
			FETCH NEXT FROM cursorTemaDaVista INTO @theme_id ,@layer_id ,@view_id_te_theme ,@name_te_theme ,@parent_id ,@priority ,@node_type ,@min_scale ,@max_scale ,@generate_attribute_where ,@generate_spatial_where ,@generate_temporal_where ,@collection_table ,@visible_rep ,@enable_visibility ,@lower_x ,@lower_y ,@upper_x ,@upper_y ,@creation_time
			WHILE @@FETCH_STATUS = 0
			BEGIN
				BEGIN
					INSERT INTO dbo.te_theme (layer_id ,view_id ,name ,parent_id ,priority ,node_type ,min_scale ,max_scale ,generate_attribute_where ,generate_spatial_where ,generate_temporal_where ,collection_table ,visible_rep ,enable_visibility ,lower_x ,lower_y ,upper_x ,upper_y ,creation_time) 
						VALUES (@layer_id ,@view_id_daVistaRecemInserida ,@name_te_theme ,@parent_id ,@priority ,@node_type ,@min_scale ,@max_scale ,@generate_attribute_where ,@generate_spatial_where ,@generate_temporal_where ,@collection_table ,@visible_rep ,@enable_visibility ,@lower_x ,@lower_y ,@upper_x ,@upper_y ,@creation_time)
					SET @theme_id_doTemaRecemInserido = SCOPE_IDENTITY()
					PRINT (STR(@theme_id_doTemaRecemInserido) + ' ' + STR(@theme_id) + ' ' + STR(@view_id_daVistaRecemInserida) + ' ' + @name_te_theme)
					DECLARE cursorLegenda CURSOR FOR
						SELECT legend_id, theme_id ,group_id ,num_objs
							FROM dbo.te_legend
							WHERE theme_id = @theme_id
					OPEN cursorLegenda
					FETCH NEXT FROM cursorLegenda INTO @legend_id, @legendId_theme_id, @legendId_group_id, @legendId_num_objs
					WHILE @@FETCH_STATUS = 0
					BEGIN
						BEGIN
							INSERT INTO dbo.te_legend (theme_id, group_id, num_objs)
								VALUES (@theme_id_doTemaRecemInserido, @legendId_group_id, @legendId_num_objs)
							SET @legend_id_daLegendaRecemInserida = SCOPE_IDENTITY()
							PRINT (STR(@legend_id_daLegendaRecemInserida) + ' ' + STR(@legend_id) + ' ' + STR(@legendId_theme_id) + ' ' + STR(@legendId_group_id) + ' ' + STR(@legendId_num_objs))
							DECLARE cursorTeVisual CURSOR FOR
								SELECT legend_id ,geom_type ,symb_id ,red ,green ,blue ,transparency ,width ,contour_symb_id ,contour_red ,contour_green ,contour_blue ,contour_transp ,contour_width ,size_value ,pt_angle ,family ,bold ,italic ,alignment_vert ,alignment_horiz ,tab_size ,line_space ,fixed_size
  									FROM dbo.te_visual
									WHERE legend_id = @legend_id
							OPEN cursorTeVisual
							FETCH NEXT FROM cursorTeVisual INTO @legend_idDaTeVisual, @geom_type ,@symb_id ,@red ,@green ,@blue ,@transparency ,@width ,@contour_symb_id ,@contour_red ,@contour_green ,@contour_blue ,@contour_transp ,@contour_width ,@size_value ,@pt_angle ,@family ,@bold ,@italic ,@alignment_vert ,@alignment_horiz ,@tab_size ,@line_space ,@fixed_size
							WHILE @@FETCH_STATUS = 0
							BEGIN
								BEGIN
									INSERT INTO dbo.te_visual (legend_id ,geom_type ,symb_id ,red ,green ,blue ,transparency ,width ,contour_symb_id ,contour_red ,contour_green ,contour_blue ,contour_transp ,contour_width ,size_value ,pt_angle ,family ,bold ,italic ,alignment_vert ,alignment_horiz ,tab_size ,line_space ,fixed_size)
										VALUES (@legend_id_daLegendaRecemInserida, @geom_type ,@symb_id ,@red ,@green ,@blue ,@transparency ,@width ,@contour_symb_id ,@contour_red ,@contour_green ,@contour_blue ,@contour_transp ,@contour_width ,@size_value ,@pt_angle ,@family ,@bold ,@italic ,@alignment_vert ,@alignment_horiz ,@tab_size ,@line_space ,@fixed_size)
									PRINT (STR(@legend_id_daLegendaRecemInserida) + ' ' + STR(@geom_type))
								END
								FETCH NEXT FROM cursorTeVisual INTO @legend_idDaTeVisual, @geom_type ,@symb_id ,@red ,@green ,@blue ,@transparency ,@width ,@contour_symb_id ,@contour_red ,@contour_green ,@contour_blue ,@contour_transp ,@contour_width ,@size_value ,@pt_angle ,@family ,@bold ,@italic ,@alignment_vert ,@alignment_horiz ,@tab_size ,@line_space ,@fixed_size
							END
							CLOSE cursorTeVisual
							DEALLOCATE cursorTeVisual
						END
						FETCH NEXT FROM cursorLegenda INTO @legend_id, @legendId_theme_id, @legendId_group_id, @legendId_num_objs
					END
					CLOSE cursorLegenda
					DEALLOCATE cursorLegenda
				END
				FETCH NEXT FROM cursorTemaDaVista INTO @theme_id ,@layer_id ,@view_id_te_theme ,@name_te_theme ,@parent_id ,@priority ,@node_type ,@min_scale ,@max_scale ,@generate_attribute_where ,@generate_spatial_where ,@generate_temporal_where ,@collection_table ,@visible_rep ,@enable_visibility ,@lower_x ,@lower_y ,@upper_x ,@upper_y ,@creation_time
			END
			CLOSE cursorTemaDaVista
			DEALLOCATE cursorTemaDaVista
		END
		FETCH NEXT FROM cursorVistasDoUsuario INTO @view_id, @projection_id, @name, @visibility, @lower_x, @lower_y, @upper_x, @upper_y 
	END
	CLOSE cursorVistasDoUsuario
	DEALLOCATE cursorVistasDoUsuario
END