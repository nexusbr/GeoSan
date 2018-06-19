USE [ArturNogueira-B-1]
GO
/****** Object:  Table [dbo].[WaterComponentsTypes]    Script Date: 02/10/2016 08:45:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[WaterComponentsTypes](
	[id_Type] [int] IDENTITY(1,1) NOT NULL,
	[Description_] [varchar](25) NULL,
	[Specification_] [varchar](100) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
SET IDENTITY_INSERT [dbo].[WaterComponentsTypes] ON														/* rodar esta linha antes de realizar as inserções */
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (43, N'VÁLVULA', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (44, N'BOMBA POT', N'PUMP')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (45, N'BOMBA CARGA', N'PUMP')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (46, N'BOMBA CURVA', N'PUMP')
SET IDENTITY_INSERT [dbo].[WaterComponentsTypes] OFF

INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 1, N'ALT INICIAL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'NINICIAL')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 2, N'ALT MIN', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'NMINIMO')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 1, N'POTÊNCIA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'POWER')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 2, N'CARGA (ALT)', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'CARGA')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 3, N'VAZÃO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'VAZAO')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 4, N'RENDIMENTO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'RENDIM')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 5, N'CURVA DA BOMBA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'CURBOMBA')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 6, N'CURVA DE REND', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'CURREND')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 3, N'ALT MAX', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'NMAXIMO')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 4, N'DIÂMETRO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'DIAMETER')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 5, N'CURVA DE VOLUME', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'VOLCURVE')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (32, 1, N'PARAM CONTROLE', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PARCONT')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (36, 1, N'NÍVEL ÁGUA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'HEAD')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (36, 2, N'PADRÃO DE NÍVEL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PATTERN')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (32, 2, N'NOME', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (36, 3, N'NOME', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 7, N'DESCRIÇÃO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 6, N'DESCRIÇÃO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (43, 1, N'DIAMETRO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'DIAMETER')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (43, 2, N'TIPO', 1, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'TYPE')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (43, 3, N'PARAM CONTR', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'SETTING')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (43, 4, N'COEF PERDA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'MINORLOSS')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (44, 1, N'POTÊNCIA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'POWER')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (44, 2, N'REG VELOC', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'SPEED')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (44, 3, N'PADRÃO TEMPORAL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PATTERN')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (44, 4, N'DESCRIÇÃO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 1, N'ALTURA (CARGA)', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'CARGA')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 2, N'VAZÃO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'VAZAO')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 3, N'REG VELOC', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'SPEED')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 4, N'PADRÃO TEMPORAL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PATTERN')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 5, N'DESCRIÇÃO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (46, 1, N'CURVA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'CURBOMBA')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (46, 2, N'REG VELOC', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'SPEED')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (46, 3, N'PADRÃO TEMPORAL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PATTERN')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (46, 4, N'DESC', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')

/* Selecione no SQLServer Management Studio: Tools - Options - Designer e então retire a seleção de Prevent saving changes that require the table re-creation */
ALTER TABLE [dbo].[WaterComponentsSelections] ALTER COLUMN [Value_] [nvarchar](50) 

INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'PRV-RED PRESSÃO', N'1', N'Válvula redutora de pressão')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'PSV-SUST PRESSÃO', N'2', N'Válvula Sustent. de Pressão')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'PBV-PERDA CARGA', N'3', N'Válvula de perda de carga fixa')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'FCV-REG VAZÃO', N'4', N'Válvula reguladora de vazão')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'TCV-CONT PERDA CARGA', N'5', N'Válvula contr de perda carga')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'GPV-GENÉRICA', N'6', N'Válvula genérica')
