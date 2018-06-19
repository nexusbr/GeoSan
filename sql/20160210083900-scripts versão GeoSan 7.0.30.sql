USE [ArturNogueira-B]
GO
/****** Object:  Table [dbo].[WaterComponentsTypes]    Script Date: 02/10/2016 08:40:47 ******/
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
SET IDENTITY_INSERT [dbo].[WaterComponentsTypes] ON 													
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (0, N'DESC.', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (1, N'ADAPTADOR', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (2, N'ATUADOR', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (3, N'BOMBA', N'PUMP')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (4, N'BOOSTER', N'PUMP')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (5, N'CAP', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (6, N'CAPTAÇÃO AGUA BRUTA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (7, N'CRUZETA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (8, N'CURVA 11-15', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (9, N'CURVA 22-30', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (10, N'CURVA 45', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (11, N'CURVA 90', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (12, N'DESCARGA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (13, N'ELEVATORIA DE AGUA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (14, N'FILTRO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (15, N'HIDRANTE', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (16, N'HIDRÔMETRO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (17, N'JUNÇÃO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (18, N'JUNTA DE ADAPTAÇÃO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (19, N'LUVA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (20, N'MACRO MEDIDOR', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (21, N'MEDIDOR DE NÍVEL', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (22, N'MEDIDOR PRESSÃO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (23, N'MEDIDOR VAZÃO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (24, N'NÃO IDENTIFICADO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (25, N'RNF-Poço Profundo', N'RNF')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (26, N'REDUÇÃO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (27, N'REGISTRO', N'REGISTER')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (28, N'RNV-Reserv Nível Variável', N'RNV')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (29, N'TE', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (30, N'VALVULA RETENÇÃO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (31, N'VENTOSA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (32, N'PRV-Válvula Red Pres', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (35, N'REG. FIXO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (36, N'RNF-Reserv Nível Fixo', N'RNF')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (38, N'PSV-Válvula Sus Pres', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (39, N'PBV-Válvula Perda Carga F', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (40, N'FCV-Válvula Reg Vazão', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (41, N'TCV-Válvula Contr Perda C', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (42, N'GPV-Válvula Genérica', N'REGISTER')
SET IDENTITY_INSERT [dbo].[WaterComponentsTypes] OFF
/****** Object:  Table [dbo].[WaterComponentsSubTypes]    Script Date: 02/10/2016 08:40:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[WaterComponentsSubTypes](
	[id_Type] [int] NOT NULL,
	[id_SubType] [int] NOT NULL,
	[Description_] [varchar](50) NULL,
	[Selection_] [bit] NOT NULL,
	[Max_] [numeric](18, 4) NULL,
	[Min_] [numeric](18, 4) NULL,
	[DefaultValue] [nvarchar](50) NOT NULL,
	[DataType] [int] NULL,
	[EPAREF] [nvarchar](10) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (27, 1, N'SITUAÇÃO', 1, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (27, 2, N'NÚMERO DO REGISTRO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 129, NULL)
/****** Object:  Table [dbo].[WaterComponentsSelections]    Script Date: 02/10/2016 08:40:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[WaterComponentsSelections](
	[id_Type] [int] NOT NULL,
	[id_SubType] [int] NOT NULL,
	[Option_] [varchar](25) NOT NULL,
	[Value_] [tinyint] NOT NULL,
	[Description_] [varchar](30) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (27, 1, N'ABERTO', 1, N'')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (27, 1, N'DESCONHECIDA', 0, N'')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (27, 1, N'FECHADO', 2, N'')
