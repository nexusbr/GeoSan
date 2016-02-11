/* Para criar as tabelas */

USE [ArturNogueira-B]

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE XXXX
GO
SET ANSI_PADDING OFF
GO


/* insere os tipos de componentes os quais um n� pode vir a ser */

/****** Object:  Table [dbo].[WaterComponentsTypes]    Script Date: 02/09/2016 18:07:15 ******/

CREATE TABLE [dbo].[WaterComponentsTypes](
	[id_Type] [int] IDENTITY(1,1) NOT NULL,
	[Description_] [varchar](25) NULL,
	[Specification_] [varchar](100) NULL
) ON [PRIMARY]

SET IDENTITY_INSERT [dbo].[WaterComponentsTypes] ON
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (0, N'DESC.', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (1, N'ADAPTADOR', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (2, N'ATUADOR', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (3, N'BOMBA', N'PUMP')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (4, N'BOOSTER', N'PUMP')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (5, N'CAP', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (6, N'CAPTA��O AGUA BRUTA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (7, N'CRUZETA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (8, N'CURVA 11-15', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (9, N'CURVA 22-30', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (10, N'CURVA 45', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (11, N'CURVA 90', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (12, N'DESCARGA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (13, N'ELEVATORIA DE AGUA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (14, N'FILTRO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (15, N'HIDRANTE', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (16, N'HIDR�METRO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (17, N'JUN��O', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (18, N'JUNTA DE ADAPTA��O', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (19, N'LUVA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (20, N'MACRO MEDIDOR', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (21, N'MEDIDOR DE N�VEL', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (22, N'MEDIDOR PRESS�O', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (23, N'MEDIDOR VAZ�O', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (24, N'N�O IDENTIFICADO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (25, N'RNF-Po�o Profundo', N'RNF')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (26, N'REDU��O', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (27, N'REGISTRO', N'REGISTER')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (28, N'RNV-Reserv N�vel Vari�vel', N'RNV')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (29, N'TE', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (30, N'VALVULA RETEN��O', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (31, N'VENTOSA', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (32, N'PRV-V�lvula Red Pres', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (35, N'REG. FIXO', N'JUNCTION')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (36, N'RNF-Reserv N�vel Fixo', N'RNF')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (38, N'PSV-V�lvula Sus Pres', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (39, N'PBV-V�lvula Perda Carga F', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (40, N'FCV-V�lvula Reg Vaz�o', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (41, N'TCV-V�lvula Contr Perda C', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (42, N'GPV-V�lvula Gen�rica', N'REGISTER')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (43, N'V�LVULA', N'VALVE')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (44, N'BOMBA POT', N'PUMP')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (45, N'BOMBA CARGA', N'PUMP')
INSERT [dbo].[WaterComponentsTypes] ([id_Type], [Description_], [Specification_]) VALUES (46, N'BOMBA CURVA', N'PUMP')
SET IDENTITY_INSERT [dbo].[WaterComponentsTypes] OFF


/* Apaga todos os metadados dos n�s */

delete from WaterComponentsTypes where id_Type = 0 
delete from WaterComponentsTypes where id_Type = 1 
delete from WaterComponentsTypes where id_Type = 2
delete from WaterComponentsTypes where id_Type = 3 
delete from WaterComponentsTypes where id_Type = 4 
delete from WaterComponentsTypes where id_Type = 5 
delete from WaterComponentsTypes where id_Type = 6 
delete from WaterComponentsTypes where id_Type = 7 
delete from WaterComponentsTypes where id_Type = 8 
delete from WaterComponentsTypes where id_Type = 9
delete from WaterComponentsTypes where id_Type = 10
delete from WaterComponentsTypes where id_Type = 11
delete from WaterComponentsTypes where id_Type = 12
delete from WaterComponentsTypes where id_Type = 13
delete from WaterComponentsTypes where id_Type = 14
delete from WaterComponentsTypes where id_Type = 15
delete from WaterComponentsTypes where id_Type = 16
delete from WaterComponentsTypes where id_Type = 17
delete from WaterComponentsTypes where id_Type = 17
delete from WaterComponentsTypes where id_Type = 19
delete from WaterComponentsTypes where id_Type = 20
delete from WaterComponentsTypes where id_Type = 21
delete from WaterComponentsTypes where id_Type = 22
delete from WaterComponentsTypes where id_Type = 23
delete from WaterComponentsTypes where id_Type = 24
delete from WaterComponentsTypes where id_Type = 25
delete from WaterComponentsTypes where id_Type = 26
delete from WaterComponentsTypes where id_Type = 27
delete from WaterComponentsTypes where id_Type = 28
delete from WaterComponentsTypes where id_Type = 29
delete from WaterComponentsTypes where id_Type = 30
delete from WaterComponentsTypes where id_Type = 31
delete from WaterComponentsTypes where id_Type = 32
delete from WaterComponentsTypes where id_Type = 33
delete from WaterComponentsTypes where id_Type = 34
delete from WaterComponentsTypes where id_Type = 35
delete from WaterComponentsTypes where id_Type = 36
delete from WaterComponentsTypes where id_Type = 37
delete from WaterComponentsTypes where id_Type = 38
delete from WaterComponentsTypes where id_Type = 39
delete from WaterComponentsTypes where id_Type = 40
delete from WaterComponentsTypes where id_Type = 41
delete from WaterComponentsTypes where id_Type = 42
delete from WaterComponentsTypes where id_Type = 43
delete from WaterComponentsTypes where id_Type = 44
delete from WaterComponentsTypes where id_Type = 45
delete from WaterComponentsTypes where id_Type = 46

/* insere par�metros de todos subtipos */

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

INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (27, 1, N'SITUA��O', 1, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (27, 2, N'N�MERO DO REGISTRO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, NULL)
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 1, N'ALT INICIAL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'NINICIAL')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 2, N'ALT MIN', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'NMINIMO')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 1, N'POT�NCIA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'POWER')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 2, N'CARGA (ALT)', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'CARGA')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 3, N'VAZ�O', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'VAZAO')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 4, N'RENDIMENTO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'RENDIM')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 5, N'CURVA DA BOMBA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'CURBOMBA')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 6, N'CURVA DE REND', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'CURREND')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 3, N'ALT MAX', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'NMAXIMO')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 4, N'DI�METRO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'DIAMETER')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 5, N'CURVA DE VOLUME', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'VOLCURVE')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (32, 1, N'PARAM CONTROLE', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PARCONT')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (36, 1, N'N�VEL �GUA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'HEAD')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (36, 2, N'PADR�O DE N�VEL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PATTERN')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (32, 2, N'NOME', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (36, 3, N'NOME', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (3, 7, N'DESCRI��O', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (28, 6, N'DESCRI��O', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (43, 1, N'DIAMETRO', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'DIAMETER')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (43, 2, N'TIPO', 1, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'TYPE')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (43, 3, N'PARAM CONTR', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'SETTING')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (43, 4, N'COEF PERDA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'MINORLOSS')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (44, 1, N'POT�NCIA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'POWER')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (44, 2, N'REG VELOC', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'SPEED')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (44, 3, N'PADR�O TEMPORAL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PATTERN')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (44, 4, N'DESCRI��O', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 1, N'ALTURA (CARGA)', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'CARGA')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 2, N'VAZ�O', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'VAZAO')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 3, N'REG VELOC', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'SPEED')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 4, N'PADR�O TEMPORAL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PATTERN')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (45, 5, N'DESCRI��O', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (46, 1, N'CURVA', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'CURBOMBA')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (46, 2, N'REG VELOC', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 5, N'SPEED')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (46, 3, N'PADR�O TEMPORAL', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 2, N'PATTERN')
INSERT [dbo].[WaterComponentsSubTypes] ([id_Type], [id_SubType], [Description_], [Selection_], [Max_], [Min_], [DefaultValue], [DataType], [EPAREF]) VALUES (46, 4, N'DESC', 0, CAST(0.0000 AS Numeric(18, 4)), CAST(0.0000 AS Numeric(18, 4)), N'0', 202, N'DESC')

/* Modificado o formato do dado da coluna Value_ de WaterComponentsSelection */

CREATE TABLE [dbo].[WaterComponentsSelections](
	[id_Type] [int] NOT NULL,
	[id_SubType] [int] NOT NULL,
	[Option_] [varchar](25) NOT NULL,
	[Value_] [nvarchar](50) NOT NULL,
	[Description_] [varchar](30) NULL
) ON [PRIMARY]

GO

INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (27, 1, N'ABERTO', N'1', N'')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (27, 1, N'DESCONHECIDA', N'0', N'')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (27, 1, N'FECHADO', N'2', N'')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'PRV-RED PRESS�O', N'1', N'V�lvula redutora de press�o')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'PSV-SUST PRESS�O', N'2', N'V�lvula Sustent. de Press�o')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'PBV-PERDA CARGA', N'3', N'V�lvula de perda de carga fixa')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'FCV-REG VAZ�O', N'4', N'V�lvula reguladora de vaz�o')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'TCV-CONT PERDA CARGA', N'5', N'V�lvula contr de perda carga')
INSERT [dbo].[WaterComponentsSelections] ([id_Type], [id_SubType], [Option_], [Value_], [Description_]) VALUES (43, 2, N'GPV-GEN�RICA', N'6', N'V�lvula gen�rica')

