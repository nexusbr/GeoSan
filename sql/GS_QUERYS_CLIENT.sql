USE [geosan]
GO

/****** Object:  Table [dbo].[GS_QUERYS_CLIENT]    Script Date: 06/15/2023 13:47:48 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[GS_QUERYS_CLIENT](
	[CLIENT_ID] [int] NOT NULL,
	[QUERY_ID] [int] NOT NULL,
	[QUERYSTRING] [varchar](1000) NULL,
 CONSTRAINT [PK_GS_QUERYS_CLIENT] PRIMARY KEY CLUSTERED 
(
	[CLIENT_ID] ASC,
	[QUERY_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO