USE [SAP_GESTION_TEST]
GO

/****** Object:  Table [dbo].[MOVIMENTS]    Script Date: 13/06/2022 12:32:33 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MOVIMENTS](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Article] [nvarchar](50) NULL,
	[Quantitat] [float] NULL,
	[Magatzem] [nchar](10) NULL,
	[TipusMov] [nchar](10) NULL,
	[Data] [date] NULL,
	[Observacions] [nvarchar](254) NULL,
	[Importat] [nchar](1) NULL CONSTRAINT [DF_MOVIMENTS_Importat]  DEFAULT (N'N'),
	[Processat] [datetime] NULL,
	[Error] [nvarchar](max) NULL,
	[IdLinComanda] [int] NULL,
	[Lot] [nvarchar](40) NULL,
	[Facturable] [nchar](1) NULL CONSTRAINT [DF_MOVIMENTS_Facturable]  DEFAULT (N'N'),
	[Trams] [nvarchar](100) NULL,
 CONSTRAINT [PK_MOVIMENTS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO


