USE [SAP_GESTION_TEST]
GO

/****** Object:  Table [dbo].[ARTICLES]    Script Date: 13/06/2022 12:32:16 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ARTICLES](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Codi] [nvarchar](50) NULL,
	[Descripcio] [nvarchar](100) NULL,
	[Familia] [nchar](3) NULL,
	[Desc_Impressio] [nvarchar](250) NULL,
	[Importat] [nchar](1) NULL CONSTRAINT [DF_ARTICLES_Importat]  DEFAULT (N'N'),
	[Processat] [datetime] NULL,
	[Error] [nvarchar](max) NULL,
	[Inactiu] [nchar](1) NULL,
	[GestionaLots] [nchar](1) NULL CONSTRAINT [DF_ARTICLES_GestionaLots]  DEFAULT (N'S'),
 CONSTRAINT [PK_ARTICLES] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO


