USE [SAP_GESTION_TEST]
GO

/****** Object:  Table [dbo].[MOVIMENTS]    Script Date: 25/01/2022 12:09:32 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MOVIMENTS](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Article] [nvarchar](50) NULL,
	[Quantitat] [float] NULL,
	[TipusMov] [nchar](10) NULL,
	[Importat] [nchar](1) NULL,
	[Magatzem] [nchar](10) NULL,
	[Processat] [datetime] NULL,
	[Error] [nvarchar](max) NULL,
	[Observacions] [nvarchar](254) NULL,
	[Data] [date] NULL,
 CONSTRAINT [PK_MOVIMENTS] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO


