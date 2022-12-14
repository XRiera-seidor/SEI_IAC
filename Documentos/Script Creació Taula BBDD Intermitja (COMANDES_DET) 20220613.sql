USE [SAP_GESTION_TEST]
GO

/****** Object:  Table [dbo].[COMANDES_DET]    Script Date: 13/06/2022 12:32:29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COMANDES_DET](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ID_CAP] [int] NOT NULL,
	[Linia] [int] NOT NULL,
	[Article] [nvarchar](50) NULL,
	[Quantitat] [float] NULL,
	[Magatzem] [nchar](10) NULL,
	[Preu] [float] NULL,
	[Ordre] [int] NULL,
	[Estat] [nchar](20) NULL,
	[Eliminar] [nchar](1) NULL CONSTRAINT [DF_COMANDES_DET_Eliminar]  DEFAULT (N'N'),
	[DataEntrega] [date] NULL,
	[Cost] [float] NULL
) ON [PRIMARY]

GO


