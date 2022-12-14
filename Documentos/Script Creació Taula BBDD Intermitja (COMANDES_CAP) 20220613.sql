USE [SAP_GESTION_TEST]
GO

/****** Object:  Table [dbo].[COMANDES_CAP]    Script Date: 13/06/2022 12:32:24 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COMANDES_CAP](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[NumComanda] [nvarchar](15) NULL,
	[Client] [nvarchar](15) NULL,
	[DataComanda] [date] NULL,
	[DataEntrega] [date] NULL,
	[Comentaris] [nvarchar](254) NULL,
	[Importat] [nchar](1) NULL CONSTRAINT [DF_COMANDES_CAP_Importat]  DEFAULT (N'N'),
	[Processat] [datetime] NULL,
	[Error] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO


