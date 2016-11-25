USE [BPSS_MASTER]
GO

/****** Object:  Table [dbo].[T_REF_BFC_RE_HEADERS]    Script Date: 21/10/2016 11:41:15 ******/
DROP TABLE [dbo].[T_REF_BFC_RE_HEADERS]
GO

/****** Object:  Table [dbo].[T_REF_BFC_RE_HEADERS]    Script Date: 21/10/2016 11:41:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[T_REF_BFC_RE_HEADERS](
	[INDICE_COL] [int] IDENTITY(1,1) NOT NULL,
	[RE_INVESTMENT] [varchar](80) NOT NULL,
	[RE_OWN_USE] [varchar](80) NOT NULL,
	[ALIAS_COL] [varchar](80) NULL,
	[CADIS_SYSTEM_INSERTED] [datetime] NULL DEFAULT (getdate()),
	[CADIS_SYSTEM_UPDATED] [datetime] NULL DEFAULT (getdate()),
	[CADIS_SYSTEM_CHANGEDBY] [nvarchar](50) NULL DEFAULT ('UNKNOWN'),
	[CADIS_SYSTEM_PRIORITY] [int] NULL DEFAULT ((1)),
PRIMARY KEY CLUSTERED 
(
	[INDICE_COL] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

USE [BPSS_MASTER]
GO

INSERT INTO [dbo].[T_REF_BFC_RE_HEADERS]
           ([RE_INVESTMENT]
           ,[RE_OWN_USE]
		   ,[ALIAS_COL])
     VALUES ('Entity','Entity',''),
			('Building','Building',''),
			('Amount to split by building','Amount to split by building',''),
			('13111 Real Estate, Land - Cost','12051 Land (own use) - Cost',''),
			('13121 Real Estate, Building - Cost','12011 Building (own use) - Cost',''),
			('13131 Real Estate, Fittings & Fixtures - Cost','12021 Buildings fixtures & fittings (own use) - Cost',''),
			('13112 Real Estate, Land - Amortisation','12052 Land (own use) - Amortisation',''),
			('13122 Real Estate, Building - Amortisation','12012 Building (own use) - Amortisation',''),
			('13132 Real Estate, Fittings & Fixtures - Amortisation','12022 Buildings fixtures & fittings (own use) - Amort',''),
			('13115 Real Estate, Land - Impairment','12055 Land (own use) - Impairment',''),
			('13125 Real Estate, Building - Impairment','12015 Building (own use) - Impairment',''),
			('13135 Real Estate, Fittings & Fixtures - Impairment','12025 Buildings fixtures & fittings (own use) - Imprmt',''),
			('13141 Real Estate, Construction in progress - Cost','12031 Fixed asset in Progress (own use) - Cost',''),
			('Real estate net booking value','Tangible assets - net booking value',''),
			('Date of valuation','Date of valuation',''),
			('Real estate - external valuation','Real estate - external valuation',''),
			('Real estate - WIP since last valuation','Real estate - WIP since last valuation',''),
			('Real estate - Fair value','Real estate - Fair value',''),
			('Real estate - URGL','Real estate - URGL',''),
			('21321 Other financial debt - Real Estate','21321 Other financial debt - Real Estate',''),
			('21331 Accrued interest on other finan debt - Real Estate','21331 Accrued interest on other finan debt - Real Estate',''),
			('21341 Issuance premium and fees -  Real estate debt','21341 Issuance premium and fees -  Real estate debt',''),
			('Total Real Estate Financing','Total Real Estate Financing',''),
			('21320 Other financial debt','21320 Other financial debt',''),
			('21330 Accrued interest on other financial debt','21330 Accrued interest on other financial debt',''),
			('21340 Issuance premium and fees - Other debt','21340 Issuance premium and fees - Other debt',''),
			('Total Other Financial Debt','Total Other Financial Debt',''),
			('Amount Pledged','Amount Pledged',''),
			('Closing Period','Closing Period','')
GO