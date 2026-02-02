-- Test Script for UserForm: ufPortfolioDiversification
-- Target SP: sp_GetPortfolioDiversification

-- Basic Params
DECLARE @SourceTableVolVal       varchar(100) = '12345';      -- REPLACE with valid Investor ID
DECLARE @Metric                  varchar(100) = 'Net Asset Value';
DECLARE @ViewCurrencyCode        varchar(20)  = 'USD';
DECLARE @FilterSourceCurrency    varchar(20)  = NULL;
DECLARE @Date                    date         = '2025-12-31';
DECLARE @AIVFundGroupID          int          = NULL;

-- Portfolio Filters (Set to NULL if not testing specific filter)
DECLARE @EntryFund               varchar(100) = NULL;
DECLARE @Manager                 varchar(100) = NULL;
DECLARE @Portfolio               varchar(100) = NULL;
DECLARE @PortfolioCloseYear      int          = NULL;
DECLARE @PortfolioCommitmentYear int          = NULL;
DECLARE @PortfolioGeography      varchar(100) = NULL;
DECLARE @PortfolioGeographyBroad varchar(100) = NULL;
DECLARE @PortfolioGeographyL3    varchar(100) = NULL;
DECLARE @PortfolioGeographyL5    varchar(100) = NULL;
DECLARE @PortfolioIndustry       varchar(100) = NULL;
DECLARE @PortfolioIndustryL1     varchar(100) = NULL;
DECLARE @PortfolioStage          varchar(100) = NULL;
DECLARE @PortfolioStageBroad     varchar(100) = NULL;
DECLARE @PortfolioStatus         varchar(100) = NULL;
DECLARE @PortfolioTypeBroadID    varchar(100) = NULL;
DECLARE @PortfolioVintageYear    int          = NULL;

PRINT 'Executing sp_GetPortfolioDiversification with dummy parameters...';

EXEC sp_GetPortfolioDiversification
    @SourceTableVolVal       = @SourceTableVolVal,
    @ViewCurrencyCode        = @ViewCurrencyCode,
    @FilterSourceCurrency    = @FilterSourceCurrency,
    @Metric                  = @Metric,
    @Date                    = @Date,
    @AIVFundGroupID          = @AIVFundGroupID,
    @EntryFund               = @EntryFund,
    @Manager                 = @Manager,
    @Portfolio               = @Portfolio,
    @PortfolioCloseYear      = @PortfolioCloseYear,
    @PortfolioCommitmentYear = @PortfolioCommitmentYear,
    @PortfolioGeography      = @PortfolioGeography,
    @PortfolioGeographyBroad = @PortfolioGeographyBroad,
    @PortfolioGeographyL3    = @PortfolioGeographyL3,
    @PortfolioGeographyL5    = @PortfolioGeographyL5,
    @PortfolioIndustry       = @PortfolioIndustry,
    @PortfolioIndustryL1     = @PortfolioIndustryL1,
    @PortfolioStage          = @PortfolioStage,
    @PortfolioStageBroad     = @PortfolioStageBroad,
    @PortfolioStatus         = @PortfolioStatus,
    @PortfolioTypeBroadID    = @PortfolioTypeBroadID,
    @PortfolioVintageYear    = @PortfolioVintageYear;
