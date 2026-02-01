CREATE OR ALTER PROCEDURE sp_GetPortfolioDiversification
    (
        @Account varchar(100) = NULL,
        @Client varchar(100) = NULL,
        @Investor varchar(100) = NULL,
        -- Dates Removed
        @ViewCurrencyCode varchar(20) = 'USD',
        @FilterSourceCurrency varchar(20) = NULL,
        @Metric varchar(100) = NULL,
        @Date date = NULL,
        @AIVFundGroupID int = NULL,
        @EntryFund varchar(100) = NULL,
        @Manager varchar(100) = NULL,
        @Portfolio varchar(100) = NULL,
        @PortfolioCloseYear int = NULL,
        @PortfolioCommitmentYear int = NULL,
        @PortfolioGeography varchar(100) = NULL,
        @PortfolioGeographyBroad varchar(100) = NULL,
        @PortfolioGeographyL3 varchar(100) = NULL,
        @PortfolioGeographyL5 varchar(100) = NULL,
        @PortfolioIndustry varchar(100) = NULL,
        @PortfolioIndustryL1 varchar(100) = NULL,
        @PortfolioStage varchar(100) = NULL,
        @PortfolioStageBroad varchar(100) = NULL,
        @PortfolioStatus varchar(100) = NULL,
        @PortfolioTypeBroadID varchar(100) = NULL,
        @PortfolioVintageYear int = NULL
    )
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        'PortfolioDiversification' as report_type,
        @Account as account,
        @Client as client,
        @Metric as metric,
        @FilterSourceCurrency as source_currency_filter,
        @ViewCurrencyCode as view_currency
END
GO
