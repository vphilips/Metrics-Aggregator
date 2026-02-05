CREATE OR ALTER PROCEDURE sp_GetPortfolioDiversification
    (
        @SourceTableVolVal varchar(100) = NULL,
        -- Dates Removed
        @ViewCurrencyID int = 1,
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
        @SourceTableVolVal as source_table_vol_val,
        @Metric as metric,
        @FilterSourceCurrency as source_currency_filter,
        @ViewCurrencyID as view_currency
END
GO
