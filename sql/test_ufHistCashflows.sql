-- Test Script for UserForm: ufHistCashflows
-- Target SP: edw.pr_usp_investor_transactions2

DECLARE @SourceTableVolVal    varchar(100) = '12345';      -- REPLACE with valid Investor ID
DECLARE @MetricName           varchar(100) = 'Pro-Rata Distribution';
DECLARE @StartDate            date         = '2010-01-01';
DECLARE @EndDate              date         = '2025-12-31';
DECLARE @ViewCurrencyID       int          = 1;            -- 1 = USD (Check exchange_rates table)
DECLARE @FilterSourceCurrency varchar(20)  = NULL;         -- Optional: 'EUR', 'GBP', etc.
DECLARE @AIVFundGroupID       int          = NULL;         -- Optional: Filter by Fund Group

-- Optional Parameters (not always populated by form)
DECLARE @FundTypesExclude     varchar(400) = NULL;
DECLARE @InvestorGroupID      int          = NULL;
DECLARE @MaxRows              int          = 1000;

PRINT 'Executing edw.pr_usp_investor_transactions2 with dummy parameters...';

EXEC edw.pr_usp_investor_transactions2
    @MetricName           = @MetricName,
    @SourceTableVolVal    = @SourceTableVolVal,
    @StartDate            = @StartDate,
    @EndDate              = @EndDate,
    @ViewCurrencyID       = @ViewCurrencyID,
    @FilterSourceCurrency = @FilterSourceCurrency,
    @FundTypesExclude     = @FundTypesExclude,
    @InvestorGroupID      = @InvestorGroupID,
    @AIVFundGroupID       = @AIVFundGroupID,
    @MaxRows              = @MaxRows,
    @OrderBy              = 'date';
