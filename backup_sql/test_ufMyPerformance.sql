-- Test Script for UserForm: ufMyPerformance
-- Target SP: edw.pr_usp_investor_transactions3

DECLARE @SourceTableVolVal    varchar(100) = '12345';      -- REPLACE with valid Investor ID
DECLARE @MetricName           varchar(100) = 'Pro-Rata Distribution';
DECLARE @Date                 date         = '2025-12-31'; -- As-Of Date
DECLARE @ViewCurrencyID       int          = 1;            -- 1 = USD
DECLARE @FilterSourceCurrency varchar(20)  = NULL;         -- Optional
DECLARE @AIVFundGroupID       int          = NULL;         -- Optional

-- Optional Parameters
DECLARE @InvestorGroupID      int          = NULL;
DECLARE @MaxRows              int          = 1000;

PRINT 'Executing edw.pr_usp_investor_transactions3 with dummy parameters...';

EXEC edw.pr_usp_investor_transactions3
    @MetricName           = @MetricName,
    @SourceTableVolVal    = @SourceTableVolVal,
    @Date                 = @Date,
    @ViewCurrencyID       = @ViewCurrencyID,
    @FilterSourceCurrency = @FilterSourceCurrency,
    @InvestorGroupID      = @InvestorGroupID,
    @AIVFundGroupID       = @AIVFundGroupID,
    @MaxRows              = @MaxRows,
    @OrderBy              = 'date';
