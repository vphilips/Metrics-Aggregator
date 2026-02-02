-- Test Script for UserForm: ufCompanyDiversification
-- Target SP: edw.pr_usp_investor_transactions_cd

-- Basic Params
DECLARE @SourceTableVolVal    varchar(100) = '12345';      -- REPLACE with valid Investor ID
DECLARE @MetricName           varchar(100) = 'Net Asset Value';
DECLARE @Date                 date         = '2025-12-31';
DECLARE @ViewCurrencyCode     varchar(20)  = 'USD';
DECLARE @FilterSourceCurrency varchar(20)  = NULL;
DECLARE @AIVFundGroupID       int          = NULL;
DECLARE @FundTypesExclude     varchar(400) = NULL;
DECLARE @InvestorGroupID      int          = NULL;
DECLARE @MaxRows              int          = 1000;

-- Company Filters (Set to NULL if not testing specific filter)
DECLARE @CompExpGeoBroad      varchar(100) = NULL;
DECLARE @CompExpGeoCountry    varchar(100) = NULL;
DECLARE @CompExpIndId         varchar(100) = NULL;
DECLARE @CompExpIndBroad      varchar(100) = NULL;
DECLARE @CompExpIndCategory   varchar(100) = NULL;
DECLARE @CompExpInvType       varchar(100) = NULL;
DECLARE @CompExpInvYear       int          = NULL;
DECLARE @CompExpIsPublic      bit          = NULL;
DECLARE @CompExpStage         varchar(100) = NULL;
DECLARE @CompExpStageBroad    varchar(100) = NULL;


PRINT 'Executing edw.pr_usp_investor_transactions_cd with dummy parameters...';

EXEC edw.pr_usp_investor_transactions_cd
    @MetricName           = @MetricName,
    @SourceTableVolVal    = @SourceTableVolVal,
    @Date                 = @Date,
    @ViewCurrencyCode     = @ViewCurrencyCode,
    @FilterSourceCurrency = @FilterSourceCurrency,
    @FundTypesExclude     = @FundTypesExclude,
    @InvestorGroupID      = @InvestorGroupID,
    @AIVFundGroupID       = @AIVFundGroupID,
    @MaxRows              = @MaxRows,
    @OrderBy              = 'date',
    -- New Company Filters
    @CompExpGeoBroad      = @CompExpGeoBroad,
    @CompExpGeoCountry    = @CompExpGeoCountry,
    @CompExpIndId         = @CompExpIndId,
    @CompExpIndBroad      = @CompExpIndBroad,
    @CompExpIndCategory   = @CompExpIndCategory,
    @CompExpInvType       = @CompExpInvType,
    @CompExpInvYear       = @CompExpInvYear,
    @CompExpIsPublic      = @CompExpIsPublic,
    @CompExpStage         = @CompExpStage,
    @CompExpStageBroad    = @CompExpStageBroad;
