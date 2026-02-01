CREATE OR ALTER PROCEDURE edw.pr_usp_investor_transactions_cd
    (
        @MetricName         varchar(100) = 'Pro-Rata Distribution',
        @SourceTableVolVal  varchar(100),
        @Date               date,                     -- As-Of Date
        @ViewCurrencyCode   varchar(20)  = 'USD',     
        @FilterSourceCurrency varchar(20)= NULL,      -- Filter by specific source currency
        @FundTypesExclude   varchar(400) = NULL,      -- comma-separated list
        @InvestorGroupID    int          = NULL,
        @AIVFundGroupID     int          = NULL,
        @MaxRows            int          = 0,
        @OrderBy            varchar(20)  = 'date',
        
        -- New Company Filters
        @CompExpGeoBroad    varchar(100) = NULL,
        @CompExpGeoCountry  varchar(100) = NULL,
        @CompExpIndId       varchar(100) = NULL,
        @CompExpIndBroad    varchar(100) = NULL,
        @CompExpIndCategory varchar(100) = NULL,
        @CompExpInvType     varchar(100) = NULL,
        @CompExpInvYear     int          = NULL,
        @CompExpIsPublic    bit          = NULL,
        @CompExpStage       varchar(100) = NULL,
        @CompExpStageBroad  varchar(100) = NULL
    )
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @StartKey int = 19000101;
    DECLARE @EndKey   int = CONVERT(int, CONVERT(char(8), @Date, 112));

    DECLARE @EffectiveMaxRows int = CASE WHEN @MaxRows IS NULL OR @MaxRows <= 0 THEN 2147483647 ELSE @MaxRows END;

    -- Validate Order By
    SET @OrderBy = LOWER(ISNULL(@OrderBy, 'date'));
    IF @OrderBy NOT IN ('date','fund','investor','amount') SET @OrderBy = 'date';

    -- 1. Get View Currency ID
    DECLARE @ViewCurrencyID int;
    SELECT @ViewCurrencyID = Currency_Id FROM edw.currency WHERE symbol = @ViewCurrencyCode;
    
    -- 2. Get Filter Source Currency ID
    DECLARE @FilterSourceCurrencyID int = NULL;
    IF @FilterSourceCurrency IS NOT NULL AND @FilterSourceCurrency <> ''
        SELECT @FilterSourceCurrencyID = Currency_Id FROM edw.currency WHERE symbol = @FilterSourceCurrency;

    DECLARE @sql nvarchar(max) = N'
;WITH investor_lookup AS
(
    SELECT edw_key
    FROM edw.global_edw_key_to_iqid
    WHERE source_table = ''investors''
      AND source_table_col_val = @SourceTableVolVal
),
fundtype_exclude AS
(
    SELECT LTRIM(RTRIM(value)) AS fund_type
    FROM STRING_SPLIT(@FundTypesExclude, '','')
    WHERE @FundTypesExclude IS NOT NULL AND LTRIM(RTRIM(value)) <> ''''
)
SELECT TOP (@EffectiveMaxRows)
    f.Metric_ID                          AS metric_id,
    @MetricName                          AS metric_name,
    NULL                                 AS investor_name_id, -- Not directly available in valuation table?

    ''' + @ViewCurrencyCode + '''        AS metric_currency_name,
    ''' + @ViewCurrencyCode + '''        AS metric_currency_code,

    cal.calendar_date                    AS transaction_date,

    -- Company & Investment Info
    c.Short_Name                         AS company_name,
    inv.investment_year                  AS investment_year,
    ind.industry_description             AS industry,
    ind.industry_broad_code              AS industry_broad,
    ind.industry_category_code           AS industry_category,
    
    fund.Short_Name                      AS fund_name,
    fund.AIV_fund_group_id               AS aiv_fund_group_id,
    fund2.Type                           AS fund_type,

    jc.name                              AS join_currency_name,

    -- FX CONVERSION APPLIED HERE: Source -> Fund -> View
    SUM(f.amount * ISNULL(xr1.Rate, 1) * ISNULL(xr2.Rate, 1))   AS aggregated_amount

FROM edw.fact_company_valuation f
JOIN edw.dim_fund fund
  ON f.fund_id = fund.fund_id
LEFT JOIN edw.dim_fund fund2
  ON f.fund_id = fund2.fund_id
LEFT JOIN edw.calendar cal
  ON f.date_id = cal.date_id
LEFT JOIN edw.currency jc
  ON f.currency_id = jc.Currency_Id

-- JOIN: Investment (Link Fund-Company)
LEFT JOIN edw.investment inv
  ON f.direct_company_id = inv.direct_company_id
  AND f.fund_id = inv.fund_id

-- JOIN: Company
LEFT JOIN edw.company c
  ON f.direct_company_id = c.company_id

-- JOIN: Industry
LEFT JOIN edw.industry ind
  ON inv.industry_id = ind.industry_id

-- FX JOIN 1
LEFT JOIN edw.exchange_rates xr1
  ON xr1.From_Currency_ID = f.currency_id
  AND xr1.To_Currency_ID = f.fund_currency_id -- Assuming fund_currency_id exists in fact_company_valuation?
  AND xr1.Date = cal.calendar_date

-- FX JOIN 2
LEFT JOIN edw.exchange_rates xr2
  ON xr2.From_Currency_ID = f.fund_currency_id -- verify column name
  AND xr2.To_Currency_ID = @ViewCurrencyID
  AND xr2.Date = cal.calendar_date

WHERE
    f.date_id <> -1
    AND f.date_id BETWEEN @StartKey AND @EndKey
';

    -- Optional filters
    IF @FilterSourceCurrencyID IS NOT NULL
        SET @sql += N' AND f.currency_id = @FilterSourceCurrencyID ';

    IF @AIVFundGroupID IS NOT NULL
        SET @sql += N' AND fund.AIV_fund_group_id = @AIVFundGroupID ';

    -- Fund type exclusion
    SET @sql += N'
    AND NOT EXISTS (SELECT 1 FROM fundtype_exclude fe WHERE fe.fund_type = fund2.Type)';
    
    -- New Company Filters (Mapped to verified tables)
    IF @CompExpGeoBroad IS NOT NULL SET @sql += N' AND f.company_geography_code = @CompExpGeoBroad '; -- Using fact table col per screenshot
    IF @CompExpGeoCountry IS NOT NULL SET @sql += N' AND inv.country_id = @CompExpGeoCountry '; -- From Investment table
    IF @CompExpIndId IS NOT NULL SET @sql += N' AND ind.industry_id = @CompExpIndId ';
    IF @CompExpIndBroad IS NOT NULL SET @sql += N' AND ind.industry_broad_code = @CompExpIndBroad ';
    IF @CompExpIndCategory IS NOT NULL SET @sql += N' AND ind.industry_category_code = @CompExpIndCategory ';
    -- IF @CompExpInvType IS NOT NULL SET @sql += N' AND c.Investment_Type = @CompExpInvType '; -- Not clear in images, commenting out
    IF @CompExpInvYear IS NOT NULL SET @sql += N' AND inv.investment_year = @CompExpInvYear ';
    -- IF @CompExpIsPublic IS NOT NULL SET @sql += N' AND c.Is_Public = @CompExpIsPublic '; -- Not clear, commenting out
    IF @CompExpStage IS NOT NULL SET @sql += N' AND inv.stage_code = @CompExpStage ';
    IF @CompExpStageBroad IS NOT NULL SET @sql += N' AND inv.broad_stage = @CompExpStageBroad ';

    SET @sql += N'
GROUP BY
    f.Metric_ID,
    cal.calendar_date,
    c.Short_Name,
    inv.investment_year,
    ind.industry_description,
    ind.industry_broad_code,
    ind.industry_category_code,
    fund.Short_Name,
    fund.AIV_fund_group_id,
    fund2.Type,
    jc.name
ORDER BY ';

    SET @sql += CASE @OrderBy
        WHEN 'date'     THEN N' cal.calendar_date '
        WHEN 'fund'     THEN N' fund.Short_Name '
        WHEN 'amount'   THEN N' SUM(f.amount * ISNULL(xr1.Rate, 1) * ISNULL(xr2.Rate, 1)) ' 
        ELSE                 N' cal.calendar_date '
    END + N' DESC;';

    EXEC sp_executesql
        @sql,
        N'@MetricName varchar(100),
          @SourceTableVolVal varchar(100),
          @StartKey int, @EndKey int,
          @FilterSourceCurrencyID int,
          @InvestorGroupID int,
          @AIVFundGroupID int,
          @FundTypesExclude varchar(400),
          @ViewCurrencyCode varchar(20),
          @ViewCurrencyID int,
          @EffectiveMaxRows int,
          @CompExpGeoBroad varchar(100),
          @CompExpGeoCountry varchar(100),
          @CompExpIndId varchar(100),
          @CompExpIndBroad varchar(100),
          @CompExpIndCategory varchar(100),
          @CompExpInvType varchar(100),
          @CompExpInvYear int,
          @CompExpIsPublic bit,
          @CompExpStage varchar(100),
          @CompExpStageBroad varchar(100)',
        @MetricName=@MetricName,
        @SourceTableVolVal=@SourceTableVolVal,
        @StartKey=@StartKey, @EndKey=@EndKey,
        @FilterSourceCurrencyID=@FilterSourceCurrencyID,
        @InvestorGroupID=@InvestorGroupID,
        @AIVFundGroupID=@AIVFundGroupID,
        @FundTypesExclude=@FundTypesExclude,
        @ViewCurrencyCode=@ViewCurrencyCode,
        @ViewCurrencyID=@ViewCurrencyID,
        @EffectiveMaxRows=@EffectiveMaxRows,
        @CompExpGeoBroad=@CompExpGeoBroad,
        @CompExpGeoCountry=@CompExpGeoCountry,
        @CompExpIndId=@CompExpIndId,
        @CompExpIndBroad=@CompExpIndBroad,
        @CompExpIndCategory=@CompExpIndCategory,
        @CompExpInvType=@CompExpInvType,
        @CompExpInvYear=@CompExpInvYear,
        @CompExpIsPublic=@CompExpIsPublic,
        @CompExpStage=@CompExpStage,
        @CompExpStageBroad=@CompExpStageBroad;
END
GO
