CREATE OR ALTER PROCEDURE edw.pr_usp_my_performance
    (
        @MetricName         varchar(100) = 'Pro-Rata Distribution',
        @SourceTableVolVal  varchar(100),
        -- Dates Removed
        @ViewCurrencyCode   varchar(20)  = 'USD',     
        @FilterSourceCurrency varchar(20)= NULL,      
        @FundTypesExclude   varchar(400) = NULL,      
        @InvestorGroupID    int          = NULL,
        @AIVFundGroupID     int          = NULL,
        @MaxRows            int          = 0,
        @OrderBy            varchar(20)  = 'date'
    )
AS
BEGIN
    SET NOCOUNT ON;

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
    f.investor_name_id                   AS investor_name_id,

    ''' + @ViewCurrencyCode + '''        AS metric_currency_name,
    ''' + @ViewCurrencyCode + '''        AS metric_currency_code,

    cal.calendar_date                    AS transaction_date,

    CASE WHEN inv.Part_Of_HV_Staff = 1
         THEN 1444282
         ELSE inv.Parent_Investor_Id
    END                                  AS investor_group_id,

    CASE WHEN (grp.Part_Of_HV_Staff = 1 AND grp.investor_name_id <> -1)
         THEN ''(Restricted)''
         ELSE grp.Short_Name
    END                                  AS investor_group_name,

    inv.Short_Name                       AS investor_name,
    inv.Address_Region                   AS investor_region,

    it.Short_Name                        AS investor_type_name,
    fund.Short_Name                      AS fund_name,
    fund.AIV_fund_group_id               AS aiv_fund_group_id,
    fund2.Type                           AS fund_type,

    jc.name                              AS join_currency_name,

    -- FX CONVERSION APPLIED HERE: Source -> Fund -> View
    SUM(f.amount * ISNULL(xr1.Rate, 1) * ISNULL(xr2.Rate, 1))   AS aggregated_amount

FROM edw.fact_investor_transactions f
JOIN investor_lookup il
  ON f.investor_name_id = il.edw_key
JOIN edw.dim_investor inv
  ON f.investor_name_id = inv.investor_name_id
JOIN edw.dim_fund fund
  ON f.fund_id = fund.fund_id
LEFT JOIN edw.dim_fund fund2
  ON f.fund_id = fund2.fund_id
LEFT JOIN edw.calendar cal
  ON f.date_id = cal.date_id
LEFT JOIN edw.dim_investor_type it
  ON it.investor_type = f.investor_type
LEFT JOIN edw.currency jc
  ON f.fund_currency_id = jc.Currency_Id
LEFT JOIN edw.dim_investor grp
  ON (CASE WHEN inv.Part_Of_HV_Staff=1 THEN 1444282 ELSE inv.Parent_Investor_Id END) = grp.investor_name_id
  
-- FX JOIN 1: Source (f.currency_id) -> Fund (f.fund_currency_id)
LEFT JOIN edw.exchange_rates xr1
  ON xr1.From_Currency_ID = f.currency_id
  AND xr1.To_Currency_ID = f.fund_currency_id
  AND xr1.Date = cal.calendar_date

-- FX JOIN 2: Fund (f.fund_currency_id) -> View (@ViewCurrencyID)
LEFT JOIN edw.exchange_rates xr2
  ON xr2.From_Currency_ID = f.fund_currency_id
  AND xr2.To_Currency_ID = @ViewCurrencyID
  AND xr2.Date = cal.calendar_date

WHERE
    f.exclude_transaction = 0
    AND f.date_id <> -1
    -- removed date filter
';

    -- Optional filters
    IF @FilterSourceCurrencyID IS NOT NULL
        SET @sql += N' AND f.currency_id = @FilterSourceCurrencyID ';

    IF @InvestorGroupID IS NOT NULL
        SET @sql += N' AND (CASE WHEN inv.Part_Of_HV_Staff=1 THEN 1444282 ELSE inv.Parent_Investor_Id END) = @InvestorGroupID ';

    IF @AIVFundGroupID IS NOT NULL
        SET @sql += N' AND fund.AIV_fund_group_id = @AIVFundGroupID ';

    -- Fund type exclusion
    SET @sql += N'
    AND NOT EXISTS (SELECT 1 FROM fundtype_exclude fe WHERE fe.fund_type = fund2.Type)';

    SET @sql += N'
GROUP BY
    f.Metric_ID,
    f.investor_name_id,
    cal.calendar_date,
    CASE WHEN inv.Part_Of_HV_Staff=1 THEN 1444282 ELSE inv.Parent_Investor_Id END,
    CASE WHEN (grp.Part_Of_HV_Staff=1 AND grp.investor_name_id <> -1) THEN ''(Restricted)'' ELSE grp.Short_Name END,
    inv.Address_Region,
    inv.Short_Name,
    it.Short_Name,
    fund.Short_Name,
    fund.AIV_fund_group_id,
    fund2.Type,
    jc.name
ORDER BY ';

    SET @sql += CASE @OrderBy
        WHEN 'date'     THEN N' cal.calendar_date '
        WHEN 'fund'     THEN N' fund.Short_Name '
        WHEN 'investor' THEN N' inv.Short_Name '
        WHEN 'amount'   THEN N' SUM(f.amount * ISNULL(xr1.Rate, 1) * ISNULL(xr2.Rate, 1)) ' 
        ELSE                 N' cal.calendar_date '
    END + N' DESC;';

    EXEC sp_executesql
        @sql,
        N'@MetricName varchar(100),
          @SourceTableVolVal varchar(100),
          @FilterSourceCurrencyID int,
          @InvestorGroupID int,
          @AIVFundGroupID int,
          @FundTypesExclude varchar(400),
          @ViewCurrencyCode varchar(20),
          @ViewCurrencyID int,
          @EffectiveMaxRows int',
        @MetricName=@MetricName,
        @SourceTableVolVal=@SourceTableVolVal,
        @FilterSourceCurrencyID=@FilterSourceCurrencyID,
        @InvestorGroupID=@InvestorGroupID,
        @AIVFundGroupID=@AIVFundGroupID,
        @FundTypesExclude=@FundTypesExclude,
        @ViewCurrencyCode=@ViewCurrencyCode,
        @ViewCurrencyID=@ViewCurrencyID,
        @EffectiveMaxRows=@EffectiveMaxRows;
END
GO
