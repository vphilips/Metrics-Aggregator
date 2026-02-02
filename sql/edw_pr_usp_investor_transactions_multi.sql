CREATE OR ALTER PROCEDURE edw.pr_usp_investor_transactions_multi
    (
        @MetricName         varchar(100) = 'Pro-Rata Distribution',
        @SourceTableVolVal  varchar(max), -- CSV of Accounts
        @StartDate          date         = '2012-01-01',
        @EndDate            date         = '2026-01-19',
        @AttributeList      varchar(max) = NULL, -- CSV of Attributes to select
        @ViewCurrencyID     int          = 1 
    )
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @StartKey int = CONVERT(int, CONVERT(char(8), @StartDate, 112));
    DECLARE @EndKey   int = CONVERT(int, CONVERT(char(8), @EndDate, 112));

    -- 1. Base Query with Common Joins
    DECLARE @sql nvarchar(max);
    
    SET @sql = N'
    ;WITH investor_lookup AS
    (
        SELECT edw_key
        FROM edw.global_edw_key_to_iqid
        WHERE source_table = ''investors''
          AND source_table_col_val IN (SELECT LTRIM(RTRIM(value)) FROM STRING_SPLIT(@SourceTableVolVal, '',''))
    )
    SELECT 
        -- Always include Key Metrics
        f.Metric_ID                          AS [Metric ID],
        @MetricName                          AS [Metric Name],
        
        -- Default Amount (Aggregated)
        SUM(f.amount * ISNULL(xr1.fx_rate, 1) * ISNULL(xr2.fx_rate, 1))   AS [Amount]
    ';

    -- 2. Dynamic Columns based on @AttributeList
    -- We assume the user passes display names, we map to DB columns.
    
    IF @AttributeList IS NOT NULL AND @AttributeList <> ''
    BEGIN
        -- Account (Investor)
        IF CHARINDEX('Account', @AttributeList) > 0 OR CHARINDEX('Investor', @AttributeList) > 0
            SET @sql += N', inv.Short_Name AS [Account] ';

        -- Client (Group)
        IF CHARINDEX('Client', @AttributeList) > 0
            SET @sql += N', CASE WHEN (grp.Part_Of_HV_Staff = 1 AND grp.investor_name_id <> -1) THEN ''(Restricted)'' ELSE grp.Short_Name END AS [Client] ';

        -- Fund
        IF CHARINDEX('Fund', @AttributeList) > 0 OR CHARINDEX('Fund Investor Presentation AIV', @AttributeList) > 0
            SET @sql += N', fund.Short_Name AS [Fund] ';

        -- Date
        IF CHARINDEX('Date', @AttributeList) > 0 OR CHARINDEX('Investor Transaction Date', @AttributeList) > 0
            SET @sql += N', cal.calendar_date AS [Transaction Date] ';

        -- Quarter
        IF CHARINDEX('Quarter', @AttributeList) > 0 OR CHARINDEX('Investor Transaction Quarter', @AttributeList) > 0
            SET @sql += N', cal.calendar_quarter AS [Transaction Quarter] '; -- Assuming calendar_quarter exists, or derive it

        -- Currency
        IF CHARINDEX('Currency', @AttributeList) > 0
            SET @sql += N', jc.name AS [Currency] ';
            
        -- Region
        IF CHARINDEX('Region', @AttributeList) > 0
            SET @sql += N', inv.Address_Region AS [Region] ';
    END
    ELSE
    BEGIN
        -- Default Columns if no attributes selected (Fall back to standard view)
        SET @sql += N', inv.Short_Name AS [Account], fund.Short_Name AS [Fund], cal.calendar_date AS [Date] ';
    END

    -- 3. FROM Clause
    SET @sql += N'
    FROM edw.fact_investor_transactions f
    JOIN investor_lookup il ON f.investor_name_id = il.edw_key
    JOIN edw.dim_investor inv ON f.investor_name_id = inv.investor_name_id
    JOIN edw.dim_fund fund ON f.fund_id = fund.fund_id
    LEFT JOIN edw.calendar cal ON f.date_id = cal.date_id
    LEFT JOIN edw.currency jc ON f.fund_currency_id = jc.Currency_Id
    LEFT JOIN edw.dim_investor grp ON (CASE WHEN inv.Part_Of_HV_Staff=1 THEN 1444282 ELSE inv.Parent_Investor_Id END) = grp.investor_name_id
    
    -- FX Joins
    LEFT JOIN edw.exchange_rates xr1 ON xr1.from_currency_id = f.currency_id AND xr1.to_currency_id = f.fund_currency_id AND xr1.date_id = cal.date_id
    LEFT JOIN edw.exchange_rates xr2 ON xr2.from_currency_id = f.fund_currency_id AND xr2.to_currency_id = @ViewCurrencyID AND xr2.date_id = cal.date_id

    WHERE f.exclude_transaction = 0
      AND f.date_id BETWEEN @StartKey AND @EndKey
    ';

    -- 4. GROUP BY Calculation
    -- We need to group by all non-aggregated columns.
    -- This is tricky with dynamic SQL inside a single string.
    -- Strategy: We repeat the conditional logic for GROUP BY.
    
    SET @sql += N' GROUP BY f.Metric_ID ';

    IF @AttributeList IS NOT NULL AND @AttributeList <> ''
    BEGIN
        IF CHARINDEX('Account', @AttributeList) > 0 OR CHARINDEX('Investor', @AttributeList) > 0
            SET @sql += N', inv.Short_Name ';

        IF CHARINDEX('Client', @AttributeList) > 0
            SET @sql += N', CASE WHEN (grp.Part_Of_HV_Staff = 1 AND grp.investor_name_id <> -1) THEN ''(Restricted)'' ELSE grp.Short_Name END ';

        IF CHARINDEX('Fund', @AttributeList) > 0 OR CHARINDEX('Fund Investor Presentation AIV', @AttributeList) > 0
            SET @sql += N', fund.Short_Name ';

        IF CHARINDEX('Date', @AttributeList) > 0 OR CHARINDEX('Investor Transaction Date', @AttributeList) > 0
            SET @sql += N', cal.calendar_date ';

        IF CHARINDEX('Quarter', @AttributeList) > 0 OR CHARINDEX('Investor Transaction Quarter', @AttributeList) > 0
            SET @sql += N', cal.calendar_quarter ';

        IF CHARINDEX('Currency', @AttributeList) > 0
            SET @sql += N', jc.name ';
            
        IF CHARINDEX('Region', @AttributeList) > 0
            SET @sql += N', inv.Address_Region ';
    END
    ELSE
    BEGIN
        SET @sql += N', inv.Short_Name, fund.Short_Name, cal.calendar_date ';
    END

    -- 5. Execution
    EXEC sp_executesql @sql,
        N'@MetricName varchar(100), @SourceTableVolVal varchar(max), @StartKey int, @EndKey int, @ViewCurrencyID int',
        @MetricName, @SourceTableVolVal, @StartKey, @EndKey, @ViewCurrencyID;
END
GO
