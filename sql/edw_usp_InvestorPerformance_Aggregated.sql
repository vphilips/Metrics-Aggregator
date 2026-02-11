USE [Harbourview_EDW_PRJ3]
GO
/****** Object:  StoredProcedure [edw].[usp_InvestorPerformance_Aggregated]    Script Date: 2/11/2026 7:17:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER   PROC [edw].[usp_InvestorPerformance_Aggregated]
(
    @Metric                 nvarchar(200),          -- single metric name
    @InvestorNameId         int,                    -- single investor_name_id (always one)
    @FundIdCsv              nvarchar(max) = NULL,   -- csv of fund_id
    @AsOfDate               int,                    -- yyyymmdd (report as-of)
    @OutputFieldsCsv        nvarchar(max) = NULL,   -- csv of output dims (plain labels); NULL/empty => total only
    @ReportingCurrencyId    int                     -- convert Fund CCY -> this CCY BEFORE aggregating
)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    ----------------------------------------------------------------------------------
    -- Basic validation
    ----------------------------------------------------------------------------------
    IF @InvestorNameId IS NULL
        THROW 50200, '@InvestorNameId is required (single investor).', 1;

    ----------------------------------------------------------------------------------
    -- Constants / defaults
    ----------------------------------------------------------------------------------
    DECLARE @AsOfDateDt date = CONVERT(date, CONVERT(char(8), @AsOfDate), 112);
    DECLARE @AivSubPerspectiveId int = 14813193; -- per requirement

    ----------------------------------------------------------------------------------
    -- Fund filter flags (handled inside dynamic SQL; no table variable)
    ----------------------------------------------------------------------------------
    DECLARE @HasFundFilter bit = 
        CASE WHEN NULLIF(LTRIM(RTRIM(@FundIdCsv)), N'') IS NULL THEN 0 ELSE 1 END;

    ----------------------------------------------------------------------------------
    -- Treat NULL/empty OutputFieldsCsv as "no dimensions" => total only
    ----------------------------------------------------------------------------------
    DECLARE @CleanCsv nvarchar(max) = NULLIF(LTRIM(RTRIM(@OutputFieldsCsv)), N'');
    DECLARE @NoDims bit = CASE WHEN @CleanCsv IS NULL THEN 1 ELSE 0 END;

    ----------------------------------------------------------------------------------
    -- Parse output fields -> labels (only if we have a CSV)
    ----------------------------------------------------------------------------------
    DECLARE @Out TABLE (label nvarchar(200) PRIMARY KEY);
    IF @NoDims = 0
    BEGIN
        INSERT INTO @Out(label)
        SELECT DISTINCT LTRIM(RTRIM(value))
        FROM STRING_SPLIT(@CleanCsv, ',')
        WHERE LTRIM(RTRIM(value)) <> '';
    END

    ----------------------------------------------------------------------------------
    -- Detect AIV usage (plain labels)
    ----------------------------------------------------------------------------------
    DECLARE @NeedAIV bit = 
    CASE WHEN @NoDims = 0 AND EXISTS
    (
        SELECT 1 
        FROM @Out
        WHERE label IN (N'Fund Investor Presentation AIV', N'Fund Investor Presentation AIV Currency')
    ) THEN 1 ELSE 0 END; -- toggles AIV joins/columns -- based on the original code's AIV toggle logic

    ----------------------------------------------------------------------------------
    -- Validate Metric (whitelist)
    ----------------------------------------------------------------------------------
    IF @Metric NOT IN
    (
        N'Investor Commitment',
        N'Investor Contribution',
        N'Investor Distribution',
        N'Investor Transfer of Interest',
        N'D/C (Sales)',
        N'Net Asset Value (Sales)',
        N'Total Value (Sales)',
        N'Gain/(Loss) (Sales)',
        N'Investor TV/C',
        N'Investor Unfunded',
        N'TV/F(Sales)',
        N'Investor IRR(Sales)',
        N'Investor IRR(Sales) - 1 Year',
        N'Investor IRR(Sales) - 3 Year',
        N'Investor IRR(Sales) - 5 Year',
        N'Investor IRR(Sales) - 10 Year'
    )
    BEGIN
        THROW 50202, 'Unsupported @Metric value.', 1;
    END;

    -- Disallow IRR with Investor Transaction Quarter
    IF @Metric LIKE N'Investor IRR%'
        AND @NoDims = 0
        AND EXISTS (SELECT 1 FROM @Out WHERE label = N'Investor Transaction Quarter')
    BEGIN
        THROW 50203, 'IRR metrics are not supported with output field Investor Transaction Quarter.', 1;
    END;

    ----------------------------------------------------------------------------------
    -- Build dynamic SELECT/GROUP BY fragments
    --   Inner = raw columns (safe for GROUP BY / inside rollups)
    --   Outer = user-facing aliases (final select only)
    ----------------------------------------------------------------------------------
    DECLARE @SelectDimsInner nvarchar(max) = N'';
    DECLARE @SelectDimsOuter nvarchar(max) = N'';
    DECLARE @GroupDims       nvarchar(max) = N'';

    IF @NoDims = 0
    BEGIN
        ;WITH map AS
        (
            -- label, sel_inner, sel_outer, grp
            SELECT N'Investor' AS label,
                   N't.investor_name_id, t.Investor' AS sel_inner,
                   N't.investor_name_id, t.Investor' AS sel_outer,
                   N't.investor_name_id, t.Investor' AS grp
            UNION ALL
            SELECT N'Fund',
                   N't.Fund_Id, t.Fund' AS sel_inner,
                   N't.Fund_Id, t.Fund' AS sel_outer,
                   N't.Fund_Id, t.Fund' AS grp
            UNION ALL
            -- AIV TYPE
            SELECT N'Fund Investor Presentation AIV',
                   N't.fund_investor_presentation_aiv_type'                   AS sel_inner,
                   N't.fund_investor_presentation_aiv_type AS [Fund Investor Presentation AIV]' AS sel_outer,
                   N't.fund_investor_presentation_aiv_type'                   AS grp
            UNION ALL
            -- AIV CURRENCY
            SELECT N'Fund Investor Presentation AIV Currency',
                   N't.fund_investor_presentation_aiv_currency_id, t.fund_investor_presentation_aiv_currency_name' AS sel_inner,
                   N't.fund_investor_presentation_aiv_currency_id, t.fund_investor_presentation_aiv_currency_name' AS sel_outer,
                   N't.fund_investor_presentation_aiv_currency_id, t.fund_investor_presentation_aiv_currency_name' AS grp
            UNION ALL
            SELECT N'Investor Transaction Quarter',
                   N't.txn_quarter' AS sel_inner,
                   N't.txn_quarter' AS sel_outer,
                   N't.txn_quarter' AS grp
        )
        SELECT 
            @SelectDimsInner = STRING_AGG(m.sel_inner, N', '),
            @SelectDimsOuter = STRING_AGG(m.sel_outer, N', '),
            @GroupDims       = STRING_AGG(m.grp,       N', ')
        FROM @Out o
        JOIN map m ON m.label = o.label;

        -- If CSV was provided but none of the labels were valid, error out
        IF NULLIF(@SelectDimsInner, N'') IS NULL
            THROW 50204, 'No valid output fields found in @OutputFieldsCsv.', 1;
    END

    -- Build context-aware clauses so commas / GROUP BY appear only when needed
    DECLARE @SelectListInner nvarchar(max) = CASE WHEN @NoDims = 1 THEN N'' ELSE @SelectDimsInner + N', ' END;
    DECLARE @SelectListOuter nvarchar(max) = CASE WHEN @NoDims = 1 THEN N'' ELSE @SelectDimsOuter + N', ' END;
    DECLARE @GroupByClause   nvarchar(max) = CASE WHEN @NoDims = 1 THEN N'' ELSE N'GROUP BY ' + @GroupDims END;

    ----------------------------------------------------------------------------------
    -- Dynamic SQL (two branches)
    ----------------------------------------------------------------------------------
    DECLARE @Sql    nvarchar(max) = N'';
    DECLARE @IsIrr bit = CASE WHEN @Metric LIKE N'Investor IRR%' THEN 1 ELSE 0 END;

    IF @IsIrr = 0
    BEGIN
        ------------------------------------------------------------------------------
        -- Non-IRR branch
        ------------------------------------------------------------------------------
        SET @Sql = N'
        ;WITH TxnBase AS
        (
            SELECT
                it.date_id AS txn_date_id,
                it.Fund_Id,
                it.investor_name_id,
                it.fund_currency_id,
                it.Currency_Id,
                it.Metric_ID,
                it.amount,
                it.exclude_transaction,
                it.is_transfered,
                it.is_transfer_within_group,
                f.Lock_Date_Eqt,
                f.Short_Name AS Fund,
                CASE WHEN inv.Part_Of_HV_Staff = 1 AND inv.investor_name_id <> -1 THEN ''(Restricted)'' ELSE inv.Short_Name END AS Investor'
                + CASE WHEN @NeedAIV = 1 THEN N',
                a16.[Type] AS fund_investor_presentation_aiv_type,
                a16.Currency_Id AS fund_investor_presentation_aiv_currency_id,
                cur.[Name] AS fund_investor_presentation_aiv_currency_name'
                ELSE N'' END + N',
                CASE
                    WHEN ' + CASE WHEN (@NoDims = 0 AND EXISTS (SELECT 1 FROM @Out WHERE label = N'Investor Transaction Quarter')) THEN '1' ELSE '0' END + N' = 1
                    THEN CONCAT(CONVERT(char(4), it.date_id/10000), ''-Q'', (((it.date_id/100)%100 - 1) / 3) + 1)
                    ELSE NULL
                END AS txn_quarter
            FROM bi.fact_investor_transactions it
            JOIN bi.dim_fund f
              ON it.Fund_Id = f.Fund_Id
            LEFT JOIN bi.dim_investor inv
              ON it.investor_name_id = inv.investor_name_id'
              + CASE WHEN @NeedAIV = 1 THEN N'
            LEFT JOIN bi.dim_fund a16
              ON f.AIV_fund_group_id = a16.Fund_Id
            LEFT JOIN bi.currency cur
              ON a16.Currency_Id = cur.Currency_Id'
              ELSE N'' END + N'
            WHERE
                it.date_id <= @AsOfDate
                AND it.investor_name_id = @InvestorNameId
                AND f.[Type] NOT IN (''Third Party Investor'')
                AND (
                    @HasFundFilter = 0
                    OR it.Fund_Id IN (SELECT TRY_CONVERT(int, LTRIM(RTRIM(s.value)))
                                      FROM STRING_SPLIT(@FundIdCsv, '','') AS s
                                      WHERE TRY_CONVERT(int, LTRIM(RTRIM(s.value))) IS NOT NULL)
                )'
                + CASE WHEN @NeedAIV = 1 THEN N'
                AND EXISTS
                (
                    SELECT 1
                    FROM bi.fact_fund_sub_perspective_funds fsp
                    WHERE fsp.fund_id = it.fund_id
                      AND fsp.fund_sub_perspective_id = @AivSubPerspectiveId
                      AND fsp.active_ind = 1
                      AND fsp.start_eff_date <= @AsOfDateDt
                      AND (fsp.end_eff_date IS NULL OR fsp.end_eff_date > @AsOfDateDt)
                )' 
                ELSE N'' END + N'
        ),
        TxnFx AS
        (
            SELECT
                b.*,
                @ReportingCurrencyId AS reporting_currency_id,
                CASE
                    WHEN b.fund_currency_id = @ReportingCurrencyId THEN 1.0
                    ELSE COALESCE(er.fx_rate, 0.0)
                END AS fx_rate
            FROM TxnBase b
            LEFT JOIN bi.exchange_rates er
              ON er.date_id = b.txn_date_id
              AND er.from_currency_id = b.Currency_Id
              AND er.to_currency_id = @ReportingCurrencyId
        ),
        TxnAggBase AS
        (
            SELECT
                ' + @SelectListInner + N'
                -- metrics
                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID IN (12,13,14,16,185,214,215)
                         THEN amount * fx_rate ELSE 0 END) AS InvestorContribution,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID IN (40,41,42,43,44,45,46,47,48,49,52,56,60,63,186,187,188,189,210,211)
                         THEN amount * fx_rate ELSE 0 END) AS InvestorDistribution,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID = 6
                         THEN amount * fx_rate ELSE 0 END) AS InvestorCommitment,

                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID IN (12,13,14,15,16,17,18,185,214,215)
                         THEN amount * fx_rate ELSE 0 END) AS Nav_Contrib_All,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID IN (40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63)
                         THEN amount * fx_rate ELSE 0 END) AS Nav_Dist_40_63,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID IN (456,457)
                         THEN amount * fx_rate ELSE 0 END) AS Nav_Redemptions,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID = 458
                         THEN amount * fx_rate ELSE 0 END) AS Nav_MgmtFee,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND txn_date_id <= ISNULL(Lock_Date_Eqt, 99991231)
                         AND Metric_ID IN
                         (
                            83,86,87,88,89,90,91,92,93,94,96,103,104,105,106,107,110,111,120,
                            108,109,164,119,95,85,98,100,102,121,122,123,124,125,126,127,128,129,
                            131,133,84,97,99,101,112,113,114,115,116,117,118,130,132,233,
                            134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150,
                            151,152,153,154,155,156,157,158,159,160,161,162,163,165,166,167,168,169,
                            170,171,172,173,174,175,176,177,234,235,236,237,238,239,240,241,245,246,
                            247,248,249,347,366,367,368,369,372,375,371,374,377,370,373,376,
                            378,379,380,381,382,383,384,385,386,398,399,400,346,
                            514,510,515,485,460,517,459,461,462,463,464,465
                         )
                         THEN amount * fx_rate ELSE 0 END) AS Nav_IncomeExpense,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND txn_date_id > ISNULL(Lock_Date_Eqt, 99991231)
                         AND is_transfered = 2
                         AND Metric_ID IN (134,135,136,137,138,139,140,141,142,234,143,144,145,146,147,148,149,150,
                            151,152,153,154,155,156,157,158,159,160,161,162,163,165,166,167,168,169,
                            170,171,172,173,174,175,176,177,378,379,380,381,383,385)
                         THEN amount * fx_rate ELSE 0 END) AS UnlockedIE_TransferredIn,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND txn_date_id > ISNULL(Lock_Date_Eqt, 99991231)
                         AND is_transfered = 1
                         AND Metric_ID IN (134,135,136,137,138,139,140,141,142,234,143,144,145,146,147,148,149,150,
                            151,152,153,154,155,156,157,158,159,160,161,162,163,165,166,167,168,169,
                            170,171,172,173,174,175,176,177,378,379,380,381,383,385)
                         THEN amount * fx_rate ELSE 0 END) AS UnlockedIE_TransferredOut,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID IN (326,327)
                         THEN amount * fx_rate ELSE 0 END) AS DistributionAdjustment,
                SUM(CASE WHEN exclude_transaction = 0 
                         AND Metric_ID IN (319,320)
                         THEN amount * fx_rate ELSE 0 END) AS ContributionAdjustment,
                (
                    SUM(CASE WHEN exclude_transaction = 0 
                             AND is_transfer_within_group = 0 AND is_transfered = 2 AND Metric_ID IN (15,16)
                             THEN amount * fx_rate ELSE 0 END)
                    -
                    SUM(CASE WHEN exclude_transaction = 0 
                             AND is_transfer_within_group = 0 AND is_transfered = 2 AND Metric_ID IN (60,61)
                             THEN amount * fx_rate ELSE 0 END)
                    -
                    SUM(CASE WHEN exclude_transaction = 0 
                             AND is_transfer_within_group = 0 AND is_transfered = 2 AND Metric_ID IN (55,56)
                             THEN amount * fx_rate ELSE 0 END)
                    +
                    SUM(CASE WHEN exclude_transaction = 0 
                             AND is_transfered = 2 AND Metric_ID = 164
                             THEN amount * fx_rate ELSE 0 END)
                ) AS TransferInNAV,
                (
                    SUM(CASE WHEN exclude_transaction = 0 
                             AND is_transfer_within_group = 0 AND is_transfered = 1 AND Metric_ID IN (15,16)
                             THEN amount * fx_rate ELSE 0 END)
                    -
                    SUM(CASE WHEN exclude_transaction = 0 
                             AND is_transfer_within_group = 0 AND is_transfered = 1 AND Metric_ID IN (60,61)
                             THEN amount * fx_rate ELSE 0 END)
                    -
                    SUM(CASE WHEN exclude_transaction = 0 
                             AND is_transfer_within_group = 0 AND is_transfered = 1 AND Metric_ID IN (55,56)
                             THEN amount * fx_rate ELSE 0 END)
                    +
                    SUM(CASE WHEN exclude_transaction = 0 
                             AND is_transfer_within_group = 0 AND is_transfered = 1
                             AND Metric_ID IN (134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150,151,152,153,154,
                                               155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174,175,176,177,
                                               234,378,379,380,381,382,383,384,385,386)
                             THEN amount * fx_rate ELSE 0 END)
                ) AS TransferOutNAVExternal
            FROM TxnFx
            ' + CASE WHEN @NoDims = 1 
                    THEN N'' -- no GROUP BY => total across all rows
                    ELSE N'GROUP BY ' + @GroupDims 
                END + N'
        ),
        Rollup AS
        (
            SELECT
                ' + @SelectListInner + N'
                SUM(InvestorCommitment)       AS InvestorCommitment,
                SUM(InvestorContribution)     AS InvestorContribution,
                SUM(InvestorDistribution)     AS InvestorDistribution,
                SUM(Nav_Contrib_All)          AS Nav_Contrib_All,
                SUM(Nav_Dist_40_63)           AS Nav_Dist_40_63,
                SUM(Nav_Redemptions)          AS Nav_Redemptions,
                SUM(Nav_MgmtFee)              AS Nav_MgmtFee,
                SUM(Nav_IncomeExpense)        AS Nav_IncomeExpense,
                SUM(UnlockedIE_TransferredIn) AS UnlockedIE_TransferredIn,
                SUM(UnlockedIE_TransferredOut)AS UnlockedIE_TransferredOut,
                SUM(DistributionAdjustment)   AS DistributionAdjustment,
                SUM(ContributionAdjustment)   AS ContributionAdjustment,
                SUM(TransferInNAV)            AS TransferInNAV,
                SUM(TransferOutNAVExternal)   AS TransferOutNAVExternal
            FROM TxnAggBase t
            ' + @GroupByClause + N'
        )
        SELECT 
            ' + @SelectListOuter + N'
            @Metric AS MetricName,
            CAST(
              CASE
                WHEN @Metric = N''Investor Commitment''           THEN t.InvestorCommitment
                WHEN @Metric = N''Investor Contribution''         THEN t.InvestorContribution
                WHEN @Metric = N''Investor Distribution''         THEN t.InvestorDistribution
                WHEN @Metric = N''Investor Transfer of Interest'' THEN t.TransferOutNAVExternal
                WHEN @Metric = N''Net Asset Value (Sales)''       THEN 
                    (t.Nav_Contrib_All - t.Nav_Dist_40_63 - t.Nav_Redemptions - t.Nav_MgmtFee + t.Nav_IncomeExpense)
                WHEN @Metric = N''Total Value (Sales)''           THEN 
                    (t.Nav_Contrib_All - t.Nav_Dist_40_63 - t.Nav_Redemptions - t.Nav_MgmtFee + t.Nav_IncomeExpense)
                    + t.InvestorDistribution
                WHEN @Metric = N''Gain/(Loss) (Sales)''           THEN 
                    ((t.Nav_Contrib_All - t.Nav_Dist_40_63 - t.Nav_Redemptions - t.Nav_MgmtFee + t.Nav_IncomeExpense)
                    + t.InvestorDistribution) - t.InvestorContribution
                WHEN @Metric = N''D/C (Sales)''                   THEN 
                    CASE WHEN NULLIF(t.InvestorContribution,0) IS NULL THEN 0.0
                         ELSE t.InvestorDistribution / NULLIF(t.InvestorContribution,0) END
                WHEN @Metric = N''Investor TV/C''                 THEN 
                    CASE WHEN NULLIF(t.InvestorContribution,0) IS NULL THEN 0.0
                         ELSE ((t.Nav_Contrib_All - t.Nav_Dist_40_63 - t.Nav_Redemptions - t.Nav_MgmtFee + t.Nav_IncomeExpense)
                               + t.InvestorDistribution) / NULLIF(t.InvestorContribution,0) END
                WHEN @Metric = N''Investor Unfunded''             THEN 
                    (t.InvestorCommitment - t.InvestorContribution)
                WHEN @Metric = N''TV/F(Sales)''                   THEN
                    CASE WHEN NULLIF((t.InvestorContribution + t.ContributionAdjustment + t.TransferInNAV),0) IS NULL THEN 0.0
                         ELSE
                            (
                              (t.Nav_Contrib_All - t.Nav_Dist_40_63 - t.Nav_Redemptions - t.Nav_MgmtFee + t.Nav_IncomeExpense)
                              + t.UnlockedIE_TransferredIn
                              + t.UnlockedIE_TransferredOut
                              + (t.InvestorDistribution + t.DistributionAdjustment)
                            ) / NULLIF((t.InvestorContribution + t.ContributionAdjustment + t.TransferInNAV),0) END
                ELSE NULL
              END
            AS float) AS MetricValue
        FROM Rollup t;'
    END
    ELSE
    BEGIN
        ------------------------------------------------------------------------------
        -- IRR branch
        ------------------------------------------------------------------------------
        SET @Sql = N'
        ;WITH Periods AS
        (
            SELECT ''1Y'' AS period_code, 1 AS years_back UNION ALL
            SELECT ''3Y'', 3 UNION ALL
            SELECT ''5Y'', 5 UNION ALL
            SELECT ''10Y'', 10 UNION ALL
            SELECT ''INC'', NULL
        ),
        Rollups AS
        (
            SELECT p.period_code, txn.date_id AS rollup_date_id, rd.date_id AS rollup_to_date_id, -1 AS irr_calc_seq
            FROM Periods p
            JOIN bi.calendar rd ON rd.date_id = @AsOfDate
            JOIN bi.calendar txn 
              ON p.years_back IS NOT NULL
              AND txn.calendar_date = EOMONTH(DATEADD(year, -p.years_back, rd.calendar_date))
            UNION ALL
            SELECT p.period_code, txn.date_id, rd.date_id, 0
            FROM Periods p
            JOIN bi.calendar rd ON rd.date_id = @AsOfDate
            JOIN bi.calendar txn
              ON p.years_back IS NOT NULL
              AND txn.date_id <= rd.date_id
              AND txn.calendar_date > EOMONTH(DATEADD(year, -p.years_back, rd.calendar_date))
            UNION ALL
            SELECT p.period_code, txn.date_id, rd.date_id, 0
            FROM Periods p
            JOIN bi.calendar rd ON rd.date_id = @AsOfDate
            JOIN bi.calendar txn
              ON p.years_back IS NULL
              AND txn.date_id <= rd.date_id
            UNION ALL
            SELECT p.period_code, rd.date_id, rd.date_id, 1
            FROM Periods p
            JOIN bi.calendar rd ON rd.date_id = @AsOfDate
        ),
        IrrAll AS
        (
            SELECT
                ' + @SelectListInner + N'
                r.rollup_to_date_id AS date_id,
                ISNULL(
                  wct.xirr(
                      CASE 
                        WHEN r.irr_calc_seq = -1 AND fi.Metric_ID = 227 THEN -1 * fi.amount * fx.fx_rate
                        WHEN r.irr_calc_seq =  0 AND fi.Metric_ID = 217 THEN -1 * fi.amount * fx.fx_rate
                        WHEN r.irr_calc_seq =  0 AND fi.Metric_ID = 218 THEN      fi.amount * fx.fx_rate
                        WHEN r.irr_calc_seq =  0 AND fi.Metric_ID = 220 THEN -1 * fi.amount * fx.fx_rate
                        WHEN r.irr_calc_seq =  0 AND fi.Metric_ID = 230 THEN -1 * fi.amount * fx.fx_rate
                        WHEN r.irr_calc_seq =  1 AND fi.Metric_ID = 227 THEN      fi.amount * fx.fx_rate
                        WHEN r.irr_calc_seq =  1 AND fi.Metric_ID = 332 THEN      fi.amount * fx.fx_rate
                      END,
                      CASE
                        WHEN fi.date_id = -1 THEN NULL
                        ELSE CONVERT(datetime, CONVERT(varchar(8), fi.date_id), 112)
                      END,
                      -0.01
                  ),
                  0
                ) AS irr_raw
            FROM bi.fact_irr_investor fi
            JOIN Rollups r
              ON fi.date_id = r.rollup_date_id
            JOIN bi.dim_fund f
              ON fi.Fund_Id = f.Fund_Id
            LEFT JOIN bi.dim_investor inv
              ON fi.investor_name_id = inv.investor_name_id'
              + CASE WHEN @NeedAIV = 1 THEN N'
            LEFT JOIN bi.dim_fund a16
              ON f.AIV_fund_group_id = a16.Fund_Id
            LEFT JOIN bi.currency cur
              ON a16.Currency_Id = cur.Currency_Id'
              ELSE N'' END + N'
            LEFT JOIN bi.exchange_rates er
              ON er.date_id = fi.date_id
              AND er.from_currency_id = fi.Currency_Id
              AND er.to_currency_id = @ReportingCurrencyId
            CROSS APPLY
            (
                SELECT CASE 
                         WHEN fi.Currency_Id = @ReportingCurrencyId THEN 1.0
                         ELSE COALESCE(er.fx_rate, 0.0)
                       END AS fx_rate
            ) fx
            WHERE 
                r.rollup_to_date_id = @AsOfDate
                AND fi.investor_name_id = @InvestorNameId
                AND f.[Type] NOT IN (''Third Party Investor'')
                AND (
                    @HasFundFilter = 0
                    OR fi.Fund_Id IN (SELECT TRY_CONVERT(int, LTRIM(RTRIM(s.value)))
                                      FROM STRING_SPLIT(@FundIdCsv, '','') AS s
                                      WHERE TRY_CONVERT(int, LTRIM(RTRIM(s.value))) IS NOT NULL)
                )' 
                + CASE WHEN @NeedAIV = 1 THEN N'
                AND EXISTS
                (
                    SELECT 1
                    FROM bi.fact_fund_sub_perspective_funds fsp
                    WHERE fsp.fund_id = fi.fund_id
                      AND fsp.fund_sub_perspective_id = @AivSubPerspectiveId
                      AND fsp.active_ind = 1
                      AND fsp.start_eff_date <= @AsOfDateDt
                      AND (fsp.end_eff_date IS NULL OR fsp.end_eff_date > @AsOfDateDt)
                )' 
                ELSE N'' END + N'
            ' + CASE WHEN @NoDims = 1 THEN N'' ELSE N'GROUP BY ' + @GroupDims + N', r.rollup_to_date_id' END + N'
        ),
        IrrPivot AS
        (
            SELECT 
                ' + @SelectListInner + N'
                date_id,
                MAX(CASE WHEN period_code = ''INC'' THEN irr_raw END)   AS InvestorIRRSales,
                MAX(CASE WHEN period_code = ''1Y''  THEN irr_raw END)   AS InvestorIRRSales_1Y,
                MAX(CASE WHEN period_code = ''3Y''  THEN irr_raw END)   AS InvestorIRRSales_3Y,
                MAX(CASE WHEN period_code = ''5Y''  THEN irr_raw END)   AS InvestorIRRSales_5Y,
                MAX(CASE WHEN period_code = ''10Y'' THEN irr_raw END)   AS InvestorIRRSales_10Y
            FROM IrrAll i
            JOIN Rollups r ON r.rollup_to_date_id = i.date_id
            '
            + CASE WHEN @NoDims = 1 THEN N'GROUP BY date_id'
                   ELSE N'GROUP BY ' + @GroupDims + N', date_id'
              END + N'
        ),
        IrrRollup AS
        (
            SELECT 
                ' + @SelectListInner + N'
                MAX(InvestorIRRSales)       AS InvestorIRRSales,
                MAX(InvestorIRRSales_1Y)    AS InvestorIRRSales_1Y,
                MAX(InvestorIRRSales_3Y)    AS InvestorIRRSales_3Y,
                MAX(InvestorIRRSales_5Y)    AS InvestorIRRSales_5Y,
                MAX(InvestorIRRSales_10Y)   AS InvestorIRRSales_10Y
            FROM IrrPivot t
            ' + @GroupByClause + N'
        )
        SELECT 
            ' + @SelectListOuter + N'
            @Metric AS MetricName,
            CAST(
              CASE
                WHEN @Metric = N''Investor IRR(Sales)''           THEN t.InvestorIRRSales
                WHEN @Metric = N''Investor IRR(Sales) - 1 Year''  THEN t.InvestorIRRSales_1Y
                WHEN @Metric = N''Investor IRR(Sales) - 3 Year''  THEN t.InvestorIRRSales_3Y
                WHEN @Metric = N''Investor IRR(Sales) - 5 Year''  THEN t.InvestorIRRSales_5Y
                WHEN @Metric = N''Investor IRR(Sales) - 10 Year'' THEN t.InvestorIRRSales_10Y
              END
            AS float) AS MetricValue
        FROM IrrRollup t;';
    END

    -- Execute the dynamic statement
    EXEC sys.sp_executesql 
        @Sql, 
        N'@AsOfDate int, 
          @AsOfDateDt date, 
          @Metric nvarchar(200),
          @InvestorNameId int,
          @ReportingCurrencyId int,
          @AivSubPerspectiveId int,
          @HasFundFilter bit,
          @FundIdCsv nvarchar(max)',
        @AsOfDate            = @AsOfDate,
        @AsOfDateDt          = @AsOfDateDt,
        @Metric              = @Metric,
        @InvestorNameId      = @InvestorNameId,
        @ReportingCurrencyId = @ReportingCurrencyId,
        @AivSubPerspectiveId = @AivSubPerspectiveId,
        @HasFundFilter       = @HasFundFilter,
        @FundIdCsv           = @FundIdCsv;
END
