CREATE OR ALTER PROCEDURE sp_GetCompanyDiversification
    (
        @Account varchar(100) = NULL,
        @Client varchar(100) = NULL,
        @Investor varchar(100) = NULL,
        -- Dates removed
        @ViewCurrencyCode varchar(20) = 'USD',
        @ViewingCoy varchar(100) = NULL,
        @Metric varchar(100) = NULL,
        @Date date = NULL,
        @AIVFundGroupID int = NULL,
        @CompExpGeoBroad varchar(100) = NULL,
        @CompExpGeoCountry varchar(100) = NULL,
        @CompExpIndId varchar(100) = NULL,
        @CompExpIndBroad varchar(100) = NULL,
        @CompExpIndCategory varchar(100) = NULL,
        @CompExpInvType varchar(100) = NULL,
        @CompExpInvYear int = NULL,
        @CompExpIsPublic bit = NULL,
        @CompExpStage varchar(100) = NULL,
        @CompExpStageBroad varchar(100) = NULL
    )
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Skeleton logic for now
    SELECT 
        'CompanyDiversification' as report_type,
        @Account as account,
        @Client as client,
        @Metric as metric
    -- Add actual query here
END
GO
