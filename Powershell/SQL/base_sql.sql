
/*** Changeable Values ***/
DECLARE @Version AS VARCHAR(10) = '{0}'
DECLARE @YearContext AS VARCHAR(10) = '{1}'
DECLARE @MonthNumber AS NVARCHAR(20) = '{2}'
DECLARE @ForecastType AS NVARCHAR(20) = '{3}'
DECLARE @FactTable AS NVARCHAR(25) = '{4}'
DECLARE @MCBOnly AS TINYINT = {5}
/*** Changeable Values ***/


/****
        Variable Notes
        @Version: Possibilities: 'Actuals', 'Forecast', 'Budget'. Note Budget and Forecast should be 'CY' Only.
        @YearContext: Possibilities:  'PY2' = 2 Years Ago, 'PY' = 1 Year Ago, 'CY' = Current Year, 'FY' = future Year.
        @MonthNumber: Either single integer like, '1' or '2' or  groups '1,2,3' or '1,2,3,4,5,6,7,8,9,10,11,12'
        @ForecastType: 'Current' or 'Historical'
        @FactTable: Options: 'byDetail', 'byProduct', 'byPC'
    ***/

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED
SET NOCOUNT ON

DECLARE @CurrentFiscalYearNumber AS INT = (SELECT
    FiscalYearNumber
FROM
    [FiscalMonth] B WITH (NOLOCK)
WHERE
            CurrentFiscalMonthFlagId = 1)

DECLARE @FiscalQuarter AS VARCHAR(10) = 'FY' + RIGHT(@CurrentFiscalYearNumber, 2)

IF @ForecastType = 'Current'
        SET @ForecastType = 1
    ELSE
        SET @ForecastType = 0

DECLARE @FY AS NVARCHAR(4)
IF @YearContext = 'CY'
        SET @FY = 'FY' + RIGHT(@CurrentFiscalYearNumber, 2)

IF @YearContext = 'PY'
        SET @FY = 'FY' + RIGHT(@CurrentFiscalYearNumber - 1, 2)

IF @YearContext = 'PY2'
        SET @FY = 'FY' + RIGHT(@CurrentFiscalYearNumber - 2, 2)

IF @YearContext = 'FY'
        SET @FY = 'FY' + RIGHT(@CurrentFiscalYearNumber + 1, 2)

-- Version
DECLARE @VersionIdLower AS INT
DECLARE @VersionIdUpper AS INT

IF @Version = 'Actuals'
        BEGIN
    SET @VersionIdLower = 14
    SET @VersionIdUpper = 14
END

IF @Version = 'Budget'
        BEGIN
    SET @VersionIdLower = 13
    SET @VersionIdUpper = 13
END

IF @Version = 'Forecast'
        BEGIN
    SET @VersionIdLower = 1
    SET @VersionIdUpper = 12
END

DECLARE @Budget AS NVARCHAR(MAX) = '
            /**** Budget Spend ****/
            SELECT
                A.[FiscalMonthId]               [FiscalMonthId],
                A.[ProfitCenterId]              [ProfitCenterId],
                A.[FinancialAccountId]          [FinancialAccountId],
                A.[SAPCompanyId]                [SAPCompanyId],
                A.[ProductFamilyId]             [ProductFamilyId],
                A.[AllocationRuleDetailId]      [AllocationRuleDetailId],
                A.[BudgetVersionId]             [BudgetVersionId],
                A.[InternalOrderId]             [InternalOrderId],
                A.[RestatementDetailId]         [RestatementDetailId],
                A.[BundleStatusId]              [BundleStatusId],
                A.[CurrencyId]                  [CurrencyId],
                A.[CurrencyTypeHierarchyId]     [CurrencyTypeHierarchyId],
                [ForecastVersionId]             = 13,
                /* Made up 13 = Budget (1-12 are Forecast)*/
                A.[ForecastTypeId]              [ForecastTypeId],
                A.[StaffingResourceTypeId]      [StaffingResourceTypeId],
                A.[BudgetAmount]                [Amount],
                A.[ExternalSegmentId]           [ExternalSegmentId],
                A.[ExternalSegmentBReportingId] [ExternalSegmentBReportingId]
            FROM [vwFactIncomeStatementBudget] A WITH (NOLOCK)
            LEFT JOIN [FiscalMonth] B WITH (NOLOCK)
                ON A.[FiscalMonthId] = B.[FiscalMonthId]
            WHERE B.[FiscalMonthNumber] IN (@MonthNumber)
                AND A.[CurrencyTypeHierarchyId] = 5

            UNION ALL

            /**** Budget Headcount ****/
            SELECT
                A.[FiscalMonthId]                       [FiscalMonthId],
                A.[ProfitCenterId]                      [ProfitCenterId],
                A.[FinancialAccountId]                  [FinancialAccountId],
                A.[SAPCompanyId]                        [SAPCompanyId],
                A.[ProductFamilyId]                     [ProductFamilyId],
                A.[AllocationRuleDetailId]              [AllocationRuleDetailId],
                A.[BudgetVersionId]                     [BudgetVersionId],
                A.[InternalOrderId]                     [InternalOrderId],
                A.[RestatementDetailId]                 [RestatementDetailId],
                A.[BundleStatusId]                      [BundleStatusId],
                A.[CurrencyId]                          [CurrencyId],
                [CurrencyTypeHierarchyId]               = 5,
                [ForecastVersionId]                     = 13,
                /* Made up 13 = Budget (1-12 are Forecast)*/
                A.[ForecastTypeId]                      [ForecastTypeId],
                A.[StaffingResourceTypeId]              [StaffingResourceTypeId],
                A.[BudgetPositionCount]                 [Amount],
                [ExternalSegmentId]                     = 0,
                [ExternalSegmentBReportingId]           = 0
            FROM [vwFactHeadCountActuals] AS A WITH (NOLOCK)
            LEFT JOIN [FiscalMonth] B WITH (NOLOCK)
                ON A.[FiscalMonthId] = B.[FiscalMonthId]
            WHERE B.[FiscalmonthNumber] IN (@MonthNumber)'

DECLARE @Forecast AS NVARCHAR(MAX) = '
            /*** Forecast Spending ***/
            SELECT
                A.[FiscalMonthId]                       [FiscalMonthId],
                A.[ProfitCenterId]                      [ProfitCenterId],
                A.[FinancialAccountId]                  [FinancialAccountId],
                A.[SAPCompanyId]                        [SAPCompanyId],
                A.[ProductFamilyId]                     [ProductFamilyId],
                A.[AllocationRuleDetailId]              [AllocationRuleDetailId],
                A.[BudgetVersionId]                     [BudgetVersionId],
                A.[InternalOrderId]                     [InternalOrderId],
                A.[RestatementDetailId]                 [RestatementDetailId],
                A.[BundleStatusId]                      [BundleStatusId],
                A.[CurrencyId]                          [CurrencyId],
                A.[CurrencyTypeHierarchyId]             [CurrencyTypeHierarchyId],
                A.[ForecastVersionId]                   [ForecastVersionId],
                A.[ForecastTypeId]                      [ForecastTypeId],
                A.[StaffingResourceTypeId]              [StaffingResourceTypeId],
                A.[ForecastAmount]                      [Amount],
                A.[ExternalSegmentId]                   [ExternalSegmentId],
                A.[ExternalSegmentBReportingId]         [ExternalSegmentBReportingId]
            FROM [Forecast]..[vwFactForecastActuals] A WITH (NOLOCK)
            LEFT JOIN [FiscalMonth] B WITH (NOLOCK)
                ON A.[FiscalMonthId] = B.[FiscalMonthId]
            /*Account Related*/
            LEFT JOIN [FinancialAccount] E WITH (NOLOCK)
                ON A.[FinancialAccountId] = E.[FinancialAccountId]
            LEFT JOIN [Forecast]..[ForecastVersion] fv
                ON fv.[ForecastVersionId] = A.[ForecastVersionId]
            WHERE
                B.[FiscalMonthNumber] IN (@MonthNumber)
                AND A.[CurrencyTypeHierarchyId] = 5
                AND ((E.[ChannelClassId] IN (37, 13) /* 16 = Expense; 13 = Headcount */ AND A.ForecastTypeId = 3 /* 3 = Microsoft */)
                OR E.[ChannelClassId] = 16 /*This is net revenue*/)
                AND A.[PostingType] = ''Forecast''
                AND fv.[CurrentForecastVersionFlagId] = @ForecastType

            UNION ALL

            /*** Forecast Headcount ***/
            SELECT
                A.[FiscalMonthId]                              [FiscalMonthId],
                A.[ProfitCenterId]                             [ProfitCenterId],
                A.[FinancialAccountId]                         [FinancialAccountId],
                A.[SAPCompanyId]                               [SAPCompanyId],
                A.[ProductFamilyId]                            [ProductFamilyId],
                A.[AllocationRuleDetailId]                     [AllocationRuleDetailId],
                A.[BudgetVersionId]                            [BudgetVersionId],
                A.[InternalOrderId]                            [InternalOrderId],
                A.[RestatementDetailId]                        [RestatementDetailId],
                A.[BundleStatusId]                             [BundleStatusId],
                A.[CurrencyId]                                 [CurrencyId],
                [CurrencyTypeHierarchyId]                       = 5,
                A.[ForecastVersionId]                          [ForecastVersionId],
                A.[ForecastTypeId]                             [ForecastTypeId],
                A.[StaffingResourceTypeId]                     [StaffingResourceTypeId],
                A.[ForecastPeopleCount]                        [Amount],
                [ExternalSegmentId]                             = 0,
                [ExternalSegmentBReportingId]                   = 0
            FROM [Forecast]..[vwFactForecastHeadCount] AS A WITH (NOLOCK)
            LEFT JOIN [Forecast]..[ForecastVersion] fv
                ON fv.[ForecastVersionId] = A.[ForecastVersionId]
            LEFT JOIN [FiscalMonth] B WITH (NOLOCK)
                ON A.[FiscalMonthId] = B.[FiscalMonthId]
            WHERE B.[FiscalMonthNumber] IN (@MonthNumber)
                AND fv.[CurrentForecastVersionFlagId] = @ForecastType
                AND (
                    B.[FiscalYearNumber] = @CurrentFiscalYearNumber + 1
                    OR ( (fv.[ForecastVersionId] = 3 AND B.[FiscalmonthNumber] > 2)   /* Sep Fcst */
                    OR (fv.[ForecastVersionId] = 4 AND B.[FiscalmonthNumber] > 3)     /* Oct Fcst */
                    OR (fv.[ForecastVersionId] = 6 AND B.[FiscalmonthNumber] > 5)     /* Dec Fcst */
                    OR (fv.[ForecastVersionId] = 7 AND B.[FiscalmonthNumber] > 6)     /*Jan Fcst */
                    OR (fv.[ForecastVersionId] = 9 AND B.[FiscalmonthNumber] > 8)     /* Mar Fcst */
                    OR (fv.[ForecastVersionId] = 10 AND B.[FiscalmonthNumber] > 9)    /* Apr Fcst */)
                )'

DECLARE @Actuals AS NVARCHAR(MAX) = '
            /**** Actuals Spend ****/
            SELECT
                A.[FiscalMonthId]               [FiscalMonthId],
                A.[ProfitCenterId]              [ProfitCenterId],
                A.[FinancialAccountId]          [FinancialAccountId],
                A.[SAPCompanyId]                [SAPCompanyId],
                A.[ProductFamilyId]             [ProductFamilyId],
                A.[AllocationRuleDetailId]      [AllocationRuleDetailId],
                A.[BudgetVersionId]             [BudgetVersionId],
                A.[InternalOrderId]             [InternalOrderId],
                A.[RestatementDetailId]         [RestatementDetailId],
                A.[BundleStatusId]              [BundleStatusId],
                A.[CurrencyId]                  [CurrencyId],
                A.[CurrencyTypeHierarchyId]     [CurrencyTypeHierarchyId],
                [ForecastVersionId]              = 14,
                /* Made up 14 = Actuals (1-12 are Forecast)*/
                [ForecastTypeId]                = 0,
                A.[StaffingResourceTypeId]      [StaffingResourceTypeId],
                A.[ActualAmount]                [Amount],
                A.[ExternalSegmentId]           [ExternalSegmentId],
                A.[ExternalSegmentBReportingId] [ExternalSegmentBReportingId]
            FROM [vwFactIncomeStatementActuals] A WITH (NOLOCK)
            LEFT JOIN [FiscalMonth] B WITH (NOLOCK)
                ON A.[FiscalMonthId] = B.[FiscalMonthId]
            WHERE B.[FiscalmonthNumber] IN (@MonthNumber)
                AND B.[FiscalYearName] = @FY
                AND A.[CurrencyTypeHierarchyId] = 5

            UNION ALL

            /**** Actuals Headcount ****/
            SELECT
               A.[FiscalMonthId]                [FiscalMonthId],
               A.[ProfitCenterId]               [ProfitCenterId],
               A.[FinancialAccountId]           [FinancialAccountId],
               A.[SAPCompanyId]                 [SAPCompanyId],
               A.[ProductFamilyId]              [ProductFamilyId],
               A.[AllocationRuleDetailId]       [AllocationRuleDetailId],
               A.[BudgetVersionId]              [BudgetVersionId],
               A.[InternalOrderId]              [InternalOrderId],
               A.[RestatementDetailId]          [RestatementDetailId],
               A.[BundleStatusId]               [BundleStatusId],
               A.[CurrencyId]                   [CurrencyId],
               [CurrencyTypeHierarchyId]        = 5,
               [ForecastVersionId]              = 14,
               /* Made up 14 = Actuals (1-12 are Forecast)*/
               A.[ForecastTypeId]               [ForecastTypeId],
               A.[StaffingResourceTypeId]       [StaffingResourceTypeId],
               A.[ActualPeopleCount]            [Amount],
               [ExternalSegmentId]              = 0,
               [ExternalSegmentBReportingId]    = 0
            FROM [vwFactHeadCountActuals] AS A WITH (NOLOCK)
            LEFT JOIN [FiscalMonth] B WITH (NOLOCK)
                ON A.[FiscalMonthId] = B.[FiscalMonthId]
            WHERE B.[FiscalmonthNumber] IN (@MonthNumber)
                AND B.[FiscalYearName] = @FY'

DECLARE @RawData AS NVARCHAR(MAX) = '
        SELECT * FROM (
        SELECT
            ISNULL(B.[FiscalMonthNumber], 0)            [FiscalMonthNumber],
            ISNULL(A.[ProfitCenterId], 0)               [ProfitCenterId],
            ISNULL(A.[SAPCompanyId], 0)                 [SAPCompanyId],
            ISNULL(A.[ProductFamilyId], 0)              [ProductFamilyId],
            ISNULL(A.[AllocationRuleDetailId], 0)       [AllocationRuleDetailId],
            ISNULL(A.[InternalOrderId], 0)              [InternalOrderId],
            ISNULL(A.[RestatementDetailId], 0)          [RestatementDetailId],
            ISNULL(A.[CurrencyId], 0)                   [CurrencyId],
            ISNULL(A.[CurrencyTypeHierarchyId], 0)      [CurrencyTypeHierarchyId],
            ISNULL(A.[ForecastTypeId], 0)               [ForecastTypeId],
            ISNULL(A.[StaffingResourceTypeId], 0)       [StaffingResourceTypeId],
            ISNULL(A.[ExternalSegmentId], 0)            [ExternalSegmentId],
            ISNULL(A.[ExternalSegmentBReportingId], 0)  [ExternalSegmentBReportingId],
            CASE
                    WHEN E.[ChannelClassId] = 16 THEN ROUND([Amount], 4) * -1
                    ELSE ROUND([Amount], 4)
            END                                         [Amount],
            (CASE SEC.[SECFunctionalAreaCode]
                WHEN ''FX'' THEN 2
                WHEN ''FR'' THEN 3
                WHEN ''FS'' THEN 4
                WHEN ''FG'' THEN 5
                WHEN ''FB'' THEN 6
                WHEN ''SI'' THEN 7
                WHEN ''ME'' THEN 8
                ELSE 9
            END) * 1000000 + ISNULL(A.[financialaccountId], 0)                                                                              [AccountId],
            CASE
                WHEN RIGHT(N.ChannelSubClassDesc, 3) != ''Adj'' AND FT.ForecastTypeDesc IN (''N/A'', ''CFO'', ''Microsoft'') THEN ''Yes''
                WHEN CF.ChannelOrgSummaryDesc IN (''WW - Consumer'', ''WW - OEM non-Field'') AND RIGHT(N.ChannelSubClassDesc, 3) = ''Adj'' AND FT.ForecastTypeDesc = ''Field'' THEN ''Yes''
                WHEN CF.ChannelOrgSummaryDesc IN (''WW - Consumer'', ''WW - OEM non-Field'') AND RIGHT(N.ChannelSubClassDesc, 3) = ''Adj'' AND FT.ForecastTypeDesc = ''N/A'' THEN ''Yes''
                WHEN CF.ChannelOrgSummaryDesc NOT IN (''WW - Consumer'', ''WW - OEM non-Field'') AND RIGHT(N.ChannelSubClassDesc, 3) = ''Adj'' AND FT.ForecastTypeDesc IN (''N/A'', ''CFO'', ''Microsoft'') THEN ''Yes'' ELSE ''No''
            END                                                                                                                             [FCST_TYPE],
            CONCAT(ISNULL(A.[forecastversionId], 0), ''-'', RIGHT(B.[FiscalYearNumber], 2), ''-'', ISNULL(B.[fiscalmonthNumber], 0))        [VersionId],
                /*
                SubClassId: 103 = Traffic Acquisition Costs (TRAC)
                SubClassId: 102 = Shared Operations (SHOP)
                ExecOrgId: 646 = Brand, Advertising and Research (FESB)
                ExecOrgDetailId: 10964 = One GDC
                External Segment: 3 = Cloud and Enterprise; 6 = Dynamics
                ProductDivision: 5674 = Outlook.com Advertising; 5980 = Outlook.com Subscription; 6206 OneDrive
                */
            CASE
                WHEN e.ChannelSubClassId = 103 AND M.ExecOrgId = 646 THEN 1
                WHEN e.FinancialAccountCode = ''724069'' AND M.ExecOrgDetailId = 10964 THEN 2
                WHEN a.ExternalSegmentId IN (3, 6) AND e.ChannelSubClassId = 102 THEN 3
                WHEN pd.ProductDivisionId IN (5674, 5980, 6206) THEN 4
                ELSE (CASE SEC.[SECFunctionalAreaCode]
                        WHEN ''FX'' THEN 2
                        WHEN ''FR'' THEN 3
                        WHEN ''FS'' THEN 4
                        WHEN ''FG'' THEN 5
                        WHEN ''FB'' THEN 6
                        WHEN ''SI'' THEN 7
                        WHEN ''ME'' THEN 8
                        ELSE 9
                END) * 1000000 + ISNULL(A.[financialaccountId], 0)
            END                                                                                                                             [CACKey],
            /*Additional Security*/
            ISNULL(M.SalesLocationId, 0)                    [SalesLocationId],
            ISNULL(CF.ChannelOrgSummaryId, 0)               [ChannelOrgSummaryId],
            ISNULL(M.ExecOrgSummaryId, 0)                   [ExecOrgSummaryId]
            FROM ( [[Inner Selection]] ) A
        LEFT JOIN [FiscalMonth] B WITH (NOLOCK)
            ON A.FiscalMonthId = B.FiscalMonthId
        /*Account Related*/
        LEFT JOIN [FinancialAccount] E WITH (NOLOCK)
            ON A.[FinancialAccountId] = E.[FinancialAccountId]
        LEFT JOIN [ChannelSubClass] n WITH (NOLOCK)
            ON E.[ChannelSubClassId] = n.[ChannelSubClassId]
        LEFT JOIN [vwMSRBundleStatus] J WITH (NOLOCK)
            ON A.[BundleStatusId] = J.[BundleStatusId]
        /*Profit Center Realated*/
        LEFT JOIN [ProfitCenter] M WITH (NOLOCK)
            ON A.ProfitCenterId = M.ProfitCenterId
        LEFT JOIN [SAPCompany] O WITH (NOLOCK)
            ON O.[SAPCompanyId] = A.[SAPCompanyId]
        LEFT JOIN [vwSecFunctionalAreaDenorm] SEC
            ON SEC.ProfitCenterId = A.ProfitCenterId
        LEFT JOIN [ChannelClass] AS CC
            ON CC.[ChannelClassId] = E.[ChannelClassId]
        LEFT JOIN [VwChannelFunctionDenorm] CF
            ON CF.ProfitCenterId = A.ProfitCenterId
        LEFT JOIN [Forecast].[dbo].[ForecastType] FT
            ON FT.ForecastTypeId = A.[forecasttypeId]
        /*Product Related*/
        LEFT JOIN [ProductFamily] pd WITH (NOLOCK)
            ON pd.ProductFamilyId = a.ProductFamilyId
        WHERE A.[CurrencyTypeHierarchyId] IN (5) /* 3 = USD and 5 = CD */
            AND J.[BundleStatusName] IN (''AS Shipped'')
            AND A.[BudgetVersionId] IN (0, 2) /* Reporting - 01'',''N/A'' */
            AND E.[ChannelClassId] IN (37, 16, 13) /* Expenses, Net Revenue, Headcount */
            AND E.[CostElementTypeId] IN (1, 0, 3) /* Excludes Intercompany*/
            AND O.[SAPCompanyTypeDetailId] IN (0, 2)
            AND ROUND(A.[Amount], 4) != 0
            AND A.[ForecastVersionId] BETWEEN @VersionIdLower AND @VersionIdUpper
            AND M.[ExecOrgSummaryId] IN (107) )  A
        [[Footer]]
        OPTION (HASH GROUP)'

/*Group By statements*/
DECLARE @byPCSQL_Header AS NVARCHAR(MAX) = '
        SELECT
            A.[FiscalMonthNumber],
            A.[ProfitCenterId],
            A.[CurrencyId],
            A.[CurrencyTypeHierarchyId],
            A.[ForecastTypeId],
            A.[StaffingResourceTypeId],
            A.[ExternalSegmentId],
            A.[ExternalSegmentBReportingId],
            A.[AccountId],
            A.[FCST_TYPE],
            A.[VersionId],
            A.[CACKey],
            /*Additional Security*/
            A.[SalesLocationId],
            A.[ChannelOrgSummaryId],
            A.[ExecOrgSummaryId],
            /*Amount*/
            ROUND(SUM(A.[Amount]), 4)       [Amount]
    FROM'

DECLARE @byPCSQL_Footer AS NVARCHAR(MAX) = '
        GROUP BY
            A.[FiscalMonthNumber],
            A.[ProfitCenterId],
            A.[CurrencyTypeHierarchyId],
            A.[CurrencyId],
            A.[ForecastTypeId],
            A.[StaffingResourceTypeId],
            A.[ExternalSegmentId],
            A.[ExternalSegmentBReportingId],
            A.[AccountId],
            A.[FCST_TYPE],
            A.[VersionId],
            A.[CACKey],
            A.[SalesLocationId],
            A.[ChannelOrgSummaryId],
            A.[ExecOrgSummaryId]
        HAVING
            ROUND(SUM(A.[Amount]), 4) <> 0'

DECLARE @byProdSQL_Header AS NVARCHAR(MAX) = '
        SELECT
            A.[FiscalMonthNumber],
            A.[ProductFamilyId],
            A.[CurrencyId],
            A.[CurrencyTypeHierarchyId],
            A.[ForecastTypeId],
            A.[StaffingResourceTypeId],
            A.[ExternalSegmentId],
            A.[ExternalSegmentBReportingId],
            A.[AccountId],
            A.[FCST_TYPE],
            A.[VersionId],
            A.[CACKey],
            /*Additional Security*/
            A.[SalesLocationId],
            A.[ChannelOrgSummaryId],
            A.[ExecOrgSummaryId],
            /*Amount*/
            ROUND(SUM(A.[Amount]), 4)       [Amount]
        FROM'

DECLARE @byProdSQL_Footer AS NVARCHAR(MAX) = '
        GROUP BY
            A.[FiscalMonthNumber],
            A.[ProductFamilyId],
            A.[CurrencyTypeHierarchyId],
            A.[CurrencyId],
            A.[ForecastTypeId],
            A.[StaffingResourceTypeId],
            A.[ExternalSegmentId],
            A.[ExternalSegmentBReportingId],
            A.[AccountId],
            A.[FCST_TYPE],
            A.[VersionId],
            A.[CACKey],
            A.[SalesLocationId],
            A.[ChannelOrgSummaryId],
            A.[ExecOrgSummaryId]
        HAVING ROUND(SUM(A.[Amount]), 4) <> 0'

/*Replace inner query with budget, actuals or forecast*/
DECLARE @SQL AS NVARCHAR(MAX) = ''
IF @Version = 'Budget'
        SET @SQL = REPLACE(@RawData, '[[Inner Selection]]', (SELECT
    @Budget))
IF @Version = 'Forecast'
        SET @SQL = REPLACE(@RawData, '[[Inner Selection]]', (SELECT
    @Forecast))
IF @Version = 'Actuals'
        SET @SQL = REPLACE(@RawData, '[[Inner Selection]]', (SELECT
    @Actuals))

/*Replace Variables with Values*/
SET @SQL = REPLACE(@SQL, '@MonthNumber', (SELECT
    @MonthNumber))
SET @SQL = REPLACE(@SQL, '@VersionIdLower', (SELECT
    @VersionIdLower))
SET @SQL = REPLACE(@SQL, '@VersionIdUpper', (SELECT
    @VersionIdUpper))
SET @SQL = REPLACE(@SQL, '@CurrentFiscalYearNumber', (SELECT
    @CurrentFiscalYearNumber))
SET @SQL = REPLACE(@SQL, '@FY', (SELECT
    '''' + @FY + ''''))
SET @SQL = REPLACE(@SQL, '@ForecastType', (SELECT
    @ForecastType))

/*Replace Header and Footer of RawData to Group By if neccesay*/
IF @FactTable = 'byDetail'
        SET @SQL = REPLACE(@SQL, '[[Footer]]', '')
IF @FactTable = 'byPC'
        SET @SQL = REPLACE(@SQL, 'SELECT * FROM', (SELECT
    @byPCSQL_Header))
IF @FactTable = 'byPC'
        SET @SQL = REPLACE(@SQL, '[[Footer]]', (SELECT
    @byPCSQL_Footer))
IF @FactTable = 'byProduct'
        SET @SQL = REPLACE(@SQL, 'SELECT * FROM', (SELECT
    @byProdSQL_Header))
IF @FactTable = 'byProduct'
        SET @SQL = REPLACE(@SQL, '[[Footer]]', (SELECT
    @byProdSQL_Footer))

/*This is for the Light Cube.  By Default this is included*/
IF (@MCBOnly = 0)
        BEGIN
    SET @SQL = REPLACE(@SQL, 'AND M.[ExecOrgSummaryId] IN (107)', '')
END

-- SELECT @SQL

EXECUTE [sys].sp_executesql @SQL