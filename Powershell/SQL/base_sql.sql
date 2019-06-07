
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
