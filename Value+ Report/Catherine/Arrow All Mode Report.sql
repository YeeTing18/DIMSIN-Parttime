USE [dbARP]
GO
/****** Object:  StoredProcedure [dbo].[SP_GetRptData_SIN_Air_Sales_Incentive_Program]    Script Date: 6/1/2026 5:26:01 pm ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[SP_GetRptData_SIN_Air_Sales_Incentive_Program]
(
    @DBStation AS VARCHAR(10) = '',
	 @LotFrom AS VARCHAR(MAX)='',
	@LotTo AS VARCHAR(MAX)=''
)
AS
BEGIN
    SET NOCOUNT ON;

DECLARE @SQL AS NVARCHAR(MAX);
DECLARE @eChainVP AS NVARCHAR(MAX);
DECLARE @LastMonthDate AS DATE = DATEADD(M, -4, GETDATE());
DECLARE @YYMM_Input VARCHAR(4) = LEFT(@LotFrom, 4);

-- 2. Convert that YYMM to a proper Start Date
-- We prefix with '20' to get '2025-08-01'
DECLARE @ReportMonthStart AS DATE = CAST('20' + LEFT(@YYMM_Input, 2) + '-' + RIGHT(@YYMM_Input, 2) + '-01' AS DATE);
DECLARE @FiveYearsAgo AS INT = YEAR(@ReportMonthStart) - 5;
DECLARE @OneYearThreshold AS DATE = DATEADD(YEAR, -1, @ReportMonthStart); -- Standard 1-year lookback
SET @eChainVP = '';
EXEC SP_GetDBName @DBStation, @eChainVP OUTPUT;

-- Create and populate a temp table for the active salespeople.
IF OBJECT_ID('tempdb..#ActiveSalesPeople') IS NOT NULL DROP TABLE #ActiveSalesPeople;
CREATE TABLE #ActiveSalesPeople (
    SalesPersonName VARCHAR(max)
);

INSERT INTO #ActiveSalesPeople (SalesPersonName)
SELECT
    u.fullname
FROM
    [ReSM].[dbo].[SMUser] u

-- Step 1: Create a temp table to hold the detailed data for the last month.
IF OBJECT_ID('tempdb..#LastMonthData') IS NOT NULL DROP TABLE #LastMonthData;
CREATE TABLE #LastMonthData (
    StationName VARCHAR(max),HQID varchar(max), CustomerID varchar(max), CustomerName VARCHAR(max), Mode VARCHAR(max), ProductLine VARCHAR(max),
    MBL VARCHAR(max), HBL VARCHAR(max), Lot VARCHAR(max), PReceipt VARCHAR(max), POL VARCHAR(max), POD VARCHAR(max),
    GP DECIMAL(20,4), LotYear varchar(max), SalesPerson VARCHAR(max)
);

-- Step 2: Build the dynamic SQL to insert the detailed records from last month.
SET @SQL = N'
INSERT INTO #LastMonthData (StationName,HQID, CustomerID, CustomerName, Mode,ProductLine, MBL, HBL, Lot, PReceipt, POL, POD, GP, LotYear, SalesPerson)
SELECT 
    (SELECT s.StationCode FROM ReSM..SMStation s WHERE s.StationID = A.STATIONID) AS StationName,
	E.CustomerID AS HQDID,
    E.Customercode AS CustomerID,
    E.CustomerName AS CustomerName,
    ''AE'' As Mode,
	x.productline as ProductLine,
    A.MAWBNo AS MBL, 
    B.HAWBNo AS HBL,
    A.LOTNO AS Lot,
    NULL AS PReceipt,
    dbo.fnAECityCodeByID(TRY_CAST(B.portofdept AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(B.portofdstn AS INT)) AS POD,
    ISNULL(CAST(((C.FRT_Sales_PP + C.Extra_Sales_PP + C.Other_Sales_PP + C.FRT_Sales_CC + C.Extra_Sales_CC + C.Other_Sales_CC) - (C.FRT_ECOST + C.Other_ECOST + C.Extra_Other_ECost + C.Extra_FRT_ECost)) as decimal(10,2)),0) AS GP,
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.LotNo, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
    D.FULLNAME AS SalesPerson
FROM [' + @eChainVP + '].dbo.AEMAWB A 
LEFT JOIN [' + @eChainVP + '].dbo.aehawb B ON B.MAWBID = A.ID AND B.STATIONID = A.STATIONID 
left join [' + @eChainVP + '].dbo.fmopsource x on x.house = B.HAWBNo
LEFT JOIN [' + @eChainVP + '].dbo.AEAWBSUM C ON C.SourceID = b.ID AND C.StationID = A.STATIONID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne E on E.CustomerID = b.customer 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne F on F.CustomerID = b.shipper 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne G on G.CustomerID = b.CNEE 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUSER D ON D.USERID = TRY_CAST(b.sales AS varchar(max))
WHERE a.Status <> ''VOID'' and LotNo between ''' + @LotFrom + ''' AND ''' + @LotTo + ''' 

UNION ALL

SELECT 
    (SELECT s.StationCode FROM ReSM..SMStation s WHERE s.StationID = A.STATIONID) AS StationName,
	E.CustomerID AS HQID,
    e.customercode AS CustomerID,
    E.CustomerName AS CustomerName,
    ''AI'' As Mode,
	x.productline as ProductLine,
    A.MAWBNo AS MBL, 
    B.HAWBNo AS HBL,
    A.LOTNO AS Lot,
    NULL AS PReceipt,
    dbo.fnAECityCodeByID(TRY_CAST(B.portofdept AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(B.portofdstn AS INT)) AS POD,
    ISNULL(CAST(((C.FRT_Sales_PP + C.Extra_Sales_PP + C.Other_Sales_PP + C.FRT_Sales_CC + C.Extra_Sales_CC + C.Other_Sales_CC) - (C.FRT_ECOST + C.Other_ECOST + C.Extra_Other_ECost + C.Extra_FRT_ECost)) as decimal(10,2)),0) AS GP,
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.LotNo, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
    D.FULLNAME AS SalesPerson
FROM [' + @eChainVP + '].dbo.AiMAWB A 
LEFT JOIN [' + @eChainVP + '].dbo.aihawb B ON B.MAWBID = A.ID AND B.STATIONID = A.STATIONID 
left join [' + @eChainVP + '].dbo.fmopsource x on x.house = b.hawbno 
LEFT JOIN [' + @eChainVP + '].dbo.AiAWBSUM C ON C.SourceID = B.ID AND C.StationID = A.STATIONID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne E on E.CustomerID = b.customer 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne F on F.CustomerID = b.shipper 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne G on G.CustomerID = b.CNEE 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUSER D ON D.USERID = TRY_CAST(b.sales AS varchar(max))
WHERE a.Status <> ''VOID'' and LotNo between ''' + @LotFrom + ''' AND ''' + @LotTo + ''' 

UNION ALL

SELECT 
    (SELECT s.StationCode FROM ReSM..SMStation s WHERE s.StationID = A.STATIONID) AS StationName,
	E.CustomerID AS HQID,
    e.customercode AS CustomerID,
    E.CustomerName AS CustomerName,
    ''CC'' As Mode,
	x.productline as ProductLine,
    A.Reference1 AS MBL, 
    a.Reference2 AS HBL,
    A.LOTNO AS Lot,
    NULL AS PReceipt,
    dbo.fnAECityCodeByID(TRY_CAST(a.DEPT AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(a.dstn AS INT)) AS POD,
    ISNULL(CAST(((C.FRT_Sales_PP + C.Extra_Sales_PP + C.Other_Sales_PP + C.FRT_Sales_CC + C.Extra_Sales_CC + C.Other_Sales_CC) - (C.FRT_ECOST + C.Other_ECOST + C.Extra_Other_ECost + C.Extra_FRT_ECost)) as decimal(10,2)),0) AS GP,
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.LotNo, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
    D.FULLNAME AS SalesPerson
FROM [' + @eChainVP + '].dbo.CBBrokerage A 
LEFT JOIN [' + @eChainVP + '].dbo.CCAWBSUM C ON C.SourceID = A.ID AND C.StationID = A.STATIONID 
left join [' + @eChainVP + '].dbo.fmopsource x on x.house = a.Reference2
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne E on E.CustomerID = a.customer 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne F on F.CustomerID = a.shpr 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne G on G.CustomerID = a.CNEE 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUSER D ON D.USERID = TRY_CAST(a.salesperson AS varchar(max))
WHERE a.LotStatus <> ''V'' and LotNo between ''' + @LotFrom + ''' AND ''' + @LotTo + ''' 

UNION ALL

SELECT 
    (SELECT s.StationCode FROM ReSM..SMStation s WHERE s.StationID = A.STATIONID) AS StationName,
	E.CustomerID AS HQID,
    e.customercode AS CustomerID,
    E.CustomerName AS CustomerName,
    ''3A'' As Mode,
	x.productline as ProductLine,
    a.MasterNo AS MBL,
    A.HouseNo AS HBL,
    A.LOT AS Lot,
    NULL AS PReceipt, 
    dbo.fnAECityCodeByID(TRY_CAST(a.ORGN AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(a.DSTN AS INT)) AS POD, 
    ISNULL(CAST(((C.FRT_Sales_PP + C.Extra_Sales_PP + C.Other_Sales_PP + C.FRT_Sales_CC + C.Extra_Sales_CC + C.Other_Sales_CC) - (C.FRT_ECOST + C.Other_ECOST + C.Extra_Other_ECost + C.Extra_FRT_ECost)) as decimal(10,2)),0) AS GP,
    CASE
        WHEN A.Lot LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.Lot, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
    D.FULLNAME AS SalesPerson
FROM [' + @eChainVP + '].dbo.TPAIRData A 
left join [' + @eChainVP + '].dbo.fmopsource x on x.house = A.HouseNo
LEFT JOIN [' + @eChainVP + '].dbo.TPAAWBSUM C ON C.SourceID = A.ID AND C.StationID = A.STATIONID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne E on E.CustomerID = a.CustomerID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne F on F.CustomerID = a.Shipper 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne G on G.CustomerID = a.CNEE 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUSER D ON D.USERID = TRY_CAST(a.salesperson AS varchar(max))
WHERE a.LotStatus <> ''VOID'' and A.Lot between ''' + @LotFrom + ''' AND ''' + @LotTo + ''' 

UNION ALL

SELECT 
    (SELECT s.StationCode FROM ReSM..SMStation s WHERE s.StationID = A.STATIONID) AS StationName,
	E.CustomerID AS HQID,
    e.customercode AS CustomerID,
    E.CustomerName AS CustomerName,
    ''3O'' As Mode,
	x.Productline as ProductLine,
    a.MasterNo AS MBL,
    A.HouseNo AS HBL,
    A.LOT AS Lot,
    dbo.fnAECityCodeByID(TRY_CAST(a.Receipt AS INT)) AS PReceipt,
    dbo.fnAECityCodeByID(TRY_CAST(a.Receipt AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(a.Delivery AS INT)) AS POD,
    ISNULL(CAST(((C.FRT_Sales_PP + C.Extra_Sales_PP + C.Other_Sales_PP + C.FRT_Sales_CC + C.Extra_Sales_CC + C.Other_Sales_CC) - (C.FRT_ECOST + C.Other_ECOST + C.Extra_Other_ECost + C.Extra_FRT_ECost)) as decimal(10,2)),0) AS GP,
    CASE
        WHEN A.Lot LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.Lot, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
    D.FULLNAME AS SalesPerson
FROM [' + @eChainVP + '].dbo.TPOceanData A 
left join [' + @eChainVP + '].dbo.fmopsource x on x.house = A.HouseNo
LEFT JOIN [' + @eChainVP + '].dbo.TPOAWBSUM C ON C.SourceID = A.ID AND C.StationID = A.STATIONID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne E on E.CustomerID = a.CustomerID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne F on F.CustomerID = a.Shipper 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne G on G.CustomerID = a.CNEE 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUSER D ON D.USERID = TRY_CAST(a.salesperson AS varchar(max))
WHERE a.LotStatus <> ''VOID'' and A.Lot between ''' + @LotFrom + ''' AND ''' + @LotTo + ''' 

UNION ALL

SELECT 
    (SELECT s.StationCode FROM ReSM..SMStation s WHERE s.StationID = A.STATIONID) AS StationName,
	cus.CustomerID AS HQID,
    cus.customercode AS CustomerID,
    Cus.CustomerName AS CustomerName,
    ''TO'' As Mode,
	x.productline as ProductLine,
    b.MAWBNO AS MBL,
    b.HAWBNO AS HBL,
    a1.LotNo as Lot,
    NULL AS PReceipt,
    d.LocNo AS POL,
    e.LocNo AS POD, 
    TotalSales_Home-(TotalCost_Home+TruckCost_Home+MainDriverCost_Home+BackupDriverCost_Home) as GP,
    CASE
        WHEN a1.LotNo LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(a1.LotNo, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
    '''' AS SalesPerson
FROM [' + @eChainVP + '].dbo.TKStockDetailSummary a 
INNER JOIN [' + @eChainVP + '].dbo.TKOrderMain b on a.OrderID=b.ID 
left join [' + @eChainVP + '].dbo.fmopsource x on x.house = b.HAWBNO
INNER JOIN [' + @eChainVP + '].dbo.TKStockMain a1 on a.StockID=a1.StockID AND a.truckid = isnull(a1.truckid,a1.vendor_truck_id) 
INNER JOIN [' + @eChainVP + '].dbo.TKTruck c on a.TruckID=c.ID 
INNER JOIN [' + @eChainVP + '].dbo.TKdicDEPTDSTN d on b.DeptID=d.ID 
INNER JOIN [' + @eChainVP + '].dbo.TKdicDEPTDSTN e on b.DSTNID=e.ID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne f on b.SendCargoCompanyID=f.CustomerID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne g on b.GetCargoCompanyID=g.CustomerID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne h on b.DeliveryCargoCompanyID=h.CustomerID 
LEFT JOIN [' + @eChainVP + '].dbo.SMCustomerOne Cus on Cus.CustomerID = B.CustomerID 
WHERE a.Status <> ''VOID'' and a1.LotNo between ''' + @LotFrom + ''' AND ''' + @LotTo + ''' 
';
-- Pass in the Lot range parameters
EXEC sp_executesql @SQL, N'@LotFrom VARCHAR(8), @LotTo VARCHAR(8)', @LotFrom = @LotFrom, @LotTo = @LotTo;


-- Step 3: Create history tables
IF OBJECT_ID('tempdb..#CustomerHistory') IS NOT NULL DROP TABLE #CustomerHistory;
CREATE TABLE #CustomerHistory (CustomerID varchar(max), MinYear INT, MaxShipmentDate DATE);

IF OBJECT_ID('tempdb..#LaneHistory') IS NOT NULL DROP TABLE #LaneHistory;
CREATE TABLE #LaneHistory (CustomerID varchar(max), POL VARCHAR(100), POD VARCHAR(100), MinYear INT, MaxShipmentDate DATE);

IF OBJECT_ID('tempdb..#LaneOwnershipHistory') IS NOT NULL DROP TABLE #LaneOwnershipHistory;
CREATE TABLE #LaneOwnershipHistory (
    CustomerID varchar(max), POL VARCHAR(100), POD VARCHAR(100),
    DistinctSalespersonCount INT, SingleOwnerName VARCHAR(100), SingleOwnerMinYear INT
);

IF OBJECT_ID('tempdb..#LaneHistory_PReceipt') IS NOT NULL DROP TABLE #LaneHistory_PReceipt;
CREATE TABLE #LaneHistory_PReceipt (CustomerID varchar(max), PReceipt VARCHAR(100), POD VARCHAR(100), MinYear INT, MaxShipmentDate DATE);

IF OBJECT_ID('tempdb..#LaneOwnershipHistory_PReceipt') IS NOT NULL DROP TABLE #LaneOwnershipHistory_PReceipt;
CREATE TABLE #LaneOwnershipHistory_PReceipt (
    CustomerID varchar(max), PReceipt VARCHAR(100), POD VARCHAR(100),
    DistinctSalespersonCount INT, SingleOwnerName VARCHAR(100), SingleOwnerMinYear INT
);

IF OBJECT_ID('tempdb..#AllHistoricalData') IS NOT NULL DROP TABLE #AllHistoricalData;
CREATE TABLE #AllHistoricalData (
    CustomerID varchar(max), PReceipt VARCHAR(100), POL VARCHAR(100), POD VARCHAR(100), SalesPerson VARCHAR(100), LotYear INT, ShipmentDate DATE
);

-- Step 4: Populate history tables with data from previous years
SET @SQL = N'
INSERT INTO #AllHistoricalData (CustomerID, PReceipt, POL, POD, SalesPerson, LotYear, ShipmentDate)
SELECT 
    B.customer AS CustomerID,
    NULL AS PReceipt,
    dbo.fnAECityCodeByID(TRY_CAST(B.portofdept AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(B.portofdstn AS INT)) AS POD,
    D.FULLNAME AS SalesPerson,
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.LotNo, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9][0-9][0-9]%''
        THEN TRY_CAST(''20'' + SUBSTRING(A.LotNo, 1, 2) + ''-'' + SUBSTRING(A.LotNo, 3, 2) + ''-01'' AS DATE)
        ELSE NULL
    END AS ShipmentDate
FROM [' + @eChainVP + '].dbo.AEMAWB A 
LEFT JOIN [' + @eChainVP + '].dbo.aehawb B ON B.MAWBID = A.ID AND B.STATIONID = A.STATIONID 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUSER D ON D.USERID = TRY_CAST(b.sales AS varchar(max))
WHERE a.Status <> ''VOID'' AND 
    CASE WHEN A.LotNo LIKE ''[0-9][0-9][0-9][0-9]%'' THEN TRY_CAST(''20'' + SUBSTRING(A.LotNo, 1, 2) + ''-'' + SUBSTRING(A.LotNo, 3, 2) + ''-01'' AS DATE) ELSE NULL END < @ReportMonthStart

UNION ALL

SELECT 
    B.customer AS CustomerID,
    NULL AS PReceipt,
    dbo.fnAECityCodeByID(TRY_CAST(B.portofdept AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(B.portofdstn AS INT)) AS POD,
    D.FULLNAME AS SalesPerson,
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.LotNo, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9][0-9][0-9]%''
        THEN TRY_CAST(''20'' + SUBSTRING(A.LotNo, 1, 2) + ''-'' + SUBSTRING(A.LotNo, 3, 2) + ''-01'' AS DATE)
        ELSE NULL
    END AS ShipmentDate
FROM [' + @eChainVP + '].dbo.AiMAWB A 
LEFT JOIN [' + @eChainVP + '].dbo.aihawb B ON B.MAWBID = A.ID AND B.STATIONID = A.STATIONID 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUSER D ON D.USERID = TRY_CAST(b.sales AS varchar(max))
WHERE a.Status <> ''VOID'' AND 
    CASE WHEN A.LotNo LIKE ''[0-9][0-9][0-9][0-9]%'' THEN TRY_CAST(''20'' + SUBSTRING(A.LotNo, 1, 2) + ''-'' + SUBSTRING(A.LotNo, 3, 2) + ''-01'' AS DATE) ELSE NULL END < @ReportMonthStart

UNION ALL

SELECT 
    A.customer AS CustomerID,
    NULL AS PReceipt,
    dbo.fnAECityCodeByID(TRY_CAST(a.DEPT AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(a.dstn AS INT)) AS POD,
    I.FullName AS SalesPerson,
 
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.LotNo, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
 
    CASE
        WHEN A.LotNo LIKE ''[0-9][0-9][0-9][0-9]%''
        THEN TRY_CAST(''20'' + SUBSTRING(A.LotNo, 1, 2) + ''-'' + SUBSTRING(A.LotNo, 3, 2) + ''-01'' AS DATE)
        ELSE NULL
    END AS ShipmentDate
FROM [' + @eChainVP + '].dbo.CBBrokerage A 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUser I ON I.UserID = TRY_CAST(a.salesperson AS varchar(max))
WHERE a.LotStatus <> ''V'' AND 
    CASE WHEN A.LotNo LIKE ''[0-9][0-9][0-9][0-9]%'' THEN TRY_CAST(''20'' + SUBSTRING(A.LotNo, 1, 2) + ''-'' + SUBSTRING(A.LotNo, 3, 2) + ''-01'' AS DATE) ELSE NULL END < @ReportMonthStart

UNION ALL

SELECT 
    A.CustomerID AS CustomerID,
    NULL AS PReceipt, 
    dbo.fnAECityCodeByID(TRY_CAST(a.ORGN AS INT)) AS POL, 
    dbo.fnAECityCodeByID(TRY_CAST(a.DSTN AS INT)) AS POD, 
    I.FullName AS SalesPerson,
 
    CASE
        WHEN A.Lot LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.Lot, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
 
    CASE
        WHEN A.Lot LIKE ''[0-9][0-9][0-9][0-9]%''
        THEN TRY_CAST(''20'' + SUBSTRING(A.Lot, 1, 2) + ''-'' + SUBSTRING(A.Lot, 3, 2) + ''-01'' AS DATE)
        ELSE NULL
    END AS ShipmentDate
FROM [' + @eChainVP + '].dbo.TPAIRData A 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUser I ON I.UserID = TRY_CAST(a.salesperson AS varchar(max))
WHERE a.LotStatus <> ''VOID'' AND 
    CASE WHEN A.Lot LIKE ''[0-9][0-9][0-9][0-9]%'' THEN TRY_CAST(''20'' + SUBSTRING(A.Lot, 1, 2) + ''-'' + SUBSTRING(A.Lot, 3, 2) + ''-01'' AS DATE) ELSE NULL END < @ReportMonthStart

UNION ALL

SELECT 
    A.CustomerID AS CustomerID,
    dbo.fnAECityCodeByID(TRY_CAST(a.Receipt AS INT)) AS PReceipt,
    dbo.fnAECityCodeByID(TRY_CAST(a.Receipt AS INT)) AS POL,
    dbo.fnAECityCodeByID(TRY_CAST(a.Delivery AS INT)) AS POD,
    I.FullName AS SalesPerson,
    
    CASE
        WHEN A.Lot LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(A.Lot, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,
 
    CASE
        WHEN A.Lot LIKE ''[0-9][0-9][0-9][0-9]%''
        THEN TRY_CAST(''20'' + SUBSTRING(A.Lot, 1, 2) + ''-'' + SUBSTRING(A.Lot, 3, 2) + ''-01'' AS DATE)
        ELSE NULL
    END AS ShipmentDate
FROM [' + @eChainVP + '].dbo.TPOceanData A 
-- **FIX APPLIED HERE**
LEFT JOIN RESM..SMUser I ON I.UserID = TRY_CAST(a.salesperson AS varchar(max))
WHERE a.LotStatus <> ''VOID'' AND 
    CASE WHEN A.Lot LIKE ''[0-9][0-9][0-9][0-9]%'' THEN TRY_CAST(''20'' + SUBSTRING(A.Lot, 1, 2) + ''-'' + SUBSTRING(A.Lot, 3, 2) + ''-01'' AS DATE) ELSE NULL END < @ReportMonthStart

UNION ALL

SELECT 
    B.CustomerID AS CustomerID,
    NULL AS PReceipt,
    d.LocNo AS POL,
    e.LocNo AS POD, 
    '''' AS SalesPerson,
 
    CASE
        WHEN a1.LotNo LIKE ''[0-9][0-9]%''
        THEN TRY_CAST(LEFT(a1.LotNo, 2) AS INT) + 2000
        ELSE NULL
    END AS LotYear,

    CASE
        WHEN a1.LotNo LIKE ''[0-9][0-9][0-9][0-9]%''
        THEN TRY_CAST(''20'' + SUBSTRING(a1.LotNo, 1, 2) + ''-'' + SUBSTRING(a1.LotNo, 3, 2) + ''-01'' AS DATE)
        ELSE NULL
    END AS ShipmentDate
FROM [' + @eChainVP + '].dbo.TKStockDetailSummary a 
INNER JOIN [' + @eChainVP + '].dbo.TKOrderMain b on a.OrderID=b.ID 
INNER JOIN [' + @eChainVP + '].dbo.TKStockMain a1 on a.StockID=a1.StockID AND a.truckid = isnull(a1.truckid,a1.vendor_truck_id) 
INNER JOIN [' + @eChainVP + '].dbo.TKdicDEPTDSTN d on b.DeptID=d.ID 
INNER JOIN [' + @eChainVP + '].dbo.TKdicDEPTDSTN e on b.DSTNID=e.ID 
WHERE a.Status <> ''VOID'' AND 
    CASE WHEN a1.LotNo LIKE ''[0-9][0-9][0-9][0-9]%'' THEN TRY_CAST(''20'' + SUBSTRING(a1.LotNo, 1, 2) + ''-'' + SUBSTRING(a1.LotNo, 3, 2) + ''-01'' AS DATE) ELSE NULL END < @ReportMonthStart;
';

EXEC sp_executesql @SQL, N'@ReportMonthStart DATE', @ReportMonthStart;

-- Step 5: Process and Select the final report data (No changes required here)

-- Customer History (General)
INSERT INTO #CustomerHistory (CustomerID, MinYear, MaxShipmentDate)
SELECT CustomerID, MIN(LotYear), MAX(ShipmentDate) FROM #AllHistoricalData WHERE CustomerID IS NOT NULL AND LotYear IS NOT NULL GROUP BY CustomerID;

-- Lane History (POL -> POD)
INSERT INTO #LaneHistory (CustomerID, POL, POD, MinYear, MaxShipmentDate)
SELECT CustomerID, POL, POD, MIN(LotYear), MAX(ShipmentDate) FROM #AllHistoricalData WHERE CustomerID IS NOT NULL AND POL IS NOT NULL AND POD IS NOT NULL AND LotYear IS NOT NULL GROUP BY CustomerID, POL, POD;

-- Lane Ownership History (POL -> POD)
INSERT INTO #LaneOwnershipHistory (CustomerID, POL, POD, DistinctSalespersonCount, SingleOwnerName, SingleOwnerMinYear)
SELECT CustomerID, POL, POD, COUNT(DISTINCT SalesPerson), CASE WHEN COUNT(DISTINCT SalesPerson) = 1 THEN MAX(SalesPerson) ELSE NULL END, CASE WHEN COUNT(DISTINCT SalesPerson) = 1 THEN MIN(LotYear) ELSE NULL END
FROM #AllHistoricalData WHERE CustomerID IS NOT NULL AND POL IS NOT NULL AND POD IS NOT NULL AND SalesPerson IS NOT NULL AND SalesPerson <> '' GROUP BY CustomerID, POL, POD;

-- Lane History (PReceipt -> POD)
INSERT INTO #LaneHistory_PReceipt (CustomerID, PReceipt, POD, MinYear, MaxShipmentDate)
SELECT CustomerID, PReceipt, POD, MIN(LotYear), MAX(ShipmentDate) FROM #AllHistoricalData WHERE CustomerID IS NOT NULL AND PReceipt IS NOT NULL AND POD IS NOT NULL AND LotYear IS NOT NULL GROUP BY CustomerID, PReceipt, POD;

-- Lane Ownership History (PReceipt -> POD)
INSERT INTO #LaneOwnershipHistory_PReceipt (CustomerID, PReceipt, POD, DistinctSalespersonCount, SingleOwnerName, SingleOwnerMinYear)
SELECT CustomerID, PReceipt, POD, COUNT(DISTINCT SalesPerson), CASE WHEN COUNT(DISTINCT SalesPerson) = 1 THEN MAX(SalesPerson) ELSE NULL END, CASE WHEN COUNT(DISTINCT SalesPerson) = 1 THEN MIN(LotYear) ELSE NULL END
FROM #AllHistoricalData WHERE CustomerID IS NOT NULL AND PReceipt IS NOT NULL AND POD IS NOT NULL AND SalesPerson IS NOT NULL AND SalesPerson <> '' GROUP BY CustomerID, PReceipt, POD;
 
-- Step 6: Select and display the final report
SELECT
    lmd.StationName,
    lmd.SalesPerson,
    '' as Type,
	lmd.ProductLine,
	lmd.HQID,
    lmd.CustomerID,
    lmd.CustomerName,
    lmd.Mode,
    lmd.MBL,
    lmd.HBL,
    lmd.Lot,
    lmd.POL,
    lmd.POD,
    lmd.GP,
    ' ' AS Ratio,
    ' ' AS [GP for Incentive]
FROM #LastMonthData lmd
LEFT JOIN #CustomerHistory ch ON lmd.CustomerID = ch.CustomerID
LEFT JOIN #LaneHistory lh ON lmd.CustomerID = lh.CustomerID AND lmd.POL = lh.POL AND lmd.POD = lh.POD
LEFT JOIN #LaneOwnershipHistory loh ON lmd.CustomerID = loh.CustomerID AND lmd.POL = loh.POL AND lmd.POD = loh.POD
LEFT JOIN #LaneHistory_PReceipt lhp ON lmd.CustomerID = lhp.CustomerID AND lmd.PReceipt = lhp.PReceipt AND lmd.POD = lhp.POD
LEFT JOIN #LaneOwnershipHistory_PReceipt lohp ON lmd.CustomerID = lohp.CustomerID AND lmd.PReceipt = lohp.PReceipt AND lmd.POD = lohp.POD
CROSS APPLY (
    SELECT
        Type = CASE
            -- Customer Status checks
            WHEN ch.CustomerID IS NULL THEN 'FreeHand (New Customer)'
            -- Lane Status checks (considers both POL->POD and PReceipt->POD)
            WHEN lh.CustomerID IS NULL AND lhp.CustomerID IS NULL THEN 'FreeHand (New Lane)'
            WHEN ch.MaxShipmentDate < @OneYearThreshold THEN 'FreeHand (Reactivated Customer)'
            WHEN (CASE WHEN lh.MaxShipmentDate > lhp.MaxShipmentDate THEN lh.MaxShipmentDate ELSE ISNULL(lhp.MaxShipmentDate, lh.MaxShipmentDate) END) < @OneYearThreshold THEN 'FreeHand (Reactivated Lane)'
             
            -- House Account logic (precedence to POL->POD lane)
            WHEN lh.CustomerID IS NOT NULL AND (loh.DistinctSalespersonCount > 1 OR lmd.SalesPerson <> loh.SingleOwnerName) THEN 'House Account (New Sales Person)'
            WHEN lh.CustomerID IS NOT NULL AND loh.SingleOwnerMinYear < @FiveYearsAgo THEN 'House Account (>= 5 Years)'
            WHEN lh.CustomerID IS NOT NULL THEN 'House Account (< 5 Years)'
             
            -- House Account logic (fallback to PReceipt->POD lane)
            WHEN lhp.CustomerID IS NOT NULL AND (lohp.DistinctSalespersonCount > 1 OR lmd.SalesPerson <> lohp.SingleOwnerName) THEN 'House Account (New Sales Person)'
            WHEN lhp.CustomerID IS NOT NULL AND lohp.SingleOwnerMinYear < @FiveYearsAgo THEN 'House Account (>= 5 Years)'
            ELSE 'House Account (< 5 Years)'
        END
) AS TypeCalc
CROSS APPLY (
    SELECT
        Ratio = CASE TypeCalc.Type
            WHEN 'FreeHand (New Customer)' THEN '100%'
            WHEN 'FreeHand (New Lane)' THEN '100%'
            WHEN 'FreeHand (Reactivated Customer)' THEN '100%'
            WHEN 'FreeHand (Reactivated Lane)' THEN '100%'
            WHEN 'House Account (New Sales Person)' THEN '30%'
            WHEN 'House Account (>= 5 Years)' THEN '80%'
            WHEN 'House Account (< 5 Years)' THEN '100%'
            ELSE '0%'
        END
) AS Calcs

-- Step 7: Clean up temporary tables
DROP TABLE #LastMonthData;
DROP TABLE #CustomerHistory;
DROP TABLE #LaneHistory;
DROP TABLE #LaneOwnershipHistory;
DROP TABLE #AllHistoricalData;
DROP TABLE #LaneHistory_PReceipt;
DROP TABLE #LaneOwnershipHistory_PReceipt;
DROP TABLE #ActiveSalesPeople;

end
