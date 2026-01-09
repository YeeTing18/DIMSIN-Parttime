SELECT
    'DMER' AS CarrierID,
    'AEASGD' AS ClientID,
    'D' AS InvoiceType,
    inv.InvoiceNo AS [Invoice Number],
    inv.InvoiceAmount AS [Invoice Total],
    inv.InvoiceDate AS [Invoice Date],
    aeH.HAWBNo AS [Shipment Number],
	CONVERT(DATETIME,(SELECT TOP 1 CONVERT(VARCHAR(11),MilestoneTime,106) FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeh.HAWBNo AND PRODUCTLINE = 2 AND MileStoneID = 500 AND Status = 1 ),101) AS 'Ship Date',
    'AEAS' as [Bill To Account],
	'K' as [Weight Unit],
	aeH.CWT AS [Billed Weight],
	aeH.CWT AS [Actual Weight],
	'' as [DIM Weight],
	'' as [Volume],
	aeh.ActPCS as [PCS],
	'' AS [Rejected Pieces],
	AEH.ActPCSUOM AS [Package Type],
	'' AS [Loading meters],
	'LTL' AS [Service Type],
	'' as [Service Zone],
	'I' as [Direction],
	'Air' as [Mode],
	'' as [Payment Terms],
	'' as [Inco terms],
	'' as [Distance qualifier],
	'' as [Distance],
	'' as [Bill Of Lading],
	'' as [MAWB],
	'' as [HAWB],
	'' as [PO Number],
	'Air' as [Reference number],
    ISNULL(
        (SELECT TOP 1 MilestoneTime FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeH.HAWBNo AND PRODUCTLINE = 2 AND MileStoneID = 2970 AND Status = 1),
        (SELECT TOP 1 MilestoneTime FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeH.HAWBNo AND PRODUCTLINE = 2 AND MileStoneID = 3000 AND Status = 1)
    ) AS [Delivery Date],
	'' as [Delivery Time],
	'' as [POD Name],
	CASE WHEN aeH.ActPCSUOM = 'PLT' THEN aeH.ActPCS ELSE NULL END AS [Notes1 (Pallets)],
    CASE WHEN aeH.ActPCSUOM = 'CTN' THEN aeH.ActPCS ELSE NULL END AS [Notes2 (Cartons)],
	aeH.hawbno as [Notes3],
	'' as [Notes4],
	'' as [Notes5],
	'' as [Notes6],
	'' as[Notes7],
	'' as [Notes8],
	'' as [Notes9],
	'' as [Notes10],
	'' as [Notes11],
	'' as [Notes12],
	'' as [Notes13],
	'' as [Notes14],
	'' as [Notes15],
	'' as [Notes16],
	 OriginAirport.AirportCode as [Origin(Air)Port], 
	 DestAirport.AirportCode AS [Destination (Air)Port],
	 '' as [Shipper Location],
	 shp.customername as [Shipper Name],
	 shp.customername as [Shipper Company],
	 shp.customeraddress1 + shp.CustomerAddress2 as [Shipper Address1],
	 '' as [Shipper Address2],
	 '' as [Shipper Address3],
	 city.CityName as [Shipper City],
	 '' as [Shipper state],
	 '' as [Shipper Postcode],
	 city.CityName as [Shipper Country],
	 '' as [Consignee Location],
	  cnee.customername as [Consignee Name],
	  cnee.customername as [Consignee Company],
	  cnee.CustomerAddress1 +cnee.CustomerAddress2 as [Consignee Address1],
	  '' as [Consignee Address2],
	  '' as [Consignee Address3],
	  dstn.CityName as [Consignee City],
	  '' as [Consignee State],
	  cnee.CustomerAddress1 +cnee.CustomerAddress2 as [Consignee Postcode],
	  country.countryname as [Consignee Country],
	  inv.InvoiceAmount as [Amount Billed],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 13 ) as [Freight Amount - FRT],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 13 ) as [Total Shipment VAT Amount],
	  'FSC' AS [Accessorial Charge- FSC],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 14 ) as [Accessorial Charge-FSC Amount1 - Fuel Surcharge],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 14 ) as [VAT liable flag 1],
	  'SSC' as [Accessorial Charge-SSC],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 15 ) as [Accessorial Charge-SSC Amount2],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 15 ) as [VAT liable flag 2],
	  '' as [Accessorial Charge-P/U Code3],
	  '' as [Accessorial Charge- Amount3],
	  '' as [VAT liable flag 3],
	  'H/C' AS [Accessorial Charge-Origin Terminal Code4],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 400 ) as [Accessorial Charge-Origin Terminal Amount4],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 400 ) as [VAT liable flag 4],
	  '' AS [Accessorial Charge-CUS Code5],
	  '' AS [Accessorial Charge-CUS Amount5],
	  '' AS [VAT liable flag 5],
	  'H/C' as [Accessorial Charge-Dest Terminal Code6],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 005 ) as [Accessorial Charge-dest Terminal Amount6],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 005 ) as [VAT liable flag 6],
	  'DEL' AS [Accessorial Charge Code7],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 113 ) as [Accessorial Charge Amount7],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 113 ) as [VAT liable flag 7],
	  'CUS' AS [Accessorial Charge Code8],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 001 ) as [Accessorial Charge Amount8],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 001 ) as [VAT liable flag 8],
	  '' as [Accessorial Charge Code9],
	  '' as [Accessorial Charge Amount9],
	  '' as [VAT liable flag 9],
	  '' as [Accessorial Charge Code10],
	  '' as [Accessorial Charge Amount10],
	  '' as[Accessorial 10 VAT code],
	  'SGD' AS [Currency],
	  '' AS [Bill to Location],
	  'Arrow Electronics Asia (S) Pte Ltd' AS [Bill To Company],
	  'NIC, 5 Tai Seng Drive #06-01' AS [Bill To Address1],
	  '' AS [Bill To Address2],
	  'Singapore' AS [Bill To City],
	  '' AS [Bill To State],
	  '535217' AS [Bill To Postcode],
	  'SG' AS [Bill To Country],
	  '' AS [Equipment Type],
	   '' AS [Container number],
	   '' AS [Trailer/Car ID],
	   '' AS [carrier VAT number],
	   '' AS [CLIENT VAT NUMBER],
	   '' AS [VAT percentage],
	   (SELECT TOP 1 STransactionRate FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID) AS [Currency Exchange Rate],
	   '' AS [Overtime hours],
	   '' AS [Waiting hours],
	   '' AS [Urgent fee location count]
	  
FROM
    FMOPSource AS inv 
LEFT JOIN
    AEHAWB AS aeH ON inv.SourceID = aeH.ID AND inv.StationID = aeH.StationID
LEFT JOIN
    dbo.smcustomerone AS cus ON aeH.Customer = cus.CustomerID
LEFT JOIN 
	RESM..SMAirPort as DestAirport ON DestAirport.HQID = aeH.airportofdstn
LEFT JOIN 
	RESM..SMAirPort as OriginAirport ON OriginAirport.HQID = aeH.airportofdept
LEFT JOIN
	dbo.smcustomerone AS shp ON aeH.Shipper = shp.CustomerID
Left Join
	RESM..SMCITY as city on city.HQID = aeh.PlaceOfRCPT
LEFT JOIN
    dbo.smcustomerone AS cnee ON aeH.CNEE = cnee.CustomerID
Left Join
	RESM..SMCITY as dstn on dstn.HQID = aeh.PortOfDSTN
Left Join
	RESM..smcountry as country on country.HQID = cnee.Country
Left join
	dbo.smcustomerone as billto on inv.billto = billto.CustomerID
WHERE
    inv.ModeCode = 'AE'        
    AND cus.CustomerCode = 'ARROWE001'
	and billto.CustomerCode='ARROWSG'
    AND ISNULL(inv.InvoiceNo, '') <> '' 
UNION ALL
SELECT
    'DMER' AS CarrierID,
    'AEASGD' AS ClientID,
    'D' AS InvoiceType,
    inv.InvoiceNo AS [Invoice Number],
    inv.InvoiceAmount AS [Invoice Total],
    inv.InvoiceDate AS [Invoice Date],
    aeH.HAWBNo AS [Shipment Number],
	CONVERT(DATETIME,(SELECT TOP 1 CONVERT(VARCHAR(11),MilestoneTime,106) FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeh.HAWBNo AND PRODUCTLINE = 2 AND MileStoneID = 500 AND Status = 1 ),101) AS 'Ship Date',
    'AEAS' as [Bill To Account],
	'K' as [Weight Unit],
	aeH.CWT AS [Billed Weight],
	aeH.CWT AS [Actual Weight],
	'' as [DIM Weight],
	'' as [Volume],
	aeh.ActPCS as [PCS],
	'' AS [Rejected Pieces],
	AEH.ActPCSUOM AS [Package Type],
	'' AS [Loading meters],
	'LTL' AS [Service Type],
	'' as [Service Zone],
	'I' as [Direction],
	'Air' as [Mode],
	'' as [Payment Terms],
	'' as [Inco terms],
	'' as [Distance qualifier],
	'' as [Distance],
	'' as [Bill Of Lading],
	'' as [MAWB],
	'' as [HAWB],
	'' as [PO Number],
	'Air' as [Reference number],
    ISNULL(
        (SELECT TOP 1 MilestoneTime FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeH.HAWBNo AND PRODUCTLINE = 2 AND MileStoneID = 2970 AND Status = 1),
        (SELECT TOP 1 MilestoneTime FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeH.HAWBNo AND PRODUCTLINE = 2 AND MileStoneID = 3000 AND Status = 1)
    ) AS [Delivery Date],
	'' as [Delivery Time],
	'' as [POD Name],
	CASE WHEN aeH.ActPCSUOM = 'PLT' THEN aeH.ActPCS ELSE NULL END AS [Notes1 (Pallets)],
    CASE WHEN aeH.ActPCSUOM = 'CTN' THEN aeH.ActPCS ELSE NULL END AS [Notes2 (Cartons)],
	aeH.hawbno as [Notes3],
	'' as [Notes4],
	'' as [Notes5],
	'' as [Notes6],
	'' as[Notes7],
	'' as [Notes8],
	'' as [Notes9],
	'' as [Notes10],
	'' as [Notes11],
	'' as [Notes12],
	'' as [Notes13],
	'' as [Notes14],
	'' as [Notes15],
	'' as [Notes16],
	 OriginAirport.AirportCode as [Origin(Air)Port], 
	 DestAirport.AirportCode AS [Destination (Air)Port],
	 '' as [Shipper Location],
	 shp.customername as [Shipper Name],
	 shp.customername as [Shipper Company],
	 shp.customeraddress1 + shp.CustomerAddress2 as [Shipper Address1],
	 '' as [Shipper Address2],
	 '' as [Shipper Address3],
	 city.CityName as [Shipper City],
	 '' as [Shipper state],
	 '' as [Shipper Postcode],
	 city.CityName as [Shipper Country],
	 '' as [Consignee Location],
	  cnee.customername as [Consignee Name],
	  cnee.customername as [Consignee Company],
	  cnee.CustomerAddress1 +cnee.CustomerAddress2 as [Consignee Address1],
	  '' as [Consignee Address2],
	  '' as [Consignee Address3],
	  dstn.CityName as [Consignee City],
	  '' as [Consignee State],
	  cnee.CustomerAddress1 +cnee.CustomerAddress2 as [Consignee Postcode],
	  country.countryname as [Consignee Country],
	  inv.InvoiceAmount as [Amount Billed],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 13 ) as [Freight Amount - FRT],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 13 ) as [Total Shipment VAT Amount],
	  'FSC' AS [Accessorial Charge- FSC],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 14 ) as [Accessorial Charge-FSC Amount1 - Fuel Surcharge],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 14 ) as [VAT liable flag 1],
	  'SSC' as [Accessorial Charge-SSC],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 15 ) as [Accessorial Charge-SSC Amount2],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 15 ) as [VAT liable flag 2],
	  '' as [Accessorial Charge-P/U Code3],
	  '' as [Accessorial Charge- Amount3],
	  '' as [VAT liable flag 3],
	  'H/C' AS [Accessorial Charge-Origin Terminal Code4],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 400 ) as [Accessorial Charge-Origin Terminal Amount4],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 400 ) as [VAT liable flag 4],
	  '' AS [Accessorial Charge-CUS Code5],
	  '' AS [Accessorial Charge-CUS Amount5],
	  '' AS [VAT liable flag 5],
	  'H/C' as [Accessorial Charge-Dest Terminal Code6],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 005 ) as [Accessorial Charge-dest Terminal Amount6],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 005 ) as [VAT liable flag 6],
	  'DEL' AS [Accessorial Charge Code7],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 113 ) as [Accessorial Charge Amount7],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 113 ) as [VAT liable flag 7],
	  'CUS' AS [Accessorial Charge Code8],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 001 ) as [Accessorial Charge Amount8],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 001 ) as [VAT liable flag 8],
	  '' as [Accessorial Charge Code9],
	  '' as [Accessorial Charge Amount9],
	  '' as [VAT liable flag 9],
	  '' as [Accessorial Charge Code10],
	  '' as [Accessorial Charge Amount10],
	  '' as[Accessorial 10 VAT code],
	  'SGD' AS [Currency],
	  '' AS [Bill to Location],
	  'Arrow Electronics Asia (S) Pte Ltd' AS [Bill To Company],
	  'NIC, 5 Tai Seng Drive #06-01' AS [Bill To Address1],
	  '' AS [Bill To Address2],
	  'Singapore' AS [Bill To City],
	  '' AS [Bill To State],
	  '535217' AS [Bill To Postcode],
	  'SG' AS [Bill To Country],
	  '' AS [Equipment Type],
	   '' AS [Container number],
	   '' AS [Trailer/Car ID],
	   '' AS [carrier VAT number],
	   '' AS [CLIENT VAT NUMBER],
	   '' AS [VAT percentage],
	   (SELECT TOP 1 STransactionRate FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID) AS [Currency Exchange Rate],
	   '' AS [Overtime hours],
	   '' AS [Waiting hours],
	   '' AS [Urgent fee location count]
	  
FROM
    FMOPSource AS inv 
LEFT JOIN
    AIHAWB AS aeH ON inv.SourceID = aeH.ID AND inv.StationID = aeH.StationID
LEFT JOIN
    dbo.smcustomerone AS cus ON aeH.Customer = cus.CustomerID
LEFT JOIN 
	RESM..SMAirPort as DestAirport ON DestAirport.HQID = aeH.airportofdstn
LEFT JOIN 
	RESM..SMAirPort as OriginAirport ON OriginAirport.HQID = aeH.airportofdept
LEFT JOIN
	dbo.smcustomerone AS shp ON aeH.Shipper = shp.CustomerID
Left Join
	RESM..SMCITY as city on city.HQID = aeh.PlaceOfRCPT
LEFT JOIN
    dbo.smcustomerone AS cnee ON aeH.CNEE = cnee.CustomerID
Left Join
	RESM..SMCITY as dstn on dstn.HQID = aeh.PortOfDSTN
Left Join
	RESM..smcountry as country on country.HQID = cnee.Country
Left join
	dbo.smcustomerone as billto on inv.billto = billto.CustomerID
WHERE
    inv.ModeCode = 'AI'        
    AND cus.CustomerCode = 'ARROWE001'
	and billto.CustomerCode='ARROWSG'
    AND ISNULL(inv.InvoiceNo, '') <> '' 
UNION ALL
SELECT
    'DMER' AS CarrierID,
    'AEASGD' AS ClientID,
    'D' AS InvoiceType,
    inv.InvoiceNo AS [Invoice Number],
    inv.InvoiceAmount AS [Invoice Total],
    inv.InvoiceDate AS [Invoice Date],
    aeH.HouseNo AS [Shipment Number],
	CONVERT(DATETIME,(SELECT TOP 1 CONVERT(VARCHAR(11),MilestoneTime,106) FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeh.HouseNo AND PRODUCTLINE = 2 AND MileStoneID = 500 AND Status = 1 ),101) AS 'Ship Date',
    'AEAS' as [Bill To Account],
	'K' as [Weight Unit],
	aeH.CWT AS [Billed Weight],
	aeH.CWT AS [Actual Weight],
	'' as [DIM Weight],
	'' as [Volume],
	aeh.ActPCS as [PCS],
	'' AS [Rejected Pieces],
	AEH.ActPCSUOM AS [Package Type],
	'' AS [Loading meters],
	'LTL' AS [Service Type],
	'' as [Service Zone],
	'I' as [Direction],
	'Air' as [Mode],
	'' as [Payment Terms],
	'' as [Inco terms],
	'' as [Distance qualifier],
	'' as [Distance],
	'' as [Bill Of Lading],
	'' as [MAWB],
	'' as [HAWB],
	'' as [PO Number],
	'Air' as [Reference number],
    ISNULL(
        (SELECT TOP 1 MilestoneTime FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeH.HouseNo AND PRODUCTLINE = 2 AND MileStoneID = 2970 AND Status = 1),
        (SELECT TOP 1 MilestoneTime FROM RESM..SMMilestoneHistory WHERE KEYVALUE = aeH.HouseNo AND PRODUCTLINE = 2 AND MileStoneID = 3000 AND Status = 1)
    ) AS [Delivery Date],
	'' as [Delivery Time],
	'' as [POD Name],
	CASE WHEN aeH.ActPCSUOM = 'PLT' THEN aeH.ActPCS ELSE NULL END AS [Notes1 (Pallets)],
    CASE WHEN aeH.ActPCSUOM = 'CTN' THEN aeH.ActPCS ELSE NULL END AS [Notes2 (Cartons)],
	aeH.HouseNo as [Notes3],
	'' as [Notes4],
	'' as [Notes5],
	'' as [Notes6],
	'' as[Notes7],
	'' as [Notes8],
	'' as [Notes9],
	'' as [Notes10],
	'' as [Notes11],
	'' as [Notes12],
	'' as [Notes13],
	'' as [Notes14],
	'' as [Notes15],
	'' as [Notes16],
	 OriginAirport.AirportCode as [Origin(Air)Port], 
	 DestAirport.AirportCode AS [Destination (Air)Port],
	 '' as [Shipper Location],
	 shp.customername as [Shipper Name],
	 shp.customername as [Shipper Company],
	 shp.customeraddress1 + shp.CustomerAddress2 as [Shipper Address1],
	 '' as [Shipper Address2],
	 '' as [Shipper Address3],
	 city.CityName as [Shipper City],
	 '' as [Shipper state],
	 '' as [Shipper Postcode],
	 city.CityName as [Shipper Country],
	 '' as [Consignee Location],
	  cnee.customername as [Consignee Name],
	  cnee.customername as [Consignee Company],
	  cnee.CustomerAddress1 +cnee.CustomerAddress2 as [Consignee Address1],
	  '' as [Consignee Address2],
	  '' as [Consignee Address3],
	  dstn.CityName as [Consignee City],
	  '' as [Consignee State],
	  cnee.CustomerAddress1 +cnee.CustomerAddress2 as [Consignee Postcode],
	  country.countryname as [Consignee Country],
	  inv.InvoiceAmount as [Amount Billed],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 13 ) as [Freight Amount - FRT],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 13 ) as [Total Shipment VAT Amount],
	  'FSC' AS [Accessorial Charge- FSC],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 14 ) as [Accessorial Charge-FSC Amount1 - Fuel Surcharge],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 14 ) as [VAT liable flag 1],
	  'SSC' as [Accessorial Charge-SSC],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 15 ) as [Accessorial Charge-SSC Amount2],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 15 ) as [VAT liable flag 2],
	  '' as [Accessorial Charge-P/U Code3],
	  '' as [Accessorial Charge- Amount3],
	  '' as [VAT liable flag 3],
	  'H/C' AS [Accessorial Charge-Origin Terminal Code4],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 400 ) as [Accessorial Charge-Origin Terminal Amount4],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 400 ) as [VAT liable flag 4],
	  '' AS [Accessorial Charge-CUS Code5],
	  '' AS [Accessorial Charge-CUS Amount5],
	  '' AS [VAT liable flag 5],
	  'H/C' as [Accessorial Charge-Dest Terminal Code6],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 005 ) as [Accessorial Charge-dest Terminal Amount6],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 005 ) as [VAT liable flag 6],
	  'DEL' AS [Accessorial Charge Code7],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 113 ) as [Accessorial Charge Amount7],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 113 ) as [VAT liable flag 7],
	  'CUS' AS [Accessorial Charge Code8],
	  (SELECT SUM(SAmount) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 001 ) as [Accessorial Charge Amount8],
	  (SELECT SUM(SLocalEquiv) FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID AND ChargeID = 001 ) as [VAT liable flag 8],
	  '' as [Accessorial Charge Code9],
	  '' as [Accessorial Charge Amount9],
	  '' as [VAT liable flag 9],
	  '' as [Accessorial Charge Code10],
	  '' as [Accessorial Charge Amount10],
	  '' as[Accessorial 10 VAT code],
	  'SGD' AS [Currency],
	  '' AS [Bill to Location],
	  'Arrow Electronics Asia (S) Pte Ltd' AS [Bill To Company],
	  'NIC, 5 Tai Seng Drive #06-01' AS [Bill To Address1],
	  '' AS [Bill To Address2],
	  'Singapore' AS [Bill To City],
	  '' AS [Bill To State],
	  '535217' AS [Bill To Postcode],
	  'SG' AS [Bill To Country],
	  '' AS [Equipment Type],
	   '' AS [Container number],
	   '' AS [Trailer/Car ID],
	   '' AS [carrier VAT number],
	   '' AS [CLIENT VAT NUMBER],
	   '' AS [VAT percentage],
	   (SELECT TOP 1 STransactionRate FROM FMInvoiceDetail WHERE InvoiceNo = inv.InvoiceNo AND StationID = inv.StationID) AS [Currency Exchange Rate],
	   '' AS [Overtime hours],
	   '' AS [Waiting hours],
	   '' AS [Urgent fee location count]
	  
FROM
    FMOPSource AS inv 
LEFT JOIN
    TPAIRData AS aeH ON inv.SourceID = aeH.ID AND inv.StationID = aeH.StationID
LEFT JOIN
    dbo.smcustomerone AS cus ON aeH.CustomerID = cus.CustomerID
LEFT JOIN 
	RESM..SMAirPort as DestAirport ON DestAirport.HQID = aeH.airportofdstn
LEFT JOIN 
	RESM..SMAirPort as OriginAirport ON OriginAirport.HQID = aeH.airportofdept
LEFT JOIN
	dbo.smcustomerone AS shp ON aeH.Shipper = shp.CustomerID
Left Join
	RESM..SMCITY as city on city.HQID = aeh.Receipt
LEFT JOIN
    dbo.smcustomerone AS cnee ON aeH.CNEE = cnee.CustomerID
Left Join
	RESM..SMCITY as dstn on dstn.HQID = aeh.DSTN
Left Join
	RESM..smcountry as country on country.HQID = cnee.Country
Left join
	dbo.smcustomerone as billto on inv.billto = billto.CustomerID
WHERE
    inv.ModeCode = '3A'        
    AND cus.CustomerCode = 'ARROWE001'
	and billto.CustomerCode='ARROWSG'
	AND ISNULL(inv.InvoiceNo, '') <> '' 
