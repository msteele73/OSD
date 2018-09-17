IF OBJECT_ID('tempdb..#TMP_Devices') IS NOT NULL
DROP TABLE #TMP_Devices

--GET DATA FROM H01
--USE CM_H01

;WITH SRR
(ResourceID, School, ComputerName, Vendor, Model, BuildNumber, OperatingSystem) AS
(
	SELECT
		cm.ResourceID,
		cm.AD_Site_Name0 AS School,
		cm.Name0 AS Computername,
		dev.Manufacturer0 AS Vendor,
		dev.Model0 AS Model,
		cm.Build01 AS BuildNumber,
		os.Caption0 AS OperatingSystem
	FROM
		(
			SELECT
				ResourceID,
				AD_Site_Name0,
				Build01,
				CASE
					WHEN Name0 = 'Unknown' THEN 'Unknown (ResourceID: ' + CONVERT(varchar, ResourceID) + ')'
					ELSE Name0
				END AS Name0
			FROM [CM_H01].[dbo].[v_R_System] vrs
			where vrs.Operating_System_Name_and0 like '%workstation%'
		) AS cm
		JOIN
		(
			SELECT
				ResourceID,
				Manufacturer0,
				CASE
					WHEN Model0 LIKE '%23-c021a%' THEN 'HP ENVY 23-c021a'
					WHEN Model0 LIKE '23-n100a' THEN 'HP ENVY 23-N100A'
					WHEN Model0 LIKE '24-b014a' THEN 'HP Pavilion All-in-One (Touch)'
					WHEN Model0 LIKE '500-410a' THEN 'HP Pavilion Desktop - 500-410a'
					WHEN Model0 LIKE '500-1010a' THEN 'HP TouchSmart 520-1010a Desktop PC'
					WHEN Model0 LIKE 'Compaq 621' THEN 'Compaq 621 Notebook PC'
					WHEN Model0 LIKE '%HP dx2810 MT%' THEN 'HP Compaq dx2810 Microtower PC'
					WHEN Model0 LIKE '%TM6595T%' THEN 'TravelMate 6595T'
					WHEN Model0 LIKE '%TMP455-MG%' THEN 'TravelMate P455-MG'
					WHEN Model0 LIKE '%EX5630%' THEN 'Extensa 5630'
					WHEN Model0 LIKE '%AO533%' THEN 'Aspire ONE 533'
					WHEN Model0 LIKE '%4333%' THEN 'IdeaPad S10e'
					WHEN Model0 LIKE '%80TG%' OR Model0 LIKE '%80TL%' THEN 'IdeaPad V110'
					WHEN Model0 LIKE '%2038%' THEN 'IdeaPad Y330'
					WHEN Model0 LIKE '%80Q7%' THEN 'IdeaPad 300'
					WHEN Model0 LIKE '%80U2%' THEN 'Yoga 310'
					WHEN Model0 LIKE '%2032%' THEN 'Miix 2 11'
					WHEN Model0 LIKE '%80U1%' THEN 'Miix 2 510'
					WHEN Model0 LIKE '%9211%' THEN 'ThinkCentre M52'
					WHEN Model0 LIKE '%8813%' THEN 'ThinkCentre M55'
					WHEN Model0 LIKE '%6072%' OR Model0 LIKE '%6074%' OR Model0 LIKE '%6077%' THEN 'ThinkCentre M57'
					WHEN Model0 LIKE '%7627%' OR Model0 LIKE '%7360%' OR Model0 LIKE '%7373%' OR Model0 LIKE '%7826%' THEN 'ThinkCentre M58'
					WHEN Model0 LIKE '%7483%' OR Model0 LIKE '%6234%' THEN 'ThinkCentre M58p'
					WHEN Model0 LIKE '%10KQ%' THEN 'ThinkCentre M700'
					WHEN Model0 LIKE '%10M8%' THEN 'ThinkCentre M710s'
					WHEN Model0 LIKE '%3132%' OR Model0 LIKE '%3134%' OR Model0 LIKE '%3140%' THEN 'ThinkCentre M71e'
					WHEN Model0 LIKE '%10SU%' THEN 'ThinkCentre M720'				
					WHEN Model0 LIKE '%6578%' THEN 'ThinkCentre M53'
					WHEN Model0 LIKE '%8808%' OR Model0 LIKE '%8814%' OR Model0 LIKE '%8816%' THEN 'ThinkCentre M55'
					WHEN Model0 LIKE '%9632%' THEN 'ThinkCentre M55e'
					WHEN Model0 LIKE '%6081%' THEN 'ThinkCentre M57'
					WHEN Model0 LIKE '%10B7%' THEN 'ThinkCentre M57p'
					WHEN Model0 LIKE '%7628%' THEN 'ThinkCentre M58'
					WHEN Model0 LIKE '%7479%' OR Model0 LIKE '%7630%' THEN 'ThinkCentre M58p'
					WHEN Model0 LIKE '%10FX%' THEN 'ThinkCentre M800 SFF'
					WHEN Model0 LIKE '%5049%' THEN 'ThinkCentre M81'
					WHEN Model0 LIKE '%2756%' THEN 'ThinkCentre M82'
					WHEN Model0 LIKE '%10FG%' OR Model0 LIKE '%10FL%' THEN 'ThinkCentre M900 SFF'
					WHEN Model0 LIKE '%5536%' OR Model0 LIKE '%4524%' OR Model0 LIKE '%5498%' THEN 'ThinkCentre M90p'
					WHEN Model0 LIKE '%5205%' THEN 'ThinkCentre M90z - ALL-in-ONE'
					WHEN Model0 LIKE '%10MU%' THEN 'ThinkCentre M910q TFF'
					WHEN Model0 LIKE '%10ML%' THEN 'ThinkCentre M910s SFF'
					WHEN Model0 LIKE '%4518%' OR Model0 LIKE '%4524%' THEN 'ThinkCentre M91p'
					WHEN Model0 LIKE '%3227%' THEN 'ThinkCentre M92p'
					WHEN Model0 LIKE '%3238%' OR Model0 LIKE '%3209%' THEN 'ThinkCentre M92p TFF'
					WHEN Model0 LIKE '%10A8%' OR Model0 LIKE '%10AA%' THEN 'ThinkCentre M93p SFF(or TFF)'
					WHEN Model0 LIKE '%1S75%' OR Model0 LIKE '%F0CU%' THEN 'ThinkCentre ALL-in-ONE'
					WHEN Model0 LIKE '%20C3%' THEN 'ThinkPad 10 Tablet G1'
					WHEN Model0 LIKE '%20E4%' THEN 'ThinkPad 10 Tablet G2'
					WHEN Model0 LIKE '%20E6%' OR Model0 LIKE '%20E8%' THEN 'ThinkPad Yoga 11e'
					WHEN Model0 LIKE '%20DA%' THEN 'ThinkPad Yoga 11e G2'
					WHEN Model0 LIKE '%20G9%' THEN 'ThinkPad Yoga 11e G3'
					WHEN Model0 LIKE '%20HS%' OR Model0 LIKE '%20HU%' THEN 'ThinkPad Yoga 11e G4'
					WHEN Model0 LIKE '%20LN%' THEN 'ThinkPad Yoga 11e G5'
					WHEN Model0 LIKE '%0328%' THEN 'ThinkPad Edge 11'
					WHEN Model0 LIKE '%1141%' THEN 'ThinkPad Edge E420'
					WHEN Model0 LIKE '%3259%' THEN 'ThinkPad Edge E530'
					WHEN Model0 LIKE '%6885%' THEN 'ThinkPad Edge E531'
					WHEN Model0 LIKE '%20H1%' THEN 'ThinkPad E470'
					WHEN Model0 LIKE '%20M7%' OR Model0 LIKE '%20M8%' THEN 'ThinkPad Yoga L380'
					WHEN Model0 LIKE '%2468%' THEN 'ThinkPad L430'
					WHEN Model0 LIKE '%2597%' OR Model0 LIKE '%2598%' THEN 'ThinkPad L512'
					WHEN Model0 LIKE '%5019%' THEN 'ThinkPad L520'
					WHEN Model0 LIKE '%3506%' OR Model0 LIKE '%3507%' THEN 'ThinkPad Mini 10'
					WHEN Model0 LIKE '%2912%' OR Model0 LIKE '%2522%' OR Model0 LIKE '%2537%' THEN 'ThinkPad T410'
					WHEN Model0 LIKE '%2924%' THEN 'ThinkPad T410s'
					WHEN Model0 LIKE '%2356%' OR Model0 LIKE '%2344%' THEN 'ThinkPad T430'
					WHEN Model0 LIKE '%20B7%' THEN 'ThinkPad T440'
					WHEN Model0 LIKE '%20AN%' THEN 'ThinkPad T440p'
					WHEN Model0 LIKE '%20AR%' THEN 'ThinkPad T440s'
					WHEN Model0 LIKE '%20BU%' THEN 'ThinkPad T450'
					WHEN Model0 LIKE '%20FM%' OR Model0 LIKE '%20FN%' THEN 'ThinkPad T460'
					WHEN Model0 LIKE '%20JN%' OR Model0 LIKE '%20HE%' THEN 'ThinkPad T470'
					WHEN Model0 LIKE '%20L5%' OR Model0 LIKE '%20L6%' THEN 'ThinkPad T480'
					WHEN Model0 LIKE '%2087%' THEN 'ThinkPad T500'
					WHEN Model0 LIKE '%4240%' OR Model0 LIKE '%4242%' THEN 'ThinkPad T520'
					WHEN Model0 LIKE '%2394%' OR Model0 LIKE '%2355%' THEN 'ThinkPad T530'
					WHEN Model0 LIKE '%20KJ%' OR Model0 LIKE '%20KK%' THEN 'ThinkPad X1 Tablet G3'
					WHEN Model0 LIKE '%1293%' OR Model0 LIKE '%3448%' THEN 'ThinkPad X1 Carbon G1'
					WHEN Model0 LIKE '%20BT%' THEN 'ThinkPad X1 Carbon G2'
					WHEN Model0 LIKE '%20A8%' THEN 'ThinkPad X1 Carbon G3'
					WHEN Model0 LIKE '%20FB%' OR Model0 LIKE '%20FC%' THEN 'ThinkPad X1 Carbon G4'
					WHEN Model0 LIKE '%20K3%' OR Model0 LIKE '%20K4%' THEN 'ThinkPad X1 Carbon G5'
					WHEN Model0 LIKE '%20FR%' OR Model0 LIKE '%20FQ%' THEN 'ThinkPad X1 Yoga G1'
					WHEN Model0 LIKE '%2338%' OR Model0 LIKE '%FL9B%' OR Model0 LIKE '%FL9X%' THEN 'ThinkPad x130e'
					WHEN Model0 LIKE '%3367%' OR Model0 LIKE '%3369%' THEN 'ThinkPad X131e'
					WHEN Model0 LIKE '%3093%' OR Model0 LIKE '%3239%' OR Model0 LIKE '%3357%' THEN 'ThinkPad X201 Tablet'
					WHEN Model0 LIKE '%4298%' THEN 'ThinkPad X220 Tablet'
					WHEN Model0 LIKE '%3701%' THEN 'ThinkPad Helix'
					WHEN Model0 LIKE '%30AJ%' THEN 'ThinkStation P300'
					WHEN Model0 LIKE '%30AU%' THEN 'ThinkStation P310'
					WHEN Model0 LIKE '%4157%' THEN 'ThinkStation S20'
					WHEN Model0 LIKE '%0569%' THEN 'ThinkStation S30'
					WHEN Model0 LIKE '%4222%' THEN 'ThinkStation E20'
					WHEN Model0 LIKE '%7783%' THEN 'ThinkStation E30'
					WHEN Model0 LIKE '%80FY%' THEN 'Lenovo G40-30'
					WHEN Model0 LIKE '%80E3%' THEN 'Lenovo G50-45'
					WHEN Model0 LIKE '%80DU%' THEN 'Lenovo Y70 Touch'
					WHEN Model0 LIKE '%4637%' THEN 'Docking Station'
					WHEN Model0 LIKE '%5409%' OR Model0 LIKE '%10SK%' OR Model0 LIKE '%L3D0%' OR Model0 LIKE '%5454%' OR Model0 LIKE '%L3E1%' THEN 'CANNOT BE IDENTIFIED'
				ELSE Model0
				END AS Model0
			FROM [CM_H01].[dbo].[v_GS_COMPUTER_SYSTEM]
		) AS dev ON dev.ResourceID = cm.ResourceID
		JOIN
		(
			SELECT
				ResourceID,
				Caption0
			FROM [CM_H01].[dbo].[v_GS_OPERATING_SYSTEM]
		) AS os ON os.ResourceID = cm.ResourceID
)
-- Select out a Consolidated Row
SELECT
	ResourceID,
	School,
	ComputerName,
	Vendor,
	Model,
	CASE
		WHEN BuildNumber = '6.0.6002' THEN 6002 -- Windows Vista SP2
		WHEN BuildNumber = '6.1.7600' THEN 7600 -- Windows 7 
		WHEN BuildNumber = '6.1.7601' THEN 7601 -- Windows 7 SP1
		WHEN BuildNumber = '6.2.9200' THEN 9200 -- Windows 8
		WHEN BuildNumber = '6.3.9600' THEN 9600 -- Windows 8.1
		WHEN BuildNumber = '10.0.10240' THEN 1507 -- Windows 10
		WHEN BuildNumber = '10.0.10586' THEN 1511 -- Windows 10
		WHEN BuildNumber = '10.0.14393' THEN 1607 -- Windows 10
		WHEN BuildNumber = '10.0.15063' THEN 1703 -- Windows 10
		WHEN BuildNumber = '10.0.16299' THEN 1709 -- Windows 10
		WHEN BuildNumber = '10.0.17134' THEN 1803 -- Windows 10
		WHEN BuildNumber = '10.0.17751' THEN 1809 -- Windows 10
		WHEN BuildNumber = '10.0.17713' THEN 1809 -- Windows 10
		WHEN BuildNumber = '10.0.17723' THEN 1809 -- Windows 10
		ELSE BuildNumber
	END AS BuildNumber,
	OperatingSystem
INTO #TMP_Devices
FROM SRR
WHERE Model <> 'Virtual Machine'

-------------------------------------------------------------------------------
-------------------------------------------------------------------------------

--GET DATA FROM H02
--USE CM_H02

;WITH SRR
(ResourceID, School, ComputerName, Vendor, Model, BuildNumber, OperatingSystem) AS
(
	SELECT
		cm.ResourceID,
		cm.AD_Site_Name0 AS School,
		cm.Name0 AS Computername,
		dev.Manufacturer0 AS Vendor,
		dev.Model0 AS Model,
		cm.Build01 AS BuildNumber,
		os.Caption0 AS OperatingSystem
	FROM
		(
			SELECT
				ResourceID,
				AD_Site_Name0,
				Build01,
				CASE
					WHEN Name0 = 'Unknown' THEN 'Unknown (ResourceID: ' + CONVERT(varchar, ResourceID) + ')'
					ELSE Name0
				END AS Name0
			FROM [CM_H02].[dbo].[v_R_System] vrs
			where vrs.Operating_System_Name_and0 like '%workstation%'
		) AS cm
		JOIN
		(
			SELECT
				ResourceID,
				Manufacturer0,
				CASE
					WHEN Model0 LIKE '%23-c021a%' THEN 'HP ENVY 23-c021a'
					WHEN Model0 LIKE '23-n100a' THEN 'HP ENVY 23-N100A'
					WHEN Model0 LIKE '24-b014a' THEN 'HP Pavilion All-in-One (Touch)'
					WHEN Model0 LIKE '500-410a' THEN 'HP Pavilion Desktop - 500-410a'
					WHEN Model0 LIKE '500-1010a' THEN 'HP TouchSmart 520-1010a Desktop PC'
					WHEN Model0 LIKE 'Compaq 621' THEN 'Compaq 621 Notebook PC'
					WHEN Model0 LIKE '%HP dx2810 MT%' THEN 'HP Compaq dx2810 Microtower PC'
					WHEN Model0 LIKE '%TM6595T%' THEN 'TravelMate 6595T'
					WHEN Model0 LIKE '%TMP455-MG%' THEN 'TravelMate P455-MG'
					WHEN Model0 LIKE '%EX5630%' THEN 'Extensa 5630'
					WHEN Model0 LIKE '%AO533%' THEN 'Aspire ONE 533'
					WHEN Model0 LIKE '%4333%' THEN 'IdeaPad S10e'
					WHEN Model0 LIKE '%80TG%' OR Model0 LIKE '%80TL%' THEN 'IdeaPad V110'
					WHEN Model0 LIKE '%2038%' THEN 'IdeaPad Y330'
					WHEN Model0 LIKE '%80Q7%' THEN 'IdeaPad 300'
					WHEN Model0 LIKE '%80U2%' THEN 'Yoga 310'
					WHEN Model0 LIKE '%2032%' THEN 'Miix 2 11'
					WHEN Model0 LIKE '%80U1%' THEN 'Miix 2 510'
					WHEN Model0 LIKE '%9211%' THEN 'ThinkCentre M52'
					WHEN Model0 LIKE '%8813%' THEN 'ThinkCentre M55'
					WHEN Model0 LIKE '%6072%' OR Model0 LIKE '%6074%' OR Model0 LIKE '%6077%' THEN 'ThinkCentre M57'
					WHEN Model0 LIKE '%7627%' OR Model0 LIKE '%7360%' OR Model0 LIKE '%7373%' OR Model0 LIKE '%7826%' THEN 'ThinkCentre M58'
					WHEN Model0 LIKE '%7483%' OR Model0 LIKE '%6234%' THEN 'ThinkCentre M58p'
					WHEN Model0 LIKE '%10KQ%' THEN 'ThinkCentre M700'
					WHEN Model0 LIKE '%10M8%' THEN 'ThinkCentre M710s'
					WHEN Model0 LIKE '%3132%' OR Model0 LIKE '%3134%' OR Model0 LIKE '%3140%' THEN 'ThinkCentre M71e'
					WHEN Model0 LIKE '%10SU%' THEN 'ThinkCentre M720'				
					WHEN Model0 LIKE '%6578%' THEN 'ThinkCentre M53'
					WHEN Model0 LIKE '%8808%' OR Model0 LIKE '%8814%' OR Model0 LIKE '%8816%' THEN 'ThinkCentre M55'
					WHEN Model0 LIKE '%9632%' THEN 'ThinkCentre M55e'
					WHEN Model0 LIKE '%6081%' THEN 'ThinkCentre M57'
					WHEN Model0 LIKE '%10B7%' THEN 'ThinkCentre M57p'
					WHEN Model0 LIKE '%7628%' THEN 'ThinkCentre M58'
					WHEN Model0 LIKE '%7479%' OR Model0 LIKE '%7630%' THEN 'ThinkCentre M58p'
					WHEN Model0 LIKE '%10FX%' THEN 'ThinkCentre M800 SFF'
					WHEN Model0 LIKE '%5049%' THEN 'ThinkCentre M81'
					WHEN Model0 LIKE '%2756%' THEN 'ThinkCentre M82'
					WHEN Model0 LIKE '%10FG%' OR Model0 LIKE '%10FL%' THEN 'ThinkCentre M900 SFF'
					WHEN Model0 LIKE '%5536%' OR Model0 LIKE '%4524%' OR Model0 LIKE '%5498%' THEN 'ThinkCentre M90p'
					WHEN Model0 LIKE '%5205%' THEN 'ThinkCentre M90z - ALL-in-ONE'
					WHEN Model0 LIKE '%10MU%' THEN 'ThinkCentre M910q TFF'
					WHEN Model0 LIKE '%10ML%' THEN 'ThinkCentre M910s SFF'
					WHEN Model0 LIKE '%4518%' OR Model0 LIKE '%4524%' THEN 'ThinkCentre M91p'
					WHEN Model0 LIKE '%3227%' THEN 'ThinkCentre M92p'
					WHEN Model0 LIKE '%3238%' OR Model0 LIKE '%3209%' THEN 'ThinkCentre M92p TFF'
					WHEN Model0 LIKE '%10A8%' OR Model0 LIKE '%10AA%' THEN 'ThinkCentre M93p SFF(or TFF)'
					WHEN Model0 LIKE '%1S75%' OR Model0 LIKE '%F0CU%' THEN 'ThinkCentre ALL-in-ONE'
					WHEN Model0 LIKE '%20C3%' THEN 'ThinkPad 10 Tablet G1'
					WHEN Model0 LIKE '%20E4%' THEN 'ThinkPad 10 Tablet G2'
					WHEN Model0 LIKE '%20E6%' OR Model0 LIKE '%20E8%' THEN 'ThinkPad Yoga 11e'
					WHEN Model0 LIKE '%20DA%' THEN 'ThinkPad Yoga 11e G2'
					WHEN Model0 LIKE '%20G9%' THEN 'ThinkPad Yoga 11e G3'
					WHEN Model0 LIKE '%20HS%' OR Model0 LIKE '%20HU%' THEN 'ThinkPad Yoga 11e G4'
					WHEN Model0 LIKE '%20LN%' THEN 'ThinkPad Yoga 11e G5'
					WHEN Model0 LIKE '%0328%' THEN 'ThinkPad Edge 11'
					WHEN Model0 LIKE '%1141%' THEN 'ThinkPad Edge E420'
					WHEN Model0 LIKE '%3259%' THEN 'ThinkPad Edge E530'
					WHEN Model0 LIKE '%6885%' THEN 'ThinkPad Edge E531'
					WHEN Model0 LIKE '%20H1%' THEN 'ThinkPad E470'
					WHEN Model0 LIKE '%20M7%' OR Model0 LIKE '%20M8%' THEN 'ThinkPad Yoga L380'
					WHEN Model0 LIKE '%2468%' THEN 'ThinkPad L430'
					WHEN Model0 LIKE '%2597%' OR Model0 LIKE '%2598%' THEN 'ThinkPad L512'
					WHEN Model0 LIKE '%5019%' THEN 'ThinkPad L520'
					WHEN Model0 LIKE '%3506%' OR Model0 LIKE '%3507%' THEN 'ThinkPad Mini 10'
					WHEN Model0 LIKE '%2912%' OR Model0 LIKE '%2522%' OR Model0 LIKE '%2537%' THEN 'ThinkPad T410'
					WHEN Model0 LIKE '%2924%' THEN 'ThinkPad T410s'
					WHEN Model0 LIKE '%2356%' OR Model0 LIKE '%2344%' THEN 'ThinkPad T430'
					WHEN Model0 LIKE '%20B7%' THEN 'ThinkPad T440'
					WHEN Model0 LIKE '%20AN%' THEN 'ThinkPad T440p'
					WHEN Model0 LIKE '%20AR%' THEN 'ThinkPad T440s'
					WHEN Model0 LIKE '%20BU%' THEN 'ThinkPad T450'
					WHEN Model0 LIKE '%20FM%' OR Model0 LIKE '%20FN%' THEN 'ThinkPad T460'
					WHEN Model0 LIKE '%20JN%' OR Model0 LIKE '%20HE%' THEN 'ThinkPad T470'
					WHEN Model0 LIKE '%20L5%' OR Model0 LIKE '%20L6%' THEN 'ThinkPad T480'
					WHEN Model0 LIKE '%2087%' THEN 'ThinkPad T500'
					WHEN Model0 LIKE '%4240%' OR Model0 LIKE '%4242%' THEN 'ThinkPad T520'
					WHEN Model0 LIKE '%2394%' OR Model0 LIKE '%2355%' THEN 'ThinkPad T530'
					WHEN Model0 LIKE '%20KJ%' OR Model0 LIKE '%20KK%' THEN 'ThinkPad X1 Tablet G3'
					WHEN Model0 LIKE '%1293%' OR Model0 LIKE '%3448%' THEN 'ThinkPad X1 Carbon G1'
					WHEN Model0 LIKE '%20BT%' THEN 'ThinkPad X1 Carbon G2'
					WHEN Model0 LIKE '%20A8%' THEN 'ThinkPad X1 Carbon G3'
					WHEN Model0 LIKE '%20FB%' OR Model0 LIKE '%20FC%' THEN 'ThinkPad X1 Carbon G4'
					WHEN Model0 LIKE '%20K3%' OR Model0 LIKE '%20K4%' THEN 'ThinkPad X1 Carbon G5'
					WHEN Model0 LIKE '%20FR%' OR Model0 LIKE '%20FQ%' THEN 'ThinkPad X1 Yoga G1'
					WHEN Model0 LIKE '%2338%' OR Model0 LIKE '%FL9B%' OR Model0 LIKE '%FL9X%' THEN 'ThinkPad x130e'
					WHEN Model0 LIKE '%3367%' OR Model0 LIKE '%3369%' THEN 'ThinkPad X131e'
					WHEN Model0 LIKE '%3093%' OR Model0 LIKE '%3239%' OR Model0 LIKE '%3357%' THEN 'ThinkPad X201 Tablet'
					WHEN Model0 LIKE '%4298%' THEN 'ThinkPad X220 Tablet'
					WHEN Model0 LIKE '%3701%' THEN 'ThinkPad Helix'
					WHEN Model0 LIKE '%30AJ%' THEN 'ThinkStation P300'
					WHEN Model0 LIKE '%30AU%' THEN 'ThinkStation P310'
					WHEN Model0 LIKE '%4157%' THEN 'ThinkStation S20'
					WHEN Model0 LIKE '%0569%' THEN 'ThinkStation S30'
					WHEN Model0 LIKE '%4222%' THEN 'ThinkStation E20'
					WHEN Model0 LIKE '%7783%' THEN 'ThinkStation E30'
					WHEN Model0 LIKE '%80FY%' THEN 'Lenovo G40-30'
					WHEN Model0 LIKE '%80E3%' THEN 'Lenovo G50-45'
					WHEN Model0 LIKE '%80DU%' THEN 'Lenovo Y70 Touch'
					WHEN Model0 LIKE '%4637%' THEN 'Docking Station'
					WHEN Model0 LIKE '%5409%' OR Model0 LIKE '%10SK%' OR Model0 LIKE '%L3D0%' OR Model0 LIKE '%5454%' OR Model0 LIKE '%L3E1%' THEN 'CANNOT BE IDENTIFIED'
				ELSE Model0
				END AS Model0
			FROM [CM_H02].[dbo].[v_GS_COMPUTER_SYSTEM]
		) AS dev ON dev.ResourceID = cm.ResourceID
		JOIN
		(
			SELECT
				ResourceID,
				Caption0
			FROM [CM_H02].[dbo].[v_GS_OPERATING_SYSTEM]
		) AS os ON os.ResourceID = cm.ResourceID
)
-- Select out a Consolidated Row
INSERT INTO #TMP_Devices
SELECT
	ResourceID,
	School,
	ComputerName,
	Vendor,
	Model,
	CASE
		WHEN BuildNumber = '6.0.6002' THEN 6002 -- Windows Vista SP2
		WHEN BuildNumber = '6.1.7600' THEN 7600 -- Windows 7 
		WHEN BuildNumber = '6.1.7601' THEN 7601 -- Windows 7 SP1
		WHEN BuildNumber = '6.2.9200' THEN 9200 -- Windows 8
		WHEN BuildNumber = '6.3.9600' THEN 9600 -- Windows 8.1
		WHEN BuildNumber = '10.0.10240' THEN 1507 -- Windows 10
		WHEN BuildNumber = '10.0.10586' THEN 1511 -- Windows 10
		WHEN BuildNumber = '10.0.14393' THEN 1607 -- Windows 10
		WHEN BuildNumber = '10.0.15063' THEN 1703 -- Windows 10
		WHEN BuildNumber = '10.0.16299' THEN 1709 -- Windows 10
		WHEN BuildNumber = '10.0.17134' THEN 1803 -- Windows 10
		WHEN BuildNumber = '10.0.17751' THEN 1809 -- Windows 10
		WHEN BuildNumber = '10.0.17713' THEN 1809 -- Windows 10
		WHEN BuildNumber = '10.0.17723' THEN 1809 -- Windows 10
		ELSE BuildNumber
	END AS BuildNumber,
	OperatingSystem
FROM SRR
WHERE Model <> 'Virtual Machine'

-------------------------------------------------------------------------------
-- Select out the Consolidated Rows
-------------------------------------------------------------------------------
IF OBJECT_ID('tempdb..#TMP_Year') IS NOT NULL
DROP TABLE #TMP_Year

SELECT
	ResourceID,
	School,
	ComputerName,
	Vendor,
	Model,
	BuildNumber,
	OperatingSystem,
	CASE
					---------- ACER ----------
					WHEN Model = 'Extensa 5630' OR Model = 'EX5630' THEN 2008
					WHEN Model = 'TravelMate 6592' THEN 2010
					WHEN Model = 'TravelMate 6595T' THEN 2012
					WHEN Model = 'TravelMate 8572' THEN 2011
					WHEN Model = 'TravelMate 8572T' THEN 2010
					WHEN Model = 'TravelMate 8572G' THEN 2010
					WHEN Model = 'TravelMate B113' THEN 2013
					WHEN Model = 'TravelMate B115-M' THEN 2014
					WHEN Model = 'TravelMate B117-M' THEN 2016
					WHEN Model = 'TravelMate P653-V' THEN 2013
					WHEN Model = 'Veriton S6630G' OR Model = 'Veriton S6630' OR Model = 'VeritonS6630G' OR Model = 'VS6630G' OR Model = 'S6630G' THEN 2014
					WHEN Model = 'Veriton M6630G' THEN 2015
					WHEN Model = 'Veriton L4630G' THEN 2015
					WHEN Model = 'Aspire ONE 533' THEN 2010
					WHEN Model = 'Aspire 5253' THEN 2011
					WHEN Model = 'Aspire 5542' THEN 2010
					WHEN Model = 'Aspire 5733Z' THEN 2012
					WHEN Model = 'Aspire 5738' THEN 2009
					WHEN Model = 'Aspire 5750' THEN 2012
					WHEN Model = 'Aspire 8942G' THEN 2010
					WHEN Model = 'Aspire E1-571' THEN 2012
					WHEN Model = 'Aspire E5-553G' THEN 2017
					WHEN Model = 'Aspire F5-573G' THEN 2016
					WHEN Model = 'Aspire V3-371' THEN 2015
					WHEN Model = 'Aspire V3-571G' THEN 2012
					WHEN Model = 'Aspire X3950' THEN 2010
					WHEN Model = 'Aspire X3990' THEN 2008
					WHEN Model = 'Aspire 7720' THEN 2007
					WHEN Model = 'Aspire ES1-522' THEN 2013
					WHEN Model = 'Aspire Z5710' THEN 2010
					WHEN Model = 'Veriton Z6611G' THEN 2011
					WHEN Model = 'Veriton L4620G' THEN 2015
					WHEN Model = 'Aspire ES1-523' THEN 2017
					WHEN Model = 'Aspire 7720' THEN 2007
					WHEN Model = 'TravelMate8572TG' THEN 2010
					WHEN Model = 'TravelMate P653' THEN 2012
					WHEN Model = 'TRAVELMATE B115' THEN 2014
					WHEN Model = 'TravelMate 8571' THEN 2009
					WHEN Model = 'TravelMate 6595TG' THEN 2012
					WHEN Model = 'TravelMate 5742' THEN 2011
					WHEN Model = 'TravelMate 5730' THEN 2008
					WHEN Model = 'ICONIA W700' THEN 2013
					WHEN Model = 'TravelMate P455-MG' THEN 2017
					---------- DELL ----------
					WHEN Model = 'Inspiron 13-5378' THEN 2017
					WHEN Model = 'Inspiron 15-3565' THEN 2017
					WHEN Model = 'Inspiron 5379' THEN 2018
					WHEN Model = 'Inspiron 5567' THEN 2012
					WHEN Model = 'Inspiron 5559' THEN 2016
					WHEN Model = 'Inspiron 1545' THEN 2009
					WHEN Model = 'Inspiron 545s' THEN 2009
					WHEN Model = 'Latitude 3330' THEN 2014
					WHEN Model = 'Latitude E5440' THEN 2014
					WHEN Model = 'Latitude E5540' THEN 2014
					WHEN Model = 'Latitude E6220' THEN 2011
					WHEN Model = 'Latitude E6230' THEN 2012
					WHEN Model = 'Latitude E6320' THEN 2011
					WHEN Model = 'Latitude E6330' THEN 2012
					WHEN Model = 'Latitude E6420' THEN 2011
					WHEN Model = 'Latitude E6430' THEN 2012
					WHEN Model = 'Latitude E6440' THEN 2013
					WHEN Model = 'Latitude E6520' THEN 2011
					WHEN Model = 'Latitude E7470' THEN 2016
					WHEN Model = 'OptiPlex GX620' THEN 2005
					WHEN Model = 'Optiplex 3020' THEN 2014
					WHEN Model = 'Optiplex 330' THEN 2007
					WHEN Model = 'Optiplex 360' THEN 2008
					WHEN Model = 'Optiplex 7010' THEN 2012
					WHEN Model = 'Optiplex 745' THEN 2007
					WHEN Model = 'Optiplex 7040' THEN 2012
					WHEN Model = 'Optiplex 755' THEN 2008
					WHEN Model = 'Optiplex 760' THEN 2009
					WHEN Model = 'Optiplex 780' THEN 2010 
					WHEN Model = 'Optiplex 790' THEN 2011
					WHEN Model = 'Optiplex 9010' THEN 2012
					WHEN Model = 'Optiplex 9020' THEN 2014
					WHEN Model = 'Optiplex 960' THEN 2009
					WHEN Model = 'Optiplex 980' THEN 2010
					WHEN Model = 'Optiplex 990' THEN 2011
					WHEN Model = 'Studio Slim 540s' THEN 2009
					WHEN Model = 'Studio 1537' THEN 2009
					WHEN Model = 'Venue 11 Pro 7130 MS' THEN 2014
					WHEN Model = 'Venue 11 Pro 7140' THEN 2014
					WHEN Model = 'Vostro 200' THEN 2009
					WHEN Model = 'Vostro 260s' THEN 2011
					WHEN Model = 'Dell System Vostro 3750' THEN 2011
					----------- HP -----------
					WHEN Model = 'HP 210 G1 Notebook' THEN 2015
					WHEN Model = 'HP Elite x2 1012 G1' OR Model = 'HP Elite x2 1012 G1 Tablet' OR Model = 'HP ELite x2 1012 G1 Tablet' THEN 2017
					WHEN Model = 'HP Elite x2 1012 G2' THEN 2017
					WHEN Model = 'HP EliteBook 840 G1' OR Model = 'HP EliteBook 840 G1 Corp' THEN 2015
					WHEN Model = 'HP EliteBook 840 G2' THEN 2016
					WHEN Model = 'HP EliteBook 840 G3' THEN 2017
					WHEN Model = 'HP EliteBook 840 G5' THEN 2018
					WHEN Model = 'HP EliteBook Folio 1040 G1' THEN 2015
					WHEN Model = 'HP EliteBook Folio 1040 G2' THEN 2016
					WHEN Model = 'HP EliteBook Folio 1040 G3' THEN 2017
					WHEN Model = 'EliteBook Folio 1040 G4' OR Model = 'HP EliteBook 1040 G4' THEN 2018
					WHEN Model = 'HP EliteBook Folio G1' THEN 2017
					WHEN Model = 'HP EliteDesk 700 G1 SFF' OR Model = 'EliteDesk 700 G1 SFF' OR Model = 'EliteDesk 700 G1' OR Model = 'HP EliteDesk 700 G1' THEN 2015
					WHEN Model = 'HP EliteDesk 800 G1 DM' OR Model = 'HP EliteDesk 800 G1 SFF' OR Model = 'EliteDesk 800 G1 SFF' THEN 2015
					WHEN Model = 'HP EliteDesk 800 G2 DM 35W' OR Model = 'HP EliteDesk 800 G2 SFF' OR Model = 'EliteDesk 800 G2 SFF' OR Model = 'ELITEDESK 800 G2' OR Model = 'HP EliteDesk 800 G2' OR Model = 'HP EliteDesk 800 G2  SFF' OR Model = 'HP EliteDesk 800 G2 SFF CORP' OR Model = ' HP EliteDesk 800 G2 SFF' OR Model = 'HP EliteDesk 800G2 SFF' OR Model =' HP EliteDesk 800  G2 SFF' THEN 2016
					WHEN Model = 'HP EliteDesk 800 G3 DM 35W' THEN 2017
					WHEN Model = 'HP EliteDesk 800 G4 SFF' THEN 2018
					WHEN Model = 'HP Pro x2 612 G1 Tablet' THEN 2015
					WHEN Model = 'HP ProBook 11 G1' THEN 2016
					WHEN Model = 'HP ProBook 11 G2' OR Model = 'HP ProBook 11 G2 EE' OR Model = 'HP ProBook 11 G2 HP ProBook 11 G2' THEN 2017
					WHEN Model = 'HP ProBook 430 G2' THEN 2015
					WHEN Model = 'HP ProBook 430 G3' OR Model = 'HP ProBook 430 G3 Notebook PC' OR Model = 'HP 430 G3' THEN 2016
					WHEN Model = 'HP ProBook 430 G4' THEN 2017
					WHEN Model = 'HP ProBook 430 G5' THEN 2018
					WHEN Model = 'HP ProBook 470 G5' THEN 2009
					WHEN Model = 'HP ProBook x360 11 EE G1' OR Model = 'HP ProBook x360 11 G1 EE' THEN 2018
					WHEN Model = 'HP ProBook 11 EE' THEN 2009
					WHEN Model = 'HP PROBOOK 11 EE G2' OR Model = 'HP PROBOOK 11 EE G2 NOTEBOOK PC' THEN 2017
					WHEN Model = 'HP ProDesk 600 G2 SFF' OR Model = 'HP ProDesk 600 G2 SMALL Form Factor PC' OR Model = 'HP ProDesk 600 G2SFF' OR Model = ' HP ProDesk 600 SFF G2' OR Model = 'HP Prodesk 600 SFF G2' OR Model = 'HP ProDesk 600G2' OR Model = 'HP ProDesk 600G2 SFF' OR Model = 'HP ProDesk 600 G2 SFF PC' OR Model = 'HP ProDesk 600 G2' OR Model = 'ProDesk 600 SFF G2' OR Model = 'HP  ProDesk 600 G2' OR Model = 'HP ProDesk 600' OR Model = 'HP Prodersk 600 G2 SFF' THEN 2016
					WHEN Model = 'HP ProDesk 600 G4 SFF' THEN 2018
					WHEN Model = 'HP Z230 SFF Workstation' THEN 2015
					WHEN Model = 'HP Z240 SFF Workstation' THEN 2016
					WHEN Model = 'HP Compaq Elite 8300 SFF' THEN 2012
					WHEN Model = '20-c002a' THEN 2016
					WHEN Model = 'HP ENVY 23-c021a' THEN 2012
					WHEN Model = 'HP ENVY 23-N100A' THEN 2014
					WHEN Model = 'HP Pavilion All-in-One (Touch)' THEN 2016
					WHEN Model = 'HP Pavilion Desktop - 500-410a' THEN 2014
					WHEN Model = 'HP TouchSmart 520-1010a Desktop PC' THEN 2011
					WHEN Model = 'Compaq 621 Notebook PC' THEN 2011
					WHEN Model = 'Compaq CQ45 Notebook PC' THEN 2008
					WHEN Model = 'Compaq Presario C700 Notebook PC' THEN 2008
					WHEN Model = 'HP 2000 Notebook PC' THEN 2013
					WHEN Model = 'HP 250 G5 Notebook PC' THEN 2017
					WHEN Model = 'HP Compaq 6000 Pro AiO Business PC' THEN 2010
					WHEN Model = 'HP Compaq 6000 Pro MT PC' THEN 2009
					WHEN Model = 'HP Compaq 6000 Pro SFF PC' THEN 2010
					WHEN Model = 'HP Compaq 8000 Elite SFF PC' THEN 2009
					WHEN Model = 'HP Compaq 8100 Elite SFF PC' THEN 2010
					WHEN Model = 'HP Compaq dc7600 Small Form Factor' THEN 2005
					WHEN Model = 'HP Compaq dc7700 Small Form Factor' OR Model = 'HP Compaq dc7700p Small Form Factor' THEN 2006	
					WHEN Model = 'HP Compaq dc7800 Small Form Factor' OR Model = 'HP Compaq dc7800p Small Form Factor' THEN 2008	
					WHEN Model = 'HP Compaq dc7900 Small Form Factor' OR Model = 'HP Compaq dc7900 Ultra-Slim Desktop' THEN 2008
					WHEN Model = 'HP Compaq dx2810 Microtower PC' THEN 2009
					WHEN Model = 'HP EliteBook Folio 9470m' THEN 2008
					WHEN Model = 'HP Laptop 15-bw0xx' THEN 2017
					WHEN Model = 'HP Compaq 6910p' THEN 2007
					WHEN Model = 'HP ProBook 6570b' THEN 2009
					WHEN Model = 'HP ProBook 4520s' THEN 2009			
					--------- LENOVO ---------
					WHEN Model = 'IdeaPad S10e' THEN 2009
					WHEN Model = 'IdeaPad V110' THEN 2016
					WHEN Model = 'IdeaPad Y330' THEN 2008
					WHEN Model = 'IdeaPad 300' THEN 2015
					WHEN Model = 'Yoga 310' THEN 2016
					WHEN Model = 'Miix 2 11' THEN 2014
					WHEN Model = 'Miix 2 510' THEN 2016
					WHEN Model = 'ThinkCentre M52' THEN 2006
					WHEN Model = 'ThinkCentre M55' THEN 2007
					WHEN Model = 'ThinkCentre M57' THEN 2008
 					WHEN Model = 'ThinkCentre M58' THEN 2010
					WHEN Model = 'ThinkCentre M58p' THEN 2008
					WHEN Model = 'ThinkCentre M700 SFF' THEN 2016
					WHEN Model = 'ThinkCentre M710s' THEN 2017
					WHEN Model = 'ThinkCentre M71e' THEN 2011
					WHEN Model = 'ThinkCentre M720'	THEN 2018			
					WHEN Model = 'ThinkCentre M53' THEN 2015
					WHEN Model = 'ThinkCentre M55e' THEN 2006
					WHEN Model = 'ThinkCentre M57p' THEN 2008
					WHEN Model = 'ThinkCentre M800 SFF' THEN 2016
					WHEN Model = 'ThinkCentre M81' THEN 2011 
					WHEN Model = 'ThinkCentre M82' THEN 2012
					WHEN Model = 'ThinkCentre M900 SFF' THEN 2016
					WHEN Model = 'ThinkCentre M90p' THEN 2010
					WHEN Model = 'ThinkCentre M90z - ALL-in-ONE' THEN 2010
					WHEN Model = 'ThinkCentre M910q TFF' THEN 2017
					WHEN Model = 'ThinkCentre M910s SFF' THEN 2017
					WHEN Model = 'ThinkCentre M91p' THEN 2011
					WHEN Model = 'ThinkCentre M92p' THEN 2012
					WHEN Model = 'ThinkCentre M92p TFF' THEN 2012
					WHEN Model = 'ThinkCentre M93p SFF(or TFF)' THEN 2014
					WHEN Model = 'ThinkCentre ALL-in-ONE' THEN 2010
					WHEN Model = 'ThinkPad 10 Tablet G1' THEN 2015
					WHEN Model = 'ThinkPad 10 Tablet G2' THEN 2016
					WHEN Model = 'ThinkPad 11e'  THEN 2014
					WHEN Model = 'ThinkPad 11e G2' THEN 2015
					WHEN Model = 'ThinkPad 11e G3' THEN 2016
					WHEN Model = 'ThinkPad 11e G4' THEN 2017
					WHEN Model = 'ThinkPad 11e G5' THEN 2017
					WHEN Model = 'ThinkPad Edge 11' THEN 2010
					WHEN Model = 'ThinkPad Edge E420' THEN 2011
					WHEN Model = 'ThinkPad Edge E530' THEN 2012
					WHEN Model = 'ThinkPad Edge E531' THEN 2013
					WHEN Model = 'ThinkPad E470' THEN 2016
					WHEN Model = 'ThinkPad Yoga L380' THEN 2018
					WHEN Model = 'ThinkPad L430' THEN 2012
					WHEN Model = 'ThinkPad L512' THEN 2010
					WHEN Model = 'ThinkPad L520' THEN 2011
					WHEN Model = 'ThinkPad Mini 10' THEN 2010
					WHEN Model = 'ThinkPad T410' THEN 2010
					WHEN Model = 'ThinkPad T410s' THEN 2010
					WHEN Model = 'ThinkPad T430' THEN 2012
					WHEN Model = 'ThinkPad T440' THEN 2014
					WHEN Model = 'ThinkPad T440p' THEN 2013
					WHEN Model = 'ThinkPad T440s' THEN 2013
					WHEN Model = 'ThinkPad T450' THEN 2015
					WHEN Model = 'ThinkPad T460' THEN 2016
					WHEN Model = 'ThinkPad T470' THEN 2017
					WHEN Model = 'ThinkPad T480' THEN 2018
					WHEN Model = 'ThinkPad T500' THEN 2008
					WHEN Model = 'ThinkPad T520' THEN 2012
					WHEN Model = 'ThinkPad T530' THEN 2013
					WHEN Model = 'ThinkPad X1 Tablet G3' THEN 2011
					WHEN Model = 'ThinkPad X1 Carbon G1' THEN 2012
					WHEN Model = 'ThinkPad X1 Carbon G2' THEN 2014
					WHEN Model = 'ThinkPad X1 Carbon G3' THEN 2015
					WHEN Model = 'ThinkPad X1 Carbon G4' THEN 2016
					WHEN Model = 'ThinkPad X1 Carbon G5' THEN 2017
					WHEN Model = 'ThinkPad X1 Yoga G1' THEN 2016
					WHEN Model = 'ThinkPad Yoga 11e G2' THEN 2013
					WHEN Model = 'ThinkPad x130e' THEN 2012
					WHEN Model = 'ThinkPad X131e' THEN 2013
					WHEN Model = 'ThinkPad X201 Tablet' THEN 2011
					WHEN Model = 'ThinkPad X220 Tablet' THEN 2011
					WHEN Model = 'ThinkPad Helix' THEN 2013
					WHEN Model = 'ThinkStation P300' THEN 2015
					WHEN Model = 'ThinkStation P310' THEN 2016
					WHEN Model = 'ThinkStation S20' THEN 2009
					WHEN Model = 'ThinkStation S30' THEN 2013
					WHEN Model = 'ThinkStation E20' THEN 2010
					WHEN Model = 'ThinkStation E30' THEN 2011
					WHEN Model = 'Lenovo G40-30' THEN 2014
					WHEN Model = 'Lenovo G50-45' THEN 2015
					WHEN Model = 'Lenovo Y70 Touch' THEN 2014
					--------- MICROSOFT ---------
					WHEN Model = 'Surface Pro' THEN 2017
					WHEN Model = 'Surface Pro 3' THEN 2014
					WHEN Model = 'Surface Pro 4' THEN 2014
					WHEN Model = 'Surface Go' THEN 2018
					WHEN Model = 'Surface Book' THEN 2017
					WHEN Model = 'Surface Book 2' THEN 2017
					WHEN Model = 'Surface 3' THEN 2015
	END AS RelYear,
	CASE -- Windows 7 
					---------- ACER ----------
					WHEN Model = 'Extensa 5630' OR Model = 'EX5630' THEN 'Y'
					WHEN Model = 'TravelMate 6592' THEN 'Y'
					WHEN Model = 'TravelMate 6595T' THEN 'Y'
					WHEN Model = 'TravelMate 8572' THEN 'Y'
					WHEN Model = 'TravelMate 8572T' THEN 'Y'
					WHEN Model = 'TravelMate 8572G' THEN 'Y'
					WHEN Model = 'TravelMate B113' THEN 'Y'
					WHEN Model = 'TravelMate B115-M' THEN 'Y'
					WHEN Model = 'TravelMate B117-M' THEN 'Y'
					WHEN Model = 'TravelMate P653-V' THEN 'Y'
					WHEN Model = 'Veriton S6630G' OR Model = 'Veriton S6630' OR Model = 'VeritonS6630G' OR Model = 'VS6630G' OR Model = 'S6630G' THEN 'Y'
					WHEN Model = 'Veriton M6630G' THEN 'Y'
					WHEN Model = 'Veriton L4630G' THEN 'Y'
					WHEN Model = 'Aspire ONE 533' THEN 'Y'
					WHEN Model = 'Aspire 5253' THEN 'Y'
					WHEN Model = 'Aspire 5542' THEN 'Y'
					WHEN Model = 'Aspire 5733Z' THEN 'Y'
					WHEN Model = 'Aspire 5738' THEN 'Y'
					WHEN Model = 'Aspire 5750' THEN 'Y'
					WHEN Model = 'Aspire 8942G' THEN 'Y'
					WHEN Model = 'Aspire E1-571' THEN 'Y'
					WHEN Model = 'Aspire E5-553G' THEN 'N'
					WHEN Model = 'Aspire F5-573G' THEN 'Y'
					WHEN Model = 'Aspire V3-371' THEN 'Y'
					WHEN Model = 'Aspire V3-571G' THEN 'Y'
					WHEN Model = 'Aspire X3950' THEN 'Y'
					WHEN Model = 'Aspire X3990' THEN 'Y'
					WHEN Model = 'Aspire 7720' THEN 'Y'
					WHEN Model = 'Aspire ES1-522' THEN 'Y'
					WHEN Model = 'Aspire Z5710' THEN 'Y'
					WHEN Model = 'Veriton Z6611G' THEN 'Y'
					WHEN Model = 'Veriton L4620G' THEN 'Y'
					WHEN Model = 'Aspire ES1-523' THEN 'N'
					WHEN Model = 'Aspire 7720' THEN 'Y'
					WHEN Model = 'TravelMate8572TG' THEN 'Y'
					WHEN Model = 'TravelMate P653' THEN 'Y'
					WHEN Model = 'TRAVELMATE B115' THEN 'Y'
					WHEN Model = 'TravelMate 8571' THEN 'Y'
					WHEN Model = 'TravelMate 6595TG' THEN 'Y'
					WHEN Model = 'TravelMate 5742' THEN 'Y'
					WHEN Model = 'TravelMate 5730' THEN 'Y'
					WHEN Model = 'ICONIA W700' THEN 'Y'
					WHEN Model = 'TravelMate P455-MG' THEN 'N'
					---------- DELL ----------
					WHEN Model = 'Inspiron 13-5378' THEN 'N'
					WHEN Model = 'Inspiron 15-3565' THEN 'N'
					WHEN Model = 'Inspiron 5379' THEN 'N'
					WHEN Model = 'Inspiron 5567' THEN 'Y'
					WHEN Model = 'Inspiron 5559' THEN 'Y'
					WHEN Model = 'Inspiron 1545' THEN 'Y'
					WHEN Model = 'Inspiron 545s' THEN 'Y'
					WHEN Model = 'Latitude 3330' THEN 'Y'
					WHEN Model = 'Latitude E5440' THEN 'Y'
					WHEN Model = 'Latitude E5540' THEN 'Y'
					WHEN Model = 'Latitude E6220' THEN 'Y'
					WHEN Model = 'Latitude E6230' THEN 'Y'
					WHEN Model = 'Latitude E6320' THEN 'Y'
					WHEN Model = 'Latitude E6330' THEN 'Y'
					WHEN Model = 'Latitude E6420' THEN 'Y'
					WHEN Model = 'Latitude E6430' THEN 'Y'
					WHEN Model = 'Latitude E6440' THEN 'Y'
					WHEN Model = 'Latitude E6520' THEN 'Y'
					WHEN Model = 'Latitude E7470' THEN 'Y'
					WHEN Model = 'OptiPlex GX620' THEN 'Y'
					WHEN Model = 'Optiplex 3020' THEN 'Y'
					WHEN Model = 'Optiplex 330' THEN 'Y'
					WHEN Model = 'Optiplex 360' THEN 'Y'
					WHEN Model = 'Optiplex 7010' THEN 'Y'
					WHEN Model = 'Optiplex 745' THEN 'Y'
					WHEN Model = 'Optiplex 7040' THEN 'Y'
					WHEN Model = 'Optiplex 755' THEN 'Y'
					WHEN Model = 'Optiplex 760' THEN 'Y'
					WHEN Model = 'Optiplex 780' THEN 'Y' 
					WHEN Model = 'Optiplex 790' THEN 'Y'
					WHEN Model = 'Optiplex 9010' THEN 'Y'
					WHEN Model = 'Optiplex 9020' THEN 'Y'
					WHEN Model = 'Optiplex 960' THEN 'Y'
					WHEN Model = 'Optiplex 980' THEN 'Y'
					WHEN Model = 'Optiplex 990' THEN 'Y'
					WHEN Model = 'Studio Slim 540s' THEN 'Y'
					WHEN Model = 'Studio 1537' THEN 'Y'
					WHEN Model = 'Venue 11 Pro 7130 MS' THEN 'Y'
					WHEN Model = 'Venue 11 Pro 7140' THEN 'Y'
					WHEN Model = 'Vostro 200' THEN 'Y'
					WHEN Model = 'Vostro 260s' THEN 'Y'
					WHEN Model = 'Dell System Vostro 3750' THEN 'Y'
					----------- HP -----------
					WHEN Model = 'HP 210 G1 Notebook' THEN 'Y'
					WHEN Model = 'HP Elite x2 1012 G1' OR Model = 'HP Elite x2 1012 G1 Tablet' OR Model = 'HP ELite x2 1012 G1 Tablet' THEN 'N'
					WHEN Model = 'HP Elite x2 1012 G2' THEN 'N'
					WHEN Model = 'HP EliteBook 840 G1' OR Model = 'HP EliteBook 840 G1 Corp' THEN 'Y'
					WHEN Model = 'HP EliteBook 840 G2' THEN 'Y'
					WHEN Model = 'HP EliteBook 840 G3' THEN 'N'
					WHEN Model = 'HP EliteBook 840 G5' THEN 'N'
					WHEN Model = 'HP EliteBook Folio 1040 G1' THEN 'Y'
					WHEN Model = 'HP EliteBook Folio 1040 G2' THEN 'Y'
					WHEN Model = 'HP EliteBook Folio 1040 G3' THEN 'N'
					WHEN Model = 'EliteBook Folio 1040 G4' THEN 'N'
					WHEN Model = 'HP EliteBook Folio G1' THEN 'N'
					WHEN Model = 'HP EliteDesk 700 G1 SFF' OR Model = 'EliteDesk 700 G1 SFF' OR Model = 'EliteDesk 700 G1' OR Model = 'HP EliteDesk 700 G1' THEN 'Y'
					WHEN Model = 'HP EliteDesk 800 G1 DM' OR Model = 'HP EliteDesk 800 G1 SFF' THEN 'Y'
					WHEN Model = 'HP EliteDesk 800 G2 DM 35W' OR Model = 'HP EliteDesk 800 G2 SFF' OR Model = 'EliteDesk 800 G2 SFF' OR Model = 'ELITEDESK 800 G2' OR Model = 'HP EliteDesk 800 G2' OR Model = 'HP EliteDesk 800 G2  SFF' OR Model = 'HP EliteDesk 800 G2 SFF CORP' OR Model = ' HP EliteDesk 800 G2 SFF' OR Model = 'HP EliteDesk 800G2 SFF' OR Model =' HP EliteDesk 800  G2 SFF' THEN 'Y'
					WHEN Model = 'HP EliteDesk 800 G3 DM 35W' THEN 'Y'
					WHEN Model = 'HP EliteDesk 800 G4 SFF' THEN 'N'
					WHEN Model = 'HP Pro x2 612 G1 Tablet' THEN 'Y'
					WHEN Model = 'HP ProBook 11 G1' THEN 'Y'
					WHEN Model = 'HP ProBook 11 G2' OR Model = 'HP ProBook 11 G2 EE' OR Model = 'HP ProBook 11 G2 HP ProBook 11 G2' THEN 'N'
					WHEN Model = 'HP ProBook 430 G2' THEN 'Y'
					WHEN Model = 'HP ProBook 430 G3' OR Model = 'HP ProBook 430 G3 Notebook PC' OR Model = 'HP 430 G3' THEN 'Y'
					WHEN Model = 'HP ProBook 430 G4' THEN 'N'
					WHEN Model = 'HP ProBook 430 G5' THEN 'N'
					WHEN Model = 'HP ProBook 470 G5' THEN 'Y'
					WHEN Model = 'HP ProBook x360 11 EE G1' THEN 'N'
					WHEN Model =  'HP ProBook 11 EE' THEN 'Y'
					WHEN Model = 'HP PROBOOK 11 EE G2' OR Model = 'HP PROBOOK 11 EE G2 NOTEBOOK PC' THEN 'Y'
					WHEN Model = 'HP ProDesk 600 G2 SFF' OR Model = 'HP ProDesk 600 G2 SMALL Form Factor PC' OR Model = 'HP ProDesk 600 G2SFF' OR Model = ' HP ProDesk 600 SFF G2' OR Model = 'HP Prodesk 600 SFF G2' OR Model = 'HP ProDesk 600G2' OR Model = 'HP ProDesk 600G2 SFF' OR Model = 'HP ProDesk 600 G2 SFF PC' OR Model = 'HP ProDesk 600 G2' OR Model = 'ProDesk 600 SFF G2' OR Model = 'HP  ProDesk 600 G2' OR Model = 'HP ProDesk 600' OR Model = 'HP Prodersk 600 G2 SFF' THEN 'Y'
					WHEN Model = 'HP ProDesk 600 G4 SFF' THEN 'N'
					WHEN Model = 'HP Z230 SFF Workstation' THEN 'Y'
					WHEN Model = 'HP Z240 SFF Workstation' THEN 'Y'
					WHEN Model = '20-c002a' THEN 'Y'
					WHEN Model = 'HP ENVY 23-c021a' THEN 'Y'
					WHEN Model = 'HP ENVY 23-N100A' THEN 'Y'
					WHEN Model = 'HP Pavilion All-in-One (Touch)' THEN 'Y'
					WHEN Model = 'HP Pavilion Desktop - 500-410a' THEN 'Y'
					WHEN Model = 'HP TouchSmart 520-1010a Desktop PC' THEN 'Y'
					WHEN Model = 'Compaq 621 Notebook PC' THEN 'Y'
					WHEN Model = 'Compaq CQ45 Notebook PC' THEN 'Y'
					WHEN Model = 'Compaq Presario C700 Notebook PC' THEN 'Y'
					WHEN Model = 'HP 2000 Notebook PC' THEN 'Y'
					WHEN Model = 'HP 250 G5 Notebook PC' THEN 'Y'
					WHEN Model = 'HP Compaq 6000 Pro AiO Business PC' THEN 'Y'
					WHEN Model = 'HP Compaq 6000 Pro MT PC' THEN 'Y'
					WHEN Model = 'HP Compaq 6000 Pro SFF PC' THEN 'Y'
					WHEN Model = 'HP Compaq 8000 Elite SFF PC' THEN 'Y'
					WHEN Model = 'HP Compaq 8100 Elite SFF PC' THEN 'Y'
					WHEN Model = 'HP Compaq dc7600 Small Form Factor' THEN 'Y'
					WHEN Model = 'HP Compaq dc7700 Small Form Factor' OR Model = 'HP Compaq dc7700p Small Form Factor' THEN 'Y'
					WHEN Model = 'HP Compaq dc7800 Small Form Factor' OR Model = 'HP Compaq dc7800p Small Form Factor' THEN 'Y'
					WHEN Model = 'HP Compaq dc7900 Small Form Factor' OR Model = 'HP Compaq dc7900 Ultra-Slim Desktop' THEN 'Y'
					WHEN Model = 'HP Compaq dx2810 Microtower PC' THEN 'Y'
					WHEN Model = 'HP EliteBook Folio 9470m' THEN 'Y'
					WHEN Model = 'HP Laptop 15-bw0xx' THEN 'N'
					WHEN Model = 'HP Compaq 6910p' THEN 'Y'
					WHEN Model = 'HP ProBook 6570b' THEN 'Y'
					WHEN Model = 'HP ProBook 4520s' THEN 'Y'
					--------- LENOVO ---------
					WHEN Model = 'IdeaPad S10e' THEN 'Y'
					WHEN Model = 'IdeaPad V110' THEN 'Y'
					WHEN Model = 'IdeaPad Y330' THEN 'Y'
					WHEN Model = 'IdeaPad 300' THEN 'Y'
					WHEN Model = 'Yoga 310' THEN 'Y'
					WHEN Model = 'Miix 2 11' THEN 'Y'
					WHEN Model = 'Miix 2 510' THEN 'Y'
					WHEN Model = 'ThinkCentre M52' THEN 'Y'
					WHEN Model = 'ThinkCentre M55' THEN 'Y'
					WHEN Model = 'ThinkCentre M57' THEN 'Y'
 					WHEN Model = 'ThinkCentre M58' THEN 'Y'
					WHEN Model = 'ThinkCentre M58p' THEN 'Y'
					WHEN Model = 'ThinkCentre M700 SFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M710s' THEN 'N'
					WHEN Model = 'ThinkCentre M71e' THEN 'Y'
					WHEN Model = 'ThinkCentre M720'	THEN 'N'			
					WHEN Model = 'ThinkCentre M53' THEN 'Y'
					WHEN Model = 'ThinkCentre M55e' THEN 'Y'
					WHEN Model = 'ThinkCentre M57p' THEN 'Y'
					WHEN Model = 'ThinkCentre M800 SFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M81' THEN 'Y' 
					WHEN Model = 'ThinkCentre M82' THEN 'Y'
					WHEN Model = 'ThinkCentre M900 SFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M90p' THEN 'Y'
					WHEN Model = 'ThinkCentre M90z - ALL-in-ONE' THEN 'Y'
					WHEN Model = 'ThinkCentre M910q TFF' THEN 'N'
					WHEN Model = 'ThinkCentre M910s SFF' THEN 'N'
					WHEN Model = 'ThinkCentre M91p' THEN 'Y'
					WHEN Model = 'ThinkCentre M92p' THEN 'Y'
					WHEN Model = 'ThinkCentre M92p TFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M93p SFF(or TFF)' THEN 'Y'
					WHEN Model = 'ThinkCentre ALL-in-ONE' THEN 'Y'
					WHEN Model = 'ThinkPad 10 Tablet G1' THEN 'Y'
					WHEN Model = 'ThinkPad 10 Tablet G2' THEN 'Y'
					WHEN Model = 'ThinkPad 11e'  THEN 'Y'
					WHEN Model = 'ThinkPad 11e G2' THEN 'Y'
					WHEN Model = 'ThinkPad 11e G3' THEN 'Y'
					WHEN Model = 'ThinkPad 11e G4' THEN 'N'
					WHEN Model = 'ThinkPad 11e G5' THEN 'N'
					WHEN Model = 'ThinkPad Edge 11' THEN 'Y'
					WHEN Model = 'ThinkPad Edge E420' THEN 'Y'
					WHEN Model = 'ThinkPad Edge E530' THEN 'Y'
					WHEN Model = 'ThinkPad Edge E531' THEN 'Y'
					WHEN Model = 'ThinkPad E470' THEN 'Y'
					WHEN Model = 'ThinkPad Yoga L380' THEN 'N'
					WHEN Model = 'ThinkPad L430' THEN 'Y'
					WHEN Model = 'ThinkPad L512' THEN 'Y'
					WHEN Model = 'ThinkPad L520' THEN 'Y'
					WHEN Model = 'ThinkPad Mini 10' THEN 'Y'
					WHEN Model = 'ThinkPad T410' THEN 'Y'
					WHEN Model = 'ThinkPad T410s' THEN 'Y'
					WHEN Model = 'ThinkPad T430' THEN 'Y'
					WHEN Model = 'ThinkPad T440' THEN 'Y'
					WHEN Model = 'ThinkPad T440p' THEN 'Y'
					WHEN Model = 'ThinkPad T440s' THEN 'Y'
					WHEN Model = 'ThinkPad T450' THEN 'Y'
					WHEN Model = 'ThinkPad T460' THEN 'Y'
					WHEN Model = 'ThinkPad T470' THEN 'N'
					WHEN Model = 'ThinkPad T480' THEN 'N'
					WHEN Model = 'ThinkPad T500' THEN 'Y'
					WHEN Model = 'ThinkPad T520' THEN 'Y'
					WHEN Model = 'ThinkPad T530' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Tablet G3' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G1' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G2' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G3' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G4' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G5' THEN 'N'
					WHEN Model = 'ThinkPad X1 Yoga G1' THEN 'Y'
					WHEN Model = 'ThinkPad Yoga 11e G2' THEN 'Y'
					WHEN Model = 'ThinkPad x130e' THEN 'Y'
					WHEN Model = 'ThinkPad X131e' THEN 'Y'
					WHEN Model = 'ThinkPad X201 Tablet' THEN 'Y'
					WHEN Model = 'ThinkPad X220 Tablet' THEN 'Y'
					WHEN Model = 'ThinkPad Helix' THEN 'Y'
					WHEN Model = 'ThinkStation P300' THEN 'Y'
					WHEN Model = 'ThinkStation P310' THEN 'Y'
					WHEN Model = 'ThinkStation S20' THEN 'Y'
					WHEN Model = 'ThinkStation S30' THEN 'Y'
					WHEN Model = 'ThinkStation E20' THEN 'Y'
					WHEN Model = 'ThinkStation E30' THEN 'Y'
					WHEN Model = 'Lenovo G40-30' THEN 'Y'
					WHEN Model = 'Lenovo G50-45' THEN 'Y'
					WHEN Model = 'Lenovo Y70 Touch' THEN 'Y'
					--------- MICROSOFT ---------
					WHEN Model = 'Surface Pro' THEN 'N'
					WHEN Model = 'Surface Pro 3' THEN 'N'
					WHEN Model = 'Surface Pro 4' THEN 'N'
					WHEN Model = 'Surface Go' THEN 'N'
					WHEN Model = 'Surface Book' THEN 'N'
					WHEN Model = 'Surface Book 2' THEN 'N'
					WHEN Model = 'Surface 3' THEN 'N'
	END AS Win7,
CASE -- Windows 10
					---------- ACER ----------
					WHEN Model = 'Extensa 5630' OR Model = 'EX5630' THEN 'N'
					WHEN Model = 'TravelMate 6592' THEN 'N'
					WHEN Model = 'TravelMate 6595T' THEN 'Y'
					WHEN Model = 'TravelMate 8572' THEN 'N'
					WHEN Model = 'TravelMate 8572T' THEN 'N'
					WHEN Model = 'TravelMate 8572G' THEN 'N'
					WHEN Model = 'TravelMate B113' THEN 'Y'
					WHEN Model = 'TravelMate B115-M' THEN 'Y'
					WHEN Model = 'TravelMate B117-M' THEN 'Y'
					WHEN Model = 'TravelMate P653-V' THEN 'Y'
					WHEN Model = 'Veriton S6630G' OR Model = 'Veriton S6630' OR Model = 'VeritonS6630G' OR Model = 'VS6630G' OR Model = 'S6630G' THEN 'Y'
					WHEN Model = 'Veriton M6630G' THEN 'Y'
					WHEN Model = 'Veriton L4630G' THEN 'Y'
					WHEN Model = 'Aspire ONE 533' THEN 'N'
					WHEN Model = 'Aspire 5253' THEN 'N'
					WHEN Model = 'Aspire 5542' THEN 'N'
					WHEN Model = 'Aspire 5733Z' THEN 'Y'
					WHEN Model = 'Aspire 5738' THEN 'N'
					WHEN Model = 'Aspire 5750' THEN 'Y'
					WHEN Model = 'Aspire 8942G' THEN 'N'
					WHEN Model = 'Aspire E1-571' THEN 'Y'
					WHEN Model = 'Aspire E5-553G' THEN 'Y'
					WHEN Model = 'Aspire F5-573G' THEN 'Y'
					WHEN Model = 'Aspire V3-371' THEN 'Y'
					WHEN Model = 'Aspire V3-571G' THEN 'Y'
					WHEN Model = 'Aspire X3950' THEN 'N'
					WHEN Model = 'Aspire X3990' THEN 'N'
					WHEN Model = 'Aspire 7720' THEN 'N'
					WHEN Model = 'Aspire ES1-522' THEN 'Y'
					WHEN Model = 'Aspire Z5710' THEN 'N'
					WHEN Model = 'Veriton Z6611G' THEN 'N'
					WHEN Model = 'Veriton L4620G' THEN 'Y'
					WHEN Model = 'Aspire ES1-523' THEN 'Y'
					WHEN Model = 'Aspire 7720' THEN 'N'
					WHEN Model = 'TravelMate8572TG' THEN 'N'
					WHEN Model = 'TravelMate P653' THEN 'Y'
					WHEN Model = 'TRAVELMATE B115' THEN 'Y'
					WHEN Model = 'TravelMate 8571' THEN 'N'
					WHEN Model = 'TravelMate 6595TG' THEN 'Y'
					WHEN Model = 'TravelMate 5742' THEN 'N'
					WHEN Model = 'TravelMate 5730' THEN 'N'
					WHEN Model = 'ICONIA W700' THEN 'Y'
					WHEN Model = 'TravelMate P455-MG' THEN 'Y'
					---------- DELL ----------
					WHEN Model = 'Inspiron 13-5378' THEN 'Y'
					WHEN Model = 'Inspiron 15-3565' THEN 'Y'
					WHEN Model = 'Inspiron 5379' THEN 'Y'
					WHEN Model = 'Inspiron 5567' THEN 'Y'
					WHEN Model = 'Inspiron 5559' THEN 'Y'
					WHEN Model = 'Inspiron 1545' THEN 'N'
					WHEN Model = 'Inspiron 545s' THEN 'N'
					WHEN Model = 'Latitude 3330' THEN 'Y'
					WHEN Model = 'Latitude E5440' THEN 'Y'
					WHEN Model = 'Latitude E5540' THEN 'Y'
					WHEN Model = 'Latitude E6220' THEN 'N'
					WHEN Model = 'Latitude E6230' THEN 'Y'
					WHEN Model = 'Latitude E6320' THEN 'N'
					WHEN Model = 'Latitude E6330' THEN 'Y'
					WHEN Model = 'Latitude E6420' THEN 'N'
					WHEN Model = 'Latitude E6430' THEN 'Y'
					WHEN Model = 'Latitude E6440' THEN 'Y'
					WHEN Model = 'Latitude E6520' THEN 'N'
					WHEN Model = 'Latitude E7470' THEN 'Y'
					WHEN Model = 'OptiPlex GX620' THEN 'N'
					WHEN Model = 'Optiplex 3020' THEN 'Y'
					WHEN Model = 'Optiplex 330' THEN 'N'
					WHEN Model = 'Optiplex 360' THEN 'N'
					WHEN Model = 'Optiplex 7010' THEN 'Y'
					WHEN Model = 'Optiplex 745' THEN 'N'
					WHEN Model = 'Optiplex 7040' THEN 'Y'
					WHEN Model = 'Optiplex 755' THEN 'N'
					WHEN Model = 'Optiplex 760' THEN 'N'
					WHEN Model = 'Optiplex 780' THEN 'N' 
					WHEN Model = 'Optiplex 790' THEN 'N'
					WHEN Model = 'Optiplex 9010' THEN 'Y'
					WHEN Model = 'Optiplex 9020' THEN 'Y'
					WHEN Model = 'Optiplex 960' THEN 'N'
					WHEN Model = 'Optiplex 980' THEN 'N'
					WHEN Model = 'Optiplex 990' THEN 'N'
					WHEN Model = 'Studio Slim 540s' THEN 'N'
					WHEN Model = 'Studio 1537' THEN 'N'
					WHEN Model = 'Venue 11 Pro 7130 MS' THEN 'Y'
					WHEN Model = 'Venue 11 Pro 7140' THEN 'Y'
					WHEN Model = 'Vostro 200' THEN 'N'
					WHEN Model = 'Vostro 260s' THEN 'N'
					WHEN Model = 'Dell System Vostro 3750' THEN 'N'
					----------- HP -----------
					WHEN Model = 'HP 210 G1 Notebook' THEN 'Y'
					WHEN Model = 'HP Elite x2 1012 G1' OR Model = 'HP Elite x2 1012 G1 Tablet' OR Model = 'HP ELite x2 1012 G1 Tablet' THEN 'Y'
					WHEN Model = 'HP Elite x2 1012 G2' THEN 'Y'
					WHEN Model = 'HP EliteBook 840 G1' OR Model = 'HP EliteBook 840 G1 Corp' THEN 'Y'
					WHEN Model = 'HP EliteBook 840 G2' THEN 'Y'
					WHEN Model = 'HP EliteBook 840 G3' THEN 'Y'
					WHEN Model = 'HP EliteBook 840 G5' THEN 'Y'
					WHEN Model = 'HP EliteBook Folio 1040 G1' THEN 'Y'
					WHEN Model = 'HP EliteBook Folio 1040 G2' THEN 'Y'
					WHEN Model = 'HP EliteBook Folio 1040 G3' THEN 'Y'
					WHEN Model = 'EliteBook Folio 1040 G4' THEN 'Y'
					WHEN Model = 'HP EliteBook Folio G1' THEN 'Y'
					WHEN Model = 'HP EliteDesk 700 G1 SFF' OR Model = 'EliteDesk 700 G1 SFF' OR Model = 'EliteDesk 700 G1' OR Model = 'HP EliteDesk 700 G1' THEN 'Y'
					WHEN Model = 'HP EliteDesk 800 G1 DM' OR Model = 'HP EliteDesk 800 G1 SFF' THEN 'Y'
					WHEN Model = 'HP EliteDesk 800 G2 DM 35W' OR Model = 'HP EliteDesk 800 G2 SFF' OR Model = 'EliteDesk 800 G2 SFF' OR Model = 'ELITEDESK 800 G2' OR Model = 'HP EliteDesk 800 G2' OR Model = 'HP EliteDesk 800 G2  SFF' OR Model = 'HP EliteDesk 800 G2 SFF CORP' OR Model = ' HP EliteDesk 800 G2 SFF' OR Model = 'HP EliteDesk 800G2 SFF' OR Model =' HP EliteDesk 800  G2 SFF' THEN 'Y'
					WHEN Model = 'HP EliteDesk 800 G3 DM 35W' THEN 'Y'
					WHEN Model = 'HP EliteDesk 800 G4 SFF' THEN 'Y'
					WHEN Model = 'HP Pro x2 612 G1 Tablet' THEN 'Y'
					WHEN Model = 'HP ProBook 11 G1' THEN 'Y'
					WHEN Model = 'HP ProBook 11 G2' OR Model = 'HP ProBook 11 G2 EE' OR Model = 'HP ProBook 11 G2 HP ProBook 11 G2' THEN 'Y'
					WHEN Model = 'HP ProBook 430 G2' THEN 'Y'
					WHEN Model = 'HP ProBook 430 G3' OR Model = 'HP ProBook 430 G3 Notebook PC' OR Model = 'HP 430 G3' THEN 'Y'
					WHEN Model = 'HP ProBook 430 G4' THEN 'Y'
					WHEN Model = 'HP ProBook 430 G5' THEN 'Y'
					WHEN Model = 'HP ProBook 470 G5' THEN 'N'
					WHEN Model = 'HP ProBook x360 11 EE G1' OR Model = 'HP ProBook x360 11 G1 EE' THEN 'Y'
					WHEN Model = 'HP ProBook 11 EE' THEN 'N'
					WHEN Model = 'HP PROBOOK 11 EE G2' OR Model = 'HP PROBOOK 11 EE G2 NOTEBOOK PC' THEN 'Y'
					WHEN Model = 'HP ProDesk 600 G2 SFF' OR Model = 'HP ProDesk 600 G2 SMALL Form Factor PC' OR Model = 'HP ProDesk 600 G2SFF' OR Model = ' HP ProDesk 600 SFF G2' OR Model = 'HP Prodesk 600 SFF G2' OR Model = 'HP ProDesk 600G2' OR Model = 'HP ProDesk 600G2 SFF' OR Model = 'HP ProDesk 600 G2 SFF PC' OR Model = 'HP ProDesk 600 G2' OR Model = 'ProDesk 600 SFF G2' OR Model = 'HP  ProDesk 600 G2' OR Model = 'HP ProDesk 600' OR Model = 'HP Prodersk 600 G2 SFF' THEN 'Y'
					WHEN Model = 'HP ProDesk 600 G4 SFF' THEN 'Y'
					WHEN Model = 'HP Z230 SFF Workstation' THEN 'Y'
					WHEN Model = 'HP Z240 SFF Workstation' THEN 'Y'
					WHEN Model = '20-c002a' THEN 'Y'
					WHEN Model = 'HP ENVY 23-c021a' THEN 'Y'
					WHEN Model = 'HP ENVY 23-N100A' THEN 'Y'
					WHEN Model = 'HP Pavilion All-in-One (Touch)' THEN 'Y'
					WHEN Model = 'HP Pavilion Desktop - 500-410a' THEN 'Y'
					WHEN Model = 'HP TouchSmart 520-1010a Desktop PC' THEN 'N'
					WHEN Model = 'Compaq 621 Notebook PC' THEN 'N'
					WHEN Model = 'Compaq CQ45 Notebook PC' THEN 'N'
					WHEN Model = 'Compaq Presario C700 Notebook PC' THEN 'N'
					WHEN Model = 'HP 2000 Notebook PC' THEN 'Y'
					WHEN Model = 'HP 250 G5 Notebook PC' THEN 'Y'
					WHEN Model = 'HP Compaq 6000 Pro AiO Business PC' THEN 'N'
					WHEN Model = 'HP Compaq 6000 Pro MT PC' THEN 'N'
					WHEN Model = 'HP Compaq 6000 Pro SFF PC' THEN 'N'
					WHEN Model = 'HP Compaq 8000 Elite SFF PC' THEN 'N'
					WHEN Model = 'HP Compaq 8100 Elite SFF PC' THEN 'N'
					WHEN Model = 'HP Compaq dc7600 Small Form Factor' THEN 'N'
					WHEN Model = 'HP Compaq dc7700 Small Form Factor' OR Model = 'HP Compaq dc7700p Small Form Factor' THEN 'N'
					WHEN Model = 'HP Compaq dc7800 Small Form Factor' OR Model = 'HP Compaq dc7800p Small Form Factor' THEN 'N'
					WHEN Model = 'HP Compaq dc7900 Small Form Factor' OR Model = 'HP Compaq dc7900 Ultra-Slim Desktop' THEN 'N'
					WHEN Model = 'HP Compaq dx2810 Microtower PC' THEN 'N'
					WHEN Model = 'HP EliteBook Folio 9470m' THEN 'N'
					WHEN Model = 'HP Laptop 15-bw0xx' THEN 'Y'
					WHEN Model = 'HP Compaq 6910p' THEN 'N'
					WHEN Model = 'HP ProBook 6570b' THEN 'Y'
					WHEN Model = 'HP ProBook 4520s' THEN 'N'
					--------- LENOVO ---------
					WHEN Model = 'IdeaPad S10e' THEN 'N'
					WHEN Model = 'IdeaPad V110' THEN 'Y'
					WHEN Model = 'IdeaPad Y330' THEN 'N'
					WHEN Model = 'IdeaPad 300' THEN 'Y'
					WHEN Model = 'Yoga 310' THEN 'Y'
					WHEN Model = 'Miix 2 11' THEN 'Y'
					WHEN Model = 'Miix 2 510' THEN 'Y'
					WHEN Model = 'ThinkCentre M52' THEN 'N'
					WHEN Model = 'ThinkCentre M55' THEN 'N'
					WHEN Model = 'ThinkCentre M57' THEN 'N'
 					WHEN Model = 'ThinkCentre M58' THEN 'N'
					WHEN Model = 'ThinkCentre M58p' THEN 'N'
					WHEN Model = 'ThinkCentre M700 SFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M710s' THEN 'Y'
					WHEN Model = 'ThinkCentre M71e' THEN 'N'
					WHEN Model = 'ThinkCentre M720'	THEN 'Y'			
					WHEN Model = 'ThinkCentre M53' THEN 'Y'
					WHEN Model = 'ThinkCentre M55e' THEN 'N'
					WHEN Model = 'ThinkCentre M57p' THEN 'N'
					WHEN Model = 'ThinkCentre M800 SFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M81' THEN 'N' 
					WHEN Model = 'ThinkCentre M82' THEN 'Y'
					WHEN Model = 'ThinkCentre M900 SFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M90p' THEN 'N'
					WHEN Model = 'ThinkCentre M90z - ALL-in-ONE' THEN 'N'
					WHEN Model = 'ThinkCentre M910q TFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M910s SFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M91p' THEN 'N'
					WHEN Model = 'ThinkCentre M92p' THEN 'Y'
					WHEN Model = 'ThinkCentre M92p TFF' THEN 'Y'
					WHEN Model = 'ThinkCentre M93p SFF(or TFF)' THEN 'Y'
					WHEN Model = 'ThinkCentre ALL-in-ONE' THEN 'N'
					WHEN Model = 'ThinkPad 10 Tablet G1' THEN 'Y'
					WHEN Model = 'ThinkPad 10 Tablet G2' THEN 'Y'
					WHEN Model = 'ThinkPad 11e'  THEN 'Y'
					WHEN Model = 'ThinkPad 11e G2' THEN 'Y'
					WHEN Model = 'ThinkPad 11e G3' THEN 'Y'
					WHEN Model = 'ThinkPad 11e G4' THEN 'Y'
					WHEN Model = 'ThinkPad 11e G5' THEN 'Y'
					WHEN Model = 'ThinkPad Edge 11' THEN 'N'
					WHEN Model = 'ThinkPad Edge E420' THEN 'N'
					WHEN Model = 'ThinkPad Edge E530' THEN 'Y'
					WHEN Model = 'ThinkPad Edge E531' THEN 'Y'
					WHEN Model = 'ThinkPad E470' THEN 'Y'
					WHEN Model = 'ThinkPad Yoga L380' THEN 'Y'
					WHEN Model = 'ThinkPad L430' THEN 'Y'
					WHEN Model = 'ThinkPad L512' THEN 'N'
					WHEN Model = 'ThinkPad L520' THEN 'N'
					WHEN Model = 'ThinkPad Mini 10' THEN 'N'
					WHEN Model = 'ThinkPad T410' THEN 'N'
					WHEN Model = 'ThinkPad T410s' THEN 'N'
					WHEN Model = 'ThinkPad T430' THEN 'Y'
					WHEN Model = 'ThinkPad T440' THEN 'Y'
					WHEN Model = 'ThinkPad T440p' THEN 'Y'
					WHEN Model = 'ThinkPad T440s' THEN 'Y'
					WHEN Model = 'ThinkPad T450' THEN 'Y'
					WHEN Model = 'ThinkPad T460' THEN 'Y'
					WHEN Model = 'ThinkPad T470' THEN 'Y'
					WHEN Model = 'ThinkPad T480' THEN 'Y'
					WHEN Model = 'ThinkPad T500' THEN 'N'
					WHEN Model = 'ThinkPad T520' THEN 'Y'
					WHEN Model = 'ThinkPad T530' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Tablet G3' THEN 'N'
					WHEN Model = 'ThinkPad X1 Carbon G1' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G2' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G3' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G4' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Carbon G5' THEN 'Y'
					WHEN Model = 'ThinkPad X1 Yoga G1' THEN 'Y'
					WHEN Model = 'ThinkPad Yoga 11e G2' THEN 'N'
					WHEN Model = 'ThinkPad x130e' THEN 'Y'
					WHEN Model = 'ThinkPad X131e' THEN 'Y'
					WHEN Model = 'ThinkPad X201 Tablet' THEN 'N'
					WHEN Model = 'ThinkPad X220 Tablet' THEN 'N'
					WHEN Model = 'ThinkPad Helix' THEN 'Y'
					WHEN Model = 'ThinkStation P300' THEN 'Y'
					WHEN Model = 'ThinkStation P310' THEN 'Y'
					WHEN Model = 'ThinkStation S20' THEN 'N'
					WHEN Model = 'ThinkStation S30' THEN 'Y'
					WHEN Model = 'ThinkStation E20' THEN 'N'
					WHEN Model = 'ThinkStation E30' THEN 'N'
					WHEN Model = 'Lenovo G40-30' THEN 'Y'
					WHEN Model = 'Lenovo G50-45' THEN 'Y'
					WHEN Model = 'Lenovo Y70 Touch' THEN 'Y'
					--------- MICROSOFT ---------
					WHEN Model = 'Surface Pro' THEN 'Y'
					WHEN Model = 'Surface Pro 3' THEN 'Y'
					WHEN Model = 'Surface Pro 4' THEN 'Y'
					WHEN Model = 'Surface Go' THEN 'Y'
					WHEN Model = 'Surface Book' THEN 'Y'
					WHEN Model = 'Surface Book 2' THEN 'Y'
					WHEN Model = 'Surface 3' THEN 'Y'
	END AS Win10
FROM #TMP_Devices
--WHERE ((Vendor = 'HP' OR Vendor = 'Hewlett-Packard'))
ORDER BY School, Vendor, Model, RelYear, BuildNumber ASC
