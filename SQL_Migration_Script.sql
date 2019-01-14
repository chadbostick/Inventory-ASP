/******** DMA Schema Migration Deployment Script      Script Date: 1/14/2019 12:24:25 AM ********/
-- 14 object(s) with recommendations identified during assessment. Please review these objects before deploying.


/****** Object:  Table [dbo].[tblSuppliers]    Script Date: 1/14/2019 12:24:23 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblSuppliers]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblSuppliers](
	[SupplierID] [int] IDENTITY(1,1) NOT NULL,
	[SupplierName] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ContactName] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ContactTitle] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Address] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[City] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PostalCode] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[StateOrProvince] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Country] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PhoneNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FaxNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblSuppliers] PRIMARY KEY CLUSTERED 
(
	[SupplierID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spGetSupplierBySupplierId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSupplierBySupplierId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSupplierBySupplierId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSupplierBySupplierId]

@ErrorMessage	varchar(255)	OUTPUT,
@SupplierId	int

AS

IF( @SupplierId = 0 )
BEGIN
	SELECT '' AS SupplierId, '' AS SupplierName, '' AS ContactName, '' AS ContactTitle, '' AS Address, '' AS City, '' AS PostalCode, '' AS StateOrProvince,
	'' AS Country, '' AS PhoneNumber, '' AS FaxNumber
END
ELSE
BEGIN
	SELECT tblSuppliers.SupplierId, SupplierName, ContactName, ContactTitle, Address, City, PostalCode, StateOrProvince,
	Country, PhoneNumber, FaxNumber
	FROM tblSuppliers
	WHERE tblSuppliers.SupplierId = @SupplierId
END



GO
/****** Object:  Table [dbo].[tblLocations]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblLocations]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblLocations](
	[LocationId] [int] IDENTITY(1,1) NOT NULL,
	[Location] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteLocationByLocationId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteLocationByLocationId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteLocationByLocationId] AS' 
END
GO


ALTER PROCEDURE [dbo].[spDeleteLocationByLocationId]

@ErrorMessage		varchar(255)	OUTPUT,
@LocationId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblLocations WHERE LocationId = @LocationId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblLocations WHERE LocationId = @LocationId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0
GO
/****** Object:  StoredProcedure [dbo].[spGetSupplierList]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSupplierList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSupplierList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSupplierList]

AS

SELECT SupplierId, SupplierName FROM tblSuppliers ORDER BY SupplierName



GO
/****** Object:  Table [dbo].[tblCategories]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblCategories]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblCategories](
	[CategoryID] [int] IDENTITY(1,1) NOT NULL,
	[CategoryName] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblCategories] PRIMARY KEY CLUSTERED 
(
	[CategoryID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblProductOwners]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblProductOwners]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblProductOwners](
	[ProductOwnerID] [int] IDENTITY(1,1) NOT NULL,
	[ProductOwner] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblProductOwners] PRIMARY KEY CLUSTERED 
(
	[ProductOwnerID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblProducts]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblProducts.ProductDescription uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 5, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblProducts]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblProducts](
	[ProductID] [int] IDENTITY(1,1) NOT NULL,
	[PartNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ProductName] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ProductDescription] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ProductShortDescription] [varchar](3000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CategoryID] [int] NULL,
	[SerialNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[UnitPrice] [money] NULL,
	[ReorderLevel] [int] NULL,
	[LeadTime] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DrawingID] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ProductOwnerID] [int] NOT NULL,
 CONSTRAINT [PK_tblProducts] PRIMARY KEY CLUSTERED 
(
	[ProductID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblProducts_ProductOwnerID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblProducts] ADD  CONSTRAINT [DF_tblProducts_ProductOwnerID]  DEFAULT (1) FOR [ProductOwnerID]
END

GO
/****** Object:  Table [dbo].[tblShippingMethods]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblShippingMethods]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblShippingMethods](
	[ShippingMethodID] [int] IDENTITY(1,1) NOT NULL,
	[ShippingMethod] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblShippingMethods] PRIMARY KEY CLUSTERED 
(
	[ShippingMethodID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblPurchaseOrders]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblPurchaseOrders.InternalNotes uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 16, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblPurchaseOrders]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblPurchaseOrders](
	[PurchaseOrderID] [int] IDENTITY(1,1) NOT NULL,
	[PurchaseOrderNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PurchaseOrderDescription] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[SupplierID] [int] NULL,
	[EmployeeID] [int] NULL,
	[OrderDate] [smalldatetime] NULL,
	[DateRequired] [smalldatetime] NULL,
	[DatePromised] [smalldatetime] NULL,
	[ShipDate] [smalldatetime] NULL,
	[DateShippedToSupplier] [smalldatetime] NULL,
	[ShippingMethodID] [int] NULL,
	[FreightCharge] [money] NULL,
	[TrackingNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DateClosed] [smalldatetime] NULL,
	[InternalNotes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PODescription] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblPurchaseOrders] PRIMARY KEY CLUSTERED 
(
	[PurchaseOrderID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblInventoryTransactions]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblInventoryTransactions.TransactionDescription uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 6, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblInventoryTransactions]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblInventoryTransactions](
	[TransactionID] [int] IDENTITY(1,1) NOT NULL,
	[TransactionDate] [smalldatetime] NULL,
	[ProductID] [int] NULL,
	[PurchaseOrderID] [int] NULL,
	[TransactionDescription] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[UnitPrice] [money] NULL,
	[UnitsOrdered] [numeric](10, 2) NULL,
	[DateReceived] [smalldatetime] NULL,
	[UnitsReceived] [numeric](10, 2) NULL,
	[UnitsSold] [numeric](10, 2) NULL,
	[UnitsShrinkage] [numeric](10, 2) NULL,
 CONSTRAINT [PK_tblInventoryTransactions] PRIMARY KEY CLUSTERED 
(
	[TransactionID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeletePOTransactionByTransactionId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeletePOTransactionByTransactionId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeletePOTransactionByTransactionId] AS' 
END
GO


ALTER PROCEDURE [dbo].[spDeletePOTransactionByTransactionId]

@ErrorMessage	varchar(255)	OUTPUT,
@TransactionId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblInventoryTransactions WHERE TransactionId = @TransactionId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblInventoryTransactions WHERE TransactionId = @TransactionId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0

GO
/****** Object:  StoredProcedure [dbo].[spGetSuppliers]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSuppliers]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSuppliers] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSuppliers]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'SupplierName',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'SupplierName'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetSupplierss FROM (
SELECT TOP 100 PERCENT tblSuppliers.SupplierId, MAX( SupplierName ) AS SupplierName, MAX( ContactName ) AS ContactName,
MAX( ContactTitle ) AS ContactTitle, MAX( PhoneNumber ) AS PhoneNumber, MAX( FaxNumber ) AS FaxNumber
FROM tblSuppliers
WHERE
CASE @OrderBy
	WHEN 'SupplierName' THEN SupplierName
	WHEN 'ContactName' THEN ContactName
	WHEN 'ContactTitle' THEN ContactTitle
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblSuppliers.SupplierId
ORDER BY
CASE @OrderBy WHEN 'SupplierName' THEN MAX( SupplierName ) ELSE NULL END,
CASE @OrderBy WHEN 'ContactName' THEN MAX( ContactName ) ELSE NULL END,
CASE @OrderBy WHEN 'ContactTitle' THEN MAX( ContactTitle ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetSupplierss WHERE 
	CASE @OrderBy
		WHEN 'SupplierName' THEN SUBSTRING( SupplierName, 1, 1 )
		WHEN 'ContactName' THEN SUBSTRING( ContactName, 1, 1 )
		WHEN 'ContactTitle' THEN SUBSTRING( ContactTitle, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, SupplierId, SupplierName, ContactName, ContactTitle, PhoneNumber, FaxNumber
FROM #tblTempGetSupplierss
WHERE (
CASE @OrderBy
	WHEN 'SupplierName' THEN SUBSTRING( SupplierName, 1, 1 )
	WHEN 'ContactName' THEN SUBSTRING( ContactName, 1, 1 )
	WHEN 'ContactTitle' THEN SUBSTRING( ContactTitle, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetSupplierss



GO
/****** Object:  StoredProcedure [dbo].[spDeleteProductByProductId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteProductByProductId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteProductByProductId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteProductByProductId]

@ErrorMessage	varchar(255)	OUTPUT,
@ProductId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblProducts WHERE ProductId = @ProductId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblProducts WHERE ProductId = @ProductId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0


GO
/****** Object:  Table [dbo].[tblProjectTypes]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblProjectTypes]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblProjectTypes](
	[ProjectTypeId] [int] IDENTITY(1000,1) NOT NULL,
	[ProjectType] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblProjectTypes] PRIMARY KEY CLUSTERED 
(
	[ProjectTypeId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblSpindles]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblSpindles.BalancingRequirements uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 66, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblSpindles]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblSpindles](
	[SpindleId] [int] IDENTITY(1000,1) NOT NULL,
	[SpindleType] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[SpindleCategoryId] [int] NULL,
	[DrawingNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RPM] [numeric](10, 2) NULL,
	[Weight] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Vibration] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[VibrationRear] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantFlow] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GSE] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GSERear] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ShaftTemp] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantTempIncomingSet] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantTempActual] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FrontTemp] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RearTemp] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantPressureIncoming] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BreakInTime] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolingMethod] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Volts] [numeric](10, 2) NULL,
	[HP] [numeric](10, 2) NULL,
	[Amps] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Phase] [varchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Hz] [numeric](10, 2) NULL,
	[Thermistor] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Poles] [numeric](10, 2) NULL,
	[AmpDraw] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ConnectorCtrl] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ConnectorPower] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Converter] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolHolder] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PullPin] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EMDimension] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EjectionPath] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolOutPressure] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ReturnPressure] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DrawbarForce] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolChangeFunction] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ProximitySwitchFunction] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Lubrication] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Grease] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilMist] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilJet] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilGreaseType] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[IntervalDPM] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[MainPressure] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TubePressure] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[LubeNotes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Preload] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RadialPlay] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[AxialPlay] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFrontLocation] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront2] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront2Location] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRearLocation] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear2] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear2Location] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolContact] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolContactRear] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolGap] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolGapRear] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Other] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BalancingRequirements] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BearingInformation] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GeneralSpindleNotes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DEPRICATED_DrawForce] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DEPRICATED_Category] [varchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblSpindles] PRIMARY KEY CLUSTERED 
(
	[SpindleId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblCustomers]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblCustomers.Notes uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 10, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblCustomers]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblCustomers](
	[CustomerId] [int] IDENTITY(1000,1) NOT NULL,
	[Customer] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Address] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[City] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[State] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Country] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Zip] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DateEstablished] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Notes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[SalesRepId] [int] NULL,
	[MainContactId] [int] NULL,
 CONSTRAINT [PK_tblCustomers] PRIMARY KEY CLUSTERED 
(
	[CustomerId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblEmployees]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblEmployees]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblEmployees](
	[EmployeeID] [int] IDENTITY(1,1) NOT NULL,
	[FirstName] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[LastName] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Title] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Extension] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[WorkPhone] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EmailAddress] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblEmployees] PRIMARY KEY CLUSTERED 
(
	[EmployeeID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblProjects]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblProjects.ProjectDescription uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 7, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblProjects]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblProjects](
	[ProjectId] [int] IDENTITY(1000,1) NOT NULL,
	[ProjectName] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ProjectTypeId] [int] NULL,
	[SalesRepId] [int] NULL,
	[CustomerId] [int] NULL,
	[ProjectDescription] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CustomerPO] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[StartDate] [smalldatetime] NULL,
	[EstCompletionDate] [smalldatetime] NULL,
	[CompletionDate] [smalldatetime] NULL,
	[ProjectPriorityId] [int] NULL,
	[SpindleId] [int] NULL,
	[ProjectContactId] [int] NULL,
 CONSTRAINT [PK_tblProjects] PRIMARY KEY CLUSTERED 
(
	[ProjectId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblWorkOrders]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblWorkOrders.AdditionalInfo uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 56, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblWorkOrders]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblWorkOrders](
	[WorkOrderId] [int] IDENTITY(1000,1) NOT NULL,
	[WorkOrderNumber] [int] NOT NULL,
	[ProjectId] [int] NULL,
	[PromiseDate] [smalldatetime] NULL,
	[DateIn] [smalldatetime] NULL,
	[DateOut] [smalldatetime] NULL,
	[SerialNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[NewSpindle] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PONumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Labor] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Material] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Subcontract] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Cost] [numeric](10, 2) NULL,
	[Charge] [numeric](10, 2) NULL,
	[Date] [smalldatetime] NULL,
	[DateExp] [smalldatetime] NULL,
	[SalesRep] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ShippingMethodId] [int] NULL,
	[TrackingWaybill] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Location] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[LocationId] [int] NULL,
	[Priority] [varchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[WorkOrderPriorityId] [int] NOT NULL,
	[BoeingWorkOrderNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BoeingSpindleId] [int] NULL,
	[Parts] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Bearings] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Lube] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BalVelocity] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BalVelocityFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GSE] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GSEFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BreakIn] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BreakInFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RoomTemp] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RoomTempFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FrontTemp] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FrontTempFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RearTemp] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RearTempFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Cooling] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolingFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFrontFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Rear] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RearFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Other] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[IncomingInspection] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Remarks] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DateRec] [smalldatetime] NULL,
	[ExpectedDelDate] [smalldatetime] NULL,
	[Commission] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CommissionReceivedDate] [smalldatetime] NULL,
	[AdditionalInfo] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ActualDrawforce] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ActualDrawforceFinal] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[WorkOrderContactId] [int] NULL,
 CONSTRAINT [PK_tblWorkOrders] PRIMARY KEY CLUSTERED 
(
	[WorkOrderId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblWorkOrders_WorkOrderPriorityId]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblWorkOrders] ADD  CONSTRAINT [DF_tblWorkOrders_WorkOrderPriorityId]  DEFAULT (1000) FOR [WorkOrderPriorityId]
END

GO
/****** Object:  Table [dbo].[tblCalls]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblCalls.CallComments uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 8, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblCalls]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblCalls](
	[CallId] [int] IDENTITY(1000,1) NOT NULL,
	[CallDate] [datetime] NOT NULL,
	[CustomerId] [int] NOT NULL,
	[EmployeeId] [int] NOT NULL,
	[ProjectId] [int] NULL,
	[WorkOrderId] [int] NULL,
	[CallComments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblCalls] PRIMARY KEY CLUSTERED 
(
	[CallId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblCalls_CallDate]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblCalls] ADD  CONSTRAINT [DF_tblCalls_CallDate]  DEFAULT (getdate()) FOR [CallDate]
END

GO
/****** Object:  Table [dbo].[tblWorkOrderPOs]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblWorkOrderPOs]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblWorkOrderPOs](
	[WorkOrderId] [int] NOT NULL,
	[PurchaseOrderId] [int] NOT NULL,
 CONSTRAINT [PK_tblWorkOrderPOs] PRIMARY KEY CLUSTERED 
(
	[WorkOrderId] ASC,
	[PurchaseOrderId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  Table [dbo].[tblWorkOrderPriorities]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblWorkOrderPriorities]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblWorkOrderPriorities](
	[WorkOrderPriorityId] [int] IDENTITY(1000,1) NOT NULL,
	[WorkOrderPriority] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
END
GO
/****** Object:  Table [dbo].[tblCustomerContacts]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblCustomerContacts.Notes uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 12, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblCustomerContacts]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblCustomerContacts](
	[CustomerContactId] [int] IDENTITY(1,1) NOT NULL,
	[CustomerId] [int] NOT NULL,
	[Contact] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ContactTitle] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Department] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TelephoneNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Extension] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[MobileNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FaxNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EmailAddress] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Notes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
END
GO
/****** Object:  StoredProcedure [dbo].[spGetWorkOrderByWorkOrderId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetWorkOrderByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetWorkOrderByWorkOrderId] AS' 
END
GO






ALTER     PROCEDURE [dbo].[spGetWorkOrderByWorkOrderId]

@ErrorMessage	varchar(255)	OUTPUT,
@WorkOrderId	int,
@IsPrint	bit = 0,
@CustomerId	int	OUTPUT

AS

IF( @IsPrint = 0 )
BEGIN
	IF( @WorkOrderId = 0 )
	BEGIN
		SELECT  '' AS SpindleId, '' AS SpindleType, '' AS WorkOrderNumber, '' AS ProjectId, '' AS ProjectName, '' AS CustomerId, '' AS Customer, 
		'' AS PromiseDate, '' AS DateIn, '' AS DateOut, '' AS SerialNumber, '' AS NewSpindle, 
		'' AS PONumber,	0 AS Labor, 0 AS Material, 0 AS Subcontract, '' AS Cost, '' AS Charge, '' AS [Date], '' AS DateExp, '' AS SalesRep, 0 AS ShippingMethodId, '' AS TrackingWaybill, '' AS Location, 0 AS LocationId,
		0 AS WorkOrderPriorityId, '' AS WorkOrderPriority, '' AS Priority, '' AS BoeingWorkOrderNumber, 0 AS BoeingSpindleID, '' AS Parts, '' AS Bearings, '' AS Lube, '' AS BalVelocity, '' AS BalVelocityFinal, '' AS GSE, '' AS GSEFinal, '' AS BreakIn, '' AS BreakInFinal, '' AS RoomTemp, '' AS RoomTempFinal, 
		'' AS FrontTemp, '' AS FrontTempFinal, '' AS RearTemp, '' AS RearTempFinal, '' AS Cooling, '' AS CoolingFinal, '' AS RunoutFront, '' AS RunoutFrontFinal, '' AS Rear, '' AS RearFinal, '' AS Other, '' AS IncomingInspection, '' AS Comments, '' AS Remarks, 
		'' AS DateRec, '' AS ExpectedDelDate, '' AS Commission, '' AS CommissionReceivedDate, '' AS AdditionalInfo, '' AS ActualDrawforce, '' AS ActualDrawforceFinal, 0 AS TotalRepairCost, '' AS SalesRepName
	END
	ELSE
	BEGIN
		SELECT  tblSpindles.SpindleId AS SpindleId, SpindleType, WorkOrderNumber, tblWorkOrders.ProjectId, ProjectName, tblCustomers.CustomerId, tblCustomers.Customer,
		PromiseDate, DateIn, DateOut, SerialNumber, NewSpindle, PONumber,
		ISNULL( Labor, 0 ) AS Labor, ISNULL( Material, 0 ) AS Material, ISNULL( Subcontract, 0 ) AS Subcontract, Cost, Charge, [Date], DateExp, SalesRep, ShippingMethodId, TrackingWaybill, Location, LocationId, tblWorkOrders.WorkOrderPriorityId, WorkOrderPriority, BoeingWorkOrderNumber, BoeingSpindleId,
		Parts, Bearings, Lube, BalVelocity, BalVelocityFinal, tblWorkOrders.GSE, GSEFinal, BreakIn, BreakInFinal, RoomTemp, RoomTempFinal, tblWorkOrders.FrontTemp, FrontTempFinal, tblWorkOrders.RearTemp, RearTempFinal, Cooling, CoolingFinal, tblWorkOrders.RunoutFront, RunoutFrontFinal, Rear, RearFinal, tblWorkOrders.Other, 
		IncomingInspection, Comments, Remarks, DateRec, ExpectedDelDate, Commission, CommissionReceivedDate, AdditionalInfo, ActualDrawforce, ActualDrawforceFinal,
		( ISNULL( CAST ( Labor AS money ), 0 ) + ISNULL( CAST( Material AS money ), 0 ) + ISNULL( CAST( Subcontract AS money ), 0 ) ) AS TotalRepairCost,
		WorkOrderContactId, tblEmployees.FirstName + ' ' + tblEmployees.LastName AS SalesRepName
		FROM tblWorkOrders
		LEFT JOIN tblWorkOrderPriorities ON tblWorkOrders.WorkOrderPriorityId = tblWorkOrderPriorities.WorkOrderPriorityId
		LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
		LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
		LEFT JOIN tblEmployees ON tblCustomers.SalesRepId = tblEmployees.EmployeeId
		LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
		WHERE tblWorkOrders.WorkOrderId = @WorkOrderId
		
		SELECT tblCalls.CallId, CallDate, ( LastName + ', ' + FirstName ) AS EmployeeName, ProjectId, WorkOrderId, CallComments
		FROM tblCalls
		LEFT JOIN tblEmployees ON tblCalls.EmployeeId = tblEmployees.EmployeeId
		WHERE WorkOrderId = @WorkOrderId
		ORDER BY CallDate DESC


		SELECT @CustomerId = tblProjects.CustomerId	FROM tblProjects
		LEFT JOIN tblWorkOrders ON tblProjects.ProjectId = tblWorkOrders.ProjectId
		LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
		WHERE tblWorkOrders.WorkOrderId = @WorkOrderId


		SELECT tblWorkOrderPOs.PurchaseOrderId AS PurchaseOrderId,
		tblWorkOrderPOs.WorkOrderId AS WorkOrderId,
		tblPurchaseOrders.PurchaseOrderNumber,
		tblPurchaseOrders.PurchaseOrderDescription,
		tblPurchaseOrders.OrderDate,
		tblPurchaseOrders.DatePromised,
		tblPurchaseOrders.DateRequired,
		tblPurchaseOrders.DateClosed,
		tblSuppliers.SupplierName
		FROM tblWorkOrderPOs
		LEFT JOIN tblPurchaseOrders ON tblWorkOrderPOs.PurchaseOrderId = tblPurchaseOrders.PurchaseOrderId
		LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierId
		WHERE tblWorkOrderPOs.WorkOrderId = @WorkOrderId
	END
END
ELSE
BEGIN
	SELECT WorkOrderNumber, DateIn, Customer, tblCustomerContacts.Contact, tblCustomerContacts.TelephoneNumber, tblCustomerContacts.FaxNumber, Address, SalesRep, tblWorkOrders.WorkOrderPriorityId, WorkOrderPriority, Priority, SpindleType, SerialNumber, AdditionalInfo, PONumber, tblWorkOrders.ProjectId AS ProjectId, tblEmployees.FirstName + ' ' + tblEmployees.LastName AS SalesRepName
	FROM tblWorkOrders
	LEFT JOIN tblWorkOrderPriorities ON tblWorkOrders.WorkOrderPriorityId = tblWorkOrderPriorities.WorkOrderPriorityId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
		LEFT JOIN tblEmployees ON tblCustomers.SalesRepId = tblEmployees.EmployeeId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	LEFT JOIN tblCustomerContacts ON tblWorkOrders.WorkOrderContactId = tblCustomerContacts.CustomerContactId
	WHERE tblWorkOrders.WorkOrderId = @WorkOrderId
END

GO
/****** Object:  StoredProcedure [dbo].[spDeleteProductCategoryByProductCategoryId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteProductCategoryByProductCategoryId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteProductCategoryByProductCategoryId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteProductCategoryByProductCategoryId]

@ErrorMessage		varchar(255)	OUTPUT,
@ProductCategoryId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblCategories WHERE CategoryId = @ProductCategoryId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblCategories WHERE CategoryId = @ProductCategoryId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0



GO
/****** Object:  StoredProcedure [dbo].[spGetWorkOrderList]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetWorkOrderList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetWorkOrderList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetWorkOrderList]

AS

SELECT WorkOrderId, WorkOrderNumber FROM tblWorkOrders ORDER BY WorkOrderNumber



GO
/****** Object:  StoredProcedure [dbo].[spDeleteProductOwnerByProductOwnerId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteProductOwnerByProductOwnerId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteProductOwnerByProductOwnerId] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spDeleteProductOwnerByProductOwnerId]

@ErrorMessage		varchar(255)	OUTPUT,
@ProductOwnerId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblProductOwners WHERE ProductOwnerId = @ProductOwnerId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblProductOwners WHERE ProductOwnerId = @ProductOwnerId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0




GO
/****** Object:  StoredProcedure [dbo].[spGetWorkOrderPriorities]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetWorkOrderPriorities]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetWorkOrderPriorities] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetWorkOrderPriorities]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'WorkOrderPriority',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'WorkOrderPriority'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetWorkOrderPriorities FROM (
SELECT TOP 100 PERCENT tblWorkOrderPriorities.WorkOrderPriorityId, MAX( WorkOrderPriority ) AS WorkOrderPriority
FROM tblWorkOrderPriorities
WHERE
CASE @OrderBy
	WHEN 'WorkOrderPriority' THEN WorkOrderPriority
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblWorkOrderPriorities.WorkOrderPriorityId
ORDER BY
CASE @OrderBy WHEN 'WorkOrderPriority' THEN MAX( WorkOrderPriority ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetWorkOrderPriorities WHERE 
	CASE @OrderBy
		WHEN 'WorkOrderPriority' THEN SUBSTRING( WorkOrderPriority, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, WorkOrderPriorityId, WorkOrderPriority
FROM #tblTempGetWorkOrderPriorities
WHERE (
CASE @OrderBy
	WHEN 'WorkOrderPriority' THEN SUBSTRING( WorkOrderPriority, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetWorkOrderPriorities


GO
/****** Object:  StoredProcedure [dbo].[spDeleteProjectByProjectId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteProjectByProjectId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteProjectByProjectId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteProjectByProjectId]

@ErrorMessage	varchar(255)	OUTPUT,
@ProjectId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblProjects WHERE ProjectId = @ProjectId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblProjects WHERE ProjectId = @ProjectId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0




GO
/****** Object:  StoredProcedure [dbo].[spGetWorkOrderPriorityByWorkOrderPriorityId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetWorkOrderPriorityByWorkOrderPriorityId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetWorkOrderPriorityByWorkOrderPriorityId] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetWorkOrderPriorityByWorkOrderPriorityId]

@ErrorMessage	varchar(255)	OUTPUT,
@WorkOrderPriorityId	int

AS

IF( @WorkOrderPriorityId = 0 )
BEGIN
	SELECT '' AS WorkOrderPriorityId, '' AS WorkOrderPriority
END
ELSE
BEGIN
	SELECT tblWorkOrderPriorities.WorkOrderPriorityId, WorkOrderPriority
	FROM tblWorkOrderPriorities 
	WHERE tblWorkOrderPriorities.WorkOrderPriorityId = @WorkOrderPriorityId
END


GO
/****** Object:  Table [dbo].[tblProjectPriorities]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblProjectPriorities]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblProjectPriorities](
	[ProjectPriorityId] [int] IDENTITY(1000,1) NOT NULL,
	[ProjectPriority] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblProjectPriorities] PRIMARY KEY CLUSTERED 
(
	[ProjectPriorityId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteProjectPriorityByProjectPriorityId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteProjectPriorityByProjectPriorityId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteProjectPriorityByProjectPriorityId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteProjectPriorityByProjectPriorityId]

@ErrorMessage		varchar(255)	OUTPUT,
@ProjectPriorityId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblProjectPriorities WHERE ProjectPriorityId = @ProjectPriorityId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblProjectPriorities WHERE ProjectPriorityId = @ProjectPriorityId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0



GO
/****** Object:  StoredProcedure [dbo].[spGetWorkOrderPriorityList]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetWorkOrderPriorityList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetWorkOrderPriorityList] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetWorkOrderPriorityList]

AS

SELECT WorkOrderPriorityId, WorkOrderPriority FROM tblWorkOrderPriorities ORDER BY WorkOrderPriorityId


GO
/****** Object:  StoredProcedure [dbo].[spDeleteProjectTypeByProjectTypeId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteProjectTypeByProjectTypeId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteProjectTypeByProjectTypeId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteProjectTypeByProjectTypeId]

@ErrorMessage		varchar(255)	OUTPUT,
@ProjectTypeId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblProjectTypes WHERE ProjectTypeId = @ProjectTypeId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblProjectTypes WHERE ProjectTypeId = @ProjectTypeId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0


GO
/****** Object:  StoredProcedure [dbo].[spGetWorkOrders]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetWorkOrders]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetWorkOrders] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetWorkOrders]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'WorkOrder' OUTPUT,
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(11)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'Customer'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

IF( @OrderBy = 'DateIn' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetWorkOrdersByDateIn FROM (
	SELECT TOP 100 PERCENT MAX( WorkOrderId ) AS WorkOrderId, MAX( WorkOrderNumber ) AS WorkOrderNumber, MAX( Customer ) AS Customer, MAX( DateIn ) AS DateIn, 
	MAX( DateOut ) AS DateOut, MAX( SpindleType ) AS SpindleType, MAX( NewSpindle ) AS NewSpindle, MAX( WorkOrderPriority ) AS Priority, MAX( PONumber ) AS PONumber
	FROM tblWorkOrders
	LEFT JOIN tblWorkOrderPriorities ON tblWorkOrders.WorkOrderPriorityId = tblWorkOrderPriorities.WorkOrderPriorityId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	GROUP BY tblWorkOrders.WorkOrderId
	ORDER BY DateIn
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetWorkOrdersByDateIn
		WHERE DateIn BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, WorkOrderId, WorkOrderNumber, Customer, DateIn, DateOut, SpindleType, NewSpindle, Priority, PONumber
		FROM #tblTempGetWorkOrdersByDateIn
		WHERE ( DateIn BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SELECT RowNumber, WorkOrderId, WorkOrderNumber, Customer, DateIn, DateOut, SpindleType, NewSpindle, Priority, PONumber
		FROM #tblTempGetWorkOrdersByDateIn
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetWorkOrdersByDateIn
END
ELSE IF( @OrderBy = 'DateOut' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetWorkOrdersByDateOut FROM (
	SELECT TOP 100 PERCENT MAX( WorkOrderId ) AS WorkOrderId, MAX( WorkOrderNumber ) AS WorkOrderNumber, MAX( Customer ) AS Customer, MAX( DateIn ) AS DateIn, 
	MAX( DateOut ) AS DateOut, MAX( SpindleType ) AS SpindleType, MAX( NewSpindle ) AS NewSpindle, MAX( WorkOrderPriority ) AS Priority, MAX( PONumber ) AS PONumber
	FROM tblWorkOrders
	LEFT JOIN tblWorkOrderPriorities ON tblWorkOrders.WorkOrderPriorityId = tblWorkOrderPriorities.WorkOrderPriorityId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	GROUP BY tblWorkOrders.WorkOrderId
	ORDER BY DateOut
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetWorkOrdersByDateOut
		WHERE DateOut BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, WorkOrderId, WorkOrderNumber, Customer, DateIn, DateOut, SpindleType, NewSpindle, Priority, PONumber
		FROM #tblTempGetWorkOrdersByDateOut
		WHERE ( DateOut BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	END
	ELSE
	BEGIN
		SELECT RowNumber, WorkOrderId, WorkOrderNumber, Customer, DateIn, DateOut, SpindleType, NewSpindle, Priority, PONumber
		FROM #tblTempGetWorkOrdersByDateOut
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetWorkOrdersByDateOut
END
ELSE IF( @OrderBy = 'WorkOrder' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetWorkOrdersById FROM (
	SELECT TOP 100 PERCENT MAX( WorkOrderId ) AS WorkOrderId, MAX( WorkOrderNumber ) AS WorkOrderNumber, MAX( Customer ) AS Customer, MAX( DateIn ) AS DateIn, 
	MAX( DateOut ) AS DateOut, MAX( SpindleType ) AS SpindleType, MAX( NewSpindle ) AS NewSpindle, MAX( WorkOrderPriority ) AS Priority, MAX( PONumber ) AS PONumber
	FROM tblWorkOrders
	LEFT JOIN tblWorkOrderPriorities ON tblWorkOrders.WorkOrderPriorityId = tblWorkOrderPriorities.WorkOrderPriorityId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE WorkOrderNumber LIKE CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	GROUP BY tblWorkOrders.WorkOrderId
	ORDER BY WorkOrderNumber ) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetWorkOrdersById
		WHERE WorkOrderNumber BETWEEN CONVERT( int, @JumpTo ) AND 9999999 ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SET @JumpTo = 1
	END
	
	SELECT RowNumber, WorkOrderId, WorkOrderNumber, Customer, DateIn, DateOut, SpindleType, NewSpindle, Priority, PONumber
	FROM #tblTempGetWorkOrdersById
	WHERE ( WorkOrderNumber BETWEEN CONVERT( int, @JumpTo ) AND 9999999 )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	DROP TABLE #tblTempGetWorkOrdersById
END
ELSE
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetWorkOrders FROM (
	SELECT TOP 100 PERCENT MAX( WorkOrderId ) AS WorkOrderId, MAX( WorkOrderNumber ) AS WorkOrderNumber, MAX( Customer ) AS Customer, MAX( DateIn ) AS DateIn, 
	MAX( DateOut ) AS DateOut, MAX( SpindleType ) AS SpindleType, MAX( NewSpindle ) AS NewSpindle, MAX( WorkOrderPriority ) AS Priority, MAX( PONumber ) AS PONumber
	FROM tblWorkOrders
	LEFT JOIN tblWorkOrderPriorities ON tblWorkOrders.WorkOrderPriorityId = tblWorkOrderPriorities.WorkOrderPriorityId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE
	CASE @OrderBy
		WHEN 'Customer' THEN Customer
		WHEN 'SpindleType' THEN SpindleType
		WHEN 'Priority' THEN Priority
		WHEN 'NewSpindle' THEN NewSpindle
		WHEN 'PONumber' THEN PONumber
	END
	LIKE
	CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	GROUP BY tblWorkOrders.WorkOrderId
	ORDER BY
	CASE @OrderBy WHEN 'Customer' THEN MAX( Customer ) ELSE NULL END,
	CASE @OrderBy WHEN 'SpindleType' THEN MAX( SpindleType ) ELSE NULL END,
	CASE @OrderBy WHEN 'Priority' THEN MAX( Priority ) ELSE NULL END,
	CASE @OrderBy WHEN 'NewSpindle' THEN MAX( NewSpindle ) ELSE NULL END,
	CASE @OrderBy WHEN 'PONumber' THEN MAX( PONumber ) ELSE NULL END
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetWorkOrders WHERE 
		CASE @OrderBy
			WHEN 'Customer' THEN SUBSTRING( Customer, 1, 1 )
			WHEN 'SpindleType' THEN SUBSTRING( SpindleType, 1, 1 )
			WHEN 'Priority' THEN SUBSTRING( Priority, 1, 1 )
			WHEN 'NewSpindle' THEN SUBSTRING( NewSpindle, 1, 1 )
			WHEN 'PONumber' THEN SUBSTRING( PONumber, 1, 1 )
		END
		LIKE @JumpTo ORDER BY RowNumber

		SELECT RowNumber, WorkOrderId, WorkOrderNumber, Customer, DateIn, DateOut, SpindleType, NewSpindle, Priority, PONumber
		FROM #tblTempGetWorkOrders
		WHERE (
		CASE @OrderBy
			WHEN 'Customer' THEN SUBSTRING( Customer, 1, 1 )
			WHEN 'SpindleType' THEN SUBSTRING( SpindleType, 1, 1 )
			WHEN 'Priority' THEN SUBSTRING( Priority, 1, 1 )
			WHEN 'NewSpindle' THEN SUBSTRING( NewSpindle, 1, 1 )
			WHEN 'PONumber' THEN SUBSTRING( PONumber, 1, 1 )
		END
		BETWEEN @JumpTo AND 'Z' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SELECT RowNumber, WorkOrderId, WorkOrderNumber, Customer, DateIn, DateOut, SpindleType, NewSpindle, Priority, PONumber
		FROM #tblTempGetWorkOrders
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	
	DROP TABLE #tblTempGetWorkOrders
END

GO
/****** Object:  StoredProcedure [dbo].[spDeletePurchaseOrderByPurchaseOrderId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeletePurchaseOrderByPurchaseOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeletePurchaseOrderByPurchaseOrderId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeletePurchaseOrderByPurchaseOrderId]

@ErrorMessage	varchar(255)	OUTPUT,
@PurchaseOrderId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblPurchaseOrders WHERE PurchaseOrderId = @PurchaseOrderId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblPurchaseOrders WHERE PurchaseOrderId = @PurchaseOrderId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0


GO
/****** Object:  StoredProcedure [dbo].[spPrintPurchaseOrderCenterline]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spPrintPurchaseOrderCenterline]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spPrintPurchaseOrderCenterline] AS' 
END
GO






ALTER   PROCEDURE [dbo].[spPrintPurchaseOrderCenterline]

@PurchaseOrderId	int

AS

SELECT SupplierName, Address, City, PostalCode, StateOrProvince, Country, ContactName, PhoneNumber, FaxNumber, CONVERT( varchar(10), OrderDate, 101 ) AS OrderDate,
CONVERT( varchar(10), DateRequired, 101 ) AS DateRequired, tblPurchaseOrders.PurchaseOrderId AS PurchaseOrderId, PurchaseOrderNumber, ( LastName + ', ' + FirstName ) AS BuyerName,
PurchaseOrderDescription, CONVERT( varchar(10), DatePromised, 101 ) AS DatePromised, ShippingMethod
FROM tblPurchaseOrders
LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierID
LEFT JOIN tblEmployees ON tblPurchaseOrders.EmployeeId = tblEmployees.EmployeeId
LEFT JOIN tblShippingMethods ON tblPurchaseOrders.ShippingMethodId = tblShippingMethods.ShippingMethodId
WHERE tblPurchaseOrders.PurchaseOrderId = @PurchaseOrderId

SELECT tblInventoryTransactions.ProductId, ProductName, PartNumber, ProductDescription, ISNULL( UnitsOrdered, 0 ) AS UnitsOrdered, ISNULL( tblInventoryTransactions.UnitPrice, 0 ) AS UnitPrice,
( ISNULL( UnitsOrdered, 0 ) * ISNULL( tblInventoryTransactions.UnitPrice, 0 ) ) AS SubTotal
FROM tblInventoryTransactions
LEFT JOIN tblProducts ON tblInventoryTransactions.ProductId = tblProducts.ProductId
WHERE tblInventoryTransactions.PurchaseOrderId = @PurchaseOrderId


select sum (( ISNULL( UnitsOrdered, 0 ) * ISNULL( tblInventoryTransactions.UnitPrice, 0 ) )) AS OrderTotal
FROM tblInventoryTransactions
LEFT JOIN tblProducts ON tblInventoryTransactions.ProductId = tblProducts.ProductId
WHERE tblInventoryTransactions.PurchaseOrderId = @PurchaseOrderId


GO
/****** Object:  Table [dbo].[tblQCTemps]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblQCTemps.QCNotes uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 10, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblQCTemps]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblQCTemps](
	[QCTempId] [int] IDENTITY(1,1) NOT NULL,
	[QCTempLogId] [int] NOT NULL,
	[QCTime] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QCSpeed] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QCFront] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QCRear] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QCShaft] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QCTempLocation] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QCNotes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblQCTemps] PRIMARY KEY CLUSTERED 
(
	[QCTempId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteQCTempByQCTempId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteQCTempByQCTempId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteQCTempByQCTempId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteQCTempByQCTempId]

@ErrorMessage	varchar(255)	OUTPUT,
@QCTempId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblQCTemps WHERE QCTempId = @QCTempId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblQCTemps WHERE QCTempId = @QCTempId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0
GO
/****** Object:  StoredProcedure [dbo].[spPrintSubWorkCenterline]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spPrintSubWorkCenterline]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spPrintSubWorkCenterline] AS' 
END
GO





ALTER  PROCEDURE [dbo].[spPrintSubWorkCenterline]

@PurchaseOrderId	int

AS

SELECT SupplierName, Address, City, PostalCode, StateOrProvince, Country, ContactName, PhoneNumber, FaxNumber, CONVERT( varchar(10), OrderDate, 101 ) AS OrderDate,
CONVERT( varchar(10), DateRequired, 101 ) AS DateRequired, PurchaseOrderId, PurchaseOrderNumber, ( LastName + ', ' + FirstName ) AS BuyerName,
PurchaseOrderDescription, CONVERT( varchar(10), DatePromised, 101 ) AS DatePromised, ShippingMethod
FROM tblPurchaseOrders
LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierID
LEFT JOIN tblEmployees ON tblPurchaseOrders.EmployeeId = tblEmployees.EmployeeId
LEFT JOIN tblShippingMethods ON tblPurchaseOrders.ShippingMethodId = tblShippingMethods.ShippingMethodId
WHERE tblPurchaseOrders.PurchaseOrderId = @PurchaseOrderId

SELECT tblInventoryTransactions.ProductId, ProductName, TransactionDescription, ISNULL( UnitsOrdered, 0 ) AS UnitsOrdered, ISNULL( tblInventoryTransactions.UnitPrice, 0 ) AS UnitPrice,
( ISNULL( UnitsOrdered, 0 ) * ISNULL( tblInventoryTransactions.UnitPrice, 0 ) ) AS SubTotal
FROM tblInventoryTransactions
LEFT JOIN tblProducts ON tblInventoryTransactions.ProductId = tblProducts.ProductId
WHERE tblInventoryTransactions.PurchaseOrderId = @PurchaseOrderId


select sum (( ISNULL( UnitsOrdered, 0 ) * ISNULL( tblInventoryTransactions.UnitPrice, 0 ) )) AS OrderTotal
FROM tblInventoryTransactions
LEFT JOIN tblProducts ON tblInventoryTransactions.ProductId = tblProducts.ProductId
WHERE tblInventoryTransactions.PurchaseOrderId = @PurchaseOrderId






GO
/****** Object:  Table [dbo].[tblQCTempLogs]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblQCTempLogs]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblQCTempLogs](
	[QCTempLogId] [int] IDENTITY(1,1) NOT NULL,
	[WorkOrderId] [int] NOT NULL,
	[QCDate] [datetime] NULL,
	[QCMaxSpeed] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QCTotalRunTime] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblQCTempLogs] PRIMARY KEY CLUSTERED 
(
	[QCTempLogId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteQCTempLogByQCTempLogId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteQCTempLogByQCTempLogId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteQCTempLogByQCTempLogId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteQCTempLogByQCTempLogId]

@ErrorMessage	varchar(255)	OUTPUT,
@QCTempLogId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblQCTempLogs WHERE QCTempLogId = @QCTempLogId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblQCTempLogs WHERE QCTempLogId = @QCTempLogId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0
GO
/****** Object:  Table [dbo].[tblQCInspections]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblQCInspections]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblQCInspections](
	[QCInspectionId] [int] IDENTITY(1,1) NOT NULL,
	[WorkOrderId] [int] NOT NULL,
	[RPM_In] [numeric](10, 2) NULL,
	[RPM_Final] [numeric](10, 2) NULL,
	[Weight_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Weight_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Vibration_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Vibration_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[VibrationRear_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[VibrationRear_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantFlow_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantFlow_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GSE_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GSE_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GSERear_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GSERear_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ShaftTemp_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ShaftTemp_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantTempIncomingSet_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantTempIncomingSet_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantTempActual_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantTempActual_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FrontTemp_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FrontTemp_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RearTemp_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RearTemp_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantPressureIncoming_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolantPressureIncoming_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BreakInTime_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BreakInTime_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolingMethod_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CoolingMethod_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Volts_In] [numeric](10, 2) NULL,
	[Volts_Final] [numeric](10, 2) NULL,
	[HP_In] [numeric](10, 2) NULL,
	[HP_Final] [numeric](10, 2) NULL,
	[Amps_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Amps_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Phase_In] [varchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Phase_Final] [varchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Hz_In] [numeric](10, 2) NULL,
	[Hz_Final] [numeric](10, 2) NULL,
	[Thermistor_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Thermistor_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Poles_In] [numeric](10, 2) NULL,
	[Poles_Final] [numeric](10, 2) NULL,
	[AmpDraw_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[AmpDraw_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ConnectorCtrl_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ConnectorCtrl_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ConnectorPower_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ConnectorPower_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Converter_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Converter_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolHolder_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolHolder_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PullPin_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PullPin_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EMDimension_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EMDimension_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EjectionPath_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EjectionPath_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolOutPressure_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolOutPressure_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ReturnPressure_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ReturnPressure_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DrawbarForce_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DrawbarForce_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolChangeFunction_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolChangeFunction_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ProximitySwitchFunction_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ProximitySwitchFunction_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Lubrication_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Lubrication_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Grease_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Grease_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilMist_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilMist_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilJet_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilJet_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilGreaseType_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OilGreaseType_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[IntervalDPM_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[IntervalDPM_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[MainPressure_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[MainPressure_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TubePressure_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TubePressure_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Preload_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Preload_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RadialPlay_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RadialPlay_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[AxialPlay_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[AxialPlay_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFrontLocation_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFrontLocation_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront2_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront2_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront2Location_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutFront2Location_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRearLocation_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRearLocation_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear2_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear2_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear2Location_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[RunoutRear2Location_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolContact_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolContact_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolContactRear_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolContactRear_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolGap_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolGap_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolGapRear_In] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ToolGapRear_Final] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
END
GO
/****** Object:  StoredProcedure [dbo].[spReportFinalInspection]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spReportFinalInspection]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spReportFinalInspection] AS' 
END
GO







ALTER   PROCEDURE [dbo].[spReportFinalInspection]

@WorkOrderId	int

AS

SELECT Customer, SerialNumber, PONumber, IncomingInspection, Bearings, Parts,

tblQCInspections.OilGreaseType_Final,
tblQCInspections.IntervalDPM_Final,
tblQCInspections.MainPressure_Final,
tblQCInspections.TubePressure_Final,
tblQCInspections.Vibration_Final,
tblQCInspections.GSE_Final,
tblQCInspections.BreakInTime_Final,
tblQCInspections.DrawbarForce_Final,
tblQCInspections.CoolantFlow_Final,
tblQCInspections.FrontTemp_Final,
tblQCInspections.RearTemp_Final,
tblQCInspections.RunoutFront_Final,
tblQCInspections.RunoutRear_Final,
tblQCInspections.RunoutFrontLocation_Final,
tblQCInspections.RunoutRearLocation_Final,
Remarks,
DateOut, WorkOrderNumber, SpindleType
FROM tblWorkOrders
LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
LEFT JOIN tblQCInspections ON tblWorkOrders.WorkOrderId = tblQCInspections.WorkOrderId
WHERE tblWorkOrders.WorkOrderId = @WorkOrderId




GO
/****** Object:  Table [dbo].[tblQuotes]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblQuotes.MscWorkNeeded uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 21, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblQuotes]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblQuotes](
	[QuoteId] [int] IDENTITY(1000,1) NOT NULL,
	[OriginalQuoteId] [int] NULL,
	[WorkOrderId] [int] NULL,
	[WorkOrderNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[SerialNumber] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DisassemblyEvaluation] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[HoursDisassembly] [numeric](10, 2) NULL,
	[CleanAndInspect] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[HoursCleanAndInspect] [numeric](10, 2) NULL,
	[InhouseGrinding] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GrindingHours] [numeric](10, 2) NULL,
	[Balancing] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BalancingHours] [numeric](10, 2) NULL,
	[ElectricalWork] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ElectricalHours] [numeric](10, 2) NULL,
	[GreaseBearings] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[GreaseHours] [numeric](10, 2) NULL,
	[AssemblyAndTest] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[AssemblyAndTestHours] [numeric](10, 2) NULL,
	[MscWorkNeeded] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[MscWorkHours] [numeric](10, 2) NULL,
	[FreightChargeParts] [numeric](10, 2) NULL,
	[FreightChargeSubWork] [numeric](10, 2) NULL,
	[FreightLBS] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FreightChargeSub] [numeric](10, 2) NULL,
	[FreightChargeSub1] [numeric](10, 2) NULL,
	[ExpDeliveryDate] [smalldatetime] NULL,
	[Notes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PartsCommission] [numeric](10, 2) NULL,
	[BearingFreightCharge] [numeric](10, 2) NULL,
	[BearingCommission] [numeric](10, 2) NULL,
	[DateApproved] [smalldatetime] NULL,
	[DateQuoted] [smalldatetime] NULL,
	[SpacerPreparation] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[SpacerPreparationHours] [numeric](18, 0) NULL,
	[LaborCommission] [numeric](10, 2) NULL,
	[SubWorkCommission] [numeric](10, 2) NULL,
	[QuoteContactId] [int] NULL,
	[DeliveryInformation] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ExpeditedDeliveryInformation] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QuoteSpecificComments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[QuotedById] [int] NULL,
	[HandlingCharge] [numeric](10, 2) NULL,
 CONSTRAINT [PK_tblQuotes] PRIMARY KEY CLUSTERED 
(
	[QuoteId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblQuotes_PartsCommission]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblQuotes] ADD  CONSTRAINT [DF_tblQuotes_PartsCommission]  DEFAULT (0.1) FOR [PartsCommission]
END

GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblQuotes_BearingCommission]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblQuotes] ADD  CONSTRAINT [DF_tblQuotes_BearingCommission]  DEFAULT (0.1) FOR [BearingCommission]
END

GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblQuotes_DateQuoted]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblQuotes] ADD  CONSTRAINT [DF_tblQuotes_DateQuoted]  DEFAULT (convert(varchar,getdate(),101)) FOR [DateQuoted]
END

GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblQuotes_LaborCommission]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblQuotes] ADD  CONSTRAINT [DF_tblQuotes_LaborCommission]  DEFAULT (0.1) FOR [LaborCommission]
END

GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblQuotes_SubWorkCommission]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblQuotes] ADD  CONSTRAINT [DF_tblQuotes_SubWorkCommission]  DEFAULT (0.1) FOR [SubWorkCommission]
END

GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DF_tblQuotes_HandlingCharge]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblQuotes] ADD  CONSTRAINT [DF_tblQuotes_HandlingCharge]  DEFAULT (0) FOR [HandlingCharge]
END

GO
/****** Object:  Table [dbo].[tblQuoteBearings]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblQuoteBearings.ProductName uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 10, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblQuoteBearings]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblQuoteBearings](
	[QuoteBearingId] [int] IDENTITY(1,1) NOT NULL,
	[QuoteId] [int] NULL,
	[ProductId] [int] NULL,
	[CenBearingCode] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CssPrice] [money] NULL,
	[BearingCost] [money] NULL,
	[BearingMarkup] [numeric](10, 2) NULL,
	[SupplierId] [int] NULL,
	[ProductName] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BearingDescription] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Qty] [numeric](10, 2) NULL,
 CONSTRAINT [PK_tblQuoteBearings] PRIMARY KEY CLUSTERED 
(
	[QuoteBearingId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteQuoteBearingByBearingId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteQuoteBearingByBearingId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteQuoteBearingByBearingId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteQuoteBearingByBearingId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteBearingId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblQuoteBearings WHERE QuoteBearingId = @QuoteBearingId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblQuoteBearings WHERE QuoteBearingId = @QuoteBearingId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0


GO
/****** Object:  View [dbo].[vwInventoryTransactions]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwInventoryTransactions]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwInventoryTransactions
AS
SELECT     dbo.tblProducts.ProductID, SUM(ISNULL(dbo.tblInventoryTransactions.UnitsReceived, 0)) - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsShrinkage, 
                      0)) - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsSold, 0)) AS OnHand, SUM(ISNULL(dbo.tblInventoryTransactions.UnitsOrdered, 0)) 
                      - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsReceived, 0)) AS OnOrder, MAX(ISNULL(dbo.tblProducts.ReorderLevel, 0)) AS ReorderLevel, 
                      dbo.tblProductOwners.ProductOwner, dbo.tblProductOwners.ProductOwnerID, dbo.tblCategories.CategoryID, dbo.tblCategories.CategoryName, 
                      ISNULL(dbo.tblProducts.UnitPrice, 0) AS UnitPrice, dbo.tblProducts.PartNumber
FROM         dbo.tblProducts LEFT OUTER JOIN
                      dbo.tblInventoryTransactions ON dbo.tblProducts.ProductID = dbo.tblInventoryTransactions.ProductID INNER JOIN
                      dbo.tblCategories ON dbo.tblProducts.CategoryID = dbo.tblCategories.CategoryID INNER JOIN
                      dbo.tblProductOwners ON dbo.tblProducts.ProductOwnerID = dbo.tblProductOwners.ProductOwnerID
WHERE     (NOT (dbo.tblProducts.PartNumber IS NULL))
GROUP BY dbo.tblProducts.ProductID, dbo.tblProductOwners.ProductOwner, dbo.tblProductOwners.ProductOwnerID, dbo.tblCategories.CategoryID, 
                      dbo.tblCategories.CategoryName, dbo.tblProducts.UnitPrice, dbo.tblProducts.PartNumber

' 
GO
/****** Object:  View [dbo].[vwInventoryTransactionReport]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwInventoryTransactionReport]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwInventoryTransactionReport
AS
SELECT     dbo.vwInventoryTransactions.ProductID, dbo.vwInventoryTransactions.OnHand, dbo.vwInventoryTransactions.OnOrder, 
                      dbo.vwInventoryTransactions.ReorderLevel, dbo.vwInventoryTransactions.ProductOwner, dbo.vwInventoryTransactions.ProductOwnerID, 
                      dbo.vwInventoryTransactions.CategoryID, dbo.vwInventoryTransactions.CategoryName, dbo.vwInventoryTransactions.UnitPrice, 
                      dbo.vwInventoryTransactions.PartNumber, dbo.tblProducts.ProductShortDescription
FROM         dbo.vwInventoryTransactions INNER JOIN
                      dbo.tblProducts ON dbo.vwInventoryTransactions.ProductID = dbo.tblProducts.ProductID

' 
GO
/****** Object:  StoredProcedure [dbo].[spReportInventoryTransactions]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spReportInventoryTransactions]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spReportInventoryTransactions] AS' 
END
GO


ALTER  PROCEDURE [dbo].[spReportInventoryTransactions]

@ProductOwnerID		AS int = 1,
@OrderBy		AS varchar(50) = 'ProductName'

AS


SELECT ProductID, PartNumber, CategoryName, ProductShortDescription, ReorderLevel, OnHand, OnOrder, UnitPrice
FROM vwInventoryTransactionReport
WHERE ProductOwnerID = @ProductOwnerID

ORDER BY
	CASE @OrderBy WHEN 'ProductID' THEN ( ProductID ) ELSE NULL END,
	CASE @OrderBy WHEN 'PartNumber' THEN ( PartNumber ) ELSE NULL END,
	CASE @OrderBy WHEN 'CategoryName' THEN ( CategoryName ) ELSE NULL END,
	CASE @OrderBy WHEN 'ReorderLevel' THEN ( ReorderLevel ) ELSE NULL END,
	CASE @OrderBy WHEN 'OnHand' THEN ( OnHand ) ELSE NULL END,
	CASE @OrderBy WHEN 'OnOrder' THEN ( OnOrder ) ELSE NULL END,
	CASE @OrderBy WHEN 'UnitPrice' THEN ( UnitPrice ) ELSE NULL END

GO
/****** Object:  StoredProcedure [dbo].[spDeleteQuoteByQuoteId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteQuoteByQuoteId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteQuoteByQuoteId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteQuoteByQuoteId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblQuotes WHERE QuoteId = @QuoteId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/


BEGIN TRY
	DELETE FROM tblQuotes WHERE QuoteId = @QuoteId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0

GO
/****** Object:  StoredProcedure [dbo].[spReportOpenProjects]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spReportOpenProjects]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spReportOpenProjects] AS' 
END
GO



ALTER PROCEDURE [dbo].[spReportOpenProjects]

AS

SELECT ProjectId, ProjectName, ISNULL( ProjectDescription, '&nbsp;' ) AS ProjectDescription, CONVERT( varchar(10), StartDate, 101 ) AS StartDate, ISNULL( Customer, '&nbsp;' ) AS Customer, 
ISNULL( ProjectPriority, '&nbsp;' ) AS ProjectPriority
FROM tblProjects
LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
LEFT JOIN tblProjectPriorities ON tblProjects.ProjectPriorityId = tblProjectPriorities.ProjectPriorityId
WHERE CompletionDate IS NULL
ORDER BY Customer



GO
/****** Object:  Table [dbo].[tblQuoteParts]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblQuoteParts]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblQuoteParts](
	[QuotePartId] [int] IDENTITY(1,1) NOT NULL,
	[QuoteId] [int] NULL,
	[ProductId] [int] NULL,
	[PartCost] [money] NULL,
	[Markup] [numeric](10, 2) NULL,
	[PartDescription] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[SupplierId] [int] NULL,
	[Qty] [numeric](10, 2) NULL,
 CONSTRAINT [PK_tblQuoteParts] PRIMARY KEY CLUSTERED 
(
	[QuotePartId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteQuotePartByPartId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteQuotePartByPartId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteQuotePartByPartId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteQuotePartByPartId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuotePartId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblQuoteParts WHERE QuotePartId = @QuotePartId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblQuoteParts WHERE QuotePartId = @QuotePartId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0


GO
/****** Object:  StoredProcedure [dbo].[spReportOpenPurchaseOrders]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spReportOpenPurchaseOrders]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spReportOpenPurchaseOrders] AS' 
END
GO



ALTER PROCEDURE [dbo].[spReportOpenPurchaseOrders]

AS

SELECT PurchaseOrderId, ISNULL( PurchaseOrderNumber, '&nbsp;' ) AS PurchaseOrderNumber, ISNULL( PurchaseOrderDescription, '&nbsp;' ) AS PurchaseOrderDescription, 
CONVERT( varchar(10), OrderDate, 101 ) AS OrderDate, CONVERT( varchar(10), DateRequired, 101 ) AS DateRequired, ISNULL( SupplierName, '&nbsp;' ) AS SupplierName
FROM tblPurchaseOrders 
LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierId
WHERE DateClosed IS NULL
ORDER BY SupplierName


GO
/****** Object:  Table [dbo].[tblQuoteSubWork]    Script Date: 1/14/2019 12:24:24 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblQuoteSubWork.SubWorkDescription uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 6, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblQuoteSubWork]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblQuoteSubWork](
	[QuoteSubWorkId] [int] IDENTITY(1,1) NOT NULL,
	[QuoteId] [int] NULL,
	[SubWorkCost] [money] NULL,
	[SupplierId] [int] NULL,
	[SubWorkDescription] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
 CONSTRAINT [PK_tblQuoteSubWork] PRIMARY KEY CLUSTERED 
(
	[QuoteSubWorkId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteQuoteSubWorkBySubWorkId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteQuoteSubWorkBySubWorkId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteQuoteSubWorkBySubWorkId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteQuoteSubWorkBySubWorkId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteSubWorkId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblQuoteSubWork WHERE QuoteSubWorkId = @QuoteSubWorkId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblQuoteSubWork WHERE QuoteSubWorkId = @QuoteSubWorkId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0


GO
/****** Object:  StoredProcedure [dbo].[spReportOpenQuotes]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spReportOpenQuotes]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spReportOpenQuotes] AS' 
END
GO



ALTER PROCEDURE [dbo].[spReportOpenQuotes]

AS

SELECT QuoteId, tblQuotes.WorkOrderId, ISNULL( tblQuotes.WorkOrderNumber, '&nbsp;' ) AS WorkOrderNumber, CONVERT( varchar(10), DateQuoted, 101 ) AS DateQuoted, CONVERT( varchar(10), DateApproved, 101 ) AS DateApproved, 
ISNULL( Customer, '&nbsp;' ) Customer, ISNULL( SpindleType, '&nbsp;' ) AS SpindleType
FROM tblQuotes
LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
WHERE tblQuotes.DateApproved IS NULL
ORDER BY Customer

GO
/****** Object:  StoredProcedure [dbo].[spDeleteShippingMethodByShippingMethodId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteShippingMethodByShippingMethodId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteShippingMethodByShippingMethodId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteShippingMethodByShippingMethodId]

@ErrorMessage		varchar(255)	OUTPUT,
@ShippingMethodId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblShippingMethods WHERE ShippingMethodId = @ShippingMethodId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/


BEGIN TRY
	DELETE FROM tblShippingMethods WHERE ShippingMethodId = @ShippingMethodId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0

GO
/****** Object:  StoredProcedure [dbo].[spReportOpenWorkOrders]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spReportOpenWorkOrders]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spReportOpenWorkOrders] AS' 
END
GO






ALTER PROCEDURE [dbo].[spReportOpenWorkOrders]

AS

SELECT WorkOrderId, ISNULL( WorkOrderNumber, '&nbsp;' ) AS WorkOrderNumber, CONVERT( varchar(10), DateIn, 1 ) AS DateIn, ISNULL( Customer, '&nbsp;' ) AS Customer, 
WorkOrderPriority AS Priority, ISNULL( SpindleType, '&nbsp;' ) AS SpindleType, CONVERT( varchar(10), PromiseDate, 1 ) AS PromiseDate, CONVERT( varchar(10), DateRec, 1 ) AS DateRec,
CONVERT( varchar(10), ExpectedDelDate, 1 ) AS ExpectedDelDate, CONVERT( varchar(10), [Date], 1 ) AS [Date], ISNULL( PromiseDate, '12/31/2049')  AS PromiseDateSorted
FROM tblWorkOrders
LEFT JOIN tblWorkOrderPriorities ON tblWorkOrders.WorkOrderPriorityId = tblWorkOrderPriorities.WorkOrderPriorityId
LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
WHERE 
DateOut IS NULL
ORDER BY tblWorkOrders.WorkOrderPriorityId DESC, PromiseDateSorted ASC, WorkOrderNumber ASC


GO
/****** Object:  StoredProcedure [dbo].[spDeleteSpindleBySpindleId]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteSpindleBySpindleId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteSpindleBySpindleId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteSpindleBySpindleId]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblSpindles WHERE SpindleId = @SpindleId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblSpindles WHERE SpindleId = @SpindleId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0
GO
/****** Object:  View [dbo].[vwOnHandProducts]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwOnHandProducts]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwOnHandProducts
AS
SELECT     dbo.tblInventoryTransactions.ProductID, SUM(ISNULL(dbo.tblInventoryTransactions.UnitsReceived, 0)) 
                      - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsShrinkage, 0)) - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsSold, 0)) AS OnHand, 
                      SUM(ISNULL(dbo.tblInventoryTransactions.UnitsOrdered, 0)) - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsReceived, 0)) AS OnOrder, 
                      dbo.tblProducts.ReorderLevel AS ReorderLevel, dbo.tblProducts.PartNumber, dbo.tblProducts.ProductShortDescription
FROM         dbo.tblInventoryTransactions INNER JOIN
                      dbo.tblProducts ON dbo.tblInventoryTransactions.ProductID = dbo.tblProducts.ProductID
GROUP BY dbo.tblInventoryTransactions.ProductID, dbo.tblProducts.ReorderLevel, dbo.tblProducts.PartNumber, dbo.tblProducts.ProductShortDescription

' 
GO
/****** Object:  StoredProcedure [dbo].[spReportProductsNeeded]    Script Date: 1/14/2019 12:24:24 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spReportProductsNeeded]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spReportProductsNeeded] AS' 
END
GO



ALTER PROCEDURE [dbo].[spReportProductsNeeded]

AS

SELECT ProductId, OnHand, OnOrder, ReorderLevel, ProductShortDescription, PartNumber
FROM vwOnHandProducts
WHERE ( OnHand <= ReorderLevel )
ORDER BY PartNumber

GO
/****** Object:  Table [dbo].[tblSpindleCategories]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblSpindleCategories]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblSpindleCategories](
	[SpindleCategoryId] [int] IDENTITY(1,1) NOT NULL,
	[SpindleCategoryName] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteSpindleCategoryBySpindleCategoryId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteSpindleCategoryBySpindleCategoryId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteSpindleCategoryBySpindleCategoryId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteSpindleCategoryBySpindleCategoryId]

@ErrorMessage		varchar(255)	OUTPUT,
@SpindleCategoryId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblSpindleCategories WHERE SpindleCategoryId = @SpindleCategoryId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblSpindleCategories WHERE SpindleCategoryId = @SpindleCategoryId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0

GO
/****** Object:  Table [dbo].[tblSpindlesProducts]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblSpindlesProducts]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblSpindlesProducts](
	[SpindleProductId] [int] IDENTITY(1000,1) NOT NULL,
	[SpindleId] [int] NULL,
	[ProductId] [int] NULL,
	[SupplierId] [int] NULL,
	[Cost] [money] NULL,
	[Markup] [numeric](10, 2) NULL,
	[Quantity] [numeric](10, 2) NULL
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteSpindleProductBySpindleProductId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteSpindleProductBySpindleProductId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteSpindleProductBySpindleProductId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteSpindleProductBySpindleProductId]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleProductId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblSpindlesProducts WHERE SpindleProductId = @SpindleProductId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblSpindlesProducts WHERE SpindleProductId = @SpindleProductId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0

GO
/****** Object:  Table [dbo].[tblSpindlesSubWork]    Script Date: 1/14/2019 12:24:25 AM ******/
/**
Assessment issue: Deprecated data types TEXT, IMAGE or NTEXT
Categories: Compatibility, Information
Applicable compatibility levels: CompatLevel100, CompatLevel110, CompatLevel120, CompatLevel130, CompatLevel140
Impact: These data types are checked as deprecated. In some cases, using TEXT, IMAGE or NTEXT might harm performance.
Impact details: Object dbo.tblSpindlesSubWork.SubWorkDescription uses deprecated data type TEXT, IMAGE or NTEXT which will be discontinued for future versions of SQL Server. For more details, please see: Line 4, Column 5.
Recommendation: Deprecated data types are marked to be discontinued on next versions of SQL Server, should use new data types such as: (varchar(max), nvarchar(max), varbinary(max) and etc.)
More information: ntext, text, and image (Transact-SQL) (https://go.microsoft.com/fwlink/?LinkId=798558)
 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblSpindlesSubWork]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblSpindlesSubWork](
	[SpindleSubWorkId] [int] IDENTITY(1,1) NOT NULL,
	[SpindleId] [int] NULL,
	[SubWorkDescription] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[SupplierId] [int] NULL,
	[SubWorkCost] [money] NULL
)
END
GO
/****** Object:  StoredProcedure [dbo].[spDeleteSpindleSubWorkBySpindleSubWorkId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteSpindleSubWorkBySpindleSubWorkId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteSpindleSubWorkBySpindleSubWorkId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteSpindleSubWorkBySpindleSubWorkId]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleSubWorkId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblSpindlesSubWork WHERE SpindleSubWorkId = @SpindleSubWorkId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblSpindlesSubWork WHERE SpindleSubWorkId = @SpindleSubWorkId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0
GO
/****** Object:  StoredProcedure [dbo].[spDeleteSupplierBySupplierId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteSupplierBySupplierId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteSupplierBySupplierId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteSupplierBySupplierId]

@ErrorMessage	varchar(255)	OUTPUT,
@SupplierId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblSuppliers WHERE SupplierId = @SupplierId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/


BEGIN TRY
	DELETE FROM tblSuppliers WHERE SupplierId = @SupplierId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0

GO
/****** Object:  StoredProcedure [dbo].[spDeleteWorkOrderByWorkOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteWorkOrderByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteWorkOrderByWorkOrderId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteWorkOrderByWorkOrderId]

@ErrorMessage	varchar(255)	OUTPUT,
@WorkOrderId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblWorkOrders WHERE WorkOrderId = @WorkOrderId
DELETE FROM tblQCInspections WHERE WorkOrderId = @WorkOrderId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/
BEGIN TRY
	DELETE FROM tblWorkOrders WHERE WorkOrderId = @WorkOrderId
	DELETE FROM tblQCInspections WHERE WorkOrderId = @WorkOrderId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0


GO
/****** Object:  StoredProcedure [dbo].[spDeleteWorkOrderPO]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteWorkOrderPO]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteWorkOrderPO] AS' 
END
GO

ALTER PROCEDURE [dbo].[spDeleteWorkOrderPO]
@WorkOrderId		int,
@PurchaseOrderId	int

AS

BEGIN
	DELETE FROM tblWorkOrderPOs WHERE ( (WorkOrderId = @WorkOrderId) AND (PurchaseOrderId = @PurchaseOrderId) )
END

GO
/****** Object:  View [dbo].[vwCustomersWorkOrdersSpindles]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwCustomersWorkOrdersSpindles]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwCustomersWorkOrdersSpindles
AS
SELECT     dbo.tblWorkOrders.WorkOrderId, dbo.tblWorkOrders.WorkOrderNumber, dbo.tblWorkOrders.SerialNumber, dbo.tblCustomers.CustomerId, 
                      dbo.tblCustomers.Customer, dbo.tblWorkOrders.ProjectId, dbo.tblProjects.ProjectName, dbo.tblSpindles.SpindleId, dbo.tblSpindles.SpindleType, 
                      dbo.tblSpindleCategories.SpindleCategoryName
FROM         dbo.tblWorkOrders INNER JOIN
                      dbo.tblProjects ON dbo.tblWorkOrders.ProjectId = dbo.tblProjects.ProjectId INNER JOIN
                      dbo.tblCustomers ON dbo.tblProjects.CustomerId = dbo.tblCustomers.CustomerId INNER JOIN
                      dbo.tblSpindles ON dbo.tblProjects.SpindleId = dbo.tblSpindles.SpindleId INNER JOIN
                      dbo.tblSpindleCategories ON dbo.tblSpindles.SpindleCategoryId = dbo.tblSpindleCategories.SpindleCategoryId
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwCustomersWorkOrdersSpindles', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[48] 4[26] 2[16] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 0
               Left = 371
               Bottom = 404
               Right = 651
            End
            DisplayFlags = 280
            TopColumn = 34
         End
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 0
               Left = 0
               Bottom = 404
               Right = 279
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 408
               Left = 38
               Bottom = 523
               Right = 196
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblSpindles"
            Begin Extent = 
               Top = 408
               Left = 234
               Bottom = 523
               Right = 437
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblSpindleCategories"
            Begin Extent = 
               Top = 528
               Left = 38
               Bottom = 613
               Right = 225
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwCustomersWorkOrdersSpindles'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwCustomersWorkOrdersSpindles', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwCustomersWorkOrdersSpindles'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwCustomersWorkOrdersSpindles', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwCustomersWorkOrdersSpindles'
GO
/****** Object:  StoredProcedure [dbo].[uspGetCustomerWorkOrderSpindle]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[uspGetCustomerWorkOrderSpindle]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[uspGetCustomerWorkOrderSpindle] AS' 
END
GO




-- Creates a new record in the [dbo].[tblCustomers] table.
ALTER PROCEDURE [dbo].[uspGetCustomerWorkOrderSpindle]
    @p_Customer varchar(100),
    @p_SpindleType varchar(255)
    
AS
BEGIN
    SELECT * 
	FROM dbo.vwCustomersWorkOrdersSpindles 
	
	WHERE customer LIKE '%' + @p_Customer + '%' 
	AND 
	SpindleType LIKE + '%' + @p_SpindleType + '%'

END



GO
/****** Object:  StoredProcedure [dbo].[spDeleteWorkOrderPriorityByWorkOrderPriorityId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteWorkOrderPriorityByWorkOrderPriorityId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteWorkOrderPriorityByWorkOrderPriorityId] AS' 
END
GO


ALTER PROCEDURE [dbo].[spDeleteWorkOrderPriorityByWorkOrderPriorityId]

@ErrorMessage		varchar(255)	OUTPUT,
@WorkOrderPriorityId	int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblWorkOrderPriorities WHERE WorkOrderPriorityId = @WorkOrderPriorityId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/


BEGIN TRY
	DELETE FROM tblWorkOrderPriorities WHERE WorkOrderPriorityId = @WorkOrderPriorityId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0
GO
/****** Object:  StoredProcedure [dbo].[spEditCall]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditCall]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditCall] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditCall]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@CallId		int		= NULL,
@CustomerId	int,
@EmployeeId	int,
@ProjectId	int		= NULL,
@WorkOrderId	int		= NULL,
@CallComments	text		= NULL

AS
/*
SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( @EmployeeId = 0 ) SET @EmployeeId = NULL

IF( DATALENGTH( @CallId ) = 0 ) OR ( @CallId IS NULL ) OR ( @CallId = 0 )
BEGIN
	INSERT tblCalls( CustomerId, EmployeeId, ProjectId, WorkOrderId, CallComments ) 
	VALUES( @CustomerId, @EmployeeId, @ProjectId, @WorkOrderId, @CallComments )

	SET @ErrorNum	= @@ERROR
	SET @CallId 	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
		SET @CallId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT CallId FROM tblCalls WHERE CallId = @CallId ) )
	BEGIN
		UPDATE tblCalls SET CustomerId = @CustomerId, EmployeeId = @EmployeeId, ProjectId = @ProjectId, WorkOrderId = @WorkOrderId, CallComments = @CallComments
		WHERE CallId = @CallId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Call record not found'
		RETURN -1
	END
END

RETURN @CallId
*/
SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( @EmployeeId = 0 ) SET @EmployeeId = NULL

IF( DATALENGTH( @CallId ) = 0 ) OR ( @CallId IS NULL ) OR ( @CallId = 0 )
BEGIN
	INSERT tblCalls( CustomerId, EmployeeId, ProjectId, WorkOrderId, CallComments ) 
	VALUES( @CustomerId, @EmployeeId, @ProjectId, @WorkOrderId, @CallComments )

	SET @ErrorNum	= @@ERROR
	SET @CallId 	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @CallId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT CallId FROM tblCalls WHERE CallId = @CallId ) )
	BEGIN
		UPDATE tblCalls SET CustomerId = @CustomerId, EmployeeId = @EmployeeId, ProjectId = @ProjectId, WorkOrderId = @WorkOrderId, CallComments = @CallComments
		WHERE CallId = @CallId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Call record not found'
		RETURN -1
	END
END

RETURN @CallId
GO
/****** Object:  Table [dbo].[tblCompanyInformation]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblCompanyInformation]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblCompanyInformation](
	[SetupID] [int] NOT NULL,
	[CompanyName] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Address] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[City] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[StateOrProvince] [varchar](20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PostalCode] [varchar](20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[Country] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PhoneNumber] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[FaxNumber] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
END
GO
/****** Object:  StoredProcedure [dbo].[spEditCompanyInformation]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditCompanyInformation]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditCompanyInformation] AS' 
END
GO


ALTER PROCEDURE [dbo].[spEditCompanyInformation]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@CompanyName		varchar(50)	= NULL,
@Address		varchar(255)	= NULL,
@City			varchar(50)	= NULL,
@StateOrProvince	varchar(20)	= NULL,
@PostalCode		varchar(20)	= NULL,
@Country		varchar(50)	= NULL,
@PhoneNumber		varchar(30)	= NULL,
@FaxNumber		varchar(30)	= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

UPDATE tblCompanyInformation SET CompanyName = @CompanyName, Address = @Address, City = @City, StateOrProvince = @StateOrProvince, PostalCode = @PostalCode,
	Country = @Country, PhoneNumber = @PhoneNumber, FaxNumber = @FaxNumber
WHERE SetupId = 1

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END

RETURN 0


GO
/****** Object:  StoredProcedure [dbo].[spEditCustomer]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditCustomer]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditCustomer] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditCustomer]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@CustomerId		int 		= NULL,
@Customer		varchar(100),
@Address		varchar(255)	= NULL,
@City			varchar(100)	= NULL,
@State			varchar(100)	= NULL,
@Country		varchar(100)	= NULL,
@Zip			varchar(100)	= NULL,
@SalesRepId		int		= NULL,
@DateEstablished	varchar(100)	= NULL,
@MainContactId		int		= NULL,
@Notes			text		= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @CustomerId ) = 0 ) OR ( @CustomerId IS NULL ) OR ( @CustomerId = 0 )
BEGIN
	INSERT tblCustomers( Customer, Address, City, State, Country, Zip, SalesRepId, DateEstablished, MainContactId, Notes ) 
	VALUES( @Customer, @Address, @City, @State, @Country, @Zip, @SalesRepId, @DateEstablished, @MainContactId, @Notes )

	SET @ErrorNum	= @@ERROR
	SET @CustomerId = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @CustomerId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT CustomerId FROM tblCustomers WHERE CustomerId = @CustomerId ) )
	BEGIN
		UPDATE tblCustomers SET Customer = @Customer, Address = @Address, City = @City, State = @State, Country = @Country, Zip = @Zip, 
			SalesRepId = @SalesRepId, DateEstablished = @DateEstablished, MainContactId = @MainContactId, Notes = @Notes
		WHERE CustomerId = @CustomerId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Customer record not found'
		RETURN -1
	END
END

RETURN @CustomerId

GO
/****** Object:  StoredProcedure [dbo].[spEditCustomerContact]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditCustomerContact]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditCustomerContact] AS' 
END
GO

ALTER PROCEDURE [dbo].[spEditCustomerContact]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@CustomerContactId	int 		= NULL,
@CustomerId		int 		= NULL,
@Contact		varchar(100)	= NULL,
@ContactTitle		varchar(100)	= NULL,
@Department		varchar(100)	= NULL,
@TelephoneNumber	varchar(100)	= NULL,
@Extension		varchar(100)	= NULL,
@MobileNumber		varchar(100)	= NULL,
@FaxNumber		varchar(100)	= NULL,
@EmailAddress		varchar(100)	= NULL,
@Notes			text		= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @CustomerContactId ) = 0 ) OR ( @CustomerContactId IS NULL ) OR ( @CustomerContactId = 0 )
BEGIN
	INSERT tblCustomerContacts( CustomerId, Contact, ContactTitle, Department, TelephoneNumber, Extension, MobileNumber, FaxNumber,
			    EmailAddress, Notes ) VALUES( @CustomerId, @Contact, @ContactTitle, @Department, @TelephoneNumber, @Extension,
			    @MobileNumber, @FaxNumber, @EmailAddress, @Notes )

	SET @ErrorNum	= @@ERROR
	SET @CustomerContactId = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @CustomerContactId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT CustomerContactId FROM tblCustomerContacts WHERE CustomerContactId = @CustomerContactId ) )
	BEGIN
		UPDATE tblCustomerContacts SET Contact = @Contact, ContactTitle = @ContactTitle, Department = @Department,
				TelephoneNumber = @TelephoneNumber, Extension = @Extension, MobileNumber = @MobileNumber,
			       	FaxNumber = @FaxNumber, EmailAddress = @EmailAddress, Notes = @Notes
		WHERE CustomerContactId = @CustomerContactId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Customer Contact record not found'
		RETURN -1
	END
END

RETURN @CustomerContactId

GO
/****** Object:  StoredProcedure [dbo].[spEditEmployee]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditEmployee]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditEmployee] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditEmployee]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@EmployeeID		int		= NULL,
@FirstName		varchar(100),
@LastName		varchar(100),
@Title			varchar(100)	= NULL,
@Extension		varchar(100)	= NULL,
@WorkPhone		varchar(100)	= NULL,
@EmailAddress		varchar(100)	= NULL 

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @EmployeeID ) = 0 ) OR ( @EmployeeID IS NULL ) OR ( @EmployeeID = 0 )
BEGIN
	INSERT tblEmployees( LastName, FirstName, Title, WorkPhone, Extension, EmailAddress ) 
	VALUES( @LastName, @FirstName, @Title, @WorkPhone, @Extension, @EmailAddress )

	SET @ErrorNum	= @@ERROR
	SET @EmployeeID = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @EmployeeID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT EmployeeId FROM tblEmployees WHERE EmployeeId = @EmployeeID ) )
	BEGIN
		UPDATE tblEmployees SET LastName = @LastName, FirstName = @FirstName, Title = @Title, WorkPhone = @WorkPhone, Extension = @Extension, EmailAddress = @EmailAddress
		WHERE EmployeeId = @EmployeeID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Employee record not found'
		RETURN -1
	END
END

RETURN @EmployeeID


GO
/****** Object:  StoredProcedure [dbo].[spEditInventoryTransaction]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditInventoryTransaction]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditInventoryTransaction] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditInventoryTransaction]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@TransactionID		int,
@TransactionDescription	text		= NULL,
@UnitPrice		money	= NULL,
@UnitsOrdered		numeric(10,2)	= NULL,
@DateReceived		smalldatetime	= NULL,
@UnitsReceived		numeric(10,2)	= NULL,
@UnitsSold		numeric(10,2)	= NULL,
@UnitsShrinkage		numeric(10,2)	= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @TransactionID ) = 0 ) OR ( @TransactionID IS NULL ) OR ( @TransactionID = 0 )
BEGIN
	RETURN -1
	SELECT @ErrorMessage = 'Inventory Transaction record not found'
END
ELSE
BEGIN
	IF( EXISTS( SELECT TransactionId FROM tblInventoryTransactions WHERE TransactionId = @TransactionID ) )
	BEGIN
		UPDATE tblInventoryTransactions SET TransactionDescription = @TransactionDescription, UnitPrice = @UnitPrice, UnitsOrdered = @UnitsOrdered,
		DateReceived = @DateReceived, UnitsReceived = @UnitsReceived, UnitsSold = @UnitsSold, UnitsShrinkage = @UnitsShrinkage
		WHERE TransactionId = @TransactionID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Inventory Transaction record not found'
		RETURN -1
	END
END

RETURN @TransactionID

GO
/****** Object:  StoredProcedure [dbo].[spEditLocation]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditLocation]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditLocation] AS' 
END
GO


ALTER PROCEDURE [dbo].[spEditLocation]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@LocationID	int		= NULL,
@Location	varchar(100)

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @LocationID ) = 0 ) OR ( @LocationID IS NULL ) OR ( @LocationID = 0 )
BEGIN
	INSERT tblLocations( Location ) VALUES ( @Location )

	SET @ErrorNum	= @@ERROR
	SET @LocationID = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @LocationID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT LocationID FROM tblLocations WHERE LocationID = @LocationID ) )
	BEGIN
		UPDATE tblLocations SET Location = @Location
		WHERE LocationID = @LocationID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Location record not found'
		RETURN -1
	END
END

RETURN @LocationID


GO
/****** Object:  StoredProcedure [dbo].[spEditProduct]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditProduct]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditProduct] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditProduct]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@ProductID		int		= NULL,
@PartNumber		varchar(100),
@ProductName		text		= NULL,
@ProductDescription	text		= NULL,
@CategoryID		int,
@SerialNumber		varchar(100)	= NULL,
@UnitPrice		numeric(10,2)	= NULL,
@ReorderLevel		int		= NULL,
@LeadTime		varchar(100)	= NULL,
@DrawingID		varchar(100)	= NULL,
@ProductOwnerID		int

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int
DECLARE @ProductShortDescription varchar(3000)

SELECT @ProductShortDescription = CAST (@ProductDescription as varchar(3000))

IF( DATALENGTH( @ProductID ) = 0 ) OR ( @ProductID IS NULL ) OR ( @ProductID = 0 )
BEGIN
	INSERT tblProducts( PartNumber, ProductName, ProductDescription, ProductShortDescription, CategoryID, SerialNumber, UnitPrice, ReorderLevel, LeadTime, DrawingID, ProductOwnerID )
	VALUES( @PartNumber, @ProductName, @ProductDescription, @ProductShortDescription, @CategoryID, @SerialNumber, @UnitPrice, @ReorderLevel, @LeadTime, @DrawingID, @ProductOwnerID )

	SET @ErrorNum	= @@ERROR
	SET @ProductID 	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @ProductID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT ProductId FROM tblProducts WHERE ProductId = @ProductID ) )
	BEGIN
		UPDATE tblProducts SET PartNumber = @PartNumber, ProductName = @ProductName, ProductDescription = @ProductDescription, ProductShortDescription = @ProductShortDescription, CategoryID = @CategoryID, SerialNumber = @SerialNumber,
		UnitPrice = @UnitPrice, ReorderLevel = @ReorderLevel, LeadTime = @LeadTime, DrawingID = @DrawingID, ProductOwnerID = @ProductOwnerID
		WHERE ProductId = @ProductID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Product record not found'
		RETURN -1
	END
END

RETURN @ProductID

GO
/****** Object:  StoredProcedure [dbo].[spEditProductCategory]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditProductCategory]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditProductCategory] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditProductCategory]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@ProductCategoryID	int		= NULL,
@ProductCategory	varchar(100)

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @ProductCategoryID ) = 0 ) OR ( @ProductCategoryID IS NULL ) OR ( @ProductCategoryID = 0 )
BEGIN
	INSERT tblCategories( CategoryName ) VALUES ( @ProductCategory )

	SET @ErrorNum		= @@ERROR
	SET @ProductCategoryID	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @ProductCategoryID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT CategoryID FROM tblCategories WHERE CategoryID = @ProductCategoryID ) )
	BEGIN
		UPDATE tblCategories SET CategoryName = @ProductCategory
		WHERE CategoryID = @ProductCategoryID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Product category record not found'
		RETURN -1
	END
END

RETURN @ProductCategoryID


GO
/****** Object:  StoredProcedure [dbo].[spEditProductOwner]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditProductOwner]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditProductOwner] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spEditProductOwner]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@ProductOwnerID	int		= NULL,
@ProductOwner	varchar(100)

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @ProductOwnerID ) = 0 ) OR ( @ProductOwnerID IS NULL ) OR ( @ProductOwnerID = 0 )
BEGIN
	INSERT tblProductOwners( ProductOwner ) VALUES ( @ProductOwner )

	SET @ErrorNum		= @@ERROR
	SET @ProductOwnerID	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @ProductOwnerID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT ProductOwnerID FROM tblProductOwners WHERE ProductOwnerID = @ProductOwnerID ) )
	BEGIN
		UPDATE tblProductOwners SET ProductOwner = @ProductOwner
		WHERE ProductOwnerID = @ProductOwnerID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Product Owner record not found'
		RETURN -1
	END
END

RETURN @ProductOwnerID



GO
/****** Object:  StoredProcedure [dbo].[spEditProject]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditProject]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditProject] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spEditProject]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@ProjectId		int		= NULL,
@ProjectName		varchar(100),
@ProjectTypeId		int,
@CustomerId		int,
@CustomerPO		varchar(100)	= NULL,
@ProjectDescription	text		= NULL,
@StartDate		datetime	= NULL,
@EstCompletionDate	datetime	= NULL,
@CompletionDate		datetime	= NULL,
@SpindleId		int		= NULL,
@ProjectContactId	int		= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @ProjectId ) = 0 ) OR ( @ProjectId IS NULL ) OR ( @ProjectId = 0 )
BEGIN
	DECLARE @MainContactId AS int
	SELECT  @MainContactId = MainContactId FROM tblCustomers WHERE CustomerId = @CustomerId

	INSERT tblProjects( ProjectName, ProjectTypeId, CustomerId, CustomerPO, ProjectDescription, StartDate, EstCompletionDate, CompletionDate, 
	SpindleId, ProjectContactId )
	VALUES( @ProjectName, @ProjectTypeId, @CustomerId, @CustomerPO, @ProjectDescription, @StartDate, @EstCompletionDate, @CompletionDate, 
	@SpindleId, @MainContactId )

	SET @ErrorNum	= @@ERROR
	SET @ProjectId 	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @ProjectId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT ProjectId FROM tblProjects WHERE ProjectId = @ProjectId ) )
	BEGIN
		UPDATE tblProjects SET ProjectName = @ProjectName, ProjectTypeId = @ProjectTypeId, CustomerId = @CustomerId, CustomerPO = @CustomerPO,
		ProjectDescription = @ProjectDescription, StartDate = @StartDate, EstCompletionDate = @EstCompletionDate, CompletionDate = @CompletionDate, 
		SpindleId = @SpindleId, ProjectContactId = @ProjectContactId
		WHERE ProjectId = @ProjectId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Project record not found'
		RETURN -1
	END
END

RETURN @ProjectId

GO
/****** Object:  StoredProcedure [dbo].[spEditProjectPriority]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditProjectPriority]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditProjectPriority] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditProjectPriority]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@ProjectPriorityID	int		= NULL,
@ProjectPriority	varchar(100)

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @ProjectPriorityID ) = 0 ) OR ( @ProjectPriorityID IS NULL ) OR ( @ProjectPriorityID = 0 )
BEGIN
	INSERT tblProjectPriorities( ProjectPriority ) VALUES ( @ProjectPriority )

	SET @ErrorNum	= @@ERROR
	SET @ProjectPriorityID = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @ProjectPriorityID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT ProjectPriorityID FROM tblProjectPriorities WHERE ProjectPriorityID = @ProjectPriorityID ) )
	BEGIN
		UPDATE tblProjectPriorities SET ProjectPriority = @ProjectPriority
		WHERE ProjectPriorityID = @ProjectPriorityID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Project Priority record not found'
		RETURN -1
	END
END

RETURN @ProjectPriorityID


GO
/****** Object:  StoredProcedure [dbo].[spEditProjectType]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditProjectType]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditProjectType] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditProjectType]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@ProjectTypeID	int		= NULL,
@ProjectType	varchar(100)

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @ProjectTypeID ) = 0 ) OR ( @ProjectTypeID IS NULL ) OR ( @ProjectTypeID = 0 )
BEGIN
	INSERT tblProjectTypes( ProjectType ) VALUES ( @ProjectType )

	SET @ErrorNum	= @@ERROR
	SET @ProjectTypeID = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @ProjectTypeID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT ProjectTypeID FROM tblProjectTypes WHERE ProjectTypeID = @ProjectTypeID ) )
	BEGIN
		UPDATE tblProjectTypes SET ProjectType = @ProjectType
		WHERE ProjectTypeID = @ProjectTypeID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Project Type record not found'
		RETURN -1
	END
END

RETURN @ProjectTypeID


GO
/****** Object:  StoredProcedure [dbo].[spEditPurchaseOrder]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditPurchaseOrder]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditPurchaseOrder] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditPurchaseOrder]

@ErrorMessage			varchar(255) 	= NULL OUTPUT,
@PurchaseOrderID		int		= NULL,
@PurchaseOrderNumber		varchar(100),
@PurchaseOrderDescription	text		= NULL,
@SupplierID			int,
@EmployeeID			int,
@OrderDate			smalldatetime	= NULL,
@DateRequired			smalldatetime	= NULL,
@DatePromised			smalldatetime	= NULL,
@DateClosed			smalldatetime	= NULL,
@ShipDate	smalldatetime = NULL,
@DateShippedToSupplier	smalldatetime = NULL,
@ShippingMethodID		int,
@TrackingNumber		varchar(100)	= NULL,
@FreightCharge			money		= NULL,
@InternalNotes	text = NULL,
@PODescription	text = NULL


AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( @OrderDate IS NULL ) SET @OrderDate = getdate()

IF( DATALENGTH( @PurchaseOrderID ) = 0 ) OR ( @PurchaseOrderID IS NULL ) OR ( @PurchaseOrderID = 0 )
BEGIN
	INSERT tblPurchaseOrders( PurchaseOrderNumber, PurchaseOrderDescription, SupplierID, EmployeeID, OrderDate, DateRequired, DatePromised,
				  DateClosed, ShippingMethodID, FreightCharge, ShipDate, DateShippedToSupplier, TrackingNumber, InternalNotes, PODescription ) VALUES ( @PurchaseOrderNumber, @PurchaseOrderDescription, @SupplierID, 
				  @EmployeeID, @OrderDate, @DateRequired, @DateRequired, @DateClosed, @ShippingMethodID, @FreightCharge, @ShipDate, @DateShippedToSupplier, @TrackingNumber, @InternalNotes, @PODescription )

	SET @ErrorNum		= @@ERROR
	SET @PurchaseOrderID 	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @PurchaseOrderID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT PurchaseOrderID FROM tblPurchaseOrders WHERE PurchaseOrderID = @PurchaseOrderID ) )
	BEGIN
		UPDATE tblPurchaseOrders SET PurchaseOrderNumber = @PurchaseOrderNumber, PurchaseOrderDescription = @PurchaseOrderDescription, SupplierID = @SupplierID, 
					     EmployeeID = @EmployeeID, OrderDate = @OrderDate, DateRequired = @DateRequired, DatePromised = @DatePromised, ShipDate = @ShipDate,
					     DateClosed = @DateClosed, ShippingMethodID = @ShippingMethodID, FreightCharge = @FreightCharge, DateShippedToSupplier = @DateShippedToSupplier,
					     TrackingNumber = @TrackingNumber, InternalNotes = @InternalNotes, PODescription = @PODescription
		WHERE PurchaseOrderID = @PurchaseOrderID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Purchase order not found'
		RETURN -1
	END
END

RETURN @PurchaseOrderID
GO
/****** Object:  StoredProcedure [dbo].[spEditPurchaseOrderDetail]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditPurchaseOrderDetail]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditPurchaseOrderDetail] AS' 
END
GO

ALTER PROCEDURE [dbo].[spEditPurchaseOrderDetail]

@ErrorMessage			varchar(255) 	= NULL OUTPUT,
@TransactionID			int		= NULL,
@TransactionDate		smalldatetime	= NULL,
@ProductID			int,
@PurchaseOrderID		int,
@TransactionDescription		text		= NULL,
@UnitPrice			money,
@UnitsOrdered			numeric(10,2)	= 1

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( @TransactionDate IS NULL ) SET @TransactionDate = getdate()

IF( DATALENGTH( @TransactionID ) = 0 ) OR ( @TransactionID IS NULL ) OR ( @TransactionID = 0 )
BEGIN

	INSERT tblInventoryTransactions( TransactionDate, ProductId, PurchaseOrderId, TransactionDescription, UnitPrice, UnitsOrdered )
		VALUES ( @TransactionDate, @ProductID, @PurchaseOrderID, @TransactionDescription, @UnitPrice, @UnitsOrdered )

	SET @ErrorNum		= @@ERROR
	SET @TransactionID	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @TransactionID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT TransactionID FROM tblInventoryTransactions WHERE TransactionID = @TransactionID ) )
	BEGIN

		UPDATE tblInventoryTransactions SET TransactionDate = @TransactionDate, ProductId = @ProductId, PurchaseOrderId = @PurchaseOrderId, 
						TransactionDescription = @TransactionDescription, UnitPrice = @UnitPrice, UnitsOrdered = @UnitsOrdered
		WHERE TransactionID = @TransactionID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Purchase order detail record not found'
		RETURN -1
	END
END

RETURN @TransactionID


GO
/****** Object:  StoredProcedure [dbo].[spEditQCInspection]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQCInspection]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQCInspection] AS' 
END
GO





ALTER   PROCEDURE [dbo].[spEditQCInspection]

@ErrorMessage			varchar(255) 	= NULL OUTPUT,
@WorkOrderId			int		= NULL,

@RPM_In				numeric(10, 2)	= NULL,
@Weight_In			varchar(100)	= NULL,
@Vibration_In			varchar(100)	= NULL,
@VibrationRear_In		varchar(100)	= NULL,
@CoolantFlow_In			varchar(100)	= NULL,
@GSE_In				varchar(100)	= NULL,
@GSERear_In			varchar(100)	= NULL,
@ShaftTemp_In			varchar(100)	= NULL,
@CoolantTempIncomingSet_In	varchar(100)	= NULL,
@CoolantTempActual_In		varchar(100)	= NULL,
@FrontTemp_In			varchar(100)	= NULL,
@RearTemp_In			varchar(100)	= NULL,
@CoolantPressureIncoming_In	varchar(100)	= NULL,
@BreakInTime_In			varchar(100)	= NULL,
@CoolingMethod_In		varchar(100)	= NULL,
@Volts_In			numeric(10, 2)	= NULL,
@HP_In				numeric(10, 2)	= NULL,
@Amps_In			varchar(100)	= NULL,
@Phase_In			varchar(1)	= NULL,
@Hz_In				numeric(10, 2)	= NULL,
@Thermistor_In			varchar(100)	= NULL,
@Poles_In			numeric(10, 2)	= NULL,
@AmpDraw_In			varchar(100)	= NULL,
@ConnectorCtrl_In		varchar(100)	= NULL,
@ConnectorPower_In		varchar(100)	= NULL,
@Converter_In			varchar(100)	= NULL,
@ToolHolder_In			varchar(100)	= NULL,
@PullPin_In			varchar(100)	= NULL,
@EMDimension_In			varchar(100)	= NULL,
@EjectionPath_In		varchar(100)	= NULL,
@ToolOutPressure_In		varchar(100)	= NULL,
@ReturnPressure_In		varchar(100)	= NULL,
@DrawbarForce_In		varchar(100)	= NULL,
@ToolChangeFunction_In		varchar(100)	= NULL,
@ProximitySwitchFunction_In	varchar(100)	= NULL,
@Lubrication_In			varchar(100)	= NULL,
@Grease_In			varchar(100)	= NULL,
@OilMist_In			varchar(100)	= NULL,
@OilJet_In			varchar(100)	= NULL,
@OilGreaseType_In		varchar(100)	= NULL,
@IntervalDPM_In			varchar(100)	= NULL,
@MainPressure_In		varchar(100)	= NULL,
@TubePressure_In		varchar(100)	= NULL,
@Preload_In			varchar(100)	= NULL,
@RadialPlay_In			varchar(100)	= NULL,
@AxialPlay_In			varchar(100)	= NULL,
@RunoutFront_In			varchar(100)	= NULL,
@RunoutFrontLocation_In		varchar(100)	= NULL,
@RunoutFront2_In		varchar(100)	= NULL,
@RunoutFront2Location_In	varchar(100)	= NULL,
@RunoutRear_In			varchar(100)	= NULL,
@RunoutRearLocation_In		varchar(100)	= NULL,
@RunoutRear2_In			varchar(100)	= NULL,
@RunoutRear2Location_In		varchar(100)	= NULL,
@ToolContact_In			varchar(100)	= NULL,
@ToolContactRear_In		varchar(100)	= NULL,
@ToolGap_In			varchar(100)	= NULL,
@ToolGapRear_In			varchar(100)	= NULL,

@RPM_Final			numeric(10, 2)	= NULL,
@Weight_Final			varchar(100)	= NULL,
@Vibration_Final		varchar(100)	= NULL,
@VibrationRear_Final		varchar(100)	= NULL,
@CoolantFlow_Final		varchar(100)	= NULL,
@GSE_Final			varchar(100)	= NULL,
@GSERear_Final			varchar(100)	= NULL,
@ShaftTemp_Final		varchar(100)	= NULL,
@CoolantTempIncomingSet_Final	varchar(100)	= NULL,
@CoolantTempActual_Final	varchar(100)	= NULL,
@FrontTemp_Final		varchar(100)	= NULL,
@RearTemp_Final			varchar(100)	= NULL,
@CoolantPressureIncoming_Final	varchar(100)	= NULL,
@BreakInTime_Final		varchar(100)	= NULL,
@CoolingMethod_Final		varchar(100)	= NULL,
@Volts_Final			numeric(10, 2)	= NULL,
@HP_Final			numeric(10, 2)	= NULL,
@Amps_Final			varchar(100)	= NULL,
@Phase_Final			varchar(1)	= NULL,
@Hz_Final			numeric(10, 2)	= NULL,
@Thermistor_Final		varchar(100)	= NULL,
@Poles_Final			numeric(10, 2)	= NULL,
@AmpDraw_Final			varchar(100)	= NULL,
@ConnectorCtrl_Final		varchar(100)	= NULL,
@ConnectorPower_Final		varchar(100)	= NULL,
@Converter_Final		varchar(100)	= NULL,
@ToolHolder_Final		varchar(100)	= NULL,
@PullPin_Final			varchar(100)	= NULL,
@EMDimension_Final		varchar(100)	= NULL,
@EjectionPath_Final		varchar(100)	= NULL,
@ToolOutPressure_Final		varchar(100)	= NULL,
@ReturnPressure_Final		varchar(100)	= NULL,
@DrawbarForce_Final		varchar(100)	= NULL,
@ToolChangeFunction_Final	varchar(100)	= NULL,
@ProximitySwitchFunction_Final	varchar(100)	= NULL,
@Lubrication_Final		varchar(100)	= NULL,
@Grease_Final			varchar(100)	= NULL,
@OilMist_Final			varchar(100)	= NULL,
@OilJet_Final			varchar(100)	= NULL,
@OilGreaseType_Final		varchar(100)	= NULL,
@IntervalDPM_Final		varchar(100)	= NULL,
@MainPressure_Final		varchar(100)	= NULL,
@TubePressure_Final		varchar(100)	= NULL,
@Preload_Final			varchar(100)	= NULL,
@RadialPlay_Final		varchar(100)	= NULL,
@AxialPlay_Final		varchar(100)	= NULL,
@RunoutFront_Final		varchar(100)	= NULL,
@RunoutFrontLocation_Final	varchar(100)	= NULL,
@RunoutFront2_Final		varchar(100)	= NULL,
@RunoutFront2Location_Final	varchar(100)	= NULL,
@RunoutRear_Final		varchar(100)	= NULL,
@RunoutRearLocation_Final	varchar(100)	= NULL,
@RunoutRear2_Final		varchar(100)	= NULL,
@RunoutRear2Location_Final	varchar(100)	= NULL,
@ToolContact_Final		varchar(100)	= NULL,
@ToolContactRear_Final		varchar(100)	= NULL,
@ToolGap_Final			varchar(100)	= NULL,
@ToolGapRear_Final		varchar(100)	= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int
DECLARE @QCInspectionId int

SELECT @QCInspectionId = QCInspectionId FROM tblQCInspections WHERE WorkOrderId = @WorkOrderId
IF( DATALENGTH( @QCInspectionId ) = 0 ) OR ( @QCInspectionId IS NULL ) OR ( @QCInspectionId = 0 )
BEGIN

	INSERT tblQCInspections( WorkOrderId,
RPM_In,
Weight_In,
Vibration_In,
VibrationRear_In,
CoolantFlow_In,
GSE_In,
GSERear_In,
ShaftTemp_In,
CoolantTempIncomingSet_In,
CoolantTempActual_In,
FrontTemp_In,
RearTemp_In,
CoolantPressureIncoming_In,
BreakInTime_In,
CoolingMethod_In,
Volts_In,
HP_In,
Amps_In,
Phase_In,
Hz_In,
Thermistor_In,
Poles_In,
AmpDraw_In,
ConnectorCtrl_In,
ConnectorPower_In,
Converter_In,
ToolHolder_In,
PullPin_In,
EMDimension_In,
EjectionPath_In,
ToolOutPressure_In,
ReturnPressure_In,
DrawbarForce_In,
ToolChangeFunction_In,
ProximitySwitchFunction_In,
Lubrication_In,
Grease_In,
OilMist_In,
OilJet_In,
OilGreaseType_In,
IntervalDPM_In,
MainPressure_In,
TubePressure_In,
Preload_In,
RadialPlay_In,
AxialPlay_In,
RunoutFront_In,
RunoutFrontLocation_In,
RunoutFront2_In,
RunoutFront2Location_In,
RunoutRear_In,
RunoutRearLocation_In,
RunoutRear2_In,
RunoutRear2Location_In,
ToolContact_In,
ToolContactRear_In,
ToolGap_In,
ToolGapRear_In,

RPM_Final,
Weight_Final,
Vibration_Final,
VibrationRear_Final,
CoolantFlow_Final,
GSE_Final,
GSERear_Final,
ShaftTemp_Final,
CoolantTempIncomingSet_Final,
CoolantTempActual_Final,
FrontTemp_Final,
RearTemp_Final,
CoolantPressureIncoming_Final,
BreakInTime_Final,
CoolingMethod_Final,
Volts_Final,
HP_Final,
Amps_Final,
Phase_Final,
Hz_Final,
Thermistor_Final,
Poles_Final,
AmpDraw_Final,
ConnectorCtrl_Final,
ConnectorPower_Final,
Converter_Final,
ToolHolder_Final,
PullPin_Final,
EMDimension_Final,
EjectionPath_Final,
ToolOutPressure_Final,
ReturnPressure_Final,
DrawbarForce_Final,
ToolChangeFunction_Final,
ProximitySwitchFunction_Final,
Lubrication_Final,
Grease_Final,
OilMist_Final,
OilJet_Final,
OilGreaseType_Final,
IntervalDPM_Final,
MainPressure_Final,
TubePressure_Final,
Preload_Final,
RadialPlay_Final,
AxialPlay_Final,
RunoutFront_Final,
RunoutFrontLocation_Final,
RunoutFront2_Final,
RunoutFront2Location_Final,
RunoutRear_Final,
RunoutRearLocation_Final,
RunoutRear2_Final,
RunoutRear2Location_Final,
ToolContact_Final,
ToolContactRear_Final,
ToolGap_Final,
ToolGapRear_Final
) 
VALUES ( @WorkOrderId,

@RPM_In,
@Weight_In,
@Vibration_In,
@VibrationRear_In,
@CoolantFlow_In,
@GSE_In,
@GSERear_In,
@ShaftTemp_In,
@CoolantTempIncomingSet_In,
@CoolantTempActual_In,
@FrontTemp_In,
@RearTemp_In,
@CoolantPressureIncoming_In,
@BreakInTime_In,
@CoolingMethod_In,
@Volts_In,
@HP_In,
@Amps_In,
@Phase_In,
@Hz_In,
@Thermistor_In,
@Poles_In,
@AmpDraw_In,
@ConnectorCtrl_In,
@ConnectorPower_In,
@Converter_In,
@ToolHolder_In,
@PullPin_In,
@EMDimension_In,
@EjectionPath_In,
@ToolOutPressure_In,
@ReturnPressure_In,
@DrawbarForce_In,
@ToolChangeFunction_In,
@ProximitySwitchFunction_In,
@Lubrication_In,
@Grease_In,
@OilMist_In,
@OilJet_In,
@OilGreaseType_In,
@IntervalDPM_In,
@MainPressure_In,
@TubePressure_In,
@Preload_In,
@RadialPlay_In,
@AxialPlay_In,
@RunoutFront_In,
@RunoutFrontLocation_In,
@RunoutFront2_In,
@RunoutFront2Location_In,
@RunoutRear_In,
@RunoutRearLocation_In,
@RunoutRear2_In,
@RunoutRear2Location_In,
@ToolContact_In,
@ToolContactRear_In,
@ToolGap_In,
@ToolGapRear_In,

@RPM_Final,
@Weight_Final,
@Vibration_Final,
@VibrationRear_Final,
@CoolantFlow_Final,
@GSE_Final,
@GSERear_Final,
@ShaftTemp_Final,
@CoolantTempIncomingSet_Final,
@CoolantTempActual_Final,
@FrontTemp_Final,
@RearTemp_Final,
@CoolantPressureIncoming_Final,
@BreakInTime_Final,
@CoolingMethod_Final,
@Volts_Final,
@HP_Final,
@Amps_Final,
@Phase_Final,
@Hz_Final,
@Thermistor_Final,
@Poles_Final,
@AmpDraw_Final,
@ConnectorCtrl_Final,
@ConnectorPower_Final,
@Converter_Final,
@ToolHolder_Final,
@PullPin_Final,
@EMDimension_Final,
@EjectionPath_Final,
@ToolOutPressure_Final,
@ReturnPressure_Final,
@DrawbarForce_Final,
@ToolChangeFunction_Final,
@ProximitySwitchFunction_Final,
@Lubrication_Final,
@Grease_Final,
@OilMist_Final,
@OilJet_Final,
@OilGreaseType_Final,
@IntervalDPM_Final,
@MainPressure_Final,
@TubePressure_Final,
@Preload_Final,
@RadialPlay_Final,
@AxialPlay_Final,
@RunoutFront_Final,
@RunoutFrontLocation_Final,
@RunoutFront2_Final,
@RunoutFront2Location_Final,
@RunoutRear_Final,
@RunoutRearLocation_Final,
@RunoutRear2_Final,
@RunoutRear2Location_Final,
@ToolContact_Final,
@ToolContactRear_Final,
@ToolGap_Final,
@ToolGapRear_Final

 ) 

	SET @ErrorNum	= @@ERROR
	SET @QCInspectionId = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @WorkOrderId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT QCInspectionId FROM tblQCInspections WHERE WorkOrderId = @WorkOrderId ) )
	BEGIN
		UPDATE tblQCInspections SET WorkOrderId = @WorkOrderId,
RPM_In = @RPM_In,
Weight_In = @Weight_In,
Vibration_In = @Vibration_In,
VibrationRear_In = @VibrationRear_In,
CoolantFlow_In = @CoolantFlow_In,
GSE_In = @GSE_In,
GSERear_In = @GSERear_In,
ShaftTemp_In = @ShaftTemp_In,
CoolantTempIncomingSet_In = @CoolantTempIncomingSet_In,
CoolantTempActual_In = @CoolantTempActual_In,
FrontTemp_In = @FrontTemp_In,
RearTemp_In = @RearTemp_In,
CoolantPressureIncoming_In = @CoolantPressureIncoming_In,
BreakInTime_In = @BreakInTime_In,
CoolingMethod_In = @CoolingMethod_In,
Volts_In = @Volts_In,
HP_In = @HP_In,
Amps_In = @Amps_In,
Phase_In = @Phase_In,
Hz_In = @Hz_In,
Thermistor_In = @Thermistor_In,
Poles_In = @Poles_In,
AmpDraw_In = @AmpDraw_In,
ConnectorCtrl_In = @ConnectorCtrl_In,
ConnectorPower_In = @ConnectorPower_In,
Converter_In = @Converter_In,
ToolHolder_In = @ToolHolder_In,
PullPin_In = @PullPin_In,
EMDimension_In = @EMDimension_In,
EjectionPath_In = @EjectionPath_In,
ToolOutPressure_In = @ToolOutPressure_In,
ReturnPressure_In = @ReturnPressure_In,
DrawbarForce_In = @DrawbarForce_In,
ToolChangeFunction_In = @ToolChangeFunction_In,
ProximitySwitchFunction_In = @ProximitySwitchFunction_In,
Lubrication_In = @Lubrication_In,
Grease_In = @Grease_In,
OilMist_In = @OilMist_In,
OilJet_In = @OilJet_In,
OilGreaseType_In = @OilGreaseType_In,
IntervalDPM_In = @IntervalDPM_In,
MainPressure_In = @MainPressure_In,
TubePressure_In = @TubePressure_In,
Preload_In = @Preload_In,
RadialPlay_In = @RadialPlay_In,
AxialPlay_In = @AxialPlay_In,
RunoutFront_In = @RunoutFront_In,
RunoutFrontLocation_In = @RunoutFrontLocation_In,
RunoutFront2_In = @RunoutFront2_In,
RunoutFront2Location_In = @RunoutFront2Location_In,
RunoutRear_In = @RunoutRear_In,
RunoutRearLocation_In = @RunoutRearLocation_In,
RunoutRear2_In = @RunoutRear2_In,
RunoutRear2Location_In = @RunoutRear2Location_In,
ToolContact_In = @ToolContact_In,
ToolContactRear_In = @ToolContactRear_In,
ToolGap_In = @ToolGap_In,
ToolGapRear_In = @ToolGapRear_In,
RPM_Final = @RPM_Final,
Weight_Final = @Weight_Final,
Vibration_Final = @Vibration_Final,
VibrationRear_Final = @VibrationRear_Final,
CoolantFlow_Final = @CoolantFlow_Final,
GSE_Final = @GSE_Final,
GSERear_Final = @GSERear_Final,
ShaftTemp_Final = @ShaftTemp_Final,
CoolantTempIncomingSet_Final = @CoolantTempIncomingSet_Final,
CoolantTempActual_Final = @CoolantTempActual_Final,
FrontTemp_Final = @FrontTemp_Final,
RearTemp_Final = @RearTemp_Final,
CoolantPressureIncoming_Final = @CoolantPressureIncoming_Final,
BreakInTime_Final = @BreakInTime_Final,
CoolingMethod_Final = @CoolingMethod_Final,
Volts_Final = @Volts_Final,
HP_Final = @HP_Final,
Amps_Final = @Amps_Final,
Phase_Final = @Phase_Final,
Hz_Final = @Hz_Final,
Thermistor_Final = @Thermistor_Final,
Poles_Final = @Poles_Final,
AmpDraw_Final = @AmpDraw_Final,
ConnectorCtrl_Final = @ConnectorCtrl_Final,
ConnectorPower_Final = @ConnectorPower_Final,
Converter_Final = @Converter_Final,
ToolHolder_Final = @ToolHolder_Final,
PullPin_Final = @PullPin_Final,
EMDimension_Final = @EMDimension_Final,
EjectionPath_Final = @EjectionPath_Final,
ToolOutPressure_Final = @ToolOutPressure_Final,
ReturnPressure_Final = @ReturnPressure_Final,
DrawbarForce_Final = @DrawbarForce_Final,
ToolChangeFunction_Final = @ToolChangeFunction_Final,
ProximitySwitchFunction_Final = @ProximitySwitchFunction_Final,
Lubrication_Final = @Lubrication_Final,
Grease_Final = @Grease_Final,
OilMist_Final = @OilMist_Final,
OilJet_Final = @OilJet_Final,
OilGreaseType_Final = @OilGreaseType_Final,
IntervalDPM_Final = @IntervalDPM_Final,
MainPressure_Final = @MainPressure_Final,
TubePressure_Final = @TubePressure_Final,
Preload_Final = @Preload_Final,
RadialPlay_Final = @RadialPlay_Final,
AxialPlay_Final = @AxialPlay_Final,
RunoutFront_Final = @RunoutFront_Final,
RunoutFrontLocation_Final = @RunoutFrontLocation_Final,
RunoutFront2_Final = @RunoutFront2_Final,
RunoutFront2Location_Final = @RunoutFront2Location_Final,
RunoutRear_Final = @RunoutRear_Final,
RunoutRearLocation_Final = @RunoutRearLocation_Final,
RunoutRear2_Final = @RunoutRear2_Final,
RunoutRear2Location_Final = @RunoutRear2Location_Final,
ToolContact_Final = @ToolContact_Final,
ToolContactRear_Final = @ToolContactRear_Final,
ToolGap_Final = @ToolGap_Final,
ToolGapRear_Final = @ToolGapRear_Final



		WHERE QCInspectionId = @QCInspectionId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'QCInspection record not found'
		RETURN -1
	END
END

RETURN @QCInspectionId

GO
/****** Object:  StoredProcedure [dbo].[spEditQCTemp]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQCTemp]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQCTemp] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spEditQCTemp]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QCTempId		int		= NULL,
@QCTempLogId		int		= NULL,
@QCTime			varchar(100)	= NULL,
@QCSpeed		varchar(100)	= NULL,
@QCFront		varchar(100)	= NULL,
@QCRear			varchar(100)	= NULL,
@QCShaft		varchar(100)	= NULL,
@QCTempLocation		varchar(100)	= NULL,
@QCNotes		text		= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @QCTempId ) = 0 ) OR ( @QCTempId IS NULL ) OR ( @QCTempId = 0 )
BEGIN

	INSERT tblQCTemps ( QCTempLogId, QCTime, QCSpeed, QCFront, QCRear, QCShaft, QCTempLocation, QCNotes ) 
	VALUES ( @QCTempLogId, @QCTime, @QCSpeed, @QCFront, @QCRear, @QCShaft, @QCTempLocation, @QCNotes ) 

	SET @ErrorNum	= @@ERROR
	SET @QCTempId = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @QCTempId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT QCTempId FROM tblQCTemps WHERE QCTempId = @QCTempId ) )
	BEGIN
		UPDATE tblQCTemps SET 
			QCTempLogId	= @QCTempLogId, 
			QCTime		= @QCTime, 
			QCSpeed		= @QCSpeed, 
			QCFront		= @QCFront, 
			QCRear		= @QCRear, 
			QCShaft		= @QCShaft,
			QCTempLocation	= @QCTempLocation, 
			QCNotes		= @QCNotes
		WHERE QCTempId = @QCTempId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'QCTemp record not found'
		RETURN -1
	END
END

RETURN @QCTempId

GO
/****** Object:  StoredProcedure [dbo].[spEditQCTempLog]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQCTempLog]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQCTempLog] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spEditQCTempLog]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QCTempLogId		int		= NULL,
@WorkOrderId		int		= NULL,
@QCDate			smalldatetime	= NULL,
@QCMaxSpeed		varchar(100)	= NULL,
@QCTotalRunTime		varchar(100)	= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @QCTempLogId ) = 0 ) OR ( @QCTempLogId IS NULL ) OR ( @QCTempLogId = 0 )
BEGIN

	INSERT tblQCTempLogs( WorkOrderId, QCDate, QCMaxSpeed, QCTotalRunTime ) 
	VALUES ( @WorkOrderId, @QCDate, @QCMaxSpeed, @QCTotalRunTime ) 

	SET @ErrorNum	= @@ERROR
	SET @QCTempLogId = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @WorkOrderId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT QCTempLogId FROM tblQCTempLogs WHERE QCTempLogId = @QCTempLogId ) )
	BEGIN
		UPDATE tblQCTempLogs SET WorkOrderId = @WorkOrderId, QCDate = @QCDate, QCMaxSpeed = @QCMaxSpeed, QCTotalRunTime = @QCTotalRunTime
		WHERE QCTempLogId = @QCTempLogId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'QCTempLog record not found'
		RETURN -1
	END
END

RETURN @QCTempLogId



GO
/****** Object:  StoredProcedure [dbo].[spEditQuote]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQuote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQuote] AS' 
END
GO


ALTER PROCEDURE [dbo].[spEditQuote]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QuoteID		int		= NULL,
@WorkOrderId		int		= NULL,
@WorkOrderNumber	varchar(100)	= NULL,
@SerialNumber		varchar(100)	= NULL,
@DateApproved		smalldatetime	= NULL,
@DateQuoted		smalldatetime,
@Notes			text		= NULL,
@QuoteSpecificComments	text		= NULL,
@DefaultQuoteId		int		= NULL,
@QuoteContactId		int		= NULL,
@DeliveryInformation	varchar(100)	= NULL,
@ExpeditedDeliveryInformation	varchar(100)	= NULL,
@QuotedById		int		= NULL,
@HandlingCharge	numeric(10,2)	= 25.0

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @WorkOrderNumber ) = 0 ) OR ( @WorkOrderNumber IS NULL )
BEGIN
	SELECT @WorkOrderNumber = WorkOrderNumber FROM tblWorkOrders WHERE WorkOrderId = @WorkOrderId
END

IF( DATALENGTH( @QuoteID ) = 0 ) OR ( @QuoteID IS NULL ) OR ( @QuoteID = 0 )
BEGIN
	IF( @DefaultQuoteId <> 0 )
	BEGIN
		INSERT tblQuotes( OriginalQuoteId, WorkOrderId, WorkOrderNumber, SerialNumber, DisassemblyEvaluation, HoursDisassembly, CleanAndInspect,
			HoursCleanAndInspect, InhouseGrinding, GrindingHours, Balancing, BalancingHours, ElectricalWork, ElectricalHours, GreaseBearings,
			GreaseHours, AssemblyAndTest, AssemblyAndTestHours, MscWorkNeeded, MscWorkHours, FreightChargeParts, FreightChargeSubWork,
			FreightLBS, FreightChargeSub, FreightChargeSub1, ExpDeliveryDate, Notes, QuoteSpecificComments, PartsCommission, BearingFreightCharge, BearingCommission,
			SpacerPreparation, SpacerPreparationHours, QuoteContactId, DeliveryInformation, ExpeditedDeliveryInformation, QuotedById )
		SELECT OriginalQuoteId, WorkOrderId, WorkOrderNumber, SerialNumber, DisassemblyEvaluation, HoursDisassembly, CleanAndInspect,
			HoursCleanAndInspect, InhouseGrinding, GrindingHours, Balancing, BalancingHours, ElectricalWork, ElectricalHours, GreaseBearings,
			GreaseHours, AssemblyAndTest, AssemblyAndTestHours, MscWorkNeeded, MscWorkHours, FreightChargeParts, FreightChargeSubWork,
			FreightLBS, FreightChargeSub, FreightChargeSub1, ExpDeliveryDate, Notes, QuoteSpecificComments, PartsCommission, BearingFreightCharge, BearingCommission,
			SpacerPreparation, SpacerPreparationHours, QuoteContactId, DeliveryInformation, ExpeditedDeliveryInformation, QuotedById FROM tblQuotes
		WHERE QuoteID = @DefaultQuoteId

		SET @ErrorNum	= @@ERROR
		SET @QuoteID 	= @@IDENTITY

		INSERT tblQuoteParts( QuoteId, ProductId, PartCost, Markup, SupplierId, Qty )
		SELECT @QuoteID, ProductId, PartCost, Markup, SupplierId, Qty FROM tblQuoteParts
		WHERE QuoteId = @DefaultQuoteId

		INSERT tblQuoteSubWork( QuoteId, SubWorkCost, SupplierId, SubWorkDescription )
		SELECT @QuoteID, SubWorkCost, SupplierId, SubWorkDescription FROM tblQuoteSubWork
		WHERE QuoteId = @DefaultQuoteId

		INSERT tblQuoteBearings( QuoteId, ProductId, CenBearingCode, CssPrice, BearingCost, BearingMarkup, SupplierId, BearingDescription, Qty )
		SELECT @QuoteID, ProductId, CenBearingCode, CssPrice, BearingCost, BearingMarkup, SupplierId, BearingDescription, Qty FROM tblQuoteBearings
		WHERE QuoteId = @DefaultQuoteId
	END
	ELSE
	BEGIN

		DECLARE @MainContactId AS int
		SELECT  @MainContactId = MainContactId FROM tblWorkOrders
		LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
		LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
		WHERE WorkOrderId = @WorkOrderId

		INSERT tblQuotes( WorkOrderId, WorkOrderNumber, SerialNumber, DateApproved, DateQuoted, Notes, QuoteContactId, DeliveryInformation, ExpeditedDeliveryInformation, QuotedById )
		VALUES( @WorkOrderId, @WorkOrderNumber, @SerialNumber, @DateApproved, @DateQuoted, @Notes, @MainContactId, @DeliveryInformation, @ExpeditedDeliveryInformation, @QuotedById )

		SET @ErrorNum	= @@ERROR
		SET @QuoteID 	= @@IDENTITY
	END


	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @QuoteID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT QuoteId FROM tblQuotes WHERE QuoteId = @QuoteId ) )
	BEGIN
		IF( EXISTS( SELECT WorkOrderId FROM tblQuotes WHERE QuoteId = @QuoteId ) )
		BEGIN
			UPDATE tblQuotes SET WorkOrderId = @WorkOrderId, WorkOrderNumber = @WorkOrderNumber, SerialNumber = @SerialNumber, DateApproved = @DateApproved, DateQuoted = @DateQuoted, Notes = @Notes, QuoteSpecificComments = @QuoteSpecificComments,  QuoteContactId = @QuoteContactId, DeliveryInformation = @DeliveryInformation, ExpeditedDeliveryInformation = @ExpeditedDeliveryInformation, QuotedById = @QuotedById, HandlingCharge = @HandlingCharge
			WHERE QuoteId = @QuoteId
		END
		ELSE
		BEGIN
			UPDATE tblQuotes SET WorkOrderId = @WorkOrderId, WorkOrderNumber = @WorkOrderNumber, SerialNumber = @SerialNumber, DateApproved = @DateApproved, DateQuoted = @DateQuoted, Notes = @Notes, QuoteSpecificComments = @QuoteSpecificComments, QuoteContactId = @QuoteContactId, DeliveryInformation = @DeliveryInformation, ExpeditedDeliveryInformation = @ExpeditedDeliveryInformation, QuotedById = @QuotedById, HandlingCharge = @HandlingCharge
			WHERE QuoteId = @QuoteId
		END

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Quote record not found'
		RETURN -1
	END
END

RETURN @QuoteID
GO
/****** Object:  StoredProcedure [dbo].[spEditQuoteBearing]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQuoteBearing]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQuoteBearing] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditQuoteBearing]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QuoteID		int		= NULL,
@BearingCommission	numeric(10,2)	= NULL,
@BearingFreightCharge	numeric(10,2)	= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( EXISTS( SELECT QuoteId FROM tblQuotes WHERE QuoteId = @QuoteId ) )
BEGIN
	UPDATE tblQuotes SET BearingCommission = @BearingCommission, BearingFreightCharge = @BearingFreightCharge
	WHERE QuoteId = @QuoteId

	SET @ErrorNum = @@ERROR
	
	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		RETURN -1
	END
END
ELSE
BEGIN
	SELECT @ErrorMessage = 'Quote record not found'
	RETURN -1
END

RETURN @QuoteID



GO
/****** Object:  StoredProcedure [dbo].[spEditQuoteBearingDetail]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQuoteBearingDetail]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQuoteBearingDetail] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditQuoteBearingDetail]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QuoteId		int,
@QuoteBearingId		int		= NULL,
@ProductId		int,
@BearingCost		numeric(10,2)	= NULL,
@Markup			numeric(10,2)	= 1.0,
@SupplierId		int,
@Qty			numeric(10,2)	= 1

AS

SET NOCOUNT ON

DECLARE @ErrorNum		int
DECLARE @BearingDescription	varchar(255)

SELECT @BearingDescription = PartNumber FROM tblProducts WHERE ProductId = @ProductId

IF( DATALENGTH( @QuoteBearingId ) = 0 ) OR ( @QuoteBearingId IS NULL ) OR ( @QuoteBearingId = 0 )
BEGIN
	INSERT tblQuoteBearings( QuoteId, ProductId, BearingCost, BearingMarkup, SupplierId, BearingDescription, Qty )
	VALUES ( @QuoteId, @ProductId, @BearingCost, @Markup, @SupplierId, @BearingDescription, @Qty )

	SET @ErrorNum		= @@ERROR
	SET @QuoteBearingId	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @QuoteBearingId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT QuoteBearingId FROM tblQuoteBearings WHERE QuoteBearingId = @QuoteBearingId ) )
	BEGIN

		UPDATE tblQuoteBearings SET QuoteId = @QuoteId, ProductId = @ProductId, SupplierId = @SupplierId, BearingDescription = @BearingDescription,
		BearingCost = @BearingCost, BearingMarkup = @Markup, Qty = @Qty
		WHERE QuoteBearingId = @QuoteBearingId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Quote bearing detail record not found'
		RETURN -1
	END
END

RETURN @QuoteBearingId

GO
/****** Object:  StoredProcedure [dbo].[spEditQuoteLabor]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQuoteLabor]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQuoteLabor] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditQuoteLabor]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QuoteID		int		= NULL,
@DisassemblyEvaluation	varchar(100)	= NULL,
@HoursDisassembly	numeric(10,2)	= NULL,
@CleanAndInspect	varchar(100)	= NULL,
@HoursCleanAndInspect	numeric(10,2)	= NULL,
@InhouseGrinding	varchar(100)	= NULL,
@GrindingHours		numeric(10,2)	= NULL,
@Balancing		varchar(100)	= NULL,
@BalancingHours		numeric(10,2)	= NULL,
@ElectricalWork		varchar(100)	= NULL,
@ElectricalHours	numeric(10,2)	= NULL,
@GreaseBearings		varchar(100)	= NULL,
@GreaseHours		numeric(10,2)	= NULL,
@AssemblyAndTest	varchar(100)	= NULL,
@AssemblyAndTestHours	numeric(10,2)	= NULL,
@MscWorkNeeded		text		= NULL,
@MscWorkHours		numeric(10,2)	= NULL,
@LaborCommission	numeric(10,2)	= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( EXISTS( SELECT QuoteId FROM tblQuotes WHERE QuoteId = @QuoteId ) )
BEGIN
	UPDATE tblQuotes SET DisassemblyEvaluation = @DisassemblyEvaluation, HoursDisassembly = @HoursDisassembly, CleanAndInspect = @CleanAndInspect, 
	HoursCleanAndInspect = @HoursCleanAndInspect, InhouseGrinding = @InhouseGrinding, GrindingHours = @GrindingHours, Balancing = @Balancing,
	BalancingHours = @BalancingHours, ElectricalWork = @ElectricalWork, ElectricalHours = @ElectricalHours,
	GreaseBearings = @GreaseBearings, GreaseHours = @GreaseHours, AssemblyAndTest = @AssemblyAndTest, 
	AssemblyAndTestHours = @AssemblyAndTestHours, MscWorkNeeded = @MscWorkNeeded, MscWorkHours = @MscWorkHours, LaborCommission = @LaborCommission
	WHERE QuoteId = @QuoteId

	SET @ErrorNum = @@ERROR

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		RETURN -1
	END
END
ELSE
BEGIN
	SELECT @ErrorMessage = 'Quote record not found'
	RETURN -1
END

RETURN @QuoteID


GO
/****** Object:  StoredProcedure [dbo].[spEditQuotePart]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQuotePart]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQuotePart] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditQuotePart]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QuoteID		int		= NULL,
@PartsCommission	numeric(10,2)	= NULL,
@FreightChargeParts	numeric(10,2)	= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( EXISTS( SELECT QuoteId FROM tblQuotes WHERE QuoteId = @QuoteId ) )
BEGIN
	UPDATE tblQuotes SET PartsCommission = @PartsCommission, FreightChargeParts = @FreightChargeParts
	WHERE QuoteId = @QuoteId

	SET @ErrorNum = @@ERROR
	
	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		RETURN -1
	END
END
ELSE
BEGIN
	SELECT @ErrorMessage = 'Quote record not found'
	RETURN -1
END

RETURN @QuoteID



GO
/****** Object:  StoredProcedure [dbo].[spEditQuotePartDetail]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQuotePartDetail]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQuotePartDetail] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditQuotePartDetail]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QuoteId		int,
@QuotePartId		int		= NULL,
@ProductId		int,
@PartCost		numeric(10,2)	= NULL,
@Markup			numeric(10,2)	= 1.0,
@SupplierId		int,
@Qty			numeric(10,2)	= 1

AS

SET NOCOUNT ON

DECLARE @ErrorNum		int

IF( DATALENGTH( @QuotePartId ) = 0 ) OR ( @QuotePartId IS NULL ) OR ( @QuotePartId = 0 )
BEGIN
	INSERT tblQuoteParts( QuoteId, ProductId, PartCost, Markup, SupplierId, Qty )
	VALUES ( @QuoteId, @ProductId, @PartCost, @Markup, @SupplierId, @Qty )

	SET @ErrorNum		= @@ERROR
	SET @QuotePartId	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @QuotePartId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT QuotePartId FROM tblQuoteParts WHERE QuotePartId = @QuotePartId ) )
	BEGIN
		UPDATE tblQuoteParts SET QuoteId = @QuoteId, ProductId = @ProductId, PartCost = @PartCost, Markup = @Markup, 
		SupplierId = @SupplierId, Qty = @Qty
		WHERE QuotePartId = @QuotePartId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Quote part detail record not found'
		RETURN -1
	END
END

RETURN @QuotePartId

GO
/****** Object:  StoredProcedure [dbo].[spEditQuoteSubWork]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQuoteSubWork]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQuoteSubWork] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditQuoteSubWork]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QuoteID		int		= NULL,
@FreightLBS		varchar(100)	= NULL,
@FreightChargeSub	numeric(10, 2)	= NULL,
@ExpDeliveryDate	smalldatetime	= NULL,
@SubWorkCommission	numeric(10,2)	= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( EXISTS( SELECT QuoteId FROM tblQuotes WHERE QuoteId = @QuoteId ) )
BEGIN
	UPDATE tblQuotes SET FreightLBS = @FreightLBS, FreightChargeSub = @FreightChargeSub, ExpDeliveryDate = @ExpDeliveryDate, SubWorkCommission = @SubWorkCommission
	WHERE QuoteId = @QuoteId

	SET @ErrorNum = @@ERROR
	
	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		RETURN -1
	END
END
ELSE
BEGIN
	SELECT @ErrorMessage = 'Quote record not found'
	RETURN -1
END

RETURN @QuoteID



GO
/****** Object:  StoredProcedure [dbo].[spEditQuoteSubWorkDetail]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditQuoteSubWorkDetail]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditQuoteSubWorkDetail] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditQuoteSubWorkDetail]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@QuoteId		int,
@QuoteSubWorkId		int		= NULL,
@SubWorkCost		numeric(10,2)	= NULL,
@SupplierId		int,
@SubWorkDescription	text

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @QuoteSubWorkId ) = 0 ) OR ( @QuoteSubWorkId IS NULL ) OR ( @QuoteSubWorkId = 0 )
BEGIN

	INSERT tblQuoteSubWork( QuoteId, SubWorkCost, SupplierId, SubWorkDescription )
	VALUES ( @QuoteId, @SubWorkCost, @SupplierId, @SubWorkDescription )

	SET @ErrorNum		= @@ERROR
	SET @QuoteSubWorkId	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @QuoteSubWorkId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT QuoteSubWorkId FROM tblQuoteSubWork WHERE QuoteSubWorkId = @QuoteSubWorkId ) )
	BEGIN

		UPDATE tblQuoteSubWork SET QuoteId = @QuoteId, SubWorkCost = @SubWorkCost, SupplierId = @SupplierId, SubWorkDescription = @SubWorkDescription
		WHERE QuoteSubWorkId = @QuoteSubWorkId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Quote sub work detail record not found'
		RETURN -1
	END
END

RETURN @QuoteSubWorkId


GO
/****** Object:  StoredProcedure [dbo].[spEditShippingMethod]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditShippingMethod]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditShippingMethod] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditShippingMethod]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@ShippingMethodID	int		= NULL,
@ShippingMethod		varchar(100)

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @ShippingMethodID ) = 0 ) OR ( @ShippingMethodID IS NULL ) OR ( @ShippingMethodID = 0 )
BEGIN
	INSERT tblShippingMethods( ShippingMethod ) VALUES ( @ShippingMethod )

	SET @ErrorNum	= @@ERROR
	SET @ShippingMethodID = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @ShippingMethodID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT ShippingMethodID FROM tblShippingMethods WHERE ShippingMethodID = @ShippingMethodID ) )
	BEGIN
		UPDATE tblShippingMethods SET ShippingMethod = @ShippingMethod
		WHERE ShippingMethodID = @ShippingMethodID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Shipping method record not found'
		RETURN -1
	END
END

RETURN @ShippingMethodID


GO
/****** Object:  StoredProcedure [dbo].[spEditSpindle]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditSpindle]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditSpindle] AS' 
END
GO



ALTER   PROCEDURE [dbo].[spEditSpindle]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@SpindleId		int		= NULL,
@SpindleType		varchar(100)	= NULL,
@SpindleCategoryId	int		= NULL,
@DrawingNumber		varchar(100)	= NULL,
@RPM			numeric(10, 2)	= NULL,
@Weight			varchar(100)	= NULL,
@Vibration		varchar(100)	= NULL,
@VibrationRear		varchar(100)	= NULL,
@CoolantFlow		varchar(100)	= NULL,
@GSE			varchar(100)	= NULL,
@GSERear		varchar(100)	= NULL,
@ShaftTemp		varchar(100)	= NULL,
@CoolantTempIncomingSet	varchar(100)	= NULL,
@CoolantTempActual	varchar(100)	= NULL,
@FrontTemp		varchar(100)	= NULL,
@RearTemp		varchar(100)	= NULL,
@CoolantPressureIncoming	varchar(100)	= NULL,
@BreakInTime		varchar(100)	= NULL,
@CoolingMethod		varchar(100)	= NULL,
@Volts			numeric(10, 2)	= NULL,
@HP			numeric(10, 2)	= NULL,
@Amps			varchar(100)	= NULL,
@Phase			varchar(1)	= NULL,
@Hz			numeric(10, 2)	= NULL,
@Thermistor		varchar(100)	= NULL,
@Poles			numeric(10, 2)	= NULL,
@AmpDraw		varchar(100)	= NULL,
@ConnectorCtrl		varchar(100)	= NULL,
@ConnectorPower		varchar(100)	= NULL,
@Converter		varchar(100)	= NULL,
@ToolHolder		varchar(100)	= NULL,
@PullPin		varchar(100)	= NULL,
@EMDimension		varchar(100)	= NULL,
@EjectionPath		varchar(100)	= NULL,
@ToolOutPressure	varchar(100)	= NULL,
@ReturnPressure		varchar(100)	= NULL,
@DrawbarForce		varchar(100)	= NULL,
@ToolChangeFunction	varchar(100)	= NULL,
@ProximitySwitchFunction	varchar(100)	= NULL,
@Lubrication		varchar(100)	= NULL,
@Grease			varchar(100)	= NULL,
@OilMist		varchar(100)	= NULL,
@OilJet			varchar(100)	= NULL,
@OilGreaseType		varchar(100)	= NULL,
@IntervalDPM		varchar(100)	= NULL,
@MainPressure		varchar(100)	= NULL,
@TubePressure		varchar(100)	= NULL,
@LubeNotes		text		= NULL,
@Preload		varchar(100)	= NULL,
@RadialPlay		varchar(100)	= NULL,
@AxialPlay		varchar(100)	= NULL,
@RunoutFront		varchar(100)	= NULL,
@RunoutFrontLocation	varchar(100)	= NULL,
@RunoutFront2		varchar(100)	= NULL,
@RunoutFront2Location	varchar(100)	= NULL,
@RunoutRear		varchar(100)	= NULL,
@RunoutRearLocation	varchar(100)	= NULL,
@RunoutRear2		varchar(100)	= NULL,
@RunoutRear2Location	varchar(100)	= NULL,
@ToolContact		varchar(100)	= NULL,
@ToolContactRear	varchar(100)	= NULL,
@ToolGap		varchar(100)	= NULL,
@ToolGapRear		varchar(100)	= NULL,
@Other			text		= NULL,
@BalancingRequirements	text		= NULL,
@BearingInformation	text		= NULL,
@GeneralSpindleNotes	text		= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @SpindleId ) = 0 ) OR ( @SpindleId IS NULL ) OR ( @SpindleId = 0 )
BEGIN
	INSERT tblSpindles( 
SpindleType,
SpindleCategoryId,
DrawingNumber,
RPM,
Weight,
Vibration,
VibrationRear,
CoolantFlow,
GSE,
GSERear,
ShaftTemp,
CoolantTempIncomingSet,
CoolantTempActual,
FrontTemp,
RearTemp,
CoolantPressureIncoming,
BreakInTime,
CoolingMethod,
Volts,
HP,
Amps,
Phase,
Hz,
Thermistor,
Poles,
AmpDraw,
ConnectorCtrl,
ConnectorPower,
Converter,
ToolHolder,
PullPin,
EMDimension,
EjectionPath,
ToolOutPressure,
ReturnPressure,
DrawbarForce,
ToolChangeFunction,
ProximitySwitchFunction,
Lubrication,
Grease,
OilMist,
OilJet,
OilGreaseType,
IntervalDPM,
MainPressure,
TubePressure,
LubeNotes,
Preload,
RadialPlay,
AxialPlay,
RunoutFront,
RunoutFrontLocation,
RunoutFront2,
RunoutFront2Location,
RunoutRear,
RunoutRearLocation,
RunoutRear2,
RunoutRear2Location,
ToolContact,
ToolContactRear,
ToolGap,
ToolGapRear,
Other,
BalancingRequirements,
BearingInformation,
GeneralSpindleNotes
 ) 
VALUES( 
@SpindleType,
@SpindleCategoryId,
@DrawingNumber,
@RPM,
@Weight,
@Vibration,
@VibrationRear,
@CoolantFlow,
@GSE,
@GSERear,
@ShaftTemp,
@CoolantTempIncomingSet,
@CoolantTempActual,
@FrontTemp,
@RearTemp,
@CoolantPressureIncoming,
@BreakInTime,
@CoolingMethod,
@Volts,
@HP,
@Amps,
@Phase,
@Hz,
@Thermistor,
@Poles,
@AmpDraw,
@ConnectorCtrl,
@ConnectorPower,
@Converter,
@ToolHolder,
@PullPin,
@EMDimension,
@EjectionPath,
@ToolOutPressure,
@ReturnPressure,
@DrawbarForce,
@ToolChangeFunction,
@ProximitySwitchFunction,
@Lubrication,
@Grease,
@OilMist,
@OilJet,
@OilGreaseType,
@IntervalDPM,
@MainPressure,
@TubePressure,
@LubeNotes,
@Preload,
@RadialPlay,
@AxialPlay,
@RunoutFront,
@RunoutFrontLocation,
@RunoutFront2,
@RunoutFront2Location,
@RunoutRear,
@RunoutRearLocation,
@RunoutRear2,
@RunoutRear2Location,
@ToolContact,
@ToolContactRear,
@ToolGap,
@ToolGapRear,
@Other,
@BalancingRequirements,
@BearingInformation,
@GeneralSpindleNotes ) 

	SET @ErrorNum	= @@ERROR
	SET @SpindleId = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @SpindleId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT SpindleId FROM tblSpindles WHERE SpindleId = @SpindleId ) )
	BEGIN
		UPDATE tblSpindles SET
SpindleType = @SpindleType,
SpindleCategoryId = @SpindleCategoryId,
DrawingNumber = @DrawingNumber,
RPM = @RPM,
Weight = @Weight,
Vibration = @Vibration,
VibrationRear = @VibrationRear,
CoolantFlow = @CoolantFlow,
GSE = @GSE,
GSERear = @GSERear,
ShaftTemp = @ShaftTemp,
CoolantTempIncomingSet = @CoolantTempIncomingSet,
CoolantTempActual = @CoolantTempActual,
FrontTemp = @FrontTemp,
RearTemp = @RearTemp,
CoolantPressureIncoming = @CoolantPressureIncoming,
BreakInTime = @BreakInTime,
CoolingMethod = @CoolingMethod,
Volts = @Volts,
HP = @HP,
Amps = @Amps,
Phase = @Phase,
Hz = @Hz,
Thermistor = @Thermistor,
Poles = @Poles,
AmpDraw = @AmpDraw,
ConnectorCtrl = @ConnectorCtrl,
ConnectorPower = @ConnectorPower,
Converter = @Converter,
ToolHolder = @ToolHolder,
PullPin = @PullPin,
EMDimension = @EMDimension,
EjectionPath = @EjectionPath,
ToolOutPressure = @ToolOutPressure,
ReturnPressure = @ReturnPressure,
DrawbarForce = @DrawbarForce,
ToolChangeFunction = @ToolChangeFunction,
ProximitySwitchFunction = @ProximitySwitchFunction,
Lubrication = @Lubrication,
Grease = @Grease,
OilMist = @OilMist,
OilJet = @OilJet,
OilGreaseType = @OilGreaseType,
IntervalDPM = @IntervalDPM,
MainPressure = @MainPressure,
TubePressure = @TubePressure,
LubeNotes = @LubeNotes,
Preload = @Preload,
RadialPlay = @RadialPlay,
AxialPlay = @AxialPlay,
RunoutFront = @RunoutFront,
RunoutFrontLocation = @RunoutFrontLocation,
RunoutFront2 = @RunoutFront2,
RunoutFront2Location = @RunoutFront2Location,
RunoutRear = @RunoutRear,
RunoutRearLocation = @RunoutRearLocation,
RunoutRear2 = @RunoutRear2,
RunoutRear2Location = @RunoutRear2Location,
ToolContact = @ToolContact,
ToolContactRear = @ToolContactRear,
ToolGap = @ToolGap,
ToolGapRear = @ToolGapRear,
Other = @Other,
BalancingRequirements = @BalancingRequirements,
BearingInformation = @BearingInformation,
GeneralSpindleNotes = @GeneralSpindleNotes
		WHERE SpindleId = @SpindleId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Spindle record not found'
		RETURN -1
	END
END

RETURN @SpindleId

GO
/****** Object:  StoredProcedure [dbo].[spEditSpindleCategory]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditSpindleCategory]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditSpindleCategory] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditSpindleCategory]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@SpindleCategoryID	int		= NULL,
@SpindleCategory	varchar(100)

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @SpindleCategoryID ) = 0 ) OR ( @SpindleCategoryID IS NULL ) OR ( @SpindleCategoryID = 0 )
BEGIN
	INSERT tblSpindleCategories( SpindleCategoryName ) VALUES ( @SpindleCategory )

	SET @ErrorNum		= @@ERROR
	SET @SpindleCategoryID	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @SpindleCategoryID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT SpindleCategoryID FROM tblSpindleCategories WHERE SpindleCategoryID = @SpindleCategoryID ) )
	BEGIN
		UPDATE tblSpindleCategories SET SpindleCategoryName = @SpindleCategory
		WHERE SpindleCategoryID = @SpindleCategoryID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Spindle category record not found'
		RETURN -1
	END
END

RETURN @SpindleCategoryID


GO
/****** Object:  StoredProcedure [dbo].[spEditSpindleProduct]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditSpindleProduct]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditSpindleProduct] AS' 
END
GO


ALTER PROCEDURE [dbo].[spEditSpindleProduct]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@SpindleProductId	int,
@SpindleId		int		= NULL,
@ProductId		int,
@PartCost		numeric(10,2)	= NULL,
@Markup			numeric(10,2)	= 1.0,
@SupplierId		int,
@Qty			numeric(10,2)	= 1

AS

SET NOCOUNT ON

DECLARE @ErrorNum		int

IF( DATALENGTH( @SpindleProductId ) = 0 ) OR ( @SpindleProductId IS NULL ) OR ( @SpindleProductId = 0 )
BEGIN
	INSERT tblSpindlesProducts( SpindleId, ProductId, Cost, Markup, SupplierId, Quantity )
	VALUES ( @SpindleId, @ProductId, @PartCost, @Markup, @SupplierId, @Qty )

	SET @ErrorNum		= @@ERROR
	SET @SpindleProductId	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @SpindleProductId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT SpindleProductId FROM tblSpindlesProducts WHERE SpindleProductId = @SpindleProductId ) )
	BEGIN
		UPDATE tblSpindlesProducts SET SpindleId = @SpindleId, ProductId = @ProductId, Cost = @PartCost, Markup = @Markup,
		SupplierId = @SupplierId, Quantity = @Qty
		WHERE SpindleProductId = @SpindleProductId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Spindle product detail record not found'
		RETURN -1
	END
END

RETURN @SpindleProductId


GO
/****** Object:  StoredProcedure [dbo].[spEditSpindleSubWork]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditSpindleSubWork]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditSpindleSubWork] AS' 
END
GO


ALTER PROCEDURE [dbo].[spEditSpindleSubWork]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@SpindleSubWorkId	int,
@SpindleId		int		= NULL,
@SubWorkCost		numeric(10,2)	= NULL,
@SupplierId		int,
@SubWorkDesc		text		= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum		int

IF( DATALENGTH( @SpindleSubWorkId ) = 0 ) OR ( @SpindleSubWorkId IS NULL ) OR ( @SpindleSubWorkId = 0 )
BEGIN
	INSERT tblSpindlesSubWork( SpindleId, SubWorkCost, SupplierId, SubWorkDescription )
	VALUES ( @SpindleId, @SubWorkCost, @SupplierId, @SubWorkDesc )

	SET @ErrorNum		= @@ERROR
	SET @SpindleSubWorkId	= @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @SpindleSubWorkId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT SpindleSubWorkId FROM tblSpindlesSubWork WHERE SpindleSubWorkId = @SpindleSubWorkId ) )
	BEGIN
		UPDATE tblSpindlesSubWork SET SpindleId = @SpindleId, SubWorkCost = @SubWorkCost,
		SupplierId = @SupplierId, SubWorkDescription = @SubWorkDesc
		WHERE SpindleSubWorkId = @SpindleSubWorkId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Spindle subwork detail record not found'
		RETURN -1
	END
END

RETURN @SpindleSubWorkId


GO
/****** Object:  StoredProcedure [dbo].[spEditSupplier]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditSupplier]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditSupplier] AS' 
END
GO



ALTER PROCEDURE [dbo].[spEditSupplier]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@SupplierID		int		= NULL,
@SupplierName		varchar(100),
@ContactName		varchar(100)	= NULL,
@ContactTitle		varchar(100)	= NULL,
@Address		varchar(255)	= NULL,
@City			varchar(100)	= NULL,
@PostalCode		varchar(100)	= NULL,
@StateOrProvince	varchar(100)	= NULL,
@Country		varchar(100)	= NULL,
@PhoneNumber		varchar(100)	= NULL,
@FaxNumber		varchar(100)	= NULL 

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @SupplierID ) = 0 ) OR ( @SupplierID IS NULL ) OR ( @SupplierID = 0 )
BEGIN
	INSERT tblSuppliers( SupplierName, ContactName, ContactTitle, Address, City, PostalCode, StateOrProvince, Country, PhoneNumber, FaxNumber ) 
	VALUES( @SupplierName, @ContactName, @ContactTitle, @Address, @City, @PostalCode, @StateOrProvince, @Country, @PhoneNumber, @FaxNumber )

	SET @ErrorNum	= @@ERROR
	SET @SupplierID = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @SupplierID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT SupplierId FROM tblSuppliers WHERE SupplierId = @SupplierID ) )
	BEGIN
		UPDATE tblSuppliers SET SupplierName = @SupplierName, ContactName = @ContactName, ContactTitle = @ContactTitle, Address = @Address, City = @City, 
		PostalCode = @PostalCode, StateOrProvince = @StateOrProvince, Country = @Country, PhoneNumber = @PhoneNumber, FaxNumber = @FaxNumber
		WHERE SupplierId = @SupplierID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Supplier record not found'
		RETURN -1
	END
END

RETURN @SupplierID


GO
/****** Object:  StoredProcedure [dbo].[spEditWorkOrder]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditWorkOrder]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditWorkOrder] AS' 
END
GO



ALTER  PROCEDURE [dbo].[spEditWorkOrder]

@ErrorMessage		varchar(255) 	= NULL OUTPUT,
@WorkOrderId		int		= NULL,
@WorkOrderNumber	int		= NULL,
@ProjectId		int,
@PromiseDate		smalldatetime	= NULL,
@DateIn			smalldatetime	= NULL,
@DateOut		smalldatetime	= NULL,
@SerialNumber		varchar(100)	= NULL,
@NewSpindle		varchar(100)	= NULL,
@PONumber		varchar(100)	= NULL,
@Labor			varchar(100)	= NULL,
@Material		varchar(100)	= NULL,
@Subcontract		varchar(100)	= NULL,
@Cost			numeric(10, 2)	= NULL,
@Charge			numeric(10, 2)	= NULL,
@Date			smalldatetime	= NULL,
@DateExp		smalldatetime	= NULL,
@SalesRep		varchar(100)	= NULL,
@ShippingMethodId	int		= NULL,
@TrackingWaybill	varchar(100)	= NULL,
@LocationId		int		= NULL,
@WorkOrderPriorityId	int		= 1000,
@BoeingSpindleId	int		= NULL,
@BoeingWorkOrderNumber	varchar(100)	= NULL,
@Parts			text		= NULL,
@Bearings		text		= NULL,
@Lube			text		= NULL,
@BalVelocity		varchar(100)	= NULL,
@BalVelocityFinal	varchar(100)	= NULL,
@GSE			varchar(100)	= NULL,
@GSEFinal		varchar(100)	= NULL,
@BreakIn		varchar(100)	= NULL,
@BreakInFinal		varchar(100)	= NULL,
@RoomTemp		varchar(100)	= NULL,
@RoomTempFinal		varchar(100)	= NULL,
@FrontTemp		varchar(100)	= NULL,
@FrontTempFinal		varchar(100)	= NULL,
@RearTemp		varchar(100)	= NULL,
@RearTempFinal		varchar(100)	= NULL,
@Cooling		varchar(100)	= NULL,
@CoolingFinal		varchar(100)	= NULL,
@RunoutFront		varchar(100)	= NULL,
@RunoutFrontFinal	varchar(100)	= NULL,
@Rear			varchar(100)	= NULL,
@RearFinal		varchar(100)	= NULL,
@Other			text		= NULL,
@IncomingInspection	text		= NULL,
@Comments		text		= NULL,
@Remarks		text		= NULL,
@DateRec		smalldatetime	= NULL,
@ExpectedDelDate	smalldatetime	= NULL,
@Commission		varchar(100)	= NULL,
@CommissionReceivedDate	smalldatetime	= NULL,
@AdditionalInfo		text		= NULL,
@ActualDrawforce	varchar(100)	= NULL,
@ActualDrawforceFinal	varchar(100)	= NULL,
@WorkOrderContactId	int		= NULL

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @WorkOrderId ) = 0 ) OR ( @WorkOrderId IS NULL ) OR ( @WorkOrderId = 0 )
BEGIN
	IF( DATALENGTH( @WorkOrderNumber ) = 0 ) OR ( @WorkOrderNumber IS NULL ) OR ( @WorkOrderNumber = 0 )
	BEGIN
		/*Very bad programming method for obtaining next work order number, but since there cannot be two identity columns in one table, what to do?*/
		SELECT @WorkOrderNumber = ( MAX( WorkOrderNumber ) + 1 ) FROM tblWorkOrders
	END

	DECLARE @MainContactId AS int
	SELECT  @MainContactId = MainContactId FROM tblProjects
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	WHERE ProjectId = @ProjectId

	INSERT tblWorkOrders( WorkOrderNumber, ProjectId, PromiseDate, DateIn, DateOut, SerialNumber, NewSpindle, PONumber, Labor, Material,
	Subcontract, Cost, Charge, [Date], DateExp, SalesRep, ShippingMethodId, TrackingWaybill, LocationId, WorkOrderPriorityId, BoeingSpindleId, BoeingWorkOrderNumber, Parts, Bearings, Lube, BalVelocity, BalVelocityFinal, GSE, GSEFinal, BreakIn, BreakInFinal,
	RoomTemp, RoomTempFinal, FrontTemp, FrontTempFinal, RearTemp, RearTempFinal, Cooling, CoolingFinal, RunoutFront, RunoutFrontFinal, Rear, RearFinal, Other, IncomingInspection, Comments, Remarks, DateRec, ExpectedDelDate, Commission,
	CommissionReceivedDate, AdditionalInfo, ActualDrawforce, ActualDrawforceFinal, WorkOrderContactId ) VALUES ( @WorkOrderNumber, @ProjectId, @PromiseDate, @DateIn, @DateOut, @SerialNumber, 
	@NewSpindle, @PONumber, @Labor, @Material, @Subcontract, @Cost, @Charge, @Date, @DateExp, @SalesRep, @ShippingMethodId, @TrackingWaybill, @LocationId, @WorkOrderPriorityId, @BoeingSpindleId, @BoeingWorkOrderNumber, 
	@Parts, @Bearings, @Lube, @BalVelocity, @BalVelocityFinal, @GSE, @GSEFinal, @BreakIn, @BreakInFinal, @RoomTemp, @RoomTempFinal, @FrontTemp, @FrontTempFinal, @RearTemp, @RearTempFinal, @Cooling, @CoolingFinal, @RunoutFront, @RunoutFrontFinal, @Rear, @RearFinal, @Other, @IncomingInspection,
	@Comments, @Remarks, @DateRec, @ExpectedDelDate, @Commission, @CommissionReceivedDate, @AdditionalInfo, @ActualDrawforceFinal, @ActualDrawforceFinal, @MainContactId )

	SET @ErrorNum	= @@ERROR
	SET @WorkOrderId = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @WorkOrderId = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT WorkOrderId FROM tblWorkOrders WHERE WorkOrderId = @WorkOrderId ) )
	BEGIN
		UPDATE tblWorkOrders SET WorkOrderNumber = @WorkOrderNumber, ProjectId = @ProjectId, PromiseDate = @PromiseDate, DateIn = @DateIn, DateOut = @DateOut, 
		SerialNumber = @SerialNumber, NewSpindle = @NewSpindle, PONumber = @PONumber, Labor = @Labor, Material = @Material, Subcontract = @Subcontract,
		Cost = @Cost, Charge = @Charge, [Date] = @Date, DateExp = @DateExp, SalesRep = @SalesRep, ShippingMethodId = @ShippingMethodId, TrackingWaybill = @TrackingWaybill, LocationId = @LocationId, WorkOrderPriorityId = @WorkOrderPriorityId, 
		BoeingSpindleId = @BoeingSpindleId, BoeingWorkOrderNumber = @BoeingWorkOrderNumber, Parts = @Parts, Bearings = @Bearings, Lube = @Lube, BalVelocity = @BalVelocity, BalVelocityFinal = @BalVelocityFinal, GSE = @GSE, GSEFinal = @GSEFinal, BreakIn = @BreakIn, BreakInFinal = @BreakInFinal,
		RoomTemp = @RoomTemp, RoomTempFinal = @RoomTempFinal, FrontTemp = @FrontTemp, FrontTempFinal = @FrontTempFinal, RearTemp = @RearTemp, RearTempFinal = @RearTempFinal, Cooling = @Cooling, CoolingFinal = @CoolingFinal, RunoutFront = @RunoutFront, RunoutFrontFinal = @RunoutFrontFinal, Rear = @Rear, RearFinal = @RearFinal, Other = @Other, 
		IncomingInspection = @IncomingInspection, Comments = @Comments, Remarks = @Remarks, DateRec = @DateRec, ExpectedDelDate = @ExpectedDelDate, Commission = @Commission,
		CommissionReceivedDate = @CommissionReceivedDate, AdditionalInfo = @AdditionalInfo, ActualDrawforce = @ActualDrawforce, ActualDrawforceFinal = @ActualDrawforceFinal, WorkOrderContactId = @WorkOrderContactId
		WHERE WorkOrderId = @WorkOrderId

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Work Order record not found'
		RETURN -1
	END
END

RETURN @WorkOrderId

GO
/****** Object:  StoredProcedure [dbo].[spEditWorkOrderPO]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditWorkOrderPO]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditWorkOrderPO] AS' 
END
GO

ALTER PROCEDURE [dbo].[spEditWorkOrderPO]
@WorkOrderId		int,
@PurchaseOrderId	int

AS

BEGIN

IF (NOT EXISTS ( SELECT WorkOrderId FROM tblWorkOrderPOs WHERE WorkOrderId = @WorkOrderId AND PurchaseOrderId = @PurchaseOrderId ) )
BEGIN

	INSERT tblWorkOrderPOs ( WorkOrderId, PurchaseOrderId )
	VALUES (@WorkOrderId, @PurchaseOrderId)
END

END

GO
/****** Object:  StoredProcedure [dbo].[spEditWorkOrderPriority]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spEditWorkOrderPriority]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spEditWorkOrderPriority] AS' 
END
GO


ALTER PROCEDURE [dbo].[spEditWorkOrderPriority]

@ErrorMessage		varchar(255)	= NULL OUTPUT,
@WorkOrderPriorityID	int		= NULL,
@WorkOrderPriority	varchar(100)

AS

SET NOCOUNT ON

DECLARE @ErrorNum	int

IF( DATALENGTH( @WorkOrderPriorityID ) = 0 ) OR ( @WorkOrderPriorityID IS NULL ) OR ( @WorkOrderPriorityID = 0 )
BEGIN
	INSERT tblWorkOrderPriorities( WorkOrderPriority ) VALUES ( @WorkOrderPriority )

	SET @ErrorNum	= @@ERROR
	SET @WorkOrderPriorityID = @@IDENTITY

	IF( @ErrorNum <> 0 )
	BEGIN
		SELECT @ErrorMessage = ERROR_MESSAGE()
		SET @WorkOrderPriorityID = NULL
		RETURN -1
	END
END
ELSE
BEGIN
	IF( EXISTS( SELECT WorkOrderPriorityID FROM tblWorkOrderPriorities WHERE WorkOrderPriorityID = @WorkOrderPriorityID ) )
	BEGIN
		UPDATE tblWorkOrderPriorities SET WorkOrderPriority = @WorkOrderPriority
		WHERE WorkOrderPriorityID = @WorkOrderPriorityID

		SET @ErrorNum = @@ERROR
	
		IF( @ErrorNum <> 0 )
		BEGIN
			SELECT @ErrorMessage = ERROR_MESSAGE()
			RETURN -1
		END
	END
	ELSE
	BEGIN
		SELECT @ErrorMessage = 'Work Order Priority record not found'
		RETURN -1
	END
END

RETURN @WorkOrderPriorityID


GO
/****** Object:  StoredProcedure [dbo].[spGetCategoryList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetCategoryList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetCategoryList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetCategoryList]

AS

SELECT CategoryId, CategoryName FROM tblCategories WHERE NOT CategoryName IS NULL ORDER BY CategoryName



GO
/****** Object:  StoredProcedure [dbo].[spGetCompanyInformation]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetCompanyInformation]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetCompanyInformation] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetCompanyInformation]

AS

SELECT CompanyName, Address, City, StateOrProvince, PostalCode, Country, PhoneNumber, FaxNumber
FROM tblCompanyInformation
WHERE SetupId = 1


GO
/****** Object:  StoredProcedure [dbo].[spGetCustomerByCustomerId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetCustomerByCustomerId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetCustomerByCustomerId] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetCustomerByCustomerId]

@ErrorMessage	varchar(255)	OUTPUT,
@CustomerId	int

AS

SET TEXTSIZE 8000

IF( @CustomerId = 0 )
BEGIN
	SELECT '' AS CustomerId, '' AS Customer, '' AS Address, '' AS City, '' AS State, '' AS Country, '' AS Zip, '' AS Contact, '' AS ContactTitle, '' AS Department, 
	'' AS TelephoneNumber, '' AS Extension, '' AS FaxNumber, '' AS EmailAddress, '' AS SalesRepId, '' AS DateEstablished, '' AS Notes, 0 AS MainContactId
END
ELSE
BEGIN
	SELECT tblCustomers.CustomerId, Customer, Address, City, State, Country, Zip, SalesRepId, DateEstablished, MainContactId, Notes
	FROM tblCustomers 
	WHERE tblCustomers.CustomerId = @CustomerId

	SELECT tblCalls.CallId, CallDate, ( LastName + ', ' + FirstName ) AS EmployeeName, tblCalls.ProjectId, ProjectName, tblCalls.WorkOrderId, WorkOrderNumber, CallComments, tblProjects.CompletionDate
	FROM tblCalls
	LEFT JOIN tblEmployees ON tblCalls.EmployeeId = tblEmployees.EmployeeId
	LEFT JOIN tblProjects ON tblCalls.ProjectId = tblProjects.ProjectId 
	LEFT JOIN tblWorkOrders ON tblCalls.WorkOrderId = tblWorkOrders.WorkOrderId 
	WHERE tblCalls.CustomerId = @CustomerId
		AND tblProjects.CompletionDate IS NULL
	ORDER BY CallDate DESC
END

/*SELECT TransactionDate, PurchaseOrderId, TransactionDescription, UnitsOrdered, DateReceived, UnitsReceived, UnitsSold
FROM tblInventoryTransactions
WHERE tblInventoryTransactions.ProductId = @ProductId
ORDER BY TransactionDate DESC*/

GO
/****** Object:  StoredProcedure [dbo].[spGetCustomerContactByCustomerContactId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetCustomerContactByCustomerContactId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetCustomerContactByCustomerContactId] AS' 
END
GO

ALTER PROCEDURE [dbo].[spGetCustomerContactByCustomerContactId]

@ErrorMessage	varchar(255)	OUTPUT,
@CustomerContactId	int

AS

IF( @CustomerContactId = 0 )
BEGIN
	SELECT '' AS CustomerContactId, '' AS Contact, '' AS ContactTitle, '' AS Department, '' AS TelephoneNumber, '' AS Extension, '' AS FaxNumber, '' AS MobileNumber, '' AS EmailAddress, '' AS Notes
END
ELSE
BEGIN
	SELECT CustomerContactId, Contact, ContactTitle, Department, TelephoneNumber, Extension, FaxNumber, MobileNumber, EmailAddress, Notes
	FROM tblCustomerContacts 
	WHERE CustomerContactId = @CustomerContactId
END

GO
/****** Object:  StoredProcedure [dbo].[spGetCustomerContactsByCustomerId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetCustomerContactsByCustomerId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetCustomerContactsByCustomerId] AS' 
END
GO



ALTER  PROCEDURE [dbo].[spGetCustomerContactsByCustomerId]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@CustomerId	int 		= 0

AS

BEGIN

	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetCustomerContacts FROM (
	SELECT TOP 100 PERCENT CustomerContactId, Contact, ContactTitle, Department, TelephoneNumber, Extension
	FROM tblCustomerContacts
	WHERE ( CustomerId = @CustomerId )
	ORDER BY CustomerContactId
	) AS X


	BEGIN
		SELECT RowNumber, CustomerContactId, Contact, ContactTitle, Department, TelephoneNumber, Extension
		FROM #tblTempGetCustomerContacts
		ORDER BY RowNumber
	END
	
	DROP TABLE #tblTempGetCustomerContacts
	
END

GO
/****** Object:  StoredProcedure [dbo].[spGetCustomerList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetCustomerList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetCustomerList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetCustomerList]

AS

SELECT CustomerId, Customer FROM tblCustomers ORDER BY Customer



GO
/****** Object:  StoredProcedure [dbo].[spGetCustomers]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetCustomers]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetCustomers] AS' 
END
GO


ALTER   PROCEDURE [dbo].[spGetCustomers]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'CustomerName' OUTPUT,
@SearchString	varchar(255)	= '%',
@JumpTo		char(1)		= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'CustomerName'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetCustomers FROM (
SELECT TOP 100 PERCENT tblCustomers.CustomerId, MAX( Customer ) AS Customer, MAX( Address ) AS Address, MAX( City ) AS City, MAX( State ) AS State, MAX( Zip ) AS Zip,  MAX( TelephoneNumber ) AS TelephoneNumber,
MAX( DateEstablished ) AS DateEstablished,  MAX( Contact ) AS Contact, Max( FirstName + ' ' + LastName) AS SalesRep, MAX( LastName) AS LastName
FROM tblCustomers
LEFT JOIN tblEmployees ON tblCustomers.SalesRepId = tblEmployees.EmployeeId
LEFT JOIN tblCustomerContacts ON tblCustomers.MainContactId = tblCustomerContacts.CustomerContactId
WHERE
CASE @OrderBy
	WHEN 'CustomerName' THEN Customer
	WHEN 'Contact' THEN Contact
	WHEN 'SalesRep' THEN LastName
	WHEN 'Address' THEN Address
	WHEN 'City' THEN City
	WHEN 'State' THEN State
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblCustomers.CustomerId
ORDER BY
CASE @OrderBy WHEN 'CustomerName' THEN MAX( Customer ) ELSE NULL END,
CASE @OrderBy WHEN 'Contact' THEN MAX( Contact ) ELSE NULL END,
CASE @OrderBy WHEN 'SalesRep' THEN MAX( LastName  ) ELSE NULL END,
CASE @OrderBy WHEN 'Address' THEN MAX( Address ) ELSE NULL END,
CASE @OrderBy WHEN 'City' THEN MAX( City ) ELSE NULL END,
CASE @OrderBy WHEN 'State' THEN MAX( State ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetCustomers WHERE 
	CASE @OrderBy
		WHEN 'CustomerName' THEN SUBSTRING( Customer, 1, 1 )
		WHEN 'Contact' THEN SUBSTRING( Contact, 1, 1 )
		WHEN 'SalesRep' THEN SUBSTRING( LastName, 1, 1 )
		WHEN 'Address' THEN SUBSTRING( Address, 1, 1 )
		WHEN 'City' THEN SUBSTRING( City, 1, 1 )
		WHEN 'State' THEN SUBSTRING( State, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, CustomerId, Customer, Address, City, State, Zip, TelephoneNumber, Contact, SalesRep, DateEstablished
FROM #tblTempGetCustomers
WHERE (
CASE @OrderBy
	WHEN 'CustomerName' THEN SUBSTRING( Customer, 1, 1 )
	WHEN 'Contact' THEN SUBSTRING( Contact, 1, 1 )
	WHEN 'SalesRep' THEN SUBSTRING( LastName, 1, 1 )
	WHEN 'Address' THEN SUBSTRING( Address, 1, 1 )
	WHEN 'City' THEN SUBSTRING( City, 1, 1 )
	WHEN 'State' THEN SUBSTRING( State, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetCustomers

GO
/****** Object:  StoredProcedure [dbo].[spGetEmailByQuoteId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetEmailByQuoteId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetEmailByQuoteId] AS' 
END
GO


ALTER   PROCEDURE [dbo].[spGetEmailByQuoteId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteId	int

AS

BEGIN
	SELECT tblQuotes.QuoteId AS QuoteId, tblWorkOrders.PONumber AS PONum, tblCustomerContacts.EmailAddress AS ToEmail, tblEmployees.EmailAddress AS FromEmail
	FROM tblQuotes
	LEFT JOIN tblEmployees ON tblQuotes.QuotedById = tblEmployees.EmployeeId
	LEFT JOIN tblCustomerContacts ON tblQuotes.QuoteContactId = tblCustomerContacts.CustomerContactId
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	WHERE tblQuotes.QuoteId = @QuoteId
END


GO
/****** Object:  StoredProcedure [dbo].[spGetEmployeeByEmployeeId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetEmployeeByEmployeeId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetEmployeeByEmployeeId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetEmployeeByEmployeeId]

@ErrorMessage	varchar(255)	OUTPUT,
@EmployeeId	int

AS

IF( @EmployeeId = 0 )
BEGIN
	SELECT '' AS EmployeeId, '' AS LastName, '' AS FirstName, '' AS Title, '' AS WorkPhone, '' AS Extension, '' AS EmailAddress
END
ELSE
BEGIN
	SELECT tblEmployees.EmployeeId, LastName, FirstName, Title, WorkPhone, Extension, EmailAddress
	FROM tblEmployees 
	WHERE tblEmployees.EmployeeId = @EmployeeId
END



GO
/****** Object:  StoredProcedure [dbo].[spGetEmployeeList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetEmployeeList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetEmployeeList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetEmployeeList]

AS

SELECT EmployeeId, LastName + ', ' + FirstName AS EmployeeName FROM tblEmployees ORDER BY EmployeeName



GO
/****** Object:  StoredProcedure [dbo].[spGetEmployees]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetEmployees]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetEmployees] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetEmployees]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'LastName',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'LastName'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetEmployees FROM (
SELECT TOP 100 PERCENT tblEmployees.EmployeeId, MAX( LastName ) AS LastName, MAX( FirstName ) AS FirstName,
MAX( Title ) AS Title, MAX( Extension ) AS Extension, MAX( WorkPhone ) AS WorkPhone, MAX( EmailAddress ) AS EmailAddress
FROM tblEmployees
WHERE
CASE @OrderBy
	WHEN 'LastName' THEN LastName
	WHEN 'FirstName' THEN FirstName
	WHEN 'Title' THEN Title
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblEmployees.EmployeeId
ORDER BY
CASE @OrderBy WHEN 'LastName' THEN MAX( LastName ) ELSE NULL END,
CASE @OrderBy WHEN 'FirstName' THEN MAX( FirstName ) ELSE NULL END,
CASE @OrderBy WHEN 'Title' THEN MAX( Title ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetEmployees WHERE 
	CASE @OrderBy
		WHEN 'LastName' THEN SUBSTRING( LastName, 1, 1 )
		WHEN 'FirstName' THEN SUBSTRING( FirstName, 1, 1 )
		WHEN 'Title' THEN SUBSTRING( Title, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, EmployeeId, LastName, FirstName, Title, Extension, WorkPhone, EmailAddress
FROM #tblTempGetEmployees
WHERE (
CASE @OrderBy
	WHEN 'LastName' THEN SUBSTRING( LastName, 1, 1 )
	WHEN 'FirstName' THEN SUBSTRING( FirstName, 1, 1 )
	WHEN 'Title' THEN SUBSTRING( Title, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetEmployees



GO
/****** Object:  StoredProcedure [dbo].[spGetInventoryTransactionByInventoryTransactionId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetInventoryTransactionByInventoryTransactionId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetInventoryTransactionByInventoryTransactionId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetInventoryTransactionByInventoryTransactionId]

@ErrorMessage		varchar(255)	OUTPUT,
@InventoryTransactionId	int

AS

IF( @InventoryTransactionId = 0 )
BEGIN
	SELECT '' AS TransactionId, '' AS TransactionDate, '' AS ProductID, '' AS PurchaseOrderID, '' AS TransactionDescription, '' AS UnitPrice, '' AS UnitsOrdered,
	'' AS DateReceived, '' AS UnitsReceived, '' AS UnitsSold, '' AS UnitsShrinkage
END
ELSE
BEGIN
	SELECT TransactionId, TransactionDate, ProductId, PurchaseOrderID, TransactionDescription, ISNULL(UnitPrice,0) AS UnitPrice, UnitsOrdered, DateReceived, UnitsReceived, UnitsSold, UnitsShrinkage
	FROM tblInventoryTransactions 
	WHERE tblInventoryTransactions.TransactionId = @InventoryTransactionId
END

GO
/****** Object:  StoredProcedure [dbo].[spGetInventoryTransactions]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetInventoryTransactions]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetInventoryTransactions] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetInventoryTransactions]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'PartNumber' OUTPUT,
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(11)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'PartNumber'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

IF( @OrderBy = 'TransactionDate' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempInventoryTransactionsByTransDate FROM (
	SELECT TOP 100 PERCENT MAX( TransactionDate ) AS TransactionDate, TransactionId, MAX( tblInventoryTransactions.PurchaseOrderId ) AS PurchaseOrderId, 
	MAX( PurchaseOrderNumber ) AS PurchaseOrderNumber, MAX (PartNumber) AS PartNumber,
	MAX( UnitsOrdered ) AS UnitsOrdered, MAX( DateReceived ) AS DateReceived, MAX( UnitsReceived ) AS UnitsReceived, MAX( UnitsSold ) AS UnitsSold
	FROM tblInventoryTransactions
	LEFT JOIN tblProducts ON tblInventoryTransactions.ProductId = tblProducts.ProductId
	LEFT JOIN tblPurchaseOrders ON tblInventoryTransactions.PurchaseOrderId = tblPurchaseOrders.PurchaseOrderId
	GROUP BY tblInventoryTransactions.TransactionId
	ORDER BY TransactionDate
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempInventoryTransactionsByTransDate
		WHERE TransactionDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, TransactionDate, TransactionId, PurchaseOrderId, PurchaseOrderNumber, PartNumber, TransactionDescription, 
		UnitsOrdered, DateReceived, UnitsReceived, UnitsSold, tblProducts.ProductId
		FROM #tblTempInventoryTransactionsByTransDate
		LEFT JOIN tblInventoryTransactions ON #tblTempInventoryTransactionsByTransDate.TransactionId = tblInventoryTransactions.TransactionId
		WHERE ( TransactionDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SELECT  #tblTempInventoryTransactionsByTransDate.*, tblInventoryTransactions.TransactionDescription
		FROM #tblTempInventoryTransactionsByTransDate
		LEFT JOIN tblInventoryTransactions ON #tblTempInventoryTransactionsByTransDate.TransactionId = tblInventoryTransactions.TransactionId
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempInventoryTransactionsByTransDate
END
ELSE IF( @OrderBy = 'TransactionId' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempInventoryTransactionsByTransId FROM (
	SELECT TOP 100 PERCENT MAX( TransactionDate ) AS TransactionDate, TransactionId, MAX( tblInventoryTransactions.PurchaseOrderId ) AS PurchaseOrderId, 
	MAX( PurchaseOrderNumber ) AS PurchaseOrderNumber, MAX( PartNumber ) AS PartNumber,
	MAX( UnitsOrdered ) AS UnitsOrdered, MAX( DateReceived ) AS DateReceived, MAX( UnitsReceived ) AS UnitsReceived, MAX( UnitsSold ) AS UnitsSold
	FROM tblInventoryTransactions
	LEFT JOIN tblProducts ON tblInventoryTransactions.ProductId = tblProducts.ProductId
	LEFT JOIN tblPurchaseOrders ON tblInventoryTransactions.PurchaseOrderId = tblPurchaseOrders.PurchaseOrderId
	GROUP BY tblInventoryTransactions.TransactionId
	ORDER BY TransactionId
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempInventoryTransactionsByTransId
		WHERE TransactionId BETWEEN CONVERT( int, @JumpTo ) AND 9999999 ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SET @JumpTo = 1
	END
	
	SELECT  #tblTempInventoryTransactionsByTransId.*, tblInventoryTransactions.TransactionDescription
	FROM #tblTempInventoryTransactionsByTransId
		LEFT JOIN tblInventoryTransactions ON #tblTempInventoryTransactionsByTransId.TransactionId = tblInventoryTransactions.TransactionId
	WHERE ( TransactionId BETWEEN CONVERT( int, @JumpTo ) AND 9999999 )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	DROP TABLE #tblTempInventoryTransactionsByTransId
END
ELSE
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempInventoryTransactions FROM (
	SELECT TOP 100 PERCENT MAX( CONVERT( varchar(15), TransactionDate, 101 ) ) AS TransactionDate, TransactionId, MAX( tblInventoryTransactions.PurchaseOrderId ) AS PurchaseOrderId, 
	MAX( PurchaseOrderNumber ) AS PurchaseOrderNumber, MAX( PartNumber ) AS PartNumber,
	MAX( UnitsOrdered ) AS UnitsOrdered, MAX( DateReceived ) AS DateReceived, MAX( UnitsReceived ) AS UnitsReceived, MAX( UnitsSold ) AS UnitsSold
	FROM tblInventoryTransactions
	LEFT JOIN tblProducts ON tblInventoryTransactions.ProductId = tblProducts.ProductId
	LEFT JOIN tblPurchaseOrders ON tblInventoryTransactions.PurchaseOrderId = tblPurchaseOrders.PurchaseOrderId
	WHERE
	CASE @OrderBy
		WHEN 'PartNumber' THEN PartNumber
		WHEN 'TransactionDescription' THEN TransactionDescription
		WHEN 'PurchaseOrderNumber' THEN PurchaseOrderNumber
	END
	LIKE
	CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	GROUP BY tblInventoryTransactions.TransactionId
	ORDER BY
	CASE @OrderBy WHEN 'PartNumber' THEN MAX( PartNumber ) ELSE NULL END,
	--CASE @OrderBy WHEN 'TransactionDescription' THEN MAX( TransactionDescription ) ELSE NULL END,
	CASE @OrderBy WHEN 'PurchaseOrderNumber' THEN MAX( PurchaseOrderNumber ) ELSE NULL END
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempInventoryTransactions WHERE 
		CASE @OrderBy
			WHEN 'PartNumber' THEN SUBSTRING( PartNumber, 1, 1 )
			--WHEN 'TransactionDescription' THEN SUBSTRING( TransactionDescription, 1, 1 )
			WHEN 'PurchaseOrderNumber' THEN SUBSTRING( PurchaseOrderNumber, 1, 1 )
		END
		LIKE @JumpTo ORDER BY RowNumber
	END
	
	SELECT #tblTempInventoryTransactions.*, tblInventoryTransactions.TransactionDescription
	FROM #tblTempInventoryTransactions
		LEFT JOIN tblInventoryTransactions ON #tblTempInventoryTransactions.TransactionId = tblInventoryTransactions.TransactionId
	WHERE (
	CASE @OrderBy
		WHEN 'PartNumber' THEN SUBSTRING( PartNumber, 1, 1 )
		--WHEN 'TransactionDescription' THEN SUBSTRING( TransactionDescription, 1, 1 )
		WHEN 'PurchaseOrderNumber' THEN SUBSTRING( PurchaseOrderNumber, 1, 1 )
	END
	BETWEEN @JumpTo AND 'Z' )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	
	DROP TABLE #tblTempInventoryTransactions
END



GO
/****** Object:  StoredProcedure [dbo].[spGetLocationByLocationId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetLocationByLocationId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetLocationByLocationId] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetLocationByLocationId]

@ErrorMessage		varchar(255)	OUTPUT,
@LocationId	int

AS

IF( @LocationId = 0 )
BEGIN
	SELECT '' AS LocationId, '' AS Location
END
ELSE
BEGIN
	SELECT tblLocations.LocationId, Location
	FROM tblLocations 
	WHERE tblLocations.LocationId = @LocationId
END



GO
/****** Object:  StoredProcedure [dbo].[spGetLocationList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetLocationList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetLocationList] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetLocationList]

AS

SELECT LocationId, Location FROM tblLocations ORDER BY Location



GO
/****** Object:  StoredProcedure [dbo].[spGetLocations]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetLocations]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetLocations] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetLocations]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'Location',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'Location'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetLocations FROM (
SELECT TOP 100 PERCENT tblLocations.LocationId, MAX( Location ) AS Location
FROM tblLocations
WHERE
CASE @OrderBy
	WHEN 'Location' THEN Location
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblLocations.LocationId
ORDER BY
CASE @OrderBy WHEN 'Location' THEN MAX( Location ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetLocations WHERE 
	CASE @OrderBy
		WHEN 'Location' THEN SUBSTRING( Location, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, LocationId, Location
FROM #tblTempGetLocations
WHERE (
CASE @OrderBy
	WHEN 'Location' THEN SUBSTRING( Location, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetLocations



GO
/****** Object:  StoredProcedure [dbo].[spGetProductByProductId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProductByProductId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProductByProductId] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetProductByProductId]

@ErrorMessage	varchar(255)	OUTPUT,
@ProductId	int

AS

IF( @ProductId = 0 )
BEGIN
	SELECT '' AS ProductId, '' AS ProductName, '' AS ProductDescription, '' AS CategoryName, '' AS DrawingId, '' AS LeadTime, '' AS ReorderLevel,
	0 AS OnOrder, 0 AS OnHand, 0 AS UnitPrice, '' AS SerialNumber, '' AS ProductOwner, '' AS PartNumber
END
ELSE
BEGIN
	SELECT * INTO #tblTempGetProductByProductId FROM (
	SELECT TOP 100 PERCENT MAX( tblProducts.ProductId ) AS ProductId, MAX( PartNumber ) AS PartNumber, 
	MAX( CategoryName ) AS CategoryName, MAX( ProductOwner ) AS ProductOwner, MAX( DrawingId ) AS DrawingId, MAX( ISNULL(LeadTime,0) ) AS LeadTime, MAX( ReorderLevel ) AS ReorderLevel,
	( SUM( ISNULL( UnitsOrdered, 0 ) ) - SUM( ISNULL( UnitsReceived, 0 ) ) ) AS OnOrder,
	( SUM( ISNULL( UnitsReceived, 0 ) ) - SUM( ISNULL( UnitsShrinkage, 0 ) ) - SUM( ISNULL( UnitsSold, 0 ) ) ) AS OnHand,
	MAX( ISNULL( tblProducts.UnitPrice, 0 ) ) AS UnitPrice, MAX( SerialNumber ) AS SerialNumber
	FROM tblProducts 
	LEFT JOIN tblCategories ON tblProducts.CategoryId = tblCategories.CategoryId
	LEFT JOIN tblProductOwners ON tblProducts.ProductOwnerId = tblProductOwners.ProductOwnerId
	LEFT JOIN tblInventoryTransactions on tblProducts.ProductId = tblInventoryTransactions.ProductId
	WHERE tblProducts.ProductId = @ProductId
	) AS X

	SELECT #tblTempGetProductByProductId.*, ProductDescription, ProductName
	FROM #tblTempGetProductByProductId
	LEFT JOIN tblProducts ON #tblTempGetProductByProductId.ProductId = tblProducts.ProductId
	
	
	SELECT TransactionDate, PurchaseOrderId, TransactionDescription, UnitsOrdered, DateReceived, UnitsReceived, UnitsSold
	FROM tblInventoryTransactions
	WHERE tblInventoryTransactions.ProductId = @ProductId
	ORDER BY TransactionDate DESC
END

GO
/****** Object:  StoredProcedure [dbo].[spGetProductCategories]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProductCategories]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProductCategories] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProductCategories]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'ProductCategory',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'ProductCategory'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProductCategories FROM (
SELECT TOP 100 PERCENT tblCategories.CategoryId AS ProductCategoryId, MAX( CategoryName ) AS ProductCategory
FROM tblCategories
WHERE
CASE @OrderBy
	WHEN 'ProductCategory' THEN CategoryName
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblCategories.CategoryId
ORDER BY
CASE @OrderBy WHEN 'ProductCategory' THEN MAX( CategoryName ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProductCategories WHERE 
	CASE @OrderBy
		WHEN 'ProductCategory' THEN SUBSTRING( ProductCategory, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, ProductCategoryId, ProductCategory
FROM #tblTempGetProductCategories
WHERE (
CASE @OrderBy
	WHEN 'ProductCategory' THEN SUBSTRING( ProductCategory, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetProductCategories



GO
/****** Object:  StoredProcedure [dbo].[spGetProductCategoryByProductCategoryId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProductCategoryByProductCategoryId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProductCategoryByProductCategoryId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProductCategoryByProductCategoryId]

@ErrorMessage		varchar(255)	OUTPUT,
@ProductCategoryId	int

AS

IF( @ProductCategoryId = 0 )
BEGIN
	SELECT '' AS ProductCategoryId, '' AS ProductCategory
END
ELSE
BEGIN
	SELECT tblCategories.CategoryId AS ProductCategoryId, CategoryName AS ProductCategory
	FROM tblCategories
	WHERE tblCategories.CategoryId = @ProductCategoryId
END




GO
/****** Object:  StoredProcedure [dbo].[spGetProductList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProductList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProductList] AS' 
END
GO





ALTER   PROCEDURE [dbo].[spGetProductList]

AS

SELECT ProductId, PartNumber, ProductName, UnitPrice FROM tblProducts ORDER BY PartNumber





GO
/****** Object:  StoredProcedure [dbo].[spGetProductListByName]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProductListByName]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProductListByName] AS' 
END
GO





ALTER   PROCEDURE [dbo].[spGetProductListByName]

AS

SELECT ProductId, PartNumber, ProductName, UnitPrice FROM tblProducts ORDER BY PartNumber




GO
/****** Object:  StoredProcedure [dbo].[spGetProductOwnerByProductOwnerId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProductOwnerByProductOwnerId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProductOwnerByProductOwnerId] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetProductOwnerByProductOwnerId]

@ErrorMessage		varchar(255)	OUTPUT,
@ProductOwnerId	int

AS

IF( @ProductOwnerId = 0 )
BEGIN
	SELECT '' AS ProductOwnerId, '' AS ProductOwner
END
ELSE
BEGIN
	SELECT tblProductOwners.ProductOwnerId AS ProductOwnerId, ProductOwner AS ProductOwner
	FROM tblProductOwners
	WHERE tblProductOwners.ProductOwnerId = @ProductOwnerId
END



GO
/****** Object:  StoredProcedure [dbo].[spGetProductOwnerList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProductOwnerList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProductOwnerList] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetProductOwnerList]

AS

SELECT ProductOwnerId, ProductOwner FROM tblProductOwners ORDER BY ProductOwner

GO
/****** Object:  StoredProcedure [dbo].[spGetProductOwners]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProductOwners]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProductOwners] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetProductOwners]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'ProductOwner',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'ProductOwner'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProductOwners FROM (
SELECT TOP 100 PERCENT tblProductOwners.ProductOwnerId AS ProductOwnerId, MAX( ProductOwner ) AS ProductOwner
FROM tblProductOwners
WHERE
CASE @OrderBy
	WHEN 'ProductOwner' THEN ProductOwner
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblProductOwners.ProductOwnerId
ORDER BY
CASE @OrderBy WHEN 'ProductOwner' THEN MAX( ProductOwner ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProductOwners WHERE 
	CASE @OrderBy
		WHEN 'ProductOwner' THEN SUBSTRING( ProductOwner, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, ProductOwnerId, ProductOwner
FROM #tblTempGetProductOwners
WHERE (
CASE @OrderBy
	WHEN 'ProductOwner' THEN SUBSTRING( ProductOwner, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetProductOwners



GO
/****** Object:  StoredProcedure [dbo].[spGetProducts]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProducts]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProducts] AS' 
END
GO


ALTER    PROCEDURE [dbo].[spGetProducts]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'PartNumber' OUTPUT,
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'PartNumber'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'


IF( @OrderBy = 'ProductId' )
BEGIN
	SELECT * INTO #tblTempGetProductsAggregateId FROM (
	SELECT TOP 100 PERCENT tblProducts.ProductId AS ProductID, ( SUM( ISNULL( UnitsOrdered, 0 ) ) - SUM( ISNULL( UnitsReceived, 0 ) ) ) AS OnOrder,
	( SUM( ISNULL( UnitsReceived, 0 ) ) - SUM( ISNULL( UnitsShrinkage, 0 ) ) - SUM( ISNULL( UnitsSold, 0 ) ) ) AS OnHand
	FROM tblProducts
	LEFT JOIN tblInventoryTransactions ON tblProducts.ProductId = tblInventoryTransactions.ProductId
	LEFT JOIN tblCategories ON tblProducts.CategoryId = tblCategories.CategoryId
	WHERE tblProducts.ProductID
	 LIKE CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	GROUP BY tblProducts.ProductID
	ORDER BY tblProducts.ProductID
	) AS Y

	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProductsById FROM (
	SELECT TOP 100 PERCENT tblProducts.ProductID AS ProductID, PartNumber,
	CategoryName, 
	ReorderLevel, OnOrder, OnHand,
	ISNULL( tblProducts.UnitPrice, 0 ) AS UnitPrice, ProductShortDescription
	FROM tblProducts
	LEFT JOIN #tblTempGetProductsAggregateId ON tblProducts.ProductID = #tblTempGetProductsAggregateId.ProductID
	LEFT JOIN tblCategories ON tblProducts.CategoryId = tblCategories.CategoryId
	WHERE tblProducts.ProductId LIKE CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	ORDER BY tblProducts.ProductID ) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProductsById
		WHERE ProductId BETWEEN CONVERT( int, @JumpTo ) AND 9999999 ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SET @JumpTo = 1
	END
	
	SELECT RowNumber, ProductID, PartNumber, CategoryName, ReorderLevel, OnOrder, OnHand, UnitPrice, ProductShortDescription
	FROM #tblTempGetProductsById
	WHERE ( ProductId BETWEEN CONVERT( int, @JumpTo ) AND 9999999 )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	DROP TABLE #tblTempGetProductsById
	DROP TABLE #tblTempGetProductsAggregateId
END
ELSE
BEGIN
	SELECT * INTO #tblTempGetProductsAggregate FROM (
	SELECT TOP 100 PERCENT tblProducts.ProductId AS ProductID, ( SUM( ISNULL( UnitsOrdered, 0 ) ) - SUM( ISNULL( UnitsReceived, 0 ) ) ) AS OnOrder,
	( SUM( ISNULL( UnitsReceived, 0 ) ) - SUM( ISNULL( UnitsShrinkage, 0 ) ) - SUM( ISNULL( UnitsSold, 0 ) ) ) AS OnHand
	FROM tblProducts
	LEFT JOIN tblInventoryTransactions ON tblProducts.ProductId = tblInventoryTransactions.ProductId
	LEFT JOIN tblCategories ON tblProducts.CategoryId = tblCategories.CategoryId
	WHERE
	CASE @OrderBy
		WHEN 'PartNumber' THEN PartNumber
		WHEN 'CategoryName' THEN CategoryName
		WHEN 'ProductShortDescription' THEN ProductShortDescription
	END
	 LIKE CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	GROUP BY tblProducts.ProductID
	ORDER BY tblProducts.ProductID
	) AS Y
	
	SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProducts FROM (
	SELECT TOP 100 PERCENT tblProducts.ProductID AS ProductID, PartNumber,
	CategoryName,
	ReorderLevel, OnOrder, OnHand,
	ISNULL( tblProducts.UnitPrice, 0 ) AS UnitPrice, ProductShortDescription
	FROM tblProducts
	LEFT JOIN #tblTempGetProductsAggregate ON tblProducts.ProductID = #tblTempGetProductsAggregate.ProductID
	LEFT JOIN tblCategories ON tblProducts.CategoryId = tblCategories.CategoryId
	WHERE
	CASE @OrderBy
		WHEN 'PartNumber' THEN PartNumber
		WHEN 'CategoryName' THEN CategoryName
		WHEN 'ProductShortDescription' THEN ProductShortDescription
	END
	LIKE
	CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	ORDER BY
	CASE @OrderBy WHEN 'PartNumber' THEN PartNumber ELSE NULL END,
	CASE @OrderBy WHEN 'CategoryName' THEN CategoryName ELSE NULL END,
	CASE @OrderBy WHEN 'ProductShortDescription' THEN ProductShortDescription ELSE NULL END
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProducts WHERE 
		CASE @OrderBy
			WHEN 'PartNumber' THEN SUBSTRING( PartNumber, 1, 1 )
			WHEN 'CategoryName' THEN SUBSTRING( CategoryName, 1, 1 )
			WHEN 'ProductShortDescription' THEN SUBSTRING( ProductShortDescription, 1, 1 )
		END
		LIKE @JumpTo ORDER BY RowNumber
	END
	
	SELECT RowNumber, ProductID, PartNumber, CategoryName, ReorderLevel, OnOrder, OnHand, UnitPrice, ProductShortDescription
	FROM #tblTempGetProducts
	WHERE (
	CASE @OrderBy
		WHEN 'PartNumber' THEN SUBSTRING( PartNumber, 1, 1 )
		WHEN 'CategoryName' THEN SUBSTRING( CategoryName, 1, 1 )
		WHEN 'ProductShortDescription' THEN SUBSTRING( ProductShortDescription, 1, 1 )
	END
	BETWEEN @JumpTo AND 'Z' )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	DROP TABLE #tblTempGetProducts
	DROP TABLE #tblTempGetProductsAggregate
END

GO
/****** Object:  StoredProcedure [dbo].[spGetProjectByProjectId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectByProjectId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectByProjectId] AS' 
END
GO


ALTER PROCEDURE [dbo].[spGetProjectByProjectId]

@ErrorMessage	varchar(255)	OUTPUT,
@ProjectId	int,
@CustomerId	int	OUTPUT,
@SpindleId	int	OUTPUT

AS

IF( @ProjectId = 0 )
BEGIN
	SELECT '' AS ProjectId, '' AS ProjectName, '' AS ProjectTypeId, '' AS SalesRepId, '' AS EmployeeName, '' AS CustomerId, '' AS Customer, '' AS CustomerPO, '' AS ProjectDescription,
	'' AS StartDate, '' AS EstCompletionDate, '' AS CompletionDate, '' AS ProjectPriorityId, 0 AS SpindleId, '' AS SpindleType, 0 AS ProjectContactId
END
ELSE
BEGIN
	SELECT tblProjects.ProjectId, ProjectName, ProjectTypeId, tblCustomers.SalesRepId AS SalesRepId, ( tblEmployees.LastName + ', ' + tblEmployees.FirstName ) AS EmployeeName, tblProjects.CustomerId, tblCustomers.Customer, CustomerPO, ProjectDescription, StartDate, EstCompletionDate, CompletionDate, 
	ProjectPriorityId, tblProjects.SpindleId, SpindleType, ProjectContactId
	FROM tblProjects
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblEmployees ON tblCustomers.SalesRepId = tblEmployees.EmployeeId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE tblProjects.ProjectId = @ProjectId

	SELECT @CustomerId = tblProjects.CustomerId, @SpindleId = tblProjects.SpindleId	FROM tblProjects
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	WHERE tblProjects.ProjectId = @ProjectId

	
	SELECT tblCalls.CallId, CallDate, ( LastName + ', ' + FirstName ) AS EmployeeName, tblCalls.ProjectId, ProjectName, tblCalls.WorkOrderId, WorkOrderNumber, CallComments, tblWorkOrders.DateOut
	FROM tblCalls
	LEFT JOIN tblEmployees ON tblCalls.EmployeeId = tblEmployees.EmployeeId
	LEFT JOIN tblProjects ON tblCalls.ProjectId = tblProjects.ProjectId 
	LEFT JOIN tblWorkOrders ON tblCalls.WorkOrderId = tblWorkOrders.WorkOrderId 
	WHERE tblCalls.ProjectId = @ProjectId
		AND tblWorkOrders.DateOut IS NULL
	ORDER BY CallDate DESC

	SELECT tblWorkOrders.WorkOrderId, WorkOrderNumber, PromiseDate, DateIn, DateOut, ProjectId, SerialNumber
	FROM tblWorkOrders
	WHERE ProjectId = @ProjectId
END

GO
/****** Object:  StoredProcedure [dbo].[spGetProjectList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProjectList]

AS

SELECT ProjectId, ProjectName FROM tblProjects ORDER BY ProjectName



GO
/****** Object:  StoredProcedure [dbo].[spGetProjectPriorities]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectPriorities]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectPriorities] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProjectPriorities]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'ProjectPriority',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'ProjectPriority'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProjectPriorities FROM (
SELECT TOP 100 PERCENT tblProjectPriorities.ProjectPriorityId, MAX( ProjectPriority ) AS ProjectPriority
FROM tblProjectPriorities
WHERE
CASE @OrderBy
	WHEN 'ProjectPriority' THEN ProjectPriority
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblProjectPriorities.ProjectPriorityId
ORDER BY
CASE @OrderBy WHEN 'ProjectPriority' THEN MAX( ProjectPriority ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProjectPriorities WHERE 
	CASE @OrderBy
		WHEN 'ProjectPriority' THEN SUBSTRING( ProjectPriority, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, ProjectPriorityId, ProjectPriority
FROM #tblTempGetProjectPriorities
WHERE (
CASE @OrderBy
	WHEN 'ProjectPriority' THEN SUBSTRING( ProjectPriority, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetProjectPriorities




GO
/****** Object:  StoredProcedure [dbo].[spGetProjectPriorityByProjectPriorityId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectPriorityByProjectPriorityId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectPriorityByProjectPriorityId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProjectPriorityByProjectPriorityId]

@ErrorMessage	varchar(255)	OUTPUT,
@ProjectPriorityId	int

AS

IF( @ProjectPriorityId = 0 )
BEGIN
	SELECT '' AS ProjectPriorityId, '' AS ProjectPriority
END
ELSE
BEGIN
	SELECT tblProjectPriorities.ProjectPriorityId, ProjectPriority
	FROM tblProjectPriorities 
	WHERE tblProjectPriorities.ProjectPriorityId = @ProjectPriorityId
END






GO
/****** Object:  StoredProcedure [dbo].[spGetProjectPriorityList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectPriorityList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectPriorityList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProjectPriorityList]

AS

SELECT ProjectPriorityId, ProjectPriority FROM tblProjectPriorities ORDER BY ProjectPriority



GO
/****** Object:  StoredProcedure [dbo].[spGetProjectTypeByProjectTypeId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectTypeByProjectTypeId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectTypeByProjectTypeId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProjectTypeByProjectTypeId]

@ErrorMessage	varchar(255)	OUTPUT,
@ProjectTypeId	int

AS

IF( @ProjectTypeId = 0 )
BEGIN
	SELECT '' AS ProjectTypeId, '' AS ProjectType
END
ELSE
BEGIN
	SELECT tblProjectTypes.ProjectTypeId, ProjectType
	FROM tblProjectTypes 
	WHERE tblProjectTypes.ProjectTypeId = @ProjectTypeId
END





GO
/****** Object:  StoredProcedure [dbo].[spGetProjectTypeList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectTypeList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectTypeList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProjectTypeList]

AS

SELECT ProjectTypeId, ProjectType FROM tblProjectTypes ORDER BY ProjectType



GO
/****** Object:  StoredProcedure [dbo].[spGetProjectTypes]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectTypes]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectTypes] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetProjectTypes]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'ProjectType',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'ProjectType'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProjectTypes FROM (
SELECT TOP 100 PERCENT tblProjectTypes.ProjectTypeId, MAX( ProjectType ) AS ProjectType
FROM tblProjectTypes
WHERE
CASE @OrderBy
	WHEN 'ProjectType' THEN ProjectType
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblProjectTypes.ProjectTypeId
ORDER BY
CASE @OrderBy WHEN 'ProjectType' THEN MAX( ProjectType ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProjectTypes WHERE 
	CASE @OrderBy
		WHEN 'ProjectType' THEN SUBSTRING( ProjectType, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, ProjectTypeId, ProjectType
FROM #tblTempGetProjectTypes
WHERE (
CASE @OrderBy
	WHEN 'ProjectType' THEN SUBSTRING( ProjectType, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetProjectTypes



GO
/****** Object:  StoredProcedure [dbo].[spGetProjects]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjects]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjects] AS' 
END
GO







ALTER      PROCEDURE [dbo].[spGetProjects]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'ProjectName' OUTPUT,
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(11)	= '%'

AS
SET TEXTSIZE 8000

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'StartDate'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

IF( @OrderBy = 'StartDate' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProjectsByStartDate FROM (
	SELECT TOP 100 PERCENT ProjectId, ProjectName, ProjectType, Customer, 
	( LastName + ', ' + FirstName ) AS SalesRep, StartDate, EstCompletionDate, CompletionDate, 
	ProjectDescription, CustomerPO, SpindleType
	FROM tblProjects
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblEmployees ON tblProjects.SalesRepId = tblEmployees.EmployeeId
	LEFT JOIN tblProjectTypes ON tblProjects.ProjectTypeId = tblProjectTypes.ProjectTypeId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	ORDER BY StartDate
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProjectsByStartDate
		WHERE StartDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription, CustomerPO, SpindleType
		FROM #tblTempGetProjectsByStartDate
		WHERE ( StartDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription, CustomerPO, SpindleType
		FROM #tblTempGetProjectsByStartDate
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetProjectsByStartDate
END
ELSE IF( @OrderBy = 'EstCompletionDate' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProjectsByEstCompletionDate FROM (
	SELECT TOP 100 PERCENT ProjectId, ProjectName, ProjectType, Customer, 
	( LastName + ', ' + FirstName ) AS SalesRep, StartDate, EstCompletionDate, CompletionDate, 
	ProjectDescription, CustomerPO, SpindleType
	FROM tblProjects
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblEmployees ON tblProjects.SalesRepId = tblEmployees.EmployeeId
	LEFT JOIN tblProjectTypes ON tblProjects.ProjectTypeId = tblProjectTypes.ProjectTypeId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	ORDER BY EstCompletionDate
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProjectsByEstCompletionDate
		WHERE EstCompletionDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription, CustomerPO, SpindleType
		FROM #tblTempGetProjectsByEstCompletionDate
		WHERE ( EstCompletionDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	END
	ELSE
	BEGIN
		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription, CustomerPO, SpindleType
		FROM #tblTempGetProjectsByEstCompletionDate
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetProjectsByEstCompletionDate
END
ELSE IF( @OrderBy = 'CompletionDate' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProjectsByCompletionDate FROM (
	SELECT TOP 100 PERCENT ProjectId, ProjectName, ProjectType, Customer, 
	( LastName + ', ' + FirstName ) AS SalesRep, StartDate, EstCompletionDate, CompletionDate, 
	ProjectDescription, CustomerPO, SpindleType
	FROM tblProjects
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblEmployees ON tblProjects.SalesRepId = tblEmployees.EmployeeId
	LEFT JOIN tblProjectTypes ON tblProjects.ProjectTypeId = tblProjectTypes.ProjectTypeId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	ORDER BY CompletionDate
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProjectsByCompletionDate
		WHERE CompletionDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription, CustomerPO, SpindleType
		FROM #tblTempGetProjectsByCompletionDate
		WHERE ( CompletionDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	END
	ELSE
	BEGIN
		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription, CustomerPO, SpindleType
		FROM #tblTempGetProjectsByCompletionDate
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetProjectsByCompletionDate
END
ELSE
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProjects FROM (
	SELECT TOP 100 PERCENT ProjectId, ProjectName, ProjectType, Customer, 
	( LastName + ', ' + FirstName ) AS SalesRep, LastName, FirstName, StartDate, EstCompletionDate, CompletionDate, 
	ProjectDescription, CustomerPO, SpindleType
	FROM tblProjects
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblEmployees ON tblProjects.SalesRepId = tblEmployees.EmployeeId
	LEFT JOIN tblProjectTypes ON tblProjects.ProjectTypeId = tblProjectTypes.ProjectTypeId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE
	CASE @OrderBy
		WHEN 'Customer' THEN Customer
		WHEN 'ProjectType' THEN ProjectType
		WHEN 'ProjectName' THEN ProjectName
		WHEN 'ProjectDescription' THEN ProjectDescription
		WHEN 'SpindleType' THEN SpindleType
		WHEN 'SalesRep' THEN LastName
		WHEN 'CustomerPO' THEN CustomerPO
	END
	LIKE
	CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	ORDER BY
	CASE @OrderBy WHEN 'Customer' THEN ( Customer ) ELSE NULL END,
	CASE @OrderBy WHEN 'ProjectType' THEN ( ProjectType ) ELSE NULL END,
	CASE @OrderBy WHEN 'ProjectName' THEN ( ProjectName ) ELSE NULL END,
	CASE @OrderBy WHEN 'SpindleType' THEN ( SpindleType ) ELSE NULL END,
	CASE @OrderBy WHEN 'SalesRep' THEN ( LastName ) ELSE NULL END,
	CASE @OrderBy WHEN 'CustomerPO' THEN ( CustomerPO ) ELSE NULL END
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetProjects WHERE 
		CASE @OrderBy
			WHEN 'Customer' THEN SUBSTRING( Customer, 1, 1 )
			WHEN 'ProjectType' THEN SUBSTRING( ProjectType, 1, 1 )
			WHEN 'ProjectName' THEN SUBSTRING( ProjectName, 1, 1 )
			WHEN 'ProjectDescription' THEN SUBSTRING( ProjectDescription, 1, 1 )
			WHEN 'SpindleType' THEN SUBSTRING( SpindleType, 1, 1 )
			WHEN 'SalesRep' THEN SUBSTRING( LastName, 1, 1 )
			WHEN 'CustomerPO' THEN SUBSTRING( CustomerPO, 1, 1 )
		END
		LIKE @JumpTo ORDER BY RowNumber

		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription, CustomerPO, SpindleType
		FROM #tblTempGetProjects
		WHERE (
		CASE @OrderBy
			WHEN 'Customer' THEN SUBSTRING( Customer, 1, 1 )
			WHEN 'ProjectType' THEN SUBSTRING( ProjectType, 1, 1 )
			WHEN 'ProjectName' THEN SUBSTRING( ProjectName, 1, 1 )
			WHEN 'ProjectDescription' THEN SUBSTRING( ProjectDescription, 1, 1 )
			WHEN 'SpindleType' THEN SUBSTRING( SpindleType, 1, 1 )
			WHEN 'SalesRep' THEN SUBSTRING( LastName, 1, 1 )
			WHEN 'CustomerPO' THEN SUBSTRING( CustomerPO, 1, 1 )
		END
		BETWEEN @JumpTo AND 'Z' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription, CustomerPO, SpindleType
		FROM #tblTempGetProjects
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	
	DROP TABLE #tblTempGetProjects
END



GO
/****** Object:  StoredProcedure [dbo].[spGetProjectsByCustomerId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetProjectsByCustomerId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetProjectsByCustomerId] AS' 
END
GO



ALTER  PROCEDURE [dbo].[spGetProjectsByCustomerId]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@CustomerId	int 		= 0

AS

BEGIN

	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetProjectsByStartDate FROM (
	SELECT TOP 100 PERCENT ProjectId, ProjectName, ProjectType, Customer, tblProjects.CustomerId AS CustomerId, 
	( LastName + ', ' + FirstName ) AS SalesRep, StartDate, EstCompletionDate, CompletionDate, 
	ProjectDescription
	FROM tblProjects
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblEmployees ON tblProjects.SalesRepId = tblEmployees.EmployeeId
	LEFT JOIN tblProjectTypes ON tblProjects.ProjectTypeId = tblProjectTypes.ProjectTypeId
	WHERE ( tblProjects.CustomerId = @CustomerId )
	ORDER BY CompletionDate
	) AS X


BEGIN
		SELECT RowNumber, ProjectId, ProjectName, ProjectType, Customer, SalesRep, StartDate, EstCompletionDate, CompletionDate, ProjectDescription
		FROM #tblTempGetProjectsByStartDate
		ORDER BY RowNumber
	END
	
	DROP TABLE #tblTempGetProjectsByStartDate
	
END



GO
/****** Object:  StoredProcedure [dbo].[spGetPurchaseOrderByPurchaseOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetPurchaseOrderByPurchaseOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetPurchaseOrderByPurchaseOrderId] AS' 
END
GO




ALTER PROCEDURE [dbo].[spGetPurchaseOrderByPurchaseOrderId]

@ErrorMessage		varchar(255)	OUTPUT,
@PurchaseOrderId	int

AS

IF( @PurchaseOrderId = 0 )
BEGIN
	SELECT '' AS PurchaseOrderID, '' AS PurchaseOrderNumber, '' AS PurchaseOrderDescription, '' AS SupplierID, '' AS EmployeeID, '' AS OrderDate, '' AS ShippingMethodID, 
	'' AS DateClosed, '' AS DatePromised, '' AS DateRequired, '' AS ShipDate, '' AS DateShippedToSupplier, '' AS TrackingNumber, '' AS InternalNotes, '' AS PODescription
END
ELSE
BEGIN
	SELECT PurchaseOrderID, PurchaseOrderNumber, PurchaseOrderDescription, SupplierID, EmployeeID, OrderDate, ShippingMethodID, DateClosed, DatePromised, DateRequired, ShipDate, DateShippedToSupplier, TrackingNumber, InternalNotes, PODescription
	FROM tblPurchaseOrders WHERE PurchaseOrderID = @PurchaseOrderId
	
	SELECT TransactionId, TransactionDate, tblInventoryTransactions.ProductID, ProductName, TransactionDescription, UnitsOrdered, tblInventoryTransactions.UnitPrice, ( UnitsOrdered * tblInventoryTransactions.UnitPrice ) AS SubTotal
	FROM tblInventoryTransactions
	LEFT JOIN tblProducts ON tblInventoryTransactions.ProductID = tblProducts.ProductID
	WHERE PurchaseOrderID = @PurchaseOrderId
	ORDER BY TransactionDate DESC, TransactionId DESC

	SELECT tblWorkOrderPOs.PurchaseOrderId AS PurchaseOrderId,
	tblWorkOrderPOs.WorkOrderId AS WorkOrderId,
	tblWorkOrders.WorkOrderNumber AS WorkOrderNumber,
	tblWorkOrders.DateIn AS DateIn,
	tblWorkOrders.DateOut AS DateOut,
	tblWorkOrders.SerialNumber AS SerialNumber,
	tblWorkOrders.BoeingWorkOrderNumber AS BoeingWorkOrderNumber,
	tblSpindles.SpindleType AS SpindleType,
	tblSpindles.SpindleId AS SpindleId
	FROM tblWorkOrderPOs
	LEFT JOIN tblWorkOrders ON tblWorkOrderPOs.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE tblWorkOrderPOs.PurchaseOrderId = @PurchaseOrderId
		
END

GO
/****** Object:  StoredProcedure [dbo].[spGetPurchaseOrderList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetPurchaseOrderList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetPurchaseOrderList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetPurchaseOrderList]

AS

SELECT PurchaseOrderID, PurchaseOrderNumber FROM tblPurchaseOrders ORDER BY PurchaseOrderNumber



GO
/****** Object:  StoredProcedure [dbo].[spGetPurchaseOrders]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetPurchaseOrders]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetPurchaseOrders] AS' 
END
GO



ALTER   PROCEDURE [dbo].[spGetPurchaseOrders]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'OrderNumber' OUTPUT,
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(11)	= '%'

AS

SET TEXTSIZE 8000

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'OrderId'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

IF( @OrderBy = 'OrderNumber' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetPurchaseOrdersByOrderNumber FROM (
	SELECT TOP 100 PERCENT tblPurchaseOrders.PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, 
	OrderDate, DatePromised, SupplierName, ( LastName + ', ' + FirstName ) AS BuyerName,
	DateClosed
	FROM tblPurchaseOrders
	LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblEmployees ON tblPurchaseOrders.EmployeeId = tblEmployees.EmployeeId
	WHERE PurchaseOrderNumber LIKE CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	ORDER BY PurchaseOrderNumber
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetPurchaseOrdersByOrderNumber WHERE 
		SUBSTRING( PurchaseOrderNumber, 1, 1 ) LIKE @JumpTo ORDER BY RowNumber
	END
	
	SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
	FROM #tblTempGetPurchaseOrdersByOrderNumber
	WHERE ( SUBSTRING( PurchaseOrderNumber, 1, 1 ) BETWEEN @JumpTo AND 'Z' )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	DROP TABLE #tblTempGetPurchaseOrdersByOrderNumber
END
ELSE IF( @OrderBy = 'SupplierName' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetPurchaseOrdersBySupplierName FROM (
	SELECT TOP 100 PERCENT tblPurchaseOrders.PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, 
	OrderDate, DatePromised, SupplierName, ( LastName + ', ' + FirstName ) AS BuyerName,
	DateClosed
	FROM tblPurchaseOrders
	LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblEmployees ON tblPurchaseOrders.EmployeeId = tblEmployees.EmployeeId
	WHERE SupplierName LIKE CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	ORDER BY SupplierName
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetPurchaseOrdersBySupplierName WHERE 
		SUBSTRING( SupplierName, 1, 1 ) LIKE @JumpTo ORDER BY RowNumber
	END
	
	SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
	FROM #tblTempGetPurchaseOrdersBySupplierName
	WHERE ( SUBSTRING( SupplierName, 1, 1 ) BETWEEN @JumpTo AND 'Z' )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	DROP TABLE #tblTempGetPurchaseOrdersBySupplierName
END
ELSE IF( @OrderBy = 'OrderId' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetPurchaseOrdersByOrderId FROM (
	SELECT TOP 100 PERCENT tblPurchaseOrders.PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, 
	OrderDate, DatePromised, SupplierName, ( LastName + ', ' + FirstName ) AS BuyerName,
	DateClosed
	FROM tblPurchaseOrders
	LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblEmployees ON tblPurchaseOrders.EmployeeId = tblEmployees.EmployeeId
	WHERE PurchaseOrderId LIKE CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	ORDER BY PurchaseOrderId
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetPurchaseOrdersByOrderId
		WHERE PurchaseOrderId BETWEEN CONVERT( int, @JumpTo ) AND 9999999 ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SET @JumpTo = 1
	END

	SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
	FROM #tblTempGetPurchaseOrdersByOrderId
	WHERE ( PurchaseOrderId BETWEEN CONVERT( int, @JumpTo ) AND 9999999 )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	DROP TABLE #tblTempGetPurchaseOrdersByOrderId

	RETURN 0
END
ELSE IF( @OrderBy = 'OrderDate' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetPurchaseOrdersByOrderDate FROM (
	SELECT TOP 100 PERCENT tblPurchaseOrders.PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, 
	OrderDate, DatePromised, SupplierName, ( LastName + ', ' + FirstName ) AS BuyerName,
	DateClosed
	FROM tblPurchaseOrders
	LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblEmployees ON tblPurchaseOrders.EmployeeId = tblEmployees.EmployeeId
	ORDER BY OrderDate
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetPurchaseOrdersByOrderDate
		WHERE OrderDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
		FROM #tblTempGetPurchaseOrdersByOrderDate
		WHERE ( OrderDate BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
		FROM #tblTempGetPurchaseOrdersByOrderDate
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetPurchaseOrdersByOrderDate

	RETURN 0
END
ELSE IF( @OrderBy = 'DatePromised' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetPurchaseOrdersByDatePromised FROM (
	SELECT TOP 100 PERCENT tblPurchaseOrders.PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, 
	OrderDate, DatePromised, SupplierName, ( LastName + ', ' + FirstName ) AS BuyerName,
	DateClosed
	FROM tblPurchaseOrders
	LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblEmployees ON tblPurchaseOrders.EmployeeId = tblEmployees.EmployeeId
	ORDER BY DatePromised
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetPurchaseOrdersByDatePromised
		WHERE DatePromised BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
		FROM #tblTempGetPurchaseOrdersByDatePromised
		WHERE ( DatePromised BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
		FROM #tblTempGetPurchaseOrdersByDatePromised
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetPurchaseOrdersByDatePromised

	RETURN 0
END
ELSE IF( @OrderBy = 'DateClosed' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetPurchaseOrdersByDateClosed FROM (
	SELECT TOP 100 PERCENT tblPurchaseOrders.PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, 
	OrderDate, DatePromised, SupplierName, ( LastName + ', ' + FirstName ) AS BuyerName,
	DateClosed
	FROM tblPurchaseOrders
	LEFT JOIN tblSuppliers ON tblPurchaseOrders.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblEmployees ON tblPurchaseOrders.EmployeeId = tblEmployees.EmployeeId
	ORDER BY DateClosed
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetPurchaseOrdersByDateClosed
		WHERE DateClosed BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
		FROM #tblTempGetPurchaseOrdersByDateClosed
		WHERE ( DateClosed BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SELECT RowNumber, PurchaseOrderId, PurchaseOrderNumber, PurchaseOrderDescription, OrderDate, DatePromised, SupplierName, BuyerName, DateClosed
		FROM #tblTempGetPurchaseOrdersByDateClosed
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetPurchaseOrdersByDateClosed

	RETURN 0
END

GO
/****** Object:  View [dbo].[vwQCInspections]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwQCInspections]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwQCInspections
AS
SELECT     dbo.tblQCInspections.QCInspectionId, dbo.tblQCInspections.RPM_In, dbo.tblQCInspections.RPM_Final, dbo.tblQCInspections.Weight_In, 
                      dbo.tblQCInspections.Weight_Final, dbo.tblQCInspections.Vibration_In, dbo.tblQCInspections.Vibration_Final, dbo.tblQCInspections.VibrationRear_In, 
                      dbo.tblQCInspections.VibrationRear_Final, dbo.tblQCInspections.CoolantFlow_In, dbo.tblQCInspections.CoolantFlow_Final, 
                      dbo.tblQCInspections.GSE_In, dbo.tblQCInspections.GSE_Final, dbo.tblQCInspections.GSERear_In, dbo.tblQCInspections.GSERear_Final, 
                      dbo.tblQCInspections.ShaftTemp_In, dbo.tblQCInspections.ShaftTemp_Final, dbo.tblQCInspections.CoolantTempIncomingSet_In, 
                      dbo.tblQCInspections.CoolantTempIncomingSet_Final, dbo.tblQCInspections.CoolantTempActual_In, dbo.tblQCInspections.CoolantTempActual_Final, 
                      dbo.tblQCInspections.FrontTemp_In, dbo.tblQCInspections.FrontTemp_Final, dbo.tblQCInspections.RearTemp_In, 
                      dbo.tblQCInspections.RearTemp_Final, dbo.tblQCInspections.CoolantPressureIncoming_In, dbo.tblQCInspections.CoolantPressureIncoming_Final, 
                      dbo.tblQCInspections.BreakInTime_In, dbo.tblQCInspections.BreakInTime_Final, dbo.tblQCInspections.CoolingMethod_In, 
                      dbo.tblQCInspections.CoolingMethod_Final, dbo.tblQCInspections.Volts_In, dbo.tblQCInspections.Volts_Final, dbo.tblQCInspections.HP_In, 
                      dbo.tblQCInspections.HP_Final, dbo.tblQCInspections.Amps_In, dbo.tblQCInspections.Amps_Final, dbo.tblQCInspections.Phase_In, 
                      dbo.tblQCInspections.Phase_Final, dbo.tblQCInspections.Hz_In, dbo.tblQCInspections.Hz_Final, dbo.tblQCInspections.Thermistor_In, 
                      dbo.tblQCInspections.Thermistor_Final, dbo.tblQCInspections.Poles_In, dbo.tblQCInspections.Poles_Final, dbo.tblQCInspections.AmpDraw_In, 
                      dbo.tblQCInspections.AmpDraw_Final, dbo.tblQCInspections.ConnectorCtrl_In, dbo.tblQCInspections.ConnectorCtrl_Final, 
                      dbo.tblQCInspections.ConnectorPower_In, dbo.tblQCInspections.ConnectorPower_Final, dbo.tblQCInspections.Converter_In, 
                      dbo.tblQCInspections.Converter_Final, dbo.tblQCInspections.ToolHolder_In, dbo.tblQCInspections.ToolHolder_Final, dbo.tblQCInspections.PullPin_In, 
                      dbo.tblQCInspections.PullPin_Final, dbo.tblQCInspections.EMDimension_In, dbo.tblQCInspections.EMDimension_Final, 
                      dbo.tblQCInspections.EjectionPath_In, dbo.tblQCInspections.EjectionPath_Final, dbo.tblQCInspections.ToolOutPressure_In, 
                      dbo.tblQCInspections.ToolOutPressure_Final, dbo.tblQCInspections.ReturnPressure_In, dbo.tblQCInspections.ReturnPressure_Final, 
                      dbo.tblQCInspections.DrawbarForce_In, dbo.tblQCInspections.DrawbarForce_Final, dbo.tblQCInspections.ToolChangeFunction_In, 
                      dbo.tblQCInspections.ToolChangeFunction_Final, dbo.tblQCInspections.ProximitySwitchFunction_In, 
                      dbo.tblQCInspections.ProximitySwitchFunction_Final, dbo.tblQCInspections.Lubrication_In, dbo.tblQCInspections.Lubrication_Final, 
                      dbo.tblQCInspections.Grease_In, dbo.tblQCInspections.Grease_Final, dbo.tblQCInspections.OilMist_In, dbo.tblQCInspections.OilMist_Final, 
                      dbo.tblQCInspections.OilJet_In, dbo.tblQCInspections.OilJet_Final, dbo.tblQCInspections.OilGreaseType_In, 
                      dbo.tblQCInspections.OilGreaseType_Final, dbo.tblQCInspections.IntervalDPM_In, dbo.tblQCInspections.IntervalDPM_Final, 
                      dbo.tblQCInspections.MainPressure_In, dbo.tblQCInspections.MainPressure_Final, dbo.tblQCInspections.TubePressure_In, 
                      dbo.tblQCInspections.TubePressure_Final, dbo.tblQCInspections.Preload_In, dbo.tblQCInspections.Preload_Final, dbo.tblQCInspections.RadialPlay_In, 
                      dbo.tblQCInspections.RadialPlay_Final, dbo.tblQCInspections.AxialPlay_In, dbo.tblQCInspections.AxialPlay_Final, 
                      dbo.tblQCInspections.RunoutFront_In, dbo.tblQCInspections.RunoutFront_Final, dbo.tblQCInspections.RunoutFrontLocation_In, 
                      dbo.tblQCInspections.RunoutFrontLocation_Final, dbo.tblQCInspections.RunoutFront2_In, dbo.tblQCInspections.RunoutFront2_Final, 
                      dbo.tblQCInspections.RunoutFront2Location_In, dbo.tblQCInspections.RunoutFront2Location_Final, dbo.tblQCInspections.RunoutRear_In, 
                      dbo.tblQCInspections.RunoutRear_Final, dbo.tblQCInspections.RunoutRearLocation_In, dbo.tblQCInspections.RunoutRearLocation_Final, 
                      dbo.tblQCInspections.RunoutRear2_In, dbo.tblQCInspections.RunoutRear2_Final, dbo.tblQCInspections.RunoutRear2Location_In, 
                      dbo.tblQCInspections.RunoutRear2Location_Final, dbo.tblQCInspections.ToolContact_In, dbo.tblQCInspections.ToolContact_Final, 
                      dbo.tblQCInspections.ToolContactRear_In, dbo.tblQCInspections.ToolContactRear_Final, dbo.tblQCInspections.ToolGap_In, 
                      dbo.tblQCInspections.ToolGap_Final, dbo.tblQCInspections.ToolGapRear_In, dbo.tblQCInspections.ToolGapRear_Final, dbo.tblSpindles.SpindleId, 
                      dbo.tblSpindles.SpindleType, dbo.tblSpindles.SpindleCategoryId, dbo.tblSpindles.RPM, dbo.tblSpindles.Weight, dbo.tblSpindles.Vibration, 
                      dbo.tblSpindles.DrawingNumber, dbo.tblSpindles.VibrationRear, dbo.tblSpindles.GSE, dbo.tblSpindles.CoolantFlow, dbo.tblSpindles.ShaftTemp, 
                      dbo.tblSpindles.GSERear, dbo.tblSpindles.CoolantTempIncomingSet, dbo.tblSpindles.CoolantTempActual, dbo.tblSpindles.FrontTemp, 
                      dbo.tblSpindles.RearTemp, dbo.tblSpindles.CoolantPressureIncoming, dbo.tblSpindles.BreakInTime, dbo.tblSpindles.CoolingMethod, 
                      dbo.tblSpindles.Volts, dbo.tblSpindles.HP, dbo.tblSpindles.Amps, dbo.tblSpindles.Phase, dbo.tblSpindles.Hz, dbo.tblSpindles.Thermistor, 
                      dbo.tblSpindles.Poles, dbo.tblSpindles.AmpDraw, dbo.tblSpindles.ConnectorCtrl, dbo.tblSpindles.ConnectorPower, dbo.tblSpindles.Converter, 
                      dbo.tblSpindles.ToolHolder, dbo.tblSpindles.PullPin, dbo.tblSpindles.EMDimension, dbo.tblSpindles.EjectionPath, dbo.tblSpindles.ToolOutPressure, 
                      dbo.tblSpindles.ReturnPressure, dbo.tblSpindles.DrawbarForce, dbo.tblSpindles.ToolChangeFunction, dbo.tblSpindles.ProximitySwitchFunction, 
                      dbo.tblSpindles.Lubrication, dbo.tblSpindles.Grease, dbo.tblSpindles.OilMist, dbo.tblSpindles.OilJet, dbo.tblSpindles.OilGreaseType, 
                      dbo.tblSpindles.IntervalDPM, dbo.tblSpindles.MainPressure, dbo.tblSpindles.TubePressure, dbo.tblSpindles.Preload, dbo.tblSpindles.RadialPlay, 
                      dbo.tblSpindles.AxialPlay, dbo.tblSpindles.RunoutFront, dbo.tblSpindles.RunoutFrontLocation, dbo.tblSpindles.RunoutFront2, 
                      dbo.tblSpindles.RunoutFront2Location, dbo.tblSpindles.RunoutRear, dbo.tblSpindles.RunoutRearLocation, dbo.tblSpindles.RunoutRear2, 
                      dbo.tblSpindles.RunoutRear2Location, dbo.tblSpindles.ToolContact, dbo.tblSpindles.ToolContactRear, dbo.tblSpindles.ToolGap, 
                      dbo.tblSpindles.ToolGapRear, dbo.tblWorkOrders.WorkOrderId, dbo.tblWorkOrders.WorkOrderNumber
FROM         dbo.tblWorkOrders LEFT OUTER JOIN
                      dbo.tblQCInspections ON dbo.tblWorkOrders.WorkOrderId = dbo.tblQCInspections.WorkOrderId LEFT OUTER JOIN
                      dbo.tblProjects ON dbo.tblWorkOrders.ProjectId = dbo.tblProjects.ProjectId LEFT OUTER JOIN
                      dbo.tblSpindles ON dbo.tblProjects.SpindleId = dbo.tblSpindles.SpindleId


' 
GO
/****** Object:  StoredProcedure [dbo].[spGetQCInspectionByWorkOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQCInspectionByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQCInspectionByWorkOrderId] AS' 
END
GO

ALTER   PROCEDURE [dbo].[spGetQCInspectionByWorkOrderId]

@ErrorMessage	varchar(255)	OUTPUT,
@WorkOrderId	int = NULL

AS

	
	BEGIN
		SELECT *
		FROM vwQCInspections
		WHERE vwQCInspections.WorkOrderId = @WorkOrderId
	END

GO
/****** Object:  StoredProcedure [dbo].[spGetQCTempByQCTempId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQCTempByQCTempId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQCTempByQCTempId] AS' 
END
GO


ALTER   PROCEDURE [dbo].[spGetQCTempByQCTempId]

@ErrorMessage	varchar(255)	OUTPUT,
@QCTempId	int = NULL

AS

	IF( ( @QCTempId = 0 ) OR ( @QCTempId IS NULL ) )
	BEGIN
		SELECT '' AS QCTime, '' AS QCSpeed, '' AS QCFront, '' AS QCRear, '' AS QCShaft, '' AS QCTempLocation, '' AS QCNotes
	END
	ELSE
	BEGIN


		SELECT QCTempLogId, QCTempId, QCTime, QCSpeed, QCFront, QCRear, QCShaft, QCTempLocation, QCNotes
		FROM tblQCTemps
		WHERE tblQCTemps.QCTempId = @QCTempId
	END

GO
/****** Object:  StoredProcedure [dbo].[spGetQCTempLogByQCTempLogId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQCTempLogByQCTempLogId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQCTempLogByQCTempLogId] AS' 
END
GO






ALTER   PROCEDURE [dbo].[spGetQCTempLogByQCTempLogId]

@ErrorMessage	varchar(255)	OUTPUT,
@QCTempLogId	int = NULL

AS

	IF( ( @QCTempLogId = 0 ) OR ( @QCTempLogId IS NULL ) )
	BEGIN
		SELECT '' AS QCDate, '' AS QCMaxSpeed, '' AS QCTotalRunTime
	END
	ELSE
	BEGIN
		SELECT tblQCTempLogs.QCTempLogId, tblQCTempLogs.QCDate, tblQCTempLogs.QCMaxSpeed, tblQCTempLogs.QCTotalRunTime,
		tblQCTempLogs.WorkOrderId AS WorkOrderId, tblWorkOrders.WorkOrderNumber
		FROM tblQCTempLogs
		LEFT JOIN tblWorkOrders ON tblQCTempLogs.WorkOrderId = tblWorkOrders.WorkOrderId
		WHERE tblQCTempLogs.QCTempLogId = @QCTempLogId


		SELECT QCTempLogId, QCTempId, QCTime, QCSpeed, QCFront, QCRear, QCShaft, QCTempLocation, QCNotes
		FROM tblQCTemps
		WHERE tblQCTemps.QCTempLogId = @QCTempLogId
	END

GO
/****** Object:  StoredProcedure [dbo].[spGetQCTempLogsByWorkOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQCTempLogsByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQCTempLogsByWorkOrderId] AS' 
END
GO

ALTER   PROCEDURE [dbo].[spGetQCTempLogsByWorkOrderId]

@ErrorMessage	varchar(255)	OUTPUT,
@WorkOrderId	int = NULL

AS

	
	BEGIN
		SELECT tblQCTempLogs.QCTempLogId, tblQCTempLogs.QCDate, tblQCTempLogs.QCMaxSpeed, tblQCTempLogs.QCTotalRunTime
		FROM tblQCTempLogs
		WHERE tblQCTempLogs.WorkOrderId = @WorkOrderId
	END

GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteBearingDetailsByQuoteBearingId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteBearingDetailsByQuoteBearingId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteBearingDetailsByQuoteBearingId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetQuoteBearingDetailsByQuoteBearingId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteBearingId	int = NULL

AS


BEGIN
	
	
	SELECT QuoteBearingId, tblProducts.ProductName, tblProducts.ProductId, SupplierName,
	ISNULL( BearingCost, 0 ) AS BearingCost, BearingMarkup AS Markup, Qty, tblQuoteBearings.SupplierId
	FROM tblQuoteBearings
	LEFT JOIN tblSuppliers ON tblQuoteBearings.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblProducts ON tblQuoteBearings.ProductId = tblProducts.ProductId
	WHERE tblQuoteBearings.QuoteBearingId = @QuoteBearingId
END

GO
/****** Object:  View [dbo].[vwQuotedBearings]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwQuotedBearings]'))
EXEC dbo.sp_executesql @statement = N'
CREATE VIEW dbo.vwQuotedBearings
AS
SELECT     TOP 100 PERCENT dbo.tblProducts.ProductID, SUM(ISNULL(dbo.tblQuoteBearings.Qty, 0)) AS QtyQuoted, dbo.tblProducts.PartNumber
FROM         dbo.tblProducts INNER JOIN
                      dbo.tblQuoteBearings ON dbo.tblProducts.ProductID = dbo.tblQuoteBearings.ProductId INNER JOIN
                      dbo.tblQuotes ON dbo.tblQuoteBearings.QuoteId = dbo.tblQuotes.QuoteId
WHERE     (dbo.tblQuotes.DateApproved IS NULL)
GROUP BY dbo.tblProducts.ProductID, dbo.tblProducts.PartNumber, dbo.tblProducts.PartNumber
ORDER BY dbo.tblProducts.ProductID

' 
GO
/****** Object:  View [dbo].[vwQuotedParts]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwQuotedParts]'))
EXEC dbo.sp_executesql @statement = N'
CREATE VIEW dbo.vwQuotedParts
AS
SELECT     TOP 100 PERCENT dbo.tblProducts.ProductID, SUM(ISNULL(dbo.tblQuoteParts.Qty, 0)) AS QtyQuoted, dbo.tblProducts.PartNumber
FROM         dbo.tblProducts INNER JOIN
                      dbo.tblQuoteParts ON dbo.tblProducts.ProductID = dbo.tblQuoteParts.ProductId INNER JOIN
                      dbo.tblQuotes ON dbo.tblQuoteParts.QuoteId = dbo.tblQuotes.QuoteId
WHERE     (dbo.tblQuotes.DateApproved IS NULL)
GROUP BY dbo.tblProducts.ProductID, dbo.tblProducts.PartNumber, dbo.tblProducts.PartNumber
ORDER BY dbo.tblProducts.ProductID

' 
GO
/****** Object:  View [dbo].[vwQuotedProducts]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwQuotedProducts]'))
EXEC dbo.sp_executesql @statement = N'
CREATE VIEW dbo.vwQuotedProducts
AS
SELECT     ProductID, PartNumber, SUM(QtyQuoted) AS QtyQuoted
FROM         (SELECT     *
                       FROM          vwQuotedParts
                       UNION ALL
                       SELECT     *
                       FROM         vwQuotedBearings) QUOTED
GROUP BY ProductID, PartNumber

' 
GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteBearingsByQuoteId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteBearingsByQuoteId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteBearingsByQuoteId] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetQuoteBearingsByQuoteId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteId	int = NULL,
@IsPrint	bit = 0

AS

IF( @IsPrint = 0 )
BEGIN
	IF( ( @QuoteId = 0 ) OR ( @QuoteId IS NULL ) )
	BEGIN
		SELECT '' AS BearingCommission, 0 AS FreightChargeBearings, '' AS TotalBearingCost, '' AS TotalBearingCharge
	
		/*SELECT '' AS QuoteBearingId, '' AS BearingId, '' AS BearingSourceId, '' AS ProductId, '' AS CenBearingCode, '' AS CSSPrice, '' AS BearingCost, '' AS Markup*/
	END
	ELSE
	BEGIN
		SELECT MAX( BearingCommission ) AS BearingCommission , MAX( ISNULL( BearingFreightCharge, 0 ) ) AS BearingFreightCharge, 
		( SUM( ISNULL( BearingCost, 0 ) * ISNULL( Qty, 0 ) ) ) AS TotalBearingCost, 
		( SUM( ISNULL( BearingCost, 0 ) * ISNULL( BearingMarkup, 0 ) * ISNULL( Qty, 0 ) ) ) + ISNULL( MAX( BearingFreightCharge ), 0 )  AS TotalBearingCharge,
		MAX( SpindleType ) AS SpindleType, MAX(tblProjects.SpindleId) AS SpindleId
		FROM tblQuotes 
		LEFT JOIN tblQuoteBearings ON tblQuotes.QuoteId = tblQuoteBearings.QuoteId
		LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
		LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
		LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
		WHERE tblQuotes.QuoteId = @QuoteId
	
		SELECT QuoteBearingId, tblQuoteBearings.ProductId AS ProductId, SupplierName,
		ISNULL( BearingCost, 0 ) AS BearingCost, BearingMarkup AS Markup, Qty, CenBearingCode, ISNULL( CSSPrice, 0 ) AS CSSPrice,
		ISNULL (OnHand,0) AS OnHand, ISNULL (vwOnHandProducts.ReorderLevel, 0) AS ReorderLevel, ISNULL (vwQuotedProducts.QtyQuoted, 0) AS OnOrder,
		tblProducts.ProductName, tblProducts.ProductDescription, tblProducts.PartNumber
		FROM tblQuoteBearings
		LEFT JOIN tblSuppliers ON tblQuoteBearings.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN vwOnHandProducts ON tblQuoteBearings.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN vwQuotedProducts ON tblQuoteBearings.ProductId = vwQuotedProducts.ProductId
		LEFT JOIN tblProducts ON tblQuoteBearings.ProductId = tblProducts.ProductId
		WHERE tblQuoteBearings.QuoteId = @QuoteId
	END
END
ELSE
BEGIN
	SELECT WorkOrderNumber FROM tblQuotes WHERE tblQuotes.QuoteId = @QuoteId

	SELECT tblQuoteBearings.ProductName AS CenBearingCode, tblQuoteBearings.ProductId, BearingDescription, SupplierName, ISNULL( UnitPrice, 0 ) AS CSSPrice, ISNULL( BearingCost, 0 ) AS BearingCost
	FROM tblQuoteBearings
	LEFT JOIN tblSuppliers ON tblQuoteBearings.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblProducts ON tblQuoteBearings.ProductId = tblProducts.ProductId
	WHERE tblQuoteBearings.QuoteId = @QuoteId
END

GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteBearingsByWorkOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteBearingsByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteBearingsByWorkOrderId] AS' 
END
GO

ALTER PROCEDURE [dbo].[spGetQuoteBearingsByWorkOrderId]
@ErrorMessage	varchar(255) = NULL OUTPUT,
@WorkOrderId	int = 0

AS

SELECT QuoteBearingId, tblQuoteBearings.ProductId AS ProductId, SupplierName, Qty,
		ISNULL (OnHand,0) AS OnHand, ISNULL (vwOnHandProducts.ReorderLevel, 0) AS ReorderLevel,
		tblProducts.PartNumber, tblProducts.ProductName, tblProducts.ProductDescription
		FROM tblQuoteBearings
		LEFT JOIN tblSuppliers ON tblQuoteBearings.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN tblQuotes ON tblQuoteBearings.QuoteId = tblQuotes.QuoteId
		LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
		LEFT JOIN vwOnHandProducts ON tblQuoteBearings.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN tblProducts ON tblQuoteBearings.ProductId = tblProducts.ProductId
		WHERE tblWorkOrders.WorkOrderId = @WorkOrderId

GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteBearingsList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteBearingsList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteBearingsList] AS' 
END
GO

ALTER PROCEDURE [dbo].[spGetQuoteBearingsList]
@ErrorMessage	varchar(255) = NULL OUTPUT,
@QuoteId	int = 0

AS

SELECT QuoteBearingId, tblQuoteBearings.ProductId AS ProductId, SupplierName, Qty,
		ISNULL (OnHand,0) AS OnHand, ISNULL (vwOnHandProducts.ReorderLevel, 0) AS ReorderLevel,
		tblProducts.PartNumber, tblProducts.ProductName, tblProducts.ProductDescription
		FROM tblQuoteBearings
		LEFT JOIN tblSuppliers ON tblQuoteBearings.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN tblQuotes ON tblQuoteBearings.QuoteId = tblQuotes.QuoteId
		LEFT JOIN vwOnHandProducts ON tblQuoteBearings.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN tblProducts ON tblQuoteBearings.ProductId = tblProducts.ProductId
		WHERE tblQuotes.QuoteId = @QuoteId

GO
/****** Object:  Table [dbo].[tblMultipliers]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblMultipliers]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[tblMultipliers](
	[multiplier_id] [int] IDENTITY(1,1) NOT NULL,
	[name] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[value] [numeric](10, 2) NULL,
 CONSTRAINT [PK_tblMultipliers] PRIMARY KEY CLUSTERED 
(
	[multiplier_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
END
GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteByQuoteId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteByQuoteId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteByQuoteId] AS' 
END
GO





ALTER PROCEDURE [dbo].[spGetQuoteByQuoteId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteId	int	= 0,
@CustomerId	int OUTPUT

AS

DECLARE @TotalLaborCost AS money
DECLARE @TotalBearingCharge AS money
DECLARE @TotalPartCharge AS money
DECLARE @TotalReworkCharge AS money
/*addition of multiplier table - cen 20070922*/
DECLARE @MultiplierLabor	numeric(10,2)		SET @MultiplierLabor = 0.00
DECLARE @MultiplierParts	numeric(10,2)		SET @MultiplierParts = 0.00
DECLARE @MultiplierBearings	numeric(10,2)		SET @MultiplierBearings = 0.00
DECLARE @MultiplierRework	numeric(10,2)		SET @MultiplierRework = 0.00

SET NOCOUNT ON 

IF( @QuoteId = 0 )
BEGIN
	SELECT '' AS QuoteId, '' AS CustomerId, '' AS Customer, '' AS ProjectId, '' AS ProjectName, '' AS WorkOrderId, '' AS WorkOrderNumber, '' AS SpindleId, '' AS SpindleType, '' AS SerialNumber, 
	0 AS FreightChargeParts, 0 AS FreightChargeSubWork, '' AS FreightLBS, 0 AS FreightChargeSub, '' AS ExpDeliveryDate, '' AS Notes, '' AS QuoteSpecificComments,
	0 AS PartsCommission, 0 AS BearingFreightCharge, 0 AS BearingCommission, '' AS DateApproved, getdate() AS DateQuoted, 0 AS TotalLaborCost, 
	0 AS TotalBearingCharge, 0 AS TotalPartCharge, 0 AS TotalReworkCharge, 0 AS TotalFreightCharge, 0 AS FreightChargeSubWork, 0 AS TotalCommissionCharge, 
	0 AS TotalRepairPrice, 0 AS TotalExpeditingCharge, 0 AS QuotedById, '' AS DeliveryInformation, '' AS ExpeditedDeliveryInformation, 0 AS QuotedById, 25.00 AS HandlingCharge
END
ELSE
BEGIN
	BEGIN
		SELECT @MultiplierLabor = value FROM tblMultipliers WHERE name = 'labor'
		SELECT @MultiplierParts = value FROM tblMultipliers WHERE name = 'parts'
		SELECT @MultiplierBearings = value FROM tblMultipliers WHERE name = 'bearings'
		SELECT @MultiplierRework = value FROM tblMultipliers WHERE name = 'rework'
	END
	SELECT @TotalLaborCost = ROUND ( ( ( MAX( ISNULL( HoursDisassembly, 0 ) + ISNULL( HoursCleanAndInspect, 0 ) + ISNULL( GrindingHours, 0 ) + 
	ISNULL( BalancingHours, 0 ) + ISNULL( ElectricalHours, 0 ) + ISNULL( GreaseHours, 0 ) + ISNULL( AssemblyAndTestHours, 0 ) 
	+ ISNULL( MscWorkHours, 0 ) ) * @MultiplierLabor ) ), 0 )
	FROM tblQuotes
	WHERE tblQuotes.QuoteId = @QuoteId

	SELECT @TotalBearingCharge = ROUND ( ( ( SUM( ISNULL( BearingCost, 0 ) * ISNULL( Qty, 0 ) * ISNULL( BearingMarkup, 0 ) ) * @MultiplierBearings ) + MAX( ISNULL( BearingFreightCharge, 0 ) ) ), 0 )
	FROM tblQuotes
	LEFT JOIN tblQuoteBearings ON tblQuotes.QuoteId = tblQuoteBearings.QuoteId
	WHERE tblQuotes.QuoteId = @QuoteId
	
	SELECT @TotalPartCharge = ROUND ( ( ( SUM( ISNULL( PartCost, 0 ) * ISNULL( Qty, 0 ) * ISNULL( Markup, 0 ) ) * @MultiplierParts ) + MAX( ISNULL( FreightChargeParts, 0 ) ) ), 0 )
	FROM tblQuotes
	LEFT JOIN tblQuoteParts ON tblQuotes.QuoteId = tblQuoteParts.QuoteId
	WHERE tblQuotes.QuoteId = @QuoteId

	SELECT @TotalReworkCharge = ROUND ( ( ( SUM( ISNULL( SubWorkCost, 0 ) * @MultiplierRework ) ) + MAX( ISNULL( FreightChargeSubWork, 0 ) ) + MAX( ISNULL( FreightChargeSub, 0 ) ) ), 0 )
	FROM tblQuotes
	LEFT JOIN tblQuoteSubWork ON tblQuotes.QuoteId = tblQuoteSubWork.QuoteId
	WHERE tblQuotes.QuoteId = @QuoteId

	SELECT QuoteId, tblCustomers.CustomerId, tblCustomers.Customer, tblProjects.ProjectId, tblProjects.ProjectName,  tblQuotes.WorkOrderId, tblQuotes.WorkOrderNumber, tblSpindles.SpindleId AS SpindleId, SpindleType, tblWorkOrders.SerialNumber, ISNULL( FreightChargeParts, 0 ) AS FreightChargeParts, 
	ISNULL( FreightChargeSubWork, 0 ) AS FreightChargeSubWork, FreightLBS, ISNULL( FreightChargeSub, 0 ) AS FreightChargeSub, ExpDeliveryDate, tblQuotes.Notes, tblQuotes.QuoteSpecificComments, HandlingCharge,
	(ISNULL( FreightChargeSub, 0) + ISNULL( FreightChargeSub1, 0) + ISNULL( BearingFreightCharge, 0) + ISNULL( FreightChargeParts, 0)) AS TotalFreightCharge,
	ISNULL( PartsCommission, .1 ) AS PartsCommission, ISNULL( BearingFreightCharge, 0 ) AS BearingFreightCharge, ISNULL( BearingCommission, .1 ) AS BearingCommission, 
	ISNULL( LaborCommission, .1 ) AS LaborCommission, ISNULL( SubWorkCommission, 0 ) AS SubWorkCommission,
	DateApproved, ISNULL( DateQuoted, '' ) AS DateQuoted, @TotalLaborCost AS TotalLaborCost, @TotalLaborCost + HandlingCharge AS TotalLaborAndHandling, @TotalBearingCharge AS TotalBearingCharge, 
	@TotalPartCharge AS TotalPartCharge, @TotalReworkCharge AS TotalReworkCharge, ( ISNULL( FreightChargeSub, 0 ) + ISNULL( FreightChargeSub1, 0 ) ) AS FreightChargeSubWork, @TotalReworkCharge + ( ISNULL( FreightChargeSub, 0 ) + ISNULL( FreightChargeSub1, 0 ) ) AS TotalReworkWithFreightCharge,
	( @TotalLaborCost * ISNULL( LaborCommission, .1 ) ) + ( @TotalPartCharge * ISNULL( PartsCommission, .1 ) ) + ( @TotalReworkCharge * ISNULL( SubWorkCommission, .1 ) ) + ( @TotalBearingCharge * ISNULL( BearingCommission, .1 ) ) AS TotalCommissionCharge,
	@TotalLaborCost + @TotalPartCharge + @TotalReworkCharge + @TotalBearingCharge + HandlingCharge AS TotalRepairPrice,
	ROUND ( ( ( @TotalLaborCost * 0.4 ) + ( @TotalReworkCharge * 0.3 ) + ( @TotalLaborCost + @TotalPartCharge + @TotalReworkCharge + @TotalBearingCharge ) ), 0 ) + HandlingCharge AS TotalExpeditingCharge,
	QuoteContactId, PONumber, DeliveryInformation, ExpeditedDeliveryInformation, QuotedById, tblEmployees.FirstName + ' ' + tblEmployees.LastName AS QuotedBy,
	tblCustomerContacts.Contact, tblCustomerContacts.TelephoneNumber, tblCustomerContacts.FaxNumber
	FROM tblQuotes 
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	LEFT JOIN tblCustomerContacts ON tblQuotes.QuoteContactId = tblCustomerContacts.CustomerContactId
	LEFT JOIN tblEmployees ON tblQuotes.QuotedById = tblEmployees.EmployeeId
	WHERE tblQuotes.QuoteId = @QuoteId


	SELECT @CustomerId = tblProjects.CustomerId
	FROM tblQuotes 
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	WHERE tblQuotes.QuoteId = @QuoteId
END


GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteByQuoteIdNoRounding]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteByQuoteIdNoRounding]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteByQuoteIdNoRounding] AS' 
END
GO






ALTER   PROCEDURE [dbo].[spGetQuoteByQuoteIdNoRounding]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteId	int	= 0,
@CustomerId	int OUTPUT

AS

DECLARE @TotalLaborCost AS money
DECLARE @TotalBearingCharge AS money
DECLARE @TotalPartCharge AS money
DECLARE @TotalReworkCharge AS money
/*addition of multiplier table - cen 20070922*/
DECLARE @MultiplierLabor	numeric(10,2)		SET @MultiplierLabor = 0.00
DECLARE @MultiplierParts	numeric(10,2)		SET @MultiplierParts = 0.00
DECLARE @MultiplierBearings	numeric(10,2)		SET @MultiplierBearings = 0.00
DECLARE @MultiplierRework	numeric(10,2)		SET @MultiplierRework = 0.00

SET NOCOUNT ON 

IF( @QuoteId = 0 )
BEGIN
	SELECT '' AS QuoteId, '' AS CustomerId, '' AS Customer, '' AS ProjectId, '' AS ProjectName, '' AS WorkOrderId, '' AS WorkOrderNumber, '' AS SpindleId, '' AS SpindleType, '' AS SerialNumber, 
	0 AS FreightChargeParts, 0 AS FreightChargeSubWork, '' AS FreightLBS, 0 AS FreightChargeSub, '' AS ExpDeliveryDate, '' AS Notes, '' AS QuoteSpecificComments,
	0 AS PartsCommission, 0 AS BearingFreightCharge, 0 AS BearingCommission, '' AS DateApproved, getdate() AS DateQuoted, 0 AS TotalLaborCost, 
	0 AS TotalBearingCharge, 0 AS TotalPartCharge, 0 AS TotalReworkCharge, 0 AS TotalFreightCharge, 0 AS FreightChargeSubWork, 0 AS TotalCommissionCharge, 
	0 AS TotalRepairPrice, 0 AS TotalExpeditingCharge, 0 AS QuotedById, '' AS DeliveryInformation, '' AS ExpeditedDeliveryInformation, 0 AS QuotedById
END
ELSE
BEGIN
	BEGIN
		SELECT @MultiplierLabor = value FROM tblMultipliers WHERE name = 'labor'
		SELECT @MultiplierParts = value FROM tblMultipliers WHERE name = 'parts'
		SELECT @MultiplierBearings = value FROM tblMultipliers WHERE name = 'bearings'
		SELECT @MultiplierRework = value FROM tblMultipliers WHERE name = 'rework'
	END
	SELECT @TotalLaborCost = ( ( MAX( ISNULL( HoursDisassembly, 0 ) + ISNULL( HoursCleanAndInspect, 0 ) + ISNULL( GrindingHours, 0 ) + 
	ISNULL( BalancingHours, 0 ) + ISNULL( ElectricalHours, 0 ) + ISNULL( GreaseHours, 0 ) + ISNULL( AssemblyAndTestHours, 0 ) 
	+ ISNULL( MscWorkHours, 0 ) ) * @MultiplierLabor ) )
	FROM tblQuotes
	WHERE tblQuotes.QuoteId = @QuoteId

	SELECT @TotalBearingCharge =  ( ( SUM( ISNULL( BearingCost, 0 ) * ISNULL( Qty, 0 ) * ISNULL( BearingMarkup, 0 ) ) * @MultiplierBearings ) + MAX( ISNULL( BearingFreightCharge, 0 ) ) )
	FROM tblQuotes
	LEFT JOIN tblQuoteBearings ON tblQuotes.QuoteId = tblQuoteBearings.QuoteId
	WHERE tblQuotes.QuoteId = @QuoteId
	
	SELECT @TotalPartCharge =  ( ( SUM( ISNULL( PartCost, 0 ) * ISNULL( Qty, 0 ) * ISNULL( Markup, 0 ) ) * @MultiplierParts ) + MAX( ISNULL( FreightChargeParts, 0 ) ) )
	FROM tblQuotes
	LEFT JOIN tblQuoteParts ON tblQuotes.QuoteId = tblQuoteParts.QuoteId
	WHERE tblQuotes.QuoteId = @QuoteId

	SELECT @TotalReworkCharge =  ( ( SUM( ISNULL( SubWorkCost, 0 ) * @MultiplierRework ) ) + MAX( ISNULL( FreightChargeSubWork, 0 ) ) + MAX( ISNULL( FreightChargeSub, 0 ) ) )
	FROM tblQuotes
	LEFT JOIN tblQuoteSubWork ON tblQuotes.QuoteId = tblQuoteSubWork.QuoteId
	WHERE tblQuotes.QuoteId = @QuoteId

	SELECT QuoteId, tblCustomers.CustomerId, tblCustomers.Customer, tblProjects.ProjectId, tblProjects.ProjectName,  tblQuotes.WorkOrderId, tblQuotes.WorkOrderNumber, tblSpindles.SpindleId AS SpindleId, SpindleType, tblWorkOrders.SerialNumber, ISNULL( FreightChargeParts, 0 ) AS FreightChargeParts, 
	ISNULL( FreightChargeSubWork, 0 ) AS FreightChargeSubWork, FreightLBS, ISNULL( FreightChargeSub, 0 ) AS FreightChargeSub, ExpDeliveryDate, tblQuotes.Notes, tblQuotes.QuoteSpecificComments,
	(ISNULL( FreightChargeSub, 0) + ISNULL( FreightChargeSub1, 0) + ISNULL( BearingFreightCharge, 0) + ISNULL( FreightChargeParts, 0)) AS TotalFreightCharge,
	ISNULL( PartsCommission, .1 ) AS PartsCommission, ISNULL( BearingFreightCharge, 0 ) AS BearingFreightCharge, ISNULL( BearingCommission, .1 ) AS BearingCommission, 
	ISNULL( LaborCommission, .1 ) AS LaborCommission, ISNULL( SubWorkCommission, 0 ) AS SubWorkCommission,
	DateApproved, ISNULL( DateQuoted, '' ) AS DateQuoted, @TotalLaborCost AS TotalLaborCost, @TotalBearingCharge AS TotalBearingCharge, 
	@TotalPartCharge AS TotalPartCharge, @TotalReworkCharge AS TotalReworkCharge, ( ISNULL( FreightChargeSub, 0 ) + ISNULL( FreightChargeSub1, 0 ) ) AS FreightChargeSubWork, @TotalReworkCharge + ( ISNULL( FreightChargeSub, 0 ) + ISNULL( FreightChargeSub1, 0 ) ) AS TotalReworkWithFreightCharge,
	( @TotalLaborCost * ISNULL( LaborCommission, .1 ) ) + ( @TotalPartCharge * ISNULL( PartsCommission, .1 ) ) + ( @TotalReworkCharge * ISNULL( SubWorkCommission, .1 ) ) + ( @TotalBearingCharge * ISNULL( BearingCommission, .1 ) ) AS TotalCommissionCharge,
	@TotalLaborCost + @TotalPartCharge + @TotalReworkCharge + @TotalBearingCharge AS TotalRepairPrice,
	( ( @TotalLaborCost * 0.4 ) + ( @TotalReworkCharge * 0.3 ) + ( @TotalLaborCost + @TotalPartCharge + @TotalReworkCharge + @TotalBearingCharge ) ) AS TotalExpeditingCharge,
	QuoteContactId, PONumber, DeliveryInformation, ExpeditedDeliveryInformation, QuotedById,
	tblCustomerContacts.Contact, tblCustomerContacts.TelephoneNumber, tblCustomerContacts.FaxNumber
	FROM tblQuotes 
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	LEFT JOIN tblCustomerContacts ON tblQuotes.QuoteContactId = tblCustomerContacts.CustomerContactId
	WHERE tblQuotes.QuoteId = @QuoteId


	SELECT @CustomerId = tblProjects.CustomerId
	FROM tblQuotes 
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	WHERE tblQuotes.QuoteId = @QuoteId
END





GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteLaborByQuoteId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteLaborByQuoteId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteLaborByQuoteId] AS' 
END
GO





ALTER PROCEDURE [dbo].[spGetQuoteLaborByQuoteId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteId	int	= 0

AS

DECLARE 
@MultiplierLabor	numeric(10,2)		SET @MultiplierLabor = 0.00

IF( @QuoteId = 0 )
BEGIN
	SELECT '' AS QuoteId, '' AS WorkOrderNumber, '' AS DisassemblyEvaluation, '' AS HoursDisassembly, '' AS CleanAndInspect, '' AS HoursCleanAndInspect, '' AS InhouseGrinding, 
	'' AS GrindingHours, '' AS Balancing, '' AS BalancingHours, '' AS ElectricalWork, '' AS ElectricalHours, '' AS GreaseBearings, '' AS GreaseHours, 
	'' AS AssemblyAndTest, '' AS AssemblyAndTestHours, '' AS MscWorkNeeded, '' AS MscWorkHours, '' AS TotalLaborHours, '' AS TotalLaborCost, '' AS SpindleType, '' AS LaborCommission
END
ELSE
BEGIN
	BEGIN
		SELECT @MultiplierLabor = value FROM tblMultipliers WHERE name = 'labor'
	END
	SELECT QuoteId, tblQuotes.WorkOrderNumber, DisassemblyEvaluation, HoursDisassembly, CleanAndInspect, HoursCleanAndInspect, InhouseGrinding, GrindingHours, Balancing, BalancingHours, 
	ElectricalWork, ElectricalHours, SpacerPreparation, SpacerPreparationHours, GreaseBearings, GreaseHours, AssemblyAndTest, AssemblyAndTestHours, MscWorkNeeded, MscWorkHours,
	( ISNULL( HoursDisassembly, 0 ) + ISNULL( HoursCleanAndInspect, 0 ) + ISNULL( GrindingHours, 0 ) + ISNULL( BalancingHours, 0 ) + ISNULL( ElectricalHours, 0 )
	 + ISNULL( GreaseHours, 0 ) + ISNULL( AssemblyAndTestHours, 0 ) + ISNULL( MscWorkHours, 0 ) ) AS TotalLaborHours,
	( ( ISNULL( HoursDisassembly, 0 ) + ISNULL( HoursCleanAndInspect, 0 ) + ISNULL( GrindingHours, 0 ) + ISNULL( BalancingHours, 0 ) + ISNULL( ElectricalHours, 0 )
	 + ISNULL( GreaseHours, 0 ) + ISNULL( AssemblyAndTestHours, 0 ) + ISNULL( MscWorkHours, 0 ) ) * @MultiplierLabor ) AS TotalLaborCost,
	SpindleType, tblProjects.SpindleId AS SpindleId, LaborCommission
	FROM tblQuotes 
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE tblQuotes.QuoteId = @QuoteId
END




GO
/****** Object:  StoredProcedure [dbo].[spGetQuotePartDetailsByQuotePartId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuotePartDetailsByQuotePartId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuotePartDetailsByQuotePartId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetQuotePartDetailsByQuotePartId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuotePartId	int = NULL


AS

BEGIN	
	SELECT QuotePartId, tblProducts.ProductId, ISNULL( PartCost, 0 ) AS PartCost, Markup, SupplierName, Qty, tblQuoteParts.SupplierId,
	tblProducts.ProductName, tblProducts.ProductDescription, tblProducts.PartNumber
	FROM tblQuoteParts
	LEFT JOIN tblSuppliers ON tblQuoteParts.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblProducts ON tblQuoteParts.ProductId = tblProducts.ProductId
	WHERE tblQuoteParts.QuotePartId = @QuotePartId
END

GO
/****** Object:  StoredProcedure [dbo].[spGetQuotePartsByQuoteId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuotePartsByQuoteId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuotePartsByQuoteId] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetQuotePartsByQuoteId]

@ErrorMessage	varchar(255)	OUTPUT,
@QuoteId	int = NULL,
@IsPrint	bit = 0

AS

IF( @IsPrint = 0 )
BEGIN
	IF( ( @QuoteId = 0 ) OR ( @QuoteId IS NULL ) )
	BEGIN
		SELECT '' AS PartsCommission, 0 AS FreightChargeParts, 0 AS TotalPartCost, 0 AS TotalPartCharge
	END
	ELSE
	BEGIN
		SELECT MAX( PartsCommission ) AS PartsCommission , MAX( ISNULL( FreightChargeParts, 0 ) ) AS FreightChargeParts, 
		( SUM( ISNULL( PartCost, 0 ) * ISNULL( Qty, 0 ) ) ) AS TotalPartCost, 
		( SUM( ISNULL( PartCost, 0 ) * ISNULL( Markup, 0 ) * ISNULL( Qty, 0 ) ) ) + ISNULL( MAX( FreightChargeParts ), 0 )  AS TotalPartCharge,
		MAX( SpindleType ) AS SpindleType, MAX(tblProjects.SpindleId) AS SpindleId
		FROM tblQuotes 
		LEFT JOIN tblQuoteParts ON tblQuotes.QuoteId = tblQuoteParts.QuoteId
		LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
		LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
		LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
		WHERE tblQuotes.QuoteId = @QuoteId
	
		SELECT QuotePartId, tblQuoteParts.ProductId AS ProductId, ISNULL( PartCost, 0 ) AS PartCost, Markup, SupplierName, Qty,
		ISNULL (OnHand,0) AS OnHand, ISNULL (vwOnHandProducts.ReorderLevel, 0) AS ReorderLevel, ISNULL (vwQuotedProducts.QtyQuoted, 0) AS OnOrder,
		tblProducts.PartNumber, tblProducts.ProductName, tblProducts.ProductDescription
		FROM tblQuoteParts
		LEFT JOIN tblSuppliers ON tblQuoteParts.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN vwOnHandProducts ON tblQuoteParts.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN vwQuotedProducts ON tblQuoteParts.ProductId = vwQuotedProducts.ProductId
		LEFT JOIN tblProducts ON tblQuoteParts.ProductId = tblProducts.ProductId
		WHERE tblQuoteParts.QuoteId = @QuoteId
	END
END
ELSE
BEGIN
	SELECT WorkOrderNumber FROM tblQuotes WHERE tblQuotes.QuoteId = @QuoteId

	SELECT tblQuoteParts.ProductId, SupplierName, tblProducts.PartNumber, tblProducts.ProductName, tblProducts.ProductDescription
	FROM tblQuoteParts
	LEFT JOIN tblSuppliers ON tblQuoteParts.SupplierId = tblSuppliers.SupplierId
	LEFT JOIN tblProducts ON tblQuoteParts.ProductId = tblProducts.ProductId
	WHERE tblQuoteParts.QuoteId = @QuoteId
END

GO
/****** Object:  StoredProcedure [dbo].[spGetQuotePartsByWorkOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuotePartsByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuotePartsByWorkOrderId] AS' 
END
GO

ALTER PROCEDURE [dbo].[spGetQuotePartsByWorkOrderId]
@ErrorMessage	varchar(255) = NULL OUTPUT,
@WorkOrderId	int = 0

AS

SELECT QuotePartId, tblQuoteParts.ProductId AS ProductId, SupplierName, Qty,
		ISNULL (OnHand,0) AS OnHand, ISNULL (vwOnHandProducts.ReorderLevel, 0) AS ReorderLevel,
		tblProducts.PartNumber, tblProducts.ProductName, tblProducts.ProductDescription
		FROM tblQuoteParts
		LEFT JOIN tblSuppliers ON tblQuoteParts.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN tblQuotes ON tblQuoteParts.QuoteId = tblQuotes.QuoteId
		LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
		LEFT JOIN vwOnHandProducts ON tblQuoteParts.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN tblProducts ON tblQuoteParts.ProductId = tblProducts.ProductId
		WHERE tblWorkOrders.WorkOrderId = @WorkOrderId

GO
/****** Object:  StoredProcedure [dbo].[spGetQuotePartsList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuotePartsList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuotePartsList] AS' 
END
GO

ALTER PROCEDURE [dbo].[spGetQuotePartsList]
@ErrorMessage	varchar(255) = NULL OUTPUT,
@QuoteId	int = 0

AS

SELECT QuotePartId, tblQuoteParts.ProductId AS ProductId, SupplierName, Qty,
		ISNULL (OnHand,0) AS OnHand, ISNULL (vwOnHandProducts.ReorderLevel, 0) AS ReorderLevel,
		tblProducts.PartNumber, tblProducts.ProductName, tblProducts.ProductDescription
		FROM tblQuoteParts
		LEFT JOIN tblSuppliers ON tblQuoteParts.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN tblQuotes ON tblQuoteParts.QuoteId = tblQuotes.QuoteId
		LEFT JOIN vwOnHandProducts ON tblQuoteParts.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN tblProducts ON tblQuoteParts.ProductId = tblProducts.ProductId
		WHERE tblQuotes.QuoteId = @QuoteId

GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteSubWorkByQuoteId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteSubWorkByQuoteId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteSubWorkByQuoteId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetQuoteSubWorkByQuoteId]

@ErrorMessage	varchar(255) OUTPUT,
@QuoteId	int = NULL,
@IsPrint	bit = 0

AS

IF( @IsPrint = 0 )
BEGIN
	IF( ( @QuoteId = 0 ) OR ( @QuoteId IS NULL ) )
	BEGIN
		SELECT '' AS FreightLBS, 0 AS FreightChargeSub, 0 AS FreightChargeSubWork, 0 AS TotalReworkCost, 0 AS TotalReworkCharge, '' AS ExpDeliveryDate, 0.1 AS SubWorkCommission
	END
	ELSE
	BEGIN
		SELECT MAX( FreightLBS ) AS FreightLBS, MAX( ISNULL( FreightChargeSub, 0 ) ) AS FreightChargeSub, 
		MAX( ISNULL( FreightChargeSub, 0 ) ) + MAX( ISNULL( FreightChargeSub1, 0 ) ) AS FreightChargeSubWork,
		SUM( ISNULL( SubWorkCost, 0 ) ) AS TotalReworkCost, 
		( SUM( ISNULL( SubWorkCost, 0 ) ) ) + ISNULL( MAX( FreightChargeSub1 ), 0 )  AS TotalReworkCharge,
		MAX( ExpDeliveryDate ) AS ExpDeliveryDate, MAX( SpindleType ) AS SpindleType, MAX( tblProjects.SpindleId ) AS SpindleId,
		MAX( SubWorkCommission) AS SubWorkCommission
		FROM tblQuotes 
		LEFT JOIN tblQuoteSubWork ON tblQuotes.QuoteId = tblQuoteSubWork.QuoteId
		LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
		LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
		LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
		WHERE tblQuotes.QuoteId = @QuoteId
	
		SELECT QuoteSubWorkId, SubWorkDescription, SupplierName, 
		ISNULL( SubWorkCost, 0 ) AS SubWorkCost
		FROM tblQuoteSubWork
		LEFT JOIN tblSuppliers ON tblQuoteSubWork.SupplierId = tblSuppliers.SupplierId
		WHERE tblQuoteSubWork.QuoteId = @QuoteId
	END
END
ELSE
BEGIN
	SELECT tblQuotes.WorkOrderNumber, ExpDeliveryDate, Customer, SpindleType, FreightLBS, ISNULL( FreightChargeSub, 0 ) AS FreightChargeSub, SubWorkCommission
	FROM tblQuotes
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE tblQuotes.QuoteId = @QuoteId

	SELECT SubWorkDescription, SupplierName
	FROM tblQuoteSubWork
	LEFT JOIN tblSuppliers ON tblQuoteSubWork.SupplierId = tblSuppliers.SupplierId
	WHERE tblQuoteSubWork.QuoteId = @QuoteId
END


GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteSubWorkDetailsByQuoteSubWorkId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteSubWorkDetailsByQuoteSubWorkId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteSubWorkDetailsByQuoteSubWorkId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetQuoteSubWorkDetailsByQuoteSubWorkId]

@ErrorMessage	varchar(255) OUTPUT,
@QuoteSubWorkId	int = NULL

AS


BEGIN
	SELECT QuoteSubWorkId, SubWorkDescription, SupplierName, 
	ISNULL( SubWorkCost, 0 ) AS SubWorkCost, tblQuoteSubWork.SupplierId
	FROM tblQuoteSubWork
	LEFT JOIN tblSuppliers ON tblQuoteSubWork.SupplierId = tblSuppliers.SupplierId
	WHERE tblQuoteSubWork.QuoteSubWorkId = @QuoteSubWorkId
END


GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteSubworkByWorkOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteSubworkByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteSubworkByWorkOrderId] AS' 
END
GO

ALTER PROCEDURE [dbo].[spGetQuoteSubworkByWorkOrderId]
@ErrorMessage	varchar(255) = NULL OUTPUT,
@WorkOrderId	int = 0

AS

SELECT SubWorkDescription
		FROM tblQuoteSubWork
		LEFT JOIN tblQuotes ON tblQuoteSubWork.QuoteId = tblQuotes.QuoteId
		LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
		WHERE tblWorkOrders.WorkOrderId = @WorkOrderId

GO
/****** Object:  StoredProcedure [dbo].[spGetQuoteSubworkList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuoteSubworkList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuoteSubworkList] AS' 
END
GO

ALTER PROCEDURE [dbo].[spGetQuoteSubworkList]
@ErrorMessage	varchar(255) = NULL OUTPUT,
@QuoteId	int = 0

AS

SELECT SubWorkDescription
		FROM tblQuoteSubWork
		LEFT JOIN tblQuotes ON tblQuoteSubWork.QuoteId = tblQuotes.QuoteId
		LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
		WHERE tblQuotes.QuoteId = @QuoteId

GO
/****** Object:  StoredProcedure [dbo].[spGetQuotes]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuotes]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuotes] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetQuotes]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'WorkOrderNumber' OUTPUT,
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(11)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'Customer'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

IF( @OrderBy = 'DateQuoted' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetQuotesByDateQuoted FROM (
	SELECT TOP 100 PERCENT tblQuotes.QuoteId, MAX( tblQuotes.WorkOrderId ) AS WorkOrderId, MAX( tblQuotes.WorkOrderNumber ) AS WorkOrderNumber, MAX( Customer ) AS Customer, MAX( SpindleType ) AS SpindleType, 
	MAX( DateQuoted ) AS DateQuoted, MAX( DateApproved ) AS DateApproved, MAX( tblQuotes.SerialNumber ) AS SerialNumber
	FROM tblQuotes
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	GROUP BY tblQuotes.QuoteId
	ORDER BY DateQuoted
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount

	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetQuotesByDateQuoted
		WHERE DateQuoted BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' ORDER BY RowNumber

		SELECT RowNumber, QuoteId, WorkOrderId, WorkOrderNumber, Customer, SpindleType, DateQuoted, DateApproved, SerialNumber
		FROM #tblTempGetQuotesByDateQuoted
		WHERE ( DateQuoted BETWEEN CONVERT( smalldatetime, @JumpTo ) AND '06/06/2079' )
		AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END
	BEGIN
		SELECT RowNumber, QuoteId, WorkOrderId, WorkOrderNumber, Customer, SpindleType, DateQuoted, DateApproved, SerialNumber
		FROM #tblTempGetQuotesByDateQuoted
		WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetQuotesByDateQuoted
END
ELSE IF( @OrderBy = 'QuoteId' )
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetQuotesByQuoteId FROM (
	SELECT TOP 100 PERCENT tblQuotes.QuoteId, MAX( tblQuotes.WorkOrderId ) AS WorkOrderId, MAX( tblQuotes.WorkOrderNumber ) AS WorkOrderNumber, MAX( Customer ) AS Customer, MAX( SpindleType ) AS SpindleType, 
	MAX( DateQuoted ) AS DateQuoted, MAX( DateApproved ) AS DateApproved, MAX( tblQuotes.SerialNumber ) AS SerialNumber
	FROM tblQuotes
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE QuoteId LIKE CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	GROUP BY tblQuotes.QuoteId
	ORDER BY QuoteId ) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetQuotesByQuoteId
		WHERE QuoteId BETWEEN CONVERT( int, @JumpTo ) AND 9999999 ORDER BY RowNumber
	END
	ELSE
	BEGIN
		SET @JumpTo = 1
	END
	
	SELECT RowNumber, QuoteId, WorkOrderId, WorkOrderNumber, Customer, SpindleType, DateQuoted, DateApproved, SerialNumber
	FROM #tblTempGetQuotesByQuoteId
	WHERE ( QuoteId BETWEEN CONVERT( int, @JumpTo ) AND 9999999 )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

	DROP TABLE #tblTempGetQuotesByQuoteId
END
ELSE
BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetQuotes FROM (
	SELECT TOP 100 PERCENT tblQuotes.QuoteId, MAX( tblQuotes.WorkOrderId ) AS WorkOrderId, MAX( tblQuotes.WorkOrderNumber ) AS WorkOrderNumber, MAX( Customer ) AS Customer, MAX( SpindleType ) AS SpindleType, 
	MAX( DateQuoted ) AS DateQuoted, MAX( DateApproved ) AS DateApproved, MAX( tblQuotes.SerialNumber ) AS SerialNumber
	FROM tblQuotes
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE
	CASE @OrderBy
		WHEN 'Customer' THEN Customer
		WHEN 'SpindleType' THEN SpindleType
		WHEN 'WorkOrderNumber' THEN tblQuotes.WorkOrderNumber
		WHEN 'SerialNumber' THEN tblQuotes.SerialNumber
	END
	LIKE
	CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
	GROUP BY tblQuotes.QuoteId
	ORDER BY
	CASE @OrderBy WHEN 'Customer' THEN MAX( Customer ) ELSE NULL END,
	CASE @OrderBy WHEN 'SpindleType' THEN MAX( SpindleType ) ELSE NULL END,
	CASE @OrderBy WHEN 'WorkOrderNumber' THEN MAX( tblQuotes.WorkOrderNumber ) ELSE NULL END,
	CASE @OrderBy WHEN 'SerialNumber' THEN MAX( tblQuotes.SerialNumber ) ELSE NULL END
	) AS X
	
	SELECT @@ROWCOUNT AS TotalRowCount
	
	IF( @JumpTo <> '%' )
	BEGIN
		SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetQuotes WHERE 
		CASE @OrderBy
			WHEN 'Customer' THEN SUBSTRING( Customer, 1, 1 )
			WHEN 'SpindleType' THEN SUBSTRING( SpindleType, 1, 1 )
			WHEN 'WorkOrderNumber' THEN SUBSTRING( WorkOrderNumber, 1, 1 )
			WHEN 'SerialNumber' THEN SUBSTRING( SerialNumber, 1, 1 )
		END
		LIKE @JumpTo ORDER BY RowNumber
	END
	
	SELECT RowNumber, QuoteId, WorkOrderId, WorkOrderNumber, Customer, SpindleType, DateQuoted, DateApproved, SerialNumber
	FROM #tblTempGetQuotes
	WHERE (
	CASE @OrderBy
		WHEN 'Customer' THEN SUBSTRING( Customer, 1, 1 )
		WHEN 'SpindleType' THEN SUBSTRING( SpindleType, 1, 1 )
		WHEN 'WorkOrderNumber' THEN SUBSTRING( WorkOrderNumber, 1, 1 )
		WHEN 'SerialNumber' THEN SUBSTRING( SerialNumber, 1, 1 )
	END
	BETWEEN @JumpTo AND 'Z' )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
	
	DROP TABLE #tblTempGetQuotes
END

GO
/****** Object:  StoredProcedure [dbo].[spCheckSpindleProduct]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCheckSpindleProduct]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCheckSpindleProduct] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCheckSpindleProduct]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleId	int = NULL,
@ProductId	int = NULL

AS

DECLARE @Spindley	int

	BEGIN
		SELECT @Spindley = tblSpindlesProducts.SpindleProductId
		FROM tblSpindlesProducts
		WHERE tblSpindlesProducts.SpindleId = @SpindleId AND tblSpindlesProducts.ProductId = @ProductId
	END


IF (@Spindley > 0)
	return 1;
ELSE
	return 0;


GO
/****** Object:  StoredProcedure [dbo].[spGetQuotesByWorkOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetQuotesByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetQuotesByWorkOrderId] AS' 
END
GO




ALTER  PROCEDURE [dbo].[spGetQuotesByWorkOrderId]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@WorkOrderId	int		= 0

AS

BEGIN
	SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetQuotesByDateQuoted FROM (
	SELECT TOP 100 PERCENT tblQuotes.QuoteId, MAX( tblQuotes.WorkOrderId ) AS WorkOrderId, MAX( tblQuotes.WorkOrderNumber ) AS WorkOrderNumber, MAX( Customer ) AS Customer, MAX( SpindleType ) AS SpindleType, 
	MAX( DateQuoted ) AS DateQuoted, MAX( DateApproved ) AS DateApproved, MAX( tblQuotes.SerialNumber ) AS SerialNumber
	FROM tblQuotes
	LEFT JOIN tblWorkOrders ON tblQuotes.WorkOrderId = tblWorkOrders.WorkOrderId
	LEFT JOIN tblProjects ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblCustomers ON tblProjects.CustomerId = tblCustomers.CustomerId
	LEFT JOIN tblSpindles ON tblProjects.SpindleId = tblSpindles.SpindleId
	WHERE ( tblQuotes.WorkOrderId = @WorkOrderId )
	GROUP BY tblQuotes.QuoteId
	ORDER BY  DateQuoted
	) AS X

	BEGIN
		SELECT RowNumber, QuoteId, WorkOrderId, WorkOrderNumber, Customer, SpindleType, DateQuoted, DateApproved, SerialNumber
		FROM #tblTempGetQuotesByDateQuoted
		ORDER BY RowNumber
	END

	DROP TABLE #tblTempGetQuotesByDateQuoted
END



GO
/****** Object:  StoredProcedure [dbo].[spCheckSpindleSubWorkDesc]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCheckSpindleSubWorkDesc]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCheckSpindleSubWorkDesc] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCheckSpindleSubWorkDesc]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleId	int = NULL,
@SubWorkDesc	text = NULL

AS

DECLARE @retVal	int

	BEGIN
		SELECT @retVal = tblSpindlesSubWork.SpindleSubWorkId
		FROM tblSpindlesSubWork
		WHERE tblSpindlesSubWork.SpindleId = @SpindleId AND tblSpindlesSubWork.SubWorkDescription LIKE @SubWorkDesc
	END


IF (@retVal > 0)
	return 1;
ELSE
	return 0;


GO
/****** Object:  StoredProcedure [dbo].[spGetShippingMethodByShippingMethodId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetShippingMethodByShippingMethodId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetShippingMethodByShippingMethodId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetShippingMethodByShippingMethodId]

@ErrorMessage		varchar(255)	OUTPUT,
@ShippingMethodId	int

AS

IF( @ShippingMethodId = 0 )
BEGIN
	SELECT '' AS ShippingMethodId, '' AS ShippingMethod
END
ELSE
BEGIN
	SELECT tblShippingMethods.ShippingMethodId, ShippingMethod
	FROM tblShippingMethods 
	WHERE tblShippingMethods.ShippingMethodId = @ShippingMethodId
END



GO
/****** Object:  StoredProcedure [dbo].[spCopySpindlesProductsData]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCopySpindlesProductsData]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCopySpindlesProductsData] AS' 
END
GO

ALTER PROCEDURE [dbo].[spCopySpindlesProductsData] AS

INSERT INTO tblSpindlesProducts
SELECT tblProjects.SpindleId AS SpindleId, tblQuoteParts.ProductId AS ProductId, SupplierId, PartCost AS Cost, Markup, Qty AS Quantity
FROM tblQuoteParts
LEFT JOIN tblQuotes ON tblQuotes.QuoteId = tblQuoteParts.QuoteId
LEFT JOIN tblWorkOrders ON tblWorkOrders.WorkOrderId = tblQuotes.WorkOrderId
LEFT JOIN tblProjects ON tblProjects.ProjectId = tblWorkOrders.WorkOrderId
WHERE (SpindleId is not null) AND (ProductId is not NULL)
ORDER BY SpindleId, ProductId

GO
/****** Object:  StoredProcedure [dbo].[spGetShippingMethodList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetShippingMethodList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetShippingMethodList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetShippingMethodList]

AS

SELECT ShippingMethodId, ShippingMethod FROM tblShippingMethods ORDER BY ShippingMethod



GO
/****** Object:  StoredProcedure [dbo].[spCreateCustomersTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateCustomersTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateCustomersTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateCustomersTable]

AS

IF EXISTS ( SELECT * FROM dbo.sysobjects WHERE id = object_id( N'[dbo].[tblCustomers]' ) AND OBJECTPROPERTY( id, N'IsUserTable' ) = 1 )
DROP TABLE [dbo].[tblCustomers]

CREATE TABLE [dbo].[tblCustomers] (
	[CustomerId] [int] IDENTITY (1000, 1) NOT NULL ,
	[Customer] [varchar] (50) NULL ,
	[Address] [varchar] (255) NULL ,
	[City] [varchar] (30) NULL ,
	[State] [varchar] (20) NULL ,
	[Country] [varchar] (30) NULL ,
	[Zip] [varchar] (10) NULL ,
	[Contact] [varchar] (20) NULL ,
	[ContactTitle] [varchar] (30) NULL ,
	[Department] [varchar] (30) NULL ,
	[TelephoneNumber] [varchar] (20) NULL ,
	[Extension] [varchar] (5) NULL ,
	[FaxNumber] [varchar] (20) NULL ,
	[EmailAddress] [varchar] (64) NULL ,
	[DateEstablished] [varchar] (25) NULL ,
	[Notes] [text] NULL 
)

INSERT tblCustomers( Customer, Address, City, State, Zip, Contact, TelephoneNumber, FaxNumber, DateEstablished )
SELECT tblRepairQuoteSpindle.Customer, MAX( Address ) AS Address, MAX( City ) AS City, MAX( State ) AS State, MAX( Zip ) AS Zip, MAX( Contact ) AS Contact, 
MAX( TelephoneNumber ) AS TelephoneNumber, MAX( FaxNumber ) AS FaxNumber, MAX( DateEstablished ) AS DateEstablished
FROM tblSpindle
RIGHT OUTER JOIN ( SELECT Customer FROM tblSpindle GROUP BY Customer
UNION SELECT Customer FROM tblRepairQuote WHERE NOT tblRepairQuote.Customer IS NULL ) AS tblRepairQuoteSpindle 
ON tblRepairQuoteSpindle.Customer = tblSpindle.Customer
GROUP BY tblRepairQuoteSpindle.Customer
ORDER BY tblRepairQuoteSpindle.Customer



GO
/****** Object:  StoredProcedure [dbo].[spGetShippingMethods]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetShippingMethods]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetShippingMethods] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetShippingMethods]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'ShippingMethod',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'ShippingMethod'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetShippingMethods FROM (
SELECT TOP 100 PERCENT tblShippingMethods.ShippingMethodId, MAX( ShippingMethod ) AS ShippingMethod
FROM tblShippingMethods
WHERE
CASE @OrderBy
	WHEN 'ShippingMethod' THEN ShippingMethod
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblShippingMethods.ShippingMethodId
ORDER BY
CASE @OrderBy WHEN 'ShippingMethod' THEN MAX( ShippingMethod ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetShippingMethods WHERE 
	CASE @OrderBy
		WHEN 'ShippingMethod' THEN SUBSTRING( ShippingMethod, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, ShippingMethodId, ShippingMethod
FROM #tblTempGetShippingMethods
WHERE (
CASE @OrderBy
	WHEN 'ShippingMethod' THEN SUBSTRING( ShippingMethod, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetShippingMethods



GO
/****** Object:  StoredProcedure [dbo].[spCreateProjectsTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateProjectsTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateProjectsTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateProjectsTable]

AS

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblProjects]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblProjects]

CREATE TABLE [dbo].[tblProjects] (
	[ProjectId] [int] IDENTITY (1000, 1) NOT NULL ,
	[ProjectName] [varchar] (50) NULL ,
	[ProjectTypeId] [int] NULL ,
	[SalesRepId] [int] NULL ,
	[CustomerId] [int] NULL ,
	[ProjectDescription] [varchar] (255) NULL ,
	[CustomerPO] [varchar] (15) NULL ,
	[StartDate] [smalldatetime] NULL ,
	[EstCompletionDate] [smalldatetime] NULL ,
	[CompletionDate] [smalldatetime] NULL ,
	[ProjectPriorityId] [int] NULL ,
	[SpindleId] [int] NULL 
) ON [PRIMARY]

INSERT tblProjects( ProjectName, ProjectTypeId, SalesRepId, CustomerId, CustomerPO, StartDate, CompletionDate, ProjectPriorityId, SpindleId ) 
SELECT CONVERT( varchar(50), CONVERT( int, WorkOrderNumber ) ) AS ProjectName, 1000 AS ProjectTypeId, 1 AS SalesRepId, CustomerId, 
PONumber AS CustomerPO, DateIn AS StartDate, DateOut AS CompletionDate, 1001 AS ProjectPriorityId, SpindleId
FROM tblSpindle
LEFT JOIN tblCustomers ON tblSpindle.Customer = tblCustomers.Customer
LEFT JOIN tblSpindles ON tblSpindle.SpindleType = tblSpindles.SpindleType
ORDER BY ProjectName



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleBySpindleId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleBySpindleId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleBySpindleId] AS' 
END
GO





ALTER   PROCEDURE [dbo].[spGetSpindleBySpindleId]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleId	int

AS

IF( ( @SpindleId = 0 ) OR ( @SpindleId IS NULL ) )
BEGIN
SELECT '' AS SpindleId,
'' AS SpindleType,
'' AS SpindleCategoryId,
'' AS DrawingNumber,
'' AS RPM,
'' AS Weight,
'' AS Vibration,
'' AS VibrationRear,
'' AS CoolantFlow,
'' AS GSE,
'' AS GSERear,
'' AS ShaftTemp,
'' AS CoolantTempIncomingSet,
'' AS CoolantTempActual,
'' AS FrontTemp,
'' AS RearTemp,
'' AS CoolantPressureIncoming,
'' AS BreakInTime,
'' AS CoolingMethod,
'' AS Volts,
'' AS HP,
'' AS Amps,
'' AS Phase,
'' AS Hz,
'' AS Thermistor,
'' AS Poles,
'' AS AmpDraw,
'' AS ConnectorCtrl,
'' AS ConnectorPower,
'' AS Converter,
'' AS ToolHolder,
'' AS PullPin,
'' AS EMDimension,
'' AS EjectionPath,
'' AS ToolOutPressure,
'' AS ReturnPressure,
'' AS DrawbarForce,
'' AS ToolChangeFunction,
'' AS ProximitySwitchFunction,
'' AS Lubrication,
'' AS Grease,
'' AS OilMist,
'' AS OilJet,
'' AS OilGreaseType,
'' AS IntervalDPM,
'' AS MainPressure,
'' AS TubePressure,
'' AS LubeNotes,
'' AS Preload,
'' AS RadialPlay,
'' AS AxialPlay,
'' AS RunoutFront,
'' AS RunoutFrontLocation,
'' AS RunoutFront2,
'' AS RunoutFront2Location,
'' AS RunoutRear,
'' AS RunoutRearLocation,
'' AS RunoutRear2,
'' AS RunoutRear2Location,
'' AS ToolContact,
'' AS ToolContactRear,
'' AS ToolGap,
'' AS ToolGapRear,
'' AS Other,
'' AS BalancingRequirements,
'' AS BearingInformation,
'' AS GeneralSpindleNotes

END
ELSE
BEGIN
	SELECT tblSpindles.SpindleId, 
SpindleType,
tblSpindles.SpindleCategoryId AS SpindleCategoryId,
DrawingNumber,
RPM,
Weight,
Vibration,
VibrationRear,
CoolantFlow,
GSE,
GSERear,
ShaftTemp,
CoolantTempIncomingSet,
CoolantTempActual,
FrontTemp,
RearTemp,
CoolantPressureIncoming,
BreakInTime,
CoolingMethod,
Volts,
HP,
Amps,
Phase,
Hz,
Thermistor,
Poles,
AmpDraw,
ConnectorCtrl,
ConnectorPower,
Converter,
ToolHolder,
PullPin,
EMDimension,
EjectionPath,
ToolOutPressure,
ReturnPressure,
DrawbarForce,
ToolChangeFunction,
ProximitySwitchFunction,
Lubrication,
Grease,
OilMist,
OilJet,
OilGreaseType,
IntervalDPM,
MainPressure,
TubePressure,
LubeNotes,
Preload,
RadialPlay,
AxialPlay,
RunoutFront,
RunoutFrontLocation,
RunoutFront2,
RunoutFront2Location,
RunoutRear,
RunoutRearLocation,
RunoutRear2,
RunoutRear2Location,
ToolContact,
ToolContactRear,
ToolGap,
ToolGapRear,
Other,
BalancingRequirements,
BearingInformation,
GeneralSpindleNotes,
SpindleCategoryName
	FROM tblSpindles
	LEFT JOIN tblSpindleCategories ON tblSpindleCategories.SpindleCategoryId = tblSpindles.SpindleCategoryId
	WHERE tblSpindles.SpindleId = @SpindleId
END



GO
/****** Object:  StoredProcedure [dbo].[spCreateQuoteBearingsTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateQuoteBearingsTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateQuoteBearingsTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateQuoteBearingsTable]

AS

IF EXISTS ( SELECT * FROM dbo.sysobjects WHERE id = object_id( N'[dbo].[tblQuoteBearings]' ) AND OBJECTPROPERTY( id, N'IsUserTable' ) = 1 )
DROP TABLE [dbo].[tblQuoteBearings]

CREATE TABLE [dbo].[tblQuoteBearings] (
	[QuoteBearingId] int IDENTITY (1000, 1) NOT NULL,
	[QuoteId] int,
	[ProductId] int NULL,
	[CenBearingCode] varchar(30) NULL,
	[CssPrice] money NULL,
	[BearingCost] money NULL,
	[BearingMarkup] numeric(10,2) NULL,
	[SupplierId] int,
	[BearingDescription] varchar(255) NULL
)

INSERT tblQuoteBearings( QuoteId, ProductId, CenBearingCode, CssPrice, BearingCost, BearingMarkup, SupplierId, BearingDescription )
SELECT QuoteId, CAST( ProductId1 AS int ) AS ProductId, CenBearingCode1 AS CenBearingCode, 
CssPrice1 AS CssPrice, BearingCost1 AS BearingCost, BearingMarkup1 AS BearingMarkup, SupplierId, Bearing1 AS BearingDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.BearingSource1 = tblSuppliers.SupplierName
WHERE NOT Bearing1 IS NULL
UNION ALL SELECT QuoteId, CAST( ProductId2 AS int ) AS ProductId, CenBearingCode2 AS CenBearingCode, 
CssPrice2 AS CssPrice, BearingCost2 AS BearingCost, BearingMarkup2 AS BearingMarkup, SupplierId, Bearing2 AS BearingDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.BearingSource2 = tblSuppliers.SupplierName
WHERE NOT Bearing2 IS NULL
UNION ALL SELECT QuoteId, CAST( ProductId3 AS int ) AS ProductId, CenBearingCode3 AS CenBearingCode, 
CssPrice3 AS CssPrice, BearingCost3 AS BearingCost, BearingMarkup3 AS BearingMarkup, SupplierId, Bearing3 AS BearingDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.BearingSource3 = tblSuppliers.SupplierName
WHERE NOT Bearing3 IS NULL
UNION ALL SELECT QuoteId, CAST( ProductId4 AS int ) AS ProductId, CenBearingCode4 AS CenBearingCode, 
CssPrice4 AS CssPrice, BearingCost4 AS BearingCost, BearingMarkup4 AS BearingMarkup, SupplierId, Bearing4 AS BearingDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.BearingSource4 = tblSuppliers.SupplierName
WHERE NOT Bearing4 IS NULL
UNION ALL SELECT QuoteId, CAST( ProductId5 AS int ) AS ProductId, CenBearingCode5 AS CenBearingCode, 
CssPrice5 AS CssPrice, BearingCost5 AS BearingCost, BearingMarkup5 AS BearingMarkup, SupplierId, Bearing5 AS BearingDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.BearingSource5 = tblSuppliers.SupplierName
WHERE NOT Bearing5 IS NULL
UNION ALL SELECT QuoteId, CAST( ProductId6 AS int ) AS ProductId, CenBearingCode6 AS CenBearingCode, 
CssPrice6 AS CssPrice, BearingCost6 AS BearingCost, BearingMarkup6 AS BearingMarkup, SupplierId, Bearing6 AS BearingDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.BearingSource6 = tblSuppliers.SupplierName
WHERE NOT Bearing6 IS NULL
ORDER BY QuoteId



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleCategories]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleCategories]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleCategories] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleCategories]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'SpindleCategory',
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'SpindleCategory'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetSpindleCategories FROM (
SELECT TOP 100 PERCENT tblSpindleCategories.SpindleCategoryId AS SpindleCategoryId, MAX( SpindleCategoryName ) AS SpindleCategory
FROM tblSpindleCategories
WHERE
CASE @OrderBy
	WHEN 'SpindleCategory' THEN SpindleCategoryName
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblSpindleCategories.SpindleCategoryId
ORDER BY
CASE @OrderBy WHEN 'SpindleCategory' THEN MAX( SpindleCategoryName ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetSpindleCategories WHERE 
	CASE @OrderBy
		WHEN 'SpindleCategory' THEN SUBSTRING( SpindleCategory, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber
END

SELECT RowNumber, SpindleCategoryId, SpindleCategory
FROM #tblTempGetSpindleCategories
WHERE (
CASE @OrderBy
	WHEN 'SpindleCategory' THEN SUBSTRING( SpindleCategory, 1, 1 )
END
BETWEEN @JumpTo AND 'Z' )
AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber

DROP TABLE #tblTempGetSpindleCategories


GO
/****** Object:  StoredProcedure [dbo].[spCreateQuotePartsTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateQuotePartsTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateQuotePartsTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateQuotePartsTable]

AS

IF EXISTS ( SELECT * FROM dbo.sysobjects WHERE id = object_id( N'[dbo].[tblQuoteParts]' ) AND OBJECTPROPERTY( id, N'IsUserTable' ) = 1 )
DROP TABLE [dbo].[tblQuoteParts]

CREATE TABLE [dbo].[tblQuoteParts] (
	[QuotePartId] int IDENTITY (1000, 1) NOT NULL,
	[QuoteId] int,
	[ProductId] int NULL,
	[PartCost] money NULL,
	[Markup] numeric(10,2) NULL,
	[PartDescription] varchar(255) NULL,
	[SupplierId] int
)

INSERT tblQuoteParts( QuoteId, ProductId, PartCost, Markup, PartDescription, SupplierId )
SELECT QuoteId, CAST( ProductId7 AS int ) AS ProductId, CostPart1 AS PartCost, Markup1 AS Markup, Part1 AS PartDescription, SupplierId FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourcePart1 = tblSuppliers.SupplierName
WHERE NOT Part1 IS NULL
UNION SELECT QuoteId, CAST( ProductId8 AS int ) AS ProductId, CostPart2 AS PartCost, Markup2 AS Markup, Part2 AS PartDescription, SupplierId FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourcePart2 = tblSuppliers.SupplierName
WHERE NOT Part2 IS NULL
UNION SELECT QuoteId, CAST( ProductId9 AS int ) AS ProductId, CostPart3 AS PartCost, Markup3 AS Markup, Part3 AS PartDescription, SupplierId FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourcePart3 = tblSuppliers.SupplierName
WHERE NOT Part3 IS NULL
UNION SELECT QuoteId, CAST( ProductId10 AS int ) AS ProductId, CostPart4 AS PartCost, Markup4 AS Markup, Part4 AS PartDescription, SupplierId FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourcePart4 = tblSuppliers.SupplierName
WHERE NOT Part4 IS NULL
UNION SELECT QuoteId, CAST( ProductId11 AS int ) AS ProductId, CostPart5 AS PartCost, Markup5 AS Markup, Part5 AS PartDescription, SupplierId FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourcePart5 = tblSuppliers.SupplierName
WHERE NOT Part5 IS NULL
UNION SELECT QuoteId, CAST( ProductId12 AS int ) AS ProductId, CostPart6 AS PartCost, Markup6 AS Markup, Part6 AS PartDescription, SupplierId FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourcePart6 = tblSuppliers.SupplierName
WHERE NOT Part6 IS NULL
ORDER BY QuoteId



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleCategoryBySpindleCategoryId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleCategoryBySpindleCategoryId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleCategoryBySpindleCategoryId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleCategoryBySpindleCategoryId]

@ErrorMessage		varchar(255)	OUTPUT,
@SpindleCategoryId	int

AS

IF( @SpindleCategoryId = 0 )
BEGIN
	SELECT '' AS SpindleCategoryId, '' AS SpindleCategory
END
ELSE
BEGIN
	SELECT tblSpindleCategories.SpindleCategoryId AS SpindleCategoryId, SpindleCategoryName AS SpindleCategory
	FROM tblSpindleCategories
	WHERE tblSpindleCategories.SpindleCategoryId = @SpindleCategoryId
END


GO
/****** Object:  StoredProcedure [dbo].[spCreateQuoteSubWorkTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateQuoteSubWorkTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateQuoteSubWorkTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateQuoteSubWorkTable]

AS

IF EXISTS ( SELECT * FROM dbo.sysobjects WHERE id = object_id( N'[dbo].[tblQuoteSubWork]' ) AND OBJECTPROPERTY( id, N'IsUserTable' ) = 1 )
DROP TABLE [dbo].[tblQuoteSubWork]

CREATE TABLE [dbo].[tblQuoteSubWork] (
	[QuoteSubWorkId] int IDENTITY (1000, 1) NOT NULL,
	[QuoteId] int,
	[SubWorkCost] money NULL,
	[SupplierId] int,
	[SubWorkDescription] varchar(255) NULL
)

INSERT tblQuoteSubWork( QuoteId, SubWorkCost, SupplierId, SubWorkDescription )
SELECT QuoteId, CostSubWork1 AS CostSubWork, SupplierId, SubWork1 AS SubWorkDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourceWubWork1 = tblSuppliers.SupplierName
WHERE NOT SubWork1 IS NULL
UNION SELECT QuoteId, CostSubWork2 AS CostSubWork, SupplierId, SubWork2 AS SubWorkDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourceSubWork2 = tblSuppliers.SupplierName
WHERE NOT SubWork2 IS NULL
UNION SELECT QuoteId, CostSubWork3 AS CostSubWork, SupplierId, SubWork3 AS SubWorkDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourceSubWork3 = tblSuppliers.SupplierName
WHERE NOT SubWork3 IS NULL
UNION SELECT QuoteId, CostSubWork4 AS CostSubWork, SupplierId, SubWork4 AS SubWorkDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourceSubWork4 = tblSuppliers.SupplierName
WHERE NOT SubWork4 IS NULL
UNION SELECT QuoteId, CostSubWork5 AS CostSubWork, SupplierId, SubWork5 AS SubWorkDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourceSubWork5 = tblSuppliers.SupplierName
WHERE NOT SubWork5 IS NULL
UNION SELECT QuoteId, CostSubWork6 AS CostSubWork, SupplierId, SubWork6 AS SubWorkDescription
FROM tblRepairQuote
LEFT JOIN tblQuotes ON tblRepairQuote.OriginalQuoteId = tblQuotes.OriginalQuoteId
LEFT JOIN tblSuppliers ON tblRepairQuote.SourceSubWork6 = tblSuppliers.SupplierName
WHERE NOT SubWork6 IS NULL
ORDER BY QuoteId



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleCategoryList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleCategoryList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleCategoryList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleCategoryList]

AS

SELECT SpindleCategoryId, SpindleCategoryName FROM tblSpindleCategories ORDER BY SpindleCategoryName


GO
/****** Object:  StoredProcedure [dbo].[spCreateQuotesTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateQuotesTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateQuotesTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateQuotesTable]

AS

IF EXISTS ( SELECT * FROM dbo.sysobjects WHERE id = object_id( N'[dbo].[tblQuotes]' ) AND OBJECTPROPERTY( id, N'IsUserTable' ) = 1 )
DROP TABLE [dbo].[tblQuotes]

SELECT IDENTITY( int, 1000, 1 ) AS QuoteId, OriginalQuoteId, WorkOrderId, tblRepairQuote.WorkOrderNumber, tblRepairQuote.SerialNumber, 
DisassemblyEvaluation, HoursDisassembly, CleanAndInspect, HoursCleanAndInspect, InhouseGrinding, GrindingHours, Balancing, BalancingHours, ElectricalWork, 
ElectricalHours, GreaseBearings, GreaseHours, AssemblyAndTest, AssemblyAndTestHours, MscWorkNeeded, MscWorkHours, FreightChargeParts, FreightChargeSubWork, 
FreightLBS, FreightChargeSub, FreightChargeSub1, ExpDeliveryDate, tblRepairQuote.Notes, PartsCommission, BearingFreightCharge, BearingCommission, DateApproved, DateQuoted
INTO tblQuotes
FROM tblRepairQuote
LEFT JOIN tblWorkOrders ON tblRepairQuote.WorkOrderNumber = CONVERT( varchar(10), tblWorkOrders.WorkOrderNumber )
ORDER BY tblRepairQuote.OriginalQuoteId



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleDetailsBySpindleId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleDetailsBySpindleId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleDetailsBySpindleId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleDetailsBySpindleId]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleId	int = NULL

AS


BEGIN
	
	BEGIN
		--Return the recordset for the previous used parts
		SELECT SpindleProductId, tblSpindlesProducts.ProductId, ProductDescription,  ISNULL( Cost, 0 ) AS PartCost, Markup, Quantity AS Qty, SupplierName
		FROM tblSpindlesProducts
		LEFT JOIN tblSuppliers ON tblSpindlesProducts.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN tblProducts ON tblSpindlesProducts.ProductId = tblProducts.ProductId
		WHERE tblSpindlesProducts.SpindleId = @SpindleId AND tblProducts.CategoryId <> 3 AND tblProducts.CategoryId <> 9

		--Return the recordset for the previous used bearings
		SELECT SpindleProductId, tblSpindlesProducts.ProductId, ProductDescription AS BearingDescription,  ISNULL( Cost, 0 ) AS BearingCost, Markup, Quantity AS Qty, SupplierName
		FROM tblSpindlesProducts
		LEFT JOIN tblSuppliers ON tblSpindlesProducts.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN tblProducts ON tblSpindlesProducts.ProductId = tblProducts.ProductId
		WHERE tblSpindlesProducts.SpindleId = @SpindleId AND tblProducts.CategoryId = 3

		--Return the recordset for the previous used subwork
		SELECT SpindleSubWorkId, SubWorkDescription,  ISNULL( SubWorkCost, 0 ) AS SubWorkCost, SupplierName
		FROM tblSpindlesSubWork
		LEFT JOIN tblSuppliers ON tblSpindlesSubWork.SupplierId = tblSuppliers.SupplierId
		WHERE tblSpindlesSubWork.SpindleId = @SpindleId
		
	END
END


GO
/****** Object:  StoredProcedure [dbo].[spCreateRepairBearingSourcesTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateRepairBearingSourcesTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateRepairBearingSourcesTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateRepairBearingSourcesTable]

AS

INSERT INTO tblSuppliers( SupplierName ) SELECT BearingSource FROM ( SELECT BearingSource1 AS BearingSource FROM tblRepairQuote WHERE NOT BearingSource1 IS NULL
UNION SELECT BearingSource2 AS BearingSource FROM tblRepairQuote WHERE NOT BearingSource2 IS NULL
UNION SELECT BearingSource3 AS BearingSource FROM tblRepairQuote WHERE NOT BearingSource3 IS NULL
UNION SELECT BearingSource4 AS BearingSource FROM tblRepairQuote WHERE NOT BearingSource4 IS NULL
UNION SELECT BearingSource5 AS BearingSource FROM tblRepairQuote WHERE NOT BearingSource5 IS NULL
UNION SELECT BearingSource6 AS BearingSource FROM tblRepairQuote WHERE NOT BearingSource6 IS NULL ) AS tblRepairBearingSources
WHERE BearingSource NOT IN ( SELECT SupplierName FROM tblSuppliers ) ORDER BY BearingSource



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleInformationByWorkOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleInformationByWorkOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleInformationByWorkOrderId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleInformationByWorkOrderId]

@ErrorMessage	varchar(255)	OUTPUT,
@WorkOrderId	int

AS

IF( ( @WorkOrderId = 0 ) OR ( @WorkOrderId IS NULL ) )
BEGIN
	SELECT '' AS SpindleId, '' AS SpindleType, '' AS Weight, '' AS RPM, '' AS Volts, '' AS Hz, '' AS HP, '' AS Poles, '' AS SpindleCategoryName, '' AS Amps,
	'' AS Phase, '' AS ToolHolder, '' AS PullPin, '' AS EMDimension, '' AS EjectionPath, '' AS ToolOutPressure,
	'' AS ReturnPressure, '' AS Lubrication, '' AS Grease, '' AS OilMist, '' AS OilJet, '' AS OilGreaseType, '' AS MainPressure,
	'' AS TubePressure, '' AS IntervalDPM, '' AS LubeNotes, '' AS CoolingMethod, '' AS GeneralSpindleNotes, '' AS SerialNumber
END
ELSE
BEGIN
	SELECT tblSpindles.SpindleId, SpindleType, Weight, RPM, Volts, Hz, HP, Poles, SpindleCategoryName, Amps, Phase, ToolHolder,
	PullPin, EMDimension, EjectionPath, ToolOutPressure, ReturnPressure, Lubrication, Grease, OilMist, OilJet, OilGreaseType,
	MainPressure, TubePressure, IntervalDPM, LubeNotes, CoolingMethod, GeneralSpindleNotes, SerialNumber,
	tblWorkOrders.GSE, GSEFinal,
	BalVelocity, BalVelocityFinal,
	BreakIn, BreakInFinal,
	ActualDrawforce, ActualDrawforceFinal,
	Cooling, CoolingFinal,
	RoomTemp, RoomTempFinal,
	tblWorkOrders.FrontTemp, FrontTempFinal,
	tblWorkOrders.RearTemp, RearTempFinal,
	tblWorkOrders.RunoutFront, RunoutFrontFinal,
	Rear, RearFinal
	FROM tblSpindles
	LEFT JOIN tblProjects ON tblSpindles.SpindleId = tblProjects.SpindleId
	LEFT JOIN tblWorkOrders ON tblWorkOrders.ProjectId = tblProjects.ProjectId
	LEFT JOIN tblSpindleCategories ON tblSpindleCategories.SpindleCategoryId = tblSpindles.SpindleCategoryId
	WHERE tblWorkOrders.WorkOrderId = @WorkOrderId
END

GO
/****** Object:  StoredProcedure [dbo].[spCreateRepairPartSourcesTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateRepairPartSourcesTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateRepairPartSourcesTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateRepairPartSourcesTable]

AS

INSERT INTO tblSuppliers( SupplierName ) SELECT SourcePart FROM ( SELECT SourcePart1 AS SourcePart FROM tblRepairQuote WHERE NOT SourcePart1 IS NULL
UNION SELECT SourcePart2 AS SourcePart FROM tblRepairQuote WHERE NOT SourcePart2 IS NULL
UNION SELECT SourcePart3 AS SourcePart FROM tblRepairQuote WHERE NOT SourcePart3 IS NULL
UNION SELECT SourcePart4 AS SourcePart FROM tblRepairQuote WHERE NOT SourcePart4 IS NULL
UNION SELECT SourcePart5 AS SourcePart FROM tblRepairQuote WHERE NOT SourcePart5 IS NULL
UNION SELECT SourcePart6 AS SourcePart FROM tblRepairQuote WHERE NOT SourcePart6 IS NULL ) AS tblRepairPartSources
WHERE SourcePart NOT IN ( SELECT SupplierName FROM tblSuppliers ) ORDER BY SourcePart



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleProductDetailsBySpindleProductId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleProductDetailsBySpindleProductId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleProductDetailsBySpindleProductId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleProductDetailsBySpindleProductId]

@ErrorMessage		varchar(255)	OUTPUT,
@SpindleProductId	int = NULL

AS


BEGIN
	
	BEGIN
		--Return the recordset for the specified part/bearing
		SELECT SpindleProductId, tblSpindlesProducts.ProductId, ProductDescription,  ISNULL( Cost, 0 ) AS PartCost, Markup, Quantity AS Qty, SupplierName, tblSpindlesProducts.SupplierId
		FROM tblSpindlesProducts
		LEFT JOIN tblSuppliers ON tblSpindlesProducts.SupplierId = tblSuppliers.SupplierId
		LEFT JOIN tblProducts ON tblSpindlesProducts.ProductId = tblProducts.ProductId
		WHERE tblSpindlesProducts.SpindleProductId = @SpindleProductId
		
	END
END


GO
/****** Object:  StoredProcedure [dbo].[spCreateRepairSubWorkSourcesTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateRepairSubWorkSourcesTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateRepairSubWorkSourcesTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateRepairSubWorkSourcesTable]

AS

INSERT INTO tblSuppliers( SupplierName ) SELECT SourceSubWork FROM ( SELECT SourceWubWork1 AS SourceSubWork FROM tblRepairQuote WHERE NOT SourceWubWork1 IS NULL
UNION SELECT SourceSubWork2 AS SourceSubWork FROM tblRepairQuote WHERE NOT SourceSubWork2 IS NULL
UNION SELECT SourceSubWork3 AS SourceSubWork FROM tblRepairQuote WHERE NOT SourceSubWork3 IS NULL
UNION SELECT SourceSubWork4 AS SourceSubWork FROM tblRepairQuote WHERE NOT SourceSubWork4 IS NULL
UNION SELECT SourceSubWork5 AS SourceSubWork FROM tblRepairQuote WHERE NOT SourceSubWork5 IS NULL
UNION SELECT SourceSubWork6 AS SourceSubWork FROM tblRepairQuote WHERE NOT SourceSubWork6 IS NULL ) AS tblRepairSubWorkSources
WHERE SourceSubWork NOT IN ( SELECT SupplierName FROM tblSuppliers ) ORDER BY SourceSubWork



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleProductsByPurchaseOrderId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleProductsByPurchaseOrderId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleProductsByPurchaseOrderId] AS' 
END
GO

ALTER PROCEDURE [dbo].[spGetSpindleProductsByPurchaseOrderId]

@ErrorMessage	varchar(255)	OUTPUT,
@PurchaseOrderId	int = NULL

AS



		SELECT tblSpindlesProducts.ProductID, tblProducts.PartNumber, ProductName, tblSpindlesProducts.SupplierId AS SupplierId, SupplierName, tblSpindlesProducts.Cost AS Cost, Markup, Quantity, SpindleProductId, tblProducts.ProductShortDescription, CategoryId,
		ISNULL (vwOnHandProducts.OnHand,0) AS OnHand, ISNULL (vwQuotedProducts.QtyQuoted, 0) AS OnOrder
		FROM tblSpindlesProducts
		LEFT JOIN tblProducts ON tblProducts.ProductId = tblSpindlesProducts.ProductId
		LEFT JOIN tblSuppliers ON tblSuppliers.SupplierId = tblSpindlesProducts.SupplierId
		LEFT JOIN vwOnHandProducts ON tblProducts.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN vwQuotedProducts ON tblProducts.ProductId = vwQuotedProducts.ProductId
		LEFT JOIN tblProjects ON tblSpindlesProducts.SpindleId = tblProjects.SpindleId
		LEFT JOIN tblWorkOrders ON tblProjects.ProjectId = tblWorkOrders.ProjectId
		LEFT JOIN tblWorkOrderPOs ON tblWorkOrders.WorkOrderId = tblWorkOrderPOs.WorkOrderId
		
		WHERE tblWorkOrderPOs.PurchaseOrderId = @PurchaseOrderId

GO
/****** Object:  StoredProcedure [dbo].[spCreateSpindlesTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateSpindlesTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateSpindlesTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateSpindlesTable]

AS

IF EXISTS ( SELECT * FROM dbo.sysobjects WHERE id = object_id( N'[dbo].[tblSpindles]' ) AND OBJECTPROPERTY( id, N'IsUserTable' ) = 1 )
DROP TABLE [dbo].[tblSpindles]

CREATE TABLE [dbo].[tblSpindles] (
	[SpindleId] [int] IDENTITY (1000, 1) NOT NULL ,
	[SpindleType] [varchar] (50) NULL ,
	[Weight] [varchar] (10) NULL ,
	[RPM] [numeric](10, 2) NULL ,
	[Volts] [numeric](10, 2) NULL ,
	[Hz] [numeric](10, 2) NULL ,
	[HP] [numeric](10, 2) NULL ,
	[Poles] [numeric](10, 2) NULL ,
	[Category] [varchar] (10) NULL ,
	[Amps] [varchar] (10) NULL ,
	[Phase] [varchar] (1) NULL ,
	[DrawForce] [varchar] (12) NULL
)

INSERT tblSpindles( SpindleType, Weight, RPM, Volts, Hz, HP, Poles, Category, Amps, Phase, DrawForce )
SELECT tblRepairQuoteSpindle.SpindleType, MAX( Weight ) AS Weight, MAX( RPM ) AS RPM, MAX( Volts ) AS Volts, MAX( Hz ) AS Hz, 
MAX( HP ) AS HP, MAX( Poles ) AS Poles, MAX( Category ) AS Category, MAX( Amps ) AS Amps, MAX( Phase ) AS Phase, MAX( DrawForce ) AS DrawForce
FROM tblSpindle
RIGHT OUTER JOIN ( SELECT SpindleType FROM tblSpindle WHERE NOT tblSpindle.SpindleType IS NULL GROUP BY SpindleType
UNION SELECT SpindleType FROM tblRepairQuote WHERE NOT tblRepairQuote.SpindleType IS NULL ) AS tblRepairQuoteSpindle
ON tblRepairQuoteSpindle.SpindleType = tblSpindle.SpindleType
GROUP BY tblRepairQuoteSpindle.SpindleType
ORDER BY tblRepairQuoteSpindle.SpindleType



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleProductsBySpindleId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleProductsBySpindleId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleProductsBySpindleId] AS' 
END
GO






ALTER    PROCEDURE [dbo].[spGetSpindleProductsBySpindleId]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleId	int = NULL,
@PartType	int = NULL

AS


IF (@PartType = -1)
	--This selects all parts except bearings and rework
	BEGIN
		SELECT tblSpindlesProducts.ProductID, tblProducts.PartNumber, ProductName, tblSpindlesProducts.SupplierId AS SupplierId, SupplierName, Cost, Markup, Quantity, SpindleProductId, tblProducts.ProductShortDescription, CategoryId,
		ISNULL (vwOnHandProducts.OnHand,0) AS OnHand, ISNULL (vwQuotedProducts.QtyQuoted, 0) AS OnOrder
		FROM tblSpindlesProducts
		LEFT JOIN tblProducts ON tblProducts.ProductId = tblSpindlesProducts.ProductId
		LEFT JOIN tblSuppliers ON tblSuppliers.SupplierId = tblSpindlesProducts.SupplierId
		LEFT JOIN vwOnHandProducts ON tblProducts.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN vwQuotedProducts ON tblProducts.ProductId = vwQuotedProducts.ProductId
		WHERE tblSpindlesProducts.SpindleId = @SpindleId AND tblProducts.CategoryID <> 3 AND tblProducts.CategoryID <> 9
	END

ELSE IF (@PartType > 0)
	--This selects the specific part type
	BEGIN
		SELECT tblSpindlesProducts.ProductID, tblProducts.PartNumber, ProductName, tblSpindlesProducts.SupplierId AS SupplierId, SupplierName, Cost, Markup, Quantity, SpindleProductId, tblProducts.ProductShortDescription, CategoryId,
		ISNULL (vwOnHandProducts.OnHand,0) AS OnHand, ISNULL (vwQuotedProducts.QtyQuoted, 0) AS OnOrder
		FROM tblSpindlesProducts
		LEFT JOIN tblProducts ON tblProducts.ProductId = tblSpindlesProducts.ProductId
		LEFT JOIN tblSuppliers ON tblSuppliers.SupplierId = tblSpindlesProducts.SupplierId
		LEFT JOIN vwOnHandProducts ON tblProducts.ProductId = vwOnHandProducts.ProductId
		LEFT JOIN vwQuotedProducts ON tblProducts.ProductId = vwQuotedProducts.ProductId
		WHERE tblSpindlesProducts.SpindleId = @SpindleId AND tblProducts.CategoryId = @PartType
	END

GO
/****** Object:  StoredProcedure [dbo].[spCreateWorkOrdersTable]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spCreateWorkOrdersTable]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spCreateWorkOrdersTable] AS' 
END
GO



ALTER PROCEDURE [dbo].[spCreateWorkOrdersTable]

AS

/*IF EXISTS ( SELECT * FROM dbo.sysobjects WHERE id = object_id( N'[dbo].[tblWorkOrders]' ) AND OBJECTPROPERTY( id, N'IsUserTable' ) = 1 )
DROP TABLE [dbo].[tblWorkOrders]

SELECT IDENTITY( int, 1000, 1 ) AS WorkOrderId, CAST( WorkOrderNumber AS int ) AS WorkOrderNumber, CustomerId, PromiseDate, DateIn, DateOut,
SpindleId, SerialNumber, NewSpindle, PONumber, Labor, Material, Subcontract, Cost, Charge, Date, DateExp, SalesRep, Location, Priority, BoeingWorkOrderNumber, Parts, Bearings, Lube, BalVelocity, GSE, BreakIn, RoomTemp, FrontTemp, RearTemp,
Cooling, RunoutFront, Rear, Other, IncomingInspection, Comments, Remarks, DateRec, ExpectedDelDate, Commission, CommissionReceivedDate, 
AdditionalInfo
INTO tblWorkOrders
FROM tblSpindle
LEFT JOIN tblCustomers ON tblSpindle.Customer = tblCustomers.Customer
LEFT JOIN tblSpindles ON tblSpindle.SpindleType = tblSpindles.SpindleType
ORDER BY WorkOrderNumber
GO*/

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblWorkOrders]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblWorkOrders]

CREATE TABLE [dbo].[tblWorkOrders] (
	[WorkOrderId] [int] IDENTITY (1000, 1) NOT NULL ,
	[WorkOrderNumber] [int] NOT NULL ,
	[ProjectId] [int] NULL ,
	[PromiseDate] [smalldatetime] NULL ,
	[DateIn] [smalldatetime] NULL ,
	[DateOut] [smalldatetime] NULL ,
	[SerialNumber] [varchar] (20) NULL ,
	[NewSpindle] [varchar] (10) NULL ,
	[PONumber] [varchar] (15) NULL ,
	[Labor] [varchar] (10) NULL ,
	[Material] [varchar] (15) NULL ,
	[Subcontract] [varchar] (10) NULL ,
	[Cost] [numeric](10, 2) NULL ,
	[Charge] [numeric](10, 2) NULL ,
	[Date] [smalldatetime] NULL ,
	[DateExp] [smalldatetime] NULL ,
	[SalesRep] [varchar] (4) NULL ,
	[Location] [varchar] (50) NULL ,
	[Priority] [varchar] (10) NULL ,
	[BoeingWorkOrderNumber] [varchar] (15) NULL ,
	[Parts] [varchar] (150) NULL ,
	[Bearings] [varchar] (150) NULL ,
	[Lube] [varchar] (150) NULL ,
	[BalVelocity] [varchar] (20) NULL ,
	[GSE] [varchar] (10) NULL ,
	[BreakIn] [varchar] (10) NULL ,
	[RoomTemp] [varchar] (3) NULL ,
	[FrontTemp] [varchar] (4) NULL ,
	[RearTemp] [varchar] (4) NULL ,
	[Cooling] [varchar] (50) NULL ,
	[RunoutFront] [varchar] (50) NULL ,
	[Rear] [varchar] (50) NULL ,
	[Other] [varchar] (200) NULL ,
	[IncomingInspection] [varchar] (200) NULL ,
	[Comments] [varchar] (254) NULL ,
	[Remarks] [varchar] (254) NULL ,
	[DateRec] [smalldatetime] NULL ,
	[ExpectedDelDate] [smalldatetime] NULL ,
	[Commission] [varchar] (10) NULL ,
	[CommissionReceivedDate] [smalldatetime] NULL ,
	[AdditionalInfo] [varchar] (100) NULL 
) ON [PRIMARY]

INSERT tblWorkOrders( WorkOrderNumber, ProjectId, PromiseDate, DateIn, DateOut, SerialNumber, NewSpindle, PONumber, Labor, Material, Subcontract, Cost, Charge, [Date], DateExp, 
SalesRep, Location, Priority, BoeingWorkOrderNumber, Parts, Bearings, Lube, BalVelocity, GSE, BreakIn, RoomTemp, FrontTemp, RearTemp, Cooling, RunoutFront, Rear, Other, 
IncomingInspection, Comments, Remarks, DateRec, ExpectedDelDate, Commission, CommissionReceivedDate, AdditionalInfo )
SELECT CAST( WorkOrderNumber AS int ) AS WorkOrderNumber, ProjectId, PromiseDate, DateIn, DateOut,
SerialNumber, NewSpindle, PONumber, Labor, Material, Subcontract, Cost, Charge, Date, DateExp, SalesRep, Location, Priority, BoeingWorkOrderNumber, Parts, Bearings, Lube, BalVelocity, GSE, BreakIn, RoomTemp, FrontTemp, RearTemp,
Cooling, RunoutFront, Rear, Other, IncomingInspection, Comments, Remarks, DateRec, ExpectedDelDate, Commission, CommissionReceivedDate, 
AdditionalInfo
FROM tblSpindle
LEFT JOIN tblProjects ON tblSpindle.WorkOrderNumber = tblProjects.ProjectName
ORDER BY WorkOrderNumber



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleSubWorkBySpindleId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleSubWorkBySpindleId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleSubWorkBySpindleId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleSubWorkBySpindleId]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleId	int = NULL

AS
	BEGIN
		SELECT SubWorkDescription, tblSpindlesSubWork.SupplierId AS SupplierId, SupplierName, SubWorkCost, SpindleSubWorkId
		FROM tblSpindlesSubWork
		LEFT JOIN tblSuppliers ON tblSuppliers.SupplierId = tblSpindlesSubWork.SupplierId
		WHERE tblSpindlesSubWork.SpindleId = @SpindleId
	END


GO
/****** Object:  StoredProcedure [dbo].[spDeleteCallByCallId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteCallByCallId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteCallByCallId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteCallByCallId]

@ErrorMessage	varchar(255)	OUTPUT,
@CallId		int

AS
/*
DECLARE @ErrorNum AS int

DELETE FROM tblCalls WHERE CallId = @CallId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END
*/

BEGIN TRY
	DELETE from tblCalls WHERE CallId = @CallId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH


RETURN 0



GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleSubWorkDetailsBySpindleSubWorkId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleSubWorkDetailsBySpindleSubWorkId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleSubWorkDetailsBySpindleSubWorkId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleSubWorkDetailsBySpindleSubWorkId]

@ErrorMessage	varchar(255)	OUTPUT,
@SpindleSubWorkId	int = NULL

AS
	BEGIN
		SELECT SubWorkDescription, tblSpindlesSubWork.SupplierId AS SupplierId, SupplierName, SubWorkCost, SpindleSubWorkId
		FROM tblSpindlesSubWork
		LEFT JOIN tblSuppliers ON tblSuppliers.SupplierId = tblSpindlesSubWork.SupplierId
		WHERE tblSpindlesSubWork.SpindleSubWorkId = @SpindleSubWorkId
	END


GO
/****** Object:  StoredProcedure [dbo].[spDeleteCustomerByCustomerId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteCustomerByCustomerId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteCustomerByCustomerId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteCustomerByCustomerId]

@ErrorMessage	varchar(255)	OUTPUT,
@CustomerId	int

AS

/*
DECLARE @ErrorNum AS int

DELETE FROM tblCustomers WHERE CustomerId = @CustomerId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0

*/

BEGIN TRY
	DELETE FROM tblCustomers WHERE CustomerId = @CustomerId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH


RETURN 0

GO
/****** Object:  StoredProcedure [dbo].[spGetSpindleTypeList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindleTypeList]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindleTypeList] AS' 
END
GO



ALTER PROCEDURE [dbo].[spGetSpindleTypeList]

AS

SELECT SpindleId, SpindleType FROM tblSpindles ORDER BY SpindleType



GO
/****** Object:  StoredProcedure [dbo].[spDeleteCustomerContactByCustomerContactId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteCustomerContactByCustomerContactId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteCustomerContactByCustomerContactId] AS' 
END
GO

ALTER PROCEDURE [dbo].[spDeleteCustomerContactByCustomerContactId]

@ErrorMessage	varchar(255)	OUTPUT,
@CustomerContactId	int

AS

/*
DECLARE @ErrorNum AS int

DELETE FROM tblCustomerContacts WHERE CustomerContactId = @CustomerContactId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/

BEGIN TRY
	DELETE FROM tblCustomerContacts WHERE CustomerContactId = @CustomerContactId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH


RETURN 0
GO
/****** Object:  StoredProcedure [dbo].[spGetSpindles]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spGetSpindles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spGetSpindles] AS' 
END
GO




ALTER     PROCEDURE [dbo].[spGetSpindles]

@ErrorMessage	varchar(255)	= NULL OUTPUT,
@RowStart	int 		= 1 OUTPUT,
@RecordCount	int 		= 15 OUTPUT,
@OrderBy	varchar(64)	= 'SpindleType' OUTPUT,
@SearchString	varchar(255)	= '%',
@JumpTo		varchar(9)	= '%'

AS

IF( @RowStart = 0 ) OR ( @RowStart IS NULL ) SET @RowStart = 1
IF( @RecordCount = 0 ) OR ( @RecordCount IS NULL ) SET @RecordCount = 15
IF( DATALENGTH( @OrderBy ) = 0 ) OR ( @OrderBy IS NULL ) SET @OrderBy = 'SpindleType'
IF( DATALENGTH( @SearchString ) = 0 ) OR ( @SearchString IS NULL ) SET @SearchString = '%'
IF( DATALENGTH( @JumpTo ) = 0 ) OR ( @JumpTo IS NULL ) SET @JumpTo = '%'

SET @JumpTo = SUBSTRING( @JumpTo, 1, 1 )

SELECT IDENTITY( int, 1, 1 ) AS RowNumber, * INTO #tblTempGetSpindles FROM (
SELECT TOP 100 PERCENT tblSpindles.SpindleId, MAX( SpindleType ) AS SpindleType, MAX( Weight ) AS Weight, MAX( RPM ) AS RPM, MAX( Volts ) AS Volts, 
MAX( Hz ) AS Hz, MAX( HP ) AS HP, MAX( tblSpindleCategories.SpindleCategoryName ) AS SpindleCategoryName, MAX( Amps ) AS Amps
FROM tblSpindles
	LEFT JOIN tblSpindleCategories ON tblSpindles.SpindleCategoryId = tblSpindleCategories.SpindleCategoryId
WHERE
CASE @OrderBy
	WHEN 'SpindleType' THEN SpindleType
	WHEN 'SpindleCategoryName' THEN SpindleCategoryName
END
LIKE
CASE @SearchString WHEN '%' THEN @SearchString ELSE '%' + @SearchString + '%' END
GROUP BY tblSpindles.SpindleId
ORDER BY
CASE @OrderBy WHEN 'SpindleType' THEN MAX( SpindleType ) ELSE NULL END,
CASE @OrderBy WHEN 'SpindleCategoryName' THEN MAX( SpindleCategoryName ) ELSE NULL END
) AS X

SELECT @@ROWCOUNT AS TotalRowCount

IF( @JumpTo <> '%' )
BEGIN
	SELECT TOP 1 @RowStart = ( ( @RowStart - 1 ) + RowNumber ) FROM #tblTempGetSpindles WHERE 
	CASE @OrderBy
		WHEN 'SpindleType' THEN SUBSTRING( SpindleType, 1, 1 )
		WHEN 'SpindleCategoryName' THEN SUBSTRING( SpindleCategoryName, 1, 1 )
	END
	LIKE @JumpTo ORDER BY RowNumber

	SELECT RowNumber, SpindleId, SpindleType, Weight, RPM, Volts, Hz, HP, SpindleCategoryName, Amps
	FROM #tblTempGetSpindles
	WHERE (
	CASE @OrderBy
		WHEN 'SpindleType' THEN SUBSTRING( SpindleType, 1, 1 )
		WHEN 'SpindleCategoryName' THEN SUBSTRING( SpindleCategoryName, 1, 1 )
	END
	BETWEEN @JumpTo AND 'Z' )
	AND ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
END
ELSE
BEGIN
	SELECT RowNumber, SpindleId, SpindleType, Weight, RPM, Volts, Hz, HP, SpindleCategoryName, Amps
	FROM #tblTempGetSpindles
	WHERE ( ( RowNumber >= @RowStart ) AND ( RowNumber < ( @RowStart + @RecordCount ) ) ) ORDER BY RowNumber
END

DROP TABLE #tblTempGetSpindles




GO
/****** Object:  StoredProcedure [dbo].[spDeleteEmployeeByEmployeeId]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spDeleteEmployeeByEmployeeId]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[spDeleteEmployeeByEmployeeId] AS' 
END
GO



ALTER PROCEDURE [dbo].[spDeleteEmployeeByEmployeeId]

@ErrorMessage	varchar(255)	OUTPUT,
@EmployeeId	int

AS

/*
DECLARE @ErrorNum AS int

DELETE FROM tblEmployees WHERE EmployeeId = @EmployeeId

SET @ErrorNum = @@ERROR

IF( @ErrorNum <> 0 )
BEGIN
	SELECT @ErrorMessage = master.dbo.sysmessages.description FROM master.dbo.sysmessages WHERE master.dbo.sysmessages.error = @ErrorNum
	RETURN -1
END

RETURN 0
*/
BEGIN TRY
	DELETE FROM tblEmployees WHERE EmployeeId = @EmployeeId
END TRY
BEGIN CATCH
	SELECT @ErrorMessage = ERROR_MESSAGE()
	RETURN -1
END CATCH

RETURN 0

GO
/****** Object:  View [dbo].[vwOpenOrders]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwOpenOrders]'))
EXEC dbo.sp_executesql @statement = N'
CREATE VIEW dbo.vwOpenOrders
AS
SELECT     dbo.tblQuoteParts.ProductId AS ProductId, COUNT(dbo.tblQuoteParts.ProductId) AS QtyNeeded, 
                      SUM(ISNULL(dbo.tblInventoryTransactions.UnitsReceived, 0)) - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsShrinkage, 0)) 
                      - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsSold, 0)) AS OnHand, SUM(ISNULL(dbo.tblInventoryTransactions.UnitsOrdered, 0)) 
                      - SUM(ISNULL(dbo.tblInventoryTransactions.UnitsReceived, 0)) AS OnOrder, MAX(ISNULL(dbo.tblProducts.ReorderLevel, 0)) AS ReorderLevel, 
                      MAX(dbo.tblProducts.ProductName) AS ProductName
FROM         dbo.tblQuoteParts LEFT OUTER JOIN
                      dbo.tblInventoryTransactions ON dbo.tblQuoteParts.ProductId = dbo.tblInventoryTransactions.ProductID LEFT OUTER JOIN
                      dbo.tblProducts ON dbo.tblQuoteParts.ProductId = dbo.tblProducts.ProductID LEFT OUTER JOIN
                      dbo.tblQuotes ON dbo.tblQuoteParts.QuoteId = dbo.tblQuotes.QuoteId LEFT OUTER JOIN
                      dbo.tblWorkOrders ON dbo.tblQuotes.WorkOrderId = dbo.tblWorkOrders.WorkOrderId
WHERE     (NOT (dbo.tblQuoteParts.ProductId IS NULL)) AND (NOT (dbo.tblProducts.ProductName IS NULL)) AND (dbo.tblWorkOrders.DateOut IS NULL)
GROUP BY dbo.tblQuoteParts.ProductId

' 
GO
/****** Object:  View [dbo].[vwOpenOrderReport]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwOpenOrderReport]'))
EXEC dbo.sp_executesql @statement = N'
CREATE VIEW dbo.vwOpenOrderReport

AS



SELECT
vwOpenOrders.ProductId, vwOpenOrders.QtyNeeded, vwOpenOrders.OnHand,
vwOpenOrders.OnOrder, vwOpenOrders.ProductName, vwOpenOrders.ReorderLevel, ProductDescription
FROM
vwOpenOrders
LEFT JOIN tblProducts ON vwOpenOrders.ProductId = tblProducts.ProductId


' 
GO
/****** Object:  View [dbo].[vwContactChristmasList]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwContactChristmasList]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwContactChristmasList
AS
SELECT     dbo.tblCustomers.CustomerId, dbo.tblCustomers.Customer, dbo.tblCustomers.Address, dbo.tblCustomers.City, dbo.tblCustomers.State, 
                      dbo.tblCustomers.Country, dbo.tblCustomers.Zip, dbo.tblCustomers.DateEstablished, dbo.tblCustomers.Notes, dbo.tblCustomerContacts.Contact, 
                      dbo.tblCustomerContacts.ContactTitle, dbo.tblCustomerContacts.Department, dbo.tblCustomerContacts.TelephoneNumber, 
                      dbo.tblCustomerContacts.Extension, dbo.tblCustomerContacts.EmailAddress
FROM         dbo.tblCustomers LEFT OUTER JOIN
                      dbo.tblCustomerContacts ON dbo.tblCustomers.MainContactId = dbo.tblCustomerContacts.CustomerContactId
' 
GO
/****** Object:  View [dbo].[vwJKM_Spindles_Customers_Work_Orders]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_Spindles_Customers_Work_Orders]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW [dbo].[vwJKM_Spindles_Customers_Work_Orders]
AS
SELECT     dbo.tblSpindles.SpindleType, dbo.tblCustomers.Customer, dbo.tblWorkOrders.WorkOrderNumber, dbo.tblWorkOrders.SerialNumber, 
                      dbo.tblWorkOrders.WorkOrderId, dbo.tblSpindles.RPM, dbo.tblCustomers.State, dbo.tblProjects.SalesRepId, dbo.tblSpindles.SpindleId, 
                      dbo.tblWorkOrders.DateIn, dbo.tblWorkOrders.DateOut, dbo.tblSpindles.SpindleCategoryId, dbo.tblSpindles.DrawingNumber, 
                      dbo.tblCustomers.CustomerId, dbo.tblProjects.ProjectId
FROM         dbo.tblCustomers INNER JOIN
                      dbo.tblProjects ON dbo.tblCustomers.CustomerId = dbo.tblProjects.CustomerId INNER JOIN
                      dbo.tblSpindles ON dbo.tblProjects.SpindleId = dbo.tblSpindles.SpindleId INNER JOIN
                      dbo.tblWorkOrders ON dbo.tblProjects.ProjectId = dbo.tblWorkOrders.ProjectId
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_Spindles_Customers_Work_Orders', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[33] 4[21] 2[25] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 0
               Left = 0
               Bottom = 221
               Right = 144
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 0
               Left = 404
               Bottom = 277
               Right = 576
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblSpindles"
            Begin Extent = 
               Top = 0
               Left = 644
               Bottom = 455
               Right = 847
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 115
               Left = 157
               Bottom = 457
               Right = 360
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1545
         Alias = 900
         Table = 1695
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_Spindles_Customers_Work_Orders'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_Spindles_Customers_Work_Orders', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_Spindles_Customers_Work_Orders'
GO
/****** Object:  View [dbo].[vwJKM_Customer_Salesman_Email]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_Customer_Salesman_Email]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW [dbo].[vwJKM_Customer_Salesman_Email]
AS
SELECT     dbo.tblCustomers.Customer, dbo.tblCustomerContacts.EmailAddress, dbo.tblCustomerContacts.Contact, dbo.tblEmployees.LastName, 
                      dbo.tblEmployees.FirstName, dbo.tblEmployees.EmployeeID, dbo.tblCustomers.SalesRepId
FROM         dbo.tblCustomers INNER JOIN
                      dbo.tblCustomerContacts ON dbo.tblCustomers.CustomerId = dbo.tblCustomerContacts.CustomerId INNER JOIN
                      dbo.tblEmployees ON dbo.tblCustomers.SalesRepId = dbo.tblEmployees.EmployeeID
WHERE     (dbo.tblCustomerContacts.EmailAddress IS NOT NULL)
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_Customer_Salesman_Email', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[24] 4[37] 2[12] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 291
               Right = 196
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomerContacts"
            Begin Extent = 
               Top = 6
               Left = 234
               Bottom = 309
               Right = 409
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblEmployees"
            Begin Extent = 
               Top = 6
               Left = 447
               Bottom = 310
               Right = 599
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 2685
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_Customer_Salesman_Email'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_Customer_Salesman_Email', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_Customer_Salesman_Email'
GO
/****** Object:  View [dbo].[vwJKM_failureReportHistory]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_failureReportHistory]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_failureReportHistory
AS
SELECT     dbo.tblCustomers.CustomerId, dbo.tblWorkOrders.ProjectId, dbo.tblWorkOrders.IncomingInspection, dbo.tblWorkOrders.PONumber, dbo.tblWorkOrders.SerialNumber, 
                      dbo.tblWorkOrders.DateIn, dbo.tblWorkOrders.DateOut, dbo.tblWorkOrders.WorkOrderNumber
FROM         dbo.tblCustomers CROSS JOIN
                      dbo.tblWorkOrders
WHERE     (dbo.tblCustomers.CustomerId = 1539) AND (dbo.tblWorkOrders.ProjectId = 4901)
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_failureReportHistory', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[31] 4[23] 2[9] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1[50] 2[25] 3) )"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1 [56] 4 [18] 2))"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 241
               Right = 195
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 6
               Left = 233
               Bottom = 241
               Right = 541
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1440
         Width = 1440
         Width = 27660
         Width = 1440
         Width = 1440
         Width = 1440
         Width = 1440
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_failureReportHistory'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_failureReportHistory', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_failureReportHistory'
GO
/****** Object:  View [dbo].[vwJKM_call_log_searching]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_call_log_searching]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_call_log_searching
AS
SELECT     CallComments, CustomerId, CallDate, CallId, EmployeeId
FROM         dbo.tblCalls
WHERE     (EmployeeId = 57)
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_call_log_searching', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[29] 4[15] 2[10] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblCalls"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 310
               Right = 255
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 9945
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_call_log_searching'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_call_log_searching', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_call_log_searching'
GO
/****** Object:  View [dbo].[vwJKM_WORKING_qte_by_spindleWILDCARD]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_WORKING_qte_by_spindleWILDCARD]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_WORKING_qte_by_spindleWILDCARD
AS
SELECT     dbo.tblSpindles.SpindleType, dbo.tblQuotes.DisassemblyEvaluation, dbo.tblQuotes.HoursDisassembly, dbo.tblQuotes.CleanAndInspect, 
                      dbo.tblQuotes.HoursCleanAndInspect, dbo.tblQuotes.InhouseGrinding, dbo.tblQuotes.GrindingHours, dbo.tblQuotes.Balancing, 
                      dbo.tblQuotes.BalancingHours, dbo.tblQuotes.ElectricalWork, dbo.tblQuotes.ElectricalHours, dbo.tblQuotes.GreaseBearings, 
                      dbo.tblQuotes.GreaseHours, dbo.tblQuotes.AssemblyAndTest, dbo.tblQuotes.AssemblyAndTestHours, dbo.tblQuotes.MscWorkNeeded, 
                      dbo.tblQuotes.MscWorkHours, dbo.tblQuotes.FreightChargeParts, dbo.tblQuotes.FreightChargeSubWork, dbo.tblQuotes.FreightChargeSub, 
                      dbo.tblQuotes.FreightChargeSub1, dbo.tblQuotes.PartsCommission, dbo.tblQuotes.BearingFreightCharge, dbo.tblQuotes.BearingCommission, 
                      dbo.tblQuotes.SpacerPreparation, dbo.tblQuotes.SpacerPreparationHours, dbo.tblQuotes.LaborCommission, dbo.tblQuotes.SubWorkCommission, 
                      dbo.tblQuotes.HandlingCharge
FROM         dbo.tblQuotes INNER JOIN
                      dbo.tblWorkOrders ON dbo.tblQuotes.WorkOrderId = dbo.tblWorkOrders.WorkOrderId INNER JOIN
                      dbo.tblProjects ON dbo.tblWorkOrders.ProjectId = dbo.tblProjects.ProjectId INNER JOIN
                      dbo.tblSpindles ON dbo.tblProjects.SpindleId = dbo.tblSpindles.SpindleId
WHERE     (dbo.tblSpindles.SpindleType LIKE ''%toyoda%630%'')
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_WORKING_qte_by_spindleWILDCARD', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[6] 4[46] 2[19] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblQuotes"
            Begin Extent = 
               Top = 0
               Left = 0
               Bottom = 453
               Right = 174
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 0
               Left = 423
               Bottom = 389
               Right = 595
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 0
               Left = 195
               Bottom = 389
               Right = 398
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblSpindles"
            Begin Extent = 
               Top = 6
               Left = 633
               Bottom = 389
               Right = 836
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 30
         Width = 284
         Width = 3705
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Widt' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_WORKING_qte_by_spindleWILDCARD'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_WORKING_qte_by_spindleWILDCARD', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'h = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_WORKING_qte_by_spindleWILDCARD'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_WORKING_qte_by_spindleWILDCARD', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_WORKING_qte_by_spindleWILDCARD'
GO
/****** Object:  View [dbo].[vwJKM_project_inquiries]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_project_inquiries]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_project_inquiries
AS
SELECT     dbo.tblProjects.ProjectId, dbo.tblWorkOrders.WorkOrderNumber, dbo.tblWorkOrders.DateOut, dbo.tblWorkOrders.SerialNumber, 
                      dbo.tblWorkOrders.PONumber, dbo.tblWorkOrders.IncomingInspection, dbo.tblWorkOrders.Comments, dbo.tblWorkOrders.Remarks
FROM         dbo.tblProjects INNER JOIN
                      dbo.tblWorkOrders ON dbo.tblProjects.ProjectId = dbo.tblWorkOrders.ProjectId
WHERE     (dbo.tblProjects.ProjectId = 4872)
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_project_inquiries', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[41] 4[31] 2[13] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 310
               Right = 210
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 3
               Left = 434
               Bottom = 307
               Right = 637
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 11
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_project_inquiries'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_project_inquiries', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_project_inquiries'
GO
/****** Object:  View [dbo].[vwJKM_spindles-serialNumbers]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_spindles-serialNumbers]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.[vwJKM_spindles-serialNumbers]
AS
SELECT     dbo.tblProjects.CustomerId, dbo.tblProjects.SpindleId, dbo.tblWorkOrders.SerialNumber, dbo.tblProjects.ProjectName, dbo.tblProjects.ProjectTypeId, 
                      dbo.tblSpindles.SpindleType
FROM         dbo.tblProjects INNER JOIN
                      dbo.tblWorkOrders ON dbo.tblProjects.ProjectId = dbo.tblWorkOrders.ProjectId INNER JOIN
                      dbo.tblSpindles ON dbo.tblProjects.SpindleId = dbo.tblSpindles.SpindleId
WHERE     (dbo.tblProjects.CustomerId = 1539)
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_spindles-serialNumbers', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[28] 4[25] 2[12] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 0
               Left = 38
               Bottom = 209
               Right = 210
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 0
               Left = 248
               Bottom = 209
               Right = 451
            End
            DisplayFlags = 280
            TopColumn = 48
         End
         Begin Table = "tblSpindles"
            Begin Extent = 
               Top = 6
               Left = 489
               Bottom = 121
               Right = 692
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 160
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 3180
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 15' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_spindles-serialNumbers'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_spindles-serialNumbers', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'00
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_spindles-serialNumbers'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_spindles-serialNumbers', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_spindles-serialNumbers'
GO
/****** Object:  View [dbo].[vwJKM_new_customers]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_new_customers]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_new_customers
AS
SELECT     TOP 100 PERCENT dbo.tblCustomers.CustomerId, dbo.tblCustomers.Customer, dbo.tblCustomers.Address, dbo.tblCustomers.City, dbo.tblCustomers.State, 
                      dbo.tblCustomers.Zip, dbo.tblCustomers.DateEstablished, dbo.tblCustomers.SalesRepId, dbo.tblCustomers.MainContactId, dbo.tblCustomers.Customer AS Expr1, 
                      dbo.tblCustomers.CustomerId AS Expr2, dbo.tblCustomers.City AS Expr3, dbo.tblCustomers.State AS Expr4, dbo.tblCustomers.DateEstablished AS Expr5, 
                      dbo.tblCustomers.SalesRepId AS Expr7, dbo.tblCustomerContacts.Contact, dbo.tblCustomerContacts.ContactTitle, dbo.tblCustomerContacts.Department, 
                      dbo.tblCustomerContacts.MobileNumber, dbo.tblCustomerContacts.EmailAddress
FROM         dbo.tblCustomers INNER JOIN
                      dbo.tblCustomerContacts ON dbo.tblCustomers.CustomerId = dbo.tblCustomerContacts.CustomerId
WHERE     (dbo.tblCustomerContacts.EmailAddress IS NOT NULL)
ORDER BY dbo.tblCustomers.DateEstablished
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_new_customers', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[43] 4[14] 2[25] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 356
               Right = 449
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomerContacts"
            Begin Extent = 
               Top = 6
               Left = 693
               Bottom = 317
               Right = 876
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 29
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_new_customers'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_new_customers', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_new_customers'
GO
/****** Object:  View [dbo].[vwJKM_open_quotes_dollars_salesman]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_open_quotes_dollars_salesman]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_open_quotes_dollars_salesman
AS
SELECT     TOP 100 PERCENT dbo.tblWorkOrders.SalesRep, dbo.tblWorkOrders.DateIn, ISNULL(dbo.tblQuotes.WorkOrderNumber, ''&nbsp;'') AS WorkOrderNumber, 
                      CONVERT(varchar(10), dbo.tblQuotes.DateQuoted, 101) AS DateQuoted, CONVERT(varchar(10), dbo.tblQuotes.DateApproved, 101) AS DateApproved, 
                      ISNULL(dbo.tblCustomers.Customer, ''&nbsp;'') AS Customer, ISNULL(dbo.tblSpindles.SpindleType, ''&nbsp;'') AS SpindleType, dbo.tblQuotes.DeliveryInformation, 
                      dbo.tblQuotes.ExpeditedDeliveryInformation
FROM         dbo.tblQuotes LEFT OUTER JOIN
                      dbo.tblWorkOrders ON dbo.tblQuotes.WorkOrderId = dbo.tblWorkOrders.WorkOrderId LEFT OUTER JOIN
                      dbo.tblProjects ON dbo.tblWorkOrders.ProjectId = dbo.tblProjects.ProjectId LEFT OUTER JOIN
                      dbo.tblCustomers ON dbo.tblProjects.CustomerId = dbo.tblCustomers.CustomerId LEFT OUTER JOIN
                      dbo.tblSpindles ON dbo.tblProjects.SpindleId = dbo.tblSpindles.SpindleId
WHERE     (dbo.tblQuotes.DateApproved IS NULL)
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_open_quotes_dollars_salesman', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[28] 4[30] 2[42] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblQuotes"
            Begin Extent = 
               Top = 0
               Left = 74
               Bottom = 782
               Right = 298
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 0
               Left = 575
               Bottom = 783
               Right = 778
            End
            DisplayFlags = 280
            TopColumn = 10
         End
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 4
               Left = 812
               Bottom = 286
               Right = 984
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 286
               Left = 813
               Bottom = 768
               Right = 971
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblSpindles"
            Begin Extent = 
               Top = 5
               Left = 1002
               Bottom = 593
               Right = 1205
            End
            DisplayFlags = 280
            TopColumn = 34
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 11
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin Colu' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_open_quotes_dollars_salesman'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_open_quotes_dollars_salesman', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'mnWidths = 11
         Column = 4545
         Alias = 1905
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_open_quotes_dollars_salesman'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_open_quotes_dollars_salesman', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_open_quotes_dollars_salesman'
GO
/****** Object:  View [dbo].[vwJKM_customerscontacts]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_customerscontacts]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_customerscontacts
AS
SELECT     dbo.tblCustomers.Customer, dbo.tblCustomers.MainContactId, dbo.tblCustomerContacts.Contact, dbo.tblCustomerContacts.ContactTitle, 
                      dbo.tblCustomerContacts.Department, dbo.tblCustomerContacts.TelephoneNumber, dbo.tblCustomerContacts.Extension, dbo.tblCustomerContacts.MobileNumber, 
                      dbo.tblCustomerContacts.EmailAddress, dbo.tblCustomers.CustomerId
FROM         dbo.tblCustomers INNER JOIN
                      dbo.tblCustomerContacts ON dbo.tblCustomers.CustomerId = dbo.tblCustomerContacts.CustomerId
WHERE     (dbo.tblCustomers.CustomerId = 3533)
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customerscontacts', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[28] 4[43] 2[12] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 357
               Right = 204
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomerContacts"
            Begin Extent = 
               Top = 6
               Left = 242
               Bottom = 355
               Right = 425
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 19
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customerscontacts'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customerscontacts', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customerscontacts'
GO
/****** Object:  View [dbo].[vwJKM_warranty_by_spindle_ID]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_warranty_by_spindle_ID]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_warranty_by_spindle_ID
AS
SELECT     dbo.tblWorkOrders.WorkOrderNumber, dbo.tblSpindles.SpindleId, dbo.tblWorkOrders.PONumber, dbo.tblWorkOrders.IncomingInspection, 
                      dbo.tblWorkOrders.Comments, dbo.tblWorkOrders.Remarks
FROM         dbo.tblCustomers INNER JOIN
                      dbo.tblProjects ON dbo.tblCustomers.CustomerId = dbo.tblProjects.CustomerId INNER JOIN
                      dbo.tblSpindles ON dbo.tblProjects.SpindleId = dbo.tblSpindles.SpindleId INNER JOIN
                      dbo.tblWorkOrders ON dbo.tblProjects.ProjectId = dbo.tblWorkOrders.ProjectId
WHERE     (dbo.tblSpindles.SpindleId = 3034) AND (dbo.tblWorkOrders.PONumber LIKE ''%warranty%'') OR
                      (dbo.tblWorkOrders.IncomingInspection LIKE ''%warranty%'') OR
                      (dbo.tblWorkOrders.Comments LIKE ''%warranty%'') OR
                      (dbo.tblWorkOrders.Remarks LIKE ''%warranty%'')
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_warranty_by_spindle_ID', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[41] 4[19] 2[16] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 5
               Left = 0
               Bottom = 356
               Right = 331
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblSpindles"
            Begin Extent = 
               Top = 2
               Left = 347
               Bottom = 353
               Right = 675
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 0
               Left = 688
               Bottom = 351
               Right = 1074
            End
            DisplayFlags = 280
            TopColumn = 39
         End
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 6
               Left = 1112
               Bottom = 354
               Right = 1365
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_warranty_by_spindle_ID'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_warranty_by_spindle_ID', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_warranty_by_spindle_ID'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_warranty_by_spindle_ID', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_warranty_by_spindle_ID'
GO
/****** Object:  View [dbo].[vwJKM_RPT_open_quotes]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_RPT_open_quotes]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_RPT_open_quotes
AS
SELECT     TOP 100 PERCENT dbo.tblQuotes.QuoteId AS QID, ISNULL(dbo.tblQuotes.WorkOrderNumber, ''&nbsp;'') AS WO, CONVERT(varchar(10), dbo.tblQuotes.DateQuoted, 101) 
                      AS Quoted, CONVERT(varchar(10), dbo.tblQuotes.DateApproved, 101) AS DateApproved, ISNULL(dbo.tblCustomers.Customer, ''&nbsp;'') AS Customer, 
                      ISNULL(dbo.tblSpindles.SpindleType, ''&nbsp;'') AS Spindle, dbo.tblCustomers.SalesRepId AS SalesRep
FROM         dbo.tblProjects RIGHT OUTER JOIN
                      dbo.tblQuotes LEFT OUTER JOIN
                      dbo.tblWorkOrders ON dbo.tblQuotes.WorkOrderId = dbo.tblWorkOrders.WorkOrderId ON dbo.tblProjects.ProjectId = dbo.tblWorkOrders.ProjectId LEFT OUTER JOIN
                      dbo.tblCustomers ON dbo.tblProjects.CustomerId = dbo.tblCustomers.CustomerId LEFT OUTER JOIN
                      dbo.tblSpindles ON dbo.tblProjects.SpindleId = dbo.tblSpindles.SpindleId
WHERE     (dbo.tblQuotes.DateApproved IS NULL)
ORDER BY dbo.tblCustomers.SalesRepId
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_RPT_open_quotes', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[36] 4[30] 2[9] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblQuotes"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 283
               Right = 270
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 6
               Left = 308
               Bottom = 283
               Right = 519
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 0
               Left = 567
               Bottom = 276
               Right = 747
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 153
               Left = 787
               Bottom = 427
               Right = 953
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblSpindles"
            Begin Extent = 
               Top = 138
               Left = 989
               Bottom = 257
               Right = 1200
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 12
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_RPT_open_quotes'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_RPT_open_quotes', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N' = 
      Begin ColumnWidths = 11
         Column = 4545
         Alias = 1545
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_RPT_open_quotes'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_RPT_open_quotes', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_RPT_open_quotes'
GO
/****** Object:  View [dbo].[vwJKM_customer_by_90day_no_call_log]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_customer_by_90day_no_call_log]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_customer_by_90day_no_call_log
AS
SELECT DISTINCT 
                      t1.CustomerId, dbo.tblCustomers.Customer, dbo.tblCustomers.SalesRepId, dbo.tblCustomers.City, dbo.tblCustomers.State, dbo.tblCustomers.MainContactId
FROM         dbo.tblCalls AS t1 INNER JOIN
                      dbo.tblCustomers ON t1.CustomerId = dbo.tblCustomers.CustomerId INNER JOIN
                      dbo.tblCustomerContacts ON dbo.tblCustomers.CustomerId = dbo.tblCustomerContacts.CustomerId INNER JOIN
                      dbo.tblEmployees ON t1.EmployeeId = dbo.tblEmployees.EmployeeID
WHERE     (NOT EXISTS
                          (SELECT     1 AS Expr1
                            FROM          dbo.tblCalls
                            WHERE      (CallDate > DATEADD(day, - 180, GETDATE())) AND (CustomerId = t1.CustomerId)))
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customer_by_90day_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[23] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "t1"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 265
               Right = 198
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 6
               Left = 236
               Bottom = 278
               Right = 402
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomerContacts"
            Begin Extent = 
               Top = 6
               Left = 440
               Bottom = 338
               Right = 623
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblEmployees"
            Begin Extent = 
               Top = 6
               Left = 661
               Bottom = 290
               Right = 851
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 14
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customer_by_90day_no_call_log'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customer_by_90day_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'= 1410
         GroupBy = 1350
         Filter = 2745
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customer_by_90day_no_call_log'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customer_by_90day_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customer_by_90day_no_call_log'
GO
/****** Object:  View [dbo].[vwJKM_customer_by_30day_no_call_log]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_customer_by_30day_no_call_log]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_customer_by_30day_no_call_log
AS
SELECT DISTINCT t1.CustomerId, dbo.tblCustomers.SalesRepId
FROM         dbo.tblCalls AS t1 INNER JOIN
                      dbo.tblCustomers ON t1.CustomerId = dbo.tblCustomers.CustomerId INNER JOIN
                      dbo.tblCustomerContacts ON dbo.tblCustomers.CustomerId = dbo.tblCustomerContacts.CustomerId INNER JOIN
                      dbo.tblEmployees ON t1.EmployeeId = dbo.tblEmployees.EmployeeID
WHERE     (NOT EXISTS
                          (SELECT     1 AS Expr1
                            FROM          dbo.tblCalls
                            WHERE      (CallDate > DATEADD(day, - 30, GETDATE())) AND (CustomerId = t1.CustomerId)))
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customer_by_30day_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "t1"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 339
               Right = 198
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 6
               Left = 236
               Bottom = 333
               Right = 402
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomerContacts"
            Begin Extent = 
               Top = 6
               Left = 440
               Bottom = 346
               Right = 623
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblEmployees"
            Begin Extent = 
               Top = 6
               Left = 661
               Bottom = 345
               Right = 821
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customer_by_30day_no_call_log'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customer_by_30day_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customer_by_30day_no_call_log'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customer_by_30day_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customer_by_30day_no_call_log'
GO
/****** Object:  View [dbo].[vwJKM_shipping_log]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_shipping_log]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_shipping_log
AS
SELECT     TOP 100 PERCENT dbo.tblWorkOrders.WorkOrderNumber, dbo.tblWorkOrders.DateOut, dbo.tblWorkOrders.TrackingWaybill, dbo.tblCustomers.Customer, 
                      dbo.tblWorkOrders.DateRec, dbo.tblWorkOrders.DateIn, dbo.tblWorkOrders.Comments
FROM         dbo.tblWorkOrders INNER JOIN
                      dbo.tblProjects ON dbo.tblWorkOrders.ProjectId = dbo.tblProjects.ProjectId INNER JOIN
                      dbo.tblCustomers ON dbo.tblProjects.CustomerId = dbo.tblCustomers.CustomerId
WHERE     (dbo.tblWorkOrders.DateOut > DATEADD(Year, - 1, GETDATE()))
ORDER BY dbo.tblWorkOrders.DateOut DESC
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_shipping_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[41] 4[17] 2[24] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblWorkOrders"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 340
               Right = 249
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 0
               Left = 589
               Bottom = 342
               Right = 784
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblProjects"
            Begin Extent = 
               Top = 6
               Left = 330
               Bottom = 316
               Right = 510
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 4755
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 3165
         Width = 4215
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1485
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 2715
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_shipping_log'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_shipping_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_shipping_log'
GO
/****** Object:  View [dbo].[vwJKM_customer_mapping]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_customer_mapping]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_customer_mapping
AS
SELECT     Customer, Address, City, State, Country, Zip, SalesRepId
FROM         dbo.tblCustomers
WHERE     (SalesRepId = 42)
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customer_mapping', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[3] 4[3] 2[3] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 324
               Right = 204
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 4305
         Width = 1500
         Width = 1740
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customer_mapping'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_customer_mapping', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_customer_mapping'
GO
/****** Object:  View [dbo].[vwJKM_contacts_by_180_no_call_log]    Script Date: 1/14/2019 12:24:25 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vwJKM_contacts_by_180_no_call_log]'))
EXEC dbo.sp_executesql @statement = N'CREATE VIEW dbo.vwJKM_contacts_by_180_no_call_log
AS
SELECT DISTINCT 
                      dbo.tblCustomerContacts.CustomerContactId, dbo.tblCustomers.MainContactId, dbo.tblCustomers.SalesRepId, dbo.tblCustomers.Customer, dbo.tblCustomers.Address, 
                      dbo.tblCustomers.City, dbo.tblCustomers.State, dbo.tblCustomers.Zip, dbo.tblCustomerContacts.Contact, dbo.tblCustomerContacts.TelephoneNumber, 
                      dbo.tblCustomerContacts.Extension, dbo.tblCustomerContacts.EmailAddress, dbo.tblCustomerContacts.MobileNumber, dbo.tblCustomerContacts.ContactTitle, 
                      dbo.tblCustomers.Country
FROM         dbo.tblCalls AS t1 INNER JOIN
                      dbo.tblCustomers ON t1.CustomerId = dbo.tblCustomers.CustomerId INNER JOIN
                      dbo.tblCustomerContacts ON dbo.tblCustomers.CustomerId = dbo.tblCustomerContacts.CustomerId INNER JOIN
                      dbo.tblEmployees ON t1.EmployeeId = dbo.tblEmployees.EmployeeID
WHERE     (NOT EXISTS
                          (SELECT     1 AS Expr1
                            FROM          dbo.tblCalls
                            WHERE      (CallDate > DATEADD(day, - 180, GETDATE())) AND (CustomerId = t1.CustomerId))) AND (dbo.tblCustomers.Country = ''mexico'')
' 
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane1' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_contacts_by_180_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[28] 4[37] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "t1"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 231
               Right = 198
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomers"
            Begin Extent = 
               Top = 6
               Left = 236
               Bottom = 230
               Right = 402
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblCustomerContacts"
            Begin Extent = 
               Top = 6
               Left = 440
               Bottom = 229
               Right = 623
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "tblEmployees"
            Begin Extent = 
               Top = 6
               Left = 661
               Bottom = 218
               Right = 821
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 16
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_contacts_by_180_no_call_log'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPane2' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_contacts_by_180_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_contacts_by_180_no_call_log'
GO
IF NOT EXISTS (SELECT * FROM sys.fn_listextendedproperty(N'MS_DiagramPaneCount' , N'SCHEMA',N'dbo', N'VIEW',N'vwJKM_contacts_by_180_no_call_log', NULL,NULL))
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwJKM_contacts_by_180_no_call_log'
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProducts_tblCategories]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProducts]'))
ALTER TABLE [dbo].[tblProducts]  WITH CHECK ADD  CONSTRAINT [FK_tblProducts_tblCategories] FOREIGN KEY([CategoryID])
REFERENCES [dbo].[tblCategories] ([CategoryID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProducts_tblCategories]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProducts]'))
ALTER TABLE [dbo].[tblProducts] CHECK CONSTRAINT [FK_tblProducts_tblCategories]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProducts_tblProductOwners]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProducts]'))
ALTER TABLE [dbo].[tblProducts]  WITH CHECK ADD  CONSTRAINT [FK_tblProducts_tblProductOwners] FOREIGN KEY([ProductOwnerID])
REFERENCES [dbo].[tblProductOwners] ([ProductOwnerID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProducts_tblProductOwners]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProducts]'))
ALTER TABLE [dbo].[tblProducts] CHECK CONSTRAINT [FK_tblProducts_tblProductOwners]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblPurchaseOrders_tblShippingMethods]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblPurchaseOrders]'))
ALTER TABLE [dbo].[tblPurchaseOrders]  WITH CHECK ADD  CONSTRAINT [FK_tblPurchaseOrders_tblShippingMethods] FOREIGN KEY([ShippingMethodID])
REFERENCES [dbo].[tblShippingMethods] ([ShippingMethodID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblPurchaseOrders_tblShippingMethods]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblPurchaseOrders]'))
ALTER TABLE [dbo].[tblPurchaseOrders] CHECK CONSTRAINT [FK_tblPurchaseOrders_tblShippingMethods]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblPurchaseOrders_tblSuppliers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblPurchaseOrders]'))
ALTER TABLE [dbo].[tblPurchaseOrders]  WITH CHECK ADD  CONSTRAINT [FK_tblPurchaseOrders_tblSuppliers] FOREIGN KEY([SupplierID])
REFERENCES [dbo].[tblSuppliers] ([SupplierID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblPurchaseOrders_tblSuppliers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblPurchaseOrders]'))
ALTER TABLE [dbo].[tblPurchaseOrders] CHECK CONSTRAINT [FK_tblPurchaseOrders_tblSuppliers]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblInventoryTransactions_tblProducts]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblInventoryTransactions]'))
ALTER TABLE [dbo].[tblInventoryTransactions]  WITH CHECK ADD  CONSTRAINT [FK_tblInventoryTransactions_tblProducts] FOREIGN KEY([ProductID])
REFERENCES [dbo].[tblProducts] ([ProductID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblInventoryTransactions_tblProducts]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblInventoryTransactions]'))
ALTER TABLE [dbo].[tblInventoryTransactions] CHECK CONSTRAINT [FK_tblInventoryTransactions_tblProducts]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblInventoryTransactions_tblPurchaseOrders]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblInventoryTransactions]'))
ALTER TABLE [dbo].[tblInventoryTransactions]  WITH CHECK ADD  CONSTRAINT [FK_tblInventoryTransactions_tblPurchaseOrders] FOREIGN KEY([PurchaseOrderID])
REFERENCES [dbo].[tblPurchaseOrders] ([PurchaseOrderID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblInventoryTransactions_tblPurchaseOrders]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblInventoryTransactions]'))
ALTER TABLE [dbo].[tblInventoryTransactions] CHECK CONSTRAINT [FK_tblInventoryTransactions_tblPurchaseOrders]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProjects_tblCustomers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProjects]'))
ALTER TABLE [dbo].[tblProjects]  WITH CHECK ADD  CONSTRAINT [FK_tblProjects_tblCustomers] FOREIGN KEY([CustomerId])
REFERENCES [dbo].[tblCustomers] ([CustomerId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProjects_tblCustomers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProjects]'))
ALTER TABLE [dbo].[tblProjects] CHECK CONSTRAINT [FK_tblProjects_tblCustomers]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProjects_tblEmployees]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProjects]'))
ALTER TABLE [dbo].[tblProjects]  WITH CHECK ADD  CONSTRAINT [FK_tblProjects_tblEmployees] FOREIGN KEY([SalesRepId])
REFERENCES [dbo].[tblEmployees] ([EmployeeID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProjects_tblEmployees]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProjects]'))
ALTER TABLE [dbo].[tblProjects] CHECK CONSTRAINT [FK_tblProjects_tblEmployees]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProjects_tblProjectTypes]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProjects]'))
ALTER TABLE [dbo].[tblProjects]  WITH CHECK ADD  CONSTRAINT [FK_tblProjects_tblProjectTypes] FOREIGN KEY([ProjectTypeId])
REFERENCES [dbo].[tblProjectTypes] ([ProjectTypeId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProjects_tblProjectTypes]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProjects]'))
ALTER TABLE [dbo].[tblProjects] CHECK CONSTRAINT [FK_tblProjects_tblProjectTypes]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProjects_tblSpindles]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProjects]'))
ALTER TABLE [dbo].[tblProjects]  WITH CHECK ADD  CONSTRAINT [FK_tblProjects_tblSpindles] FOREIGN KEY([SpindleId])
REFERENCES [dbo].[tblSpindles] ([SpindleId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblProjects_tblSpindles]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblProjects]'))
ALTER TABLE [dbo].[tblProjects] CHECK CONSTRAINT [FK_tblProjects_tblSpindles]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblWorkOrders_tblProjects]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblWorkOrders]'))
ALTER TABLE [dbo].[tblWorkOrders]  WITH CHECK ADD  CONSTRAINT [FK_tblWorkOrders_tblProjects] FOREIGN KEY([ProjectId])
REFERENCES [dbo].[tblProjects] ([ProjectId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblWorkOrders_tblProjects]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblWorkOrders]'))
ALTER TABLE [dbo].[tblWorkOrders] CHECK CONSTRAINT [FK_tblWorkOrders_tblProjects]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblCalls_tblCustomers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblCalls]'))
ALTER TABLE [dbo].[tblCalls]  WITH CHECK ADD  CONSTRAINT [FK_tblCalls_tblCustomers] FOREIGN KEY([CustomerId])
REFERENCES [dbo].[tblCustomers] ([CustomerId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblCalls_tblCustomers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblCalls]'))
ALTER TABLE [dbo].[tblCalls] CHECK CONSTRAINT [FK_tblCalls_tblCustomers]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblCalls_tblEmployees]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblCalls]'))
ALTER TABLE [dbo].[tblCalls]  WITH CHECK ADD  CONSTRAINT [FK_tblCalls_tblEmployees] FOREIGN KEY([EmployeeId])
REFERENCES [dbo].[tblEmployees] ([EmployeeID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblCalls_tblEmployees]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblCalls]'))
ALTER TABLE [dbo].[tblCalls] CHECK CONSTRAINT [FK_tblCalls_tblEmployees]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuotes_tblWorkOrders]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuotes]'))
ALTER TABLE [dbo].[tblQuotes]  WITH CHECK ADD  CONSTRAINT [FK_tblQuotes_tblWorkOrders] FOREIGN KEY([WorkOrderId])
REFERENCES [dbo].[tblWorkOrders] ([WorkOrderId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuotes_tblWorkOrders]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuotes]'))
ALTER TABLE [dbo].[tblQuotes] CHECK CONSTRAINT [FK_tblQuotes_tblWorkOrders]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteBearings_tblQuotes]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteBearings]'))
ALTER TABLE [dbo].[tblQuoteBearings]  WITH CHECK ADD  CONSTRAINT [FK_tblQuoteBearings_tblQuotes] FOREIGN KEY([QuoteId])
REFERENCES [dbo].[tblQuotes] ([QuoteId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteBearings_tblQuotes]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteBearings]'))
ALTER TABLE [dbo].[tblQuoteBearings] CHECK CONSTRAINT [FK_tblQuoteBearings_tblQuotes]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteBearings_tblSuppliers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteBearings]'))
ALTER TABLE [dbo].[tblQuoteBearings]  WITH CHECK ADD  CONSTRAINT [FK_tblQuoteBearings_tblSuppliers] FOREIGN KEY([SupplierId])
REFERENCES [dbo].[tblSuppliers] ([SupplierID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteBearings_tblSuppliers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteBearings]'))
ALTER TABLE [dbo].[tblQuoteBearings] CHECK CONSTRAINT [FK_tblQuoteBearings_tblSuppliers]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteParts_tblQuotes]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteParts]'))
ALTER TABLE [dbo].[tblQuoteParts]  WITH CHECK ADD  CONSTRAINT [FK_tblQuoteParts_tblQuotes] FOREIGN KEY([QuoteId])
REFERENCES [dbo].[tblQuotes] ([QuoteId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteParts_tblQuotes]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteParts]'))
ALTER TABLE [dbo].[tblQuoteParts] CHECK CONSTRAINT [FK_tblQuoteParts_tblQuotes]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteParts_tblSuppliers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteParts]'))
ALTER TABLE [dbo].[tblQuoteParts]  WITH CHECK ADD  CONSTRAINT [FK_tblQuoteParts_tblSuppliers] FOREIGN KEY([SupplierId])
REFERENCES [dbo].[tblSuppliers] ([SupplierID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteParts_tblSuppliers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteParts]'))
ALTER TABLE [dbo].[tblQuoteParts] CHECK CONSTRAINT [FK_tblQuoteParts_tblSuppliers]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteSubWork_tblQuotes]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteSubWork]'))
ALTER TABLE [dbo].[tblQuoteSubWork]  WITH CHECK ADD  CONSTRAINT [FK_tblQuoteSubWork_tblQuotes] FOREIGN KEY([QuoteId])
REFERENCES [dbo].[tblQuotes] ([QuoteId])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteSubWork_tblQuotes]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteSubWork]'))
ALTER TABLE [dbo].[tblQuoteSubWork] CHECK CONSTRAINT [FK_tblQuoteSubWork_tblQuotes]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteSubWork_tblSuppliers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteSubWork]'))
ALTER TABLE [dbo].[tblQuoteSubWork]  WITH CHECK ADD  CONSTRAINT [FK_tblQuoteSubWork_tblSuppliers] FOREIGN KEY([SupplierId])
REFERENCES [dbo].[tblSuppliers] ([SupplierID])
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_tblQuoteSubWork_tblSuppliers]') AND parent_object_id = OBJECT_ID(N'[dbo].[tblQuoteSubWork]'))
ALTER TABLE [dbo].[tblQuoteSubWork] CHECK CONSTRAINT [FK_tblQuoteSubWork_tblSuppliers]
GO

/****** Adding SQLAdmin to fixed database roles    Script Date: 1/14/2019 12:24:26 AM ******/
ALTER ROLE [db_owner] ADD MEMBER [SQLAdmin]
GO

/****** Setting default schemas for migrated users    Script Date: 1/14/2019 12:24:26 AM ******/
ALTER USER [SQLAdmin] WITH DEFAULT_SCHEMA = [dbo]
GO
ALTER USER [chad] WITH DEFAULT_SCHEMA = [dbo]
GO
ALTER USER [boss] WITH DEFAULT_SCHEMA = [dbo]
GO

/****** Database Level Permissions    Script Date: 1/14/2019 12:24:26 AM ******/
GRANT CONNECT TO [boss]
    AS [dbo]
GO
GRANT CONNECT TO [chad]
    AS [dbo]
GO
GRANT CONNECT TO [SQLAdmin]
    AS [dbo]
GO

/****** Restoring object ownership for scripted objects      Script Date: 1/14/2019 12:24:28 AM ******/
ALTER AUTHORIZATION
    ON SCHEMA::[IUSR_SKMSOFT01]
    TO [IUSR_SKMSOFT01]
GO
ALTER AUTHORIZATION
    ON SCHEMA::[IUSR_INVENTORY]
    TO [IUSR_INVENTORY]
GO
ALTER AUTHORIZATION
    ON SCHEMA::[IUSR_W2KFILESRVR]
    TO [IUSR_W2KFILESRVR]
GO
ALTER AUTHORIZATION
    ON SCHEMA::[IUSR_DELL600SC]
    TO [IUSR_DELL600SC]
GO

