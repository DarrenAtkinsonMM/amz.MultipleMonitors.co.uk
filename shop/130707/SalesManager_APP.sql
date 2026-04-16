/****** Object:  StoredProcedure [dbo].[uspAddCatPrices]    Script Date: 04/26/2011 21:18:03 ******/
ALTER PROCEDURE [dbo].[uspAddCatPrices]
@Param1 nvarchar(1000),
@Param2 nvarchar(1000),
@IDCat nvarchar(7),
@CAmount nvarchar(10),
@CType int,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0;
	
	SET NOCOUNT ON;
	IF @CType=0
		SET @query='INSERT INTO pcCC_Pricing (idcustomerCategory,IDProduct,pcCC_Price) SELECT ' + @IDCat + ',Products.idProduct,Round(Products.Price*' + @CAmount + ',2) FROM ' + @Param1 + ' WHERE (Products.idProduct NOT IN (SELECT idProduct FROM pcCC_Pricing WHERE idcustomerCategory=' + @IDCat + ')) AND ' + @Param2 + ';'
	ELSE
		SET @query='INSERT INTO pcCC_Pricing(idcustomerCategory,IDProduct,pcCC_Price) SELECT ' + @IDCat + ',Products.idProduct,WPrice = CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @CAmount + ',2) ELSE Round(Products.bToBPrice*' + @CAmount + ',2) END FROM ' + @Param1 + ' WHERE (Products.idProduct NOT IN (SELECT idProduct FROM pcCC_Pricing WHERE idcustomerCategory=' + @IDCat + ')) AND ' + @Param2 + ';'
		
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT
	
	/*Add sub-product Cat Prices */
	IF @CType=0
		SET @query='INSERT INTO pcCC_Pricing (idcustomerCategory,IDProduct,pcCC_Price) SELECT ' + @IDCat + ',Products.idProduct,Round(Products.Price*' + @CAmount + ',2) FROM ' + @Param1 + ' WHERE (Products.idProduct NOT IN (SELECT idProduct FROM pcCC_Pricing WHERE idcustomerCategory=' + @IDCat + ')) AND (Products.pcProd_ParentPrd IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ')) AND Products.Removed=0;'
	ELSE
		SET @query='INSERT INTO pcCC_Pricing(idcustomerCategory,IDProduct,pcCC_Price) SELECT ' + @IDCat + ',Products.idProduct,WPrice = CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @CAmount + ',2) ELSE Round(Products.bToBPrice*' + @CAmount + ',2) END FROM ' + @Param1 + ' WHERE (Products.idProduct NOT IN (SELECT idProduct FROM pcCC_Pricing WHERE idcustomerCategory=' + @IDCat + ')) AND (Products.pcProd_ParentPrd IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ')) AND Products.Removed=0;'
	
	EXEC(@query)
	
	SET @SMCount=@SMCount+@@ROWCOUNT	
		
END
GO

/****** Object:  StoredProcedure [dbo].[uspBackUpPrices]    Script Date: 04/26/2011 21:19:18 ******/
ALTER PROCEDURE [dbo].[uspBackUpPrices]
@Param1 nvarchar(1000),
@Param2 nvarchar(1000),
@SCID nvarchar(7),
@SalesID nvarchar(7),
@TPrice nvarchar(10),
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	DECLARE @tmpQuery varchar(1000)
	
	SET @SMCount=0;
	
	SET NOCOUNT ON;
	
	IF @TPrice='0'
		SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',Products.idProduct,Products.Price FROM ' + @Param1 + ' WHERE ' + @Param2 + ';'
	ELSE
		BEGIN
		IF @TPrice='-1'
			SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',Products.idProduct,Products.bToBPrice FROM ' + @Param1 + ' WHERE ' + @Param2 + ';'
		ELSE
			SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',pcCC_Pricing.idProduct,pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=' + @TPrice + ' AND (pcCC_Pricing.IdProduct IN (SELECT Products.IdProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + '));'
		END
		
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT
	
	/*Backup Sub-Products*/
	SET @tmpQuery='SELECT Products.idProduct FROM Products WHERE (Products.pcProd_ParentPrd IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ')) AND Products.Removed=0'
		
	IF @TPrice='0'
		SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',Products.idProduct,Products.Price FROM Products WHERE (Products.IDProduct IN (' + @tmpQuery + '));'
	ELSE
		BEGIN
		IF @TPrice='-1'
			SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',Products.idProduct,Products.bToBPrice FROM Products WHERE (Products.IDProduct IN (' + @tmpQuery + '));'
		ELSE
			SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',pcCC_Pricing.idProduct,pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=' + @TPrice + ' AND (pcCC_Pricing.IdProduct IN (' + @tmpQuery + '));'
		END
		
	EXEC(@query)
	
	SET @SMCount=@SMCount+@@ROWCOUNT	
	
END
GO

/****** Object:  StoredProcedure [dbo].[uspChangePrices]    Script Date: 01/06/2012 17:05:26 ******/
ALTER PROCEDURE [dbo].[uspChangePrices]
@Param1 nvarchar(1000),
@Param2 nvarchar(1000),
@TPrice nvarchar(10),
@CType nvarchar(10),
@Relative nvarchar(10),
@Amount nvarchar(10),
@CRound nvarchar(10),
@SCID nvarchar(10),
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	DECLARE @query1 varchar(8000)
	DECLARE @tmpQuery varchar(8000)
	DECLARE @HasTmp int
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query=''
	
	/*Change sub-Product Prices*/
	
	SET @tmpQuery='SELECT Products.idProduct FROM Products WHERE (Products.pcProd_ParentPrd IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ')) AND Products.Removed=0'
	
	IF @CType='0'
	BEGIN
	
		IF @TPrice='0'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price*' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price*' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (' + @tmpQuery + ');'
		END
		
		IF @TPrice='-1'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',2) ELSE Round(Products.bToBPrice*' + @Amount + ',2) END'
			ELSE
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',0) ELSE Round(Products.bToBPrice*' + @Amount + ',0) END'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (' + @tmpQuery + ');'
		END
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct IN (' + @tmpQuery + '));'
		END
		
	END
	
	IF @CType='1'
	BEGIN
	
		IF @TPrice='0'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price-' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price-' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE (Products.IdProduct IN (' + @tmpQuery + '));'
		END
		
		IF @TPrice='-1'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price-' + @Amount + ',2) ELSE Round(Products.bToBPrice-' + @Amount + ',2) END'
			ELSE
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price-' + @Amount + ',0) ELSE Round(Products.bToBPrice-' + @Amount + ',0) END'
			
			SET @query=@query + ' WHERE (Products.IdProduct IN (' + @tmpQuery + '));'
		END
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price-' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price-' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct IN (' + @tmpQuery + '));'
		END
		
	END
		
	IF @query<>''
	BEGIN
		EXEC(@query)
		SET @SMCount=@@ROWCOUNT
		
	
		SET @query='UPDATE Products SET Products.pcSC_ID=' + @SCID + ' WHERE Products.IdProduct IN (' + @tmpQuery + ');' 
		EXEC(@query)
		
	END
	
	IF @CType='2'
	BEGIN
		
		SET @HasTmp=0
		
		IF @Relative='0'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				BEGIN
					SET @query='= Round(Products.Price*' + @Amount + ',2) '
					SET @query1=' Products.IdProduct IN (' + @tmpQuery + ') '
				END
			ELSE
				BEGIN
					SET @query='= Round(Products.Price*' + @Amount + ',0) '
					SET @query1=' Products.IdProduct IN (' + @tmpQuery + ') '
				END
		END
		
		IF @Relative='-1'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				BEGIN
					SET @query='= CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',2) ELSE Round(Products.bToBPrice*' + @Amount + ',2) END '
					SET @query1=' Products.IdProduct IN (' + @tmpQuery + ') '
				END
			ELSE
				BEGIN
					SET @query='= CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',0) ELSE Round(Products.bToBPrice*' + @Amount + ',0) END '
					SET @query1=' Products.IdProduct IN (' + @tmpQuery + ') '
				END
		END
		
		IF @Relative='-2'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				BEGIN
					SET @query='= CASE A.listPrice WHEN 0 THEN Round((B.listPrice+A.pcprod_AddPrice)*' + @Amount + ',2) ELSE Round(A.listPrice*' + @Amount + ',2) END '
					SET @query1=' A.IdProduct IN (' + @tmpQuery + ') AND B.idproduct=A.pcProd_ParentPrd '
				END
			ELSE
				BEGIN
					SET @query='= CASE A.listPrice WHEN 0 THEN Round((B.listPrice+A.pcprod_AddPrice)*' + @Amount + ',0) ELSE Round(A.listPrice*' + @Amount + ',0) END '
					SET @query1=' WHERE A.IdProduct IN (' + @tmpQuery + ') AND B.idproduct=A.pcProd_ParentPrd '
				END
		END
		
		IF @Relative='-3'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				BEGIN
					SET @query='= Round(Products.cost*' + @Amount + ',2) '
					SET @query1=' Products.IdProduct IN (' + @tmpQuery + ') '
				END
			ELSE
				BEGIN
					SET @query='= Round(Products.cost*' + @Amount + ',0) '
					SET @query1=' Products.IdProduct IN (' + @tmpQuery + ') '
				END
		END
		
		IF @HasTmp=0
		BEGIN
			IF @CRound='0'
				BEGIN
					SET @query='= Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',2) '
					SET @query1=' (pcCC_Pricing.idcustomerCategory=' + @Relative + ') AND (pcCC_Pricing.IdProduct IN (' + @tmpQuery + ')) '
				END
			ELSE
				BEGIN
					SET @query='= Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',0) '
					SET @query1=' (pcCC_Pricing.idcustomerCategory=' + @Relative + ') AND (pcCC_Pricing.IdProduct IN (' + @tmpQuery + ')) '
				END
		END
		
		IF @TPrice='0'
		BEGIN
			IF @HasTmp=0
				IF @Relative='-2'
					SET @query='UPDATE Products SET Products.Price' + @query + ' FROM Products,pcCC_Pricing,Products A,Products B WHERE A.idProduct=Products.idProduct AND Products.IdProduct=pcCC_Pricing.idProduct AND ' + @query1
				ELSE
					SET @query='UPDATE Products SET Products.Price' + @query + ' FROM Products,pcCC_Pricing WHERE Products.IdProduct=pcCC_Pricing.idProduct AND ' + @query1
			ELSE
				IF @Relative='-2'
					SET @query='UPDATE Products SET Products.Price' + @query + ' FROM Products,Products A,Products B WHERE A.idProduct=Products.idProduct AND ' + @query1
				ELSE
					SET @query='UPDATE Products SET Products.Price' + @query + ' FROM Products WHERE ' + @query1
		END

		
		IF @TPrice='-1'
		BEGIN
			IF @HasTmp=0
				IF @Relative='-2'
					SET @query='UPDATE Products SET Products.bToBPrice' + @query + ' FROM Products,pcCC_Pricing,Products A,Products B WHERE A.idProduct=Products.idProduct AND Products.IdProduct=pcCC_Pricing.idProduct AND ' + @query1
				ELSE
					SET @query='UPDATE Products SET Products.bToBPrice' + @query + ' FROM Products,pcCC_Pricing WHERE Products.IdProduct=pcCC_Pricing.idProduct AND ' + @query1
			ELSE
				IF @Relative='-2'
					SET @query='UPDATE Products SET Products.bToBPrice' + @query + ' FROM Products,Products A,Products B WHERE A.idProduct=Products.idProduct AND ' + @query1
				ELSE
					SET @query='UPDATE Products SET Products.bToBPrice' + @query + ' FROM Products WHERE ' + @query1
		END
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
		BEGIN
			IF @HasTmp=0
				IF @Relative='-2'
					SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price' + @query + ' FROM pcCC_Pricing,Products A, Products B WHERE A.idProduct=pcCC_Pricing.idProduct AND (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND ' + @query1
				ELSE
					SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price' + @query + ' FROM pcCC_Pricing WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND ' + @query1
			ELSE
				IF @Relative='-2'
					SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price' + @query + ' FROM pcCC_Pricing,Products,Products A, Products B WHERE A.idProduct=Products.idProduct AND (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct=Products.IDProduct) AND ' + @query1
				ELSE
					SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price' + @query + ' FROM pcCC_Pricing,Products WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct=Products.IDProduct) AND ' + @query1
		END
		
		EXEC(@query)
		SET @SMCount=@@ROWCOUNT
	
		
		
		SET @query='UPDATE Products SET Products.pcSC_ID=' + @SCID + ' WHERE Products.IdProduct IN (' + @tmpQuery + ');' 
		EXEC(@query)
		
	END
	
	/*Update Standard & Apparel Parent Products*/
	
	SET @query=''
	
	IF @CType='0'
	BEGIN
	
		IF @TPrice='0'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price*' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price*' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF @TPrice='-1'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',2) ELSE Round(Products.bToBPrice*' + @Amount + ',2) END'
			ELSE
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',0) ELSE Round(Products.bToBPrice*' + @Amount + ',0) END'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + '));'
		END
		
	END
	
	IF @CType='1'
	BEGIN
	
		IF @TPrice='0'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price-' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price-' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF @TPrice='-1'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price-' + @Amount + ',2) ELSE Round(Products.bToBPrice-' + @Amount + ',2) END'
			ELSE
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price-' + @Amount + ',0) ELSE Round(Products.bToBPrice-' + @Amount + ',0) END'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price-' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price-' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + '));'
		END
		
	END
		
	IF @query<>''
	BEGIN
		EXEC(@query)
		SET @SMCount=@SMCount+@@ROWCOUNT
		
		SET @query='UPDATE Products SET Products.pcSC_ID=' + @SCID + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');' 
		EXEC(@query)
	END
	
	IF @CType='2'
	BEGIN
		
		SET @HasTmp=0
		
		IF @Relative='0'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				BEGIN
					SET @query='= Round(Products.Price*' + @Amount + ',2) '
					SET @query1=' Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ') ' 
				END
			ELSE
				BEGIN
					SET @query='= Round(Products.Price*' + @Amount + ',0) '
					SET @query1=' Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ') '
				END
		END
		
		IF @Relative='-1'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				BEGIN
					SET @query='= CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',2) ELSE Round(Products.bToBPrice*' + @Amount + ',2) END '
					SET @query1=' Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ') ' 
				END
			ELSE
				BEGIN
					SET @query='= CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',0) ELSE Round(Products.bToBPrice*' + @Amount + ',0) END '
					SET @query1=' Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ') '
				END
		END
		
		IF @Relative='-2'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				BEGIN
					SET @query='= Round(Products.listPrice*' + @Amount + ',2) '
					SET @query1=' Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ') ' 
				END
			ELSE
				BEGIN
					SET @query='= Round(Products.listPrice*' + @Amount + ',0) '
					SET @query1=' Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ') '
				END
		END
		
		IF @Relative='-3'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				BEGIN
					SET @query='= Round(Products.cost*' + @Amount + ',2) '
					SET @query1=' Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ') ' 
				END
			ELSE
				BEGIN
					SET @query='= Round(Products.cost*' + @Amount + ',0) '
					SET @query1=' Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ') '
				END
		END
		
		IF @HasTmp=0
		BEGIN
			IF @CRound='0'
				BEGIN
					SET @query='= Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',2) '
					SET @query1=' (pcCC_Pricing.idcustomerCategory=' + @Relative + ') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ')) ' 
				END
			ELSE
				BEGIN
					SET @query='= Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',0) '
					SET @query1=' (pcCC_Pricing.idcustomerCategory=' + @Relative + ') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ')) '
				END
		END
		
		IF @TPrice='0'
		BEGIN
			IF @HasTmp=0
				SET @query='UPDATE Products SET Products.Price' + @query + ' FROM Products,pcCC_Pricing WHERE Products.IdProduct=pcCC_Pricing.idProduct AND ' + @query1
			ELSE
				SET @query='UPDATE Products SET Products.Price' + @query + ' FROM Products WHERE ' + @query1
		END
		
		IF @TPrice='-1'
		BEGIN
			IF @HasTmp=0
				SET @query='UPDATE Products SET Products.bToBPrice' + @query + ' FROM Products,pcCC_Pricing WHERE Products.IdProduct=pcCC_Pricing.idProduct AND ' + @query1
			ELSE
				SET @query='UPDATE Products SET Products.bToBPrice' + @query + ' FROM Products WHERE ' + @query1
		END
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
		BEGIN
			IF @HasTmp=0
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price' + @query + ' FROM pcCC_Pricing WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND ' + @query1
			ELSE
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price' + @query + ' FROM pcCC_Pricing,Products WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct=Products.IDProduct) AND ' + @query1
		END
		
		EXEC(@query)
		SET @SMCount=@SMCount+@@ROWCOUNT
		
		SET @query='UPDATE Products SET Products.pcSC_ID=' + @SCID + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');' 
		EXEC(@query)
		
	END
	
	print @SMCount
	
END
GO

/****** Object:  StoredProcedure [dbo].[uspGetPrdCount]    Script Date: 04/26/2011 21:21:24 ******/
ALTER PROCEDURE [dbo].[uspGetPrdCount]
@Param1 nvarchar(1000) ,
@Param2 nvarchar(1000),
@SMCount int Output,
@SMCountSub int Output
AS
BEGIN
	DECLARE @query varchar(8000),@tmpCount int,@tmpCount1 int
	
	SET @SMCount=0;
	SET @SMCountSub=0;
	
	SET NOCOUNT ON;
	
	SET @query='SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ';'
	EXEC(@query)
	
	SET @tmpCount=@@ROWCOUNT
	
	SET @query='SELECT Products.idProduct FROM Products WHERE (Products.pcProd_ParentPrd IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ')) AND Products.Removed=0;'
	EXEC(@query)
	
	SET @tmpCount1=@@ROWCOUNT
	
	SET @SMCount=@tmpCount
	SET @SMCountSub=@tmpCount1

END
GO

/****** Object:  StoredProcedure [dbo].[uspRmvPrdFromSale]    Script Date: 04/26/2011 21:23:31 ******/
ALTER PROCEDURE [dbo].[uspRmvPrdFromSale]
@SCID nvarchar(10),
@IDPrd nvarchar(10),
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	DECLARE @tmpQuery varchar(1000)
	DECLARE @TPrice int,@Apparel int
	
	SET @SMCount=0;
	
	SET NOCOUNT ON;
	
	SET @query='UPDATE Products SET Products.pcSC_ID=Products.active,Products.active=0 WHERE Products.IDProduct =' + @IDPrd + ';'
	EXEC(@query)
	
	SELECT TOP 1 @TPrice=pcSales_TargetPrice FROM pcSales_BackUp WHERE pcSC_ID=@SCID
	
	IF @TPrice=0
		SET @query='UPDATE Products SET Products.Price=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcSales_BackUp.IDProduct=' + @IDPrd + ';'
	
	IF @TPrice=-1
		SET @query='UPDATE Products SET Products.bToBPrice=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcSales_BackUp.IDProduct=' + @IDPrd + ';'
		
	IF @TPrice>0
		SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=pcSales_BackUp.pcSB_Price FROM pcCC_Pricing, pcSales_BackUp WHERE pcCC_Pricing.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcCC_Pricing.idcustomerCategory=' + @TPrice + ' AND pcSales_BackUp.IDProduct=' + @IDPrd + ';'
		
	EXEC(@query)
	
	SET @query='UPDATE Products SET Products.active=Products.pcSC_ID,Products.pcSC_ID=0 WHERE Products.IDProduct =' + @IDPrd + ';'
	EXEC(@query)
	
	SET @query='DELETE FROM pcSales_BackUp WHERE pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcSales_BackUp.IDProduct=' + @IDPrd + ';'
	EXEC(@query)
	
	SET @query='UPDATE pcSales_Completed SET pcSC_BUTotal=(SELECT Count(*) FROM Products WHERE Products.pcSC_ID=' + @SCID + ') WHERE pcSales_Completed.pcSC_ID=' + @SCID + ';'
	EXEC(@query)
	
	/*Remove Sub-Products from Sale*/
	SELECT TOP 1 @Apparel=pcProd_Apparel FROM Products WHERE Products.IDProduct=@IDPrd
	
	IF @Apparel=1
	BEGIN
	
		SET @tmpQuery='SELECT Products.idProduct FROM Products WHERE Products.pcProd_ParentPrd = ' + @IDPrd + ' AND Products.Removed=0'
		
		IF @TPrice=0
			SET @query='UPDATE Products SET Products.Price=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND (pcSales_BackUp.IDProduct IN (' + @tmpQuery + '));'
	
		IF @TPrice=-1
			SET @query='UPDATE Products SET Products.bToBPrice=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND (pcSales_BackUp.IDProduct IN (' + @tmpQuery + '));'
		
		IF @TPrice>0
			SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=pcSales_BackUp.pcSB_Price FROM pcCC_Pricing, pcSales_BackUp WHERE pcCC_Pricing.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcCC_Pricing.idcustomerCategory=' + @TPrice + ' AND (pcSales_BackUp.IDProduct IN (' + @tmpQuery + '));'
		
		EXEC(@query)
	
		SET @query='UPDATE Products SET Products.pcSC_ID=0 WHERE (Products.IDProduct IN (' + @tmpQuery + '));'
		EXEC(@query)
	
		SET @query='DELETE FROM pcSales_BackUp WHERE pcSales_BackUp.pcSC_ID=' + @SCID + ' AND (pcSales_BackUp.IDProduct IN (' + @tmpQuery + '));'
		EXEC(@query)
	
		SET @query='UPDATE pcSales_Completed SET pcSC_BUTotal=(SELECT Count(*) FROM Products WHERE Products.pcSC_ID=' + @SCID + ') WHERE pcSales_Completed.pcSC_ID=' + @SCID + ';'
		EXEC(@query)
		
	END
	
END
GO

