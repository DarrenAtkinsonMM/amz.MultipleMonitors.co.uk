<%
Dim app_HideNotAvailableItems,app_DisplayFinalPrice,app_DisplayWaitingBox,app_WaitingMsg

'Default Settings
app_HideNotAvailableItems=1
app_DisplayFinalPrice=0
app_DisplayWaitingBox=0
app_WaitingMsg="Please wait..."

query="SELECT pcAS_HideUItems,pcAS_PriceDiff,pcAS_TurnWB,pcAS_WMsg FROM pcApparelSettings;"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	app_HideNotAvailableItems=rsQ("pcAS_HideUItems")
	if IsNull(app_HideNotAvailableItems) OR app_HideNotAvailableItems="" then
		app_HideNotAvailableItems=1
	end if
	app_DisplayFinalPrice=rsQ("pcAS_PriceDiff")
	if IsNull(app_DisplayFinalPrice) OR app_DisplayFinalPrice="" then
		app_DisplayFinalPrice=0
	end if
	app_DisplayWaitingBox=rsQ("pcAS_TurnWB")
	if IsNull(app_DisplayWaitingBox) OR app_DisplayWaitingBox="" then
		app_DisplayWaitingBox=0
	end if
	app_WaitingMsg=rsQ("pcAS_WMsg")
	if IsNull(app_WaitingMsg) OR app_WaitingMsg="" then
		app_WaitingMsg="Please wait..."
	end if
end if
set rsQ=nothing

Dim pDefaultWeight1,pDefaultWeight2,pweight1
Dim ParentWeight, ParentRW
Dim HaveSale,pcSCID,APPshowVAT

HaveSale=0
pcSCID=0
APPshowVAT=0

' If the store is using and showing VAT, show the VAT included message and price without VAT
if ptaxVAT="1" and ptaxdisplayVAT="1" and pnotax <> "-1" then
	if session("customerType")="1" AND ptaxwholesale="0" then
	else
		APPshowVAT=1
	end if
end if

Function CheckSale()
Dim tmp1,query,rsS

	tmp1=0

			query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon,pcSales_BackUp.pcSales_TargetPrice FROM (pcSales_Completed INNER JOIN Products ON pcSales_Completed.pcSC_ID=Products.pcSC_ID) INNER JOIN pcSales_BackUp ON pcSales_BackUp.pcSC_ID=pcSales_Completed.pcSC_ID WHERE Products.idproduct=" & pidProduct & " AND Products.pcSC_ID>0;"
			set rsS=Server.CreateObject("ADODB.Recordset")
			set rsS=conntemp.execute(query)
					
			if not rsS.eof then
				ShowSaleIcon=1
				pcSCID=rsS("pcSC_ID")
				pcSCName=rsS("pcSC_SaveName")
				pcSCIcon=rsS("pcSC_SaveIcon")
				pcTargetPrice=rsS("pcSales_TargetPrice")
				if (ShowSaleIcon=1) AND (pcTargetPrice="0") AND (session("customertype")="0") AND (session("customerCategory")="0")then
					tmp1=1
				else
					if (ShowSaleIcon=1) AND (pcTargetPrice="1") AND (session("customertype")="1") AND (session("customerCategory")="0")then
						tmp1=1
					end if
				end if
			end if
			set rsS=nothing
		
	CheckSale=tmp1

End Function

Public Sub GenApparelSubProducts()
Dim query,rs,rs1,pcv_GrpCount,pcv_OptCount,pcArray,intCount
Dim i,j

HaveSale=CheckSale()

query="SELECT weight, pcprod_QtyToPound FROM Products WHERE idproduct=" & pidProduct & ";"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	ParentWeight=rsQ("weight")
	ParentQtyToPound=rsQ("pcprod_QtyToPound")
	if ParentWeight="" OR IsNull(ParentWeight) then
		ParentWeight=0
	end if
	if ParentQtyToPound="" OR IsNull(ParentQtyToPound) then
		ParentQtyToPound=0
	end if
end if
set rsQ=nothing

query="SELECT iRewardPoints FROM Products WHERE idproduct=" & pidProduct & ";"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	ParentRW=rsQ("iRewardPoints")
	if ParentRW="" OR IsNull(ParentRW) then
		ParentRW=0
	end if
end if
set rsQ=nothing

pweight1=pweight
pDefaultWeight1=0
pDefaultWeight2=0
if scShipFromWeightUnit="KGS" then
	pDefaultWeight1=Int(pWeight1/1000)
	pDefaultWeight2=pWeight1-(pDefaultWeight1*1000)
else
	pDefaultWeight1=Int(pWeight1/16)
	pDefaultWeight2=pWeight1-(pDefaultWeight1*16)
end if

IF HaveDiffPrice=0 then
pPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
pBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
if pserviceSpec=true then
	pPrice=Cdbl(pPrice+iAddDefaultPrice)
	pBtoBPrice=Cdbl(pBtoBPrice+iAddDefaultWPrice)
	pPrice1=Cdbl(pPrice1+iAddDefaultPrice1)
	pBtoBPrice1=Cdbl(pBtoBPrice1+iAddDefaultWPrice1)
end if

if session("customertype")=1 and pBtoBPrice1>0 then
	pPrice1=pBtoBPrice1
end if
ELSE
	if ccur(pBtoBPrice)=0 then
		pBtoBPrice=pPrice
	end if
	pPrice1=pPrice
	pBtoBPrice1=pBtoBPrice
END IF
		

pq_Price=request.QueryString("Price")
if pq_Price="" then
	pq_Price=0
end if
pq_AddPrice=request.QueryString("AddPrice")
if pq_AddPrice="" then
	pq_AddPrice=0
end if
pq_IDField=request.QueryString("IDField")
if pq_IDField="" then
	pq_IDField=""
end if
pq_VIndex=request.QueryString("vindex")
if pq_VIndex="" then
	pq_VIndex=0
end if%>
<%if app_DisplayWaitingBox=1 then%>
<div id="waitbox" class="pcErrorMessage" style="z-index:51; position: absolute; visibility:hidden; width: 25%;	background-color: #F7F7F7; border: 1px solid #0099FF; margin: 15px;	padding: 4px; color: #0066FF; font-size:12px; font-weight: bold; text-align: center;"><br><img src="<%=pcf_getImagePath("images","pleasewait.gif")%>"> <%=app_WaitingMsg%><br><br></div>
<div id="darkenScreenObject" style="position: absolute; overflow:hidden; display:none; top: 0px; left: 0px;"></div>
<script>
function grayOut(vis, options) {
  // Pass true to gray out screen, false to ungray
  // options are optional.  This is a JSON object with the following (optional) properties
  // opacity:0-100         // Lower number = less grayout higher = more of a blackout 
  // zindex: #             // HTML elements with a higher zindex appear on top of the gray out
  // bgcolor: (#xxxxxx)    // Standard RGB Hex color code
  // grayOut(true, {'zindex':'50', 'bgcolor':'#0000FF', 'opacity':'70'});
  // Because options is JSON opacity/zindex/bgcolor are all optional and can appear
  // in any order.  Pass only the properties you need to set.
  var options = options || {}; 
  var zindex = options.zindex || 50;
  var opacity = options.opacity || 70;
  var opaque = (opacity / 100);
  var bgcolor = options.bgcolor || '#000000';
  var dark=document.getElementById('darkenScreenObject');
  if (vis) {
    // Calculate the page width and height 
    if( document.body && ( document.body.scrollWidth || document.body.scrollHeight ) ) {
        var pageWidth = document.body.scrollWidth+'px';
        var pageHeight = document.body.scrollHeight+'px';
    } else if( document.body.offsetWidth ) {
      var pageWidth = document.body.offsetWidth+'px';
      var pageHeight = document.body.offsetHeight+'px';
    } else {
       var pageWidth='100%';
       var pageHeight='100%';
    }   
    //set the shader to cover the entire page and make it visible.
    dark.style.opacity=opaque;                      
    dark.style.MozOpacity=opaque;                   
    dark.style.filter='alpha(opacity='+opacity+')'; 
    dark.style.zIndex=zindex;        
    dark.style.backgroundColor=bgcolor;  
    dark.style.width= pageWidth;
    dark.style.height= pageHeight;
    dark.style.display='block';				 
  } else {
     dark.style.display='none';
  }
}
</script>
<script>
var ie=document.all
var ns6=document.getElementById && !document.all
if (ie||ns6)
	var waitBoxobj=document.all? document.all["waitbox"] : document.getElementById? document.getElementById("waitbox") : ""

function ietruebody()
{
	return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

var winheight=ie&&!window.opera? ietruebody().clientHeight : window.innerHeight
var winwidth=ie&&!window.opera? ietruebody().clientWidth : window.innerWidth-20

</script>
<%end if%>
<script>
function setMainImg(img, largeImg)
{

	if (largeImg !== undefined) {
		$pc("#mainimg").attr('src', img);
        if (pcv_strUseEnhancedViews) {
			$pc("#mainimg").attr("href", largeImg);
			$pc("#mainimg").parent().attr("href", largeImg);
		} else {
			$pc("#mainimg").parent().attr("href", largeImg);
			$pc("#zoombutton").attr("href", largeImg);
		}

		if (pcv_strIsMojoZoomEnabled) {
			mainImgMakeZoomable(largeImg);
		}
	} else {
        $pc("#mainimg").attr('src', img);
		mainImgMakeZoomable(img);
	}
}
        function mainImgMakeZoomable(setLink)
        {
            var tmpimg = new Image();
	
			tmpimg.src=document.getElementById("mainimg").src;
		
			tmpimg.onload = function(){
			var defaultWidth = 256;
            var defaultHeight = 256;
            
            var imageWidth = $pc("#mainimg").outerWidth();
            var imageHeight = $pc("#mainimg").outerHeight();
            
            zoomWidth = defaultWidth;
            zoomHeight = defaultHeight;
            
            if (imageWidth < defaultWidth && imageHeight < defaultHeight && imageWidth > 0 && imageHeight > 0) {
                if (imageWidth < imageHeight) {
                    zoomWidth = imageWidth;
                    zoomHeight = imageWidth;
                } else {
                    zoomWidth = imageHeight;
                    zoomHeight = imageHeight;
                }
            }
        
            MojoZoom.makeZoomable(document.getElementById("mainimg"), setLink, '', zoomWidth, zoomHeight, false);
		}
	}


	function New_FormatNumber(tmpvalue)
	{
	var DifferenceTotal = new NumberFormat();
	DifferenceTotal.setNumber(tmpvalue);
	if (scDecSign==",")
	{
		DifferenceTotal.setSeparators(true,DifferenceTotal.PERIOD);
	}
	else
	{
		DifferenceTotal.setCommas(true);
	}
	DifferenceTotal.setPlaces(2);
	DifferenceTotal.setCurrency(true);
	DifferenceTotal.setCurrencyPrefix(scCurSign);
	return(DifferenceTotal.toFormatted());
	}

	function RmvComma(tmpNum)
	{
		var tmp1=tmpNum + "";
		if (CommaSign==1)
		{
			tmp1=tmp1.replace(/\./gi,"");
			tmp1=tmp1.replace(/\,/gi,".");
		}
		else
		{
			tmp1=tmp1.replace(/\,/gi,"");
		}
		<%
		tmpStr=""
		For i=1 to len(scCurSign)
			tmpStr=tmpStr & "\" & mid(scCurSign,i,1)
		Next%>
		tmp1=tmp1.replace(/<%=tmpStr%>/gi,"");
		return(Number(tmp1));
	}
	
	function RepComma(tmpNum)
	{
		var tmp1=tmpNum + "";
		if (CommaSign==1)
		{
			tmp1=tmp1.replace(/\,/gi,"!");
			tmp1=tmp1.replace(/\./gi,",");
			tmp1=tmp1.replace(/\!/gi,".");
		}
		return(tmp1);
	}
	
	<%
	pcOrgPrice=0
	if HaveSale=1 then
		query="SELECT pcSB_Price FROM pcSales_BackUp WHERE idProduct=" & pIdProduct & " AND pcSC_ID=" & pcSCID & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			pcOrgPrice=rsQ("pcSB_Price")
		end if
		set rsQ=nothing
	end if%>
	
	var scDecSign="<%=scDecSign%>";
	var scCurSign="<%=scCurSign%>";
	var MyAccept=0;
	var LowStock=0;
	var LowPrdName="";
	var LowPrdStock=0;
	<%if APPshowVAT=1 then%>
	var DefaultVAT="<%=scCurSign & money(pcf_RemoveVAT(pPrice,pIdProduct))%>";
	<%else%>
	var DefaultVAT="<%=scCurSign & money(0)%>";
	<%end if%>
	var DefaultSku="<%=psku%>";
	var DefaultWeight1=<%=replace(pDefaultWeight1,",",".")%>;
	var DefaultWeight2=<%=replace(pDefaultWeight2,",",".")%>;
	var DefaultPrice="<%=scCurSign & money(pPrice)%>";
	var DefaultBackPrice="<%=scCurSign & money(pcOrgPrice)%>";
	<%tmpSavePrice=pPrice%>
	<%if session("customerCategory")<>0 then%>
	var DefaultWPrice="<%=scCurSign & money(pPrice1)%>";
	<%tmpSaveWPrice=pPrice1%>
	<%else%>
	var DefaultWPrice="<%=scCurSign & money(pBtoBPrice1)%>";
	<%tmpSaveWPrice=pBtoBPrice1%>
	<%end if%>
	<%if instr(money(pPrice),",")>instr(money(pPrice),".") then%>
	var CommaSign=1;
	<%else%>
	var CommaSign=0;
	<%end if%>
	<%DefaultLPrice=0
	DefaultSavings=0
	DefaultSavingsP=0
	if pListPrice-pPrice>0 then
		DefaultLPrice=pListPrice
		DefaultSavings=pListPrice-pPrice
		DefaultSavingsP=round(((pListPrice-pPrice)/pListPrice)*100)
	end if%>
	var DefaultLPrice=<%=replace(pListPrice,",",".")%>;
	var DefaultSavings=<%=replace(DefaultSavings,",",".")%>;
	var DefaultSavingsP=<%=replace(DefaultSavingsP,",",".")%>;
	var DefaultIDPrd=<%=pIDProduct%>;
	var SelectedSP=0;
	var SaveList=new Array();
	var SaveSKUList=new Array();
	var SaveDescList=new Array();
	var SaveQtyList=new Array();
	var SavedCount=0;
	var subprd_InActive=0;
	var subprd_NotAvailablePrd=0;
	var subprd_OOS=0;
	var subprd_PrdPrice=0;
	
	function makeGrpLevel()
	{
		var c=this; c.opt=new Array();
		c.name=new Array();
		return c
	}
	function makeGrp()
	{
		var c=this; c.grp=new Array();
		c.count=new Array();
		return this
	}
	function checkSPList()
	{
		if (LowPrdStock<document.additem.quantity.value)
		{
			LowStock=1;
		}
		else
		{
			LowStock=0;
		}
		if ((MyAccept==0) && (SavedCount==0))
		{
			alert("<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg2")%>\n");
			return(false);
		}
		else
		{
			if ((LowStock==1) && (SavedCount==0))
			{
				alert("<%=dictLanguage.Item(Session("language")&"_instPrd_2")%>" + LowPrdName + "<%=dictLanguage.Item(Session("language")&"_instPrd_3")%>" + LowPrdStock + "<%=dictLanguage.Item(Session("language")&"_instPrd_4")%>");
				return(false);
			}	
		}
		return(true);
	}
</script>
<%
'Generate Options List

	query="SELECT idOptionGroup FROM pcProductsOptions WHERE idproduct=" & pidProduct & " ORDER BY pcProdOpt_order ASC;"
	set rs=connTemp.execute(query)
	pcv_GrpCount=0%>
	<script>
		optGrp=new makeGrp();
		<%
		if not rs.eof then
			DO WHILE not rs.eof
				pcv_tmpIDGrp=rs("idOptionGroup")
				pcv_GrpCount=pcv_GrpCount+1
				query = 		"SELECT options_optionsGroups.idoptoptgrp, options.optiondescrip "
				query = query & "FROM options_optionsGroups "
				query = query & "INNER JOIN options "
				query = query & "ON options_optionsGroups.idOption = options.idOption "
				query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcv_tmpIDGrp &" "
				query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
				query = query & "ORDER BY options_optionsGroups.sortOrder;"
				set rs1=conntemp.execute(query)
				pcv_OptCount=0
				if not rs1.eof then
					pcArray=rs1.GetRows()
					intCount=ubound(pcArray,2)
					pcv_OptCount=intCount+1
					set rs1=nothing
					%>
					optGrp.grp[<%=pcv_GrpCount-1%>]=new makeGrpLevel();
					optGrp.count[<%=pcv_GrpCount-1%>]=<%=pcv_OptCount%>;
					<%For i=0 to intcount%>
						optGrp.grp[<%=pcv_GrpCount-1%>].opt[<%=i%>]=<%=pcArray(0,i)%>;
						optGrp.grp[<%=pcv_GrpCount-1%>].name[<%=i%>]="<%=replace(replace(pcArray(1,i),"""","\"""),"&quot;","\""")%>";
					<%Next
				end if
				set rs1=nothing
				rs.MoveNext
			LOOP
		end if
		set rs=nothing%>
		GrpCount=<%=pcv_GrpCount%>;
	</script>
<%	
'End Generate Options List

'Look-up subproducts
IF pcv_Apparel="1" then
	pcv_SPCount=0
	query="SELECT idproduct,sku,imageUrl,largeImageURL,stock,price,btoBPrice,pcprod_Relationship,pcProd_SPInActive,pcProd_AddPrice,pcProd_AddWPrice,noStock,pcProd_BackOrder,pcProd_ShipNDays,weight,description,iRewardPoints,pcprod_QtyToPound,listPrice FROM Products WHERE pcprod_ParentPrd=" & pIDProduct & " AND removed=0 "
	'if pcv_ShowStockMsg="0" then
	'	query=query & " AND ((stock>0) OR (noStock<>0) OR (pcProd_BackOrder<>0))"
	'end if
	query=query & " ORDER BY idproduct ASC;"
	set rs=connTemp.execute(query)

	if not rs.eof then%>
	<script type="text/javascript" src="<%=pcf_getJSPath("../includes","formatNumber154.js")%>"></script>
	<script>    
		var ns6=document.getElementById&&!document.all
		var ie=document.all
		
		<%
		'Override
		if pcv_ShowStockMsg="2" then
			app_HideNotAvailableItems=1
		else
			if pcv_ShowStockMsg="3" then
				app_HideNotAvailableItems=0
			end if
		end if%>
		
		var app_HideItems=<%=app_HideNotAvailableItems%>;
	
		function makeLevel()
		{
			var c=this, a=arguments; c.IDProduct=a[0]||null; c.sku=a[1]||null;
			c.img=a[2]||null;c.limg=a[3]||null;c.stock=a[4]||null;c.price=a[5]||null;c.wprice=a[6]||null;
			c.opts=new Array();c.nostock=a[7]||null;c.backorder=a[8]||null;c.inactive=a[9]||null;c.addprice=a[10]||null;c.addwprice=a[11]||null;
			c.ndays=a[12]||null;c.retext=a[13]||null;c.prdname=a[14]||null;c.weight1=a[15]||null;c.weight2=a[16]||null;c.reward=a[17]||null;c.lprice=a[18]||null;c.backprice=a[19]||null;c.vat=a[20]||null;
			return c
		}

		function makeArr()
		{
			var c=this; c.o=new Array();
			return this
		}
		
		sp=new makeArr();

		<%
		strImgFiles=""
		pcArray=rs.getRows()
		lngCount=ubound(pcArray,2)
		set rs=nothing
	
		Response.Flush
		Response.Clear
		
		For i=0 to lngCount
			pcv_HaveSPs="ok"
			pcv_SPCount=pcv_SPCount+1
			pcv_SPid=pcArray(0,i)
			pcv_SPsku=pcArray(1,i)
			pcv_SPimg=pcArray(2,i)
			'if pcv_SPimg<>"" then
			'else
			'pcv_SPimg="no_image.gif"
			'end if
			pcv_SPlimg=pcArray(3,i)
			pcv_SPStock=pcArray(4,i)
			if IsNull(pcv_SPStock) or pcv_SPStock="" then
				pcv_SPStock=0
			end if
			if cdbl(pcv_SPStock)<0 then
				pcv_SPStock=0
			end if

			pcv_SPPrice=pcArray(5,i)
			pcv_SPWPrice=pcArray(6,i)
			IF HaveDiffPrice=0 THEN
			if (pcv_SPWPrice<>"") and (pcv_SPWPrice<>"0") then
			else
				pcv_SPWPrice=pcv_SPPrice
			end if
			
			pcv_SPPrice1=CheckParentPrices(pcv_SPid,pcv_SPPrice,pcv_SPWPrice,0)
			pcv_SPWPrice1=CheckParentPrices(pcv_SPid,pcv_SPPrice,pcv_SPWPrice,1)
			
			if session("customerCategory")<>0 then
				pcv_SPWPrice=pcv_SPPrice1
			else
				if (pcv_SPWPrice1>"0") and (session("customerType")=1) then
					pcv_SPWPrice=pcv_SPWPrice1
				end if
			end if
			ELSE
				pcv_SPPrice=pPrice
				pcv_SPPrice1=pPrice
				pcv_SPWPrice=pBtoBPrice
				pcv_SPWPrice1=pBtoBPrice
			END IF
	
			pcv_Relationship=pcArray(7,i)
			TempArr=split(pcv_Relationship,"_")
				
			pcv_SPInactive=pcArray(8,i)
			if IsNull(pcv_SPInactive) or pcv_SPInactive="" then
				pcv_SPInactive="0"
			end if
	
			pcv_AddPrice=pcArray(9,i)
			if IsNull(pcv_AddPrice) or pcv_AddPrice="" then
				pcv_AddPrice="0"
			end if
			if session("customerCategory")<>0 then
				pcv_AddPrice=pcv_SPWPrice-pPrice
			end if
			pcv_AddPrice1=pcv_AddPrice
			pcv_AddWPrice=pcArray(10,i)
			if IsNull(pcv_AddWPrice) or pcv_AddWPrice="" then
				pcv_AddWPrice="0"
			end if
			pcv_AddWPrice1=pcv_AddWPrice
			if HaveSale=1 then
				pcv_AddPrice=pcv_SPPrice-tmpSavePrice
				pcv_AddWPrice=pcv_SPWPrice-tmpSaveWPrice
			end if
			pcv_ListPrice=pcArray(18,i)
			if IsNull(pcv_ListPrice) or pcv_ListPrice="" then
				pcv_ListPrice="0"
			end if
			if (HaveSale=1) AND (pcv_ListPrice="0") then
				pcv_ListPrice=Cdbl(pListPrice)+Cdbl(pcv_AddWPrice1)
			end if
			IF HaveDiffPrice=1 THEN
				pcv_AddPrice=0
				pcv_AddWPrice=0
				pcv_AddWPrice=0
				pcv_AddWPrice1=0
				pcv_ListPrice=0
			END IF
			pcv_DisStock=pcArray(11,i)
			if IsNull(pcv_DisStock) or pcv_DisStock="" then
				pcv_DisStock="0"
			end if
			pcv_BackOrder=pcArray(12,i)
			if IsNull(pcv_BackOrder) or pcv_BackOrder="" then
				pcv_BackOrder="0"
			end if
			pcv_NDays=pcArray(13,i)
			if IsNull(pcv_NDays) or pcv_NDays="" then
				pcv_NDays="0"
			end if
			pcv_Weight=pcArray(14,i)
			pcv_QtyToPound=pcArray(17,i)
			if NOT isNumeric(pcv_QtyToPound) or pcv_QtyToPound="" then
				pcv_QtyToPound="0"
			end if
			if (pcv_Weight="" OR IsNull(pcv_Weight) OR pcv_Weight="0") AND pcv_QtyToPound="0" then
				pcv_Weight=ParentWeight
			end if
			pcv_w1=0
			pcv_w2=0
			if scShipFromWeightUnit="KGS" then
				pcv_w1=Int(pcv_Weight/1000)
				pcv_w2=pcv_Weight-(pcv_w1*1000)
			else
				pcv_w1=Int(pcv_Weight/16)
				pcv_w2=pcv_Weight-(pcv_w1*16)
			end if
			pcv_PrdName=pcArray(15,i)
			pcv_subReward=pcArray(16,i)
			if pcv_subReward="" OR IsNull(pcv_subReward) OR pcv_subReward="0" then
				pcv_subReward=ParentRW
			end if
			%>
			sp.o[<%=pcv_SPCount-1%>]=new makeLevel();
			sp.o[<%=pcv_SPCount-1%>].IDProduct=<%=pcv_SPid%>;
			sp.o[<%=pcv_SPCount-1%>].sku="<%=pcv_SPsku%>";
			<%if pcv_SPimg<>"" then%>
				sp.o[<%=pcv_SPCount-1%>].img="<%=pcv_SPimg%>";
			<%else%>
				sp.o[<%=pcv_SPCount-1%>].img=GeneralImg;
			<%end if%>
			<%if pcv_SPlimg<>"" then%>
				sp.o[<%=pcv_SPCount-1%>].limg="<%=pcv_SPlimg%>";
			<%else%>
				sp.o[<%=pcv_SPCount-1%>].limg=DefLargeImg;
			<%end if%>
			<%if (pcv_SPlimg<>"") then
				if instr(strImgFiles,pcv_SPlimg & "***")=0 then
					strImgFiles=strImgFiles & pcv_SPlimg & "***"%>
					//splimg<%=pcv_SPCount-1%> = new Image();
					//splimg<%=pcv_SPCount-1%>.src = "<%=pcf_getImagePath("catalog",pcv_SPlimg)%>";
				<%end if
			end if%>
		
			sp.o[<%=pcv_SPCount-1%>].stock=<%=pcv_SPStock%>;
			sp.o[<%=pcv_SPCount-1%>].price="<%=scCurSign & money(pcv_SPPrice)%>";
			sp.o[<%=pcv_SPCount-1%>].wprice="<%=scCurSign & money(pcv_SPWPrice)%>";
			<%if APPshowVAT="1" then%>
			sp.o[<%=pcv_SPCount-1%>].vat="<%=scCurSign & money(pcf_RemoveVAT(pcv_SPPrice,pcv_SPid))%>";
			<%end if%>
			<%For j=1 to ubound(TempArr)%>
				sp.o[<%=pcv_SPCount-1%>].opts[<%=j-1%>]=<%=TempArr(j)%>;
			<%Next%>
			sp.o[<%=pcv_SPCount-1%>].inactive=<%=pcv_SPInactive%>;
			sp.o[<%=pcv_SPCount-1%>].addprice=<%=replace(pcv_AddPrice,",",".")%>;
			sp.o[<%=pcv_SPCount-1%>].addwprice=<%=replace(pcv_AddWPrice,",",".")%>;
			sp.o[<%=pcv_SPCount-1%>].lprice=<%=replace(pcv_ListPrice,",",".")%>;
			sp.o[<%=pcv_SPCount-1%>].nostock=<%=pcv_DisStock%>;
			sp.o[<%=pcv_SPCount-1%>].backorder=<%=pcv_BackOrder%>;
			sp.o[<%=pcv_SPCount-1%>].ndays=<%=pcv_NDays%>;
			<%if (session("customerType")="1") OR (session("customerCategory")<>0) then
				pcv_NewPrice=pcv_SPWPrice
			else
				pcv_NewPrice=pcv_SPPrice
			end if
			pcv_NewAddPrice=cdbl(pcv_NewPrice)-(cdbl(pq_Price)-cdbl(pq_AddPrice))
			%>
			sp.o[<%=pcv_SPCount-1%>].retext="<%=pcv_SPid%>_<%=replace(pcv_NewAddPrice,",",".")%>_<%=pcv_Weight%>_<%=replace(pcv_NewPrice,",",".")%>";
			<%pcv_PrdName1=replace(pcv_PrdName,"""","\""")%>
			sp.o[<%=pcv_SPCount-1%>].prdname="<%=replace(pcv_PrdName1,"&quot;","\""")%>";
			sp.o[<%=pcv_SPCount-1%>].weight1=<%=replace(pcv_w1,",",".")%>;
			sp.o[<%=pcv_SPCount-1%>].weight2=<%=replace(pcv_w2,",",".")%>;
			sp.o[<%=pcv_SPCount-1%>].reward=<%=replace(pcv_subReward,",",".")%>;
			
			<%'Backed-Up Price
			if HaveSale=1 then
				query="SELECT pcSB_Price FROM pcSales_BackUp WHERE idProduct=" & pcv_SPid & " AND pcSC_ID=" & pcSCID & ";"
				set rsQ=connTemp.execute(query)
				pcOrgPrice=0
				if not rsQ.eof then
					pcOrgPrice=rsQ("pcSB_Price")
				end if
				set rsQ=nothing%>
				sp.o[<%=pcv_SPCount-1%>].backprice="<%=scCurSign & money(pcOrgPrice)%>";
			<%else%>
				sp.o[<%=pcv_SPCount-1%>].backprice="0";
			<%end if%>
			
			<%if pcv_SPCount mod 100=0 then%>
			</script>
			<script>
			<%end if%>
			<%if pcv_SPCount mod 50=0 then
			Response.Flush
			Response.Clear
			end if%>
		<%Next%>
	
		var SPCount=<%=pcv_SPCount%>;
		
	function PreSelect()
	{
		var j=0;
		objElems = opener.document.additem.elements;
		var m=-1;
		var IDProduct=0;
		for(j=0;j<objElems.length;j++)
		{
			if ((objElems[j].name=="<%=pq_IDField%>")  && (objElems[j].type!="radio") && (objElems[j].type!="checkbox"))
			{
					tmpA=objElems[j].value;
				  	tmpB=tmpA.split('_');
				  	IDProduct=eval(tmpB[0]);
				  	break;
			}

			if ((objElems[j].name=="<%=pq_IDField%>") && (objElems[j].type=="radio"))
			{
				m=m+1;
				if (m==<%=pq_VIndex%>)
				{
					tmpA=objElems[j].value;
				  	tmpB=tmpA.split('_');
				  	IDProduct=eval(tmpB[0]);
				  	break;
				}
			}
			if ((objElems[j].name=="<%=pq_IDField%>") && (objElems[j].type=="checkbox"))
			{
				tmpA=objElems[j].value;
			  	tmpB=tmpA.split('_');
			  	IDProduct=eval(tmpB[0]);
			  	break;
			}
		}
		
		if (IDProduct!=0)
		{
			for (var k=0; k < SPCount; k++)
			{
				if (eval(sp.o[k].IDProduct)==eval(IDProduct))
				{
					<%if pcv_ApparelRadio="0" then%>
						var i=0;
						for (i=1;i<=GrpCount;i++) eval("document.additem.idOption" + i).value=sp.o[k].opts[i-1];
					<%else%>
						var i=0;
						var j=0;
						objElems1 = document.additem.elements;
						for (j=1;j<=GrpCount;j++)
						{
							for(i=0;i<objElems1.length;i++)
							{
								if ((objElems1[i].name=="idOption" + j) && (eval(objElems1[i].value)==eval(sp.o[k].opts[j-1])))
								{
									objElems1[i].checked=true;
									break;
								}
							}
						}
					<%end if%>
					break;
				}
			}
		} //IDProduct !=0
	}
	
	function AddBack()
	{
		var tmpStr1="<%=dictLanguage.Item(Session("language")&"_viewPrd_60")&dictLanguage.Item(Session("language")&"_sds_viewprd_1")%>";
		var tmpStr2="<%=dictLanguage.Item(Session("language")&"_sds_viewprd_1b")%>";
		if (MyAccept==1)
		{
			var tmpid=document.additem.idproduct.value;
			
			for (var k=0; k <SPCount; k++)
			{
				if (eval(sp.o[k].IDProduct)==eval(tmpid))
				{
					var j=0;
					objElems = opener.document.additem.elements;
					var m=-1;
					for(j=0;j<objElems.length;j++)
					{
					
						if ((objElems[j].name=="<%=pq_IDField%>") && (objElems[j].type=="select-one"))
						{
							var oSelect=objElems[j];
							var i=0;
							for (i=0;i<oSelect.options.length;i++)
							{
								if (i==<%=pq_VIndex%>)
								{
									oSelect.options[i].value=sp.o[k].retext + "_<%=pIdProduct%>";
			  						oSelect.options[i].text=sp.o[k].prdname;
			  						oSelect.value=sp.o[k].retext + "_<%=pIdProduct%>";
									<%If (scOutofStockPurchase=-1) then
										pq_AVField=replace(pq_IDField,"CAG","AV")
										pq_OnlyIDField=replace(pq_IDField,"CAG","")%>
										if ((sp.o[k].stock<1) && (sp.o[k].nostock==0) && (sp.o[k].backorder==1) && (sp.o[k].ndays>0))
										{
											if (ie) var tmpitem=eval("opener.document.additem.<%=pq_AVField%>")
											else if (ns6) var tmpitem=opener.document.getElementById("<%=pq_AVField%>");
											if (tmpitem!=null) tmpitem.innerHTML=tmpStr1 + sp.o[k].ndays + tmpStr2;
											opener.availArr<%=pq_OnlyIDField%>[i]=tmpStr1 + sp.o[k].ndays + tmpStr2;
										}
										else
										{
											if (ie) var tmpitem=eval("opener.document.additem.<%=pq_AVField%>")
											else if (ns6) var tmpitem=opener.document.getElementById("<%=pq_AVField%>");
											if (tmpitem!=null) tmpitem.innerHTML="";
											opener.availArr<%=pq_OnlyIDField%>[i]="";
										}
										<%end if%>
				  						break;
			  					}
			  				}
			  				break;
						}
						
						
						if ((objElems[j].name=="<%=pq_IDField%>") && (objElems[j].type=="radio"))
						{
							m=m+1;
							if (m==<%=pq_VIndex%>)
							{
								opener.$("[name=<%=pq_IDField%>DESC<%=pq_VIndex%>]").html(sp.o[k].prdname);
								var strtemp=sp.o[k].prdname;
								opener.$("[name=<%=pq_IDField%>DESC<%=pq_VIndex%>]").size=strtemp.length;
								objElems[j].value=sp.o[k].retext + "_<%=pIdProduct%>";
								<%If (scOutofStockPurchase=-1) then
									pq_AVField=replace(pq_IDField,"CAG","AV")%>
									if ((sp.o[k].stock<1) && (sp.o[k].nostock==0) && (sp.o[k].backorder==1) && (sp.o[k].ndays>0))
									{
										if (ie) var tmpitem=eval("opener.document.additem.<%=pq_AVField%>P<%=pIdProduct%>")
										else if (ns6) var tmpitem=opener.document.getElementById("<%=pq_AVField%>P<%=pIdProduct%>");
										if (tmpitem!=null) tmpitem.innerHTML=tmpStr1 + sp.o[k].ndays + tmpStr2;
									}
									else
									{
										if (ie) var tmpitem=eval("opener.document.additem.<%=pq_AVField%>P<%=pIdProduct%>")
										else if (ns6) var tmpitem=opener.document.getElementById("<%=pq_AVField%>P<%=pIdProduct%>");
										if (tmpitem!=null) tmpitem.innerHTML="";
									}
								<%end if%>
								break;
							}
						}
						if ((objElems[j].name=="<%=pq_IDField%>") && (objElems[j].type=="checkbox"))
						{
							opener.$("[name=<%=pq_IDField%>DESC<%=pq_VIndex%>]").html(sp.o[k].prdname);
							var strtemp=sp.o[k].prdname;
							opener.$("[name=<%=pq_IDField%>DESC<%=pq_VIndex%>]").size=strtemp.length;
							objElems[j].value=sp.o[k].retext + "_<%=pIdProduct%>";
							<%if session("customerType")="1" then%>
                            opener.$("[name=<%=pq_IDField%>TX0]").value=sp.o[k].wprice;
							var strtemp=sp.o[k].wprice;
							<%else%>
                            opener.$("[name=<%=pq_IDField%>TX0]").value=sp.o[k].price;
							var strtemp=sp.o[k].price;
							<%end if%>
							opener.document.additem.<%=pq_IDField%>TX0.size=strtemp.length;
							<%If (scOutofStockPurchase=-1) then
								if pq_IDField<>"" then
								if len(pq_IDField)-len(CStr(pIdProduct))>0 then
									pq_AVField=replace(mid(pq_IDField,1,len(pq_IDField)-len(CStr(pIdProduct))),"CAG","AV")
								else
									pq_AVField=replace(pq_IDField,"CAG","AV")
								end if
								end if%>
								if ((sp.o[k].stock<1) && (sp.o[k].nostock==0) && (sp.o[k].backorder==1) && (sp.o[k].ndays>0))
								{
									if (ie) var tmpitem=eval("opener.document.additem.<%=pq_AVField%>P<%=pIdProduct%>")
									else if (ns6) var tmpitem=opener.document.getElementById("<%=pq_AVField%>P<%=pIdProduct%>");
									if (tmpitem!=null) tmpitem.innerHTML=tmpStr1 + sp.o[k].ndays + tmpStr2;
								}
								else
								{
									if (ie) var tmpitem=eval("opener.document.additem.<%=pq_AVField%>P<%=pIdProduct%>")
									else if (ns6) var tmpitem=opener.document.getElementById("<%=pq_AVField%>P<%=pIdProduct%>");
									if (tmpitem!=null) tmpitem.innerHTML="";
								}
							<%end if%>
							break;
						}
					}

				}
			}
		}
	
	}
	
	function new_CheckPrdActiveAvailablePrice(tmpOpt,pos)
	{
		var i=0;
		var j=0;
		var HaveInActive=0;
		var HaveNotAvailablePrd=0;
		var tmpPrdPrice=0;
		var HaveOOS=0;
		var HaveSelectedOpt=0;
		var SubCount=0;
		
		for (i=0;i<=SPCount-1;i++)
		{
			var test1=1;
			var test2=1;
			var NeedMoreTest=0;
			var FindPrd=0;
			for (j=0;j<=GrpCount-1;j++)
			{
				if (j!=pos)
				{
					<%'Radio-Box Option
					IF pcv_ApparelRadio="1" THEN%>
					var tmpvalue=new_GetRadioValue(eval("document.additem.idOption" + parseInt(j+1)));
					<%ELSE%>
					var tmpvalue=eval("document.additem.idOption" + parseInt(j+1)).value;
					<%END IF%>
					
					if (tmpvalue + "" !="")
					{
						HaveSelectedOpt=1;
						if (tmpvalue + "" != sp.o[i].opts[j] + "")
						{
							test1=0;
							break;
						}
					}
				}
			}
			if (test1==1)
			{
				if (tmpOpt + "" !="")
				{
					if (tmpOpt + "" != sp.o[i].opts[pos] + "")
					{
						test1=0;
					}
				}
				else
				{
					test1=0;
				}
			}
			if ((test2==1) && (test1==1)) // && (sp.o[i].opts[GrpCount-1] + "" == tmpOpt + "")
			{
				SubCount=SubCount+1;
				if ((sp.o[i].stock>0) || (sp.o[i].nostock!=0) || (sp.o[i].backorder>0))
				{
					<%if session("CustomerType")=1 then%>
						<%if app_DisplayFinalPrice=1 then%>
							tmpPrdPrice=sp.o[i].wprice;
						<%else%>
							tmpPrdPrice=sp.o[i].addwprice;
						<%end if%>
					<%else%>
						<%if app_DisplayFinalPrice=1 then%>
							tmpPrdPrice=sp.o[i].price;
						<%else%>
							tmpPrdPrice=sp.o[i].addprice;
						<%end if%>
					<%end if%>
					HaveOOS=0;
					FindPrd=1;
				}
				else
				{
					//if (pos!=GrpCount-1)
					HaveOOS=1;
					HaveNotAvailablePrd=0;
					//if ((i<SPCount-1) && (HaveSelectedOpt==1))
					NeedMoreTest=1;
				}
			}
			if (NeedMoreTest==0)
			{
			if (test1==1)
			{
				if (sp.o[i].inactive==0)
				{
					HaveInActive=0;
					HaveNotAvailablePrd=0;
					break;
				}
				else
				{
					HaveInActive=1;
					FindPrd=0;
					NeedMoreTest=1;
				}
			}
			if (NeedMoreTest==0)
			{
			if (test1==1)
			{
				HaveNotAvailablePrd=0;
				break;
			}
			else
			{
				if (SubCount==0)
				{
					HaveNotAvailablePrd=1;
					NeedMoreTest=1;
				}
			}
			}
			if (FindPrd==1) break;
			}
			
		}
		subprd_InActive=HaveInActive;
		subprd_NotAvailablePrd=HaveNotAvailablePrd;
		subprd_OOS=HaveOOS;
		subprd_PrdPrice=tmpPrdPrice;
	}
	
	<%if app_DisplayWaitingBox=1 then%>
	var start=1;

	function gosleep(tmpid,ctype)
	{
		grayOut(true, {'opacity':'25'});
		waitBoxobj.style.top=document.documentElement.scrollTop+(winheight-waitBoxobj.offsetHeight)/2+"px";
		waitBoxobj.style.left=document.documentElement.scrollLeft+(winwidth- waitBoxobj.offsetWidth)/2 + "px";
		waitBoxobj.style.visibility="visible";
		var t=setTimeout("new_CheckOptGroup(" + tmpid + "," + ctype + ");",0);
	}
	<%end if%>
	
	function chooseSubPrdImg(tmpOpt)
		{
		var i=0;
		var ctype=0;
			
			for (i=0;i<=SPCount-1;i++)
			{
				var test1=0;
				if (tmpOpt + "" == sp.o[i].opts[0] + "")
				{
					test1=1;
				}

				if ((test1==1) && (sp.o[i].img != "") && (sp.o[i].img != "no_image.gif"))
				{
					if (sp.o[i].limg=="")
					{
                        $pc("#mainimg").attr('src', '<%=pcv_tmpNewPath%>catalog/' + sp.o[i].img ); 
						if (ie) show_10.style.display="none"
						else if (ns6) document.getElementById("show_10").style.display="none";
					}
					else
					{
						if (ie) show_10.style.display=""
						else if (ns6) document.getElementById("show_10").style.display="";
						LargeImg=sp.o[i].limg;
                        setMainImg("<%=pcv_tmpNewPath%>catalog/" + sp.o[i].img, "<%=pcv_tmpNewPath%>catalog/" + LargeImg);
					}
					ctype=1;
					break;
				}
			}

			//linkBack();
				
		}
		
<%'Radio-Box Option
IF pcv_ApparelRadio="1" THEN%>
	function new_GetRadioValue(tmpList)
	{
		var i=0;
		var j=tmpList.length;
		for(var i = 0; i < j; i++)
		{
			if(tmpList[i].checked)
			{
				return(tmpList[i].value);
			}
		}
		return("");
	}
	
	function new_SynChecked(tmpList)
	{
		var i=0;
		var j=tmpList.length;
		for(var i = 0; i < j; i++)
		{
			if (!tmpList[i].checked) $(tmpList[i]).removeAttr('is_che');
		}
	}
	
	function new_EnableRadioValue(tmpList,tmpvalue)
	{
		var i=0;
		var j=tmpList.length;
		for(var i = 0; i < j; i++)
		{
			if(parseInt(tmpList[i].value)==parseInt(tmpvalue))
			{
				tmpList[i].disabled=false;
				if (app_HideItems==1)
				{
					if (ie) var tmpitem=eval("Opt_"+tmpList[i].value+"_TABLE")
					else if (ns6) var tmpitem=document.getElementById("Opt_"+tmpList[i].value+"_TABLE");
					if (tmpitem!=null) { tmpitem.style.display=""; }
				}
				break;
			}
		}
	}
	
	function new_DisableRadioValue(tmpList,tmpvalue)
	{
		var i=0;
		var j=tmpList.length;
		for(var i = 0; i < j; i++)
		{
			if(parseInt(tmpList[i].value)==parseInt(tmpvalue))
			{
				tmpList[i].disabled=true;
				tmpList[i].checked=false;
				if (app_HideItems==1)
				{
					if (ie) var tmpitem=eval("Opt_"+tmpList[i].value+"_TABLE")
					else if (ns6) var tmpitem=document.getElementById("Opt_"+tmpList[i].value+"_TABLE");
					if (tmpitem!=null) { tmpitem.style.display="none"; }
				}
				break;
			}
		}
	}
	
	function new_SetRadioValue(tmpList,tmpvalue)
	{
		var i=0;
		var j=tmpList.length;
		for(var i = 0; i < j; i++)
		{
			if ((tmpList[i].value==tmpvalue) && (tmpList[i].disabled==false))
			{
				if (tmpvalue!="")
				{
					if (ie) var tmpitem=eval("Opt_"+tmpList[i].value+"_TABLE")
					else if (ns6) var tmpitem=document.getElementById("Opt_"+tmpList[i].value+"_TABLE");
					if (tmpitem!=null) { if (tmpitem.style.display=="none") break;
					tmpList[i].checked=true; }
					return(true);
				}
			}
		}
		//try
		//{
			if (tmpList.disabled==false)
			{
				tmpList.checked=true;
			}
			if (tmpList[0].disabled==false)
			{
				tmpList[0].checked=true;
			}
		//}
		//catch(er)
		//{
		//}
	}
	
	function new_HideRadioList(tmpList)
	{
		var i=0;
		var j=tmpList.length;
		for(var i = 0; i < j; i++)
		{
			if(tmpList[i].value!="")
			{
				if (ie) var tmpitem=eval("Opt_"+tmpList[i].value+"_TABLE")
				else if (ns6) var tmpitem=document.getElementById("Opt_"+tmpList[i].value+"_TABLE");
				if (tmpitem!=null) { tmpitem.style.display="none"; }
			}
		}
	}
	
	
	function new_clearRadioList(tmpid,nosub)
	{
		var SelectA=eval("document.additem.idOption" + tmpid);
		var savevalue="";
		savevalue=new_GetRadioValue(SelectA);
		
		<%if pcv_ShowStockMsg<>"1" then%>
			eval("document.additem.idOption" + tmpid + "_0_TXT").value="<%=pcv_StockMsg%>";
			//try
			//{
				SelectA[0].checked=true;
				SelectA.checked=true;
			//}
			//catch(er)
			//{
			//}
			new_HideRadioList(SelectA);
		<%else%>
			new_HideRadioList(SelectA);
			eval("document.additem.idOption" + tmpid + "_0_TXT").value="<%=dictLanguage.Item(Session("language")&"_viewPrd_61")%>";
			var idradio=tmpid-1;
			var radiocount=optGrp.count[idradio]-1;
			for(i=0;i<=radiocount;i++)
			{
				var AddP1="";
				var AddP="";
				var tmpMsg="";
				var tmpPrice="";
				subprd_InActive=0;
				subprd_NotAvailablePrd=0;
				subprd_PrdPrice=0;
				subprd_OOS=0;
				new_CheckPrdActiveAvailablePrice(optGrp.grp[idradio].opt[i],idradio);
				tmpPrice=subprd_PrdPrice;
				if (nosub==0)
				{
					<%if app_DisplayFinalPrice=1 then%>
					if (tmpPrice==DefaultPrice)
					{
						tmpPrice=0;
					}
					<%end if%>
					if (tmpPrice!=0)
					{
						<%if app_DisplayFinalPrice=1 then%>
							if (tmpPrice!=DefaultPrice)
							{
								AddP1=" - "
							}
							AddP=tmpPrice;
						<%else%>
							var PriceAdd = new NumberFormat();
							PriceAdd.setNumber(tmpPrice);
							<%if scDecSign="," then%>
								PriceAdd.setSeparators(true,PriceAdd.PERIOD);
							<%else%>
								PriceAdd.setCommas(true);
							<%end if%>
							PriceAdd.setPlaces(2);
							PriceAdd.setCurrency(true);
							PriceAdd.setCurrencyPrefix("<%=scCurSign%>");
							AddP=PriceAdd.toFormatted();
						<%end if%>
						if (tmpPrice > 0)
						{
							<%if app_DisplayFinalPrice=1 then%>
								AddP1=" - "
							<%else%>
								AddP1=" - <% response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg3")) %>"
							<%end if%>
						}
						else
						{
							if (tmpPrice < 0)
							{
								<%if app_DisplayFinalPrice=1 then%>
									AddP1=" - "
								<%else%>
									AddP1=" - <% response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg4")) %>"
								<%end if%>
							}
						}
					}
				}

				tmpMsg=" (<%=pcv_StockMsg%>)";

				if (subprd_InActive==0)
				{
					//try
					//{
						if (ie) var tmpitem=eval("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE")
						else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE");
						if (tmpitem!=null) { tmpitem.style.display=""; }
					//}
					//catch(er)
					//{
					//}
					if (subprd_NotAvailablePrd==1)
					{
						//try
						//{
							if (ie) var tmpitem=eval("document.additem.Opt_"+optGrp.grp[idradio].opt[i]+"_TXT")
							else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TXT");
							if (tmpitem!=null)
							{
								tmpitem.value=optGrp.grp[idradio].name[i];
								var mStr=tmpitem.value;
								tmpitem.size=mStr.length;
							}
						//}
						//catch(e) {}
						new_DisableRadioValue(eval("document.additem.idOption" + tmpid),optGrp.grp[idradio].opt[i]);
					}
					else
					{
						if (ie) var tmpitem=eval("document.additem.Opt_"+optGrp.grp[idradio].opt[i]+"_TXT")
						else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TXT");
						if (tmpitem!=null) { tmpitem.value=optGrp.grp[idradio].name[i] + AddP1 + AddP + tmpMsg;
						var mStr=tmpitem.value;
						tmpitem.size=mStr.length+1; }
					}
				}
			}
			new_SetRadioValue(SelectA,savevalue);
		<%end if%>
	}
	
	function new_GenRadioList(tmpid,alist,nosub)
	{
		var i=0;
		var j=0;
		var tmp1=alist;
		var idradio=tmpid-1;
		var radiocount=optGrp.count[idradio]-1;
		var savevalue="";
		var SelectA=eval("document.additem.idOption" + tmpid);
		savevalue=new_GetRadioValue(SelectA);
		eval("document.additem.idOption" + tmpid + "_0_TXT").value="<%=dictLanguage.Item(Session("language")&"_viewPrd_61")%>";
		
		if ((GrpCount==1) || (tmp1==""))
		{
		var tmp1="||";
		for (i=0;i<=SPCount-1;i++)
		{
			if (((sp.o[i].stock>0) || (sp.o[i].nostock!=0) || (sp.o[i].backorder>0)) && (sp.o[i].inactive==0))
			{
				tmp1=tmp1 + "" + sp.o[i].opts[idradio] + "||";
			}
		}
		}
		
		for(i=0;i<=radiocount;i++)
		{
			var AddP1="";
			var AddP="";
			var tmpMsg="";
			var tmpPrice="";
			subprd_InActive=0;
			subprd_NotAvailablePrd=0;
			subprd_PrdPrice=0;
			subprd_OOS=0;
			new_CheckPrdActiveAvailablePrice(optGrp.grp[idradio].opt[i],idradio);
			tmpPrice=subprd_PrdPrice;
			if (nosub==0)
			{
				<%if app_DisplayFinalPrice=1 then%>
				if (tmpPrice==DefaultPrice)
				{
					tmpPrice=0;
				}
				<%end if%>
				if (tmpPrice!=0)
				{
					<%if app_DisplayFinalPrice=1 then%>
						if (tmpPrice!=DefaultPrice)
						{
							AddP1=" - "
						}
						AddP=tmpPrice;
					<%else%>
						var PriceAdd = new NumberFormat();
						PriceAdd.setNumber(tmpPrice);
						<%if scDecSign="," then%>
							PriceAdd.setSeparators(true,PriceAdd.PERIOD);
						<%else%>
							PriceAdd.setCommas(true);
						<%end if%>
						PriceAdd.setPlaces(2);
						PriceAdd.setCurrency(true);
						PriceAdd.setCurrencyPrefix("<%=scCurSign%>");
						AddP=PriceAdd.toFormatted();
					<%end if%>
					if (tmpPrice > 0)
					{
						<%if app_DisplayFinalPrice=1 then%>
							AddP1=" - "
						<%else%>
							AddP1=" - <% response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg3")) %>"
						<%end if%>
					}
					else
					{
						if (tmpPrice < 0)
						{
							<%if app_DisplayFinalPrice=1 then%>
								AddP1=" - "
							<%else%>
								AddP1=" - <% response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg4")) %>"
							<%end if%>
						}
					}
				}
			}
			var tmp2="||"+optGrp.grp[idradio].opt[i]+"||";
			if (tmp1.indexOf(tmp2)==-1)
			{
				if (idradio<GrpCount-1) {}
				else
				{
					tmpMsg=" (<%=pcv_StockMsg%>)";
				}
			}
			<%if pcv_ShowStockMsg<>"1" then%>
			if ((tmpMsg=="") || (subprd_NotAvailablePrd==1) || (subprd_OOS==1))
			{
			<%else%>
			if (subprd_InActive==0)
			{
			<%end if%>
			
				if (ie) var tmpitem=eval("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE")
				else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE");
				if (tmpitem!=null) { tmpitem.style.display="none"; }
				
				<%if (pcv_ShowStockMsg<>"1") then%>
				if ((subprd_NotAvailablePrd==1) || (subprd_OOS==1))
				<%else%>
				if (subprd_NotAvailablePrd==1)
				<%end if%>
				{
					if ((app_HideItems==1) || (subprd_NotAvailablePrd==1)) {}
					else
					{
						if (ie) var tmpitem=eval("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE")
						else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE");
						if (tmpitem!=null) { tmpitem.style.display=""; }
						if (ie) var tmpitem=eval("document.additem.Opt_"+optGrp.grp[idradio].opt[i]+"_TXT")
						else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TXT");
						if (tmpitem!=null)
						{
							tmpitem.value=optGrp.grp[idradio].name[i];
							var mStr=tmpitem.value;
							tmpitem.size=mStr.length;
						}
					//}
					//catch(e) {}
					new_DisableRadioValue(eval("document.additem.idOption" + tmpid),optGrp.grp[idradio].opt[i]);
					}
				}
				else
				{
					if (ie) var tmpitem=eval("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE")
					else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE");
					if (tmpitem!=null) { tmpitem.style.display=""; }
					if (ie) var tmpitem=eval("document.additem.Opt_"+optGrp.grp[idradio].opt[i]+"_TXT")
					else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TXT");
					if (tmpitem!=null)
					{
						if (idradio<GrpCount-1)
						{
							tmpitem.value=optGrp.grp[idradio].name[i]
						}
						else
						{
							tmpitem.value=optGrp.grp[idradio].name[i] + AddP1 + AddP + tmpMsg;
						}
						var mStr=tmpitem.value;
						tmpitem.size=mStr.length+1;
					}
					new_EnableRadioValue(eval("document.additem.idOption" + tmpid),optGrp.grp[idradio].opt[i]);
				}
			}
			else
			{
					if (ie) var tmpitem=eval("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE")
					else if (ns6) var tmpitem=document.getElementById("Opt_"+optGrp.grp[idradio].opt[i]+"_TABLE");
					if (tmpitem!=null) { tmpitem.style.display="none"; }
			}
		}
		new_SetRadioValue(SelectA,savevalue);
	}
	
	function new_CheckOptGroup(tmpid,ctype)
	{
		<%if app_DisplayWaitingBox=1 then%>
		if (start==1)
		{
			start=0;
			gosleep(tmpid,ctype);
			return;
		}
		else
		{
			start=1;
		}
		<%end if%>
		<%	
		'// If we are in the admin we dont need the additional images javascripts.
		If pcv_strAdminPrefix<>"1" Then
		%>
		//linkBack();
		<% End If %>
		var tmpArr=new Array();
		var tmp1="||";
		var i=0;
		var nosub=0;
		LowStock=0;
		var InputQty=document.additem.quantity.value;
		if (InputQty=="") InputQty=0;
		var SelectA=eval("document.additem.idOption" + tmpid);
		new_SynChecked(SelectA);
		var grpvalue=new_GetRadioValue(SelectA);
		if (grpvalue=="") {MyAccept=0;}
		//else
		//{
			if (GrpCount-1==0)
			{
				nosub=0;
			}
			else
			{
				for (i=0;i<=GrpCount-2;i++)
				{
					if (new_GetRadioValue(eval("document.additem.idOption" + parseInt(i+1)))=="")
					{
						nosub=1;
						break;
					}
				}
			}
			
			for (i=0;i<=SPCount-1;i++)
			{
				var test1=1;
				for (j=0;j<=GrpCount-2;j++)
				{
					var tmpvalue=new_GetRadioValue(eval("document.additem.idOption" + parseInt(j+1)));
					if (tmpvalue != "")
					{
						if (tmpvalue + "" != sp.o[i].opts[j] + "")
						{
							test1=0;
							break;
						}
					}
				}
				if (test1==1)
				{
					if (((sp.o[i].stock>0) || (sp.o[i].nostock!=0) ||	(sp.o[i].backorder>0)) && (sp.o[i].inactive==0))
					{
						tmp1=tmp1 + "" + sp.o[i].opts[GrpCount-1] + "||";
					}
				}
			}
			
			if ((tmp1=="||") && ((tmpid!=GrpCount) || (GrpCount==1)))
			{
				new_clearRadioList(GrpCount,nosub);
				for (k=1;k<=GrpCount-1;k++) {if (k!=tmpid) new_GenRadioList(k,tmp1,nosub);}
				if (grpvalue!="") alert("<%response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg9"))%>");
				MyAccept=0;
			}
			else
			{
				for (k=1;k<=GrpCount-1;k++) {if (k!=tmpid) new_GenRadioList(k,tmp1,nosub);}
				if ((tmpid!=GrpCount) || (GrpCount==1)) new_GenRadioList(GrpCount,tmp1,nosub);
			}
			
			//Find selected Sub-Product
			MyAccept=0;
			
			if (nosub==0)
			{
				for (i=0;i<=SPCount-1;i++)
				{
					var test1=1;
					for (j=0;j<=GrpCount-1;j++)
					{
						var tmpvalue=new_GetRadioValue(eval("document.additem.idOption" + parseInt(j+1)));
						if (tmpvalue + "" != sp.o[i].opts[j] + "")
						{
							test1=0;
							break;
						}
					}
					if (test1==1)
					{
						if (((sp.o[i].stock>0) || (sp.o[i].nostock!=0) || (sp.o[i].backorder>0)) && (sp.o[i].inactive==0))
						{
							//Have selected Sub-Product
							MyAccept=1;
							LowStock=0;
							if ((sp.o[i].nostock!=0) ||	(sp.o[i].backorder>0))
							{
								LowPrdStock=9999999;
							}
							else
							{
								LowPrdStock=sp.o[i].stock;
							}
							LowPrdName=sp.o[i].prdname;
							if (checkNull('sku')) { $pc("#sku").html(sp.o[i].sku); }
							if (checkNull('appw1')) { document.getElementById("appw1").innerHTML=sp.o[i].weight1; }
							if (checkNull('appw2')) { document.getElementById("appw2").innerHTML=sp.o[i].weight2; }
							document.additem.idproduct.value=sp.o[i].IDProduct;
							if (Number(sp.o[i].lprice)>0)
							{
								var NewLPrice=sp.o[i].lprice;
							}
							else
							{
								var NewLPrice=Math.round((Number(DefaultLPrice)+Number(sp.o[i].addprice))*100)/100;
							}
							var PriceNum=sp.o[i].price;
							PriceNum=RmvComma(PriceNum);
							var NewSavings=Math.round((NewLPrice-Number(PriceNum))*100)/100;
							var NewSavingsP=Math.round(((NewLPrice-Number(PriceNum))/NewLPrice)*100);
							//try
							//{
							if (checkNull('mainprice')) { document.getElementById('mainprice').innerHTML=sp.o[i].price; }
							if (checkNull('lprice')) { document.getElementById("lprice").innerHTML=New_FormatNumber(NewLPrice) }
							if (checkNull('psavings')) { document.getElementById("psavings").innerHTML=New_FormatNumber(NewSavings) }
							if (checkNull('savingspercent')) { document.getElementById("savingspercent").innerHTML=" (" + RepComma(NewSavingsP) + "%)" }
							if (checkNull('pReward')) { document.getElementById("pReward").innerHTML=sp.o[i].reward }
							<%if HaveSale=1 then%>
							if (checkNull('backprice')) { document.getElementById('backprice').innerHTML=sp.o[i].backprice; }
							<%end if%>
							<%if APPshowVAT=1 then%>
							if (checkNull('vatspace')) { document.getElementById('vatspace').innerHTML=sp.o[i].vat; }
							<%end if%>
							//}
							//catch(err) {};
							<%if session("CustomerType")=1 or session("customerCategory")<>0 then%>
								if (checkNull('wprice')) { document.getElementById("wprice").innerHTML=sp.o[i].wprice; }
							<%end if%>
							
							SelectedSP=sp.o[i].IDProduct;
							<%if (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
							document.getElementById("AddtoList").style.display='';
							<%end if%>
							
							if (((sp.o[i].stock>0) && (sp.o[i].stock><%=pcv_ReorderLevel%>)) || ((sp.o[i].stock==0) && (sp.o[i].nostock!=0)))
							{
								<%if (pFormQuantity<>"-1" or NotForSaleOverride(session("customerCategory"))=1) then%>
								if ((sp.o[i].stock>0) && (sp.o[i].nostock==0))
								{
                                    <%if (scDisplayStock=-1) AND (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
									document.getElementById("StockMsg_TABLE").style.display='';
									<%end if%>
									var tmpfi=document.getElementById("StockMsg");
									tmpfi.value=sp.o[i].stock + "<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg5a")%>" + "<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg5")%>";
									var tmpstr=tmpfi.value;
									tmpfi.size=tmpstr.length;
								}
								else
								{
									document.getElementById("StockMsg_TABLE").style.display='none';
								}
								<%end if%>
							}
							else
							{
								if ((sp.o[i].stock>0)<%if pcv_ReorderLevel>0 then%> && (sp.o[i].stock<=<%=pcv_ReorderLevel%>)<%end if%>)
								{
                                    <%if (scDisplayStock=-1) AND (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
									document.getElementById("StockMsg_TABLE").style.display='';
									var tmpfi=document.getElementById("StockMsg");
									tmpfi.value="<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg6")%>";
									var tmpstr=tmpfi.value;
									tmpfi.size=tmpstr.length;
									<%end if%>
								}
								else
								{
									if ((sp.o[i].backorder==1) && (sp.o[i].ndays>0))
									{
                                        <%if (scDisplayStock=-1) AND (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
										document.getElementById("StockMsg_TABLE").style.display='';
										var tmpfi=document.getElementById("StockMsg");
										tmpfi.value="<%=dictLanguage.Item(Session("language")&"_sds_viewprd_1")%>" + sp.o[i].ndays + "<%=dictLanguage.Item(Session("language")&"_sds_viewprd_1b")%>";
										var tmpstr=tmpfi.value;
										tmpfi.size=tmpstr.length;
										<%end if%>
									}
									else
									if (sp.o[i].stock==0)
									{
                                        <%if (scDisplayStock=-1) AND (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
										document.getElementById("StockMsg_TABLE").style.display='';
										var tmpfi=document.getElementById("StockMsg");
										tmpfi.value="<%=pcv_StockMsg%>";
										var tmpstr=tmpfi.value;
										tmpfi.size=tmpstr.length;
										<%end if%>
									}
								}
							}
							
							if (ctype==0)
							{
								if (sp.o[i].limg=="")
								{
                                    $pc("#mainimg").attr('src', '<%=pcv_tmpNewPath%>catalog/' + sp.o[i].img ); 
									if (ie) show_10.style.display="none"
									else if (ns6) document.getElementById("show_10").style.display="none";
								}
								else
								{
									if (ie) show_10.style.display=""
									else if (ns6) document.getElementById("show_10").style.display="";
									LargeImg=sp.o[i].limg;
                                    setMainImg("<%=pcv_tmpNewPath%>catalog/" + sp.o[i].img, "<%=pcv_tmpNewPath%>catalog/" + LargeImg);
								}
							}
						}
					}
				}
			}
			
			//Don't have selected sub-product
			if (MyAccept==0)
			{
				if (checkNull('sku')) { $pc("#sku").html(DefaultSku); }
				if (checkNull('appw1')) { document.getElementById("appw1").innerHTML=DefaultWeight1; }
				if (checkNull('appw2')) { document.getElementById("appw2").innerHTML=DefaultWeight2; }
				document.additem.idproduct.value=DefaultIDPrd;
				//try
				//{
				if (checkNull('mainprice')) { document.getElementById('mainprice').innerHTML=DefaultPrice; }
				if (checkNull('lprice')) { document.getElementById("lprice").innerHTML=New_FormatNumber(DefaultLPrice) } 
				if (checkNull('psavings')) { document.getElementById("psavings").innerHTML=New_FormatNumber(DefaultSavings) } 
				if (checkNull('savingspercent')) { document.getElementById("savingspercent").innerHTML=" (" + RepComma(DefaultSavingsP) + "%)" } 
				if (checkNull('pReward')) { document.getElementById("pReward").innerHTML=DefaultReward }
				<%if HaveSale=1 then%>
					if (checkNull('backprice')) { document.getElementById('backprice').innerHTML=DefaultBackPrice; }
				<%end if%>
				<%if APPshowVAT="1" then%>
					if (checkNull('vatspace')) { document.getElementById('vatspace').innerHTML=DefaultVAT; }
				<%end if%>
				//}
				//catch(err) {};
				<%if session("CustomerType")=1 or session("customerCategory")<>0 then%>
					if (checkNull('wprice')) { document.getElementById("wprice").innerHTML=DefaultWPrice; }
				<%end if%>
				
				document.getElementById("StockMsg_TABLE").style.display='none';
				var tmpfi=document.getElementById("StockMsg");
				tmpfi.value="";
                try {
				    document.getElementById("AddtoList").style.display='none';
                } catch(err) { }
				
				var test1=0;
				for (i=0;i<=GrpCount-1;i++)
				{
					if (new_GetRadioValue(eval("document.additem.idOption" + parseInt(i+1)))!="")
					{
						test1=1;
						break;
					}
				}
				if (test1==0)
				{
					if (ctype==0)
					{
						if (DefLargeImg=="")
						{
                            $pc("#mainimg").attr('src', '<%=pcv_tmpNewPath%>catalog/' + GeneralImg );
							if (ie) show_10.style.display="none"
							else if (ns6) document.getElementById("show_10").style.display="none";
						}
						else
						{
							if (ie) show_10.style.display=""
							else if (ns6) document.getElementById("show_10").style.display="";
							LargeImg=DefLargeImg
                            setMainImg("<%=pcv_tmpNewPath%>catalog/" + GeneralImg, "<%=pcv_tmpNewPath%>catalog/" + LargeImg);
						}
					}
				}
				if (new_GetRadioValue(eval("document.additem.idOption1"))!="")
				{
					chooseSubPrdImg(new_GetRadioValue(eval("document.additem.idOption1")))
				}
			}
		//} //Have new option selected
		<%if popUpAPP=1 then%>
		AddBack();
		<%end if%>
		<%if app_DisplayWaitingBox=1 then%>
		waitBoxobj.style.visibility="hidden";
		grayOut(false, {'opacity':'25'});
		<%end if%>
		<%	
		'// If we are in the admin we dont need the additional images javascripts.
		If pcv_strAdminPrefix<>"1" Then
		%>
		//linkBack();
		clickSW=1;
		<% End If %>
	}
<%'Drop-down Option
ELSE%>

	function new_SetDropDownValue(tmpList,tmpvalue)
	{
		var i=0;
		var j=tmpList.options.length;
		for(var i = 0; i < j; i++)
		{
			if ((tmpList.options[i].value==tmpvalue) && (tmpList.options[i].style.color!="gray"))
			{
				tmpList.value=tmpvalue;
				return(true);
			}
		}
		tmpList.value=tmpList.options[0].value;
	}
	
	function new_clearDropDown(tmpid,nosub)
	{
		var SelectA=eval("document.additem.idOption" + tmpid);
		var savevalue="";
		savevalue=SelectA.value;
		SelectA.options.length = 0;
		<%if pcv_ShowStockMsg<>"1" then%>
			SelectA.options[0]=new Option("<%=pcv_StockMsg%>","");
			SelectA.value="";
		<%else%>
			var iddrop=tmpid-1;
			var dropcount=optGrp.count[iddrop]-1;
			SelectA.options[0]=new Option("<%=dictLanguage.Item(Session("language")&"_viewPrd_61")%>","");
			var count=0;
			for(i=0;i<=dropcount;i++)
			{
				var AddP1="";
				var AddP="";
				var tmpMsg="";
				var tmpPrice="";
				subprd_InActive=0;
				subprd_NotAvailablePrd=0;
				subprd_PrdPrice=0;
				subprd_OOS=0;
				new_CheckPrdActiveAvailablePrice(optGrp.grp[iddrop].opt[i],iddrop);
				tmpPrice=subprd_PrdPrice;
				if (nosub==0)
				{
					<%if app_DisplayFinalPrice=1 then%>
					if (tmpPrice==DefaultPrice)
					{
						tmpPrice=0;
					}
					<%end if%>
					if (tmpPrice!=0)
					{
						<%if app_DisplayFinalPrice=1 then%>
							if (tmpPrice!=DefaultPrice)
							{
								AddP1=" - "
							}
							AddP=tmpPrice;
						<%else%>
							var PriceAdd = new NumberFormat();
							PriceAdd.setNumber(tmpPrice);
							<%if scDecSign="," then%>
								PriceAdd.setSeparators(true,PriceAdd.PERIOD);
							<%else%>
								PriceAdd.setCommas(true);
							<%end if%>
							PriceAdd.setPlaces(2);
							PriceAdd.setCurrency(true);
							PriceAdd.setCurrencyPrefix("<%=scCurSign%>");
							AddP=PriceAdd.toFormatted();
						<%end if%>
						if (tmpPrice > 0)
						{
							<%if app_DisplayFinalPrice=1 then%>
								AddP1=" - "
							<%else%>
								AddP1=" - <% response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg3")) %>"
							<%end if%>
						}
						else
						{
							if (tmpPrice < 0)
							{
								<%if app_DisplayFinalPrice=1 then%>
									AddP1=" - "
								<%else%>
									AddP1=" - <% response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg4")) %>"
								<%end if%>
							}
						}
					}
				}

				tmpMsg=" (<%=pcv_StockMsg%>)";
				if (subprd_InActive==0)
				{
					count=count+1;
					if (subprd_NotAvailablePrd==1)
					{
						if (app_HideItems==1) {count=count-1;}
						else
						{
							SelectA.options[count]=new Option(optGrp.grp[iddrop].name[i],optGrp.grp[iddrop].opt[i]);
							SelectA.options[count].style.color="gray";
						}
					}
					else
					{
						SelectA.options[count]=new Option(optGrp.grp[iddrop].name[i] + AddP1 + AddP + tmpMsg,optGrp.grp[iddrop].opt[i])
					}
				}
			}
			new_SetDropDownValue(SelectA,savevalue);
		<%end if%>
	}
	
	function new_GenDropDown(tmpid,alist,nosub)
	{
		var i=0;
		var j=0;
		var tmp1=alist;
		var iddrop=tmpid-1;
		var dropcount=optGrp.count[iddrop]-1;
		var savevalue="";
		var SelectA=eval("document.additem.idOption" + tmpid);
		savevalue=SelectA.value;
		SelectA.options.length = 0;
		SelectA.options[0]=new Option("<%=dictLanguage.Item(Session("language")&"_viewPrd_61")%>","");
		
		if ((GrpCount==1) || (tmp1==""))
		{
		var tmp1="||";
		for (i=0;i<=SPCount-1;i++)
		{
			if (((sp.o[i].stock>0) || (sp.o[i].nostock!=0) || (sp.o[i].backorder>0)) && (sp.o[i].inactive==0))
			{
				tmp1=tmp1 + "" + sp.o[i].opts[iddrop] + "||";
			}
		}
		}
		
		var count=0;
		for(i=0;i<=dropcount;i++)
		{
			var AddP1="";
			var AddP="";
			var tmpMsg="";
			var tmpPrice="";
			subprd_InActive=0;
			subprd_NotAvailablePrd=0;
			subprd_PrdPrice=0;
			subprd_OOS=0;
			new_CheckPrdActiveAvailablePrice(optGrp.grp[iddrop].opt[i],iddrop);
			tmpPrice=subprd_PrdPrice;
			if (nosub==0)
			{
				<%if app_DisplayFinalPrice=1 then%>
				if (tmpPrice==DefaultPrice)
				{
					tmpPrice=0;
				}
				<%end if%>
				if (tmpPrice!=0)
				{
					<%if app_DisplayFinalPrice=1 then%>
					if (tmpPrice!=DefaultPrice)
					{
						AddP1=" - "
					}
					AddP=tmpPrice;
					<%else%>
					var PriceAdd = new NumberFormat();
					PriceAdd.setNumber(tmpPrice);
					<%if scDecSign="," then%>
						PriceAdd.setSeparators(true,PriceAdd.PERIOD);
					<%else%>
						PriceAdd.setCommas(true);
					<%end if%>
					PriceAdd.setPlaces(2);
					PriceAdd.setCurrency(true);
					PriceAdd.setCurrencyPrefix("<%=scCurSign%>");
					AddP=PriceAdd.toFormatted();
					<%end if%>
					if (tmpPrice > 0)
					{
						<%if app_DisplayFinalPrice=1 then%>
							AddP1=" - "
						<%else%>
							AddP1=" - <% response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg3")) %>"
						<%end if%>
					}
					else
					{
						if (tmpPrice < 0)
						{
							<%if app_DisplayFinalPrice=1 then%>
								AddP1=" - "
							<%else%>
								AddP1=" - <% response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg4")) %>"
							<%end if%>
						}
					}
				}
			}
			var tmp2="||"+optGrp.grp[iddrop].opt[i]+"||";
			if (tmp1.indexOf(tmp2)==-1)
			{
				if (iddrop<GrpCount-1) {}
				else
				{
					tmpMsg=" (<%=pcv_StockMsg%>)";
				}
			}
			<%if pcv_ShowStockMsg<>"1" then%>
			if ((tmpMsg=="") || (subprd_NotAvailablePrd==1) || (subprd_OOS==1))
			{
			<%else%>
			if (subprd_InActive==0)
			{
			<%end if%>
			count=count+1;
				<%if (pcv_ShowStockMsg<>"1") then%>
				if ((subprd_NotAvailablePrd==1) || (subprd_OOS==1))
				<%else%>
				if (subprd_NotAvailablePrd==1)
				<%end if%>
				{
					if ((app_HideItems==1) || (subprd_NotAvailablePrd==1)) {count=count-1;}
					else
					{
						SelectA.options[count]=new Option(optGrp.grp[iddrop].name[i],optGrp.grp[iddrop].opt[i]);
						SelectA.options[count].style.color="gray";
					}
				}
				else
				{
					if (iddrop<GrpCount-1)
					{
						SelectA.options[count]=new Option(optGrp.grp[iddrop].name[i],optGrp.grp[iddrop].opt[i]);
					}
					else
					{
						SelectA.options[count]=new Option(optGrp.grp[iddrop].name[i] + AddP1 + AddP + tmpMsg,optGrp.grp[iddrop].opt[i]);
					}
				}
			}
		}
		new_SetDropDownValue(SelectA,savevalue);
	}
	
	function new_CheckOptGroup(tmpid,ctype)
	{
		//try {
		
		<%if app_DisplayWaitingBox=1 then%>
		if (start==1)
		{
			start=0;
			gosleep(tmpid,ctype);
			return;
		}
		else
		{
			start=1;
		}
		<%end if%>

		var grpstyle=eval("document.additem.idOption" + tmpid + ".options[document.additem.idOption" + tmpid + ".selectedIndex]").style.color;
		if (grpstyle=="gray")
		{
			MyAccept=0;
			alert("<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg9")%>\n");
			eval("document.additem.idOption" + tmpid).value=eval("document.additem.idOption" + tmpid + ".options[0]").value;
			var grpvalue="";
		}
		else
		{
			var grpvalue=eval("document.additem.idOption" + tmpid).value;
		}
		var tmpArr=new Array();
		var tmp1="||";
		var i=0;
		var nosub=0; //Does not have enough options for a sub-product
		LowStock=0;
		var InputQty=document.additem.quantity.value;
		if (InputQty=="") InputQty=0;
		
		if (grpvalue=="") {MyAccept=0;}
		//else
		//{
			if (GrpCount-1==0)
			{
				nosub=0;
			}
			else
			{
				for (i=0;i<=GrpCount-2;i++)
				{
					if (eval("document.additem.idOption" + parseInt(i+1)).value=="")
					{
						nosub=1;
						break;
					}
				}
			}
			
			for (i=0;i<=SPCount-1;i++)
			{
				var test1=1;
				for (j=0;j<=GrpCount-2;j++)
				{
					var tmpvalue=eval("document.additem.idOption" + parseInt(j+1)).value;
					if (tmpvalue != "")
					{
						if (tmpvalue + "" != sp.o[i].opts[j] + "")
						{
							test1=0;
							break;
						}
					}
				}
				if (test1==1)
				{
					if (((sp.o[i].stock>0) || (sp.o[i].nostock!=0) ||	(sp.o[i].backorder>0)) && (sp.o[i].inactive==0))
					{
						tmp1=tmp1 + "" + sp.o[i].opts[GrpCount-1] + "||";
					}
				}
			}
						
			if ((tmp1=="||") && ((tmpid!=GrpCount) || (GrpCount==1)))
			{
				new_clearDropDown(GrpCount,nosub);
				for (k=1;k<=GrpCount-1;k++) {if (k!=tmpid) new_GenDropDown(k,tmp1,nosub);}
				if (grpvalue!="") alert("<%response.write(dictLanguage.Item(Session("language")&"_viewPrd_spmsg9"))%>");
				MyAccept=0;
			}
			else
			{
				for (k=1;k<=GrpCount-1;k++) {if (k!=tmpid) new_GenDropDown(k,tmp1,nosub);}
				if ((tmpid!=GrpCount) || (GrpCount==1))
				{
					new_GenDropDown(GrpCount,tmp1,nosub);
				}
			}
			
			//Find selected Sub-Product
			MyAccept=0;
			
			if (nosub==0)
			{
				for (i=0;i<=SPCount-1;i++)
				{
					var test1=1;
					for (j=0;j<=GrpCount-1;j++)
					{
                        var tmpvalue = $('select[name="' + 'idOption' + parseInt(j+1) + '"]').val();
						if (tmpvalue + "" != sp.o[i].opts[j] + "")
						{
							test1=0;
							break;
						}
					}
					if (test1==1)
					{
						if (((sp.o[i].stock>0) || (sp.o[i].nostock!=0) ||	(sp.o[i].backorder>0)) && (sp.o[i].inactive==0))
						{
							//Have selected Sub-Product
							MyAccept=1;
							LowStock=0;
							if ((sp.o[i].nostock!=0) ||	(sp.o[i].backorder>0))
							{
								LowPrdStock=9999999;
							}
							else
							{
								LowPrdStock=sp.o[i].stock;
							}
							LowPrdName=sp.o[i].prdname;					
								
							if (checkNull('sku')) { $pc("#sku").html(sp.o[i].sku); }
							if (checkNull('appw1')) { document.getElementById("appw1").innerHTML=sp.o[i].weight1; }
							if (checkNull('appw2')) { document.getElementById("appw2").innerHTML=sp.o[i].weight2; }
							document.additem.idproduct.value=sp.o[i].IDProduct;
							if (Number(sp.o[i].lprice)>0)
							{
								var NewLPrice=sp.o[i].lprice;
							}
							else
							{
								var NewLPrice=Math.round((Number(DefaultLPrice)+Number(sp.o[i].addprice))*100)/100;
							}
							var PriceNum=sp.o[i].price;
							PriceNum=RmvComma(PriceNum);
							var NewSavings=Math.round((NewLPrice-Number(PriceNum))*100)/100;
							var NewSavingsP=Math.round(((NewLPrice-Number(PriceNum))/NewLPrice)*100);
							//try
							//{
							if (checkNull('mainprice')) { document.getElementById('mainprice').innerHTML=sp.o[i].price; }
							if (checkNull('lprice')) { document.getElementById("lprice").innerHTML=New_FormatNumber(NewLPrice) }
							if (checkNull('psavings')) { document.getElementById("psavings").innerHTML=New_FormatNumber(NewSavings) }
							if (checkNull('savingspercent')) { document.getElementById("savingspercent").innerHTML=" (" + RepComma(NewSavingsP) + "%)" }
							if (checkNull('pReward')) { document.getElementById("pReward").innerHTML=sp.o[i].reward }
							<%if HaveSale=1 then%>
							if (checkNull('backprice')) { document.getElementById('backprice').innerHTML=sp.o[i].backprice; }
							<%end if%>
							<%if APPshowVAT=1 then%>
							if (checkNull('vatspace')) { document.getElementById('vatspace').innerHTML=sp.o[i].vat; }
							<%end if%>
							//}
							//catch(err) {};
							<%if session("CustomerType")=1 or session("customerCategory")<>0 then%>
								if (checkNull('wprice')) { document.getElementById("wprice").innerHTML=sp.o[i].wprice; }
							<%end if%>
							
							SelectedSP=sp.o[i].IDProduct;
							<%if (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
							document.getElementById("AddtoList").style.display='';
							<%end if%>
							
							if (((sp.o[i].stock>0) && (sp.o[i].stock><%=pcv_ReorderLevel%>)) || ((sp.o[i].stock==0) && (sp.o[i].nostock!=0)))
							{
								<%if (pFormQuantity<>"-1" or NotForSaleOverride(session("customerCategory"))=1) then%>
								if ((sp.o[i].stock>0) && (sp.o[i].nostock==0))
								{
                                    <%if (scDisplayStock=-1) AND (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
									document.getElementById("StockMsg_TABLE").style.display='';
									<%end if%>
									var tmpfi=document.getElementById("StockMsg");
									tmpfi.value=sp.o[i].stock + "<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg5a")%>" + "<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg5")%>";
									var tmpstr=tmpfi.value;
									tmpfi.size=tmpstr.length;
								}
								else
								{
									document.getElementById("StockMsg_TABLE").style.display='none';
								}
								<%end if%>
							}
							else
							{
								if ((sp.o[i].stock>0)<%if pcv_ReorderLevel>0 then%> && (sp.o[i].stock<=<%=pcv_ReorderLevel%>)<%end if%>)
								{
                                    <%if (scDisplayStock=-1) AND (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
									document.getElementById("StockMsg_TABLE").style.display='';
									var tmpfi=document.getElementById("StockMsg");
									tmpfi.value="<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg6")%>";
									var tmpstr=tmpfi.value;
									tmpfi.size=tmpstr.length;
									<%end if%>
								}
								else
								{
									if ((sp.o[i].backorder==1) && (sp.o[i].ndays>0))
									{
										<%if (scDisplayStock=-1) AND (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
										document.getElementById("StockMsg_TABLE").style.display='';
										var tmpfi=document.getElementById("StockMsg");
										tmpfi.value="<%=dictLanguage.Item(Session("language")&"_sds_viewprd_1")%>" + sp.o[i].ndays + "<%=dictLanguage.Item(Session("language")&"_sds_viewprd_1b")%>";
										var tmpstr=tmpfi.value;
										tmpfi.size=tmpstr.length;
										<%end if%>
									}
									else
									if (sp.o[i].stock==0)
									{
										<%if (scDisplayStock=-1) AND (pFormQuantity<>"-1" OR NotForSaleOverride(session("customerCategory"))=1) AND pcf_OutStockPurchaseAllow AND (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) then%>
										document.getElementById("StockMsg_TABLE").style.display='';
										var tmpfi=document.getElementById("StockMsg");
										tmpfi.value="<%=pcv_StockMsg%>";
										var tmpstr=tmpfi.value;
										tmpfi.size=tmpstr.length;
										<%end if%>
									}
								}
							}
							
							if (ctype==0)
							{                                
								if (sp.o[i].limg=="")
								{
                                    $pc("#mainimg").attr('src', '<%=pcv_tmpNewPath%>catalog/' + sp.o[i].img );
									if (ie) show_10.style.display="none"
									else if (ns6) document.getElementById("show_10").style.display="none";
								}
								else
								{
									if (ie) show_10.style.display=""
									else if (ns6) document.getElementById("show_10").style.display="";
									LargeImg=sp.o[i].limg;
                                    setMainImg("<%=pcv_tmpNewPath%>catalog/" + sp.o[i].img, "<%=pcv_tmpNewPath%>catalog/" + LargeImg);
								}
							}
						}
					}
				}
			}
			
			//Don't have selected sub-product
			if (MyAccept==0)
			{
				if (checkNull('sku')) { $pc("#sku").html(DefaultSku); }
				if (checkNull('appw1')) { document.getElementById("appw1").innerHTML=DefaultWeight1; }
				if (checkNull('appw2')) { document.getElementById("appw2").innerHTML=DefaultWeight2; }
				document.additem.idproduct.value=DefaultIDPrd;
				//try
				//{
				if (checkNull('mainprice')) { document.getElementById('mainprice').innerHTML=DefaultPrice; }
				if (checkNull('lprice')) { document.getElementById("lprice").innerHTML=New_FormatNumber(DefaultLPrice) }  
				if (checkNull('psavings')) { document.getElementById("psavings").innerHTML=New_FormatNumber(DefaultSavings) }  
				if (checkNull('savingspercent')) { document.getElementById("savingspercent").innerHTML=" (" + RepComma(DefaultSavingsP) + "%)" }  
				if (checkNull('pReward')) { document.getElementById("pReward").innerHTML=DefaultReward }
				<%if HaveSale=1 then%>
					if (checkNull('backprice')) { document.getElementById('backprice').innerHTML=DefaultBackPrice; }
				<%end if%>
				<%if APPshowVAT="1" then%>
					if (checkNull('vatspace')) { document.getElementById('vatspace').innerHTML=DefaultVAT; }
				<%end if%>
				//}
				//catch(err) {};
				<%if session("CustomerType")=1 or session("customerCategory")<>0 then%>
					if (checkNull('wprice')) { document.getElementById("wprice").innerHTML=DefaultWPrice; }
				<%end if%>
				
				document.getElementById("StockMsg_TABLE").style.display='none';
				var tmpfi=document.getElementById("StockMsg");
				tmpfi.value="";
                try {
				document.getElementById("AddtoList").style.display='none';
                } catch(err) { }
				
				var test1=0;
				for (i=0;i<=GrpCount-1;i++)
				{
					if (eval("document.additem.idOption" + parseInt(i+1)).value!="")
					{
						test1=1;
						break;
					}
				}
				if (test1==0)
				{
					if (ctype==0)
					{
						if (DefLargeImg=="")
						{
                            $pc("#mainimg").attr('src', '<%=pcv_tmpNewPath%>catalog/' + GeneralImg ); 
							if (ie) show_10.style.display="none"
							else if (ns6) document.getElementById("show_10").style.display="none";
						}
						else
						{
							if (ie) show_10.style.display=""
							else if (ns6) document.getElementById("show_10").style.display="";
							LargeImg=DefLargeImg
                            setMainImg("<%=pcv_tmpNewPath%>catalog/" + GeneralImg, "<%=pcv_tmpNewPath%>catalog/" + LargeImg);
						}
					}
				}
				if (eval("document.additem.idOption1").value!="")
				{
					chooseSubPrdImg(eval("document.additem.idOption1").value)
				}
			}
		//} //Have new option selected
		<%if popUpAPP=1 then%>
		AddBack();
		<%end if%>
		<%if app_DisplayWaitingBox=1 then%>
		waitBoxobj.style.visibility="hidden";
		grayOut(false, {'opacity':'25'});
		<%end if%>
		//} catch(e){}
		
	<%	
	'// If we are in the admin we dont need the additional images javascripts.
	If pcv_strAdminPrefix<>"1" Then
	%>
	//linkBack();
	clickSW=1;
	<% End If %>
	}
<%END IF
'End of Drop-Down Option%>

	function checkNull(element)
	{
	  if (document.getElementById(element)!=null) {
			return(true);
	  } else {
			return(false);
	  }
	}

	function new_AddSPtoList(tmpID)
	{
		if (LowPrdStock<document.additem.quantity.value)
		{
			LowStock=1;
		}
		else
		{
			LowStock=0;
		}
		if (LowStock==1)
		{
			alert("<%=dictLanguage.Item(Session("language")&"_instPrd_2")%>" + LowPrdName + "<%=dictLanguage.Item(Session("language")&"_instPrd_3")%>" + LowPrdStock + "<%=dictLanguage.Item(Session("language")&"_instPrd_4")%>");
		}
		else
		{
		var i=0;
		var j=0;
		var test1=0;
		for (i=0;i<=SPCount-1;i++)
		{
			if (sp.o[i].IDProduct==tmpID)
			{
				test1=0;
				if (SavedCount>0)
				{
					for (j=0;j<=SavedCount-1;j++)
					{
						if (SaveList[j] == tmpID)
						{
							var tmpvalue=document.additem.quantity.value;
							if ((tmpvalue==0) || (tmpvalue=="")) tmpvalue=1;
							SaveQtyList[j]=parseInt(SaveQtyList[j])+parseInt(tmpvalue);
							test1=1;
							break;
						}
					}
				}
				if (test1==0)
				{
					SavedCount=SavedCount+1;
					SaveList[SavedCount-1]=sp.o[i].IDProduct;
					SaveSKUList[SavedCount-1]=sp.o[i].sku;
					var tmpStr1=sp.o[i].prdname;
					var tmpStr2=tmpStr1.split("(")
					var tmpStr3=tmpStr2[1].split(")")
					SaveDescList[SavedCount-1]=tmpStr3[0];
					var tmpvalue=document.additem.quantity.value;
					if ((tmpvalue==0) || (tmpvalue=="")) tmpvalue=1;
					SaveQtyList[SavedCount-1]=tmpvalue;
					break;
				}
			}
		}
		
		if (SavedCount==0)
		{
			document.additem.SavedList.value="";
			document.additem.SavedQtyList.value="";
			new_HideSavedList();
		}
		else
		{
			var tmp1="";
			var tmp2="";
			for (i=0;i<=SavedCount-1;i++)
			{
				tmp1=tmp1 + "" + SaveList[i] + ",";
				tmp2=tmp2 + "" + SaveQtyList[i] + ",";
			}
			document.additem.SavedList.value=tmp1;
			document.additem.SavedQtyList.value=tmp2;
			new_ShowSavedList();
		}
		}
	}
	
	function new_DelSPtoList(tmpID)
	{
		var i=0;
		var tmpindex=-1;
		var tmpcount=SavedCount;
		for (i=0;i<=SavedCount-1;i++)
		{
			if (SaveList[i]==tmpID)
			{
				SavedCount=SavedCount-1;
				SaveList[i]="";
				SaveSKUList[i]="";
				SaveDescList[i]="";
				SaveQtyList[i]="";
				tmpindex=i+1;
				break;
			}
		}
		
		if (tmpindex>0)
		{
			for (i=tmpindex;i<=tmpcount-1;i++)
			{
				SaveList[i-1]=SaveList[i];
				SaveSKUList[i-1]=SaveSKUList[i];
				SaveDescList[i-1]=SaveDescList[i];
				SaveQtyList[i-1]=SaveQtyList[i];
			}
		}
		
		if (SavedCount==0)
		{
			document.additem.SavedList.value="";
			document.additem.SavedQtyList.value="";
			new_HideSavedList();
		}
		else
		{
			var tmp1="";
			var tmp2="";
			for (i=0;i<=SavedCount-1;i++)
			{
				tmp1=tmp1 + "" + SaveList[i] + ",";
				tmp2=tmp2 + "" + SaveQtyList[i] + ",";
			}
			document.additem.SavedList.value=tmp1;
			document.additem.SavedQtyList.value=tmp2;
			new_ShowSavedList();
		}
	}
	
	function new_HideSavedList()
	{
		document.getElementById("SelectedPrd_TABLE").innerText="";
		document.getElementById("SelectedPrd_TABLE").style.display="none";
	}
	
	function new_ShowSavedList()
	{
		var tmpHTML="";
		var i=0;
		tmpHTML='<div class="pcTable pcShowList">'
        tmpHTML=tmpHTML+'<div class="pcTableHeader">'
            tmpHTML=tmpHTML+'<div style="width: 60%"><%response.write dictLanguage.Item(Session("language")&"_viewPrd_spmsg1")%></div>'
            tmpHTML=tmpHTML+'<div style="width: 20%"><%response.write dictLanguage.Item(Session("language")&"_viewPrd_spmsg1a")%></div>'
            tmpHTML=tmpHTML+'<div style="width: 20%">&nbsp;</div>'
        tmpHTML=tmpHTML+'</div>';
		for (i=0;i<=SavedCount-1;i++)
		{
			tmpHTML=tmpHTML+'<div class="pcTableRow">'
            tmpHTML=tmpHTML+'<div style="width: 60%"><span class="pcSmallText">' + SaveDescList[i] + '</span></div>'
            tmpHTML=tmpHTML+'<div style="width: 20%"><span class="pcSmallText">' +  + SaveQtyList[i] + '</span></div>'
            tmpHTML=tmpHTML+'<div style="width: 20%"><a href="javascript:new_DelSPtoList(' + SaveList[i] + ');"><img src="<%=pcf_getImagePath(pcv_tmpNewPath & "images","delete2.gif")%>" class="pcSwatchImg" alt="<%response.write dictLanguage.Item(Session("language")&"_viewPrd_spmsg1b")%>"></a></div>'
            tmpHTML=tmpHTML+'</div>'
		}
        tmpHTML=tmpHTML+'</div>';
		document.getElementById("SelectedPrd_TABLE").innerHTML=tmpHTML;
		document.getElementById("SelectedPrd_TABLE").style.display="";
	}

	</script>
	<%else%>
		<script>
			function new_CheckOptGroup(tmpid,ctype)
			{
			}
		</script>
	<%end if
	set rs=nothing
END IF

End Sub

Public Sub ColorSwatches()
	Dim rs,query,rsA,OCount
	
	query="SELECT idOptionGroup FROM pcProductsOptions WHERE idproduct=" & pidProduct & " ORDER BY pcProdOpt_order ASC;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcv_tmpIDGrp=rs("idOptionGroup")
		set rs=nothing

		query = 		"SELECT options_optionsGroups.idoptoptgrp, options.optiondescrip, options.pcOpt_Img "
		query = query & "FROM options_optionsGroups "
		query = query & "INNER JOIN options "
		query = query & "ON options_optionsGroups.idOption = options.idOption "
		query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcv_tmpIDGrp &" "
		query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
		query = query & "ORDER BY options_optionsGroups.sortOrder ASC, options.optiondescrip ASC;"
		set rs=conntemp.execute(query)
		if not rs.eof then%>
			<div id="ColorSwatchesArea"> 
			<%
			Ocount=0
			do while not rs.eof
					pcv_IDOpt=rs("idoptoptgrp")
					pcv_OptName=rs("optiondescrip")
					pcv_OptImg=rs("pcOpt_Img")
					if pcv_OptImg<>"" then
						query= "SELECT idproduct from Products where (pcprod_ParentPrd=" & pidProduct &") AND  ((pcprod_Relationship like '" & pidProduct & "_" & pcv_IDOpt & "_" & "%') OR (pcprod_Relationship like '" & pidProduct & "_" & pcv_IDOpt & "')) AND removed=0 AND pcProd_SPInActive=0"
						if (pcv_ShowStockMsg="2") OR ((pcv_ShowStockMsg="0") AND (app_HideNotAvailableItems="1")) then
							query=query & " AND ((stock>0) OR (noStock<>0) OR (pcProd_BackOrder<>0))"
						end if
						query=query & " ORDER BY idproduct ASC;"
						set rsA=connTemp.execute(query)
						if (not rsA.eof) or (pcv_ShowStockMsg=0) then
							mystockmsg=""
						end if
						if (pcv_ShowStockMsg="1") then
							query= "SELECT idproduct from Products where (pcprod_ParentPrd=" & pidProduct &") AND  ((pcprod_Relationship like '" & pidProduct & "_" & pcv_IDOpt & "_" & "%') OR (pcprod_Relationship like '" & pidProduct & "_" & pcv_IDOpt & "')) AND removed=0 AND pcProd_SPInActive=0"
							query=query & " AND ((stock>0) OR (noStock<>0) OR (pcProd_BackOrder<>0))"
							query=query & " ORDER BY idproduct ASC;"
							set rsB=connTemp.execute(query)
							if rsB.eof then
								mystockmsg="alert('"&pcv_StockMsg&"');"
							end if
						end if
						if (not rsA.eof) or ((pcv_ShowStockMsg=1) and (not rsA.eof)) then
							if OCount=0 then%>
								<div>
							<%end if
							Ocount=Ocount+1%>
							<a href="javascript:click_swatch('<%=pcv_IDOpt%>');<%=mystockmsg%>"><img src="<%=pcf_getImagePath(pcv_tmpNewPath & "catalog",pcv_OptImg)%>" alt="<%=pcv_OptName%>" class="pcSwatchImg"></a>
							<%
							if Ocount=5 then
								Ocount=0%>
								</div>
							<%end if
						end if
					end if
					set rsA=nothing
				rs.MoveNext
				loop%>
				<%if Ocount>0 then
				Ocount=0%>
				</div>
				<%end if%>
				</div>
			<%end if
			set rs=nothing
	END IF
	set rs=nothing%>
	<script>
		function click_swatch(tmpOpt)
		{
		var i=0;
		var ctype=0;
		clickSW=1;

			<%if pcv_ApparelRadio="1" then%>
				new_SetRadioValue(document.additem.idOption1,tmpOpt);
			<%else%>
				document.additem.idOption1.value=tmpOpt;
				if ((document.additem.idOption1.value + ""=="") || (document.additem.idOption1.value!=tmpOpt))
				{
					for(i=2;i<=GrpCount;i++)
					{
						eval("document.additem.idOption" + i).value="";
					}
					new_GenDropDown(1,"",0);
					document.additem.idOption1.value=tmpOpt;
				}
				if (document.additem.idOption1.value + ""=="") {
                    document.additem.idOption1.value="";
                }
			<%end if%>
			
			new_CheckOptGroup(1,ctype);

		}
	</script>
<%End Sub

Public Sub CreateStockMsgArea()
    %>
	<div class="pcApparelRegion">
        <% If (pcv_SizeInfo<>"") or (pcv_SizeImg<>"") or (pcv_SizeURL<>"") Then %>
            <div class="row">
		        <div class="col-xs-12">
                    <div class="pcSizeChartLink"><span class="pcSmallText"><a href="javascript:open_win('app-sizechart.asp?idproduct=<%=pIDProduct%>');"><%=pcv_SizeLink%></a></span></div>
                </div>
            </div>
        <% End If %>
	    <div class="row">
		    <div class="col-xs-6">
                <div id="StockMsg_TABLE" style="display:none">
                    <div class="pcTable pcShowList">
                        <div class="pcTableHeader"><%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg") %></div>
                        <div><span class="pcSmallText"><input type="text" name="StockMsg" id="StockMsg" value="" readonly class="transparentField"></span></div>
                    </div>
                </div>
                <div id="AddtoList" class="pcSaveChoiceLink" style="display:none">
                    <%if popUpAPP<>"1" then%><span class="pcSmallText"><a href="javascript:new_AddSPtoList(SelectedSP);"><%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg7")%></a></span><%end if%>
                </div>
		    </div>
		    <div class="col-xs-6">
                <div id="SelectedPrd_TABLE" style="display:none"></div>
		    </div>
	    </div>
	</div>
<%End Sub
%>