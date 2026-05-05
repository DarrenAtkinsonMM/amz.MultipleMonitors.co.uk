<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>

<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/SearchConstants.asp"-->
<!--#include file="prv_incFunctions.asp"-->
<%Dim iAddDefaultPrice,	iAddDefaultWPrice%>
<!--#include file="pcCheckPricingCats.asp"-->
<%
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "viewCategories.asp"

'*******************************
' Page Settings
'*******************************
Dim pcCategoryClass, pcCategoryHover, pcProductHover
pcCategoryClass 	= "pcShowCategory"
pcCategoryHover 	= "pcShowCategoryBgHover"
pcProductHover		= "pcShowProductBgHover"

'*******************************
' START: Check store on/off, start PC session, check affiliate ID
'*******************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*******************************
' END: Check store on/off, start PC session, check affiliate ID
'*******************************

dim pTempIntSubCategory
%>
<!--#include file="prv_getSettings.asp"-->
<%
pTempIntSubCategory=session("idCategoryRedirect")
if pTempIntSubCategory = "" then
	pTempIntSubCategory=getUserInput(request("idCategory"),10)
end if

pTempIntSubCategory=70

'// Validate Category ID
	if not validNum(pTempIntSubCategory) then
		pTempIntSubCategory=""
	end if
	if pTempIntSubCategory="" or pTempIntSubCategory="0" then
		pTempIntSubCategory=1
	end if
intIdCategory=pTempIntSubCategory

'// Wholesale-only categories
If Session("customerType")=1 Then
	pcv_strTemp=""
else
	pcv_strTemp=" AND pccats_RetailHide<>1"
end if

'*******************************
' START Display Settings
'*******************************

pFeaturedCategory=0
pFeaturedCategoryImage=0

If validNum(pTempIntSubCategory) and pTempIntSubCategory<>1 then
	query="SELECT pcCats_SubCategoryView, pcCats_CategoryColumns, pcCats_CategoryRows, pcCats_PageStyle, pcCats_ProductOrder, pcCats_ProductColumns, pcCats_ProductRows, pcCats_FeaturedCategory, pcCats_FeaturedCategoryImage FROM categories WHERE (((idCategory)="&pTempIntSubCategory&")" & pcv_strTemp &");"

	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	if rs.EOF then
		set rs=nothing
		call closeDb()
		response.redirect "msg.asp?message=86"
	end if	
	
	Dim pIntSubCategoryView
	Dim pIntCategoryColumns
	Dim pIntCategoryRows
	Dim pIntProductColumns
	Dim pIntProductRows
	
	pIntSubCategoryView=rs("pcCats_SubCategoryView")
	pIntCategoryColumns=rs("pcCats_CategoryColumns")
	pIntCategoryRows=rs("pcCats_CategoryRows")
	pStrPageStyle=rs("pcCats_PageStyle")
	pStrProductOrder=rs("pcCats_ProductOrder")
	pIntProductColumns=rs("pcCats_ProductColumns")
	pIntProductRows=rs("pcCats_ProductRows")
	pFeaturedCategory=rs("pcCats_FeaturedCategory")
	pFeaturedCategoryImage=rs("pcCats_FeaturedCategoryImage")
	
	set rs=nothing
	
	Session("pStrPageStyle")=pStrPageStyle
End if
	
' START Load category-specific values. If empty, use storewide settings

' How sub-categories are displayed
' 	0 = in a list, with images
'	1 = in a list, without images
'	2 = drop-down
'	3 = default
'	4 = thumbnail only
if NOT validNum(pIntSubCategoryView) OR pIntSubCategoryView=3 then
	 pIntSubCategoryView=scCatImages
end if

' How many per row: number of columns
if NOT validNum(pIntCategoryColumns) OR pIntCategoryColumns=0 then
	pIntCategoryColumns=scCatRow
end if

' How many rows per page
if NOT validNum(pIntCategoryRows) OR pIntCategoryRows=0 then
	pIntCategoryRows=scCatRowsPerPage
end if

' How many products per row
if NOT validNum(pIntProductColumns) OR pIntProductColumns=0 then
	pIntProductColumns=scPrdRow
end if

' How many rows per page
if NOT validNum(pIntProductRows) OR pIntProductRows=0 then
	pIntProductRows=scPrdRowsPerPage
end if

' END Load category-specific values


' OVERRIDE page style: check to see if a querystring or a form is sending the page style.
Dim pcPageStyle, strSeoQueryString

pcPageStyle = LCase(getUserInput(Request("pageStyle"),1))

'// Check querystring saved to session by 404.asp
if pcPageStyle = "" then
	strSeoQueryString=lcase(session("strSeoQueryString"))
	if strSeoQueryString<>"" then
		if InStr(strSeoQueryString,"pagestyle")>0 then
			pcPageStyle=left(replace(strSeoQueryString,"pagestyle=",""),1)
		end if
	end if
end if

'// Category Level Settings
if pcPageStyle = "" then
	pcPageStyle = pStrPageStyle
end if

'// Global Settings
if isNULL(pcPageStyle) OR trim(pcPageStyle) = "" then
	pcPageStyle = LCase(bType)
end if

if pcPageStyle <> "h" and pcPageStyle <> "l" and pcPageStyle <> "m" and pcPageStyle <> "p" then
	pcPageStyle = LCase(bType)
end if

' OTHER display settings
' These variables show/hide information when products are shown with Page Style = L or M
Dim pShowSKU, pShowSmallImg
pShowSKU = scShowSKU ' If 0, then the SKU is hidden
pShowSmallImg = scShowSmallImg ' If 0, then the small image is not shown
' Note: the size of the small image is set via the css/pcStorefront.css stylesheet

'FB-S
if (session("Facebook")="1") AND (session("pcFBS_CustomDisplay")="1") then
	pIntSubCategoryView=session("pcFBS_CatImages")
	pIntCategoryColumns=session("pcFBS_CatRow")
	pIntCategoryRows=session("pcFBS_CatRowsperPage")
	pIntProductColumns=session("pcFBS_PrdRow")
	pIntProductRows=session("pcFBS_PrdRowsPerPage")
	pcPageStyle = session("pcFBS_BType")
	pShowSKU = session("pcFBS_ShowSKU")
	pShowSmallImg = session("pcFBS_ShowSmallImg")
end if
'FB-E

'// Check For Mobile Storefront Overrides
If session("Mobile")="1" Then
	pIntSubCategoryViewBAK=pIntSubCategoryView
	pIntSubCategoryView=0
	pIntCategoryColumns=1
	pIntCategoryRows=10
	pIntProductColumns=1
	pIntProductRows=10
	pcPageStyle = "h"
End If

'*******************************
' END Display Settings
'*******************************


if pFeaturedCategory<>0 then
	pcv_strTemp=pcv_strTemp&" AND idCategory<>"&pFeaturedCategory & " "
end if

dim pIdCategory, pCategoryDesc, pcStrViewAll

rMode=server.HTMLEncode(request.querystring("mode"))
if rMode="" then
	iPageSize=(pIntProductColumns*pIntProductRows)
	iCatPageSize=(pIntCategoryColumns*pIntCategoryRows)
	If Request("page")="" Then
		iPageCurrent=1
	Else
		iPageCurrent=CInt(Request("page"))
	End If
end if

'// View All
pcStrViewAll = Lcase(getUserInput(Request("viewAll"),3))
if pcStrViewAll = "yes" then
	iPageSize = 9999
end if	

if NOT validNum(iPageSize) OR iPageSize=0 then
	iPageSize=5
end if

pIdCategory=session("idCategoryRedirect")
mIdCategory=session("idCategoryRedirect")
'DA - EDIT
pIdCategory=pTempIntSubCategory
mIdCategory=pTempIntSubCategory
if pIdCategory="" then
	pIdCategory=getUserInput(request.querystring("idCategory"),10)
	mIdCategory=getUserInput(request.querystring("idCategory"),10)
	'// Validate Category ID
	if not validNum(pIdCategory) then
		pIdCategory=""          
	end if
	if not validNum(mIdCategory) then
		mIdCategory=""          
	end if
	
	if pIdCategory="" then
		pIdCategory=1
		mIdCategory=1
	end if
end if
session("idCategoryRedirect")=""

'*******************************
' get category tree array
'*******************************
if pIdCategory<>1 then %>
	<!--#include file="pcBreadCrumbs.asp"-->
<% end if

'*******************************
' End get category tree array
'*******************************

'*******************************
' Get sub-categories array
'*******************************
Dim intSubCatExist
Dim iCategoriesPageCount
intSubCatExist=0

IF pIdCategory=1 THEN
	scCatTotal=(pIntCategoryColumns*pIntCategoryRows)
	if pIntSubCategoryView="2" then
		scCatTotal=999999
	end if
	iCategoriesPageSize=scCatTotal
	if pcStrViewAll = "yes" then
		iCategoriesPageSize = 9999
	end if
	
	Dim pcInt_CategoriesPage
	pcInt_CategoriesPage=getUserInput(request("CategoriesPage"),10)
	if not validNum(pcInt_CategoriesPage) then
		iCategoriesPageCurrent=1
	Else
		iCategoriesPageCurrent=Cint(pcInt_CategoriesPage)
	End If

	query = "SELECT idCategory,categoryDesc,[image],idParentCategory,SDesc,HideDesc FROM Categories WHERE idParentCategory=1 AND idCategory<>1 AND iBTOhide=0 " & pcv_strTemp & " ORDER BY priority, categoryDesc ASC;"
	SET rs=Server.CreateObject("ADODB.RecordSet")

	rs.PageSize=iCategoriesPageSize
	pcv_strPageSize=iCategoriesPageSize
	rs.CacheSize=iCategoriesPageSize
		
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	
	'// Page Count
	iCategoriesPageCount=rs.PageCount
	If Cint(iCategoriesPageCurrent) > Cint(iCategoriesPageCount) Then iCategoriesPageCurrent=Cint(iCategoriesPageCount)
	If Cint(iCategoriesPageCurrent) < 1 Then iCategoriesPageCurrent=1	
ELSE
	scCatTotal=(pIntCategoryColumns*pIntCategoryRows)
	if pIntSubCategoryView="2" then
		scCatTotal=999999
	end if
	iCategoriesPageSize=scCatTotal
	if pcStrViewAll = "yes" then
		iCategoriesPageSize = 9999
	end if
	
	pcInt_CategoriesPage=getUserInput(request("CategoriesPage"),10)
	if not validNum(pcInt_CategoriesPage) then
		iCategoriesPageCurrent=1
	else
		iCategoriesPageCurrent=Cint(pcInt_CategoriesPage)
	end if
	
	query = "SELECT idCategory, categoryDesc FROM Categories WHERE idParentCategory = " & pIdCategory & " AND idCategory<>1 AND iBTOhide=0 " & pcv_strTemp & " ORDER BY priority, categoryDesc ASC;"
	set rs=Server.CreateObject("ADODB.RecordSet")

	rs.PageSize=iCategoriesPageSize
	pcv_strPageSize=iCategoriesPageSize
	rs.CacheSize=iCategoriesPageSize
		
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	
	'// Page Count
	iCategoriesPageCount=rs.PageCount
	If Cint(iCategoriesPageCurrent) > Cint(iCategoriesPageCount) Then iCategoriesPageCurrent=Cint(iCategoriesPageCount)
	If Cint(iCategoriesPageCurrent) < 1 Then iCategoriesPageCurrent=1	
END IF

If NOT rs.EOF Then
	rs.AbsolutePage=iCategoriesPageCurrent
	intSubCatExist=1
	SubCatArray=rs.GetRows(iCategoriesPageSize)
	intSubCatCount=ubound(SubCatArray,2)
End If

SET rs=nothing
'*******************************
' End get sub-categories array
'*******************************
%>

<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="pcValidateQty.asp"-->

	<!-- Header: pagetitle -->
    <header id="computercontent" class="computercontent">
		<div class="pc-content">
			<div class="container">
				<div class="row">
					<div class="col-md-7">
                         <div class="wow fadeInDown pt-headtext" data-wow-offset="0" data-wow-delay="0">
							<h1>Trading Computers</h1>
							<h2>Specially Crafted Computers Designed To <span>Power Your Trading Sessions</span> With Ease</h2>
						 </div>
						 <div class="wow fadeInUp" data-wow-offset="0" data-wow-delay="0">
							<p class="home-head-text text-white text-justify">Whether you run MT4 with one screen or TradeStation across twelve screens, you need a trading computer that is responsive, stable, and fit for purpose.</p>
						    <p class="home-head-text text-white text-justify">There is no 'one size fits all' trading computer spec so we have put together two specially designed options that cover virtually any trading platform requirements, let's find out which one is right for you.</p>
						</div>
                        <div class="wow fadeInDown pt-headtext" data-wow-offset="0" data-wow-delay="0">
                            <h2>Ready to go? <a href="javascript:tradingcomplpcustomjump();void(0);">Jump straight to the computers</a>.</h2>
                            </div>
                    </div>
					<div class="col-md-5">
                         <div class="wow fadeInRight text-center" data-wow-offset="0" data-wow-delay="0.1s">
							  <img src="/images/pc-bannerimage.png" alt="Trading Computers">
						 </div>
                    </div>
					
				</div>		
			</div>		
		</div>	
    </header>

	</section>	
    
    	<!--#include file="banner.asp"-->
    
    <section id="deliveryAreas" class="deliveryAreas paddingtop-40 paddingbot-40">            <div class="container"> <div class="row">      <div class="pinAreas wow fadeInLeft marginbot-20" data-wow-delay="0"> <div class="col-sm-4 pright-sm pinArea-icon">    <div class="dareaIcons ssImg1">         <img src="/images/trader-pc/trusted.jpg" alt="Trusted by Traders" />    </div> </div>     <div class="col-sm-8 pinArea-text"> <h2 class="trading">Specialist Trading Computers</h2> <p>If you are in the market for a new computer for trading then you will see a lot of conflicting information.  </p>
	          <p>Sometimes this is from companies trying to make a 'quick buck' off you, or who don't really know what they are doing.</p>
              <p>Sometimes it's from other traders claiming you do or don't need this or that.</p>
	          <p>It can be very difficult to know who to believe or what to buy. </p>
	          <p>The truth is that there is no one setup that is the best trading computer.</p>
	          <p>Different traders, using different trading platforms and tools, implementing different trading strategies, need different things from their trading computers.</p>
            <p>Recognising this our goal over the past few years has been to put together quality, in-depth information, that traders like you can use to determine which computer spec will best meet your needs.</p>
            <p>So let's dive in and discover what is the best trading computer setup for you.</p> 
</div>  
	</div>  
    <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text"> <h2 class="trading">Trading Platform Performance</h2></div></div>
    <div class="pinAreas wow fadeInRight marginbot-20" data-wow-delay="0.1s"> <div class="col-sm-5 pinArea-text">
     <p>We have done a lot of testing, we have an entire other website dedicated to our trading computer tests.
</p>
	    <p><strong>Our testing clearly shows that the number one impact on the performance of your trading software is the speed of your computers processor.</strong></p>
	    <p>Sure, you need enough memory (RAM), you should have a solid state hard drive (SSD), and your graphics setup needs to be able to run your desired number of screens.</p>
	    <p>But after those baselines have been covered, pretty much the only performance lever in a trading computer is the processor (CPU).</p>
        <p>The raw speed of a CPU is measured by a single thread speed test.</p>
	    <p>This chart with data pulled from our TraderSpec.com in-depth testing shows the single thread speed comparisons of various processors.</p>
        <p>As you can see the selection of CPU's we offer all rate very highly in this single thread speed category. </p>
		<p>The Intel 14th generation chips (green bars) are great options, but lose out to the newer Intel Core Ultra chips (blue bars).</p>
		<p>The 2 AMD chips (red bars) also score well however they are more expensive than the equivalently priced Intel options. </p>
        </div>
        <div class="col-sm-7 pinArea-icon">
        <div class="dareaIcons tcimg1 wow fadeInRight" data-wow-delay="0.1s">         <img src="/images/trader-pc/single-thread-speed.jpg" alt="Trading Computer Processors"/>  </div></div>
        <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text">  
        <p>The bottom three options are CPUs recommended on some competitors websites, as you can see they are not great options for most traders. Some are on cheap 2nd hand systems, some are sold 'as new' despite being almost 4 years old...</p>
                </div>
    </div>
    <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-5 pinArea-text">
        <p><strong>For traders running multiple platforms simultaneously, or running lots of screens, browser tabs, and charts, then multi-tasking ability becomes important.</strong></p>
		<p>Some backtesting software will also use a CPUs multi-threaded features to reduce analysis time.</p>
        <p> How suited a processor is to multi-tasking is highly related to the number of CPU Cores it has, the higher the Core count, the more work it can process simultaneously.</p>
        <p>This next chart, again created with data pulled from TraderSpec.com, shows the multi-threaded performance levels of the same CPU's.</p>
        <p>Here we see the top Intel Core Ultra 9 chip pulling ahead of the field, the AMD 9950X is also very strong in multi-threaded performance.</p>
		<p>The Core Ultra 7 265KF almost matches the 14th gen i9 and is way infront of everything apart from the two expensive AMD and Intel offerings.</p>
		<p>Our lowest cost Intel i5 pretty much matches the older Intel i9 10920X which is amazing considering the price difference.</p>
		<p>The reality is that for the vast majority of traders the Intel i5 14400F is more than you will ever need.</p>
        <p>Only traders running the most intensive software, or wanting to run 3, 4, or 5+ platforms simultaneously would need to go to the higher Intel or AMD performance levels.</p>
        </div>
        <div class="col-sm-7 pinArea-icon">
        <div class="dareaIcons tcimg1 wow fadeInRight" data-wow-delay="0.1s">         <img src="/images/trader-pc/multi-thread-speed.jpg" alt="Multi-Threaded Trading Performance"/>  </div></div>
    </div>  
	</div>  
    <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text"> <h2 class="trading">Graphics Options For Traders</h2></div></div>
    <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-5 pinArea-text">
     <p>Something we hear regularly is 'I've been told I need a powerful graphics card for trading'.</p>
	    <p>This is incorrect.</p>
	    <p>Most trading packages essentially display text, numbers and lines, they don't create 3D graphics. Line, bar and candlestick charts are made up of 2D graphics.</p>
        <p>This is important as 2D graphics are far less taxing for a computer to create and draw to your screen than a 3D model.</p>
        <p>This means even a fairly low cost and low powered graphics card can often easily handle the output of a trading platform.</p>
	    <p>This chart from TraderSpec.com shows the impact of different graphics cards on a simulated web browser based trading software workload.</p>
        <p><strong>As you can clearly see, the high end gaming graphics card (GTX 1060), the low priced card (GT 710) and the expensive 'professional class card' (NVS 510) all perform exactly the same.</strong></p>
        </div>
        <div class="col-sm-7 pinArea-icon">
        <div class="dareaIcons tcimg1 wow fadeInRight" data-wow-delay="0.1s">         <img src="/images/trader-pc/graphics-card-test.jpg" alt="Trading Graphics Cards" />  </div></div></div>
    	<div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text">  
		<p>The only times we have seen a real need for more powerful cards is for customers running a lot of charts across a lot of screens, something like NinjaTrader across 12 monitors, or Bloomberg across 6 screens with a high number of charts per screen (20+). In these situations we do recommend more capable 8GB nVidia RTX 4060 cards to handle that kind of load, but this is not commonly needed.</p>
			<p>If you are looking for good AI performance then a more powerful card will make a big difference here.</p>
			<p>We now list an AI TOPS score for each graphics card, the bigger the number the better suited the card is to running complex AI workloads. The Trader Pro PC has the more powerful graphics card options and is the machine for people needing strong AI performance levels.</p>
                </div>
    </div>
        <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text"> <h3 class="trading">Multiple Monitor Outputs</h3></div></div>
    <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-5 pinArea-icon pright-sm">
        <div class="dareaIcons tcimg1 wow fadeInRight" data-wow-delay="0.1s">         <img src="/images/trader-pc/six-array.png" alt="Monitor Arrays" />  </div></div> <div class="col-sm-7 pinArea-text">
     <p>For each screen you want to connect to your computer you need a corresponding monitor output port on the PC, higher number of ports are achieved by running multiple graphics cards.</p>
	    <p>For traders we recommend and use either Intel or nVidia  cards, these are professional class cards that run quiet, are relatively low power draw, but can easily handle multiple high resolution screens.</p>
	    <p>We have configurations that support 4, 6, 8, 10 or even 12 monitors, with many options supporting QHD, 4K and even 5K screens.</p>
        <h3 class="trading">4K / High Resolution Screens</h3>
	    <p>Standard screens run at an FHD resolution of 1920 x 1080 pixels, these screens are low cost and easy to run from your computer.</p>
        <p>There are higher resolution options now available which offer more usable space to layout your charts and programs.</p></div>
        <div class="pinAreas wow fadeInRight marginbot-20" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text">  
	    <p>So while higher resolutions can offer your more space to display your trading information they do tend to cost more, and are more taxing on computers often requiring a graphics setup that can fully support them.</p>
        <p>They are also a great source of misunderstanding.</p>
        <p>A 4K resolution technically means it has four times the number of pixels over a standard FHD screen, so in theory it can display four times the amount of content that you can fit onto a normal screen.</p>
		<p>In practice you can only use this extra space fully if you increase the size of the screen, for most people this means a 40" (or bigger) monitor.</p>
		<p>Learn exactly why this is and <a href="/pages/monitor-resolutions/" target="_new">find out more about high resolution screens on our new screen resolution information page.</a> (Opens in new tab)</p>
		<p>When it comes to trading computers, you need to make sure you can support the right number and type of screens that you want to run.</p>
        </div></div>
<div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text"> <h2 class="trading">The Rest Of The Spec</h2></div></div>
    <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text">
     <p>We have covered the CPU and the graphics setup, but what else needs to go into a trading computers build?</p>
	    <h3 class="trading">Hard Drive</h3>
        <img src="/images/trader-pc/trader-ssd.png" alt="HDD / SSD Hard Drive" style="float:right; margin-left:20px;" />
        <p>A hard drive is where Windows and your trading software is installed, it will also hold any data files you create or use.</p>
        <p>There are two main types of hard drive, the traditional 'platter' style drives which have moving parts in them, and newer solid state hard drives (SSD's).</p>
        <p>Solid state drives have no moving parts and are much quicker, meaning your computer will power on quicker and programs will open instantly.</p>
        <p>For trading computers a standard solid state drive is required.
          You can get faster SSD's however they make no difference to how fast your trading software will run at all, absolutely zero impact, so if you are on a budget don't waste your money here, get a normal SSD and know that is all you need.</p>
        <h3 class="trading">RAM / Memory</h3>
        <img src="/images/trader-pc/trader-ram.png" alt="RAM / Memory" style="float:right; margin-left:20px;" />
        <p>RAM in a computer is temporary storage that holds your open programs and files.</p>
          <p>When you open a chart or spreadsheet it is transferred into your RAM and used from there.</p>
          <p>When RAM gets full then your trading computers performance will slow down, a lot!</p>
          <p>So how much RAM do you need in your trading PC?
            It depends on how many platforms, charts and files you want to open and run at the same time.</p>
          <p>For MT4 with 4 - 6 charts, 8GB of RAM is enough.</p>
          <p>Platforms like NinjaTrader and TradeStation will use more RAM so 16GB - 32GB is recommended for these.</p>
          <p>If you use primarily use web platforms like IG, Pro-Realtime, or Trading View then it comes down to how many browser tabs you tend to open, web browsers can be hungry for RAM.</p>
          <p>8GB will cover around 4 - 6 tabs, any more than this then we would recommend going with 16GB - 32GB. We have some customers who open 40+ browser tabs, for them 32GB of RAM was needed to ensure everything still ran smoothly.</p>
          <p>In terms of RAM speed, we have never conclusively seen evidence that faster RAM makes any real world impact on trading software workload performance, so don't worry about it.</p>
        <h3 class="trading">The Other Stuff</h3>
        <p>A decent case helps with cooling, which is important for a stable trading computer.</p>
          <p>For power a 500W rated supply is usually more than enough, higher powered supplies do not make any difference to your computers performance. Only higher powered gaming graphics cards require more power, if you upgrade to one of these we automatically bump up the power supply rating for you.</p>
          <p>Finally, if you want to use a wireless Internet connection with the best possible speeds then a Wireless AX network card is what you need, they use the fastest WiFi protocol available.</p>
                </div>
                <div class="pinAreas wow fadeInRight marginbot-0 margintop-20" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text"> <h2 class="trading">The Best Trading Computer For You</h2></div></div>
    <div class="pinAreas wow fadeInRight marginbot-0" data-wow-delay="0.1s"> <div class="col-sm-12 pinArea-text">
     <p>Let's put this all into practice and decide what you need based on what type of trading you do.</p>
     <h3 class="trading"><i class="fa fa-check green-link"></i> Forex, Day Trading, MT4 & General Charting</h3>
     <p>If your trading relies on having fast access to market data then you want to be looking for a fast CPU, single thread speed is key.</p>
     <p>Budget friendly but still highly responsive, our Trader PC (see below) configured with an Intel 14th generation i5 will more than do the job for you.</p>
     <p>For these workloads they will outperform computers offered by some other companies priced at 2 - 3 times the price!</p>
     <p>If only the best is good enough then look at moving to the higher power Intel i5 or even i7 processors, they are fantastic for this kind of workload.</p>
     <h3 class="trading"><i class="fa fa-check green-link"></i> TradeStation, NinjaTrader, Backtesting, EA's / Indicators, Multiple Platforms</h3>
     <p>If you run a few platforms at once, make use of a lot of Expert Advisors / Indicators, or do some backtesting analysis then our Trader PC is still highly capable, upgrades to the CPU &amp; RAM will improve performance or going for the Trader Pro PC is another option to help power through that workload.</p>
     <p>The 14th generation i5 14600KF is well suited to this kind of work, the newer Core Ultra series chips are also great for multi-tasking. You would still be better off sticking with Intel here over AMD both in terms of performance levels and cost.</p>
     <h3 class="trading"><i class="fa fa-check green-link"></i> BloomBerg, X Trader, Lots of Charts, More Intensive Spreadsheets, Extensive Backtesting Jobs</h3>
     <p>Here you want a fast single thread speed and high core count CPU, so going for the Intel 14th generation i7 or i9 is a good option on the Trader PC. On the Trader Pro PC the Core Ultra 7 and 9 chips are very strong but do increase the cost a fair bit.</p>
     <p>In terms of RAM we would generally recommend pairing these processors with 32GB - 64GB of RAM to ensure you don't run into a bottleneck here.</p>
     <p>As mentioned above, when running some of the more intensive platforms across lots of screens with an ultra high number of charts per screen then you may want to go with the more powerful nVidia RTX 4060 or 5060 graphics card options. For ultimate AI workload performance add the nVidia RTX 5070 and 5080 options to your Trader Pro PC.</p>
                </div>
     </div></section>
                
<section id="pc-tradingsection2" class="pc-tradingsection2 paddingtop-20 paddingbot-40">
		<div class="container">
        <div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h2 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-line-chart green-link"></i> <span>Pick Your </span>Trading Computer</h1>
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1 marginbot-20">SELECT AND CONFIGURE YOUR PERFECT TRADING COMPUTER</h5>
						</div>
			    <div class="col-md-6 multsection-col mmc-product">
				    <div class="multi-submenu2 wow fadeInUp" data-wow-delay="0.1s">
					   <div class="row">	
							 <div class="col-sm-4 mmc-product-img">
							    <img src="/images/pc-1.png">
							 </div>
							 <div class="col-sm-8 mmc-product-text">
							    <h2>The Trader</h2>
								<h3>Trading PC</h3>
								<p>Highly responsive, great for MT4, web based trading, charting platforms, and packages like NinjaTrader, TradeStation &amp; Sierra Charts.</p>
								<h4>Monitors Supported: Up to 10</h4>
								<h5>Price From: <span>&pound;1,045.00</span></h5>
								<a href="/products/trader-pc/" class="btn btn-skin btn-wc semi pcnw-btn margintop-20">View &amp; Customise Your Trader PC <i class="fa fa-angle-right"></i></a>
							 </div>
                        </div>
					</div>
				</div> <!-- md 6 -->
				<div class="col-md-6 multsection-col mmc-product">
				   <div class="multi-submenu2 wow fadeInUp" data-wow-delay="0.1s">  
					    <div class="row">	
							 <div class="col-sm-4 mmc-product-img">
							    <img src="/images/pc-2.png">
							 </div>
							 <div class="col-sm-8 mmc-product-text">
							    <h2>The Trader Pro</h2>
								<h3>Trading PC</h3>
								<p>Top of the range trading computer capable of powering any trading session. Perfect for trading professionals and platform power users.</p>
								<h4>Monitors Supported: Up to 12</h4>
								<h5>Price From: <span>&pound;1,345.00</span></h5>
								<a href="/products/trader-pro-pc/" class="btn btn-skin btn-wc semi pcnw-btn margintop-20">View &amp; Customise Your Trader Pro <i class="fa fa-angle-right"></i></a>
							 </div>
                        </div>
				   </div>
				</div> <!-- md 6 -->
                </div><!-- container -->
                </section>

     <section id="product-stands" class="product-stands bg-smog product-grid paddingtop-20 paddingbot-40">
		<div class="container">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h2 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cubes green-link"></i> Trading Computer Bundles</h1>
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1 marginbot-20">Money saving deals including Trading PC, Multi-Screen Stand, &amp; Monitors, in one simple purchase</h5>
						</div>			
			 <div class="row">
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Dual 21.5 inch Monitor Array & Charter PC Bundle"><a class="" href="/products/trader-pc/?sid=287&mid=304&cid=333">Dual 21.5" Array & Trader PC</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/products/trader-pc/?sid=287&mid=304&cid=333"><img src="/images/bundles/dual-tra-bundle.jpg" alt="Dual 21.5 inch Monitor Array & Trader PC Bundle" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Dual 21.5" Widescreen Monitor Array and a responsive Trader PC.</p>
								
								<h4>From: <span>&pound;1,315.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/products/trader-pc/?sid=287&mid=304&cid=333">View Bundle Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Triple 21 inch Monitor Array & Charter PC Bundle"><a class="" href="/products/trader-pc/?sid=312&mid=304&cid=333">Triple 21.5" Array & Trader PC</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/products/trader-pc/?sid=312&mid=304&cid=333"><img src="/images/bundles/triple-pro-bundle.jpg" alt="Triple 21 inch Monitor Array & Trader PC Bundle" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Triple 21.5" Widescreen Monitor Array and a Trader PC bundle.</p>
								
								<h4>From: <span>&pound;1,450.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/products/trader-pc/?sid=312&mid=304&cid=333">View Bundle Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Quad 24 inch Monitor Array & Charter PC Bundle"><a class="" href="/products/trader-pc/?sid=313&mid=317&cid=333">Quad 24" Array & Trader PC</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/products/trader-pc/?sid=313&mid=317&cid=333"><img src="/images/bundles/quad-tra-bundle.jpg" alt="Quad 24 inch Monitor Array & Trader PC Trading Bundle" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Quad 24"  Monitor Array and a Trader Multi-Screen Computer.</p>
								
								<h4>From: <span>&pound;1,570.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/products/trader-pc/?sid=313&mid=317&cid=333">View Bundle Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">

                             	<h3 title="Triple 24 inch Monitor Array & Trader PC Bundle"><a class="" href="/products/trader-pro-pc/?sid=312&mid=317&cid=343">Triple 24" Array & Trader Pro PC</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/products/trader-pro-pc/?sid=312&mid=317&cid=343"><img src="/images/bundles/triple-pro-bundle.jpg" alt="Triple 24 inch Monitor Array & Trader Pro PC Bundle" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Triple 24" Widescreen Monitor Array and a Trader Pro Multi-Screen PC.</p>
								
								<h4>From: <span>&pound;1,765.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/products/trader-pro-pc/?sid=312&mid=317&cid=343">View Bundle Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Quad 21.5 inch Monitor Array & Trader PC Bundle"><a class="" href="/products/trader-pro-pc/?sid=313&mid=304&cid=343">Quad 21.5" Array & Trader Pro PC</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/products/trader-pro-pc/?sid=313&mid=304&cid=343"><img src="/images/bundles/quad-tra-bundle.jpg" alt="Quad 21.5 inch Monitor Array & Trader Pro PC Bundle" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Quad 21.5" Monitor Array and an Trader Pro Multi-Screen Computer.</p>
								
								<h4>From: <span>&pound;1,850.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/products/trader-pro-pc/?sid=313&mid=304&cid=343">View Bundle Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
			    <div class="col-md-4 col-sm-6 product-col">
				    <div class="product-detail wow fadeInUp" data-wow-delay="0">
					   <div class="row">	
							 <div class="col-sm-12 col-xs-8 pg-product-tytl">
                             	<h3 title="Six 24 inch Monitor Array & Trader PC Bundle"><a class="" href="/products/trader-pro-pc/?sid=338&mid=317&cid=343">Six 24" Array & Trader Pro PC</a></h3>
							 </div>
							 <div class="col-sm-4 col-xs-4 pg-product-img">
							 
								<a href="/products/trader-pro-pc/?sid=338&mid=317&cid=343"><img src="/images/bundles/six-ult-bundle.jpg" alt="Six 24 inch Monitor Array & Trader Pro PC Bundle" /></a>
							
							 </div>
							 <div class="col-sm-8 col-xs-12 pg-product-text">
								<p>A Six 24" Widescreen Monitor Array and a Trader Pro PC Bundle.</p>
								
								<h4>From: <span>&pound;2,125.00</span></h4>
			
								<div class="pg-btns">
                                	<a title="More Info" class="btn product-action btn-skin pg-blue-btn" href="/products/trader-pro-pc/?sid=338&mid=317&cid=343">View Bundle Details</a>
								</div>
							 </div>
               			</div>
					</div>
				</div> <!-- product-col -->
                <div class="text-center">
                <h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1 marginbot-20">Want to build your own perfect trading bundle?<br/><a href="/bundles/" style="text-decoration:underline;">Visit our Bundles page to get started</a>.</h5>
                </div>
	</section>
    <!-- /Section: Welcome -->
        

<!--#include file="footer_wrapper.asp"-->
<!--#include file="bulkAddToCart.asp"-->
