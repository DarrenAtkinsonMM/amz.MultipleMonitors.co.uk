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

pTempIntSubCategory=14

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

<script>
$(document).on('click',function(){
$('.collapse').collapse('hide');
})
</script> 


	<!-- Header: pagetitle -->
    <header id="computercontent" class="computercontent">
		<div class="pc-content">
			<div class="container">
				<div class="row">
					<div class="col-md-7">
                         <div class="wow fadeInDown pt-headtext" data-wow-offset="0" data-wow-delay="0">
							<h1>Multiple Monitor Computers</h1>
							<h2>Make the transition to a <span>multi-screen system</span> easily with our <span>hassle free,</span> multi monitor capable computers</h2>
						 </div>
						 <div class="wow fadeInUp" data-wow-offset="0" data-wow-delay="0">
							<p class="home-head-text text-white text-justify">Supporting more than 1 or 2 screens is not possible on the vast majority of PC's, however we have a range of dedicated computers to make your transition easy!</p>
						    <p class="home-head-text text-white text-justify">Simply select a computer below then use the options to choose the number of screens you want to support, we then build your new PC just for you.</p>
						</div>
                        <div class="wow fadeInDown pt-headtext" data-wow-offset="0" data-wow-delay="0">
                            <h2>Are you a trader? <a href="javascript:computerscustomjump();void(0);">Jump straight to our Trading Computers</a>.</h2>
                            </div>
                    </div>
					<div class="col-md-5">
                         <div class="wow fadeInRight text-center" data-wow-offset="0" data-wow-delay="0.1s">
							  <img src="/images/pc-bannerimage.png">
						 </div>
                    </div>
					
				</div>		
			</div>		
		</div>	
    </header>

	</section>

	<!--#include file="banner.asp"-->
	
    <header id="product-stands" class="bundle-wrap bg-smog">
		<div class="intro-content paddingtop-20 paddingbot-10">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> Multiple Monitor <span>Computers</span></h1>
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1 marginbot-20">OFFERING A WIDE RANGE OF CUSTOMISATION THESE MULTI MONITOR PC'S CAN MEET ANY REQUIREMENTS </h5>
					</div>				
				</div>		
			</div>
		</div>
        </div>	
    </header>
	
	<section id="pc-multisection" class="pc-multisection paddingtop-20 paddingbot-40 bg-smog">
		<div class="container">
			<div class="row">
				<div class="col-md-6 multsection-col mmc-product">
				   <div class="multi-submenu wow fadeInUp" data-wow-delay="0.1s">  
					    <div class="row">	
							 <div class="col-sm-4 mmc-product-img">
							    <img src="/images/pc-2.png">
							 </div>
							 <div class="col-sm-8 mmc-product-text">
							    <h1>Ultra</h1>
								<h2>multi Screen pc</h2>
								<p>Ultra fast performance using Intel 14th  gen. CPU's in a highly configurable &amp; virtually silent package.</p>
								<h3>Monitors Supported:  4, 6, 8 or 10</h3>
								<h4>Price From: <span>&pound;925.00</span></h4>
								<a href="/products/ultra-multi-monitor-pc/" class="btn btn-skin btn-wc semi pcnw-btn margintop-20">View &amp; Customise Your ultra PC <i class="fa fa-angle-right"></i></a>
							 </div>
                        </div>
				   </div>
				</div> <!-- md 6 -->
				<div class="col-md-6 multsection-col mmc-product">
				    <div class="multi-submenu wow fadeInUp" data-wow-delay="0.1s">
					   <div class="row">	
							 <div class="col-sm-4 mmc-product-img">
							    <img src="/images/pc-3.png">
							 </div>
							 <div class="col-sm-8 mmc-product-text">
							    <h1>Extreme</h1>
								<h2>multi Screen pc</h2>
								<p>Designed for power users requiring the ultimate multi-threaded performance.</p>
								<h3>Monitors Supported:  4, 6, 8, 10, or 12</h3>
								<h4>Price From: <span>&pound;1,195.00</span></h4>
								<a href="/products/extreme-multi-screen-computer/" class="btn btn-skin btn-wc semi pcnw-btn margintop-20">View &amp; Customise Your extreme PC <i class="fa fa-angle-right"></i></a>
							 </div>
                        </div>
					</div>
				</div> <!-- md 6 -->
				
               
				<div class="clr"></div>
			</div>
		</div>
	</section>
    <!-- /Section: Welcome -->
	
		<!-- Header: intro -->
    <header id="bundle-stands" class="bundle-wrap bg-lyt">
		<div class="intro-content paddingtop-20">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div id="tradinghead" class="text-center marginbot-40" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-line-chart green-link"></i> <span>Specialist </span>Trading Computers</h1>
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1 marginbot-20">DEDICATED TRADING COMPUTERS CAPABLE OF HANDLING YOUR WORKFLOW NO MATTER WHAT YOUR PLATFORM OF CHOICE IS</h5>
						</div>
                <div class="col-md-12 multsection-col mmc-product">
				    <div class="multi-submenu wow fadeInUp" data-wow-delay="0.1s" style="min-height:0px;">
					   <div class="row">	
							 <div class="col-sm-2 mmc-product-img">
							    <img src="/images/pc-1.png">
							 </div>
							 <div class="col-sm-8 mmc-product-text">
							    <h1>Trading Computers</h1>
								<h2>Designed Just For Traders</h2>
								<p>Two dedicated computers that can be configured to meet any type of trading requirements, supporting platforms like Trading View, IG Index, MT4, NinjaTrader, TradeStation, Bloomberg, TT Trader &amp; more.</p>
								<a href="/trading-computers/" class="btn btn-skin btn-wc semi pcnw-btn margintop-20" >Learn More &amp; Discover Your Perfect Trading Computer Here <i class="fa fa-angle-right"></i></a>
							 </div>
                        </div>
					</div>
				</div> <!-- md 6 -->
				<div class="clr"></div>
						
					</div>				
				</div>		
			</div>
		</div>		
    </header>	
		
	<header id="which-pc" class="bundle-wrap bg-smog">
		<div class="intro-content paddingtop-20 paddingbot-10">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-question green-link"></i> Which PC is right for you? </h1>
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1 marginbot-20">Unsure which PC is right for your needs? Read our guide to discover which computer will be a perfect fit for you. </h5>
							<a href="/blog/which-pc-is-best-for-you/" class="btn btn-skin btn-wc semi pcnw-btn margintop-20 marginbot-60" >Find out more about our PC's <i class="fa fa-angle-right"></i></a>
					</div>				
				</div>		
			</div>
		</div>
        </div>	
    </header>

<header id="direct-links" class="bundle-wrap bg-smog">
		<div class="intro-content paddingtop-20 paddingbot-10">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1 marginbot-20"><a href="/pages/dual-monitor-pc/">dual monitor pc's</a> | <a href="/pages/triple-monitor-pc/">Triple monitor pc's</a> | <a href="/pages/quad-monitor-pc/">quad monitor pc's</a> </h5>
					</div>				
				</div>		
			</div>
		</div>
        </div>	
    </header>
		
	<section id="callaction" class="callact-row">	
           <div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="callaction">
							<div class="row">
								<div class="col-md-10">
									<div class="wow fadeInUp" data-wow-delay="0.1s">
									<div class="cta-text">
									<h2 class="h-bold font-light disp-inline">Save Money,</h2>
									<h3 class="h-light font-light disp-inline">Get Free Cables & Free Delivery with a Bundle</h3>
									</div>
									</div>
								</div>
								<div class="col-md-2">
									<div class="wow fadeInRight" data-wow-delay="0.1s">
										<div class="cta-btn">
										<a data-toggle="lightbox" data-title="Multi-Screen Bundles" href="/pop-pages/bundles.htm" class="btn btn-outline">Learn More</a>	
										</div>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
            </div>
	</section>

<!--#include file="footer_wrapper.asp"-->
<!--#include file="bulkAddToCart.asp"-->
