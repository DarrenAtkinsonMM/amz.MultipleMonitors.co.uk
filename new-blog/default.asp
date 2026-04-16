<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../shop/includes/common.asp"-->
<!--#include file="../shop/includes/common_checkout.asp"--> 
<!--#include file="../shop/includes/CashbackConstants.asp"--> 
<!--#include file="../shop/pc/prv_incFunctions.asp"-->

<%
	'Category 999 means load all posts - others just a specific category
	Select Case request.QueryString("catid")
		case 0
			strBlogCatTitle = " | Hardware"
			strBlogCatDesc = "Learn about computer hardware and technologies."
			strquery = " AND pcCont_BlogCat=0"
		case 1
			strBlogCatTitle = " | Product Guides"
			strBlogCatDesc = "Details of our products and the occasional review."
			strquery = " AND pcCont_BlogCat=1"
		case 2
			strBlogCatTitle = " | Setup Guides"
			strBlogCatDesc = "Information about setting up your multi-screen system."
			strquery = " AND pcCont_BlogCat=2"
		case 3
			strBlogCatTitle = " | Software"
			strBlogCatDesc = "Computer software thoughts and discussions."
			strquery = " AND pcCont_BlogCat=3"
		case 4
			strBlogCatTitle = " | Stands"
			strBlogCatDesc = "All about multi-screen stands."
			strquery = " AND pcCont_BlogCat=4"
		case 5
			strBlogCatTitle = " | Trading"
			strBlogCatDesc = "A collection of posts related to trading and computers."
			strquery = " AND pcCont_BlogCat=5"
		case 999
			strBlogCatTitle = ""
			strBlogCatDesc = "Computer hardware, software and anything else multi-screen related"
			strquery = ""
	End Select
	
	'Work out paging - no qs value = 0
	'i controls start of results
	'j is a loop counter for records displayed k controls how many are shown
	'boolPaging let's us know if we need to throw in a next page link
	intBlogPaging = CInt(request.QueryString("page"))
	i = 0
	j = 0
	k = 10
	boolPaging = False
	
	pcv_DefaultTitle="The Multiple Monitors Blog" & strBlogCatTitle
		
	call openDb()
	query="SELECT pcCont_IDPage,pcCont_PageName,pcCont_Published,pcUrl,pcCont_Blog,pcCont_BlogVis,pcCont_BlogCat,pcCont_PubDate,pcCont_BlogIntro FROM pcContents WHERE pcCont_Blog=1 AND pcCont_InActive=0" & strquery & " ORDER BY pcCont_PubDate DESC,pcCont_IDPage DESC"
	set rstemp=server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)

	'loop through return records to build up list
	Do While Not rstemp.EOF
	
	if i >= intBlogPaging Then
	j = j + 1
	if j <= k then
	'Work out category
	select case rstemp("pcCont_BlogCat")
		case 0
			strBlogCat = "<a href=""/blog/hardware/""><i class=""fa fa-folder-open"" style=""margin-right:2px;""></i> Hardware</a>"
		case 1
			strBlogCat = "<a href=""/blog/product-guides/""><i class=""fa fa-folder-open"" style=""margin-right:2px;""></i> Product Guides</a>"
		case 2
			strBlogCat = "<a href=""/blog/setup-guides/""><i class=""fa fa-folder-open"" style=""margin-right:2px;""></i> Setup Guides</a>"
		case 3
			strBlogCat = "<a href=""/blog/software/""><i class=""fa fa-folder-open"" style=""margin-right:2px;""></i> Software</a>"
		case 4
			strBlogCat = "<a href=""/blog/stands/""><i class=""fa fa-folder-open"" style=""margin-right:2px;""></i> Stands</a>"
		case 5
			strBlogCat = "<a href=""/blog/trading/""><i class=""fa fa-folder-open"" style=""margin-right:2px;""></i> Trading</a>"
	end select
	
	strDate = MonthName(Month(rstemp("pcCont_PubDate"))) & ", " & Year(rstemp("pcCont_PubDate"))
	
	strBlogList = strBlogList & "<!-- Post: Start --><div class=""blog-lsit bl-row""><div class=""blog-content blog-list-content""><h3 class=""h-semi marginbot-20 margintop-0""><a href=""/blog/" & rstemp("pcUrl") & "/"">" & rstemp("pcCont_PageName") & "</a></h3><div class=""blog-meta""><div class=""h-post-meta meta-post-time""><i class=""fa fa-calendar"" style=""margin-right:3px;""></i> " & strDate & "</div><div class=""h-post-meta meta-post-category"">" & strBlogCat & "</div></div>" & rstemp("pcCont_BlogIntro") & "<a class=""btn btn-skin btn-lg text-uppercase"" href=""/blog/" & rstemp("pcUrl") & "/"">Read More</a></div></div><!-- Post: End -->"
	end if
	if j > k then
		'we need to page
		boolPaging = True
	end if
	end if
		i = i + 1
		rstemp.MoveNext()
	Loop

	set rstemp=nothing	
	call closeDB()
	
	'Create paging buttons
	if boolPaging = True then
		'Check if we have paging qs or not if we do replace, if not append
		if intBlogPaging = 0 then
			strPagingURLNext = Request.ServerVariables("HTTP_X_REWRITE_URL") & "?page=" & CStr(k)
			strPagingButPrev = ""
			strPagingButNext = "<a class=""btn btn-skin btn-lg text-uppercase marginbot-10"" style=""float:right;"" href=""" & strPagingURLNext & """>Next Page</a>"
		else
			'Replace paging query string with next set of results
			strPagingURLNext = Replace(Request.ServerVariables("HTTP_X_REWRITE_URL"),"?page=" & CStr(intBlogPaging),"?page=" & CStr(intBlogPaging + k))
			'Replace paging query string with previous set of results
			'If prev makes k = 0 then just remove page querystring altogether
			if intBlogPaging - k = 0 then
				strPagingURLPrev = Replace(Request.ServerVariables("HTTP_X_REWRITE_URL"),"?page=" & CStr(intBlogPaging),"")
			else
				strPagingURLPrev = Replace(Request.ServerVariables("HTTP_X_REWRITE_URL"),"?page=" & CStr(intBlogPaging),"?page=" & CStr(intBlogPaging - k))
			end if
			strPagingButPrev = "<a class=""btn btn-skin btn-lg text-uppercase marginbot-10"" href=""" & strPagingURLPrev & """>Previous Page</a>"
			strPagingButNext = "<a class=""btn btn-skin btn-lg text-uppercase marginbot-10"" style=""float:right;"" href=""" & strPagingURLNext & """>Next Page</a>"
		end if
	else
		'Final case, if boolPaging is false check if we have qs value set, if so it means we are at the end of the list and need the previous button, if no qs then list is single page
		if intBlogPaging >= k then
		'If prev makes k = 0 then just remove page querystring altogether
			if intBlogPaging - k = 0 then
				strPagingURLPrev = Replace(Request.ServerVariables("HTTP_X_REWRITE_URL"),"?page=" & CStr(intBlogPaging),"")
			else
				strPagingURLPrev = Replace(Request.ServerVariables("HTTP_X_REWRITE_URL"),"?page=" & CStr(intBlogPaging),"?page=" & CStr(intBlogPaging - k))
			end if
			strPagingButPrev = "<a class=""btn btn-skin btn-lg text-uppercase marginbot-10"" href=""" & strPagingURLPrev & """>Previous Page</a>"
	  end if
	end if

'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "home.asp"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="../shop/pc/pcStartSession.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check store on/off, start PC session, check affiliate ID
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<!--#include file="../shop/pc/prv_getSettings.asp"-->


<!--#include file="../shop/pc/header_wrapper.asp"-->

	<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi"><a href="/blog/">Multiple Monitors Blog</a><%=strBlogCatTitle%></h3>
							<p class="p-code marginbot-0"><%=strBlogCatDesc%></p>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->
	<!-- Section: product-detail -->
    <section id="blog-list" class="paddingtop-40 paddingbot-40 blog-page">
		<div class="container">
			<div class="row">
				<div class="col-xs-12 col-sm-8 blog-listing">
    				<%=strBlogList%>
					<%=strPagingButPrev%><%=strPagingButNext%>
				</div>
				<div class="col-sm-4 col-xs-12 blog-sidebar">
					<div class="sidebar-widget recent-posts">
						<h5 class="widget-title h-semi">Popular Posts</h5>
						<div class="sidewidget-content">
							<ul class="list-post widget-list">
								<!--#include file="popular.asp"-->
							</ul>
						</div>
					</div>
                    <div class="sidebar-widget recent-posts">
						<h5 class="widget-title h-semi">Recent Posts</h5>
						<div class="sidewidget-content">
							<ul class="list-post widget-list">
								<!--#include file="recent.asp"-->
							</ul>
						</div>
					</div>
					<div class="sidebar-widget post-categories">
						<h5 class="widget-title h-semi">Categories</h5>
						<div class="sidewidget-content">
							<ul class="list-category widget-list">
								<!--#include file="category.asp"-->
							</ul>
						</div>
					</div>
				</div>
			</div>
		</div>
	</section>

<!--#include file="../shop/pc/orderCompleteTracking.asp"-->
<!--#include file="../shop/pc/inc-Cashback.asp"-->
<!--#include file="../shop/pc/footer_wrapper.asp"-->