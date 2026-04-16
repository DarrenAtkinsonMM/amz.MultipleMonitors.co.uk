<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true
pcStrPageName="viewcontent.asp" %>
<!--#include file="../includes/common.asp"-->
<!--#include file="pcStartSession.asp"-->

<% 
dim query1, query2
dim pcPageId, pcvPageName, pcvProductsView, pcIntPageParent, pcvPageMetaTitle, pcvPageMetaDesc, pcvPageMetaKeywords, pcvPageThumbnail, pcIntHideBackButton

pcv_Url=request("url")

iPageCurrent = getUserInput(Request("page"), 0)
iPageSize	= getUserInput(Request("iPageSize"), 0)

If IsNumeric(iPageCurrent) And iPageCurrent <> "" Then
	iPageCurrent = CInt(iPageCurrent)
Else
	iPageCurrent = 1
End If
	
If IsNumeric(iPageSize) And iPageSize <> "" Then
	iPageSize = CInt(iPageSize)
Else
	scCatTotal = (scCatRow*scCatRowsPerPage)
	iPageSize = scCatTotal
End If

'// Load data from Existing Pages - START

	'// Look for subBrand
	if session("idParentContentPageRedirect") <> "" then
		pcParentPageId=session("idParentContentPageRedirect")
		query1="pcCont_Parent="&pcParentPageId
		pcPageId=pcParentPageId
		pcvPageType = "parent"
	else
		pcPageId=session("idContentPageRedirect")
		query1="pcCont_Parent=0"
		pcvPageType = ""
	end if
	
	'if not validNum(pcPageId) then
		'pcPageId=trim(getUserInput(request("idpage"),10))
		'if not validNum(pcPageId) AND pcPageId<>"" then
			'call closeDb()
			'response.redirect "default.asp"
		'end if
	'end if
	
	'// Check for admin preview
	pcIntAdminPreview = getUserInput(request("adminPreview"),2)
	if not validNum(pcIntAdminPreview) then pcIntAdminPreview=0
	
	if NOT (pcIntAdminPreview = 1 OR session("admin") <> 0) AND validNum(pcPageId) then
		'// Check customer session whether they are allowed to view the page or not
		query = "SELECT pcCont_CustomerType FROM pcContents WHERE pcCont_IDPage = " & pcPageId
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		if not rstemp.eof then
			if (rstemp("pcCont_CustomerType")="W" AND session("customerType")<>1) OR (left(rstemp("pcCont_CustomerType"),3)="CC_" AND rstemp("pcCont_CustomerType") <> "CC_" & session("customerCategory")) then
				set rstemp=nothing
				call closeDb()
				response.redirect "msg.asp?message=318"
			end if
		End if
	End If
	
	if not validNum(pcPageId) then
		query1="pcCont_Parent=0"
	else
		query1="pcCont_Parent="&pcPageId
	end if

	'// Select pages compatible with customer type
	if session("customerCategory")<>0 then ' The customer belongs to a customer category
		' Load pages accessible by ALL, plus those accessible by the customer pricing category that the customer belongs to
		if session("customerType")=0 then
			' Customer category does NOT have wholesale privileges, so exclude those pages
			query2 = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
		else
			' Customer category HAS wholesale privileges, so include wholesale-only pages
			query2 = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType = 'W' OR pcCont_CustomerType='CC_" & session("customerCategory") &"')"
		end if
	else
		if session("customerType")=0 then
			' Retail customer or customer not logged in: load pages accessible by ALL
			query2 = " AND pcCont_CustomerType = 'ALL'"
		else
			' Wholesale customer: load pages accessible by ALL and Wholesale customers only
			query2 = " AND (pcCont_CustomerType = 'ALL' OR pcCont_CustomerType = 'W')"
		end if
	end if
	
	if pcIntAdminPreview = 1 and session("admin") <> 0 then
		query3 = ""
	else
		query3 = " AND pcCont_InActive=0 AND pcCont_Published=1"
	end if
	
	query="SELECT pcCont_IDPage, pcCont_PageName, pcCont_IncHeader, pcCont_MetaTitle, pcCont_Description, pcCont_MetaDesc, pcCont_MetaKeywords, pcCont_Order, pcCont_Parent, pcCont_Published, pcCont_Thumbnail FROM pcContents WHERE " & query1 & query2 & query3 & " AND pcCont_MenuExclude=0 ORDER BY pcCont_Order, pcCont_PageName ASC;"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.PageSize=iPageSize
	rs.CacheSize=iPageSize
	rs.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
	if not rs.eof then
		dim iPageCount
		iPageCount=rs.PageCount
		If iPageCurrent > iPageCount Then iPageCurrent=iPageCount
		If iPageCurrent < 1 Then iPageCurrent=1
			
		If iPageCount=0 Then 
			' There are no pages
			set rs=nothing
			call closeDb()
			response.redirect "msg.asp?message=300"       
		End if
		
		rs.AbsolutePage=iPageCurrent
	end if

'// Load data from Existing Pages - END

'// Load Parent Page Information - START
'if pcPageId<>"" then

	query="SELECT pcCont_IncHeader, pcCont_MetaTitle, pcCont_Description, pcCont_MetaDesc, pcCont_MetaKeywords, pcCont_Order, pcCont_Parent, pcCont_Published, pcCont_PageTitle, pcCont_PageName, pcCont_HideBackButton, pcCont_Blog, pcCont_BlogCat, pcCont_PubDate, pcCont_BlogIntro FROM pcContents WHERE pcUrl='" &  pcv_Url & "'"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		set rstemp=nothing
		call closeDb()
		response.redirect "msg.asp?message=85"       
	End if
	
	parentPageTitle=pcf_PrintCharacters(rstemp("pcCont_PageTitle"))
    parentPageContent=pcf_PrintCharacters(rstemp("pcCont_Description"))
	
	pcvPageName=pcf_PrintCharacters(rstemp("pcCont_PageName"))
	
	pcInt_Parent=rstemp("pcCont_Parent")
	if not validNum(pcInt_Parent) then pcInt_Parent=0

	pcv_DefaultTitle=rstemp("pcCont_MetaTitle")
	if isNull(pcv_DefaultTitle) or trim(pcv_DefaultTitle)="" then
		pcv_DefaultTitle=ClearHTMLTags2(parentPageTitle,0)
	end if
	pcv_DefaultTitle = pcv_DefaultTitle
	daMetaDescTest=Replace(rstemp("pcCont_BlogIntro"),"<p>","")
	pcv_DefaultDescription=Replace(daMetaDescTest,"</p>","")
	'pcv_DefaultDescription=rstemp("pcCont_BlogIntro")
	pcv_DefaultKeywords=rstemp("pcCont_MetaKeywords")
	
	pcIntHideBackButton=rstemp("pcCont_HideBackButton")
	if not validNum(pcIntHideBackButton) then pcIntHideBackButton=0
	
	
		'DA Redirect from /pages/ to /blog/ or vice versa depending on admin page setting
	'First check out current url
	If InStr(Request.ServerVariables("HTTP_X_ORIGINAL_URL"), "/pages/") = 0 then
		daBlogUrl = True
	Else
		daBlogUrl = False
	End if
	
	'Now work out which page we should be on and then redirect if necessary (0 = pages / 1 = blog)
	If rstemp("pcCont_Blog") = 0 then
		'We need to be on a pages url
		if daBlogUrl = True then
			'Perform 301 redirect
			daCorrectUrl = Replace(Request.ServerVariables("HTTP_X_ORIGINAL_URL"),"/blog/","/pages/")
			Response.Status="301 Moved Permanently"
			Response.AddHeader "Location",daCorrectUrl
		end if
	Else
		'We need to be on a blog url
		if daBlogUrl = False then
			'Perform 301 redirect
			daCorrectUrl = Replace(Request.ServerVariables("HTTP_X_ORIGINAL_URL"),"/pages/","/blog/")
			Response.Status="301 Moved Permanently"
			Response.AddHeader "Location",daCorrectUrl
		end if
	End If
	
	'Set Correct Blog Category
	select case CInt(rstemp("pcCont_BlogCat"))
		case 0
			daBlogCatLink = "<a href=""/blog/hardware/""><i class=""fa fa-folder-open""></i> Hardware</a>"
		case 1
			daBlogCatLink = "<a href=""/blog/product-guides/""><i class=""fa fa-folder-open""></i> Product Guides</a>"
		case 2
			daBlogCatLink = "<a href=""/blog/setup-guides/""><i class=""fa fa-folder-open""></i> Setup Guides</a>"
		case 3
			daBlogCatLink = "<a href=""/blog/software/""><i class=""fa fa-folder-open""></i> Software</a>"
		case 4
			daBlogCatLink = "<a href=""/blog/stands/""><i class=""fa fa-folder-open""></i> Stands</a>"
		case 5
			daBlogCatLink = "<a href=""/blog/trading/""><i class=""fa fa-folder-open""></i> Trading</a>"
	end select
	
	daBlogDate = MonthName(Month(rstemp("pcCont_PubDate"))) & ", " & Year(rstemp("pcCont_PubDate"))
	
	'Blog Page callout html sections
	daBlogCallOutTradingPC = "<div class=""blogCallOut""><div class=""wow fadeInUp"" data-wow-delay=""0.1s""><div class=""cta-text""><h3 class=""h-bold font-light disp-inline""><a href=""/pages/trading-computers/"">Interested in Trading Computers?</a></h3><p>Learn the important, need to know information before you buy a new <a href=""/pages/trading-computers/"">trading computer</a>. </p></div></div></div>"

	set rstemp=nothing
	
'end if
'// Load Parent Brand Information - END

pcIntContentPageID=pcPageId
pcvContentPageName=pcvPageName

call pcGenerateSeoLinks
%>

<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="inc_addThis.asp"-->
<% 
If daBlogUrl = False then
'Pages template
If pcIntHideBackButton = 0 then
%>
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="<%=pcvPageName%>"><%=pcvPageName%></h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->

	<section id="intWarranties" class="intWarranties paddingtop-30 paddingbot-70">	
           <div class="container">
				<div class="row">
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s">
<%
End If
%>
<%=parentPageContent%>
<% 
If pcIntHideBackButton = 0 then
%>
					</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->
<%
end if
else
'Blog Page Template
%>
	<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-10 h-semi"><%=pcvPageName%></h3>
							<div class="blog-meta">
                                <div class="h-post-meta meta-post-time"><i class="fa fa-calendar" style="margin-bottom:5px;"></i> Last Updated: <%=daBlogDate%></div>
								<div class="h-post-meta meta-post-category"><%=daBlogCatLink%></div>
							</div>
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
					<!-- Post: Start -->
					<div class="blog-detail-single">
						<div class="blog-content blog-list-content">
                        <%=Replace(parentPageContent,"XXXBCOTRADEXXX",daBlogCallOutTradingPC)%>
                        <p class="blogAuthorTxt"><i class="fa fa-user"></i> Written by Darren @ Multiple Monitors</p>
                        <p class="blogAuthorTxt marginbot-20"><i class="fa fa-calendar"></i> Last Updated: <%=daBlogDate%></p>
						</div>
					</div>
					<!-- Post: End -->
				</div>
				<div class="col-sm-4 col-xs-12 blog-sidebar">
					<div class="sidebar-widget recent-posts">
						<h5 class="widget-title h-semi">Popular Posts</h5>
						<div class="sidewidget-content">
							<ul class="list-post widget-list">
								<!--#include file="../../new-blog/popular.asp"-->
							</ul>
						</div>
					</div>
					<div class="sidebar-widget recent-posts">
						<h5 class="widget-title h-semi">Recent Posts</h5>
						<div class="sidewidget-content">
							<ul class="list-post widget-list">
								<!--#include file="../../new-blog/recent.asp"-->
							</ul>
						</div>
					</div>
					<div class="sidebar-widget post-categories">
						<h5 class="widget-title h-semi">Categories</h5>
						<div class="sidewidget-content">
							<ul class="list-category widget-list">
								<!--#include file="../../new-blog/category.asp"-->
							</ul>
						</div>
					</div>
				</div>
			</div>
		</div>
	</section>
<%
end if
%>
<%
session("idParentContentPageRedirect")=""
session("idContentPageRedirect")=""
%>
<!--#include file="footer_wrapper.asp"-->
