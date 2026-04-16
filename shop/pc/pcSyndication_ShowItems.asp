<!DOCTYPE html>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/SocialNetworkWidgetConstants.asp"-->
<%Dim MaxH
MaxH=getUserInput(Request("mh"), 8)
If not validNum(MaxH) then
	MaxH=438
End If

MaxW=getUserInput(Request("mw"), 8)
If not validNum(MaxW) then
	MaxW=198
End If
%>
<html lang="en">
    <head>
        <!--#include file="inc_jquery.asp"-->
        <link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("css","pcStorefront.css")%>" />
        <link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("css","pcSyndication.css")%>" />
		<style>
			.pcSyndicationRegion {
				max-height: <%=MaxH%>px;
				width: <%=MaxW%>px;
				overflow-x: hidden;
				overflow-y: auto;
			}
			#pcSyndication {
				width: <%=MaxW-8%>px;
			}
		</style>
    </head>
    <%
    '// Get the Affiliate ID (if any)
    pcv_SavedAffiliateID=""
    pcv_SavedAffiliateID= getUserInput(Request("idaffiliate"), 8)
	If validNum(pcv_SavedAffiliateID) then
		pcInt_IdAffiliate=pcv_SavedAffiliateID
	Else
		pcInt_IdAffiliate=1
	End If
	idaffiliate = pcv_SavedAffiliateID
	if SNW_AFFILIATE="1" then
		session("idAffiliate") = pcv_SavedAffiliateID
	end if
    %>
    <body style="background-color:transparent" data-ng-controller="syndicationCtrl">
        <div id="pcSyndication" class="pcSyndicationRegion ng-cloak" data-ng-cloak>        
        
            <div id="pcProductRegion">
                <div data-ng-show="displayItems()">                
                    <p id="pcSyndicationBox" align="center" data-ng-repeat="syndicationitem in syndicationlist.syndicationitemrow">  
                        <span>
                            <img data-ng-src="{{syndicationitem.image}}"><br />
                        </span>				
                        <span class="pcSyndicationName"><a href="{{syndicationitem.url}}<% if SNW_AFFILIATE="1" then%>&idaffiliate=<%=idaffiliate%><%End If%>" target="_blank" class="SyndicationImage">{{syndicationitem.description}}</a></span>
                        <br/><span class="pcSyndicationPrice"><%=scCurSign%>{{syndicationitem.price}}</span>
                    </p>                    
              	</div>                
                <div data-ng-show="!displayItems()">
                    <div>No Items</div>			
                </div>	                
            </div>
        
        </div>        
        <script src="<%=pcf_getJSPath("../includes/javascripts","jquery.blockUI.js")%>"></script>
        <script src="<%=pcf_getJSPath("../includes/javascripts","json3.js")%>"></script>
        <script src="<%=pcf_getJSPath("../includes/javascripts","angular-1.0.8.js")%>"></script>
        <script src="<%=pcf_getJSPath("../includes/javascripts","accounting.min.js")%>"></script>
        <script src="<%=pcf_getJSPath("service/app","service.js")%>"></script>
        <script src="<%=pcf_getJSPath("service/app","syndication.js")%>"></script>
        <script src="<%=pcf_getJSPath("service/app","search.js")%>"></script>
    </body>
</html>
<% call closeDb() %>
