<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'*******************************
' Check if store is ON or OFF
'*******************************

Function SEOcheckAff()
	Dim tmp1
	tmp1=getUserInput(request.querystring("idAffiliate"),0)
	if not validNum(tmp1) then
		tmp1=1
	end if
	if tmp1=1 then
		if session("strSEOAffiliate")<>"" then
			if IsNumeric(session("strSEOAffiliate")) then
				tmp1=session("strSEOAffiliate")
			end if
		elseif session("idaffiliate")<>"" then
			if IsNumeric(session("idaffiliate")) then
				tmp1=session("idaffiliate")
			end if
		end if
	end if
	if tmp1=1 then		
		pcv_SavedAffiliateID = getUserInput(Request.Cookies("SavedAffiliateID"),0)
		If validNum(pcv_SavedAffiliateID) then
			session("idAffiliate")=trim(pcv_SavedAffiliateID)
            tmp1=session("idaffiliate")
		End If	        
	END IF
	SEOcheckAff=tmp1
End Function

'*******************************
' START ProductCart Session
'*******************************
HaveNewSession=0
if session("idcustomer")="" then
	Dim pcv_intFlagNoLocal
	pcv_intFlagNoLocal=Cint(0)
	session("idPCStore")= scID
	session("idCustomer")=Cint(0)
	session("customerCategory")=Cint(0)
	session("customerType")=Cint(0)
	session("ATBCustomer")= Cint(0)
	session("ATBPercentOff")= Cint(0)  
	session("language")=Cstr("english")
	session("pcCartIndex")=Cint(0)	
	dim pcCartArrayORG(100,45)
	session("pcCartSession")=pcCartArrayORG
	HaveNewSession=1
end if
if session("idPCStore")<>scID then
	session.Abandon()
	session("idPCStore")= scID
	session("idCustomer")=Cint(0)
	session("customerCategory")=Cint(0)
	session("customerType")=Cint(0)
	session("ATBCustomer")= Cint(0)
	session("ATBPercentOff")= Cint(0)     
	session("language")=Cstr("english")
	session("pcCartIndex")=Cint(0)
	redim pcCartArrayORG(100,45)
	session("pcCartSession")=pcCartArrayORG
	HaveNewSession=1
end if
pcCartArray=session("pcCartSession")
'*******************************
' END ProductCart Session
'*******************************
%>
<!--#include file="../includes/pcAffConstants.asp"-->
<%
'*******************************
' AFFILIATE - START
'*******************************
	dim pcInt_IdAffiliate, pcv_SavedAffiliateID
	dim pcInt_UseAffiliate, pcv_AffiliateCookiePath
	
	session("pcInt_AllowedAffOrders")=scAllowedAffOrders

	IF scAffProgramActive="1" THEN
	
		session("idAffiliate") = SEOcheckAff()

		'// Set cookie with Affiliate ID, if feature is active
		If scSaveAffiliate="1" and session("idAffiliate")<>1 Then
			Response.Cookies("SavedAffiliateID")=session("idAffiliate")
			pcInt_SaveAffiliateDays=Cint(scSaveAffiliateDays)
			if NOT validNum(pcInt_SaveAffiliateDays) then
				pcInt_SaveAffiliateDays=365
			end if
			Response.Cookies("SavedAffiliateID").Expires=Date() + pcInt_SaveAffiliateDays
		end if
		
	END IF

'*******************************
' AFFILIATE - END
'*******************************

IF HaveNewSession=1 THEN%>
<!--#include file="inc_RestoreShoppingCart.asp"-->
<%END IF%>