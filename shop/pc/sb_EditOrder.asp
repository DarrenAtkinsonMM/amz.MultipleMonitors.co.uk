<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/CashbackConstants.asp"-->
<!--#include file="chkPrices.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<!--#include file="../includes/pcSBHelperInc.asp"-->
<!--#include file="pcStartSession.asp"-->
<%
Dim SBIDOrder,pcv_strGUID,SBGuid

SBIDOrder=getUserInput(request("ID"),0)
if (SBIDOrder="") OR (not IsNumeric(SBIDOrder)) then
	response.redirect "CustPref.asp"
end if
pcv_strGUID=getUserInput(request("GUID"),0)
if (pcv_strGUID="") then
	response.redirect "CustPref.asp"
end if

	SBGuid=pcv_strGUID
	
	query="SELECT TOP 1 idOrder,SB_Terms FROM SB_Orders WHERE SB_Guid like '" & SBGuid & "' AND idOrder=" & SBIDOrder & " ORDER BY idOrder DESC;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		SBIDOrder=rs("idOrder")
		SBTerms=rs("SB_Terms")
	else
		set rs=nothing
		call closedb()
		response.redirect "CustPref.asp"
	end if
	set rs=nothing
	
	IF SBIDOrder>"0" THEN
		
		'Repeat Order/Generate shopping cart array
		ReErr=0%>
		<!--#include file="sb_inc_repeatorder.asp"-->
		<%
		
		If (ReErr=0) OR (ReErr=5) then
			Session("SBEditOrder")=SBGuid
			Session("SBEditOrderID")=SBIDOrder
			response.redirect "viewcart.asp"
		Else
			response.redirect "CustPref.asp"
		End if
	
	END IF


' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function

%>