<!--#include file="../includes/common.asp" -->
<%
	session("Facebook")="0"
	
	'//Get Facebook Settings
	tmpRedirectURL=""
	
	query="SELECT pcFBS_TurnOnOff,pcFBS_OffMsg,pcFBS_AppID,pcFBS_RedirectURL,pcFBS_Header,pcFBS_Footer,pcFBS_PageWidth,pcFBS_CustomDisplay,pcFBS_CatImages,pcFBS_CatRow,pcFBS_CatRowsperPage,pcFBS_BType,pcFBS_PrdRow,pcFBS_PrdRowsPerPage,pcFBS_ShowSKU,pcFBS_ShowSmallImg FROM pcFacebookSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		session("Facebook")=rs("pcFBS_TurnOnOff")
		if IsNull(session("Facebook")) OR session("Facebook")="" then
			session("Facebook")="0"
		end if
		session("pcFBS_OffMsg")=rs("pcFBS_OffMsg")
		session("pcFBS_AppID")=rs("pcFBS_AppID")
		tmpRedirectURL=rs("pcFBS_RedirectURL")
		session("pcFBS_Header")=rs("pcFBS_Header")
		session("pcFBS_Footer")=rs("pcFBS_Footer")
		session("pcFBS_PageWidth")=rs("pcFBS_PageWidth")
		if IsNull(session("pcFBS_PageWidth")) OR session("pcFBS_PageWidth")="" OR session("pcFBS_PageWidth")="0" then
			session("pcFBS_PageWidth")=790
		else
			session("pcFBS_PageWidth")=clng(session("pcFBS_PageWidth"))-20
		end if
		session("pcFBS_CustomDisplay")=rs("pcFBS_CustomDisplay")
		session("pcFBS_CatImages")=rs("pcFBS_CatImages")
		session("pcFBS_CatRow")=rs("pcFBS_CatRow")
		session("pcFBS_CatRowsperPage")=rs("pcFBS_CatRowsperPage")
		session("pcFBS_BType")=rs("pcFBS_BType")
		session("pcFBS_PrdRow")=rs("pcFBS_PrdRow")
		session("pcFBS_PrdRowsPerPage")=rs("pcFBS_PrdRowsPerPage")
		session("pcFBS_ShowSKU")=rs("pcFBS_ShowSKU")
		session("pcFBS_ShowSmallImg")=rs("pcFBS_ShowSmallImg")
	end if
	set rs=nothing

	if session("Facebook")="1" then
		if tmpRedirectURL<>"" then
			response.Redirect(tmpRedirectURL)
		else
			response.Redirect("viewcategories.asp")
		end if
	else
		response.write session("pcFBS_OffMsg")
	end if
%>
	
	
	
	